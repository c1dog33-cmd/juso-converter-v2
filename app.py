import io
import re
import urllib.parse
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font

# 🚀 네이버 클라우드 플랫폼 API 인증 정보
NAVER_CLIENT_ID = "bkwdzf72zn"
NAVER_CLIENT_SECRET = "rbzw7TF0T8wDDFB4HqoZ4rNaq14zKDgc8W9ogQX6"

def fix_zipcode(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).split('.')[0].strip()
    return val_str.zfill(5) if val_str else ""

# --- [ 특정 예외 주소 및 행정개편 강제 매핑 사전 ] ---
SPECIAL_EXCEPTIONS = {
    "불로동 268-2": "인천광역시 검단구 금정로 12 (불로동, 신검단중앙역풍경채어바니티)",
    "남산타운": "서울특별시 중구 다산로 32 (신당동, 남산타운)",
}

# --- [ openpyxl 제어 문자 에러 방지 함수 ] ---
def remove_illegal_chars(val):
    if isinstance(val, str):
        return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', val)
    return val

# --- [ 신주소(도로명 주소) 판별 함수 ] ---
def is_new_road_address(addr_str):
    if not addr_str:
        return False
    road_pattern = r'[로길]\s*\d+(?:번길)?\s+\d+'
    return bool(re.search(road_pattern, addr_str))

# --- [ 주소 전처리 및 동·호수 포맷팅 함수 ] ---
def pre_process_address(addr_str):
    if not addr_str:
        return addr_str
    # 1. 슬래시 형태(105/2202)를 105동 2202호로 변환
    addr_str = re.sub(r'\b(\d+)\s*/\s*(\d+)\b', r'\1동 \2호', addr_str)
    # 2. 하이픈 동·호수 형태(101-305) 변환
    addr_str = re.sub(r'\b(\d+)-(\d{3,4})\b', r'\1동 \2호', addr_str)
    # 3. 띄어쓰기 보정 (711동904호 -> 711동 904호)
    addr_str = re.sub(r'(\d+동)(\d+호)', r'\1 \2', addr_str)
    return addr_str

# --- [ 신주소 상세 정보(동·호수/층수) 맨 뒤로 정렬 함수 ] ---
def reorder_dong_ho_to_end(addr_str):
    if not addr_str:
        return addr_str
    dong_ho_pattern = r'(\b\d+\s*동\s*\d+\s*호?|\b[A-Za-z가나다라마바사아자차카타파하]\s*동\s*\d+\s*호?|\b\d+\s*층\s*\d+\s*호?|\b\d+\s*층|\b\d+\s*호)'
    matches = re.findall(dong_ho_pattern, addr_str)
    if not matches:
        return addr_str
        
    clean_str = addr_str
    for m in matches:
        clean_str = clean_str.replace(m, '')
        
    clean_str = ' '.join(clean_str.split())
    return clean_str + " " + " ".join(matches)

# --- [ 완벽한 상품명 정제 함수 ] ---
def clean_product_name(val):
    if pd.isna(val) or not str(val).strip():
        return val
    s = str(val).strip()
    
    has_film = False
    if ('풀액정필름2장' in s or '추가액정필름' in s) and '필요없음' not in s and ('풀액정' in s or '액정필름' in s):
        has_film = True

    target_texts = [s]
    if '//' in s:
        target_texts.insert(0, s.split('//')[-1])
    if '모델선택:' in s:
        target_texts.insert(0, s.split('모델선택:')[-1])
    if '/' in s:
        target_texts.insert(0, s.split('/')[-1])

    smartphone_match = None
    for t in target_texts:
        smartphone_match = re.search(r'(S\s*10\s*5G|S\s*10\s*플러스|S\s*20\s*플러스|S\s*20\s*울트라|S\s*22\s*플러스|S\s*22\s*울트라|노트\s*10\s*플러스|노트\s*20\s*울트라|노트\s*20\s*플러스|노트\s*\d+\s*울트라|노트\s*\d+\s*플러스|노트\s*\d+|S\s*\d+(?:\s*플러스|\s*울트라|\s*5G|\s*엣지)?|A\s*\d+|Z\s*(?:플립|폴드)\s*\d*|와이드\s*\d+|N\s*\d+|V\s*50S|V\s*50|V\s*\d+|벨벳|윙)', t, re.IGNORECASE)
        if smartphone_match:
            break

    if not smartphone_match:
        smartphone_match = re.search(r'(S\s*10\s*5G|S\s*10\s*플러스|S\s*20\s*플러스|S\s*20\s*울트라|S\s*22\s*플러스|S\s*22\s*울트라|노트\s*10\s*플러스|노트\s*20\s*울트라|노트\s*20\s*플러스|노트\s*\d+\s*울트라|노트\s*\d+\s*플러스|노트\s*\d+|S\s*\d+(?:\s*플러스|\s*울트라|\s*5G|\s*엣지)?|A\s*\d+|Z\s*(?:플립|폴드)\s*\d*|와이드\s*\d+|N\s*\d+|V\s*50S|V\s*50|V\s*\d+|벨벳|윙)', s, re.IGNORECASE)

    if smartphone_match:
        raw_model = smartphone_match.group(1).replace(" ", "").upper()
        if raw_model == 'S105G': target_model = 'S10 5G'
        elif raw_model == 'S10플러스': target_model = 'S10플러스'
        elif raw_model == 'S20플러스': target_model = 'S20플러스'
        elif raw_model == 'S22플러스': target_model = 'S22플러스'
        elif raw_model == '노트10플러스': target_model = '노트10플러스'
        elif raw_model in ['N960', 'SM-N960']: target_model = '노트9'
        elif raw_model == 'V50S': target_model = 'V50S'
        elif raw_model == 'V50': target_model = 'V50'
        elif raw_model == '노트20울트라': target_model = '노트20울트라'
        elif raw_model == '와이드5': target_model = 'A22'
        else: target_model = raw_model

        qty = 1
        if '☞' in s:
            parts = s.split('☞')
            after = parts[1].strip() if len(parts) > 1 else ""
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
        
        res = f"{target_model}[{qty}]" if qty > 1 else target_model
        if has_film: res += " +필름"
        if '택배' in s or '택배배송' in s: res += " 택배"
        return res
    
    if '케이스' not in s:
        res = s
        if has_film: res += " +필름"
        return res
    
    qty = 1
    if '☞' in s:
        parts = s.split('☞')
        base = parts[0].strip()
        after = parts[1].strip() if len(parts) > 1 else ""
        qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
        if qty_match:
            qty = int(qty_match.group(1))
    else:
        base = s
        
    if '모델선택:' in base:
        parts = base.split('모델선택:')
        target = parts[-1].strip()
        if '/' in target: target = target.split('/')[0].strip()
        if ',' in target: target = target.split(',')[0].strip()
    else:
        sub_parts = re.split(r'[/]{1,2}', base)
        target = sub_parts[-1].strip() if len(sub_parts) > 1 else base
        
    target = re.sub(r'\([^)]*\)', '', target).strip()
    target = re.sub(r'케이스', '', target)
    target = re.sub(r'풀액정', '', target)
    target = re.sub(r'2장', '', target)
    target = re.sub(r'클리어', '', target)
    target = re.sub(r'투명', '', target)
    target = re.sub(r'갤럭시', '', target)
    target = ' '.join(target.split())
    
    if target == '와이드5': target = 'A22'
    
    result_str = f"{target}[{qty}]" if qty > 1 else target
    if '택배' in s or '택배배송' in s: result_str = f"{result_str} 택배"
    if has_film: result_str = f"{result_str} +필름"
        
    return result_str

# --- [ 마스터 주소 변환 엔진 ] ---
def master_juso_converter(keyword, original_zip=""):
    if not keyword or pd.isna(keyword):
        return keyword
        
    kw_str = str(keyword).strip()
    kw_str_pre = pre_process_address(kw_str)
    
    # 1. 이미 신주소(도로명 주소)인 경우 상세 정보 맨 뒤로 정렬하여 반환
    if is_new_road_address(kw_str_pre):
        return reorder_dong_ho_to_end(kw_str_pre)

    # 2. 건물·아파트·빌리지·고시원·프라자 이름 추출
    bldg_match = re.search(r'([가-힣A-Za-z0-9]+(?:아파트|리버스트|빌리지|빌라|오피스텔|맨션|하이츠|타운|팰리스|주공|LH|단지|고시원|프라자|빌딩|상가))', kw_str_pre)
    orig_bldg = bldg_match.group(1) if bldg_match else ""

    # 3. 상세 동/호수, 층, 괄호 및 상호명 패턴 추출 (하이픈 지번 오인식 방지 패턴 적용)
    extra_pattern = r'(?:(?<!-)\b\d+\s*동\s*\d+\s*호?|(?<!-)\b\d+\s*동|\b\d+\s*호|\b[A-Za-z가나다라마바사아자차카타파하]\s*동\s*\d+\s*호?|\b[A-Za-z가나다라마바사아자차카타파하]\s*동|\d+층|\([^)]+\)|생활관\s*[가-힣A-Za-z0-9]+|\b[가-힣]+상가.*)'
    extra_details = re.findall(extra_pattern, kw_str_pre)
    
    # 4. API 검색용 스마트 쿼리 구성 (지번 오류 방지를 위해 건물명 필수 조합)
    temp_query_str = kw_str_pre
    for ext in extra_details:
        temp_query_str = temp_query_str.replace(ext, '')
    if orig_bldg:
        temp_query_str = re.sub(r'\b' + re.escape(orig_bldg) + r'\b', '', temp_query_str)
    temp_query_str = re.sub(r'\s+0\b', '', temp_query_str)
    base_region = ' '.join(temp_query_str.split())
    
    search_q_str = f"{base_region} {orig_bldg}" if orig_bldg else base_region

    for target_key, override_addr in SPECIAL_EXCEPTIONS.items():
        if target_key in kw_str:
            base_road_addr = override_addr
            break
    else:
        url = f"https://maps.apigw.ntruss.com/map-geocode/v2/geocode?query={urllib.parse.quote(search_q_str)}"
        headers = {
            "X-NCP-APIGW-API-KEY-ID": NAVER_CLIENT_ID,
            "X-NCP-APIGW-API-KEY": NAVER_CLIENT_SECRET
        }
        
        base_road_addr = ""
        try:
            response = requests.get(url, headers=headers, timeout=5)
            if response.status_code == 200:
                res = response.json()
                addresses = res.get('addresses', [])
                if addresses:
                    addr_info = addresses[0]
                    road_addr = addr_info.get('roadAddress', '')
                    jibun_addr = addr_info.get('jibunAddress', '')
                    base_road_addr = road_addr if road_addr else jibun_addr
        except Exception:
            pass

    # 5. API 변환 실패 시 원본 전처리 주소 유지 (동·호수 맨 뒤 정렬 적용)
    if not base_road_addr:
        return reorder_dong_ho_to_end(kw_str_pre)

    # 6. 네이버 결과에 건물명이 빠져 있으면 보존·추가
    if orig_bldg and orig_bldg not in base_road_addr:
        base_road_addr = f"{base_road_addr} {orig_bldg}"

    # 7. 상세 정보 결합 (중복 괄호 스마트 필터링)
    for detail in extra_details:
        detail_clean = pre_process_address(detail.strip())
        if detail_clean.startswith('(') and detail_clean.endswith(')'):
            if any(sub.strip() in base_road_addr for sub in detail_clean.strip('()').split(',')):
                continue
        if detail_clean and detail_clean not in base_road_addr:
            base_road_addr = f"{base_road_addr} {detail_clean}"

    # 8. 중복 연속 단어 제거
    words = base_road_addr.split()
    clean_words = []
    for w in words:
        clean_w = w.strip('(),')
        if not clean_words or clean_w != clean_words[-1].strip('(),'):
            clean_words.append(w)
            
    return pre_process_address(' '.join(clean_words))

# --- [ Streamlit 웹 UI ] ---
st.set_page_config(page_title="스마트샵 주소 변환기", page_icon="🛍️", layout="centered")

st.title("🛍️ 스마트샵 주소 변환 & 엑셀 수정 (네이버 API 연동)")
st.write("엑셀 파일을 업로드하면 네이버 주소 API를 통해 도로명 주소 변환, 아파트/빌리지/고시원/프라자명 보존, 중복 괄호 제거 및 상품명 정리가 이루어집니다.")

uploaded_file = st.file_uploader("변환할 엑셀 파일(.xlsx, .xls)을 업로드하세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("### 📄 업로드 데이터 미리보기 (상위 5건)")
    st.dataframe(df.head())

    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None

    if st.button("🚀 주소 변환 및 서식 적용 시작"):
        with st.spinner("주소를 변환하는 중입니다..."):
            if '우편번호' in df.columns:
                df['우편번호'] = df['우편번호'].apply(fix_zipcode)

            if '선택정보' in df.columns:
                df['선택정보'] = df['선택정보'].apply(clean_product_name)

            if '배송지' in df.columns:
                progress_bar = st.progress(0)
                total = len(df)
                
                converted_addrs = []
                for i, row in df.iterrows():
                    orig_zip = str(row['우편번호']).zfill(5) if '우편번호' in df.columns and not pd.isna(row['우편번호']) else ""
                    res = master_juso_converter(row['배송지'], original_zip=orig_zip)
                    converted_addrs.append(res)
                    progress_bar.progress((i + 1) / total)
                    
                df['배송지'] = converted_addrs
            else:
                st.warning("[주의] '배송지' 컬럼을 찾을 수 없습니다.")

            for col in df.columns:
                df[col] = df[col].apply(remove_illegal_chars)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                
                for col in worksheet.columns:
                    max_len = 0
                    col_letter = openpyxl.utils.get_column_letter(col[0].column)
                    for cell in col:
                        val = str(cell.value or '')
                        length = sum(2 if ord(char) > 127 else 1 for char in val)
                        if length > max_len:
                            max_len = length
                    adjusted_width = max(max_len + 4, 12)
                    worksheet.column_dimensions[col_letter].width = min(adjusted_width, 80)

                font_size_8 = Font(size=8)
                for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=len(df.columns)):
                    worksheet.row_dimensions[row[0].row].height = 18
                    for cell in row:
                        cell.font = font_size_8
                        
                if '우편번호' in df.columns:
                    zip_col_idx = df.columns.get_loc('우편번호') + 1
                    for row in range(2, worksheet.max_row + 1):
                        worksheet.cell(row=row, column=zip_col_idx).number_format = '@'

            st.session_state.excel_data = output.getvalue()
        st.success("🎉 주소 변환 및 상품명 정리가 완료되었습니다!")

    if st.session_state.excel_data is not None:
        st.download_button(
            label="📥 변환된 엑셀 파일 다운로드",
            data=st.session_state.excel_data,
            file_name="우편배송_변환결과_네이버.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
