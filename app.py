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

# --- [ 특정 예외 주소 및 상가/아파트 분리 강제 매핑 사전 ] ---
SPECIAL_EXCEPTIONS = {
    "불로동 268-2": "인천광역시 검단구 금정로 12 (불로동, 신검단중앙역풍경채어바니티)",
    "남산타운": "서울특별시 중구 다산로 32 (신당동, 남산타운)",
    "율곡동 415": "경상북도 김천시 혁신8로 66",
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

# --- [ 핵심: API 조회용 본문과 상세 주소(동·호수·상호명 등) 사전 격리 함수 ] ---
def isolate_address_details(addr_str):
    if not addr_str or pd.isna(addr_str):
        return "", "", ""
    
    s = str(addr_str).strip()
    
    # 1. 맨 끝 법정동 괄호 분리 (예: ( 문원동 ))
    trailing_paren = ""
    m_paren = re.search(r'(\([^)]+\))\s*$', s)
    if m_paren:
        trailing_paren = m_paren.group(1).strip()
        s = s[:m_paren.start()].strip()
        
    # 2. API가 누락하는 상세 정보(호, 층, 동, 빌라명, 상호명, 부서명 등)를 사전에 분리 포획
    # 주소 본문(시/도 ~ 도로명/지번 + 건물번호) 뒤에 붙는 상세 꼬리말 추출
    detail_suffix = ""
    
    # 대표적인 상세 정보 패턴 분리지점 탐색 (호, 층, 단독, 또는 특정 상호/시설명 키워드 기준)
    m_detail = re.search(r'(?:호|층|동|단독|\d+)\s+([가-힣A-Za-z0-9\.]{2,})$', s)
    if m_detail:
        potential_detail = m_detail.group(1)
        if not any(w in potential_detail for w in ['로', '길', '동', '읍', '면', '리', '시', '군', '구']) or any(k in potential_detail for k in ['샵', '씨앤비', '리본도', '관리', '팀', '총무부', '지층', '학교']):
            detail_suffix = potential_detail
            s = s[:m_detail.start() + len(m_detail.group(0)) - len(potential_detail)].strip()
            
    # 추가로 건물 중간에 들어가는 소규모 빌라명(예: 까사빌)이나 호수 보호 처리
    # API 쿼리용 순수 위치 주소(s)와 보존할 상세 정보(detail_suffix, trailing_paren) 반환
    return s, detail_suffix, trailing_paren

# --- [ 주소 전처리 및 동·호수 포맷팅 함수 ] ---
def pre_process_address(addr_str):
    if not addr_str:
        return addr_str
    # 1. 슬래시 형태(105/2202)를 105동 2202호로 변환
    addr_str = re.sub(r'\b(\d+)\s*/\s*(\d+)\b', r'\1동 \2호', addr_str)
    
    # 2. 아파트/빌라명/빌 명 뒤의 하이픈 동·호수 변환
    def replace_bldg_hyphen(match):
        prefix = match.group(1)
        d = match.group(2)
        h = match.group(3)
        return f"{prefix} {d}동 {h}호"
    
    addr_str = re.sub(r'([가-힣A-Za-z0-9]+아파트|[가-힣A-Za-z0-9]+빌라|[가-힣A-Za-z0-9]+빌|[가-힣A-Za-z0-9]+맨션|[가-힣A-Za-z0-9]+타운)\s+(\d{1,4})-(\d{1,4})호?', replace_bldg_hyphen, addr_str)
    addr_str = re.sub(r'\b(\d{1,4})-(\d{1,4})호\b', r'\1동 \2호', addr_str)

    # 3. 한글 보조 동 포맷팅 (비동, 가동 등)
    addr_str = re.sub(r'\b(\d+\s*동|[가-하A-Za-z]\s*동)\s*(\d+)(?!호)\b', r'\1 \2호', addr_str)
    addr_str = re.sub(r'\b([가-하A-Za-z]동)(\d+호)\b', r'\1 \2', addr_str)
    
    return addr_str

# --- [ 상품명 정제 함수 ] ---
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
        elif raw_model == '와이드6': target_model = 'A13'
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
    elif target == '와이드6': target = 'A13'
    
    result_str = f"{target}[{qty}]" if qty > 1 else target
    if '택배' in s or '택배배송' in s: result_str = f"{result_str} 택배"
    if has_film: result_str = f"{result_str} +필름"
        
    return result_str

# --- [ 마스터 주소 변환 엔진 (상세 주소 사전 격리 및 최종 재조합 통합) ] ---
def master_juso_converter(keyword, original_zip=""):
    if not keyword or pd.isna(keyword):
        return keyword
        
    kw_str = str(keyword).strip()
    
    # 1. API 조회용 순수 본문과 보존할 상세 정보(동·호수·상호명 등) 사전 격리
    api_query_base, detail_suffix, trailing_paren = isolate_address_details(kw_str)
    
    # 2. API 조회용 본문 전처리
    api_query_base = pre_process_address(api_query_base)
    
    # 3. 특수 예외 주소 사전 확인
    base_road_addr = ""
    for target_key, override_addr in SPECIAL_EXCEPTIONS.items():
        if target_key in kw_str:
            base_road_addr = override_addr
            break
            
    # 4. 네이버 API Geocoding 조회
    if not base_road_addr:
        if is_new_road_address(api_query_base):
            base_road_addr = api_query_base
        else:
            url = f"https://maps.apigw.ntruss.com/map-geocode/v2/geocode?query={urllib.parse.quote(api_query_base)}"
            headers = {
                "X-NCP-APIGW-API-KEY-ID": NAVER_CLIENT_ID,
                "X-NCP-APIGW-API-KEY": NAVER_CLIENT_SECRET
            }
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
                
    if not base_road_addr:
        base_road_addr = api_query_base

    # 5. 최종 재조합: 변환된 도로명 주소 + 사전에 격리해 둔 상세 정보(동·호수·상호명 등) + 법정동 괄호
    parts = [base_road_addr]
    if detail_suffix and detail_suffix not in base_road_addr:
        parts.append(detail_suffix)
    if trailing_paren and trailing_paren not in base_road_addr:
        parts.append(trailing_paren)
        
    return pre_process_address(' '.join(parts))

# --- [ Streamlit 웹 UI 실행 부문 ] ---
st.set_page_config(page_title="스마트샵 주소 변환기", page_icon="🛍️", layout="centered")

st.title("🛍️ 스마트샵 주소 변환 & 엑셀 수정 (네이버 API 연동)")
st.write("엑셀 파일을 업로드하면 네이버 주소 API를 통해 도로명 주소 변환 및 상세 동·호수·상호명 유실 방지 처리가 이루어집니다.")

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
            file_name="우편배송_변환결과_완벽보존.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
