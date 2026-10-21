import io
import re
import urllib.parse
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font

# 🚀 카카오 REST API 키 설정 완료
KAKAO_API_KEY = "880d50b6016d6833ccc3ff4695ae7f92"

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

# --- [ 상품명 정제 함수 ] ---
def clean_product_name(val):
    if pd.isna(val) or not str(val).strip():
        return val
    s = str(val).strip()
    
    has_film = False
    if ('풀액정필름2장' in s or '추가액정필름' in s) and '필요없음' not in s and ('풀액정' in s or '액정필름' in s):
        has_film = True

    if 'A20/A30' in s or ('A20' in s and 'A30' in s and '와이드4' in s):
        qty = 1
        if '☞' in s:
            parts = s.split('☞')
            after = parts[1].strip() if len(parts) > 1 else ""
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
        res = f"A20/A30[{qty}]" if qty > 1 else "A20/A30"
        if has_film: res += " +필름"
        if '택배' in s: res += " 택배"
        return res

    if '/' in s and '☞' in s:
        parts = s.split('☞')[0].split('/')
        target_p = ""
        for p in parts:
            if '+' in p and ('필름' in p or '강화유리' in p):
                target_p = p.strip()
                break
        if target_p:
            model_part = target_p.split('+')[0].strip()
            qty = 1
            after_arrow = s.split('☞')[1].strip()
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after_arrow, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
            base = f"{model_part}+필름"
            res = f"{base}[{qty}]" if qty > 1 else base
            if '택배' in s: res += " 택배"
            return res

    if '/' in s:
        slash_parts = s.split('/')
        for p in reversed(slash_parts):
            m = re.search(r'(S\s*10\s*5G|노트\s*20\s*울트라|노트\s*20\s*플러스|노트\s*10\s*플러스|노트\s*\d+\s*울트라|노트\s*\d+\s*플러스|노트\s*\d+|S\s*\d+(?:\s*플러스|\s*울트라|\s*5G|\s*엣지)?|A\s*\d+|Z\s*(?:플립|폴드)\s*\d*|와이드\s*\d+|N\s*\d+|V\s*50S|V\s*50)', p, re.IGNORECASE)
            if m:
                raw_model = m.group(1).replace(" ", "").upper()
                if raw_model == 'S105G': target_model = 'S10 5G'
                elif raw_model == 'S20플러스' or raw_model == 'S20+': target_model = 'S20플러스'
                elif raw_model == '노트10플러스': target_model = '노트10플러스'
                elif raw_model == '노트20울트라': target_model = '노트20울트라'
                elif raw_model == '와이드5': target_model = 'A22'
                else: target_model = raw_model
                
                qty = 1
                if '☞' in s:
                    after_arrow = s.split('☞')[1].strip()
                    qty_match = re.match(r'^(\d+)\s*(EA|개)?', after_arrow, re.IGNORECASE)
                    if qty_match:
                        qty = int(qty_match.group(1))
                res = f"{target_model}[{qty}]" if qty > 1 else target_model
                if has_film: res += " +필름"
                if '택배' in s: res += " 택배"
                return res

    if '색상선택:' in s or '옵션선택:' in s:
        qty = 1
        if '☞' in s:
            parts = s.split('☞')
            after = parts[1].strip() if len(parts) > 1 else ""
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
        
        opt_key = '색상선택:' if '색상선택:' in s else '옵션선택:'
        sub_part = s.split(opt_key)[1].strip()
        opt_val = re.split(r'[/,☞]', sub_part)[0].strip()
        opt_val = re.sub(r'\d+[개EA]*$', '', opt_val).strip()
        
        if opt_val == '와이드5': opt_val = 'A22'
        
        res = f"{opt_val}[{qty}]" if qty > 1 else opt_val
        if has_film: res += " +필름"
        if '택배' in s: res += " 택배"
        return res

    if '사이드미러' in s or '클리어커버' in s or '컬러커버' in s:
        if '클리어커버' in s:
            target = '클리어커버'
        elif '컬러커버' in s:
            target = '컬러커버'
        else:
            target = '사이드미러가드'
            
        last_part = s.split('/')[-1] if '/' in s else s
        qty = 1
        m_gae = re.search(r'(\d+)\s*개', last_part)
        if m_gae:
            qty = int(m_gae.group(1))
        else:
            m_ea = re.search(r'☞\s*(\d+)', last_part)
            if m_ea:
                qty = int(m_ea.group(1))
                
        res = f"{target}[{qty}]" if qty > 1 else target
        if has_film: res += " +필름"
        if '택배' in s: res += " 택배"
        return res

    smartphone_match = re.search(r'(S\s*10\s*5G|노트\s*20\s*울트라|노트\s*20\s*플러스|노트\s*10\s*플러스|노트\s*\d+\s*울트라|노트\s*\d+\s*플러스|노트\s*\d+|S\s*\d+(?:\s*플러스|\s*울트라|\s*5G|\s*엣지)?|A\s*\d+|Z\s*(?:플립|폴드)\s*\d*|와이드\s*\d+|N\s*\d+|V\s*50S|V\s*50|V\s*\d+|벨벳|윙)', s, re.IGNORECASE)
    if smartphone_match:
        raw_model = smartphone_match.group(1).replace(" ", "").upper()
        if raw_model == 'S105G': target_model = 'S10 5G'
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
        if '택배' in s: res += " 택배"
        return res
    
    if '렌즈' in s or 'Lens Guards' in s or '렌즈가드' in s:
        qty = 1
        if '☞' in s:
            parts = s.split('☞')
            after = parts[1].strip() if len(parts) > 1 else ""
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
        res = f"렌즈보호커버[{qty}]" if qty > 1 else "렌즈보호커버"
        if has_film: res += " +필름"
        return res

    if ('송풍구' in s or '차량용' in s) and ('거치대' in s or '태블릿' in s):
        qty = 1
        if '☞' in s:
            parts = s.split('☞')
            after = parts[1].strip() if len(parts) > 1 else ""
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
        res = f"태블릿 송풍구 거치대[{qty}]" if qty > 1 else "태블릿 송풍구 거치대"
        if has_film: res += " +필름"
        return res

    if '모니터' in s and '2단' in s:
        qty = 1
        if '☞' in s:
            parts = s.split('☞')
            after = parts[1].strip() if len(parts) > 1 else ""
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
        res = f"태블릿 모니터 2단 거치대[{qty}]" if qty > 1 else "태블릿 모니터 2단 거치대"
        if has_film: res += " +필름"
        return res
    
    if '케이블' in s:
        qty = 1
        if '☞' in s:
            parts = s.split('☞')
            after = parts[1].strip() if len(parts) > 1 else ""
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
                
        has_g = 'ㄱ자' in s or 'ㄱ 자' in s
        has_c_type = 'C타입' in s or 'C 타입' in s or 'C-type' in s or 'c타입' in s
        
        length_str = ""
        if '/' in s:
            cable_parts = s.split('/')
            for p in reversed(cable_parts):
                m_sub = re.search(r'(\d+(?:\.\d+)?(?:m|cm))', p, re.IGNORECASE)
                if m_sub:
                    length_str = m_sub.group(1)
                    break
        if not length_str:
            m_match = re.search(r'(\d+(?:\.\d+)?(?:m|cm))', s, re.IGNORECASE)
            if m_match:
                length_str = m_match.group(1)
        
        components = []
        if has_g: components.append('ㄱ자형')
        if has_c_type: components.append('C타입 케이블')
        else: components.append('케이블')
        if length_str: components.append(length_str)
            
        target = " ".join(components)
        res = f"{target}[{qty}]" if qty > 1 else target
        if has_film: res += " +필름"
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
    if '택배' in s: result_str = f"{result_str} 택배"
    if has_film: result_str = f"{result_str} +필름"
        
    return result_str

def remove_duplicate_words(addr_str):
    if not addr_str:
        return addr_str
    words = addr_str.split()
    clean_words = []
    for w in words:
        clean_w = w.strip('(),')
        if not clean_words or clean_w != clean_words[-1].strip('(),'):
            clean_words.append(w)
    return ' '.join(clean_words)

# --- [ 카카오 로컬 API 기반 주소 변환 엔진 ] ---
def master_juso_converter(keyword, original_zip=""):
    if not keyword or pd.isna(keyword):
        return keyword
        
    kw_str = str(keyword).strip()
    
    # 1. 하이픈 동·호수 형식 변환 (예: 115-2104 -> 115동 2104호)
    kw_str_mod = re.sub(r'\b(\d+)-(\d{3,4})\b', r'\1동 \2호', kw_str)
    
    # 2. 상세 동/호수 및 괄호 정보 추출 패턴
    extra_pattern = r'(?:\b\d+동\s*\d+호?|\b\d+동|\b\d+호|\b[가나다라마바사아자차카타파하A-Za-z]\s*동\s*\d+호?|\d+층|\([^)]+\))'
    extra_details = re.findall(extra_pattern, kw_str_mod)
    
    # 3. API 검색을 위한 순수 검색어 추출
    search_q_str = kw_str_mod
    for ext in extra_details:
        search_q_str = search_q_str.replace(ext, '')
    search_q_str = ' '.join(search_q_str.split())

    for target_key, override_addr in SPECIAL_EXCEPTIONS.items():
        if target_key in kw_str:
            base_road_addr = override_addr
            break
    else:
        url = f"https://dapi.kakao.com/v2/local/search/address.json?query={urllib.parse.quote(search_q_str)}"
        headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
        
        base_road_addr = ""
        try:
            response = requests.get(url, headers=headers, timeout=5)
            if response.status_code == 200:
                res = response.json()
                documents = res.get('documents', [])
                if documents:
                    doc = documents[0]
                    road_info = doc.get('road_address')
                    if road_info:
                        base_road_addr = road_info.get('address_name')
                        building_name = road_info.get('building_name', '')
                        if building_name and building_name not in base_road_addr:
                            base_road_addr = f"{base_road_addr} {building_name}"
                    else:
                        base_road_addr = doc.get('address_name')
        except Exception:
            pass

    if not base_road_addr:
        return kw_str

    # 4. 상세 동·호수 및 괄호 정보 결합
    for detail in extra_details:
        detail_clean = detail.strip()
        if detail_clean and detail_clean not in base_road_addr:
            base_road_addr = f"{base_road_addr} {detail_clean}"

    return remove_duplicate_words(base_road_addr)


# --- [ Streamlit 웹 UI ] ---
st.set_page_config(page_title="스마트샵 주소 변환기", page_icon="🛍️", layout="centered")

st.title("🛍️ 스마트샵 주소 변환 & 엑셀 수정 (카카오 API 연동)")
st.write("엑셀 파일을 업로드하면 카카오 주소 API를 통해 도로명 주소 변환 및 상품명 정리가 이루어집니다.")

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
        st.success("🎉 주소 변환 및 상품명 정리가 완벽하게 완료되었습니다!")

    if st.session_state.excel_data is not None:
        st.download_button(
            label="📥 변환된 엑셀 파일 다운로드",
            data=st.session_state.excel_data,
            file_name="우편배송_변환결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
