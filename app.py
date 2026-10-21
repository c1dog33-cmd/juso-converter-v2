import io
import re
import time
import urllib.parse
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font

CONFIRM_KEY = "U01TX0FVVEgyMDI2MTAyMTEwMjkzNzEyMDc5MDA="

def fix_zipcode(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).split('.')[0].strip()
    return val_str.zfill(5) if val_str else ""

SPECIAL_EXCEPTIONS = {
    "불로동 268-2": "인천광역시 검단구 금정로 12 (불로동, 신검단중앙역풍경채어바니티)",
    "남산타운": "서울특별시 중구 다산로 32 (신당동, 남산타운)",
}

def remove_illegal_chars(val):
    if isinstance(val, str):
        return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', val)
    return val

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

    if '와이드5' in s:
        s = s.replace('와이드5', 'A22')
    return s

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

# --- [ 주소 변환 엔진 (서버 차단 방지 딜레이 포함) ] ---
def master_juso_converter(keyword, original_zip=""):
    if not keyword or pd.isna(keyword):
        return keyword
        
    kw_str = str(keyword).strip()
    
    delivery_note_pattern = r'([가-힣\s]+(?:부탁드립니다?|놓아주세요?|전해주세요?|맡겨주세요?|보관해주세요?|부탁해요?))'
    delivery_notes = re.findall(delivery_note_pattern, kw_str)
    for note in delivery_notes:
        kw_str = kw_str.replace(note, '').strip()
    kw_str = re.sub(r'\s*\.\s*', ' ', kw_str).strip()

    for target_key, override_addr in SPECIAL_EXCEPTIONS.items():
        if target_key in kw_str:
            return override_addr

    extra_pattern = r'(?:\b\d+동\s*\d+호?|\b[가나다라마바사아자차카타파하A-Za-z]\s*동\s*\d+호?|\d+호|\d+층|\([^)]+\))'
    extra_details = re.findall(extra_pattern, kw_str)
    
    search_q_str = re.sub(extra_pattern, '', kw_str)
    search_q_str = ' '.join(search_q_str.split())

    base_tokens = search_q_str.split()
    sido_sigungu_dong_tokens = []
    jibeon_token = ""
    building_tokens = []
    
    for t in base_tokens:
        if re.match(r'^\d+(-\d+)?$', t) or re.match(r'^산\d+(-\d+)?$', t):
            jibeon_token = t
        elif any(t.endswith(s) for s in ['도', '시', '구', '군', '읍', '면', '동', '리', '가', '로', '길']) and t != '시':
            if not jibeon_token:
                sido_sigungu_dong_tokens.append(t)
            else:
                building_tokens.append(t)
        else:
            building_tokens.append(t)

    sido_sigungu_dong = " ".join(sido_sigungu_dong_tokens)
    
    query_candidates = []
    if sido_sigungu_dong and jibeon_token:
        query_candidates.append(f"{sido_sigungu_dong} {jibeon_token}")
    query_candidates.append(search_q_str)

    base_road_addr = ""
    api_bd_nm = ""
    
    for q in query_candidates:
        if not q.strip():
            continue
        url = f"https://business.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword={urllib.parse.quote(q)}&confmKey={CONFIRM_KEY}&resultType=json"
        
        try:
            # 🚀 [서버 차단 방지] 각 요청 사이에 0.3초 대기 시간 부여
            time.sleep(0.3)
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                res = response.json()
                juso_list = res.get('results', {}).get('juso') or []
                if juso_list:
                    base_road_addr = juso_list[0].get('roadAddr')
                    api_bd_nm = juso_list[0].get('bdNm', '')
                    if base_road_addr:
                        break
        except Exception:
            continue

    if not base_road_addr:
        return kw_str

    if api_bd_nm and api_bd_nm not in base_road_addr:
        base_road_addr = f"{base_road_addr} {api_bd_nm}"

    if delivery_notes:
        cleaned_notes = " ".join([n.strip() for n in delivery_notes if n.strip()])
        base_road_addr = f"{base_road_addr} ({cleaned_notes})"

    return remove_duplicate_words(base_road_addr)

# --- [ Streamlit 웹 UI ] ---
st.set_page_config(page_title="스마트샵 주소 변환기", page_icon="🛍️", layout="centered")

st.title("🛍️ 스마트샵 주소 변환 & 엑셀 수정")
uploaded_file = st.file_uploader("변환할 엑셀 파일(.xlsx, .xls)을 업로드하세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("### 📄 업로드 데이터 미리보기 (상위 5건)")
    st.dataframe(df.head())

    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None

    if st.button("🚀 주소 변환 및 서식 적용 시작"):
        with st.spinner("주소를 변환하는 중입니다... (서버 차단 방지 딜레이 적용)"):
            if '배송지' in df.columns:
                converted_addrs = []
                for idx, row in df.iterrows():
                    res = master_juso_converter(row['배송지'])
                    converted_addrs.append(res)
                df['배송지'] = converted_addrs

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            st.session_state.excel_data = output.getvalue()
        st.success("완료되었습니다!")

    if st.session_state.excel_data is not None:
        st.download_button(
            label="📥 변환된 엑셀 파일 다운로드",
            data=st.session_state.excel_data,
            file_name="우편배송_변환결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
