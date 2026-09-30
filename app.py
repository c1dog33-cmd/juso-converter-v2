import io
import re
import urllib.parse
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font

CONFIRM_KEY = "devU01TX0FVVEgyMDI2MDkzMDEwMTcwMDEyMDUzMjM="

def fix_zipcode(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).split('.')[0].strip()
    return val_str.zfill(5) if val_str else ""

# --- [ 특정 예외 주소 강제 매핑 사전 ] ---
SPECIAL_EXCEPTIONS = {
    "불로동 268-2": "인천광역시 검단구 금정로 12",
}

# --- [ 정제 및 텍스트 교정 함수 ] ---
def remove_duplicate_words(addr_str):
    if not addr_str:
        return addr_str
    
    # 행정구역 띄어쓰기 교정 (예: 남동 구 -> 남동구, 서 구 -> 서구)
    addr_str = re.sub(r'남동\s+구', '남동구', addr_str)
    addr_str = re.sub(r'서\s+구', '서구', addr_str)
    addr_str = re.sub(r'간\s+석동', '간석동', addr_str)
    addr_str = re.sub(r'가\s+능동', '가능동', addr_str)

    # 슬래시 및 하이픈 동/호수 교정
    addr_str = re.sub(r'(\d+)\s*/\s*(\d+)', r'\1동 \2호', addr_str)
    addr_str = re.sub(r'\b(\d{1,4})\s*-\s*(\d{3,4})호?\b', r'\1동 \2호', addr_str)
    
    # 동/호 띄어쓰기 교정 (예: 111 동 1404 -> 111동 1404, 111 동 -> 111동)
    addr_str = re.sub(r'\b(\d+)\s+동\s*(\d+)\b', r'\1동 \2', addr_str)
    addr_str = re.sub(r'\b(\d+)\s+동\b', r'\1동', addr_str)
    addr_str = re.sub(r'\b(\d+)\s+호\b', r'\1호', addr_str)
    
    # 동 바로 뒤에 숫자가 붙어 있는 경우 한 칸 띄우기 (예: 101동1406호 -> 101동 1406호)
    addr_str = re.sub(r'동(\d)', r'동 \1', addr_str)

    # 아파트 동 번호 뒤에 숫자가 있고 '호'가 없는 경우 '호' 표기 추가 (예: 111동 1404 -> 111동 1404호)
    addr_str = re.sub(r'\b(\d+동)\s+(\d+)(?!호)\b', r'\1 \2호', addr_str)

    words = addr_str.split()
    clean_words = []
    for w in words:
        clean_w = w.strip('(),')
        if not clean_words or clean_w != clean_words[-1].strip('(),'):
            clean_words.append(w)
            
    return ' '.join(clean_words)

# --- [ 만능 주소 변환 엔진 ] ---
def master_juso_converter(keyword):
    if not keyword or pd.isna(keyword):
        return keyword
        
    kw_str = str(keyword).strip()
    
    # 0. 행정구역 및 슬래시/하이픈, 동/호수 띄어쓰기 사전 전처리
    kw_str = re.sub(r'남동\s+구', '남동구', kw_str)
    kw_str = re.sub(r'서\s+구', '서구', kw_str)
    kw_str = re.sub(r'(\d+)\s*/\s*(\d+)', r'\1동 \2호', kw_str)
    kw_str = re.sub(r'\b(\d{1,4})\s*-\s*(\d{3,4})호?\b', r'\1동 \2호', kw_str)
    
    # 동/호 띄어쓰기 교정 (예: 111 동 1404 -> 111동 1404, 111 동 -> 111동)
    kw_str = re.sub(r'\b(\d+)\s+동\s*(\d+)\b', r'\1동 \2', kw_str)
    kw_str = re.sub(r'\b(\d+)\s+동\b', r'\1동', kw_str)
    kw_str = re.sub(r'\b(\d+)\s+호\b', r'\1호', kw_str)
    
    # 동 바로 뒤에 숫자가 붙어 있는 경우 한 칸 띄우기 (예: 101동1406호 -> 101동 1406호)
    kw_str = re.sub(r'동(\d)', r'동 \1', kw_str)

    # 아파트 동 번호 뒤에 숫자가 있고 '호'가 없는 경우 '호' 표기 추가 (예: 111동 1404 -> 111동 1404호)
    kw_str = re.sub(r'\b(\d+동)\s+(\d+)(?!호)\b', r'\1 \2호', kw_str)
    
    # [규칙 1] 특정 예외 매핑 체크 (예: 불로동 268-2)
    for target_key, override_addr in SPECIAL_EXCEPTIONS.items():
        if target_key in kw_str:
            extra_part = kw_str
            for part in target_key.split():
                extra_part = extra_part.replace(part, '')
            extra_part = re.sub(r'인천광역시|검단구|서구|불로동', '', extra_part).strip()
            return remove_duplicate_words(f"{override_addr} {extra_part}")

    # [규칙 2] 인천 서구 불로동 -> 검단구 불로동 강제 매핑
    if '불로동' in kw_str:
        kw_str = kw_str.replace('서구', '검단구').replace('서해구', '검단구')
        if '인천광역시 검단구' not in kw_str and '인천 검단구' not in kw_str:
            kw_str = re.sub(r'인천광역시\s+서구', '인천광역시 검단구', kw_str)
            kw_str = re.sub(r'인천\s+서구', '인천 검단구', kw_str)

    # 1. 상세 부가정보(동/호수, 괄호 내용, 병원/기관명 등) 추출 및 원본에서 분리 (불필요한 뒷자리 임의 절삭 패턴 제거)
    extra_pattern = r'(?:\b\d+동\s*\d+호?|\b[가A-Za-z]\s*동\s*\d+호?|\b[가A-Za-z]\s*동\d+|\b[가A-Za-z]+동\d+|\d+호|\d+층|B\d+호|관리실|택배보관함|물리치료실|\([^)]+\)|[가-힣]+(?:의원|병원|한의원|이비인후과|내과|외과|치과|소아과|센터))'
    extra_details = re.findall(extra_pattern, kw_str)
    
    # 검색용 쿼리 생성 시 상세 부가정보 일시 제거
    search_q_str = re.sub(extra_pattern, '', kw_str)
    search_q_str = ' '.join(search_q_str.split())

    # 2. 건물명 뒤 '103-401' 형태를 '103동 104호'로 자동 변환
    tokens_init = search_q_str.split()
    processed_tokens = []
    for i, t in enumerate(tokens_init):
        if re.match(r'^\d+-\d+$', t):
            prev_token = tokens_init[i-1] if i > 0 else ""
            if prev_token and not any(prev_token.endswith(s) for s in ['동', '리', '가', '로', '길']):
                parts = t.split('-')
                processed_tokens.append(f"{parts[0]}동 {parts[1]}호")
            else:
                processed_tokens.append(t)
        else:
            processed_tokens.append(t)
    search_q_str = " ".join(processed_tokens)
    
    # 3. 특수 예외 처리 (월산동 등)
    if '월산동 986-3' in search_q_str or '월산동 986' in search_q_str:
        extra = search_q_str.replace('광주광역시', '').replace('전남광주통합특별시', '').replace('남구', '').replace('월산동', '').replace('986-3', '').replace('986', '').strip()
        return remove_duplicate_words(f"광주광역시 남구 대남대로 363 {extra} {' '.join(extra_details)}".strip())

    # 4. 스마트 토큰 분리
    base_tokens = search_q_str.split()
    sido_sigungu_dong_tokens = []
    jibeon_token = ""
    building_tokens = []
    
    for t in base_tokens:
        if re.match(r'^\d+(-\d+)?$', t) or re.match(r'^산\d+(-\d+)?$', t):
            jibeon_token = t
        elif any(t.endswith(s) for s in ['도', '시', '구', '군', '동', '리', '가', '로', '길']) and t != '시':
            if not jibeon_token:
                sido_sigungu_dong_tokens.append(t)
            else:
                building_tokens.append(t)
        else:
            building_tokens.append(t)

    sido_sigungu_dong = " ".join(sido_sigungu_dong_tokens)
    building_name_candidate = " ".join(building_tokens)
    
    # 5. 다단계 검색 후보군 생성
    query_candidates = []
    
    if '불로동' in search_q_str:
        if sido_sigungu_dong and jibeon_token:
            query_candidates.append(f"인천광역시 검단구 불로동 {jibeon_token}")
        query_candidates.append(search_q_str.replace('서구', '검단구').replace('서해구', '검단구'))

    elif '서구' in search_q_str:
        seohae_q = search_q_str.replace('서구', '서해구')
        if sido_sigungu_dong and jibeon_token:
            seohae_dong = sido_sigungu_dong.replace('서구', '서해구')
            query_candidates.append(f"{seohae_dong} {jibeon_token}")
        query_candidates.append(seohae_q)

    if sido_sigungu_dong and jibeon_token:
        query_candidates.append(f"{sido_sigungu_dong} {jibeon_token}")
        
    if search_q_str not in query_candidates:
        query_candidates.append(search_q_str)
        
    if sido_sigungu_dong and building_name_candidate:
        query_candidates.append(f"{sido_sigungu_dong} {building_name_candidate}")

    base_road_addr = ""
    api_bd_nm = ""
    is_user_sangga = '상가' in kw_str
    
    for q in query_candidates:
        if not q.strip():
            continue
        url = f"https://business.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword={urllib.parse.quote(q)}&confmKey={CONFIRM_KEY}&resultType=json"
        
        try:
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                res = response.json()
                juso_list = res.get('results', {}).get('juso') or []
                
                if juso_list:
                    selected_juso = None
                    
                    for juso in juso_list:
                        bd_name = juso.get('bdNm', '').strip()
                        if not is_user_sangga and '상가' in bd_name:
                            continue
                        if bd_name and building_name_candidate and (bd_name in building_name_candidate or building_name_candidate in bd_name):
                            selected_juso = juso
                            break

                    if not selected_juso and not is_user_sangga:
                        for juso in juso_list:
                            if '상가' not in juso.get('bdNm', ''):
                                selected_juso = juso
                                break

                    if not selected_juso:
                        selected_juso = juso_list[0]

                    base_road_addr = selected_juso.get('roadAddr')
                    api_bd_nm = selected_juso.get('bdNm', '')

                    if base_road_addr:
                        break
        except Exception:
            continue

    if not base_road_addr:
        return remove_duplicate_words(kw_str)

    # 6. 아파트/건물명 보완 결합
    target_bd = api_bd_nm.strip() if api_bd_nm else building_name_candidate.strip()
    if target_bd and target_bd not in base_road_addr:
        base_road_addr = f"{base_road_addr} {target_bd}"

    # 7. 최종 결과 조합: 도로명 주소 맨 뒤에 중복되지 않는 상세 부가정보(extra_details) 배치
    full_result = base_road_addr
    if extra_details:
        needed_details = []
        for p in extra_details:
            p_clean = re.sub(r'[\s(),]', '', p)
            base_clean = re.sub(r'[\s(),]', '', base_road_addr)
            p_norm = p_clean.replace('LH', '엘에이치')
            base_norm = base_clean.replace('LH', '엘에이치')

            if p_clean and (p_clean in base_clean or p_norm in base_norm):
                continue

            # API 결과에 이미 포함된 아파트/건물명과 중복되는 괄호 항목 필터링
            is_redundant_building_paren = False
            if p.startswith('(') and p.endswith(')'):
                inner = p[1:-1]
                inner_norm = re.sub(r'[\s동시구군읍면리아파트빌딩단지]', '', inner).replace('LH', '엘에이치')
                base_inner_norm = re.sub(r'[\s동시구군읍면리아파트빌딩단지]', '', base_road_addr).replace('LH', '엘에이치')
                if inner_norm and inner_norm in base_inner_norm:
                    is_redundant_building_paren = True
                else:
                    for token in inner.split(','):
                        token_clean = token.strip().replace('아파트', '').replace('빌딩', '').replace('단지', '').replace('동', '')
                        token_norm = token_clean.replace('LH', '엘에이치')
                        if len(token_norm) >= 2 and token_norm in base_road_addr.replace('LH', '엘에이치'):
                            is_redundant_building_paren = True
                            break

            if is_redundant_building_paren:
                continue

            if p not in base_road_addr:
                needed_details.append(p)

        if needed_details:
            full_result = f"{base_road_addr} {' '.join(needed_details)}"

    return remove_duplicate_words(full_result)


# --- [ Streamlit 웹 UI ] ---
st.set_page_config(page_title="자동 주소 변환기", page_icon="🚚", layout="centered")

st.title("🚚 만능 주소 변환 & 엑셀 정제 웹 앱")
st.write("엑셀 파일을 업로드하면 도로명 주소 변환, 우편번호 0 보존, 엑셀 서식을 자동으로 적용해 줍니다.")

uploaded_file = st.file_uploader("변환할 엑셀 파일(.xlsx, .xls)을 업로드하세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("### 📄 업로드 데이터 미리보기 (상위 5건)")
    st.dataframe(df.head())

    if st.button("🚀 주소 변환 및 서식 적용 시작"):
        with st.spinner("주소를 변환하는 중입니다... 데이터 양에 따라 시간이 걸릴 수 있습니다."):
            if '우편번호' in df.columns:
                df['우편번호'] = df['우편번호'].apply(fix_zipcode)

            if '배송지' in df.columns:
                progress_bar = st.progress(0)
                total = len(df)
                
                converted_addrs = []
                for i, row in df.iterrows():
                    res = master_juso_converter(row['배송지'])
                    converted_addrs.append(res)
                    progress_bar.progress((i + 1) / total)
                    
                df['배송지'] = converted_addrs
            else:
                st.warning("[주의] '배송지' 컬럼을 찾을 수 없습니다.")

            target_columns = ['수취인명', '전화', '우편번호', '배송지', '선택정보', '기타', '구분']
            if len(df.columns) == len(target_columns):
                df.columns = target_columns

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                
                font_size_8 = Font(size=8)
                for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=len(df.columns)):
                    worksheet.row_dimensions[row[0].row].height = 18
                    for cell in row:
                        cell.font = font_size_8
                        
                if '우편번호' in df.columns:
                    zip_col_idx = df.columns.get_loc('우편번호') + 1
                    for row in range(2, worksheet.max_row + 1):
                        worksheet.cell(row=row, column=zip_col_idx).number_format = '@'

            excel_data = output.getvalue()

        st.success("🎉 변환이 완벽하게 완료되었습니다!")
        
        st.download_button(
            label="📥 변환된 엑셀 파일 다운로드",
            data=excel_data,
            file_name="우편배송_변환결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
