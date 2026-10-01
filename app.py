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
    
    # 0-1. 렌즈보호커버 상품 정제 규칙
    if '렌즈' in s or 'Lens Guards' in s or '렌즈가드' in s:
        qty = 1
        if '☞' in s:
            parts = s.split('☞')
            after = parts[1].strip() if len(parts) > 1 else ""
            qty_match = re.match(r'^(\d+)\s*(EA|개)?', after, re.IGNORECASE)
            if qty_match:
                qty = int(qty_match.group(1))
        if qty > 1:
            return f"렌즈보호커버[{qty}]"
        else:
            return "렌즈보호커버"

    # 0-2. 태블릿/송풍구 거치대 상품 처리
    if '송풍구' in s or '거치대' in s or '태블릿' in s:
        return "태블릿 송풍구 거치대"
    
    # 1. 케이블 상품 정제 규칙 (m 및 cm 단위 모두 인식)
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
        m_match = re.search(r'(\d+(?:\.\d+)?(?:m|cm))', s, re.IGNORECASE)
        length_str = m_match.group(1) if m_match else ""
        
        components = []
        if has_g:
            components.append('ㄱ자형')
        if has_c_type:
            components.append('C타입 케이블')
        else:
            components.append('케이블')
        if length_str:
            components.append(length_str)
            
        target = " ".join(components)
        if qty > 1:
            return f"{target}[{qty}]"
        else:
            return target

    # 2. 케이스 상품 정제 규칙 ('케이스' 글자가 포함된 경우)
    if '케이스' not in s:
        return s
    
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
        
    # '모델선택:'이 포함된 경우 기준 단어 포함하여 앞부분 모두 삭제
    if '모델선택:' in base:
        parts = base.split('모델선택:')
        target = parts[-1].strip()
        if '/' in target:
            target = target.split('/')[0].strip()
        if ',' in target:
            target = target.split(',')[0].strip()
    else:
        # 마지막 슬래시(/ 또는 //) 앞의 상품명은 전부 지우고 뒤쪽 내용만 추출
        sub_parts = re.split(r'[/]{1,2}', base)
        target = sub_parts[-1].strip() if len(sub_parts) > 1 else base
        
    # 괄호 내용(우편함배송 등) 제거
    target = re.sub(r'\([^)]*\)', '', target).strip()
    
    # 특정 모델명 간소화 규칙 적용
    if 'S10플러스' in target or '975' in target:
        target = 'S10플러스'
    elif '와이드6' in target and 'A13' in target:
        target = 'A13'
    elif 'S10' in target and '5G' in target:
        target = 'S10 5G'
    elif 'S10' in target and ('973' in target or '기본' in target):
        target = 'S10'
    
    # 요청하신 단어들 삭제 ('케이스', '풀액정', '2장', '클리어', '투명', '갤럭시')
    target = re.sub(r'케이스', '', target)
    target = re.sub(r'풀액정', '', target)
    target = re.sub(r'2장', '', target)
    target = re.sub(r'클리어', '', target)
    target = re.sub(r'투명', '', target)
    target = re.sub(r'갤럭시', '', target)
    target = ' '.join(target.split())
    
    # 수량이 2개 이상일 때 모델명 뒤에 [수량] 표기 (예: 노트20[2])
    if qty > 1:
        result_str = f"{target}[{qty}]"
    else:
        result_str = target
        
    # '택배배송' 또는 '택배' 글자가 포함된 경우 맨 뒤에 ' 택배' 표기 추가
    if '택배' in s:
        result_str = f"{result_str} 택배"
        
    return result_str

# --- [ 정제 및 텍스트 교정 함수 (위험한 하이픈 변환 규칙 제거 완료) ] ---
def remove_duplicate_words(addr_str):
    if not addr_str:
        return addr_str
    
    # 행정구역 띄어쓰기 교정
    addr_str = re.sub(r'남동\s+구', '남동구', addr_str)
    addr_str = re.sub(r'서\s+구', '서구', addr_str)
    addr_str = re.sub(r'간\s+석동', '간석동', addr_str)
    addr_str = re.sub(r'가\s+능동', '가능동', addr_str)

    # 슬래시 교정 (슬래시는 동/호수로 변환)
    addr_str = re.sub(r'(\d+)\s*/\s*(\d+)', r'\1동 \2호', addr_str)
    
    # 동/호 띄어쓰기 교정
    addr_str = re.sub(r'\b(\d+)\s+동\s*(\d+)\b', r'\1동 \2', addr_str)
    addr_str = re.sub(r'\b(\d+)\s+동\b', r'\1동', addr_str)
    addr_str = re.sub(r'\b(\d+)\s+호\b', r'\1호', addr_str)
    
    # 동 바로 뒤에 숫자가 붙어 있는 경우 한 칸 띄우기
    addr_str = re.sub(r'동(\d)', r'동 \1', addr_str)

    # 아파트 동 번호 뒤에 숫자가 있고 '호'가 없는 경우 '호' 표기 추가
    addr_str = re.sub(r'\b(\d+동)\s+(\d+)(?!호)\b', r'\1 \2호', addr_str)

    # 가동, 나동, 라동 등 한글 동 이름 뒤에 숫자가 있고 '호'가 없는 경우 '호' 자동 추가
    addr_str = re.sub(r'\b([가-힣A-Za-z]+동)\s*(\d+)(?!호)\b', r'\1 \2호', addr_str)

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
    
    # 0. 행정구역 띄어쓰기 사전 전처리
    kw_str = re.sub(r'남동\s+구', '남동구', kw_str)
    kw_str = re.sub(r'서\s+구', '서구', kw_str)
    
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

    # 1. 상세 부가정보(동/호수, 괄호 내용, 병원/기관명 등) 추출 및 원본에서 분리
    extra_pattern = r'(?:\b\d+동\s*\d+호?|\b[가-힣A-Za-z]+\s*동\s*\d+호?|\b[가-힣A-Za-z]+동\d+|\d+호|\d+층|B\d+호|관리실|택배보관함|물리치료실|\([^)]+\)|[가-힣]+(?:의원|병원|한의원|이비인후과|내과|외과|치과|소아과|센터))'
    extra_details = re.findall(extra_pattern, kw_str)
    
    # 검색용 쿼리 생성 시 상세 부가정보 일시 제거
    search_q_str = re.sub(extra_pattern, '', kw_str)
    search_q_str = ' '.join(search_q_str.split())

    # 2. 건물명 뒤에 있는 하이픈 형태(예: '103-401')를 '103동 104호'로 안전하게 변환 (앞의 진짜 지번 '148-1'은 보호됨)
    tokens_init = search_q_str.split()
    processed_tokens = []
    for i, t in enumerate(tokens_init):
        if re.match(r'^\d+-\d+$', t):
            prev_token = tokens_init[i-1] if i > 0 else ""
            if prev_token and not any(prev_token.endswith(s) for s in ['동', '리', '가', '로', '길', '시', '구', '군', '읍', '면']):
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

    # 4. 스마트 토큰 분리 (읍, 면 단위를 주소 구역으로 올바르게 인식하도록 추가)
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
    building_name_candidate = " ".join(building_tokens)
    
    # 5. 다단계 검색 후보군 생성 (지번 주소 검색을 최우선으로 배치)
    query_candidates = []
    
    if sido_sigungu_dong and jibeon_token:
        query_candidates.append(f"{sido_sigungu_dong} {jibeon_token}")

    if '불로동' in search_q_str:
        query_candidates.append(search_q_str.replace('서구', '검단구').replace('서해구', '검단구'))

    elif '서구' in search_q_str:
        seohae_q = search_q_str.replace('서구', '서해구')
        query_candidates.append(seohae_q)

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

    # 6. API 결과 건물명 결합
    api_bd = api_bd_nm.strip() if api_bd_nm else ""
    if api_bd and api_bd not in base_road_addr:
        base_road_addr = f"{base_road_addr} {api_bd}"

    # 7. 사용자가 입력한 맨 끝 상호명/이름(building_name_candidate)이 지워지지 않도록 맨 끝에 안전하게 보존
    user_bd = building_name_candidate.strip()
    if user_bd and user_bd not in base_road_addr:
        if user_bd not in extra_details:
            extra_details.append(user_bd)

    # 8. 최종 결과 조합: 동/건물명은 앞으로, 호수/층은 반드시 뒤로 가도록 정렬
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
            building_parts = []
            unit_parts = []
            for p in needed_details:
                # '호'나 '층'이 포함되어 있으면서 '동'이 포함되지 않은 경우만 호수(unit)로 분류
                if re.search(r'(\d+호|\d+층|B\d+호)', p) and not re.search(r'\d+동', p):
                    unit_parts.append(p)
                else:
                    building_parts.append(p)
            sorted_details = building_parts + unit_parts
            full_result = f"{base_road_addr} {' '.join(sorted_details)}"

    return remove_duplicate_words(full_result)


# --- [ Streamlit 웹 UI ] ---
st.set_page_config(page_title="스마트샵 주소 변환기", page_icon="🛍️", layout="centered")

st.title("🛍️ 스마트샵 주소 변환 & 엑셀 수정")
st.write("엑셀 파일을 업로드하면 도로명 주소 변환, 우편번호 0 보존, 상품명 모델명 자동 정리, 엑셀 서식을 자동으로 적용해 줍니다.")

uploaded_file = st.file_uploader("변환할 엑셀 파일(.xlsx, .xls)을 업로드하세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("### 📄 업로드 데이터 미리보기 (상위 5건)")
    st.dataframe(df.head())

    if st.button("🚀 주소 변환 및 서식 적용 시작"):
        with st.spinner("주소를 변환하고 상품명을 정리하는 중입니다... 데이터 양에 따라 시간이 걸릴 수 있습니다."):
            if '우편번호' in df.columns:
                df['우편번호'] = df['우편번호'].apply(fix_zipcode)

            if '선택정보' in df.columns:
                df['선택정보'] = df['선택정보'].apply(clean_product_name)

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

            # 엑셀 저장 시 openpyxl 제어 문자 에러(IllegalCharacterError) 방지 정제 적용
            for col in df.columns:
                df[col] = df[col].apply(remove_illegal_chars)

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

        st.success("🎉 주소 변환 및 상품명 정리가 완벽하게 완료되었습니다!")
        
        st.download_button(
            label="📥 변환된 엑셀 파일 다운로드",
            data=excel_data,
            file_name="우편배송_변환결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
