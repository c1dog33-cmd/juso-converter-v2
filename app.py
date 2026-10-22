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
    "율곡동 415": "경상북도 김천시 혁신8로 66",  # 율곡동 415번지 아파트 전용 도로명 매핑
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

# --- [ 후위 상세 정보(층수, 부서명, 메모 등) 사전 분리 함수 ] ---
def parse_trailing_info(addr_str):
    if not addr_str:
        return addr_str, ""
    words = addr_str.split()
    if len(words) <= 1:
        return addr_str, ""
    
    trailing_words = []
    for i in range(1, min(4, len(words) + 1)):
        w = words[-i]
        if any(term in w for term in ['층', '지층', '부', '팀', '문앞', '경비실', '택배함', '부재시', '비번', '총무']):
            trailing_words.insert(0, w)
        else:
            if trailing_words and (w in ['지하', '1', '2', '3', '4', '5', '상세', '사무실']):
                trailing_words.insert(0, w)
            else:
                break
                
    if trailing_words:
        trailing = " ".join(trailing_words)
        cleaned = " ".join(words[:-len(trailing_words)])
        return cleaned, trailing
        
    return addr_str, ""

# --- [ 주소 전처리 및 동·호수 포맷팅 함수 ] ---
def pre_process_address(addr_str):
    if not addr_str:
        return addr_str
    # 1. 슬래시 형태(105/2202)를 105동 2202호로 변환
    addr_str = re.sub(r'\b(\d+)\s*/\s*(\d+)\b', r'\1동 \2호', addr_str)
    
    # 2. 아파트/빌라명 뒤의 하이픈 동·호수 변환 (예: 대우아파트 304-505호 -> 304동 505호)
    def replace_bldg_hyphen(match):
        prefix = match.group(1)
        d = match.group(2)
        h = match.group(3)
        return f"{prefix} {d}동 {h}호"
    
    addr_str = re.sub(r'([가-힣A-Za-z0-9]+아파트|[가-힣A-Za-z0-9]+빌라|[가-힣A-Za-z0-9]+맨션)\s+(\d{1,4})-(\d{1,4})호?', replace_bldg_hyphen, addr_str)
    addr_str = re.sub(r'\b(\d{1,4})-(\d{1,4})호\b', r'\1동 \2호', addr_str)

    # 3. 건물 동 번호 포맷팅 (법정동 '창동' 등은 제외하고 숫자 동 및 한글 보조 동: 비동, 가동 등)
    addr_str = re.sub(r'\b(\d+\s*동|[가-하A-Za-z]\s*동)\s*(\d+)(?!호)\b', r'\1 \2호', addr_str)
    addr_str = re.sub(r'\b([가-하A-Za-z]동)(\d+호)\b', r'\1 \2', addr_str)
    
    # 4. 아파트 이름과 호수가 띄어쓰기 없이 붙어 있는 경우 포맷팅
    addr_str = re.sub(r'([가-힣A-Za-z0-9]+(?:아파트|스타힐스|캐슬|빌라|오피스텔|맨션|하이츠|타운|팰리스|뜨란채))\s*(\d+호)\b', r'\1 \2', addr_str)
    
    # 5. 하이픈 동·호수 변환 (법정동 직후에 붙는 진짜 지번 하이픈은 완벽히 보호)
    def replace_hyphen(match):
        full_match = match.group(0)
        start_idx = match.start()
        prefix = addr_str[:start_idx]
        if re.search(r'[가-힣]+(?:동|리|가)\s+$', prefix):
            return full_match  # 지번 하이픈 보호 (예: 사당동 190-152)
        return f"{match.group(1)}동 {match.group(2)}호"

    addr_str = re.sub(r'\b(\d{2,4})-(\d{3,4})\b', replace_hyphen, addr_str)
    
    # 6. 띄어쓰기 보정
    addr_str = re.sub(r'(\d+동)(\d+호)', r'\1 \2', addr_str)
    return addr_str

# --- [ 상품명 정제 함수 (와이드6 ➔ A13 매핑 포함) ] ---
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

# --- [ 마스터 주소 변환 엔진 ] ---
def master_juso_converter(keyword, original_zip=""):
    if not keyword or pd.isna(keyword):
        return keyword
        
    kw_str = str(keyword).strip()
    
    # 1. 후위 상세 정보(층수, 부서명 등) 사전 분리 추출 (Pre-extraction)
    kw_str_cleaned, trailing_info = parse_trailing_info(kw_str)
    
    # 2. 주소 전처리 적용
    kw_str_pre = pre_process_address(kw_str_cleaned)
    
    # 3. 회사명 / 법인명 안전 분리 추출
    corp_keywords = ['주식회사', '(주)', '유한회사', '합자회사']
    company_name = ""
    for kw in corp_keywords:
        if kw in kw_str_pre:
            idx = kw_str_pre.find(kw)
            prefix_part = kw_str_pre[:idx].strip()
            words = prefix_part.split()
            corp_words = []
            for w in reversed(words):
                if any(char.isdigit() for char in w) or '층' in w or '호' in w or '동' in w:
                    break
                corp_words.insert(0, w)
                if len(corp_words) >= 3:
                    break
            if corp_words:
                company_name = " ".join(corp_words) + " " + kw
            else:
                company_name = kw
            kw_str_pre = kw_str_pre.replace(company_name, "").strip()
            break

    # 4. 고객 배송 메시지/메모 사전 분리 추출
    memo_pattern = r'((?:검은색\s*)?(?:문앞|경비실|택배함|부재시|공동현관|현관|초인종|연락|전화|보관|비번|안전고리|문자)[^\(\)]*)'
    memo_matches = re.findall(memo_pattern, kw_str_pre)
    delivery_memo = ""
    if memo_matches:
        delivery_memo = max(memo_matches, key=len).strip()
        kw_str_pre = kw_str_pre.replace(delivery_memo, '').strip()

    dong_ho_pattern = r'(?:(?<!-)\b\d+\s*동\s*\d+\s*호?|(?<!-)\b\d+\s*동|\b\d+\s*호|\d+동\s*\d+호|\d+호|\b[A-Za-z가나다라마바사아자차카타파하]\s*동\s*\d+\s*호?|\b[A-Za-z가나다라마바사아자차카타파하]\s*동|\d+층)'

    # 5. 이미 신주소(도로명 주소)인 경우 처리
    if is_new_road_address(kw_str_pre):
        paren_match = re.search(r'(\([^)]+\))', kw_str_pre)
        paren_str = paren_match.group(1) if paren_match else ""
        if paren_str:
            kw_str_pre = kw_str_pre.replace(paren_str, '')

        bldg_match = re.search(r'([가-힣A-Za-z0-9]+(?:아파트|스타힐스|캐슬|리버스트|빌리지|빌라|오피스텔|맨션|하이츠|타운|팰리스|주공|LH|단지|고시원|프라자|빌딩|상가|타워|뜨란채))', kw_str_pre)
        orig_bldg = bldg_match.group(1) if bldg_match else ""
        if orig_bldg:
            kw_str_pre = re.sub(r'\b' + re.escape(orig_bldg) + r'\b', '', kw_str_pre)

        dong_ho_matches = re.findall(dong_ho_pattern, kw_str_pre)
        for dh in dong_ho_matches:
            kw_str_pre = kw_str_pre.replace(dh, '')

        base_region = re.sub(r'\s+0\b', '', kw_str_pre)
        base_region = ' '.join(base_region.split())

        parts = [base_region]
        if orig_bldg:
            parts.append(orig_bldg)
        if paren_str:
            parts.append(paren_str)
        if dong_ho_matches:
            parts.append(" ".join(dong_ho_matches))
        if company_name:
            parts.append(company_name)
        if delivery_memo:
            parts.append(delivery_memo)
        if trailing_info:
            parts.append(trailing_info)

        return pre_process_address(' '.join(parts))

    # 6. 지번 주소인 경우 (건물명 및 상세 정보 추출)
    bldg_match = re.search(r'([가-힣A-Za-z0-9]+(?:아파트|스타힐스|캐슬|리버스트|빌리지|빌라|오피스텔|맨션|하이츠|타운|팰리스|주공|LH|단지|고시원|프라자|빌딩|상가|타워|뜨란채))', kw_str_pre)
    orig_bldg = bldg_match.group(1) if bldg_match else ""

    extra_pattern = r'(?:(?<!-)\b\d+\s*동\s*\d+\s*호?|(?<!-)\b\d+\s*동|\b\d+\s*호|\d+동\s*\d+호|\d+호|\b[A-Za-z가나다라마바사아자차카타파하]\s*동\s*\d+\s*호?|\b[A-Za-z가나다라마바사아자차카타파하]\s*동|\d+층|\([^)]+\)|생활관\s*[가-힣A-Za-z0-9]+|\b[가-힣]+상가.*)'
    extra_details = re.findall(extra_pattern, kw_str_pre)
    
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
            if "상가" in kw_str:
                base_road_addr = kw_str_pre # 상가인 경우 기존 유지
            else:
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

    if not base_road_addr:
        base_road_addr = kw_str_pre

    if orig_bldg and orig_bldg not in base_road_addr:
        base_road_addr = f"{base_road_addr} {orig_bldg}"

    for detail in extra_details:
        detail_clean = pre_process_address(detail.strip())
        if detail_clean.startswith('(') and detail_clean.endswith(')'):
            if any(sub.strip() in base_road_addr for sub in detail_clean.strip('()').split(',')):
                continue
        if detail_clean and detail_clean not in base_road_addr:
            base_road_addr = f"{base_road_addr} {detail_clean}"

    if company_name and company_name not in base_road_addr:
        base_road_addr = f"{base_road_addr} {company_name}"

    if delivery_memo and delivery_memo not in base_road_addr:
        base_road_addr = f"{base_road_addr} {delivery_memo}"
        
    if trailing_info and trailing_info not in base_road_addr:
        base_road_addr = f"{base_road_addr} {trailing_info}"

    words = base_road_addr.split()
    clean_words = []
    for w in words:
        clean_w = w.strip('(),')
        if not clean_words or clean_w != clean_words[-1].strip('(),'):
            clean_words.append(w)
            
    return pre_process_address(' '.join(clean_words))
