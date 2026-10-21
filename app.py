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
    # 도로명 패턴 ('로 O번길 O' 또는 '길 O' 등)이 포함되어 있으면 신주소로 판단
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
    
    # 1. 이미 신주소(도로명 주소)인 경우 변환 없이 원본 띄어쓰기만 정돈하여 반환
    if is_new_road_address(kw_str_pre):
        return kw_str_pre

    # 2. 건물·아파트·빌리지·고시원·프라자 이름 추출
    bldg_match = re.search(r'([가-힣A-Za-z0-9]+(?:아파트|리버스트|빌리지|빌라|오피스텔|맨션|하이츠|타운|팰리스|주공|LH|단지|고시원|프라자|빌딩|상가))', kw_str_pre)
    orig_bldg = bldg_match.group(1) if bldg_match else ""

    # 3. 상세 동/호수, 층, 괄호 및 상호명 패턴 추출
    extra_pattern = r'(?:\b\d+\s*동\s*\d+\s*호?|\b\d+\s*동|\b\d+\s*호|\b[A-Za-z가나다라마바사아자차카타파하]\s*동\s*\d+\s*호?|\b[A-Za-z가나다라마바사아자차카타파하]\s*동|\d+층|\([^)]+\)|생활관\s*[가-힣A-Za-z0-9]+|\b[가-힣]+상가.*)'
    extra_details = re.findall(extra_pattern, kw_str_pre)
    
    # 4. API 검색용 순수 검색어 추출 (더미 '0' 등 제거)
    search_q_str = kw_str_pre
    for ext in extra_details:
        search_q_str = search_q_str.replace(ext, '')
    search_q_str = re.sub(r'\s+0\b', '', search_q_str)
    search_q_str = ' '.join(search_q_str.split())

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

    # 5. API 변환 실패 시(결과 없을 때) 원본 전처리 주소 그대로 유지
    if not base_road_addr:
        return kw_str_pre

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
