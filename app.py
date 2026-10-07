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

# --- [ 상품명 정제 함수 (노트20울트라 우선순위 및 케이블 길이 후위 탐지 반영) ] ---
def clean_product_name(val):
    if pd.isna(val) or not str(val).strip():
        return val
    s = str(val).strip()
    
    has_film = False
    if '추가액정필름' in s and '필요없음' not in s and ('풀액정' in s or '액정필름' in s):
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

    # 🚀 [우선 탐지] 슬래시('/') 뒤쪽(옵션 영역) 스캔 (노트20울트라 등 긴 기종 최우선 배치)
    if '/' in s:
        slash_parts = s.split('/')
        for p in reversed(slash_parts):
            m = re.search(r'(S\s*10\s*5G|노트\s*20\s*울트라|노트\s*20\s*플러스|노트\s*10\s*플러스|노트\s*\d+\s*울트라|노트\s*\d+\s*플러스|노트\s*\d+|S\s*\d+(?:\s*플러스|\s*울트라|\s*5G)?|A\s*\d+|Z\s*(?:플립|폴드)\s*\d*|와이드\s*\d+|N\s*\d+|V\s*50S|V\s*50)', p, re.IGNORECASE)
            if m:
                raw_model = m.group(1).replace(" ", "").upper()
                if raw_model == 'S105G': target_model = 'S10 5G'
                elif raw_model == 'S20플러스' or raw_model == 'S20+': target_model = 'S20플러스'
                elif raw_model == '노트10플러스': target_model = '노트10플러스'
                elif raw_model == '노트20울트라': target_model = '노트20울트라'
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
        
        res = f"{opt_val}[{qty}]" if qty > 1 else opt_val
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
        if '택배' in s: res += " 택배"
        return res

    # 🚀 [범용 스마트 탐지] 노트20울트라 등 긴 모델명 최우선 배치
    smartphone_match = re.search(r'(S\s*10\s*5G|노트\s*20\s*울트라|노트\s*20\s*플러스|노트\s*10\s*플러스|노트\s*\d+\s*울트라|노트\s*\d+\s*플러스|노트\s*\d+|S\s*\d+(?:\s*플러스|\s*울트라|\s*5G)?|A\s*\d+|Z\s*(?:플립|폴드)\s*\d*|와이드\s*\d+|N\s*\d+|V\s*50S|V\s*50|V\s*\d+|벨벳|윙)', s, re.IGNORECASE)
    if smartphone_match:
        raw_model = smartphone_match.group(1).replace(" ", "").upper()
        if raw_model == 'S105G': target_model = 'S10 5G'
        elif raw_model in ['N960', 'SM-N960']: target_model = '노트9'
        elif raw_model == 'V50S': target_model = 'V50S'
        elif raw_model == 'V50': target_model = 'V50'
        elif raw_model == '노트20울트라': target_model = '노트20울트라'
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
        
        last_segment = s.split('/')[-1] if '/' in s else s
        m_match = re.search(r'(\d+(?:\.\d+)?(?:m|cm))', last_segment, re.IGNORECASE)
        if not m_match:
            m_match = re.search(r'(\d+(?:\.\d+)?(?:m|cm))', s, re.IGNORECASE)
            
        length_str = m_match.group(1) if m_match else ""
        
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
    
    result_str = f"{target}[{qty}]" if qty > 1 else target
    if '택배' in s: result_str = f"{result_str} 택배"
    if has_film: result_str = f"{result_str} +필름"
        
    return result_str

# --- [ 정제 및 텍스트 교정 함수 ] ---
def remove_duplicate_words(addr_str):
    if not addr_str:
        return addr_str
    
    addr_str = re.sub(r'남동\s+구', '남동구', addr_str)
    addr_str = re.sub(r'서\s+구', '서구', addr_str)
    addr_str = re.sub(r'간\s+석동', '간석동', addr_str)
    addr_str = re.sub(r'가\s+능동', '가능동', addr_str)

    addr_str = re.sub(r'(\d+)\s*/\s*(\d+)', r'\1동 \2호', addr_str)
    addr_str = re.sub(r'\b(\d+)\s+동\s*(\d+)\b', r'\1동 \2', addr_str)
    addr_str = re.sub(r'\b(\d+)\s+동\b', r'\1동', addr_str)
    addr_str = re.sub(r'\b(\d+)\s+호\b', r'\1호', addr_str)
    addr_str = re.sub(r'동(\d)', r'동 \1', addr_str)
    addr_str = re.sub(r'\b(\d+동)\s+(\d+)(?!호)\b', r'\1 \2호', addr_str)
    addr_str = re.sub(r'\b([가나다라마바사아자차카타파하A-Za-z]동)\s*(\d+)(?!호)\b', r'\1 \2호', addr_str)

    words = addr_str.split()
    clean_words = []
    for w in words:
        clean_w = w.strip('(),')
        if not clean_words or clean_w != clean_words[-1].strip('(),'):
            clean_words.append(w)
            
    return ' '.join(clean_words)

# --- [ 만능 주소 변환 엔진 (우편번호 검증 안전장치 포함) ] ---
def master_juso_converter(keyword, original_zip=""):
    if not keyword or pd.isna(keyword):
        return keyword
        
    kw_str = str(keyword).strip()
    
    kw_str = re.sub(r'(\d+)\s+동\b', r'\1동', kw_str)
    kw_str = re.sub(r'(\d+)\s+호\b', r'\1호', kw_str)
    kw_str = re.sub(r'(\d+)\s+층\b', r'\1층', kw_str)
    kw_str = re.sub(r'\b(\d+)-(\d+)호\b', r'\1동 \2호', kw_str)
    
    delivery_note_pattern = r'([가-힣\s]+(?:부탁드립니다?|놓아주세요?|전해주세요?|맡겨주세요?|보관해주세요?|부탁해요?))'
    delivery_notes = re.findall(delivery_note_pattern, kw_str)
    for note in delivery_notes:
        kw_str = kw_str.replace(note, '').strip()
    kw_str = re.sub(r'\s*\.\s*', ' ', kw_str).strip()
    
    kw_str = re.sub(r'남동\s+구', '남동구', kw_str)
    kw_str = re.sub(r'서\s+구', '서구', kw_str)
    kw_str = kw_str.replace('대소면', '대소읍')

    for target_key, override_addr in SPECIAL_EXCEPTIONS.items():
        if target_key in kw_str:
            extra_part = kw_str
            for part in target_key.split():
                extra_part = extra_part.replace(part, '')
            extra_pattern = r'(?:\b\d+동\s*\d+호?|\b[가나다라마바사아자차카타파하A-Za-z]\s*동\s*\d+호?|\d+호|\d+층)'
            found_details = re.findall(extra_pattern, extra_part)
            
            building_names = []
            dongs = []
            hos = []
            for d in found_details:
                for tok in d.split():
                    if re.search(r'\d+호$\vert{}호$', tok) or tok.endswith('호'): hos.append(tok)
                    elif re.search(r'^[가-힣A-Za-z]\s*동$\vert{}\d+동$', tok) or tok.endswith('동'): dongs.append(tok)
                    else: building_names.append(tok)
            
            ordered_details = building_names + dongs + hos
            final_res = f"{override_addr} {' '.join(ordered_details)}"
            
            if delivery_notes:
                cleaned_notes = " ".join([n.strip() for n in delivery_notes if n.strip()])
                final_res = f"{final_res} ({cleaned_notes})"
                
            return remove_duplicate_words(final_res)

    if '불로동' in kw_str:
        kw_str = kw_str.replace('서구', '검단구').replace('서해구', '검단구')
        if '인천광역시 검단구' not in kw_str and '인천 검단구' not in kw_str:
            kw_str = re.sub(r'인천광역시\s+서구', '인천광역시 검단구', kw_str)
            kw_str = re.sub(r'인천\s+서구', '인천 검단구', kw_str)

    extra_pattern = r'(?:\b\d+동\s*\d+호?|\b[가나다라마바사아자차카타파하A-Za-z]\s*동\s*\d+호?|\b[가나다라마바사아자차카타파하A-Za-z]+동\d+|\d+호|\d+층|B\d+호|관리실|택배보관함|물리치료실|\([^)]+\)|[가-힣]+(?:의원|병원|한의원|이비인후과|내과|외과|치과|소아과|센터))'
    extra_details = re.findall(extra_pattern, kw_str)
    
    search_q_str = re.sub(extra_pattern, '', kw_str)
    search_q_str = ' '.join(search_q_str.split())

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
    
    if '월산동 986-3' in search_q_str or '월산동 986' in search_q_str:
        extra = search_q_str.replace('광주광역시', '').replace('전남광주통합특별시', '').replace('남구', '').replace('월산동', '').replace('986-3', '').replace('986', '').strip()
        res_str = f"광주광역시 남구 대남대로 363 {extra} {' '.join(extra_details)}".strip()
        if delivery_notes:
            cleaned_notes = " ".join([n.strip() for n in delivery_notes if n.strip()])
            res_str = f"{res_str} ({cleaned_notes})"
        return remove_duplicate_words(res_str)

    base_tokens = search_q_str.split()
    sido_sigungu_dong_tokens = []
    jibeon_token = ""
    building_tokens = []
    
    for t in base_tokens:
        if re.match(r'^\d+(-\d+)?$', t) or re.match(r'^산\d+(-\d+)?$', t):
            if t != '0' or ('2차' not in kw_str and '2단지' not in kw_str):
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
    api_zip_no = ""
    is_user_sangga = '상가' in kw_str
    has_2cha = '2차' in kw_str or '2단지' in kw_str
    
    for q in query_candidates:
        if not q.strip():
            continue
        url = f"https://business.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword={urllib.parse.quote(q)}&confmKey={CONFIRM_KEY}&resultType=json"
        
        try:
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                res = response
