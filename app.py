import io
import re
import urllib.parse
import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Font

# 🚀 발급받은 브이월드 API 키 적용 완료
VWORLD_API_KEY = "7F7E16A8-41A3-459F-8315-0C733BDBACBA"

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
        res = f"태블릿 모니터 2단 거치대[{qty}]" if qty > 1 else "태블릿 모니터 2단
