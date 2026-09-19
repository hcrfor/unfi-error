# -*- coding: utf-8 -*-
"""
산림 조사 데이터(엑셀) 오류 검증 스크립트
작성일: 2026-09-13
설명: excel_1-4 폴더 내 모든 엑셀 파일을 읽어 5대 오류 검증 규칙에 따라 정밀 분석 후 오류 리포트 엑셀 파일을 생성합니다.
"""

import os
import glob
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def clean_str(val):
    """문자열 공백 제거 및 None 처리"""
    if val is None:
        return ''
    return str(val).strip()

def parse_range(val):
    """
    범위 문자열을 파싱하여 (시작값, 끝값) 튜플을 반환합니다.
    예: '12.4~5.7' -> (12.4, 5.7)
    '~' 또는 공백인 경우 None 반환
    형식이 잘못된 경우 (예: '3.3~') 'INVALID' 반환
    """
    s = clean_str(val)
    if not s or s == '~' or s == '-':
        return None
    
    # 12.4~5.7 형태 매칭
    m = re.match(r'^([0-9]+(?:\.[0-9]+)?)\s*~\s*([0-9]+(?:\.[0-9]+)?)$', s)
    if m:
        return (float(m.group(1)), float(m.group(2)))
    
    return 'INVALID'

def to_float(val):
    """퍼센트 기호 등을 제거하고 실수로 변환"""
    s = clean_str(val).replace('%', '')
    try:
        return float(s)
    except (ValueError, TypeError):
        return None

# =============================================================
# 덩굴류(Vines/Lianas) 수종 목록 및 판별 함수
# 칡, 담쟁이덩굴, 노박덩굴 등 덩굴식물은 다른 수목을 타고 올라가 수고가 높게 측정될 수 있으나
# 자립 수관층을 형성하지 않으므로 임분현황 층위(상층/중층/하층) 구간 및 대표수종 산정에서 제외됩니다.
# =============================================================
VINE_EXACT_SPECIES = {
    '칡', '담쟁이', '담쟁이덩굴', '미국담쟁이덩굴',
    '노박덩굴', '등나무', '등', '머루', '왕머루', '개머루', '새머루',
    '다래', '개다래', '쥐다래', '으름', '으름덩굴',
    '능소화', '마삭줄', '백화등', '송악', '환삼덩굴', '청미래덩굴',
    '인동', '인동덩굴', '사위질빵', '할미밀망', '계요등', '배풍등',
    '거지덩굴', '댕댕이덩굴', '오미자', '흑오미자', '포도나무', '포도',
    '마', '참마', '단풍마', '부채마', '줄딸기'
}

def is_vine_species(val):
    """
    수종명이 덩굴류(칡, 담쟁이덩굴 등)에 해당하는지 검사합니다.
    1. 수종명에 '덩굴' 단어가 포함된 경우 (예: 담쟁이덩굴, 노박덩굴, 청미래덩굴 등)
    2. VINE_EXACT_SPECIES에 정의된 대표적 덩굴류 수종인 경우
    3. 수종명에 대표 덩굴류 키워드가 포함된 경우
    """
    s = clean_str(val).replace(' ', '')
    if not s:
        return False
    if '덩굴' in s:
        return True
    if s in VINE_EXACT_SPECIES:
        return True
    for vine_name in ['칡', '담쟁이', '등나무', '머루', '능소화', '마삭줄', '백화등', '송악', '청미래']:
        if vine_name in s:
            return True
    return False

def analyze_excel_files(target_dir='excel_1-4', output_file='오류_분석_결과.xlsx'):
    # 대상 파일 목록 수집 (임시파일 ~$ 제외)
    file_paths = sorted([
        f for f in glob.glob(os.path.join(target_dir, '*.xlsx'))
        if not os.path.basename(f).startswith('~$')
    ])
    
    total_files = len(file_paths)
    print(f"총 {total_files}개 엑셀 파일 분석을 시작합니다...")
    
    # 오류 기록 리스트
    detailed_errors = []
    sample_summary = {}
    rule_stats = {
        "규칙 1 (교란(인위적) 누락)": 0,
        "규칙 2 (종명 누락)": 0,
        "규칙 3 (피도 누락 및 불일치)": 0,
        "규칙 4 (범위 연결 오류 및 형식 오류)": 0,
        "규칙 5 (수고 최댓값과 상층 범위 불일치)": 0,
        "기타 (시트 누락 등)": 0
    }
    
    for idx, fpath in enumerate(file_paths, 1):
        fname = os.path.basename(fpath)
        sample_no_from_fname = fname.split('_')[0]
        
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True)
        except Exception as e:
            err_msg = f"엑셀 파일을 열 수 없습니다: {str(e)}"
            detailed_errors.append({
                "표본점번호": sample_no_from_fname,
                "파일명": fname,
                "검증규칙": "기타 (파일 손상)",
                "오류분류": "파일 오픈 오류",
                "오류내용": err_msg,
                "상세정보": str(e)
            })
            rule_stats["기타 (시트 누락 등)"] += 1
            continue
            
        # 1. 임분현황조사표 확인
        if '임분현황조사표' not in wb.sheetnames:
            detailed_errors.append({
                "표본점번호": sample_no_from_fname,
                "파일명": fname,
                "검증규칙": "기타 (시트 누락)",
                "오류분류": "시트 누락",
                "오류내용": "'임분현황조사표' 탭이 파일에 존재하지 않습니다.",
                "상세정보": f"존재 시트: {wb.sheetnames}"
            })
            rule_stats["기타 (시트 누락 등)"] += 1
            continue
            
        ws_stand = wb['임분현황조사표']
        headers_s = [ws_stand.cell(1, c).value for c in range(1, ws_stand.max_column + 1)]
        h_map_s = {clean_str(h): c for c, h in enumerate(headers_s, 1) if h is not None}
        
        # 표본점번호 추출 (2행의 표본점번호 열 우선, 없으면 파일명에서 추출)
        sample_no_cell = ws_stand.cell(2, h_map_s.get('표본점번호', 1)).value
        sample_no = clean_str(sample_no_cell) if sample_no_cell is not None else sample_no_from_fname
        if not sample_no:
            sample_no = sample_no_from_fname
            
        file_errors = []
        
        # -------------------------------------------------------------
        # [규칙 1] 임분현황조사표 탭에서 '교란(인위적)' 열에 값이 없으면 오류
        # -------------------------------------------------------------
        disturb_col = h_map_s.get('교란(인위적)')
        disturb_val = ws_stand.cell(2, disturb_col).value if disturb_col else None
        disturb_str = clean_str(disturb_val)
        
        if disturb_val is None or disturb_str == '':
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 1 (교란(인위적) 누락)",
                "오류분류": "교란(인위적) 미입력",
                "오류내용": "임분현황조사표의 '교란(인위적)' 열에 값이 입력되어 있지 않습니다.",
                "상세정보": f"입력값: 빈 값(None)"
            })
            rule_stats["규칙 1 (교란(인위적) 누락)"] += 1
            
        # 표본점유형 추출 (일반·토지현황조사표가 있는 경우)
        sample_type_val = ''
        if '일반·토지현황조사표' in wb.sheetnames:
            ws_g = wb['일반·토지현황조사표']
            h_g = {clean_str(ws_g.cell(1, c).value): c for c in range(1, ws_g.max_column + 1)}
            raw_pt = ws_g.cell(2, h_g.get('표본점유형')).value if h_g.get('표본점유형') else None
            if raw_pt is not None:
                if isinstance(raw_pt, (int, float)) and float(raw_pt).is_integer():
                    sample_type_val = str(int(raw_pt))
                else:
                    sample_type_val = clean_str(raw_pt)

        # -------------------------------------------------------------
        # [신규 규칙] 구조물 부재 시 특이사항 검증 (표본점유형 '5' 제외)
        # 구조물1, 2, 3 열에 내용이 모두 없다면 특이사항 열에 "구조물없음" 필수
        # -------------------------------------------------------------
        if sample_type_val != "5":
            st1_val = clean_str(ws_stand.cell(2, h_map_s.get('구조물1')).value) if h_map_s.get('구조물1') else ''
            st2_val = clean_str(ws_stand.cell(2, h_map_s.get('구조물2')).value) if h_map_s.get('구조물2') else ''
            st3_val = clean_str(ws_stand.cell(2, h_map_s.get('구조물3')).value) if h_map_s.get('구조물3') else ''
            
            if not st1_val and not st2_val and not st3_val:
                note_col = h_map_s.get('특이사항')
                note_val = clean_str(ws_stand.cell(2, note_col).value) if note_col else ''
                if '구조물없음' not in note_val.replace(' ', ''):
                    file_errors.append({
                        "표본점번호": sample_no,
                        "파일명": fname,
                        "검증규칙": "구조물 부재 시 특이사항 기재 오류",
                        "오류분류": "특이사항 미입력",
                        "오류내용": "구조물1, 구조물2, 구조물3 열에 입력된 내용이 모두 없으므로, 특이사항 열에 반드시 '구조물없음'이 입력되어 있어야 합니다.",
                        "상세정보": f"입력값: '{note_val}'"
                    })
                    rule_stats["기타 (시트 누락 등)"] += 1

        # -------------------------------------------------------------
        # [규칙 2] 상층(종명), 하층(종명) 열에 값이 없으면 오류.
        # ※ 중층(종명)은 사용자가 자율적으로 결정하므로 누락 오류를 체크하지 않음
        # -------------------------------------------------------------
        top_sp_val = ws_stand.cell(2, h_map_s.get('상층(종명)')).value if h_map_s.get('상층(종명)') else None
        bot_sp_val = ws_stand.cell(2, h_map_s.get('하층(종명)')).value if h_map_s.get('하층(종명)') else None
        
        top_sp_str = clean_str(top_sp_val)
        bot_sp_str = clean_str(bot_sp_val)
        
        has_top_sp = bool(top_sp_str and top_sp_str not in ['~', '-'])
        has_bot_sp = bool(bot_sp_str and bot_sp_str not in ['~', '-'])
        
        # 상층 종명 누락
        if not has_top_sp:
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 2 (종명 누락)",
                "오류분류": "상층(종명) 누락",
                "오류내용": "임분현황조사표의 '상층(종명)' 열에 식물종명이 입력되지 않았습니다.",
                "상세정보": f"입력값: '{top_sp_str}'"
            })
            rule_stats["규칙 2 (종명 누락)"] += 1
            
        # 하층 종명 누락
        if not has_bot_sp:
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 2 (종명 누락)",
                "오류분류": "하층(종명) 누락",
                "오류내용": "임분현황조사표의 '하층(종명)' 열에 식물종명이 입력되지 않았습니다.",
                "상세정보": f"입력값: '{bot_sp_str}'"
            })
            rule_stats["규칙 2 (종명 누락)"] += 1

        # -------------------------------------------------------------
        # [신규 규칙] 임분현황 층위(상층, 중층, 하층) 덩굴류 기재 검증
        # 덩굴류(칡, 담쟁이덩굴 등)는 수고가 아무리 높아도 임분현황조사표의 상층, 중층, 하층 층위 구간(종명)에 입력 불가
        # -------------------------------------------------------------
        mid_sp_val = ws_stand.cell(2, h_map_s.get('중층(종명)')).value if h_map_s.get('중층(종명)') else None
        layer_checks = [('상층', top_sp_str), ('중층', clean_str(mid_sp_val)), ('하층', bot_sp_str)]
        for l_name, l_sp in layer_checks:
            if l_sp and is_vine_species(l_sp):
                file_errors.append({
                    "표본점번호": sample_no,
                    "파일명": fname,
                    "검증규칙": "규칙 (덩굴류 수종 층위 기재 오류)",
                    "오류분류": f"{l_name}(종명) 덩굴류 기재",
                    "오류내용": f"임분현황조사표의 {l_name}(종명)에 덩굴류('{l_sp}')가 입력되었습니다. 덩굴류(칡, 담쟁이덩굴 등)는 상층, 중층, 하층의 구간에 입력될 수 없습니다.",
                    "상세정보": f"입력값: '{l_sp}'"
                })
                rule_stats.setdefault("규칙 (덩굴류 수종 층위 기재 오류)", 0)
                rule_stats["규칙 (덩굴류 수종 층위 기재 오류)"] += 1
            
        # -------------------------------------------------------------
        # [규칙 3] 피도 확인
        # 상층피도, 하층피도, 전체피도에 올바른 값이 입력되어야 함.
        # ※ 중층피도는 사용자가 자율적으로 결정함
        # 전체피도에는 상층, 중층(입력된 경우), 하층 피도 중 가장 큰 값이 있어야 함.
        # -------------------------------------------------------------
        top_cov_val = ws_stand.cell(2, h_map_s.get('상층피도(퍼센트)')).value if h_map_s.get('상층피도(퍼센트)') else None
        mid_cov_val = ws_stand.cell(2, h_map_s.get('중층피도(퍼센트)')).value if h_map_s.get('중층피도(퍼센트)') else None
        bot_cov_val = ws_stand.cell(2, h_map_s.get('하층피도(퍼센트)')).value if h_map_s.get('하층피도(퍼센트)') else None
        tot_cov_val = ws_stand.cell(2, h_map_s.get('전체피도(퍼센트)')).value if h_map_s.get('전체피도(퍼센트)') else None
        
        top_cov_num = to_float(top_cov_val)
        mid_cov_num = to_float(mid_cov_val)
        bot_cov_num = to_float(bot_cov_val)
        tot_cov_num = to_float(tot_cov_val)
        
        # 상층피도 누락
        if top_cov_num is None:
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 3 (피도 누락 및 불일치)",
                "오류분류": "상층피도 누락",
                "오류내용": "임분현황조사표의 '상층피도(퍼센트)' 열에 올바른 피도 수치가 입력되지 않았습니다.",
                "상세정보": f"입력값: '{clean_str(top_cov_val)}'"
            })
            rule_stats["규칙 3 (피도 누락 및 불일치)"] += 1
            
        # 하층피도 누락
        if bot_cov_num is None:
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 3 (피도 누락 및 불일치)",
                "오류분류": "하층피도 누락",
                "오류내용": "임분현황조사표의 '하층피도(퍼센트)' 열에 올바른 피도 수치가 입력되지 않았습니다.",
                "상세정보": f"입력값: '{clean_str(bot_cov_val)}'"
            })
            rule_stats["규칙 3 (피도 누락 및 불일치)"] += 1
            
        # 전체피도 누락 및 최댓값 일치 검증
        if tot_cov_num is None:
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 3 (피도 누락 및 불일치)",
                "오류분류": "전체피도 누락",
                "오류내용": "임분현황조사표의 '전체피도(퍼센트)' 열에 올바른 피도 수치가 입력되지 않았습니다.",
                "상세정보": f"입력값: '{clean_str(tot_cov_val)}'"
            })
            rule_stats["규칙 3 (피도 누락 및 불일치)"] += 1
        else:
            # 입력된 층별 피도 중 존재하는 값들의 최댓값 산출
            cov_candidates = []
            if top_cov_num is not None:
                cov_candidates.append(top_cov_num)
            if mid_cov_num is not None:
                cov_candidates.append(mid_cov_num)
            if bot_cov_num is not None:
                cov_candidates.append(bot_cov_num)
                
            if cov_candidates:
                expected_max_cov = max(cov_candidates)
                if round(tot_cov_num, 1) != round(expected_max_cov, 1):
                    file_errors.append({
                        "표본점번호": sample_no,
                        "파일명": fname,
                        "검증규칙": "규칙 3 (피도 누락 및 불일치)",
                        "오류분류": "전체피도 불일치",
                        "오류내용": f"전체피도({tot_cov_num}%)는 각 층의 피도 중 최댓값인 {expected_max_cov}%와 같아야 하지만 다른 값이 입력되어 있습니다.",
                        "상세정보": f"입력 전체피도={tot_cov_num}%, 상층={top_cov_num}%, 중층={mid_cov_num}%, 하층={bot_cov_num}% -> 최댓값={expected_max_cov}%"
                    })
                    rule_stats["규칙 3 (피도 누락 및 불일치)"] += 1

        # -------------------------------------------------------------
        # [규칙 4] 상층(범위), 중층(범위), 하층(범위) 연결성 및 형식 검증
        # -------------------------------------------------------------
        top_rng_raw = ws_stand.cell(2, h_map_s.get('상층(범위)')).value if h_map_s.get('상층(범위)') else None
        mid_rng_raw = ws_stand.cell(2, h_map_s.get('중층(범위)')).value if h_map_s.get('중층(범위)') else None
        bot_rng_raw = ws_stand.cell(2, h_map_s.get('하층(범위)')).value if h_map_s.get('하층(범위)') else None
        
        top_rng_str = clean_str(top_rng_raw)
        mid_rng_str = clean_str(mid_rng_raw)
        bot_rng_str = clean_str(bot_rng_raw)
        
        top_parsed = parse_range(top_rng_raw)
        mid_parsed = parse_range(mid_rng_raw)
        bot_parsed = parse_range(bot_rng_raw)
        
        # 형식 오류 확인
        if top_parsed == 'INVALID':
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 4 (범위 연결 오류 및 형식 오류)",
                "오류분류": "상층(범위) 형식 불량",
                "오류내용": f"상층(범위) 값이 '시작~끝' 숫자 형식(예: 12.4~5.7)이 아닙니다.",
                "상세정보": f"입력값: '{top_rng_str}'"
            })
            rule_stats["규칙 4 (범위 연결 오류 및 형식 오류)"] += 1
            
        if mid_parsed == 'INVALID':
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 4 (범위 연결 오류 및 형식 오류)",
                "오류분류": "중층(범위) 형식 불량",
                "오류내용": f"중층(범위) 값이 '시작~끝' 숫자 형식(예: 5.7~0.8)이 아닙니다.",
                "상세정보": f"입력값: '{mid_rng_str}'"
            })
            rule_stats["규칙 4 (범위 연결 오류 및 형식 오류)"] += 1
            
        if bot_parsed == 'INVALID':
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 4 (범위 연결 오류 및 형식 오류)",
                "오류분류": "하층(범위) 형식 불량",
                "오류내용": f"하층(범위) 값이 '시작~끝' 숫자 형식(예: 0.8~0.0)이 아닙니다.",
                "상세정보": f"입력값: '{bot_rng_str}'"
            })
            rule_stats["규칙 4 (범위 연결 오류 및 형식 오류)"] += 1

        # 누락 확인 (상층, 하층)
        if top_parsed is None:
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 4 (범위 연결 오류 및 형식 오류)",
                "오류분류": "상층(범위) 누락",
                "오류내용": "임분현황조사표의 '상층(범위)' 열 값이 누락되었습니다.",
                "상세정보": f"입력값: '{top_rng_str}'"
            })
            rule_stats["규칙 4 (범위 연결 오류 및 형식 오류)"] += 1
            
        if bot_parsed is None:
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 4 (범위 연결 오류 및 형식 오류)",
                "오류분류": "하층(범위) 누락",
                "오류내용": "임분현황조사표의 '하층(범위)' 열 값이 누락되었습니다.",
                "상세정보": f"입력값: '{bot_rng_str}'"
            })
            rule_stats["규칙 4 (범위 연결 오류 및 형식 오류)"] += 1

        # 연결성 검증: 상층과 하층이 모두 정상 튜플일 때
        if isinstance(top_parsed, tuple) and isinstance(bot_parsed, tuple):
            top_start, top_end = top_parsed
            bot_start, bot_end = bot_parsed
            
            # 중층이 존재하는 경우
            if isinstance(mid_parsed, tuple):
                mid_start, mid_end = mid_parsed
                # 1) 상층 끝값과 중층 시작값 일치 검사
                if round(top_end, 1) != round(mid_start, 1):
                    file_errors.append({
                        "표본점번호": sample_no,
                        "파일명": fname,
                        "검증규칙": "규칙 4 (범위 연결 오류 및 형식 오류)",
                        "오류분류": "상층-중층 범위 불일치",
                        "오류내용": f"상층의 끝값({top_end})과 중층의 시작값({mid_start})이 연결되지 않습니다.",
                        "상세정보": f"상층(범위)='{top_rng_str}', 중층(범위)='{mid_rng_str}'"
                    })
                    rule_stats["규칙 4 (범위 연결 오류 및 형식 오류)"] += 1
                # 2) 중층 끝값과 하층 시작값 일치 검사
                if round(mid_end, 1) != round(bot_start, 1):
                    file_errors.append({
                        "표본점번호": sample_no,
                        "파일명": fname,
                        "검증규칙": "규칙 4 (범위 연결 오류 및 형식 오류)",
                        "오류분류": "중층-하층 범위 불일치",
                        "오류내용": f"중층의 끝값({mid_end})과 하층의 시작값({bot_start})이 연결되지 않습니다.",
                        "상세정보": f"중층(범위)='{mid_rng_str}', 하층(범위)='{bot_rng_str}'"
                    })
                    rule_stats["규칙 4 (범위 연결 오류 및 형식 오류)"] += 1
            elif mid_parsed is None:
                # 중층이 생략된 경우: 상층 끝값과 하층 시작값이 바로 연결되어야 함
                if round(top_end, 1) != round(bot_start, 1):
                    file_errors.append({
                        "표본점번호": sample_no,
                        "파일명": fname,
                        "검증규칙": "규칙 4 (범위 연결 오류 및 형식 오류)",
                        "오류분류": "상층-하층 범위 불일치",
                        "오류내용": f"중층이 생략된 상태에서 상층 끝값({top_end})과 하층 시작값({bot_start})이 연결되지 않습니다.",
                        "상세정보": f"상층(범위)='{top_rng_str}', 중층(범위)='{mid_rng_str}', 하층(범위)='{bot_rng_str}'"
                    })
                    rule_stats["규칙 4 (범위 연결 오류 및 형식 오류)"] += 1

        # -------------------------------------------------------------
        # [규칙 5] 임목조사표(임목조사) 탭의 수고(cm) 최댓값(m 환산)과
        #          임분현황조사표 탭의 상층(범위) 첫번째 숫자 일치 검증
        # -------------------------------------------------------------
        if '임목조사표(임목조사)' not in wb.sheetnames:
            file_errors.append({
                "표본점번호": sample_no,
                "파일명": fname,
                "검증규칙": "규칙 5 (수고 최댓값과 상층 범위 불일치)",
                "오류분류": "임목조사표 시트 누락",
                "오류내용": "'임목조사표(임목조사)' 시트가 없어 수고 최댓값을 확인할 수 없습니다.",
                "상세정보": "시트 누락"
            })
            rule_stats["규칙 5 (수고 최댓값과 상층 범위 불일치)"] += 1
        else:
            ws_tree = wb['임목조사표(임목조사)']
            t_headers = [ws_tree.cell(1, c).value for c in range(1, ws_tree.max_column + 1)]
            th_map = {clean_str(h): c for c, h in enumerate(t_headers, 1) if h is not None}
            h_col = th_map.get('수고(cm)')
            
            tree_heights = []
            if h_col:
                for r in range(2, ws_tree.max_row + 1):
                    val = ws_tree.cell(r, h_col).value
                    num = to_float(val)
                    if num is not None and num > 0:
                        tree_heights.append(num)
                        
            if not tree_heights:
                file_errors.append({
                    "표본점번호": sample_no,
                    "파일명": fname,
                    "검증규칙": "규칙 5 (수고 최댓값과 상층 범위 불일치)",
                    "오류분류": "수고(cm) 데이터 부재",
                    "오류내용": "임목조사표(임목조사)에 측정된 수고(cm) 데이터가 하나도 없습니다.",
                    "상세정보": f"조사된 임목 수고 없음 (총 {ws_tree.max_row-1}개 행)"
                })
                rule_stats["규칙 5 (수고 최댓값과 상층 범위 불일치)"] += 1
            else:
                max_tree_h_cm = max(tree_heights)
                max_tree_h_m = max_tree_h_cm / 100.0  # cm -> m 단위 변환
                
                if isinstance(top_parsed, tuple):
                    top_start_m = top_parsed[0]
                    # 소수점 1자리 기준 비교 (반올림 오차 방지)
                    if round(max_tree_h_m, 1) != round(top_start_m, 1):
                        file_errors.append({
                            "표본점번호": sample_no,
                            "파일명": fname,
                            "검증규칙": "규칙 5 (수고 최댓값과 상층 범위 불일치)",
                            "오류분류": "수고 최댓값-상층범위 불일치",
                            "오류내용": f"임목조사표의 수고 최댓값({max_tree_h_m:.1f}m / {max_tree_h_cm:.0f}cm)과 상층(범위)의 시작값({top_start_m:.1f}m)이 일치하지 않습니다.",
                            "상세정보": f"수고 최댓값={max_tree_h_m:.1f}m ({max_tree_h_cm:.0f}cm) != 상층(범위)='{top_rng_str}'"
                        })
                        rule_stats["규칙 5 (수고 최댓값과 상층 범위 불일치)"] += 1
                else:
                    file_errors.append({
                        "표본점번호": sample_no,
                        "파일명": fname,
                        "검증규칙": "규칙 5 (수고 최댓값과 상층 범위 불일치)",
                        "오류분류": "상층(범위) 값 불량으로 비교 불가",
                        "오류내용": f"상층(범위)이 유효한 숫자가 아니어서 임목 수고 최댓값({max_tree_h_m:.1f}m)과 대조할 수 없습니다.",
                        "상세정보": f"상층(범위)='{top_rng_str}', 수고 최댓값={max_tree_h_m:.1f}m"
                    })
                    rule_stats["규칙 5 (수고 최댓값과 상층 범위 불일치)"] += 1
                    
        # 파일별 오류 저장
        if file_errors:
            detailed_errors.extend(file_errors)
            sample_summary[sample_no] = {
                "표본점번호": sample_no,
                "파일명": fname,
                "오류건수": len(file_errors),
                "오류목록": " | ".join([e["오류내용"] for e in file_errors])
            }
            
        if idx % 50 == 0 or idx == total_files:
            print(f"[{idx}/{total_files}] 진행 중... (현재까지 누적 오류 건수: {len(detailed_errors)})")
            
    print(f"\n분석 완료! 총 오류 건수: {len(detailed_errors)} 건, 오류 발생 파일: {len(sample_summary)} 개")
    
    # =========================================================================
    # 결과 엑셀 파일 생성 (오류 리포트) - 고품질 디자인 적용
    # =========================================================================
    wb_out = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # 시트 1: 오류_상세_목록 (사용자가 직관적으로 확인 가능한 메인 시트)
    # -------------------------------------------------------------
    ws1 = wb_out.active
    ws1.title = "오류_상세_목록"
    
    headers1 = ["연번", "표본점번호", "파일명", "검증 규칙", "오류 분류", "오류 상세 내용", "관련 입력값/근거"]
    ws1.append(headers1)
    
    for row_idx, item in enumerate(detailed_errors, 1):
        ws1.append([
            row_idx,
            item["표본점번호"],
            item["파일명"],
            item["검증규칙"],
            item["오류분류"],
            item["오류내용"],
            item["상세정보"]
        ])
        
    # -------------------------------------------------------------
    # 시트 2: 표본점별_오류_요약 (표본점번호 기준 요약)
    # -------------------------------------------------------------
    ws2 = wb_out.create_sheet(title="표본점별_오류_요약")
    headers2 = ["연번", "표본점번호", "파일명", "총 오류 건수", "오류 종합 내용"]
    ws2.append(headers2)
    
    for row_idx, (s_no, s_info) in enumerate(sorted(sample_summary.items(), key=lambda x: -x[1]["오류건수"]), 1):
        ws2.append([
            row_idx,
            s_info["표본점번호"],
            s_info["파일명"],
            s_info["오류건수"],
            s_info["오류목록"]
        ])
        
    # -------------------------------------------------------------
    # 시트 3: 검증_요약_통계 (대시보드 통계)
    # -------------------------------------------------------------
    ws3 = wb_out.create_sheet(title="검증_통계_요약")
    ws3.append(["구분", "수치", "비고"])
    ws3.append(["전체 분석 파일 수", total_files, "excel_1-4 폴더 내 전체 .xlsx"])
    ws3.append(["정상 파일 수", total_files - len(sample_summary), f"{((total_files - len(sample_summary))/total_files*100):.1f}%"])
    ws3.append(["오류 발생 파일 수", len(sample_summary), f"{(len(sample_summary)/total_files*100):.1f}%"])
    ws3.append(["총 검출된 오류 건수", len(detailed_errors), "중복 표본점 내 다중 오류 포함"])
    ws3.append(["", "", ""])
    ws3.append(["[규칙별 오류 발생 현황]", "건수", "점유율(%)"])
    
    for r_name, r_cnt in rule_stats.items():
        pct = (r_cnt / len(detailed_errors) * 100) if detailed_errors else 0
        ws3.append([r_name, r_cnt, f"{pct:.1f}%"])

    # -------------------------------------------------------------
    # 스타일 서식 적용 (폰트, 배경색, 테두리, 열너비 자동 조정)
    # -------------------------------------------------------------
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid") # 고급 네이비
    sub_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    white_font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="맑은 고딕", size=11, bold=True)
    normal_font = Font(name="맑은 고딕", size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    for ws in [ws1, ws2, ws3]:
        ws.views.sheetView[0].showGridLines = True
        
        # 헤더 서식
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28
        
        # 데이터 행 서식
        for r in range(2, ws.max_row + 1):
            ws.row_dimensions[r].height = 22
            is_zebra = (r % 2 == 0)
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                cell.font = normal_font
                cell.border = thin_border
                if is_zebra:
                    cell.fill = zebra_fill
                
                # 정렬 규칙
                if c in [1, 2]: # 연번, 표본점번호
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c == 3 and ws == ws1: # 파일명
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif c in [4, 5] and ws == ws1: # 규칙, 오류분류
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c in [6, 7] and ws == ws1: # 오류내용, 상세
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif ws == ws2 and c == 4: # 총 오류건수
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif ws == ws3:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    
        # 열 너비 자동 조정
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                # 한글 문자 길이 보정 (2바이트 계산)
                length = sum(2 if ord(ch) > 128 else 1 for ch in val_str)
                if length > max_len:
                    max_len = length
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    # 특정 긴 열 너비 제한 및 조정
    ws1.column_dimensions['F'].width = 50  # 오류 상세 내용
    ws1.column_dimensions['G'].width = 40  # 관련 입력값
    ws2.column_dimensions['E'].width = 75  # 오류 종합 내용
    
    wb_out.save(output_file)
    print(f"엑셀 오류 보고서 생성 완료: {os.path.abspath(output_file)}")
    return output_file, total_files, len(sample_summary), len(detailed_errors), rule_stats

if __name__ == '__main__':
    import excel_inspector
    print("최신 산림 도시 조사 데이터 오류 검수 엔진을 실행합니다...")
    excel_inspector.run_inspection('excel_1-4')
