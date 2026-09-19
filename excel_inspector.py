# -*- coding: utf-8 -*-
"""
산림 도시 조사 데이터(엑셀) 정밀 오류 검수 엔진
파일명: excel_inspector.py
작성자: Antigravity (시니어 풀스택 개발자)
설명:
    - 2026 도시 조사결과표(한성안).xlsx 기준 데이터를 참조하여
    - 사용자가 지정한 폴더 내의 조사 엑셀 파일들을 검수하고
    - 상세 오류 리포트 엑셀 파일을 생성합니다.
"""

import os
import glob
import re
import math
from collections import Counter
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =============================================================================
# 1. 데이터 파싱 및 유효성 검사 헬퍼 함수군
# =============================================================================

def clean_str(val):
    """문자열의 앞뒤 공백을 제거하고 None 값을 빈 문자열로 처리합니다."""
    if val is None:
        return ''
    return str(val).replace('\xa0', ' ').replace('\ufeff', '').strip()


def to_float(val):
    """숫자 또는 숫자 형식 문자열(% 기호 포함 등)을 실수(float)로 안전하게 변환합니다."""
    if val is None:
        return None
    s = clean_str(val).replace('%', '').replace(',', '')
    if not s:
        return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def is_valid_integer(val):
    """
    정수 형태인지 엄격하게 검증합니다.
    (예: 10, 10.0, '10' -> True / 10.5, 'abc', '', None -> False)
    """
    num = to_float(val)
    if num is None:
        return False
    return num.is_integer()


def is_within_one_decimal(val):
    """
    소수점 첫째자리까지 유효한 값인지 검증합니다.
    (예: 12, 12.0, 12.3, '12.3' -> True / 12.34, None, '' -> False)
    """
    num = to_float(val)
    if num is None:
        return False
    # 소수 둘째자리 이상 차이가 나는지 확인 (부동소수점 오차 감안)
    return abs(num - round(num, 1)) < 1e-6


def is_multiple_of_five(val):
    """
    5 단위의 값인지 검증합니다. (예: 25, 30, 35 ... 100)
    """
    num = to_float(val)
    if num is None:
        return False
    if not num.is_integer():
        return False
    return int(num) % 5 == 0


def is_multiple_of_half(val):
    """
    0.5 단위의 값인지 검증합니다. (예: 0.5, 1.0, 1.5, 2.0 ...)
    """
    num = to_float(val)
    if num is None or num <= 0:
        return False
    # 2를 곱했을 때 정수가 되는지 확인
    doubled = num * 2.0
    return abs(doubled - round(doubled)) < 1e-6


def parse_range(val):
    """
    범위 문자열을 파싱하여 (시작값, 끝값) 튜플을 반환합니다.
    예: '12.4~5.7' -> (12.4, 5.7)
    '~' 또는 공백인 경우 None 반환
    형식이 잘못된 경우 'INVALID' 반환
    """
    s = clean_str(val)
    if not s or s == '~' or s == '-':
        return None
    
    m = re.match(r'^([0-9]+(?:\.[0-9]+)?)\s*~\s*([0-9]+(?:\.[0-9]+)?)$', s)
    if m:
        return (float(m.group(1)), float(m.group(2)))
    
    return 'INVALID'


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


# =============================================================================
# 2. 기준 엑셀 데이터 로더 (2026 도시 조사결과표(한성안).xlsx)
# =============================================================================

class ReferenceDataLoader:
    """
    '2026 도시 조사결과표(한성안).xlsx' 파일의 '표본점목록' 탭 데이터를 메모리에 캐싱하여
    초고속 대조를 가능하게 해주는 로더 클래스입니다.
    """
    def __init__(self, ref_file_path):
        self.ref_file_path = ref_file_path
        self.sample_dict = {}  # {표본점번호: {'POINT_X': float, 'POINT_Y': float, '표고': float, ...}}
        self.team_dict = {}    # {표본점번호: {'팀장': str, '팀원': str}}
        self.is_loaded = False
        self.load_error = None

    def load(self):
        """기준 엑셀 파일 로드 및 표본점 사전 생성"""
        if self.is_loaded:
            return True

        if not os.path.exists(self.ref_file_path):
            self.load_error = f"기준 엑셀 파일을 찾을 수 없습니다: {self.ref_file_path}"
            return False

        try:
            wb = openpyxl.load_workbook(self.ref_file_path, read_only=True, data_only=True)
            if '표본점목록' not in wb.sheetnames:
                self.load_error = f"기준 엑셀 파일에 '표본점목록' 시트가 존재하지 않습니다."
                wb.close()
                return False

            ws = wb['표본점목록']
            rows = iter(ws.iter_rows(values_only=True))
            header = [clean_str(c) for c in next(rows)]
            
            # 주요 열 인덱스 매핑
            h_map = {name: idx for idx, name in enumerate(header)}
            pid_idx = h_map.get('표본점번호')
            px_idx = h_map.get('POINT_X')
            py_idx = h_map.get('POINT_Y')
            elev_idx = h_map.get('표고')

            if pid_idx is None:
                self.load_error = "기준 파일의 '표본점목록'에 '표본점번호' 열이 없습니다."
                wb.close()
                return False

            for r in rows:
                if not r or r[pid_idx] is None:
                    continue
                pid = clean_str(r[pid_idx])
                if not pid:
                    continue

                px = to_float(r[px_idx]) if px_idx is not None and px_idx < len(r) else None
                py = to_float(r[py_idx]) if py_idx is not None and py_idx < len(r) else None
                elev = to_float(r[elev_idx]) if elev_idx is not None and elev_idx < len(r) else None

                self.sample_dict[pid] = {
                    'POINT_X': px,
                    'POINT_Y': py,
                    '표고': elev
                }

            # '조사결과표' 시트에서 팀장, 팀원 정보 캐싱
            if '조사결과표' in wb.sheetnames:
                try:
                    ws_res = wb['조사결과표']
                    res_rows = iter(ws_res.iter_rows(values_only=True))
                    res_header = [clean_str(c) for c in next(res_rows)]
                    pid_c = None
                    leader_c = None
                    member_cols = []
                    for idx, h in enumerate(res_header):
                        if h == '표본점번호' and pid_c is None:
                            pid_c = idx
                        elif h == '팀장' and leader_c is None:
                            leader_c = idx
                        elif h == '팀원':
                            member_cols.append(idx)
                    if pid_c is not None:
                        for r in res_rows:
                            if not r or r[pid_c] is None:
                                continue
                            p_str = clean_str(r[pid_c])
                            l_str = clean_str(r[leader_c]) if leader_c is not None and leader_c < len(r) else ''
                            m_list = [clean_str(r[c]) for c in member_cols if c < len(r) and clean_str(r[c])]
                            self.team_dict[p_str] = {
                                '팀장': l_str,
                                '팀원': ', '.join(m_list)
                            }
                except Exception:
                    pass

            wb.close()
            self.is_loaded = True
            return True
        except Exception as e:
            self.load_error = f"기준 파일 읽기 중 오류 발생: {str(e)}"
            return False


# =============================================================================
# 3. 조사 파일 검수 엔진
# =============================================================================

class SurveyInspector:
    """
    지정된 폴더의 조사 엑셀 파일들을 읽어 규칙에 맞춰 검수하고
    상세한 오류 내역을 집계하는 클래스입니다.
    """
    def __init__(self, ref_data_loader):
        self.ref_loader = ref_data_loader
        self.errors = []
        self.file_summary = []
        self.rule_stats = Counter()
        self.current_leader = ''
        self.current_member = ''

    def add_error(self, sample_no, fname, sheet, row_no, rule_name, error_field, message, detail_val, leader=None, member=None):
        """오류 항목을 기록합니다."""
        err_item = {
            "표본점번호": str(sample_no),
            "파일명": fname,
            "팀장": leader if leader is not None else self.current_leader,
            "팀원": member if member is not None else self.current_member,
            "시트명": sheet,
            "행번호": row_no if row_no is not None else '-',
            "검증규칙": rule_name,
            "오류항목": error_field,
            "오류내용": message,
            "입력값_상세": str(detail_val) if detail_val is not None else 'None'
        }
        self.errors.append(err_item)
        self.rule_stats[rule_name] += 1
        return err_item

    def inspect_file(self, fpath):
        """개별 엑셀 파일을 열고 전체 검증 규칙을 수행합니다."""
        fname = os.path.basename(fpath)
        sample_no_from_fname = fname.split('_')[0]
        file_errors_before = len(self.errors)
        self.current_leader = ''
        self.current_member = ''

        try:
            wb = openpyxl.load_workbook(fpath, data_only=True)
        except Exception as e:
            self.add_error(
                sample_no=sample_no_from_fname,
                fname=fname,
                sheet="전체",
                row_no=None,
                rule_name="기타 (파일 오픈 오류)",
                error_field="파일 손상",
                message=f"엑셀 파일을 열 수 없습니다: {str(e)}",
                detail_val=str(e)
            )
            return

        # 0. 조사자 정보(팀장, 팀원) 추출
        leader_name = ''
        member_names = ''

        if '일반·토지현황조사표' in wb.sheetnames:
            ws_gen_pre = wb['일반·토지현황조사표']
            members_list = []
            for c in range(1, ws_gen_pre.max_column + 1):
                col_h = clean_str(ws_gen_pre.cell(1, c).value)
                val_r2 = clean_str(ws_gen_pre.cell(2, c).value)
                if col_h == '팀장' and val_r2:
                    leader_name = val_r2
                elif col_h == '팀원' and val_r2:
                    if val_r2 not in members_list:
                        members_list.append(val_r2)
            if members_list:
                member_names = ', '.join(members_list)

        # 기준 파일 fallback
        if not leader_name or not member_names:
            ref_team = self.ref_loader.team_dict.get(sample_no_from_fname, {})
            if not leader_name and ref_team.get('팀장'):
                leader_name = ref_team['팀장']
            if not member_names and ref_team.get('팀원'):
                member_names = ref_team['팀원']

        self.current_leader = leader_name
        self.current_member = member_names

        # -------------------------------------------------------------
        # 시트별 헤더 매핑 함수 (중복 열 발생 시 첫 번째 인덱스 우선 보존)
        # -------------------------------------------------------------
        def get_sheet_header_map(ws):
            h_map = {}
            for c in range(1, ws.max_column + 1):
                name = clean_str(ws.cell(1, c).value)
                if name and name not in h_map:
                    h_map[name] = c
            return h_map

        # =============================================================
        # 1. 일반·토지현황조사표 탭 검증
        # =============================================================
        sample_no = sample_no_from_fname
        sample_type_val = None

        if '일반·토지현황조사표' not in wb.sheetnames:
            self.add_error(
                sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=None,
                rule_name="시트 누락", error_field="일반·토지현황조사표",
                message="'일반·토지현황조사표' 탭이 존재하지 않습니다.", detail_val=None
            )
        else:
            ws_gen = wb['일반·토지현황조사표']
            h_gen = get_sheet_header_map(ws_gen)

            # 표본점번호 확인
            pid_c = h_gen.get('표본점번호')
            if pid_c:
                raw_pid = ws_gen.cell(2, pid_c).value
                if raw_pid:
                    sample_no = clean_str(raw_pid)

            # 표본점유형 추출
            ptype_c = h_gen.get('표본점유형')
            raw_ptype = ws_gen.cell(2, ptype_c).value if ptype_c else None
            if raw_ptype is not None:
                # 5.0 같은 float 형식도 '5'로 처리
                if isinstance(raw_ptype, (int, float)) and float(raw_ptype).is_integer():
                    sample_type_val = str(int(raw_ptype))
                else:
                    sample_type_val = clean_str(raw_ptype)
            else:
                sample_type_val = ''

            # [규칙 1-5] 표본점유형 누락 오류
            if not sample_type_val:
                self.add_error(
                    sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                    rule_name="[일반토지 5] 표본점유형 누락",
                    error_field="표본점유형",
                    message="일반·토지현황조사표의 '표본점유형' 열의 값이 누락되어 있습니다.",
                    detail_val=ws_gen.cell(2, ptype_c).value if ptype_c else None
                )

            # [규칙 1-1] 좌표 오차 검증 (반경 10m 이상 오류, 단 표본점유형 '5' 제외)
            if sample_type_val != "5":
                x_c = h_gen.get('좌표1_X')
                y_c = h_gen.get('좌표1_Y')
                cur_x = to_float(ws_gen.cell(2, x_c).value) if x_c else None
                cur_y = to_float(ws_gen.cell(2, y_c).value) if y_c else None

                ref_info = self.ref_loader.sample_dict.get(sample_no)
                if not ref_info:
                    self.add_error(
                        sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                        rule_name="[일반토지 1] 좌표 오차 검증",
                        error_field="표본점번호 대조",
                        message=f"기준 조사결과표(한성안)의 표본점목록에 해당 표본점번호({sample_no})가 없습니다.",
                        detail_val=sample_no
                    )
                else:
                    ref_x = ref_info.get('POINT_X')
                    ref_y = ref_info.get('POINT_Y')
                    if cur_x is None or cur_y is None:
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                            rule_name="[일반토지 1] 좌표 오차 검증",
                            error_field="좌표1_X / 좌표1_Y",
                            message="좌표1_X 또는 좌표1_Y 값이 올바른 숫자가 아니거나 누락되었습니다.",
                            detail_val=f"X={cur_x}, Y={cur_y}"
                        )
                    elif ref_x is None or ref_y is None:
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                            rule_name="[일반토지 1] 좌표 오차 검증",
                            error_field="기준 좌표 누락",
                            message="기준 파일에 POINT_X 또는 POINT_Y 좌표가 설정되어 있지 않습니다.",
                            detail_val=f"POINT_X={ref_x}, POINT_Y={ref_y}"
                        )
                    else:
                        distance = math.sqrt((cur_x - ref_x) ** 2 + (cur_y - ref_y) ** 2)
                        if distance >= 10.0:
                            self.add_error(
                                sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                                rule_name="[일반토지 1] 좌표 오차 검증",
                                error_field="좌표1_X, 좌표1_Y",
                                message=f"기준 좌표(POINT_X, POINT_Y)와 조사 좌표(좌표1_X, 좌표1_Y) 간의 거리가 {distance:.2f}m로 10m 이상 차이납니다.",
                                detail_val=f"기준=({ref_x}, {ref_y}) vs 조사=({cur_x}, {cur_y}) -> 거리: {distance:.2f}m"
                            )

            # [규칙 1-2] GPS수신상태 검증 ("양호" 필수)
            gps_c = h_gen.get('GPS수신상태')
            gps_val = clean_str(ws_gen.cell(2, gps_c).value) if gps_c else ''
            if gps_val != "양호":
                self.add_error(
                    sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                    rule_name="[일반토지 2] GPS수신상태 오류",
                    error_field="GPS수신상태",
                    message="GPS수신상태는 반드시 '양호'로 입력되어야 합니다.",
                    detail_val=ws_gen.cell(2, gps_c).value if gps_c else '누락'
                )

            # [규칙 1-3] 조사가능(%) 검증 (25 이상, 5 단위 필수)
            pct_c = h_gen.get('조사가능(%)')
            pct_raw = ws_gen.cell(2, pct_c).value if pct_c else None
            pct_num = to_float(pct_raw)
            if pct_num is None or pct_num < 25 or pct_num > 100 or not is_multiple_of_five(pct_num):
                self.add_error(
                    sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                    rule_name="[일반토지 3] 조사가능(%) 오류",
                    error_field="조사가능(%)",
                    message="조사가능(%) 값은 25 이상 100 이하의 5 단위 수치(25, 30, 35 ... 100)로 입력되어야 합니다.",
                    detail_val=pct_raw
                )

            # [규칙 1-4] 표본점유형이 "5"인 경우 조사가능(m) 검증 (0.5 단위 필수)
            if sample_type_val == "5":
                dist_m_c = h_gen.get('조사가능(m)')
                dist_m_raw = ws_gen.cell(2, dist_m_c).value if dist_m_c else None
                if not is_multiple_of_half(dist_m_raw):
                    self.add_error(
                        sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                        rule_name="[일반토지 4] 조사가능(m) 오류",
                        error_field="조사가능(m)",
                        message="표본점유형이 '5'인 경우 조사가능(m) 값은 0.5 단위(0.5, 1.0, 1.5 ...)로 입력되어야 합니다.",
                        detail_val=dist_m_raw
                    )

            # [규칙 1-6] 토지유형 13개 열의 값의 합이 100인지 검증
            # 기본 13개 표준 토지유형 항목 목록
            standard_land_cols = [
                '토지유형(농업용지)', '토지유형(묘지)', '토지유형(상업/공업지)', '토지유형(골프장)',
                '토지유형(공공용지)', '토지유형(다가구주택용지)', '토지유형(단독주택용지)',
                '토지유형(공원)', '토지유형(교통시설)', '토지유형(시설용지)',
                '토지유형(나대지)', '토지유형(수계/습지)', '토지유형(기타)'
            ]
            
            # 헤더에 존재하는 '토지유형' 관련 열들을 모두 수집 (표준 목록 우선 순서 유지 + 추가 열 대응)
            land_type_cols = []
            for col in standard_land_cols:
                if col in h_gen:
                    land_type_cols.append(col)
            for col in h_gen.keys():
                if col.startswith('토지유형') and col not in land_type_cols:
                    land_type_cols.append(col)

            land_sum = 0.0
            land_details = []
            for col_name in land_type_cols:
                col_idx = h_gen.get(col_name)
                val_raw = ws_gen.cell(2, col_idx).value if col_idx else None
                val_num = to_float(val_raw) or 0.0
                land_sum += val_num
                if val_num > 0:
                    land_details.append(f"{col_name}: {val_num}")

            if round(land_sum, 1) != 100.0:
                self.add_error(
                    sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                    rule_name="[일반토지 6] 토지유형 합계 100 불일치",
                    error_field=f"토지유형({len(land_type_cols)}개) 합계",
                    message=f"토지유형 {len(land_type_cols)}개 항목의 합계는 반드시 100이어야 하지만, 현재 합계는 {land_sum:.1f}입니다.",
                    detail_val=f"합계={land_sum:.1f} ({', '.join(land_details) if land_details else '모두 0 또는 비어있음'})"
                )

            # [규칙 1-7] 소속 열 검증 ("산림조합중앙회 산림자원조사본부" 필수)
            dept_c = h_gen.get('소속')
            dept_raw = ws_gen.cell(2, dept_c).value if dept_c else None
            dept_val = clean_str(dept_raw)
            # 공백 정규화 후 비교 (띄어쓰기 오차 허용)
            if dept_val.replace(' ', '') != '산림조합중앙회산림자원조사본부':
                self.add_error(
                    sample_no=sample_no, fname=fname, sheet="일반·토지현황조사표", row_no=2,
                    rule_name="[일반토지 7] 소속 표기 오류",
                    error_field="소속",
                    message="일반·토지현황조사표의 '소속' 열은 반드시 '산림조합중앙회 산림자원조사본부'로 입력되어야 합니다.",
                    detail_val=dept_raw
                )

        # =============================================================
        # 2. 임목조사표(임목조사) 탭 데이터 사전 분석 (임분현황 상층 수종 및 수고 검증용)
        # =============================================================
        has_tree_sheet = '임목조사표(임목조사)' in wb.sheetnames
        tree_records = []       # [(수종명, 수고)]
        tree_heights = []       # [수고]
        max_tree_height_cm = None
        max_tree_species = set()  # 최대 수고를 가진 수종명 집합

        if has_tree_sheet:
            ws_tree = wb['임목조사표(임목조사)']
            h_tree = get_sheet_header_map(ws_tree)
            sp_c = h_tree.get('수종명')
            th_c = h_tree.get('수고(cm)') or h_tree.get('수고')

            for r in range(2, ws_tree.max_row + 1):
                num_v = ws_tree.cell(r, h_tree.get('번호', 2)).value
                sp_v = ws_tree.cell(r, sp_c).value if sp_c else None
                if num_v is None and sp_v is None:
                    continue

                sp_str = clean_str(sp_v)
                # 덩굴류(칡, 담쟁이덩굴 등)는 임목조사 최대 수고 수종에서 제외
                if is_vine_species(sp_str):
                    continue
                th_val = to_float(ws_tree.cell(r, th_c).value) if th_c else None

                if th_val is not None and th_val > 0:
                    tree_heights.append(th_val)
                    if sp_str:
                        tree_records.append((sp_str, th_val))

            if tree_heights:
                max_tree_height_cm = max(tree_heights)
                # 최대 수고(cm)를 가진 모든 수종명 수집 (동률 수종 허용)
                max_tree_species = {
                    sp for sp, h in tree_records
                    if abs(h - max_tree_height_cm) < 1e-6
                }

        # =============================================================
        # 3. 하층임목조사표(임목조사 & 군락조사) 수종명 및 수고 수집 (하층 최대수고 수종 검증용)
        # =============================================================
        under_records = []      # [(수종명, 수고)]
        under_heights = []      # [수고]
        max_under_height = None
        max_under_species = set()  # 최대 수고를 가진 수종명 집합

        if '하층임목조사표(임목조사)' in wb.sheetnames:
            ws_under_tree = wb['하층임목조사표(임목조사)']
            h_ut = get_sheet_header_map(ws_under_tree)
            sp_c = h_ut.get('수종명')
            th_c = h_ut.get('평균수고') or h_ut.get('수고')
            for r in range(2, ws_under_tree.max_row + 1):
                pid_v = ws_under_tree.cell(r, h_ut.get('표본점번호', 1)).value
                num_v = ws_under_tree.cell(r, h_ut.get('번호', 2)).value
                sp_v = ws_under_tree.cell(r, sp_c).value if sp_c else None
                if any(v is not None for v in [pid_v, num_v, sp_v]):
                    sp_str = clean_str(sp_v)
                    # 덩굴류(칡, 담쟁이덩굴 등)는 수고가 아무리 높아도 상층/중층/하층 층위 구간 및 최대수고 수종에서 제외
                    if is_vine_species(sp_str):
                        continue
                    th_val = to_float(ws_under_tree.cell(r, th_c).value) if th_c else None
                    if th_val is not None and th_val > 0:
                        under_heights.append(th_val)
                        if sp_str:
                            under_records.append((sp_str, th_val))

        if '하층임목조사표(군락조사)' in wb.sheetnames:
            ws_under_comm = wb['하층임목조사표(군락조사)']
            h_uc = get_sheet_header_map(ws_under_comm)
            sp_c = h_uc.get('수종명')
            th_c = h_uc.get('평균수고') or h_uc.get('수고')
            for r in range(2, ws_under_comm.max_row + 1):
                pid_v = ws_under_comm.cell(r, h_uc.get('표본점번호', 1)).value
                num_v = ws_under_comm.cell(r, h_uc.get('번호', 2)).value
                sp_v = ws_under_comm.cell(r, sp_c).value if sp_c else None
                if any(v is not None for v in [pid_v, num_v, sp_v]):
                    sp_str = clean_str(sp_v)
                    # 덩굴류(칡, 담쟁이덩굴 등)는 수고가 아무리 높아도 상층/중층/하층 층위 구간 및 최대수고 수종에서 제외
                    if is_vine_species(sp_str):
                        continue
                    th_val = to_float(ws_under_comm.cell(r, th_c).value) if th_c else None
                    if th_val is not None and th_val > 0:
                        under_heights.append(th_val)
                        if sp_str:
                            under_records.append((sp_str, th_val))

        if under_heights:
            max_under_height = max(under_heights)
            # 하층 조사에서 최대 수고를 가진 모든 수종명 수집 (동률 수종 허용)
            max_under_species = {
                sp for sp, h in under_records
                if abs(h - max_under_height) < 1e-6
            }

        # =============================================================
        # 4. 임분현황조사표 탭 검증
        # =============================================================
        if '임분현황조사표' not in wb.sheetnames:
            self.add_error(
                sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=None,
                rule_name="시트 누락", error_field="임분현황조사표",
                message="'임분현황조사표' 탭이 파일에 존재하지 않습니다.", detail_val=None
            )
        else:
            ws_stand = wb['임분현황조사표']
            h_stand = get_sheet_header_map(ws_stand)

            # [규칙 2-1] 도로거리 정수 검증
            rd_c = h_stand.get('도로거리')
            rd_raw = ws_stand.cell(2, rd_c).value if rd_c else None
            if not is_valid_integer(rd_raw):
                self.add_error(
                    sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                    rule_name="[임분현황 1] 도로거리 정수 오류",
                    error_field="도로거리",
                    message="도로거리는 반드시 정수로 입력되어야 합니다. (빈 공간, 실수, 문자 불가)",
                    detail_val=rd_raw
                )

            # [규칙 2-2] 표고 정수 및 기준 엑셀 표고와 10 이상 차이 검증
            elev_c = h_stand.get('표고')
            elev_raw = ws_stand.cell(2, elev_c).value if elev_c else None
            if not is_valid_integer(elev_raw):
                self.add_error(
                    sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                    rule_name="[임분현황 2] 표고 정수 오류",
                    error_field="표고",
                    message="표고는 반드시 정수로 입력되어야 합니다.",
                    detail_val=elev_raw
                )
            else:
                elev_num = to_float(elev_raw)
                ref_info = self.ref_loader.sample_dict.get(sample_no)
                if ref_info and ref_info.get('표고') is not None:
                    ref_elev = ref_info['표고']
                    elev_diff = abs(elev_num - ref_elev)
                    if elev_diff >= 10.0:
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                            rule_name="[임분현황 2] 표고 오차(10m 이상) 오류",
                            error_field="표고",
                            message=f"조사 표고({elev_num:.0f}m)가 기준 표본점목록의 표고({ref_elev:.0f}m)와 10m 이상({elev_diff:.1f}m) 차이납니다.",
                            detail_val=f"조사={elev_num:.0f}m vs 기준={ref_elev:.0f}m (차이: {elev_diff:.1f}m)"
                        )

            # [규칙 2-3] 구조물 부재 시 특이사항 검증
            # 표본점유형이 "5"가 아닌 경우: 구조물1, 구조물2, 구조물3의 내용이 모두 없다면 특이사항 열에 "구조물없음" 필수
            # (표본점유형이 "5"인 경우에는 구조물 및 거리, 방위각에 내용이 없어도 오류가 아님)
            if sample_type_val != "5":
                s1_val = clean_str(ws_stand.cell(2, h_stand.get('구조물1')).value) if h_stand.get('구조물1') else ''
                s2_val = clean_str(ws_stand.cell(2, h_stand.get('구조물2')).value) if h_stand.get('구조물2') else ''
                s3_val = clean_str(ws_stand.cell(2, h_stand.get('구조물3')).value) if h_stand.get('구조물3') else ''

                if not s1_val and not s2_val and not s3_val:
                    note_c = h_stand.get('특이사항')
                    note_val = clean_str(ws_stand.cell(2, note_c).value) if note_c else ''
                    # 띄어쓰기 차이 허용 ("구조물 없음" 등)
                    if '구조물없음' not in note_val.replace(' ', ''):
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                            rule_name="[임분현황] 구조물 부재 시 특이사항 기재 오류",
                            error_field="특이사항",
                            message="구조물1, 구조물2, 구조물3 열에 입력된 내용이 모두 없으므로, 특이사항 열에 반드시 '구조물없음'이 입력되어 있어야 합니다.",
                            detail_val=f"특이사항='{note_val}'"
                        )

            # [규칙 2-4] 구조물1, 2, 3 거리 및 방위각 검증 (내용이 입력된 경우)
            for i in [1, 2, 3]:
                st_c = h_stand.get(f'구조물{i}')
                if not st_c:
                    continue

                dist_c = st_c + 1
                az_c = st_c + 2

                if clean_str(ws_stand.cell(1, dist_c).value) != f'거리{i}':
                    dist_c = h_stand.get(f'거리{i}')
                if clean_str(ws_stand.cell(1, az_c).value) != f'방위각{i}':
                    az_c = h_stand.get(f'방위각{i}')

                st_val = clean_str(ws_stand.cell(2, st_c).value) if st_c else ''
                if st_val:  # 구조물 내용이 입력되어 있을 경우
                    # 거리 검증 (소수점 첫째자리까지 유효 필수)
                    dist_raw = ws_stand.cell(2, dist_c).value if dist_c else None
                    if not is_within_one_decimal(dist_raw):
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                            rule_name="[임분현황 3] 구조물 거리 소수점 오류",
                            error_field=f"거리{i}",
                            message=f"구조물{i}('{st_val}')이 입력되어 있으므로, 거리{i}는 소수점 첫째자리까지 정확히 입력되어야 합니다.",
                            detail_val=dist_raw
                        )

                    # 방위각 검증 (0~359 범위의 정수 필수)
                    az_raw = ws_stand.cell(2, az_c).value if az_c else None
                    if not is_valid_integer(az_raw):
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                            rule_name="[임분현황 4] 구조물 방위각 정수 오류",
                            error_field=f"방위각{i}",
                            message=f"구조물{i}('{st_val}')이 입력되어 있으므로, 방위각{i}는 0~359 범위의 정수로 입력되어야 합니다.",
                            detail_val=az_raw
                        )
                    else:
                        az_int = int(to_float(az_raw))
                        if not (0 <= az_int <= 359):
                            self.add_error(
                                sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                                rule_name="[임분현황 4] 구조물 방위각 범위 오류",
                                error_field=f"방위각{i}",
                                message=f"방위각{i}의 값({az_int})이 유효 범위(0~359º)를 벗어났습니다.",
                                detail_val=az_raw
                            )

            # [규칙 2-5] 수고 200cm 기준 상층 / 하층 필수 기재 검증
            # ※ 중층(종명, 피도, 범위)은 사용자가 자율적으로 결정하므로 필수 여부를 강제하지 않음
            top_sp_raw = ws_stand.cell(2, h_stand.get('상층(종명)')).value if h_stand.get('상층(종명)') else None
            top_cov_raw = ws_stand.cell(2, h_stand.get('상층피도(퍼센트)')).value if h_stand.get('상층피도(퍼센트)') else None
            top_rng_raw = ws_stand.cell(2, h_stand.get('상층(범위)')).value if h_stand.get('상층(범위)') else None

            bot_sp_raw = ws_stand.cell(2, h_stand.get('하층(종명)')).value if h_stand.get('하층(종명)') else None
            bot_cov_raw = ws_stand.cell(2, h_stand.get('하층피도(퍼센트)')).value if h_stand.get('하층피도(퍼센트)') else None
            bot_rng_raw = ws_stand.cell(2, h_stand.get('하층(범위)')).value if h_stand.get('하층(범위)') else None

            has_top_all = bool(clean_str(top_sp_raw) and to_float(top_cov_raw) is not None and clean_str(top_rng_raw))
            has_bot_all = bool(clean_str(bot_sp_raw) and to_float(bot_cov_raw) is not None and clean_str(bot_rng_raw))

            if max_tree_height_cm is not None and max_tree_height_cm >= 200:
                # 200cm 이상: 상층 필수
                if not has_top_all:
                    missing_fields = []
                    if not clean_str(top_sp_raw): missing_fields.append('상층(종명)')
                    if to_float(top_cov_raw) is None: missing_fields.append('상층피도(퍼센트)')
                    if not clean_str(top_rng_raw): missing_fields.append('상층(범위)')
                    self.add_error(
                        sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                        rule_name="[임분현황 5] 수고 200 이상 상층 기재 누락",
                        error_field="상층 조사 항목",
                        message=f"임목조사의 수고 최댓값이 {max_tree_height_cm:.0f}cm(>=200cm)이므로 상층 항목({', '.join(missing_fields)})에 값이 입력되어야 합니다.",
                        detail_val=f"수고 최댓값={max_tree_height_cm:.0f}cm, 누락항목={missing_fields}"
                    )
            else:
                # 200cm 이하이거나 임목조사가 없는 경우: 하층 필수
                if not has_bot_all:
                    missing_fields = []
                    if not clean_str(bot_sp_raw): missing_fields.append('하층(종명)')
                    if to_float(bot_cov_raw) is None: missing_fields.append('하층피도(퍼센트)')
                    if not clean_str(bot_rng_raw): missing_fields.append('하층(범위)')
                    self.add_error(
                        sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                        rule_name="[임분현황 5] 수고 200 이하 하층 기재 누락",
                        error_field="하층 조사 항목",
                        message=f"임목조사 수고가 모두 200cm 이하(또는 부재)이므로 하층으로 판단되며, 하층 항목({', '.join(missing_fields)})에 값이 입력되어야 합니다.",
                        detail_val=f"수고 최댓값={max_tree_height_cm}, 누락항목={missing_fields}"
                    )

            # [규칙 2-6] 상층(종명) 최대수고 수종 검증
            # 임목조사표(임목조사)에서 수고(cm)가 가장 큰 값의 수종명이 입력되어야 함
            if max_tree_species and max_tree_height_cm is not None:
                cur_top_sp = clean_str(top_sp_raw)
                if cur_top_sp and cur_top_sp not in max_tree_species:
                    allowed_str = ', '.join(sorted(max_tree_species))
                    self.add_error(
                        sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                        rule_name="[임분현황 6] 상층(종명) 최대수고 수종 불일치",
                        error_field="상층(종명)",
                        message=f"상층(종명)은 임목조사표에서 수고(cm)가 가장 큰 수종명('{allowed_str}', {max_tree_height_cm:.0f}cm)으로 입력되어야 합니다.",
                        detail_val=f"입력값='{cur_top_sp}' != 최대수고 수종={sorted(list(max_tree_species))}({max_tree_height_cm:.0f}cm)"
                    )

            # [규칙 2-7] 하층(종명) 최대수고 수종 검증
            # 하층임목조사표(임목조사) 및 하층임목조사표(군락조사)에서 수고가 가장 큰 값의 수종명이 입력되어야 함 (덩굴류 제외)
            if max_under_species and max_under_height is not None:
                cur_bot_sp = clean_str(bot_sp_raw)
                if cur_bot_sp and cur_bot_sp not in max_under_species:
                    allowed_u_str = ', '.join(sorted(max_under_species))
                    self.add_error(
                        sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                        rule_name="[임분현황 7] 하층(종명) 최대수고 수종 불일치",
                        error_field="하층(종명)",
                        message=f"하층(종명)은 하층임목조사표(임목/군락)에서 수고가 가장 큰 수종명('{allowed_u_str}', {max_under_height:.0f})으로 입력되어야 합니다.",
                        detail_val=f"입력값='{cur_bot_sp}' != 최대수고 수종={sorted(list(max_under_species))}({max_under_height:.0f})"
                    )

            # [규칙 2-8] 임분현황 층위(상층, 중층, 하층) 덩굴류 기재 오류 검증
            # 덩굴류(칡, 담쟁이덩굴 등)는 수고가 아무리 높아도 임분현황조사표의 상층, 중층, 하층 층위 구간(종명)에 입력 불가
            layer_sp_checks = [
                ('상층', top_sp_raw),
                ('중층', ws_stand.cell(2, h_stand.get('중층(종명)')).value if h_stand.get('중층(종명)') else None),
                ('하층', bot_sp_raw)
            ]
            for layer_name, sp_val in layer_sp_checks:
                sp_clean = clean_str(sp_val)
                if sp_clean and is_vine_species(sp_clean):
                    self.add_error(
                        sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                        rule_name="[임분현황 8] 덩굴류 수종 층위 기재 오류",
                        error_field=f"{layer_name}(종명)",
                        message=f"임분현황조사표의 {layer_name}(종명)에 덩굴류('{sp_clean}')가 입력되었습니다. 덩굴류(칡, 담쟁이덩굴 등)는 수고와 무관하게 상층, 중층, 하층의 구간에 입력될 수 없습니다.",
                        detail_val=sp_clean
                    )

            # ---------------------------------------------------------
            # 기존 규칙: 교란(인위적), 피도 최댓값, 범위 연결성
            # ---------------------------------------------------------
            disturb_c = h_stand.get('교란(인위적)')
            disturb_val = clean_str(ws_stand.cell(2, disturb_c).value) if disturb_c else ''
            if not disturb_val:
                self.add_error(
                    sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                    rule_name="[임분현황 기존] 교란(인위적) 누락",
                    error_field="교란(인위적)",
                    message="임분현황조사표의 '교란(인위적)' 열에 값이 입력되지 않았습니다.",
                    detail_val=ws_stand.cell(2, disturb_c).value if disturb_c else None
                )

            # 전체피도 최댓값 일치 검증
            top_cov_num = to_float(top_cov_raw)
            mid_cov_num = to_float(ws_stand.cell(2, h_stand.get('중층피도(퍼센트)')).value if h_stand.get('중층피도(퍼센트)') else None)
            bot_cov_num = to_float(bot_cov_raw)
            tot_cov_num = to_float(ws_stand.cell(2, h_stand.get('전체피도(퍼센트)')).value if h_stand.get('전체피도(퍼센트)') else None)

            cov_candidates = [c for c in [top_cov_num, mid_cov_num, bot_cov_num] if c is not None]
            if cov_candidates and tot_cov_num is not None:
                exp_max = max(cov_candidates)
                if round(tot_cov_num, 1) != round(exp_max, 1):
                    self.add_error(
                        sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                        rule_name="[임분현황 기존] 전체피도 불일치",
                        error_field="전체피도(퍼센트)",
                        message=f"전체피도({tot_cov_num:.1f}%)는 각 층의 피도 중 최댓값인 {exp_max:.1f}%와 같아야 합니다.",
                        detail_val=f"입력 전체피도={tot_cov_num}%, 상층={top_cov_num}%, 중층={mid_cov_num}%, 하층={bot_cov_num}%"
                    )

            # 층별 범위 연결성 검증
            top_parsed = parse_range(top_rng_raw)
            mid_rng_raw = ws_stand.cell(2, h_stand.get('중층(범위)')).value if h_stand.get('중층(범위)') else None
            mid_parsed = parse_range(mid_rng_raw)
            bot_parsed = parse_range(bot_rng_raw)

            if isinstance(top_parsed, tuple) and isinstance(bot_parsed, tuple):
                if isinstance(mid_parsed, tuple):
                    if round(top_parsed[1], 1) != round(mid_parsed[0], 1):
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                            rule_name="[임분현황 기존] 범위 연결성 오류",
                            error_field="상층-중층 범위",
                            message=f"상층의 끝값({top_parsed[1]})과 중층의 시작값({mid_parsed[0]})이 연결되지 않습니다.",
                            detail_val=f"상층='{top_rng_raw}', 중층='{mid_rng_raw}'"
                        )
                    if round(mid_parsed[1], 1) != round(bot_parsed[0], 1):
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                            rule_name="[임분현황 기존] 범위 연결성 오류",
                            error_field="중층-하층 범위",
                            message=f"중층의 끝값({mid_parsed[1]})과 하층의 시작값({bot_parsed[0]})이 연결되지 않습니다.",
                            detail_val=f"중층='{mid_rng_raw}', 하층='{bot_rng_raw}'"
                        )
                elif mid_parsed is None:
                    if round(top_parsed[1], 1) != round(bot_parsed[0], 1):
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임분현황조사표", row_no=2,
                            rule_name="[임분현황 기존] 범위 연결성 오류",
                            error_field="상층-하층 범위",
                            message=f"중층 생략 상태에서 상층 끝값({top_parsed[1]})과 하층 시작값({bot_parsed[0]})이 연결되지 않습니다.",
                            detail_val=f"상층='{top_rng_raw}', 하층='{bot_rng_raw}'"
                        )

        # =============================================================
        # 5. 임목조사표(임목조사) 탭 개별 행 검증
        # =============================================================
        if has_tree_sheet:
            ws_tree = wb['임목조사표(임목조사)']
            h_tree = get_sheet_header_map(ws_tree)

            def_cols = ['줄기·근주결함', '가지결함', '잎결함']
            int_cols = ['평균흉고직경', '본수', '수고(cm)', '방위각(º)']

            for r in range(2, ws_tree.max_row + 1):
                num_v = ws_tree.cell(r, h_tree.get('번호', 2)).value
                sp_v = ws_tree.cell(r, h_tree.get('수종명', 3)).value
                if num_v is None and sp_v is None:
                    continue

                row_desc = f"행 {r} (번호: {num_v}, 수종: {clean_str(sp_v)})"

                # [규칙 3-1] 결함 열: 값이 있을 경우 정수 필수 (빈공간 허용)
                for dc in def_cols:
                    col_idx = h_tree.get(dc)
                    if col_idx:
                        raw_v = ws_tree.cell(r, col_idx).value
                        if raw_v is not None and clean_str(raw_v) != '':
                            if not is_valid_integer(raw_v):
                                self.add_error(
                                    sample_no=sample_no, fname=fname, sheet="임목조사표(임목조사)", row_no=r,
                                    rule_name="[임목조사 1] 결함 열 정수 오류",
                                    error_field=dc,
                                    message=f"{row_desc}: '{dc}'에 값이 입력된 경우 반드시 정수여야 합니다.",
                                    detail_val=raw_v
                                )

                # [규칙 3-2 & 3-3] 거리(m), 방위각(º) 및 주요 측정값 검증
                # 표본점유형 열의 값이 "5"일 경우: 거리(m)와 방위각(º)은 반드시 빈 공간이어야 함!
                if sample_type_val == "5":
                    # 방위각(º) 빈 공간 검증 (값이 있으면 오류)
                    az_col = h_tree.get('방위각(º)')
                    az_raw = ws_tree.cell(r, az_col).value if az_col else None
                    if az_raw is not None and clean_str(az_raw) != '':
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임목조사표(임목조사)", row_no=r,
                            rule_name="[임목조사 2] 표본점유형 5 방위각 빈공간 오류",
                            error_field="방위각(º)",
                            message=f"{row_desc}: 표본점유형이 '5'인 경우 '방위각(º)'은 반드시 빈 공간이어야 하지만 값이 입력되어 있습니다.",
                            detail_val=az_raw
                        )

                    # 거리(m) 빈 공간 검증 (값이 있으면 오류)
                    dist_col = h_tree.get('거리(m)')
                    dist_raw = ws_tree.cell(r, dist_col).value if dist_col else None
                    if dist_raw is not None and clean_str(dist_raw) != '':
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임목조사표(임목조사)", row_no=r,
                            rule_name="[임목조사 3] 표본점유형 5 거리 빈공간 오류",
                            error_field="거리(m)",
                            message=f"{row_desc}: 표본점유형이 '5'인 경우 '거리(m)'는 반드시 빈 공간이어야 하지만 값이 입력되어 있습니다.",
                            detail_val=dist_raw
                        )

                    # 나머지 주요 측정값(평균흉고직경, 본수, 수고(cm))은 정수 필수
                    for ic in ['평균흉고직경', '본수', '수고(cm)']:
                        col_idx = h_tree.get(ic)
                        raw_v = ws_tree.cell(r, col_idx).value if col_idx else None
                        if not is_valid_integer(raw_v):
                            self.add_error(
                                sample_no=sample_no, fname=fname, sheet="임목조사표(임목조사)", row_no=r,
                                rule_name="[임목조사 2] 측정값 정수 오류",
                                error_field=ic,
                                message=f"{row_desc}: '{ic}'는 반드시 정수로 입력되어야 합니다. (빈 공간 불가)",
                                detail_val=raw_v
                            )
                else:
                    # 표본점유형이 "5"가 아닌 경우:
                    # 평균흉고직경, 본수, 수고(cm), 방위각(º) 정수 필수 (빈공간 불가)
                    for ic in ['평균흉고직경', '본수', '수고(cm)', '방위각(º)']:
                        col_idx = h_tree.get(ic)
                        raw_v = ws_tree.cell(r, col_idx).value if col_idx else None
                        if not is_valid_integer(raw_v):
                            self.add_error(
                                sample_no=sample_no, fname=fname, sheet="임목조사표(임목조사)", row_no=r,
                                rule_name="[임목조사 2] 측정값 정수 오류",
                                error_field=ic,
                                message=f"{row_desc}: '{ic}'는 반드시 정수로 입력되어야 합니다. (빈 공간 불가)",
                                detail_val=raw_v
                            )

                    # 거리(m): 정수 또는 소수점 1자리까지 필수 (빈공간 불가)
                    dist_c = h_tree.get('거리(m)')
                    dist_raw = ws_tree.cell(r, dist_c).value if dist_c else None
                    if not is_within_one_decimal(dist_raw):
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="임목조사표(임목조사)", row_no=r,
                            rule_name="[임목조사 3] 거리(m) 소수점 오류",
                            error_field="거리(m)",
                            message=f"{row_desc}: '거리(m)'는 정수 또는 소수점 1자리까지 입력되어야 합니다. (빈 공간 불가)",
                            detail_val=dist_raw
                        )

        # =============================================================
        # 6. 하층임목조사표(임목조사) 탭 개별 행 검증
        # =============================================================
        if '하층임목조사표(임목조사)' in wb.sheetnames:
            ws_ut = wb['하층임목조사표(임목조사)']
            h_ut = get_sheet_header_map(ws_ut)
            target_int_cols = ['근원직경', '본수', '평균수고']

            for r in range(2, ws_ut.max_row + 1):
                pid_v = ws_ut.cell(r, h_ut.get('표본점번호', 1)).value
                num_v = ws_ut.cell(r, h_ut.get('번호', 2)).value
                sp_v = ws_ut.cell(r, h_ut.get('수종명', 3)).value

                # [규칙 4-1] 표본점번호, 번호, 수종명 열에 값이 있을 경우:
                # 근원직경, 본수, 평균수고의 값은 정수 필수
                if any(v is not None and clean_str(v) != '' for v in [pid_v, num_v, sp_v]):
                    row_desc = f"행 {r} (수종: {clean_str(sp_v)})"
                    for col_name in target_int_cols:
                        col_idx = h_ut.get(col_name)
                        val_raw = ws_ut.cell(r, col_idx).value if col_idx else None
                        if not is_valid_integer(val_raw):
                            self.add_error(
                                sample_no=sample_no, fname=fname, sheet="하층임목조사표(임목조사)", row_no=r,
                                rule_name="[하층임목 1] 주요 수치 정수 오류",
                                error_field=col_name,
                                message=f"{row_desc}: '{col_name}' 열의 값은 반드시 정수로 입력되어야 합니다.",
                                detail_val=val_raw
                            )

        # =============================================================
        # 7. 하층임목조사표(군락조사) 탭 개별 행 검증
        # =============================================================
        if '하층임목조사표(군락조사)' in wb.sheetnames:
            ws_uc = wb['하층임목조사표(군락조사)']
            h_uc = get_sheet_header_map(ws_uc)

            for r in range(2, ws_uc.max_row + 1):
                pid_v = ws_uc.cell(r, h_uc.get('표본점번호', 1)).value
                num_v = ws_uc.cell(r, h_uc.get('번호', 2)).value
                sp_v = ws_uc.cell(r, h_uc.get('수종명', 3)).value

                if any(v is not None and clean_str(v) != '' for v in [pid_v, num_v, sp_v]):
                    row_desc = f"행 {r} (수종: {clean_str(sp_v)})"

                    # [규칙 5-1] 평균수고, 손실률 값은 정수 필수
                    for col_name in ['평균수고', '손실률']:
                        col_idx = h_uc.get(col_name)
                        val_raw = ws_uc.cell(r, col_idx).value if col_idx else None
                        if not is_valid_integer(val_raw):
                            self.add_error(
                                sample_no=sample_no, fname=fname, sheet="하층임목조사표(군락조사)", row_no=r,
                                rule_name="[하층군락 1] 수고/손실률 정수 오류",
                                error_field=col_name,
                                message=f"{row_desc}: '{col_name}' 열의 값은 반드시 정수로 입력되어야 합니다.",
                                detail_val=val_raw
                            )

                    # [규칙 5-2] 면적 열의 값이 소수점 1자리까지 필수
                    area_idx = h_uc.get('면적')
                    area_raw = ws_uc.cell(r, area_idx).value if area_idx else None
                    if not is_within_one_decimal(area_raw):
                        self.add_error(
                            sample_no=sample_no, fname=fname, sheet="하층임목조사표(군락조사)", row_no=r,
                            rule_name="[하층군락 2] 면적 소수점 오류",
                            error_field="면적",
                            message=f"{row_desc}: '면적' 열의 값은 소수점 1자리까지 정확히 입력되어야 합니다.",
                            detail_val=area_raw
                        )

        wb.close()
        file_errors_after = len(self.errors)
        file_err_count = file_errors_after - file_errors_before
        self.file_summary.append({
            "표본점번호": sample_no,
            "파일명": fname,
            "팀장": self.current_leader,
            "팀원": self.current_member,
            "오류건수": file_err_count,
            "상태": "오류 발견" if file_err_count > 0 else "정상 (통과)"
        })


# =============================================================
# 4. 검수 결과 엑셀 리포트 생성기
# =============================================================

def generate_excel_report(inspector, output_path):
    """
    검수 결과를 시각적으로 아름답게 정리된 엑셀 보고서로 생성합니다.
    """
    wb = openpyxl.Workbook()
    default_sheet = wb.active

    # 스타일 정의
    font_title = Font(name="맑은 고딕", size=15, bold=True, color="1E293B")
    font_sub = Font(name="맑은 고딕", size=10, color="64748B")
    font_header = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="맑은 고딕", size=9, color="1E293B")

    fill_header_primary = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")  # 네이비
    fill_header_secondary = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid") # 오션블루
    fill_pass = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")  # 연녹색
    fill_fail = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")  # 연빨강
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # 연회색

    border_thin = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')

    # -------------------------------------------------------------
    # 시트 1: 표본점별_요약
    # -------------------------------------------------------------
    ws_sum = wb.create_sheet(title="표본점별_검수요약")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.cell(1, 1, "산림 도시 조사 데이터 검수 요약").font = font_title
    ws_sum.cell(2, 1, f"검수일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 총 검수 파일: {len(inspector.file_summary)}개 | 총 오류 건수: {len(inspector.errors)}건").font = font_sub

    headers_sum = ["연번", "표본점번호", "파일명", "팀장", "팀원", "오류 건수", "검수 판정"]
    for col_idx, h in enumerate(headers_sum, 1):
        cell = ws_sum.cell(4, col_idx, h)
        cell.font = font_header
        cell.fill = fill_header_primary
        cell.alignment = align_center
        cell.border = border_thin

    # 표본점별 요약 정렬 (팀장 -> 팀원 -> 표본점번호)
    def summary_sort_key(item):
        leader = clean_str(item.get("팀장", ""))
        member = clean_str(item.get("팀원", ""))
        raw_pid = clean_str(item.get("표본점번호", ""))
        pid_num = int(raw_pid) if raw_pid.isdigit() else 99999999
        return (leader, member, pid_num, raw_pid)

    inspector.file_summary.sort(key=summary_sort_key)

    for row_idx, item in enumerate(inspector.file_summary, 5):
        raw_pid = clean_str(item["표본점번호"])
        pid_val = int(raw_pid) if raw_pid.isdigit() else raw_pid

        c1 = ws_sum.cell(row_idx, 1, row_idx - 4)
        c2 = ws_sum.cell(row_idx, 2, pid_val)
        c3 = ws_sum.cell(row_idx, 3, item["파일명"])
        c4 = ws_sum.cell(row_idx, 4, item.get("팀장", ""))
        c5 = ws_sum.cell(row_idx, 5, item.get("팀원", ""))
        c6 = ws_sum.cell(row_idx, 6, item["오류건수"])
        c7 = ws_sum.cell(row_idx, 7, item["상태"])

        if isinstance(pid_val, int):
            c2.number_format = '0'

        for c in [c1, c2, c3, c4, c5, c6, c7]:
            c.font = font_body
            c.border = border_thin
            c.alignment = align_center

        c3.alignment = align_left
        c6.alignment = align_right

        if item["오류건수"] > 0:
            c7.fill = fill_fail
            c7.font = Font(name="맑은 고딕", size=9, bold=True, color="991B1B")
        else:
            c7.fill = fill_pass
            c7.font = Font(name="맑은 고딕", size=9, bold=True, color="166534")

        if (row_idx % 2 == 0) and item["오류건수"] == 0:
            for c in [c1, c2, c3, c4, c5, c6, c7]:
                c.fill = fill_zebra

    # -------------------------------------------------------------
    # 시트 2: 오류_상세내역
    # -------------------------------------------------------------
    ws_err = wb.create_sheet(title="오류_상세내역")
    ws_err.views.sheetView[0].showGridLines = True

    ws_err.cell(1, 1, "발견된 오류 상세 리스트").font = font_title
    ws_err.cell(2, 1, "각 파일 및 시트, 행별로 발생한 구체적인 오류 항목과 원인 설명입니다.").font = font_sub

    headers_err = ["연번", "표본점번호", "파일명", "팀장", "팀원", "시트명", "행번호", "검증규칙", "오류항목", "오류내용 및 사유", "실제 입력값"]
    for col_idx, h in enumerate(headers_err, 1):
        cell = ws_err.cell(4, col_idx, h)
        cell.font = font_header
        cell.fill = fill_header_primary
        cell.alignment = align_center
        cell.border = border_thin

    # 팀장별 오류 내역 정렬 (팀장 -> 팀원 -> 표본점번호 -> 시트명 -> 행번호)
    def error_sort_key(item):
        leader = clean_str(item.get("팀장", ""))
        member = clean_str(item.get("팀원", ""))
        raw_pid = clean_str(item.get("표본점번호", ""))
        pid_num = int(raw_pid) if raw_pid.isdigit() else 99999999
        sheet = clean_str(item.get("시트명", ""))
        raw_row = str(item.get("행번호", ""))
        row_num = int(raw_row) if raw_row.isdigit() else 99999999
        return (leader, member, pid_num, raw_pid, sheet, row_num)

    inspector.errors.sort(key=error_sort_key)

    for row_idx, item in enumerate(inspector.errors, 5):
        raw_err_pid = clean_str(item["표본점번호"])
        err_pid_val = int(raw_err_pid) if raw_err_pid.isdigit() else raw_err_pid

        c1 = ws_err.cell(row_idx, 1, row_idx - 4)
        c2 = ws_err.cell(row_idx, 2, err_pid_val)
        c3 = ws_err.cell(row_idx, 3, item["파일명"])
        c4 = ws_err.cell(row_idx, 4, item.get("팀장", ""))
        c5 = ws_err.cell(row_idx, 5, item.get("팀원", ""))
        c6 = ws_err.cell(row_idx, 6, item["시트명"])
        c7 = ws_err.cell(row_idx, 7, item["행번호"])
        c8 = ws_err.cell(row_idx, 8, item["검증규칙"])
        c9 = ws_err.cell(row_idx, 9, item["오류항목"])
        c10 = ws_err.cell(row_idx, 10, item["오류내용"])
        c11 = ws_err.cell(row_idx, 11, item["입력값_상세"])

        if isinstance(err_pid_val, int):
            c2.number_format = '0'

        for c in [c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11]:
            c.font = font_body
            c.border = border_thin
            c.alignment = align_center

        c3.alignment = align_left
        c8.alignment = align_left
        c9.alignment = align_left
        c10.alignment = align_left
        c11.alignment = align_left

        if row_idx % 2 == 0:
            for c in [c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11]:
                c.fill = fill_zebra

    # -------------------------------------------------------------
    # 시트 3: 규칙별_발생통계
    # -------------------------------------------------------------
    ws_stat = wb.create_sheet(title="규칙별_발생통계")
    ws_stat.views.sheetView[0].showGridLines = True

    ws_stat.cell(1, 1, "검증 규칙별 오류 발생 통계").font = font_title
    ws_stat.cell(2, 1, "어떤 오류 유형이 가장 많이 발생했는지 빈도순으로 정렬된 통계입니다.").font = font_sub

    headers_stat = ["순위", "검증 규칙명", "발생 건수", "비율(%)"]
    for col_idx, h in enumerate(headers_stat, 1):
        cell = ws_stat.cell(4, col_idx, h)
        cell.font = font_header
        cell.fill = fill_header_secondary
        cell.alignment = align_center
        cell.border = border_thin

    total_err_count = max(len(inspector.errors), 1)
    sorted_stats = inspector.rule_stats.most_common()

    for row_idx, (rule_name, count) in enumerate(sorted_stats, 5):
        ratio = (count / total_err_count) * 100.0
        c1 = ws_stat.cell(row_idx, 1, row_idx - 4)
        c2 = ws_stat.cell(row_idx, 2, rule_name)
        c3 = ws_stat.cell(row_idx, 3, count)
        c4 = ws_stat.cell(row_idx, 4, f"{ratio:.1f}%")

        for c in [c1, c2, c3, c4]:
            c.font = font_body
            c.border = border_thin
            c.alignment = align_center

        c2.alignment = align_left
        c3.alignment = align_right
        c4.alignment = align_right

    wb.remove(default_sheet)

    # 열 너비 자동 조절
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2]:
                    continue
                val_str = str(cell.value or '')
                byte_len = sum(2 if ord(ch) > 127 else 1 for ch in val_str)
                if byte_len > max_len:
                    max_len = byte_len
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output_path)
    return output_path


# =============================================================
# 5. 메인 일괄 검수 실행 함수 (웹 앱 및 CLI 공용)
# =============================================================

def run_inspection(target_dir, ref_file="2026 도시 조사결과표(한성안).xlsx", output_dir=None, progress_callback=None):
    """
    지정된 디렉터리의 엑셀 파일들을 일괄 검수하고 결과 엑셀 파일을 다운로드 폴더에 생성합니다.
    """
    if not os.path.isdir(target_dir):
        raise ValueError(f"지정된 검수 대상 폴더가 존재하지 않습니다: {target_dir}")

    # 결과 파일 저장 폴더 결정 (기본값: Windows 사용자 다운로드 폴더)
    if not output_dir:
        default_downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
        output_dir = default_downloads if os.path.exists(default_downloads) else target_dir
    
    os.makedirs(output_dir, exist_ok=True)

    # 기준 파일 경로 절대경로화 또는 탐색
    if not os.path.isabs(ref_file):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        cand_path = os.path.join(base_dir, ref_file)
        if os.path.exists(cand_path):
            ref_file = cand_path

    # 1. 기준 데이터 로드
    ref_loader = ReferenceDataLoader(ref_file)
    if progress_callback:
        progress_callback(0, 100, "", "기준 데이터(2026 도시 조사결과표) 로딩 중...")
    
    if not ref_loader.load():
        raise RuntimeError(f"기준 데이터 로드 실패: {ref_loader.load_error}")

    # 2. 검수 대상 파일 목록 수집
    file_paths = sorted([
        f for f in glob.glob(os.path.join(target_dir, '*.xlsx'))
        if not os.path.basename(f).startswith('~$') and '오류_검수_결과' not in os.path.basename(f)
    ])

    total_files = len(file_paths)
    if total_files == 0:
        raise ValueError(f"해당 폴더에 검수할 .xlsx 파일이 없습니다: {target_dir}")

    inspector = SurveyInspector(ref_loader)

    # 3. 파일별 검수 진행
    for idx, fpath in enumerate(file_paths, 1):
        fname = os.path.basename(fpath)
        if progress_callback:
            progress_callback(idx, total_files, fname, f"파일 검수 중 ({idx}/{total_files}): {fname}")
        inspector.inspect_file(fpath)

    # 4. 결과 엑셀 리포트 생성 (다운로드 폴더에 저장)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_filename = f"오류_검수_결과_{timestamp}.xlsx"
    report_path = os.path.join(output_dir, report_filename)

    if progress_callback:
        progress_callback(total_files, total_files, report_filename, "검수 결과 엑셀 리포트 생성 중...")

    generate_excel_report(inspector, report_path)

    # 대표 고정 파일명으로도 다운로드 폴더에 복사 저장
    try:
        import shutil
        fixed_path_in_output = os.path.join(output_dir, "오류_검수_결과.xlsx")
        shutil.copy2(report_path, fixed_path_in_output)
    except Exception:
        pass

    summary_data = {
        "report_filename": report_filename,
        "report_path": os.path.abspath(report_path),
        "output_dir": os.path.abspath(output_dir),
        "total_files": total_files,
        "error_files": sum(1 for item in inspector.file_summary if item["오류건수"] > 0),
        "pass_files": sum(1 for item in inspector.file_summary if item["오류건수"] == 0),
        "total_errors": len(inspector.errors),
        "file_summary": inspector.file_summary,
        "rule_stats": dict(inspector.rule_stats),
        "errors": inspector.errors
    }
    return summary_data


if __name__ == "__main__":
    import sys
    print("산림 도시 조사 엑셀 데이터 오류 검수 엔진 단독 테스트")
    target = "excel_1-4"
    if len(sys.argv) > 1:
        target = sys.argv[1]
    
    def cli_progress(cur, tot, fname, msg):
        print(f"[{cur}/{tot}] {msg}")

    result = run_inspection(target, progress_callback=cli_progress)
    print("\n[검수 완료]")
    print(f"- 검수 파일: {result['total_files']}개")
    print(f"- 오류 파일: {result['error_files']}개")
    print(f"- 총 오류 건수: {result['total_errors']}건")
    print(f"- 결과 엑셀 저장됨: {result['report_path']}")
