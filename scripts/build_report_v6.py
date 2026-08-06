"""
build_report_v6.py

v5 -> v6 변경 사항 (2026-01~03월 데이터 + 신규 지표 정의 반영):

1. 매핑 실패(경과년수가 최대전환년수를 넘어서는 경우) 시, 매핑정보(Sheet2류)
   대신 최대전환매핑(Sheet5류)의 최대전환 플랜코드로 대체 조회.
2. 사고유: 연차별 전환구분값 캐스케이드(이번 연차=0, 또는 이번 연차 공백+직전
   연차=0)로 판정하되, 이미 전환여부=TRUE(현재플랜=경과전환목표)로 확인된 계약은
   사고 기록이 있어도 사고유에서 제외한다("사고정보보다 현재 플랜정보가 더
   정확하니 사후로 검증"). 이번 연차 직접사고(=0)는 전환여부=TRUE&년수비교=FALSE
   조건으로, 캐스케이드(공백+직전=0)는 전환여부=TRUE&년수비교=TRUE 조건으로 제외
   (후자는 "대상" 쪽 matured 제외와 중복 차감되지 않도록 조건을 맞춤).
3. 대상: 2연차부터는 "이미 최종목표까지 도달하고(전환여부=TRUE) 그 시점도
   지난(년수비교=TRUE) 계약"을 제외한다.
4. 전환완료: 사고유 여부와 무관하게 전환여부=TRUE & 년수비교=FALSE 전체.
5. 전환예정: 이번 연차 전환구분값=1 & 전환여부=FALSE. 단 전환처리구분코드
   (001~004, 007)가 있는 건은 코드별 항목과 겹치므로 제외(중복 방지).
6. 미활동 = 미전환 - (전환예정 + 코드별 합계). 뺄셈으로 산출해 남는 예외
   케이스까지 자동으로 흡수한다.
7. 그룹 기준 2종: 부문별(수금부문명, 정확매칭) + 상품별(상품명, 정확매칭).
   레이아웃 구조(세부보고/합산보고)는 두 그룹 기준에 동일하게 적용.
8. 전환사고구분값1~5, 전환처리구분코드는 텍스트로 저장되어 있으므로 COUNTIFS
   비교값은 전부 따옴표로 텍스트 지정("0","1","002" 등) - 안 그러면 LibreOffice
   재계산 시 매칭 실패(v5까지 없던 이슈, 이번 신규 템플릿에서 발견/수정).

사용법: python build_report_v6.py <입력.xlsx> <출력.xlsx>
입력 파일은 시트 3개를 담고 있어야 한다:
  - 데이터 시트: '수금부문명' 헤더 보유
  - 매핑정보 시트: '*무사고기간' 헤더 보유 (전환전대표플랜코드+무사고기간 -> 전환후대표플랜코드)
  - 최대전환매핑 시트: '최대전환년수' 헤더 보유 (전환전대표플랜코드 -> 최대전환 플랜코드/최대전환년수)
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

FONT_NAME = '맑은 고딕'
COUNT_FMT = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
RATIO_FMT = '_-* #,##0.0_-;-* #,##0.0_-;_-* "-"_-;_-@_-'

THIN = Side(style='thin')
MEDIUM = Side(style='medium')

TOTAL_FILL = PatternFill('solid', fgColor='00B0F0')
GROUP_FILL = PatternFill('solid', fgColor='FFFF00')
HEADER_FONT = Font(name=FONT_NAME, size=11)
LABEL_FONT = Font(name=FONT_NAME, size=10)

CODE_MAP = [(1, '001', '연락두절'), (2, '002', '사고있음'), (3, '003', '고객거부'),
            (4, '004', '압류계약'), (7, '007', 'ARS거부')]

TIERS = [1, 2, 3, 4, 5]


# ---------------------------------------------------------------------------
# 공통 헬퍼
# ---------------------------------------------------------------------------

def header_map(ws, max_col=None):
    max_col = max_col or ws.max_column
    return {ws.cell(row=1, column=c).value: c for c in range(1, max_col + 1)
            if ws.cell(row=1, column=c).value}


def find_all_columns(ws, name, max_col=None):
    max_col = max_col or ws.max_column
    return [c for c in range(1, max_col + 1) if ws.cell(row=1, column=c).value == name]


def identify_sheets(wb):
    data_ws = mapping_ws = maxmap_ws = None
    for ws in wb.worksheets:
        hmap = header_map(ws)
        if '수금부문명' in hmap:
            data_ws = ws
        elif '최대전환년수' in hmap:
            maxmap_ws = ws
        elif '*무사고기간' in hmap:
            mapping_ws = ws
    if not (data_ws and mapping_ws and maxmap_ws):
        raise ValueError("데이터/매핑정보/최대전환매핑 시트를 모두 찾지 못했습니다.")
    return data_ws, mapping_ws, maxmap_ws


def last_data_row(ws, key_col):
    """key_col(예: 계약체결월)이 연속으로 채워진 마지막 행을 찾는다."""
    last = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=key_col).value is not None:
            last = r
        elif r > last + 50:
            # 50행 이상 연속 공백이면 데이터 끝으로 간주
            break
    return last


def get_distinct_values(ws, col, max_row):
    seen = []
    seen_set = set()
    for r in range(2, max_row + 1):
        v = ws.cell(row=r, column=col).value
        if v is None:
            continue
        if v not in seen_set:
            seen_set.add(v)
            seen.append(v)
    return seen


# ---------------------------------------------------------------------------
# 1) 매핑정보 시트: 키 컬럼 추가
# ---------------------------------------------------------------------------

def process_mapping_sheet(mapping_ws):
    hmap = header_map(mapping_ws)
    col_before = hmap['*전환전 대표플랜코드']
    col_years = hmap['*무사고기간']
    col_after = hmap['전환후 대표플랜코드']
    last_row = mapping_ws.max_row
    key_col = mapping_ws.max_column + 1
    L = get_column_letter
    mapping_ws.cell(row=1, column=key_col).value = '키'
    for r in range(2, last_row + 1):
        if mapping_ws.cell(row=r, column=col_before).value is None:
            continue
        mapping_ws.cell(row=r, column=key_col).value = (
            f'={L(col_before)}{r}&{L(col_years)}{r}'
        )
    return {
        'before': col_before, 'years': col_years, 'after': col_after,
        'key': key_col, 'last_row': last_row,
    }


def process_maxmap_sheet(maxmap_ws):
    hmap = header_map(maxmap_ws)
    col_before = hmap['*전환전 대표플랜코드']
    col_maxyears = hmap['최대전환년수']
    col_maxplan = hmap['최대전환 플랜코드']
    last_row = maxmap_ws.max_row
    return {
        'before': col_before, 'maxyears': col_maxyears, 'maxplan': col_maxplan,
        'last_row': last_row,
    }


# ---------------------------------------------------------------------------
# 2) 데이터 시트: 컬럼 확정 + 파생 컬럼(경과전환/사고전환/전환여부/최대전환/년수비교) 채우기
# ---------------------------------------------------------------------------

def process_data_sheet(data_ws):
    hmap = header_map(data_ws)
    plan_cols = find_all_columns(data_ws, '판매플랜코드')
    if len(plan_cols) < 2:
        raise ValueError(f"'판매플랜코드' 헤더가 2개 있어야 하는데 {len(plan_cols)}개 발견됨.")
    col_current = plan_cols[0]
    col_initial = plan_cols[1]

    col_elapsed_years = hmap['경과년수']
    col_elapsed_plan = hmap['전환판매플랜코드']
    col_accident_free_years = hmap['무사고년수']
    col_v = [hmap[f'전환사고구분값{i}'] for i in range(1, 6)]
    col_process_code = hmap['전환처리구분코드']
    col_dept = hmap['수금부문명']
    col_product = hmap['상품명']
    col_month = hmap['계약체결월']

    max_col = data_ws.max_column
    col_as = max_col + 1          # 전환여부
    col_accident_plan = max_col + 2  # 사고전환판매플랜코드(참고용)
    col_maxplan = max_col + 3     # 최대전환플랜코드
    col_maxyears = max_col + 4    # 최대전환년수
    col_yc = max_col + 5          # 년수비교

    data_ws.cell(row=1, column=col_as).value = '전환여부'
    data_ws.cell(row=1, column=col_accident_plan).value = '사고전환판매플랜코드'
    data_ws.cell(row=1, column=col_maxplan).value = '최대전환플랜코드'
    data_ws.cell(row=1, column=col_maxyears).value = '최대전환년수'
    data_ws.cell(row=1, column=col_yc).value = '년수비교'

    return {
        'current': col_current, 'initial': col_initial,
        'elapsed_years': col_elapsed_years, 'elapsed_plan': col_elapsed_plan,
        'accident_free_years': col_accident_free_years, 'v': col_v,
        'process_code': col_process_code, 'dept': col_dept, 'product': col_product,
        'month': col_month,
        'as_flag': col_as, 'accident_plan': col_accident_plan,
        'maxplan': col_maxplan, 'maxyears': col_maxyears, 'yc': col_yc,
    }


def fill_data_formulas(data_ws, cols, mapping_name, map_info, maxmap_name, maxmap_info, last_row):
    L = get_column_letter
    K = L(cols['initial'])
    Ly = L(cols['elapsed_years'])
    N = L(cols['accident_free_years'])
    J = L(cols['current'])
    M = L(cols['elapsed_plan'])

    map_key_col = L(map_info['key'])
    map_after_col = L(map_info['after'])
    map_last = map_info['last_row']

    max_before_col = L(maxmap_info['before'])
    max_maxplan_col = L(maxmap_info['maxplan'])
    max_maxyears_col = L(maxmap_info['maxyears'])
    max_last = maxmap_info['last_row']

    U = L(cols['as_flag'])
    V = L(cols['accident_plan'])
    W = L(cols['maxplan'])
    X = L(cols['maxyears'])

    for r in range(2, last_row + 1):
        if data_ws.cell(row=r, column=cols['month']).value is None:
            continue

        data_ws.cell(row=r, column=cols['elapsed_plan']).value = (
            f'=IFERROR(INDEX({mapping_name}!${map_after_col}:${map_after_col},'
            f'MATCH({K}{r}&{Ly}{r},{mapping_name}!${map_key_col}:${map_key_col},0)),'
            f'IFERROR(INDEX({maxmap_name}!${max_maxplan_col}:${max_maxplan_col},'
            f'MATCH({K}{r},{maxmap_name}!${max_before_col}:${max_before_col},0)),""))'
        )
        data_ws.cell(row=r, column=cols['accident_plan']).value = (
            f'=IFERROR(INDEX({mapping_name}!${map_after_col}:${map_after_col},'
            f'MATCH({K}{r}&{N}{r},{mapping_name}!${map_key_col}:${map_key_col},0)),'
            f'IFERROR(INDEX({maxmap_name}!${max_maxplan_col}:${max_maxplan_col},'
            f'MATCH({K}{r},{maxmap_name}!${max_before_col}:${max_before_col},0)),""))'
        )
        data_ws.cell(row=r, column=cols['as_flag']).value = f'=IFERROR({J}{r}={M}{r},FALSE)'
        data_ws.cell(row=r, column=cols['maxplan']).value = (
            f'=IFERROR(INDEX({maxmap_name}!${max_maxplan_col}:${max_maxplan_col},'
            f'MATCH({K}{r},{maxmap_name}!${max_before_col}:${max_before_col},0)),"")'
        )
        data_ws.cell(row=r, column=cols['maxyears']).value = (
            f'=IFERROR(INDEX({maxmap_name}!${max_maxyears_col}:${max_maxyears_col},'
            f'MATCH({K}{r},{maxmap_name}!${max_before_col}:${max_before_col},0)),"")'
        )
        data_ws.cell(row=r, column=cols['yc']).value = f'=IFERROR({X}{r}<{Ly}{r},FALSE)'


# ---------------------------------------------------------------------------
# 3) 레이아웃(세부보고/합산보고) 생성
# ---------------------------------------------------------------------------

def q(v):
    """COUNTIFS 비교값을 텍스트로 명시(따옴표 처리, v5 이후 발견된 숫자/텍스트 오인 버그 방지)."""
    return f'"{v}"'


class DataRefs:
    """데이터 시트의 컬럼 문자(letter) 캐시."""

    def __init__(self, data_ws_name, cols):
        L = get_column_letter
        self.name = data_ws_name
        self.month = L(cols['month'])
        self.dept = L(cols['dept'])
        self.product = L(cols['product'])
        self.tier = L(cols['elapsed_years'])
        self.v = [L(c) for c in cols['v']]
        self.code = L(cols['process_code'])
        self.as_flag = L(cols['as_flag'])
        self.yc = L(cols['yc'])

    def rng(self, letter):
        return f'{self.name}!${letter}:${letter}'


def base_criteria(dr, month, group_letter, group_value, tier):
    return (
        f'{dr.rng(dr.month)},{q(month)},'
        f'{dr.rng(group_letter)},{q(group_value)},'
        f'{dr.rng(dr.tier)},{tier}'
    )


def formula_target(dr, month, group_letter, group_value, tier):
    """대상(E): tier1은 그대로, 2연차부터는 이미 matured(전환여부TRUE&년수비교TRUE)인
    계약(이번 연차값=1 또는 공백)을 뺀다."""
    bc = base_criteria(dr, month, group_letter, group_value, tier)
    if tier == 1:
        return f'=COUNTIFS({bc})'
    v_this = dr.v[tier - 1]
    return (
        f'=COUNTIFS({bc})'
        f'-COUNTIFS({bc},{dr.rng(v_this)},{q("1")},{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},TRUE)'
        f'-COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},TRUE)'
    )


def formula_accident(dr, month, group_letter, group_value, tier):
    """사고유(F): 이번 연차=0 이거나 (공백&직전연차=0). 단 이미 matured로 확인된
    건은 사고 기록이 있어도 제외(사후검증)."""
    bc = base_criteria(dr, month, group_letter, group_value, tier)
    v_this = dr.v[tier - 1]
    term1 = f'COUNTIFS({bc},{dr.rng(v_this)},{q("0")})'
    term1_sub = (
        f'COUNTIFS({bc},{dr.rng(v_this)},{q("0")},'
        f'{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},FALSE)'
    )
    if tier == 1:
        return f'={term1}-{term1_sub}'
    v_prev = dr.v[tier - 2]
    term3 = f'COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(v_prev)},{q("0")})'
    term3_sub = (
        f'COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(v_prev)},{q("0")},'
        f'{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},TRUE)'
    )
    return f'=({term1}-{term1_sub})+({term3}-{term3_sub})'


def formula_done(dr, month, group_letter, group_value, tier):
    """전환완료(H): 사고유 여부와 무관하게 전환여부=TRUE & 년수비교=FALSE 전체."""
    bc = base_criteria(dr, month, group_letter, group_value, tier)
    return f'=COUNTIFS({bc},{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},FALSE)'


def formula_plan(dr, month, group_letter, group_value, tier):
    """전환예정(L): 이번 연차=1 & 전환여부=FALSE, 단 코드(001~004,007)가 있는 건은
    코드별 항목과 겹치므로 제외."""
    bc = base_criteria(dr, month, group_letter, group_value, tier)
    v_this = dr.v[tier - 1]
    raw = f'COUNTIFS({bc},{dr.rng(v_this)},{q("1")},{dr.rng(dr.as_flag)},FALSE)'
    subs = ''.join(
        f'-COUNTIFS({bc},{dr.rng(v_this)},{q("1")},{dr.rng(dr.as_flag)},FALSE,{dr.rng(dr.code)},{q(cs)})'
        for _, cs, _ in CODE_MAP
    )
    return f'={raw}{subs}'


def formula_code(dr, month, group_letter, group_value, tier, code_str):
    """코드별(연락두절/사고있음/고객거부/압류계약/ARS거부): 이번연차=1 또는
    (공백, 단 사고유 캐스케이드(공백&직전=0)는 제외) & 전환여부=FALSE & 코드일치."""
    bc = base_criteria(dr, month, group_letter, group_value, tier)
    v_this = dr.v[tier - 1]
    c1 = f'COUNTIFS({bc},{dr.rng(v_this)},{q("1")},{dr.rng(dr.as_flag)},FALSE,{dr.rng(dr.code)},{q(code_str)})'
    if tier == 1:
        return f'={c1}'
    v_prev = dr.v[tier - 2]
    c2 = f'COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(dr.as_flag)},FALSE,{dr.rng(dr.code)},{q(code_str)})'
    c3 = (
        f'COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(v_prev)},{q("0")},'
        f'{dr.rng(dr.as_flag)},FALSE,{dr.rng(dr.code)},{q(code_str)})'
    )
    return f'={c1}+{c2}-{c3}'


DETAIL_COLS = {
    'month': 2, 'group': 3, 'tier': 4, 'target': 5, 'accident': 6, 'conv_target': 7,
    'done': 8, 'done_pct': 9, 'not_done': 10, 'not_done_pct': 11,
    'plan': 12, 'code1': 13, 'code2': 14, 'code3': 15, 'code4': 16, 'code5': 17,
    'activity_sum': 18, 'activity_pct': 19, 'idle': 20, 'idle_pct': 21,
}
DETAIL_LABELS = [
    '월', '구분', '연차', '대상', '사고유', '전환대상\n(A-B)', '전환완료',
    '전환률(%)', '미전환\n(C-D)', '미전환률(%)', '전환예정', '연락두절', '사고있음',
    '고객거부', '압류계약', 'ARS거부', '현장활동확인', '현장비율(%)', '미활동', '전체比(%)',
]


HEADER_FILL = PatternFill('solid', fgColor='FFFFFF')

# 3행짜리 병합헤더에서, (열그룹 시작~끝)에 상위 라벨을 걸고 그 아래 개별 라벨을
# 붙이는 열들. 그 외 열은 3행 전체를 세로 병합해서 라벨 하나만 건다.
DETAIL_GROUPED_HEADER = ('plan', 'code5', '현장활동확인', ['전환예정', '연락두절', '사고있음', '고객거부', '압류계약', 'ARS거부'])
DETAIL_SINGLE_LABELS = {
    'month': '월', 'group': '구분', 'tier': '연차', 'target': '대상', 'accident': '사고유',
    'conv_target': '전환대상\n(A-B)', 'done': '전환완료', 'done_pct': '전환률(%)',
    'not_done': '미전환\n(C-D)', 'not_done_pct': '미전환률(%)', 'activity_sum': '현장활동확인',
    'activity_pct': '현장비율(%)', 'idle': '미활동', 'idle_pct': '전체比(%)',
}


def _style_header_cell(c):
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def write_detail_header(ws, row):
    """템플릿처럼 3행에 걸친 병합 헤더를 만든다."""
    L = get_column_letter
    DC = DETAIL_COLS
    r0 = row
    grouped_start, grouped_end, grouped_label, grouped_leaves = DETAIL_GROUPED_HEADER
    c_start, c_end = DC[grouped_start], DC[grouped_end]

    for key, col in DC.items():
        if key in DETAIL_SINGLE_LABELS:
            ws.merge_cells(start_row=r0, start_column=col, end_row=r0 + 2, end_column=col)
            c = ws.cell(row=r0, column=col, value=DETAIL_SINGLE_LABELS[key])
            _style_header_cell(c)
            for rr in range(r0, r0 + 3):
                _style_header_cell(ws.cell(row=rr, column=col))

    ws.merge_cells(start_row=r0, start_column=c_start, end_row=r0 + 1, end_column=c_end)
    c = ws.cell(row=r0, column=c_start, value=grouped_label)
    _style_header_cell(c)
    for col in range(c_start, c_end + 1):
        for rr in (r0, r0 + 1):
            _style_header_cell(ws.cell(row=rr, column=col))
    for i, leaf in enumerate(grouped_leaves):
        c = ws.cell(row=r0 + 2, column=c_start + i, value=leaf)
        _style_header_cell(c)

    return row + 3


def write_detail_group(ws, row, dr, month, group_letter, group_value, tier_rows_out=None):
    """그룹 하나(연차 1~5)의 세부 행을 쓰고, 각 tier의 데이터 행 번호를 반환."""
    L = get_column_letter
    DC = DETAIL_COLS
    tier_rows = []
    for tier in TIERS:
        r = row
        ws.cell(row=r, column=DC['month'], value=month)
        ws.cell(row=r, column=DC['group'], value=group_value)
        ws.cell(row=r, column=DC['tier'], value=tier)
        ws.cell(row=r, column=DC['target']).value = formula_target(dr, month, group_letter, group_value, tier)
        ws.cell(row=r, column=DC['accident']).value = formula_accident(dr, month, group_letter, group_value, tier)
        E = f'{L(DC["target"])}{r}'
        F = f'{L(DC["accident"])}{r}'
        G = f'{L(DC["conv_target"])}{r}'
        H = f'{L(DC["done"])}{r}'
        J = f'{L(DC["not_done"])}{r}'
        L_ = f'{L(DC["plan"])}{r}'
        R = f'{L(DC["activity_sum"])}{r}'
        T = f'{L(DC["idle"])}{r}'
        ws.cell(row=r, column=DC['conv_target']).value = f'={E}-{F}'
        ws.cell(row=r, column=DC['done']).value = formula_done(dr, month, group_letter, group_value, tier)
        ws.cell(row=r, column=DC['done_pct']).value = f'=IFERROR({H}/{G}*100,0)'
        ws.cell(row=r, column=DC['not_done']).value = f'={G}-{H}'
        ws.cell(row=r, column=DC['not_done_pct']).value = f'=IFERROR({J}/{G}*100,0)'
        ws.cell(row=r, column=DC['plan']).value = formula_plan(dr, month, group_letter, group_value, tier)
        for i, (_, cs, _) in enumerate(CODE_MAP):
            ws.cell(row=r, column=DC['code1'] + i).value = formula_code(dr, month, group_letter, group_value, tier, cs)
        code1c = L(DC['code1']); code5c = L(DC['code5'])
        ws.cell(row=r, column=DC['activity_sum']).value = f'=SUM({L_}:{code5c}{r})'
        ws.cell(row=r, column=DC['activity_pct']).value = f'=IFERROR({R}/{G}*100,0)'
        ws.cell(row=r, column=DC['idle']).value = f'={J}-{R}'
        ws.cell(row=r, column=DC['idle_pct']).value = f'=IFERROR({T}/{G}*100,0)'
        tier_rows.append(r)
        row += 1
    if tier_rows_out is not None:
        tier_rows_out[group_value] = tier_rows
    return row


def write_detail_total(ws, row, dr, month, tier_rows_by_group):
    """월 합계: 그룹 소속과 무관하게 (연차만 조건으로) 독립 재계산."""
    L = get_column_letter
    DC = DETAIL_COLS
    all_group_letter = None  # 전체(그룹 조건 없음)이므로 base_criteria 대신 직접 작성
    for tier in TIERS:
        r = row
        ws.cell(row=r, column=DC['month'], value=month)
        ws.cell(row=r, column=DC['group'], value='합계')
        ws.cell(row=r, column=DC['tier'], value=tier)
        bc = f'{dr.rng(dr.month)},{q(month)},{dr.rng(dr.tier)},{tier}'
        v_this = dr.v[tier - 1]
        if tier == 1:
            target_f = f'=COUNTIFS({bc})'
        else:
            target_f = (
                f'=COUNTIFS({bc})'
                f'-COUNTIFS({bc},{dr.rng(v_this)},{q("1")},{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},TRUE)'
                f'-COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},TRUE)'
            )
        ws.cell(row=r, column=DC['target']).value = target_f
        term1 = f'COUNTIFS({bc},{dr.rng(v_this)},{q("0")})'
        term1_sub = f'COUNTIFS({bc},{dr.rng(v_this)},{q("0")},{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},FALSE)'
        if tier == 1:
            acc_f = f'=({term1}-{term1_sub})'
        else:
            v_prev = dr.v[tier - 2]
            term3 = f'COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(v_prev)},{q("0")})'
            term3_sub = (
                f'COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(v_prev)},{q("0")},'
                f'{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},TRUE)'
            )
            acc_f = f'=({term1}-{term1_sub})+({term3}-{term3_sub})'
        ws.cell(row=r, column=DC['accident']).value = acc_f
        E = f'{L(DC["target"])}{r}'; F = f'{L(DC["accident"])}{r}'
        G = f'{L(DC["conv_target"])}{r}'; H = f'{L(DC["done"])}{r}'
        J = f'{L(DC["not_done"])}{r}'; L_ = f'{L(DC["plan"])}{r}'
        R = f'{L(DC["activity_sum"])}{r}'; T = f'{L(DC["idle"])}{r}'
        ws.cell(row=r, column=DC['conv_target']).value = f'={E}-{F}'
        ws.cell(row=r, column=DC['done']).value = f'=COUNTIFS({bc},{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},FALSE)'
        ws.cell(row=r, column=DC['done_pct']).value = f'=IFERROR({H}/{G}*100,0)'
        ws.cell(row=r, column=DC['not_done']).value = f'={G}-{H}'
        ws.cell(row=r, column=DC['not_done_pct']).value = f'=IFERROR({J}/{G}*100,0)'
        raw = f'COUNTIFS({bc},{dr.rng(v_this)},{q("1")},{dr.rng(dr.as_flag)},FALSE)'
        subs = ''.join(
            f'-COUNTIFS({bc},{dr.rng(v_this)},{q("1")},{dr.rng(dr.as_flag)},FALSE,{dr.rng(dr.code)},{q(cs)})'
            for _, cs, _ in CODE_MAP
        )
        ws.cell(row=r, column=DC['plan']).value = f'={raw}{subs}'
        for i, (_, cs, _) in enumerate(CODE_MAP):
            c1 = f'COUNTIFS({bc},{dr.rng(v_this)},{q("1")},{dr.rng(dr.as_flag)},FALSE,{dr.rng(dr.code)},{q(cs)})'
            if tier == 1:
                code_f = f'={c1}'
            else:
                v_prev = dr.v[tier - 2]
                c2 = f'COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(dr.as_flag)},FALSE,{dr.rng(dr.code)},{q(cs)})'
                c3 = (
                    f'COUNTIFS({bc},{dr.rng(v_this)},{q(" ")},{dr.rng(v_prev)},{q("0")},'
                    f'{dr.rng(dr.as_flag)},FALSE,{dr.rng(dr.code)},{q(cs)})'
                )
                code_f = f'={c1}+{c2}-{c3}'
            ws.cell(row=r, column=DC['code1'] + i).value = code_f
        code5c = L(DC['code5'])
        ws.cell(row=r, column=DC['activity_sum']).value = f'=SUM({L_}:{code5c}{r})'
        ws.cell(row=r, column=DC['activity_pct']).value = f'=IFERROR({R}/{G}*100,0)'
        ws.cell(row=r, column=DC['idle']).value = f'={J}-{R}'
        ws.cell(row=r, column=DC['idle_pct']).value = f'=IFERROR({T}/{G}*100,0)'
        row += 1
    return row


def build_detail_sheet(wb, sheet_name, data_ws_name, cols, months, group_col_key, group_values):
    """세부보고 시트 하나(부문별 또는 상품별)를 월 블록으로 쌓아서 만든다.
    반환값: {month: {group_value: [tier1_row,...,tier5_row], '합계': [...]}}"""
    ws = wb.create_sheet(sheet_name)
    dr = DataRefs(data_ws_name, cols)
    group_letter = getattr(dr, group_col_key)

    row = 1
    row = write_detail_header(ws, row)
    row_map = {}
    for month in months:
        row_map[month] = {}
        month_start = row
        for gv in group_values:
            tier_rows_out = {}
            row = write_detail_group(ws, row, dr, month, group_letter, gv, tier_rows_out)
            row_map[month][gv] = tier_rows_out[gv]
        row = write_detail_total(ws, row, dr, month, row_map[month])
        row_map[month]['합계'] = list(range(row - 5, row))
        month_end = row - 1

        # 월 셀은 그 달 블록 전체(합계행 포함)에 걸쳐 병합
        L = get_column_letter
        month_col = DETAIL_COLS['month']
        ws.merge_cells(start_row=month_start, start_column=month_col,
                        end_row=month_end, end_column=month_col)
        ws.cell(row=month_start, column=month_col).alignment = Alignment(
            horizontal='center', vertical='center')

        # 그룹명 셀은 그 그룹의 연차 1~5행에 걸쳐 병합 + 노란색 강조
        group_col = DETAIL_COLS['group']
        for gv, tier_rows in row_map[month].items():
            r0, r1 = tier_rows[0], tier_rows[-1]
            ws.merge_cells(start_row=r0, start_column=group_col, end_row=r1, end_column=group_col)
            gcell = ws.cell(row=r0, column=group_col)
            gcell.alignment = Alignment(horizontal='center', vertical='center')
            gcell.fill = TOTAL_FILL if gv == '합계' else GROUP_FILL
    return ws, row_map


SUMMARY_COLS = {
    'month': 2, 'group': 3, 'target': 4, 'accident': 5, 'conv_target': 6,
    'done': 7, 'done_pct': 8, 'not_done': 9, 'not_done_pct': 10,
    'plan': 11, 'code1': 12, 'code2': 13, 'code3': 14, 'code4': 15, 'code5': 16,
    'activity_sum': 17, 'activity_pct': 18, 'idle': 19, 'idle_pct': 20,
}
SUMMARY_LABELS = [
    '월', '구분', '대상', '사고유', '전환대상\n(A-B)', '전환완료', '전환률(%)',
    '미전환\n(C-D)', '미전환률(%)', '전환예정', '연락두절', '사고있음', '고객거부',
    '압류계약', 'ARS거부', '현장활동확인', '현장비율(%)', '미활동', '전체比(%)',
]


SUMMARY_GROUPED_HEADER = ('plan', 'code5', '현장활동확인', ['전환예정', '연락두절', '사고있음', '고객거부', '압류계약', 'ARS거부'])
SUMMARY_SINGLE_LABELS = {
    'month': '월', 'group': '구분', 'target': '대상', 'accident': '사고유',
    'conv_target': '전환대상\n(A-B)', 'done': '전환완료', 'done_pct': '전환률(%)',
    'not_done': '미전환\n(C-D)', 'not_done_pct': '미전환률(%)', 'activity_sum': '현장활동확인',
    'activity_pct': '현장비율(%)', 'idle': '미활동', 'idle_pct': '전체比(%)',
}


def write_summary_header(ws, row):
    """세부보고와 동일한 3행 병합 헤더."""
    SC = SUMMARY_COLS
    r0 = row
    grouped_start, grouped_end, grouped_label, grouped_leaves = SUMMARY_GROUPED_HEADER
    c_start, c_end = SC[grouped_start], SC[grouped_end]

    for key, col in SC.items():
        if key in SUMMARY_SINGLE_LABELS:
            ws.merge_cells(start_row=r0, start_column=col, end_row=r0 + 2, end_column=col)
            c = ws.cell(row=r0, column=col, value=SUMMARY_SINGLE_LABELS[key])
            _style_header_cell(c)
            for rr in range(r0, r0 + 3):
                _style_header_cell(ws.cell(row=rr, column=col))

    ws.merge_cells(start_row=r0, start_column=c_start, end_row=r0 + 1, end_column=c_end)
    c = ws.cell(row=r0, column=c_start, value=grouped_label)
    _style_header_cell(c)
    for col in range(c_start, c_end + 1):
        for rr in (r0, r0 + 1):
            _style_header_cell(ws.cell(row=rr, column=col))
    for i, leaf in enumerate(grouped_leaves):
        c = ws.cell(row=r0 + 2, column=c_start + i, value=leaf)
        _style_header_cell(c)

    return row + 3


def write_summary_row(ws, row, month, group_value, detail_sheet_name, tier_rows):
    """detail_sheet의 tier1~5(또는 합계 5행) 행을 SUM해서 합산보고 한 줄을 만든다."""
    L = get_column_letter
    SC = SUMMARY_COLS
    DC = DETAIL_COLS
    r0, r1 = tier_rows[0], tier_rows[-1]

    def dcol(key):
        return f"'{detail_sheet_name}'!{L(DC[key])}{r0}:{L(DC[key])}{r1}"

    ws.cell(row=row, column=SC['month'], value=month)
    ws.cell(row=row, column=SC['group'], value=group_value)
    ws.cell(row=row, column=SC['target']).value = f'=SUM({dcol("target")})'
    ws.cell(row=row, column=SC['accident']).value = f'=SUM({dcol("accident")})'
    E = f'{L(SC["target"])}{row}'; F = f'{L(SC["accident"])}{row}'
    G = f'{L(SC["conv_target"])}{row}'; H = f'{L(SC["done"])}{row}'
    J = f'{L(SC["not_done"])}{row}'; Lp = f'{L(SC["plan"])}{row}'
    R = f'{L(SC["activity_sum"])}{row}'; T = f'{L(SC["idle"])}{row}'
    ws.cell(row=row, column=SC['conv_target']).value = f'={E}-{F}'
    ws.cell(row=row, column=SC['done']).value = f'=SUM({dcol("done")})'
    ws.cell(row=row, column=SC['done_pct']).value = f'=IFERROR({H}/{G}*100,0)'
    ws.cell(row=row, column=SC['not_done']).value = f'={G}-{H}'
    ws.cell(row=row, column=SC['not_done_pct']).value = f'=IFERROR({J}/{G}*100,0)'
    ws.cell(row=row, column=SC['plan']).value = f'=SUM({dcol("plan")})'
    for i in range(5):
        key = f'code{i + 1}'
        ws.cell(row=row, column=SC['code1'] + i).value = f'=SUM({dcol(key)})'
    ws.cell(row=row, column=SC['activity_sum']).value = f'=SUM({Lp}:{L(SC["code5"])}{row})'
    ws.cell(row=row, column=SC['activity_pct']).value = f'=IFERROR({R}/{G}*100,0)'
    ws.cell(row=row, column=SC['idle']).value = f'={J}-{R}'
    ws.cell(row=row, column=SC['idle_pct']).value = f'=IFERROR({T}/{G}*100,0)'
    return row + 1


def build_summary_sheet(wb, sheet_name, detail_sheet_name, months, group_values, row_map):
    ws = wb.create_sheet(sheet_name)
    row = 1
    row = write_summary_header(ws, row)
    month_col = SUMMARY_COLS['month']
    for month in months:
        month_start = row
        for gv in group_values:
            row = write_summary_row(ws, row, month, gv, detail_sheet_name, row_map[month][gv])
        row = write_summary_row(ws, row, month, '합계', detail_sheet_name, row_map[month]['합계'])
        month_end = row - 1
        ws.merge_cells(start_row=month_start, start_column=month_col,
                        end_row=month_end, end_column=month_col)
        ws.cell(row=month_start, column=month_col).alignment = Alignment(
            horizontal='center', vertical='center')
    return ws


# ---------------------------------------------------------------------------
# 4) 서식 적용
# ---------------------------------------------------------------------------

def style_detail_sheet(ws, row_map, header_rows=3):
    """row_map: build_detail_sheet가 반환한 {month: {group_value: [tier_rows...]}}.
    합계 행은 여기서 뽑아내 파란색으로, 그 외 그룹행은 노란 그룹명 셀 유지."""
    L = get_column_letter
    DC = DETAIL_COLS
    ncols = len(DC)
    pct_cols = {DC['done_pct'], DC['not_done_pct'], DC['activity_pct'], DC['idle_pct']}
    count_cols = {DC['target'], DC['accident'], DC['conv_target'], DC['done'], DC['not_done'],
                  DC['plan'], DC['code1'], DC['code2'], DC['code3'], DC['code4'], DC['code5'],
                  DC['activity_sum'], DC['idle']}
    total_rows = set()
    for month, groups in row_map.items():
        total_rows.update(groups['합계'])

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=1 + ncols):
        for cell in row:
            cell.font = Font(name=FONT_NAME, size=10)
            cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
            if cell.column in pct_cols:
                cell.number_format = RATIO_FMT
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif cell.column in count_cols:
                cell.number_format = COUNT_FMT
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if row[0].row in total_rows:
                cell.fill = TOTAL_FILL

    widths = {'B': 6, 'C': 14, 'D': 5, 'E': 9, 'F': 8, 'G': 9, 'H': 9, 'I': 8,
              'J': 9, 'K': 9, 'L': 8, 'M': 8, 'N': 8, 'O': 8, 'P': 8, 'Q': 8,
              'R': 9, 'S': 8, 'T': 8, 'U': 8}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.sheet_view.showGridLines = False


def style_summary_sheet(ws):
    L = get_column_letter
    SC = SUMMARY_COLS
    ncols = len(SC)
    pct_cols = {SC['done_pct'], SC['not_done_pct'], SC['activity_pct'], SC['idle_pct']}
    count_cols = {SC['target'], SC['accident'], SC['conv_target'], SC['done'], SC['not_done'],
                  SC['plan'], SC['code1'], SC['code2'], SC['code3'], SC['code4'], SC['code5'],
                  SC['activity_sum'], SC['idle']}

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=1 + ncols):
        for cell in row:
            cell.font = Font(name=FONT_NAME, size=10)
            cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
            if cell.column in pct_cols:
                cell.number_format = RATIO_FMT
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif cell.column in count_cols:
                cell.number_format = COUNT_FMT
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if row[0].row > 1 and ws.cell(row=row[0].row, column=SC['group']).value == '합계':
                cell.fill = TOTAL_FILL

    for r in range(2, ws.max_row + 1):
        gcell = ws.cell(row=r, column=SC['group'])
        if gcell.value not in (None, '합계'):
            gcell.fill = GROUP_FILL

    widths = {'B': 6, 'C': 14, 'D': 9, 'E': 8, 'F': 9, 'G': 9, 'H': 8,
              'I': 9, 'J': 9, 'K': 8, 'L': 8, 'M': 8, 'N': 8, 'O': 8, 'P': 8,
              'Q': 9, 'R': 8, 'S': 8, 'T': 8}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------------------
# 5) 메인
# ---------------------------------------------------------------------------

DEPT_ORDER = ['개인사업부문', '전략사업부문', '신사업부문', '법인사업부문']


def order_values(values, preferred):
    ordered = [v for v in preferred if v in values]
    ordered += [v for v in values if v not in ordered]
    return ordered


def build(input_path, output_path):
    print(f"[1/3] 입력 파일 분석 중... ({input_path})")
    wb = openpyxl.load_workbook(input_path, data_only=False)
    data_ws, mapping_ws, maxmap_ws = identify_sheets(wb)
    print(f"  - 데이터 시트: '{data_ws.title}'")
    print(f"  - 매핑정보 시트: '{mapping_ws.title}'")
    print(f"  - 최대전환매핑 시트: '{maxmap_ws.title}'")

    map_info = process_mapping_sheet(mapping_ws)
    maxmap_info = process_maxmap_sheet(maxmap_ws)
    cols = process_data_sheet(data_ws)

    d_last_row = last_data_row(data_ws, cols['month'])
    print(f"  - 데이터 마지막 행: {d_last_row}")
    fill_data_formulas(data_ws, cols, mapping_ws.title, map_info, maxmap_ws.title, maxmap_info, d_last_row)

    months = sorted(get_distinct_values(data_ws, cols['month'], d_last_row))
    depts = get_distinct_values(data_ws, cols['dept'], d_last_row)
    depts = order_values(depts, DEPT_ORDER)
    products = sorted(get_distinct_values(data_ws, cols['product'], d_last_row))
    print(f"  - 계약체결월: {months}")
    print(f"  - 수금부문명({len(depts)}종): {depts}")
    print(f"  - 상품명({len(products)}종): {products}")

    print("[2/3] 레이아웃(세부보고/합산보고 x 부문별/상품별) 생성 중...")
    ws_detail_dept, rows_dept = build_detail_sheet(
        wb, '세부보고_부문별', data_ws.title, cols, months, 'dept', depts)
    style_detail_sheet(ws_detail_dept, rows_dept)

    ws_detail_prod, rows_prod = build_detail_sheet(
        wb, '세부보고_상품별', data_ws.title, cols, months, 'product', products)
    style_detail_sheet(ws_detail_prod, rows_prod)

    ws_summary_dept = build_summary_sheet(
        wb, '합산보고_부문별', '세부보고_부문별', months, depts, rows_dept)
    style_summary_sheet(ws_summary_dept)

    ws_summary_prod = build_summary_sheet(
        wb, '합산보고_상품별', '세부보고_상품별', months, products, rows_prod)
    style_summary_sheet(ws_summary_prod)

    print(f"[3/3] 저장 중... -> {output_path}")
    wb.save(output_path)
    wb.close()
    print("완료.")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python build_report_v6.py <입력.xlsx> <출력.xlsx>")
        sys.exit(1)
    build(sys.argv[1], sys.argv[2])
