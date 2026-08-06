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
7. 그룹 기준 2종: 부문별(수금부문명, 정확매칭, 법인사업부문 제외) + 상품별
   (상품명, 정확매칭). 레이아웃 구조(세부보고/합산보고)는 두 그룹 기준에
   동일하게 적용.
8. 전환사고구분값1~5, 전환처리구분코드는 텍스트로 저장되어 있으므로 COUNTIFS
   비교값은 전부 따옴표로 텍스트 지정("0","1","002" 등) - 안 그러면 LibreOffice
   재계산 시 매칭 실패(v5까지 없던 이슈, 이번 신규 템플릿에서 발견/수정).
9. 레이아웃(세부보고/합산보고)은 사용자가 제공한 원본 템플릿의 셀 배치를 그대로
   따른다: 세부보고는 "전사 계"가 각 월 블록 맨 위에 오고 그 값은 부문별
   행들의 연차별 합(SUM)으로 계산되며(법인 제외 3개 부문만 합산), 부문명은
   전사계만 폭 넓은 병합(C:D), 개별 부문은 D열 하나만 병합해서 들여쓰기
   효과를 준다. 헤더는 2개월마다 반복. 합산보고는 전사계 행 없이 월 구분행
   (병합) 아래 부문별(개인/전략/신사업 축약 표기) 한 줄씩만 나온다.

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

# 원본 템플릿에서 추출한 색상: 전사계/총계=테마accent5(4472C4) tint 0.8,
# 부문명 강조=순수 노랑, 헤더=테마lt1(흰색) tint -0.05(연회색)
TOTAL_FILL = PatternFill('solid', fgColor='DAE3F3')
GROUP_FILL = PatternFill('solid', fgColor='FFFF00')
HEADER_FILL = PatternFill('solid', fgColor='F2F2F2')
HEADER_FONT = Font(name=FONT_NAME, size=11)
LABEL_FONT = Font(name=FONT_NAME, size=10)

CODE_MAP = [(1, '001', '연락두절'), (2, '002', '사고있음'), (3, '003', '고객거부'),
            (4, '004', '압류계약'), (7, '007', 'ARS거부')]
CODE_KEYS = ['code1', 'code2', 'code3', 'code4', 'code5']

TIERS = [1, 2, 3, 4, 5]

# 부문별 그룹: 법인사업부문은 사용자 요청으로 제외
DEPT_ORDER = ['개인사업부문', '전략사업부문', '신사업부문']
DEPT_SHORT = {'개인사업부문': '개인', '전략사업부문': '전략', '신사업부문': '신사업'}


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

    max_before_col = L(maxmap_info['before'])
    max_maxplan_col = L(maxmap_info['maxplan'])
    max_maxyears_col = L(maxmap_info['maxyears'])

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
# 3) 연차별 COUNTIFS 수식 생성 (데이터 시트 참조, 레이아웃 컬럼 위치와 무관)
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
    """대상: tier1은 그대로, 2연차부터는 이미 matured(전환여부TRUE&년수비교TRUE)인
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
    """사고유: 이번 연차=0 이거나 (공백&직전연차=0). 단 이미 matured로 확인된
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
    """전환완료: 사고유 여부와 무관하게 전환여부=TRUE & 년수비교=FALSE 전체."""
    bc = base_criteria(dr, month, group_letter, group_value, tier)
    return f'=COUNTIFS({bc},{dr.rng(dr.as_flag)},TRUE,{dr.rng(dr.yc)},FALSE)'


def formula_plan(dr, month, group_letter, group_value, tier):
    """전환예정: 이번 연차=1 & 전환여부=FALSE, 단 코드(001~004,007)가 있는 건은
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


def tier_formulas_for_group(dr, month, group_letter, group_value):
    """그룹 하나의 연차 1~5별 원시 COUNTIFS 수식 딕셔너리."""
    out = {}
    for tier in TIERS:
        f = {
            'target': formula_target(dr, month, group_letter, group_value, tier),
            'accident': formula_accident(dr, month, group_letter, group_value, tier),
            'done': formula_done(dr, month, group_letter, group_value, tier),
            'plan': formula_plan(dr, month, group_letter, group_value, tier),
        }
        for cn, cs, _ in CODE_MAP:
            key = CODE_KEYS[[1, 2, 3, 4, 7].index(cn)]
            f[key] = formula_code(dr, month, group_letter, group_value, tier, cs)
        out[tier] = f
    return out


# ---------------------------------------------------------------------------
# 4) 세부보고 레이아웃 (사용자 제공 원본 템플릿 컬럼 배치)
#    B=월 C:D(전사계만 병합)/D(개별부문)=부문별 E=연차 F=유지계약 G=사고유
#    H=전환대상 I=전환완료 J=% K=미전환 L=% M=(현장활동확인 라벨) N=전체比
#    O=전환예정 P~T=코드5종 U=현장활동미확인 V=전체比
# ---------------------------------------------------------------------------

DETAIL_COLS = {
    'month': 2, 'group_wide': 3, 'group': 4, 'tier': 5, 'target': 6, 'accident': 7,
    'conv_target': 8, 'done': 9, 'done_pct': 10, 'not_done': 11, 'not_done_pct': 12,
    'act_label': 13, 'act_pct': 14, 'plan': 15, 'code1': 16, 'code2': 17, 'code3': 18,
    'code4': 19, 'code5': 20, 'idle': 21, 'idle_pct': 22,
}
HEADER_ROWS = 3
GROUP_BLOCK_ROWS = 6  # 합계 + 연차1~5


def _hcell(ws, r, c, value=None):
    cell = ws.cell(row=r, column=c)
    if value is not None:
        cell.value = value
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    return cell


def write_detail_header(ws, row):
    """3행 헤더(원본 템플릿 그대로): 월/연차/유지계약/사고유/전환대상/전환완료/
    미전환은 3행 세로 병합, 부문별은 C:D 2열 x 3행 병합, 전환률·미전환률은
    맨 아래행에만 '%', 현장활동확인/현장활동미확인은 중간~아래 2행 병합 라벨에
    맨 아래행 리프 라벨(전체比/전환예정/코드5종)이 따라붙는다."""
    DC = DETAIL_COLS
    r0 = row

    full_merge = {
        'month': '월', 'tier': '연차', 'target': '유지계약', 'accident': '사고유',
        'conv_target': '전환대상\n(A-B)', 'done': '전환완료', 'not_done': '미전환\n(C-D)',
    }
    for key, label in full_merge.items():
        col = DC[key]
        ws.merge_cells(start_row=r0, start_column=col, end_row=r0 + 2, end_column=col)
        for rr in range(r0, r0 + 3):
            _hcell(ws, rr, col)
        _hcell(ws, r0, col, label)

    gw, g = DC['group_wide'], DC['group']
    ws.merge_cells(start_row=r0, start_column=gw, end_row=r0 + 2, end_column=g)
    for rr in range(r0, r0 + 3):
        for col in (gw, g):
            _hcell(ws, rr, col)
    _hcell(ws, r0, gw, '부문별')

    for key in ('done_pct', 'not_done_pct'):
        col = DC[key]
        for rr in range(r0, r0 + 2):
            _hcell(ws, rr, col)
        _hcell(ws, r0 + 2, col, '%')

    for key, label in (('act_label', '현장활동\n확인'), ('idle', '현장활동\n미확인')):
        col = DC[key]
        _hcell(ws, r0, col)
        ws.merge_cells(start_row=r0 + 1, start_column=col, end_row=r0 + 2, end_column=col)
        for rr in (r0 + 1, r0 + 2):
            _hcell(ws, rr, col)
        _hcell(ws, r0 + 1, col, label)

    leaves = {
        'act_pct': '전체比', 'plan': '전환예정', 'code1': '연락두절', 'code2': '사고있음',
        'code3': '고객거부', 'code4': '압류계약', 'code5': 'ARS거부', 'idle_pct': '전체比',
    }
    for key, label in leaves.items():
        col = DC[key]
        _hcell(ws, r0, col)
        _hcell(ws, r0 + 1, col)
        _hcell(ws, r0 + 2, col, label)

    return row + HEADER_ROWS


def write_group_block(ws, row0, group_label_col, group_label, tier_formulas):
    """그룹 하나(합계+연차1~5, 6행)를 쓰고 (다음 행, {'합계':row,1:row,...,5:row})를 반환."""
    DC = DETAIL_COLS
    L = get_column_letter
    tier_rows = {tier: row0 + tier for tier in TIERS}

    for tier in TIERS:
        r = tier_rows[tier]
        ws.cell(row=r, column=DC['tier'], value=tier)
        f = tier_formulas[tier]
        ws.cell(row=r, column=DC['target']).value = f['target']
        ws.cell(row=r, column=DC['accident']).value = f['accident']
        E = f'{L(DC["target"])}{r}'
        Fc = f'{L(DC["accident"])}{r}'
        G = f'{L(DC["conv_target"])}{r}'
        H = f'{L(DC["done"])}{r}'
        Jc = f'{L(DC["not_done"])}{r}'
        ws.cell(row=r, column=DC['conv_target']).value = f'={E}-{Fc}'
        ws.cell(row=r, column=DC['done']).value = f['done']
        ws.cell(row=r, column=DC['done_pct']).value = f'=IFERROR({H}/{G}*100,0)'
        ws.cell(row=r, column=DC['not_done']).value = f'={G}-{H}'
        ws.cell(row=r, column=DC['not_done_pct']).value = f'=IFERROR({Jc}/{G}*100,0)'
        ws.cell(row=r, column=DC['plan']).value = f['plan']
        for i, key in enumerate(CODE_KEYS):
            ws.cell(row=r, column=DC['code1'] + i).value = f[key]
        act_range = f'{L(DC["plan"])}{r}:{L(DC["code5"])}{r}'
        ws.cell(row=r, column=DC['act_pct']).value = f'=IFERROR(SUM({act_range})/{G}*100,0)'
        # 미활동 = 미전환 - 현장활동확인(전환예정+코드5종) 합계. 뺄셈으로 산출해
        # 예외 케이스까지 자동 흡수(전사계 행도 동일 공식이면 SUM 분배법칙으로
        # 부문별 미활동 합과 정확히 같아진다).
        ws.cell(row=r, column=DC['idle']).value = f'={Jc}-SUM({act_range})'
        Uc = f'{L(DC["idle"])}{r}'
        ws.cell(row=r, column=DC['idle_pct']).value = f'=IFERROR({Uc}/{G}*100,0)'

    r = row0
    t1, t5 = tier_rows[1], tier_rows[5]

    def sumcol(key):
        return f'SUM({L(DC[key])}{t1}:{L(DC[key])}{t5})'

    ws.cell(row=r, column=DC['tier'], value='합계')
    ws.cell(row=r, column=DC['target']).value = f'={sumcol("target")}'
    ws.cell(row=r, column=DC['accident']).value = f'={sumcol("accident")}'
    E = f'{L(DC["target"])}{r}'
    Fc = f'{L(DC["accident"])}{r}'
    G = f'{L(DC["conv_target"])}{r}'
    H = f'{L(DC["done"])}{r}'
    Jc = f'{L(DC["not_done"])}{r}'
    ws.cell(row=r, column=DC['conv_target']).value = f'={E}-{Fc}'
    ws.cell(row=r, column=DC['done']).value = f'={sumcol("done")}'
    ws.cell(row=r, column=DC['done_pct']).value = f'=IFERROR({H}/{G}*100,0)'
    ws.cell(row=r, column=DC['not_done']).value = f'={G}-{H}'
    ws.cell(row=r, column=DC['not_done_pct']).value = f'=IFERROR({Jc}/{G}*100,0)'
    ws.cell(row=r, column=DC['plan']).value = f'={sumcol("plan")}'
    for key in CODE_KEYS:
        ws.cell(row=r, column=DC[key]).value = f'={sumcol(key)}'
    act_range = f'{L(DC["plan"])}{r}:{L(DC["code5"])}{r}'
    ws.cell(row=r, column=DC['act_pct']).value = f'=IFERROR(SUM({act_range})/{G}*100,0)'
    ws.cell(row=r, column=DC['idle']).value = f'={Jc}-SUM({act_range})'
    Uc = f'{L(DC["idle"])}{r}'
    ws.cell(row=r, column=DC['idle_pct']).value = f'=IFERROR({Uc}/{G}*100,0)'

    ws.merge_cells(start_row=row0, start_column=group_label_col, end_row=row0 + 5, end_column=DC['group'])
    ws.cell(row=row0, column=group_label_col, value=group_label)

    tier_rows['합계'] = row0
    return row0 + GROUP_BLOCK_ROWS, tier_rows


def write_month_block(ws, row, dr, month, group_letter, groups):
    """월 블록 하나: 전사 계(부문 3종 합) + 부문/상품별 블록. groups는 표시용
    라벨(짧은 이름 포함 가능) -> 실제 매칭값 매핑이 아니라, (표시라벨, 실제값)
    튜플 리스트로 받는다."""
    DC = DETAIL_COLS
    L = get_column_letter
    n = len(groups)
    total_start = row
    group_starts = [row + GROUP_BLOCK_ROWS * (i + 1) for i in range(n)]

    total_tier_formulas = {}
    for tier in TIERS:
        f = {}

        def sumexpr(key, tier=tier):
            terms = [f'{L(DC[key])}{group_starts[i] + tier}' for i in range(n)]
            return '=' + '+'.join(terms)

        f['target'] = sumexpr('target')
        f['accident'] = sumexpr('accident')
        f['done'] = sumexpr('done')
        f['plan'] = sumexpr('plan')
        for key in CODE_KEYS:
            f[key] = sumexpr(key)
        total_tier_formulas[tier] = f

    _, total_rows = write_group_block(ws, total_start, DC['group_wide'], '전사 계', total_tier_formulas)

    row_map = {'전사 계': total_rows}
    for i, (label, value) in enumerate(groups):
        tf = tier_formulas_for_group(dr, month, group_letter, value)
        _, g_rows = write_group_block(ws, group_starts[i], DC['group'], label, tf)
        row_map[label] = g_rows

    ws.cell(row=total_start, column=DC['month'], value=int(month) if str(month).isdigit() else month)
    last_row = group_starts[-1] + 5
    ws.merge_cells(start_row=total_start, start_column=DC['month'], end_row=last_row, end_column=DC['month'])

    return last_row + 1, row_map


def build_detail_sheet(wb, sheet_name, data_ws_name, cols, months, group_col_key, group_labels_values,
                        months_per_header=2):
    """세부보고 시트를 월 블록으로 쌓아서 만든다. 헤더는 months_per_header
    개월마다 반복(원본 템플릿과 동일). group_labels_values: [(표시라벨, 실제값), ...]
    반환값: {month: {표시라벨/'전사 계': {'합계':row, 1:row,...,5:row}}}"""
    ws = wb.create_sheet(sheet_name)
    dr = DataRefs(data_ws_name, cols)
    group_letter = getattr(dr, group_col_key)

    row = 1
    row_map = {}
    for i, month in enumerate(months):
        if i % months_per_header == 0:
            row = write_detail_header(ws, row)
        row, month_map = write_month_block(ws, row, dr, month, group_letter, group_labels_values)
        row_map[month] = month_map
    return ws, row_map


# ---------------------------------------------------------------------------
# 5) 합산보고 레이아웃 (전사계 행 없음, 월 구분행 + 부문/상품별 한 줄씩)
#    B:C(월 구분행 병합)/C(개별부문 라벨) D=유지계약 E=사고유 F=전환대상
#    G=전환완료 H=% I=미전환 J=% K=(현장활동확인 라벨) L=전체比 M=전환예정
#    N~R=코드5종 S=현장활동미확인 T=전체比
# ---------------------------------------------------------------------------

SUMMARY_COLS = {
    'month': 2, 'group': 3, 'target': 4, 'accident': 5, 'conv_target': 6,
    'done': 7, 'done_pct': 8, 'not_done': 9, 'not_done_pct': 10,
    'act_label': 11, 'act_pct': 12, 'plan': 13, 'code1': 14, 'code2': 15, 'code3': 16,
    'code4': 17, 'code5': 18, 'idle': 19, 'idle_pct': 20,
}


def write_summary_header(ws, row):
    SC = SUMMARY_COLS
    r0 = row
    gw, g = SC['month'], SC['group']
    ws.merge_cells(start_row=r0, start_column=gw, end_row=r0 + 2, end_column=g)
    for rr in range(r0, r0 + 3):
        for col in (gw, g):
            _hcell(ws, rr, col)
    _hcell(ws, r0, gw, '구분')

    full_merge = {
        'target': '유지계약\nA', 'accident': '사고有\nB', 'conv_target': '전환대상\nC (A-B)',
        'done': '전환완료\nD', 'not_done': '전환미완료\nE (C-D)',
    }
    for key, label in full_merge.items():
        col = SC[key]
        ws.merge_cells(start_row=r0, start_column=col, end_row=r0 + 2, end_column=col)
        for rr in range(r0, r0 + 3):
            _hcell(ws, rr, col)
        _hcell(ws, r0, col, label)

    for key in ('done_pct', 'not_done_pct'):
        col = SC[key]
        for rr in range(r0, r0 + 2):
            _hcell(ws, rr, col)
        _hcell(ws, r0 + 2, col, '%')

    for key, label in (('act_label', '현장활동\n확인'), ('idle', '현장활동\n미확인')):
        col = SC[key]
        _hcell(ws, r0, col)
        ws.merge_cells(start_row=r0 + 1, start_column=col, end_row=r0 + 2, end_column=col)
        for rr in (r0 + 1, r0 + 2):
            _hcell(ws, rr, col)
        _hcell(ws, r0 + 1, col, label)

    leaves = {
        'act_pct': '%', 'plan': '전환예정', 'code1': '연락두절', 'code2': '사고있음',
        'code3': '고객거부', 'code4': '압류계약', 'code5': 'ARS거부', 'idle_pct': '%',
    }
    for key, label in leaves.items():
        col = SC[key]
        _hcell(ws, r0, col)
        _hcell(ws, r0 + 1, col)
        _hcell(ws, r0 + 2, col, label)

    return row + HEADER_ROWS


def write_summary_row(ws, row, group_label, detail_sheet_name, tier_rows):
    """detail_sheet의 연차1~5행(합계 행 제외)을 SUM해서 합산보고 한 줄을 만든다."""
    L = get_column_letter
    SC = SUMMARY_COLS
    DC = DETAIL_COLS
    r0, r1 = tier_rows[1], tier_rows[5]

    def dcol(key):
        return f"'{detail_sheet_name}'!{L(DC[key])}{r0}:{L(DC[key])}{r1}"

    ws.cell(row=row, column=SC['group'], value=group_label)
    ws.cell(row=row, column=SC['target']).value = f'=SUM({dcol("target")})'
    ws.cell(row=row, column=SC['accident']).value = f'=SUM({dcol("accident")})'
    E = f'{L(SC["target"])}{row}'
    Fc = f'{L(SC["accident"])}{row}'
    G = f'{L(SC["conv_target"])}{row}'
    H = f'{L(SC["done"])}{row}'
    Jc = f'{L(SC["not_done"])}{row}'
    ws.cell(row=row, column=SC['conv_target']).value = f'={E}-{Fc}'
    ws.cell(row=row, column=SC['done']).value = f'=SUM({dcol("done")})'
    ws.cell(row=row, column=SC['done_pct']).value = f'=IFERROR({H}/{G}*100,0)'
    ws.cell(row=row, column=SC['not_done']).value = f'={G}-{H}'
    ws.cell(row=row, column=SC['not_done_pct']).value = f'=IFERROR({Jc}/{G}*100,0)'
    ws.cell(row=row, column=SC['plan']).value = f'=SUM({dcol("plan")})'
    for key in CODE_KEYS:
        ws.cell(row=row, column=SC[key]).value = f'=SUM({dcol(key)})'
    ws.cell(row=row, column=SC['idle']).value = f'=SUM({dcol("idle")})'
    act_range = f'{L(SC["plan"])}{row}:{L(SC["code5"])}{row}'
    ws.cell(row=row, column=SC['act_pct']).value = f'=IFERROR(SUM({act_range})/{G}*100,0)'
    Uc = f'{L(SC["idle"])}{row}'
    ws.cell(row=row, column=SC['idle_pct']).value = f'=IFERROR({Uc}/{G}*100,0)'
    return row + 1


def build_summary_sheet(wb, sheet_name, detail_sheet_name, months, group_labels, row_map):
    """합산보고: 전사계 행 없이, 월 구분행(병합) 아래 그룹별 한 줄씩."""
    SC = SUMMARY_COLS
    ws = wb.create_sheet(sheet_name)
    row = 1
    row = write_summary_header(ws, row)
    month_col = SC['month']
    for month in months:
        month_start = row
        month_label = f'{int(month) if str(month).isdigit() else month}월'
        ws.cell(row=row, column=month_col, value=month_label)
        row += 1
        for label in group_labels:
            row = write_summary_row(ws, row, label, detail_sheet_name, row_map[month][label])
        month_end = row - 1
        ws.merge_cells(start_row=month_start, start_column=month_col,
                        end_row=month_end, end_column=month_col)
        ws.cell(row=month_start, column=month_col).alignment = Alignment(
            horizontal='center', vertical='center')
    return ws


# ---------------------------------------------------------------------------
# 6) 서식 적용
# ---------------------------------------------------------------------------

def _apply_common_style(ws, ncols, pct_cols, count_cols):
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


def style_detail_sheet(ws, row_map):
    DC = DETAIL_COLS
    ncols = max(DC.values())
    pct_cols = {DC['done_pct'], DC['not_done_pct'], DC['act_pct'], DC['idle_pct']}
    count_cols = {DC['target'], DC['accident'], DC['conv_target'], DC['done'], DC['not_done'],
                  DC['plan'], DC['code1'], DC['code2'], DC['code3'], DC['code4'], DC['code5'],
                  DC['idle']}
    _apply_common_style(ws, ncols, pct_cols, count_cols)

    for month, groups in row_map.items():
        for label, tier_rows in groups.items():
            total_row = tier_rows['합계']
            rows = [total_row] + [tier_rows[t] for t in TIERS] if label == '전사 계' else [total_row]
            fill = TOTAL_FILL if label == '전사 계' else GROUP_FILL
            for r in rows:
                for c in range(2, 2 + ncols):
                    ws.cell(row=r, column=c).fill = fill

    widths = {'B': 6, 'C': 4, 'D': 12, 'E': 5, 'F': 9, 'G': 8, 'H': 9, 'I': 9, 'J': 8,
              'K': 9, 'L': 8, 'M': 3, 'N': 8, 'O': 8, 'P': 8, 'Q': 8, 'R': 8, 'S': 8,
              'T': 8, 'U': 9, 'V': 8}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.sheet_view.showGridLines = False


def style_summary_sheet(ws):
    SC = SUMMARY_COLS
    ncols = max(SC.values())
    pct_cols = {SC['done_pct'], SC['not_done_pct'], SC['act_pct'], SC['idle_pct']}
    count_cols = {SC['target'], SC['accident'], SC['conv_target'], SC['done'], SC['not_done'],
                  SC['plan'], SC['code1'], SC['code2'], SC['code3'], SC['code4'], SC['code5'],
                  SC['idle']}
    _apply_common_style(ws, ncols, pct_cols, count_cols)

    for r in range(1, ws.max_row + 1):
        gcell = ws.cell(row=r, column=SC['group'])
        if gcell.value not in (None,):
            gcell.fill = GROUP_FILL

    widths = {'B': 6, 'C': 8, 'D': 9, 'E': 8, 'F': 9, 'G': 9, 'H': 8,
              'I': 9, 'J': 8, 'K': 3, 'L': 8, 'M': 8, 'N': 8, 'O': 8, 'P': 8,
              'Q': 8, 'R': 8, 'S': 9, 'T': 8}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------------------
# 7) 메인
# ---------------------------------------------------------------------------

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
    all_depts = set(get_distinct_values(data_ws, cols['dept'], d_last_row))
    depts = [d for d in DEPT_ORDER if d in all_depts]
    products = sorted(get_distinct_values(data_ws, cols['product'], d_last_row))
    print(f"  - 계약체결월: {months}")
    print(f"  - 수금부문명({len(depts)}종, 법인 제외): {depts}")
    print(f"  - 상품명({len(products)}종): {products}")

    dept_pairs = [(DEPT_SHORT.get(d, d), d) for d in depts]
    dept_pairs_full = [(d, d) for d in depts]
    product_pairs = [(p, p) for p in products]

    print("[2/3] 레이아웃(세부보고/합산보고 x 부문별/상품별) 생성 중...")
    ws_detail_dept, rows_dept = build_detail_sheet(
        wb, '세부보고_부문별', data_ws.title, cols, months, 'dept', dept_pairs_full)
    style_detail_sheet(ws_detail_dept, rows_dept)

    ws_detail_prod, rows_prod = build_detail_sheet(
        wb, '세부보고_상품별', data_ws.title, cols, months, 'product', product_pairs)
    style_detail_sheet(ws_detail_prod, rows_prod)

    dept_labels = [label for label, _ in dept_pairs]
    rows_dept_short = {
        month: {DEPT_SHORT.get(full, full): v for full, v in groups.items() if full != '전사 계'}
        for month, groups in rows_dept.items()
    }
    ws_summary_dept = build_summary_sheet(
        wb, '합산보고_부문별', '세부보고_부문별', months, dept_labels, rows_dept_short)
    style_summary_sheet(ws_summary_dept)

    product_labels = [label for label, _ in product_pairs]
    rows_prod_no_total = {
        month: {k: v for k, v in groups.items() if k != '전사 계'}
        for month, groups in rows_prod.items()
    }
    ws_summary_prod = build_summary_sheet(
        wb, '합산보고_상품별', '세부보고_상품별', months, product_labels, rows_prod_no_total)
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
