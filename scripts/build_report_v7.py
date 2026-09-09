"""
build_report_v7.py — 신규 6시트 보고서 양식(사용자 제공, 2026-09 업로드) 레이아웃 생성.

1차 단계: 데이터/수식 없이 "빈 양식"만 만든다(셀 병합·배경색·테두리·열너비 확인용).
6개 시트: 월별_부문별 / 연차별_부문별_상세 / 월별_부문별_연차별 /
          쳬결기간별_부문별 / 상품별_부문별_연차별 / 상품별_부문별_합산

공통 지표 블록(전 시트 동일, 시작 컬럼만 다름, 폭 18칸):
  유지계약A, 사고有B, 전환대상C, %, 전환완료D, %, 전환미완료E, %,
  현장활동확인(라벨), %, 전환예정, 연락두절, 사고있음, 고객거부, 압류계약, ARS거부,
  현장활동미확인, %
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

TITLE_FONT = Font(name='맑은 고딕', size=18, bold=True)
STAMP_FONT = Font(name='맑은 고딕', size=10)


def write_title_and_stamp(ws, title, title_row, end_col, stamp_row=None):
    """시트 상단 제목(굵은 큰 글씨, 가운데정렬)과 '(YYYY.MM.DD 기준)' 스탬프를 쓴다.
    stamp_row가 없으면 title_row 바로 아래 행(헤더 시작 직전)에 찍는다."""
    ws.merge_cells(start_row=title_row, start_column=2, end_row=title_row + 3, end_column=end_col)
    cell = ws.cell(row=title_row, column=2, value=title)
    cell.font = TITLE_FONT
    cell.alignment = Alignment(horizontal='center', vertical='center')
    sr = stamp_row if stamp_row is not None else title_row + 3
    scell = ws.cell(row=sr, column=end_col, value=f'({date.today():%Y.%m.%d} 기준)')
    scell.font = STAMP_FONT
    scell.alignment = Alignment(horizontal='right', vertical='bottom')

FONT_NAME = '맑은 고딕'
THIN = Side(style='thin')
MEDIUM = Side(style='medium')

# 원본 신규 양식에서 추출한 색상
HEADER_FILL = PatternFill('solid', fgColor='F2F2F2')       # theme lt1 tint-0.05
PCT_HEADER_FILL = PatternFill('solid', fgColor='FBE5D6')   # theme accent2(ED7D31) tint0.8 - '%' 리프 라벨 전용
TOTAL_FILL = PatternFill('solid', fgColor='DAE3F3')         # theme accent5(4472C4) tint0.8 - 전사계 블록
PCT_FILL = PatternFill('solid', fgColor='E2EFDA')           # theme accent6(70AD47) tint0.8 - 비율열 데이터
DEPT_FILL = {
    '개인': PatternFill('solid', fgColor='FFFF00'),
    '전략': PatternFill('solid', fgColor='92D050'),
    '신사업': PatternFill('solid', fgColor='FFC000'),
}
HEADER_FONT = Font(name=FONT_NAME, size=11)

DEPTS_SHORT = ['개인', '전략', '신사업']

# 지표 블록 리프 정의: (오프셋, 라벨, 병합폭(1이면 단일열), 종류)
# 종류: 'plain'(값만) / 'ratio'(값+바로 다음 열이 %) / 'pct'(비율 표시 전용, ratio가 채움) /
#       'section'(2행 병합 라벨, 리프 자식들이 이어짐)
METRIC_WIDTH = 18


def _hcell(ws, r, c, value=None, fill=HEADER_FILL):
    cell = ws.cell(row=r, column=c)
    if value is not None:
        cell.value = value
    cell.font = HEADER_FONT
    cell.fill = fill
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    return cell


def write_metric_header(ws, r0, col0, target_label='유지계약\nA', pct_header_fill=PCT_HEADER_FILL):
    """3행 지표 헤더(유지계약~현장활동미확인, 폭 18칸)를 col0부터 그린다."""
    L = get_column_letter

    def full3(col, label):
        ws.merge_cells(start_row=r0, start_column=col, end_row=r0 + 2, end_column=col)
        for rr in range(r0, r0 + 3):
            _hcell(ws, rr, col)
        _hcell(ws, r0, col, label)

    full3(col0, target_label)
    full3(col0 + 1, '사고有\nB')
    full3(col0 + 2, '전환대상\nC (A-B)')
    full3(col0 + 4, '전환완료\nD')
    full3(col0 + 6, '전환미완료\nE (C-D)')

    for off in (3, 5, 7):
        col = col0 + off
        for rr in range(r0, r0 + 2):
            _hcell(ws, rr, col)
        _hcell(ws, r0 + 2, col, '%', fill=pct_header_fill)

    for off, label in ((8, '현장활동\n확인'), (16, '현장활동\n미확인')):
        col = col0 + off
        _hcell(ws, r0, col)
        ws.merge_cells(start_row=r0 + 1, start_column=col, end_row=r0 + 2, end_column=col)
        for rr in (r0 + 1, r0 + 2):
            _hcell(ws, rr, col)
        _hcell(ws, r0 + 1, col, label)

    leaves = {
        9: '%', 10: '전환예정', 11: '연락두절', 12: '사고있음',
        13: '고객거부', 14: '압류계약', 15: 'ARS거부', 17: '%',
    }
    for off, label in leaves.items():
        col = col0 + off
        _hcell(ws, r0, col)
        _hcell(ws, r0 + 1, col)
        fill = pct_header_fill if off in (9, 17) else HEADER_FILL
        _hcell(ws, r0 + 2, col, label, fill=fill)

    # 구획선(medium): 전환완료(+4) 좌측/전환률(+5) 우측, 미전환(+6) 좌측/전체比(+17) 우측
    medium_left = {col0 + 4, col0 + 6}
    medium_right = {col0 + 5, col0 + 17}
    top_from = col0 + 4
    for rr in range(r0, r0 + 3):
        for c in range(col0, col0 + METRIC_WIDTH):
            cell = ws.cell(row=rr, column=c)
            b = cell.border
            left = MEDIUM if c in medium_left else b.left
            right = MEDIUM if c in medium_right else b.right
            top = MEDIUM if (rr == r0 and c >= top_from) else b.top
            cell.border = Border(top=top, bottom=b.bottom, left=left, right=right)


def _metric_fill(ws, row, col0, fill):
    for c in range(col0, col0 + METRIC_WIDTH):
        ws.cell(row=row, column=c).fill = fill


def _metric_border(ws, row, col0):
    medium_left = {col0 + 4, col0 + 6}
    medium_right = {col0 + 5, col0 + 17}
    for c in range(col0, col0 + METRIC_WIDTH):
        left = MEDIUM if c in medium_left else THIN
        right = MEDIUM if c in medium_right else THIN
        ws.cell(row=row, column=c).border = Border(top=THIN, bottom=THIN, left=left, right=right)
        ws.cell(row=row, column=c).font = Font(name=FONT_NAME, size=10)


# 지표 블록 안에서 '%'가 표시되는 오프셋(전환대상/전환완료/미전환/전체比 x2)
PCT_OFFSETS = (3, 5, 7, 9, 17)


def _pct_tint(ws, row, col0):
    for off in PCT_OFFSETS:
        ws.cell(row=row, column=col0 + off).fill = PCT_FILL


# ---------------------------------------------------------------------------
# 1) 월별_부문별_연차별 (세부보고, 12개월 x 부문 x 연차1~5)
# ---------------------------------------------------------------------------

def build_month_dept_tier(wb):
    ws = wb.create_sheet('월별_부문별_연차별')
    ws.sheet_view.showGridLines = False
    col_month, col_gw, col_g, col_tier = 2, 3, 4, 5
    metric0 = 6
    scell = ws.cell(row=1, column=metric0 + METRIC_WIDTH - 1, value=f'({date.today():%Y.%m.%d} 기준)')
    scell.font = STAMP_FONT
    scell.alignment = Alignment(horizontal='right', vertical='bottom')

    def write_header(r0):
        ws.merge_cells(start_row=r0, start_column=col_month, end_row=r0 + 2, end_column=col_month)
        for rr in range(r0, r0 + 3):
            _hcell(ws, rr, col_month)
        _hcell(ws, r0, col_month, '월')

        ws.merge_cells(start_row=r0, start_column=col_gw, end_row=r0 + 2, end_column=col_g)
        for rr in range(r0, r0 + 3):
            for c in (col_gw, col_g):
                _hcell(ws, rr, c)
        _hcell(ws, r0, col_gw, '부문별')

        ws.merge_cells(start_row=r0, start_column=col_tier, end_row=r0 + 2, end_column=col_tier)
        for rr in range(r0, r0 + 3):
            _hcell(ws, rr, col_tier)
        _hcell(ws, r0, col_tier, '연차')

        write_metric_header(ws, r0, metric0)
        return r0 + 3

    row = 2
    tall_rows = []
    for i, month in enumerate(range(1, 13)):
        if i % 2 == 0:
            tall_rows.append(row)
            row = write_header(row)
        total_start = row
        ws.cell(row=row, column=col_tier, value='합계')
        for tier in range(1, 6):
            r = row + tier
            ws.cell(row=r, column=col_tier, value=tier)
        for r in range(total_start, total_start + 6):
            _metric_fill(ws, r, metric0, TOTAL_FILL)
            _metric_border(ws, r, metric0)
            ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
            ws.cell(row=r, column=col_tier).fill = TOTAL_FILL
        ws.cell(row=total_start, column=col_gw, value='전사 계')
        ws.merge_cells(start_row=total_start, start_column=col_gw, end_row=total_start + 5, end_column=col_g)
        row = total_start + 6
        # 개별 부문: 강조색 없이 부문별(폭넓은) 열에만 파란 스파인만 유지
        for label in DEPTS_SHORT:
            dept_start = row
            ws.cell(row=row, column=col_tier, value='합계')
            for tier in range(1, 6):
                r = row + tier
                ws.cell(row=r, column=col_tier, value=tier)
            for r in range(dept_start, dept_start + 6):
                _metric_border(ws, r, metric0)
                ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
            ws.cell(row=dept_start, column=col_g, value=label)
            ws.merge_cells(start_row=dept_start, start_column=col_g, end_row=dept_start + 5, end_column=col_g)
            row = dept_start + 6
        ws.cell(row=total_start, column=col_month, value=month)
        ws.merge_cells(start_row=total_start, start_column=col_month, end_row=row - 1, end_column=col_month)
        ws.cell(row=total_start, column=col_month).alignment = Alignment(horizontal='center', vertical='center')
        if i < 11:
            tall_rows.append(row)
            row += 1

    for r in range(1, row):
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'B': 5.625, 'D': 11.375, 'E': 5.25, 'F': 8.375, 'I': 6.375, 'K': 6.375,
              'L': 10.375, 'M': 6.375, 'N': 8.0, 'O': 7.125, 'U': 8.625, 'V': 8.375,
              'W': 7.125, 'Z': 11.125}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 2) 연차별_부문별_상세 (전체 기간 합산, 부문 x 연차1~4)
# ---------------------------------------------------------------------------

def build_tier_dept_summary(wb, max_tier=4):
    ws = wb.create_sheet('연차별_부문별_상세')
    ws.sheet_view.showGridLines = False
    col_gw, col_g, col_tier = 2, 3, 4
    metric0 = 5

    r0 = 6
    write_title_and_stamp(ws, '연차별_부문별_무사고전환율', 1, metric0 + METRIC_WIDTH - 1, stamp_row=r0 - 1)
    ws.merge_cells(start_row=r0, start_column=col_gw, end_row=r0 + 2, end_column=col_g)
    for rr in range(r0, r0 + 3):
        for c in (col_gw, col_g):
            _hcell(ws, rr, c)
    _hcell(ws, r0, col_gw, '부문별')
    ws.merge_cells(start_row=r0, start_column=col_tier, end_row=r0 + 2, end_column=col_tier)
    for rr in range(r0, r0 + 3):
        _hcell(ws, rr, col_tier)
    _hcell(ws, r0, col_tier, '연차')
    write_metric_header(ws, r0, metric0)

    row = r0 + 3
    for label in ['전사 계'] + DEPTS_SHORT:
        start = row
        ws.cell(row=row, column=col_tier, value='합계')
        for tier in range(1, max_tier + 1):
            r = row + tier
            ws.cell(row=r, column=col_tier, value=tier)
        label_col = col_gw if label == '전사 계' else col_g
        for r in range(start, start + max_tier + 1):
            _metric_border(ws, r, metric0)
            ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
            if label == '전사 계':
                _metric_fill(ws, r, metric0, TOTAL_FILL)
                ws.cell(row=r, column=col_tier).fill = TOTAL_FILL
        ws.cell(row=start, column=label_col, value=label)
        ws.merge_cells(start_row=start, start_column=label_col, end_row=start + max_tier, end_column=col_g)
        row = start + max_tier + 1

    for r in range(r0 + 3, row):
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'B': 5.625, 'C': 11.375, 'D': 5.25, 'E': 9.375, 'H': 7.125, 'I': 9.375,
              'J': 7.125, 'K': 10.375, 'L': 7.125, 'M': 8.0, 'N': 7.125, 'T': 8.625,
              'U': 8.375, 'V': 7.125}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.row_dimensions[r0].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 3) 월별_부문별 (연차 구분 없는 월별 요약: 전체합산 블록 + 월별 블록)
# ---------------------------------------------------------------------------

def build_month_dept_summary(wb, months=range(1, 13)):
    ws = wb.create_sheet('월별_부문별')
    ws.sheet_view.showGridLines = False
    col_g = 2
    metric0 = 4

    r0 = 6
    write_title_and_stamp(ws, '월별_부문별_무사고전환율', 1, metric0 + METRIC_WIDTH - 1, stamp_row=r0 - 1)
    ws.merge_cells(start_row=r0, start_column=col_g, end_row=r0 + 2, end_column=col_g + 1)
    for rr in range(r0, r0 + 3):
        for c in (col_g, col_g + 1):
            _hcell(ws, rr, c)
    _hcell(ws, r0, col_g, '구분')
    write_metric_header(ws, r0, metric0, pct_header_fill=PCT_FILL)

    row = r0 + 3
    tall_rows = [r0]

    def write_block(label):
        nonlocal row
        start = row
        ws.cell(row=row, column=col_g, value=label)
        ws.merge_cells(start_row=row, start_column=col_g, end_row=row, end_column=col_g + 1)
        _metric_border(ws, row, metric0)
        _pct_tint(ws, row, metric0)
        row += 1
        for d in DEPTS_SHORT:
            ws.cell(row=row, column=col_g + 1, value=d)
            _metric_border(ws, row, metric0)
            _pct_tint(ws, row, metric0)
            row += 1
        return start

    # 원본 템플릿: 이 요약 시트는 부문별 강조색이 전혀 없고, 비율(%)열만 항상
    # 연두색으로 칠해져 있다(전사계/부문 구분 없이 데이터 행 전부).
    write_block('전사 계')
    tall_rows.append(row)
    row += 1
    for month in months:
        write_block(f'{month}월')
        if month != list(months)[-1]:
            tall_rows.append(row)
            row += 1

    widths = {'A': 9.375, 'B': 3.125, 'F': 9.375, 'G': 6.125, 'I': 6.125, 'J': 9.5,
              'K': 6.125, 'M': 6.125, 'N': 8.75, 'O': 8.0, 'U': 6.125}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 4) 상품별_부문별_연차별 (상품명 x 부문 x 연차1~5)
# ---------------------------------------------------------------------------

def build_product_dept_tier(wb, products):
    ws = wb.create_sheet('상품별_부문별_연차별')
    ws.sheet_view.showGridLines = False
    col_p, col_gw, col_g, col_tier = 2, 3, 4, 5
    metric0 = 6

    r0 = 6
    write_title_and_stamp(ws, '상품별_부문별_연차별_무사고전환율', 1, metric0 + METRIC_WIDTH - 1, stamp_row=r0 - 1)
    ws.merge_cells(start_row=r0, start_column=col_p, end_row=r0 + 2, end_column=col_p)
    for rr in range(r0, r0 + 3):
        _hcell(ws, rr, col_p)
    _hcell(ws, r0, col_p, '상품명')
    ws.merge_cells(start_row=r0, start_column=col_gw, end_row=r0 + 2, end_column=col_g)
    for rr in range(r0, r0 + 3):
        for c in (col_gw, col_g):
            _hcell(ws, rr, c)
    _hcell(ws, r0, col_gw, '부문별')
    ws.merge_cells(start_row=r0, start_column=col_tier, end_row=r0 + 2, end_column=col_tier)
    for rr in range(r0, r0 + 3):
        _hcell(ws, rr, col_tier)
    _hcell(ws, r0, col_tier, '연차')
    write_metric_header(ws, r0, metric0)

    row = r0 + 3
    tall_rows = [r0]
    for p in products:
        p_start = row
        for label in ['전사 계'] + DEPTS_SHORT:
            g_start = row
            label_col = col_gw if label == '전사 계' else col_g
            ws.cell(row=row, column=col_tier, value='합계')
            for tier in range(1, 6):
                r = row + tier
                ws.cell(row=r, column=col_tier, value=tier)
            for r in range(g_start, g_start + 6):
                _metric_border(ws, r, metric0)
                ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
                if label == '전사 계':
                    _metric_fill(ws, r, metric0, TOTAL_FILL)
                    ws.cell(row=r, column=col_tier).fill = TOTAL_FILL
            ws.cell(row=g_start, column=label_col, value=label)
            ws.merge_cells(start_row=g_start, start_column=label_col, end_row=g_start + 5, end_column=label_col)
            row = g_start + 6
        ws.cell(row=p_start, column=col_p, value=p)
        ws.merge_cells(start_row=p_start, start_column=col_p, end_row=row - 1, end_column=col_p)
        tall_rows.append(row)
        row += 1

    for r in range(r0 + 3, row):
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'B': 5.625, 'D': 11.375, 'E': 5.25, 'F': 8.375, 'I': 6.375, 'K': 6.375,
              'L': 10.375, 'M': 6.375, 'N': 8.0, 'O': 7.125, 'U': 8.625, 'V': 8.375, 'W': 7.125}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 5) 상품별_부문별_합산 (상품명 x 부문, 연차 구분 없음)
# ---------------------------------------------------------------------------

def build_product_dept_summary(wb, products):
    ws = wb.create_sheet('상품별_부문별_합산')
    ws.sheet_view.showGridLines = False
    col_p, col_gw, col_g = 2, 3, 4
    metric0 = 5

    r0 = 6
    write_title_and_stamp(ws, '상품별_부문별_합산_무사고전환율', 1, metric0 + METRIC_WIDTH - 1, stamp_row=r0 - 1)
    ws.merge_cells(start_row=r0, start_column=col_p, end_row=r0 + 2, end_column=col_p)
    for rr in range(r0, r0 + 3):
        _hcell(ws, rr, col_p)
    _hcell(ws, r0, col_p, '상품명')
    ws.merge_cells(start_row=r0, start_column=col_gw, end_row=r0 + 2, end_column=col_g)
    for rr in range(r0, r0 + 3):
        for c in (col_gw, col_g):
            _hcell(ws, rr, c)
    _hcell(ws, r0, col_gw, '부문별')
    write_metric_header(ws, r0, metric0)

    row = r0 + 3
    tall_rows = [r0]
    for p in products:
        p_start = row
        for label in ['전사 계'] + DEPTS_SHORT:
            label_col = col_gw if label == '전사 계' else col_g
            ws.cell(row=row, column=label_col, value=label)
            ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
            if label == '전사 계':
                _metric_fill(ws, row, metric0, TOTAL_FILL)
            _metric_border(ws, row, metric0)
            row += 1
        ws.cell(row=p_start, column=col_p, value=p)
        ws.merge_cells(start_row=p_start, start_column=col_p, end_row=row - 1, end_column=col_p)
        tall_rows.append(row)
        row += 1

    widths = {'B': 5.625, 'D': 11.375, 'E': 8.375, 'H': 6.375, 'J': 6.375, 'K': 10.375,
              'L': 6.375, 'M': 8.0, 'N': 7.125, 'T': 8.625, 'U': 8.375, 'V': 7.125}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 6) 쳬결기간별_부문별 (체결구간 x 부문 x (구간키값E, 연차F))
# ---------------------------------------------------------------------------

def build_period_dept(wb, periods):
    """periods: [(구간라벨, 구간키값, [도달가능 연차,...]), ...] 오래된 순."""
    ws = wb.create_sheet('쳬결기간별_부문별')
    ws.sheet_view.showGridLines = False
    col_gw, col_g, col_period, col_key, col_tier = 2, 3, 4, 5, 6
    metric0 = 7

    r0 = 7
    write_title_and_stamp(ws, '체결기간별_부문별_무사고전환율', 2, metric0 + METRIC_WIDTH - 1, stamp_row=r0 - 1)
    ws.merge_cells(start_row=r0, start_column=col_gw, end_row=r0 + 2, end_column=col_g)
    for rr in range(r0, r0 + 3):
        for c in (col_gw, col_g):
            _hcell(ws, rr, c)
    _hcell(ws, r0, col_gw, '부문별')
    ws.merge_cells(start_row=r0, start_column=col_period, end_row=r0 + 2, end_column=col_period)
    for rr in range(r0, r0 + 3):
        _hcell(ws, rr, col_period)
    _hcell(ws, r0, col_period, '연도별')
    ws.merge_cells(start_row=r0, start_column=col_key, end_row=r0 + 2, end_column=col_tier)
    for rr in range(r0, r0 + 3):
        for c in (col_key, col_tier):
            _hcell(ws, rr, c)
    _hcell(ws, r0, col_key, '연차')
    write_metric_header(ws, r0, metric0, target_label='대상계약\nA')

    row = r0 + 3
    tall_rows = [r0]
    # 원본 템플릿: 이 시트는 부문별(전사계=B, 개별부문=C)열에만 파란 스파인이 있고,
    # 전사계를 포함해 지표 칸에는 강조색이 전혀 없다(비율열 연두색조차 없음).
    for label in ['전사 계'] + DEPTS_SHORT:
        g_start = row
        label_col = col_gw if label == '전사 계' else col_g
        for plabel, key, reach_tiers in periods:
            p_start = row
            ws.cell(row=row, column=col_period, value=plabel)
            ws.cell(row=row, column=col_key, value='전체')
            _metric_border(ws, row, metric0)
            ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
            if len(reach_tiers) > 1:
                row += 1
                for t in reach_tiers:
                    ws.cell(row=row, column=col_key, value=key)
                    ws.cell(row=row, column=col_tier, value=t)
                    _metric_border(ws, row, metric0)
                    ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
                    row += 1
                ws.merge_cells(start_row=p_start + 1, start_column=col_key,
                                end_row=row - 1, end_column=col_key)
            else:
                ws.cell(row=row, column=col_tier, value=reach_tiers[0])
                row += 1
            if plabel != periods[-1][0]:
                ws.merge_cells(start_row=p_start, start_column=col_period,
                                end_row=row - 1, end_column=col_period)
        ws.cell(row=g_start, column=label_col, value=label)
        ws.merge_cells(start_row=g_start, start_column=label_col, end_row=row - 1, end_column=label_col)

    for r in range(r0 + 3, row):
        ws.cell(row=r, column=col_key).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'A': 11.5, 'B': 5.625, 'C': 11.375, 'D': 16.25, 'E': 5.25, 'G': 8.375,
              'I': 8.625, 'J': 6.375, 'K': 8.625, 'L': 6.375, 'M': 10.375, 'N': 6.375,
              'P': 7.125, 'V': 8.625, 'W': 8.375, 'X': 7.125}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


def build_blank(output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_month_dept_summary(wb)
    build_tier_dept_summary(wb)
    build_month_dept_tier(wb)
    default_periods = [
        ('2022.07~2023.06', 1, [1, 2, 3, 4]),
        ('2023.07~2024.06', 2, [1, 2, 3]),
        ('2024.07~2025.06', 3, [1, 2]),
        ('2025.07~2026.06', 4, [1]),
    ]
    build_period_dept(wb, default_periods)
    products = ['상품A', '상품B', '상품C']
    build_product_dept_tier(wb, products)
    build_product_dept_summary(wb, products)
    wb.save(output_path)
    print('저장 완료 ->', output_path)


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else 'blank_form_v7.xlsx'
    build_blank(out)
