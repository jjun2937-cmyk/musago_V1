"""
write_report_v7.py — calc_report_v7.py로 계산한 값을, build_report_v7.py가 만든
레이아웃(서식/병합/색상)에 실제로 채워 넣어 최종 보고서 파일을 만든다.
LibreOffice/Excel 수식은 전혀 쓰지 않는다(순수 값만 기록).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font

THIN_SIDE = Side(style='thin')

import build_report_v7 as layout
from build_report_v7 import (
    METRIC_WIDTH, METRIC_WIDTH_PERIOD, PCT_OFFSETS, TOTAL_FILL, DEPTS_SHORT,
    write_metric_header, _metric_border, _metric_fill, _pct_tint, write_title_and_stamp, _hcell,
    write_metric_header_period, _metric_border_period, _metric_fill_period,
    BODY_FONT, LABEL_FONT, set_group_label, set_value_label, set_label_border,
)
import calc_report_v7 as calc
from calc_report_v7 import DEPT_ORDER, DEPT_SHORT, CODE_MAP, TIERS, PERIODS

COUNT_FMT = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
RATIO_FMT = '_-* #,##0.0_-;-* #,##0.0_-;_-* "-"_-;_-@_-'
VALUE_FONT_10 = Font(name=layout.FONT_NAME, size=10)
CODE_STRS = [cs for _, cs in CODE_MAP]


def _ratio(n, d):
    return round(n / d * 100, 4) if d else 0


def write_metric_values(ws, row, col0, m):
    """m: TierAccum.finalize_bucket()가 반환한 dict(target,accident,conv_target,
    done,not_done,plan,codes,idle)."""
    target, accident, conv_target = m['target'], m['accident'], m['conv_target']
    done, not_done, plan, idle = m['done'], m['not_done'], m['plan'], m['idle']
    codes = m['codes']
    act_sum = plan + sum(codes.values())

    def setv(off, value, is_pct=False):
        cell = ws.cell(row=row, column=col0 + off, value=value)
        cell.number_format = RATIO_FMT if is_pct else COUNT_FMT
        cell.font = BODY_FONT
        cell.alignment = Alignment(horizontal='right', vertical='center')

    setv(0, target)
    setv(1, accident)
    setv(2, conv_target)
    setv(3, _ratio(conv_target, target), True)
    setv(4, done)
    setv(5, _ratio(done, conv_target), True)
    setv(6, not_done)
    setv(7, _ratio(not_done, conv_target), True)
    setv(8, act_sum)
    setv(9, _ratio(act_sum, conv_target), True)
    setv(10, plan)
    for i, (_cn, cs) in enumerate(CODE_MAP):
        setv(11 + i, codes.get(cs, 0))
    setv(16, idle)
    setv(17, _ratio(idle, conv_target), True)


def write_period_values(ws, row, col0, m):
    """m: PeriodAccum.get()/_sum_period_parts()가 반환한 dict(target,accident,
    conv_target,done,max_done,not_done,plan,codes,idle). 체결기간별_부문별
    전용(20칸 폭, '최대전환완료' 값+% 포함)."""
    target, accident, conv_target = m['target'], m['accident'], m['conv_target']
    done, max_done = m['done'], m['max_done']
    not_done, plan, idle = m['not_done'], m['plan'], m['idle']
    codes = m['codes']
    act_sum = plan + sum(codes.values())

    def setv(off, value, is_pct=False):
        cell = ws.cell(row=row, column=col0 + off, value=value)
        cell.number_format = RATIO_FMT if is_pct else COUNT_FMT
        cell.font = BODY_FONT
        cell.alignment = Alignment(horizontal='right', vertical='center')

    setv(0, target)
    setv(1, accident)
    setv(2, conv_target)
    setv(3, _ratio(conv_target, target), True)
    setv(4, done)
    setv(5, _ratio(done, conv_target), True)
    setv(6, max_done)
    setv(7, _ratio(max_done, conv_target), True)
    setv(8, not_done)
    setv(9, _ratio(not_done, conv_target), True)
    setv(10, act_sum)
    setv(11, _ratio(act_sum, conv_target), True)
    setv(12, plan)
    for i, (_cn, cs) in enumerate(CODE_MAP):
        setv(13 + i, codes.get(cs, 0))
    setv(18, idle)
    setv(19, _ratio(idle, conv_target), True)


# ---------------------------------------------------------------------------
# 1) 월별_부문별_연차별
# ---------------------------------------------------------------------------

def build_month_dept_tier(wb, months, month_dept_acc, all_dept_acc):
    ws = wb.create_sheet('월별_부문별_연차별')
    ws.sheet_view.showGridLines = False
    col_month, col_gw, col_g, col_tier = 2, 3, 4, 5
    metric0 = 6
    from datetime import date
    scell = ws.cell(row=1, column=metric0 + METRIC_WIDTH - 1, value=f'({date.today():%Y.%m.%d} 기준)')
    scell.font = layout.STAMP_FONT
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
    for i, month in enumerate(months):
        if i % 2 == 0:
            tall_rows.append(row)
            row = write_header(row)
        total_start = row
        per_dept = {d: month_dept_acc.finalize((month, d)) for d in DEPT_ORDER}
        total_metrics = calc.sum_metrics([per_dept[d] for d in DEPT_ORDER])

        # 전사계 합계행(연차 1~5 합)
        agg = {
            'target': sum(total_metrics[t]['target'] for t in TIERS),
            'accident': sum(total_metrics[t]['accident'] for t in TIERS),
            'done': sum(total_metrics[t]['done'] for t in TIERS),
            'plan': sum(total_metrics[t]['plan'] for t in TIERS),
            'codes': {cs: sum(total_metrics[t]['codes'][cs] for t in TIERS) for cs in CODE_STRS},
        }
        agg['conv_target'] = agg['target'] - agg['accident']
        agg['not_done'] = agg['conv_target'] - agg['done']
        agg['idle'] = agg['not_done'] - (agg['plan'] + sum(agg['codes'].values()))
        set_value_label(ws, row, col_tier, '합계')
        write_metric_values(ws, row, metric0, agg)
        for r in range(row, row + 6):
            layout._metric_fill(ws, r, metric0, TOTAL_FILL)
            layout._metric_border(ws, r, metric0)
            ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
            ws.cell(row=r, column=col_tier).fill = TOTAL_FILL
            set_label_border(ws, r, (col_month, col_g, col_tier))
        for t in TIERS:
            set_value_label(ws, row + t, col_tier, t)
            write_metric_values(ws, row + t, metric0, total_metrics[t])
        set_group_label(ws, row, col_gw, '전사 계')
        ws.merge_cells(start_row=row, start_column=col_gw, end_row=row + 5, end_column=col_g)
        bottom_cell = ws.cell(row=row + 5, column=col_g)
        b = bottom_cell.border
        bottom_cell.border = Border(top=b.top, bottom=None, left=b.left, right=b.right)
        row += 6

        for d in DEPT_ORDER:
            dept_start = row
            dm = per_dept[d]
            agg = {
                'target': sum(dm[t]['target'] for t in TIERS),
                'accident': sum(dm[t]['accident'] for t in TIERS),
                'done': sum(dm[t]['done'] for t in TIERS),
                'plan': sum(dm[t]['plan'] for t in TIERS),
                'codes': {cs: sum(dm[t]['codes'][cs] for t in TIERS) for cs in CODE_STRS},
            }
            agg['conv_target'] = agg['target'] - agg['accident']
            agg['not_done'] = agg['conv_target'] - agg['done']
            agg['idle'] = agg['not_done'] - (agg['plan'] + sum(agg['codes'].values()))
            set_value_label(ws, row, col_tier, '합계')
            write_metric_values(ws, row, metric0, agg)
            for tier in TIERS:
                r = row + tier
                set_value_label(ws, r, col_tier, tier)
                write_metric_values(ws, r, metric0, dm[tier])
            for r in range(dept_start, dept_start + 6):
                layout._metric_border(ws, r, metric0)
                ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
                set_label_border(ws, r, (col_month, col_g, col_tier))
            set_group_label(ws, dept_start, col_g, DEPT_SHORT[d])
            ws.merge_cells(start_row=dept_start, start_column=col_g, end_row=dept_start + 5, end_column=col_g)
            row = dept_start + 6

        # C열(전사계+개인+전략+신사업 전체를 잇는 파란 스파인)은 이 월 블록 전체에
        # 걸쳐 내부 가로줄 없이, 맨 위/맨 아래에만 테두리를 준다.
        for r in range(total_start, row):
            set_label_border(ws, r, (col_gw,), top=(r == total_start), bottom=(r == row - 1))

        set_value_label(ws, total_start, col_month, int(month) if str(month).isdigit() else month)
        ws.merge_cells(start_row=total_start, start_column=col_month, end_row=row - 1, end_column=col_month)
        if i < len(months) - 1:
            tall_rows.append(row)
            # D열(부문별)은 다음 달 '전사 계'가 C:D 2차원 병합이라 병합 좌상단(C)의
            # 위 테두리만 실제로 그려지고 D쪽은 비어 보인다 - 그 대신 바로 위
            # 빈 구분행의 D열에 아래 테두리만(좌우 없이) 줘서 선을 이어지게 한다.
            ws.cell(row=row, column=col_g).border = Border(top=None, bottom=THIN_SIDE, left=None, right=None)
            row += 1

    for r in range(1, row):
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'B': 5.625, 'C': 3.125, 'D': 11.375, 'E': 5.25, 'F': 8.375, 'I': 6.658, 'K': 6.658,
              'L': 10.375, 'M': 6.658, 'N': 8.375, 'O': 6.658, 'U': 8.625, 'V': 8.375,
              'W': 6.658, 'Z': 11.125}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 2) 연차별_부문별_상세 (전체 기간 합산, 부문 x 연차1~4)
# ---------------------------------------------------------------------------

def build_tier_dept_summary(wb, all_dept_acc, max_tier=4):
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

    per_dept = {d: all_dept_acc.finalize(d) for d in DEPT_ORDER}
    total_metrics = calc.sum_metrics([per_dept[d] for d in DEPT_ORDER])

    def agg_over(metrics_by_tier, tiers):
        a = {
            'target': sum(metrics_by_tier[t]['target'] for t in tiers),
            'accident': sum(metrics_by_tier[t]['accident'] for t in tiers),
            'done': sum(metrics_by_tier[t]['done'] for t in tiers),
            'plan': sum(metrics_by_tier[t]['plan'] for t in tiers),
            'codes': {cs: sum(metrics_by_tier[t]['codes'][cs] for t in tiers) for cs in CODE_STRS},
        }
        a['conv_target'] = a['target'] - a['accident']
        a['not_done'] = a['conv_target'] - a['done']
        a['idle'] = a['not_done'] - (a['plan'] + sum(a['codes'].values()))
        return a

    row = r0 + 3
    table_start = row
    for label in ['전사 계'] + DEPTS_SHORT:
        start = row
        metrics = total_metrics if label == '전사 계' else per_dept[DEPT_ORDER[DEPTS_SHORT.index(label)]]
        agg = agg_over(metrics, range(1, max_tier + 1))
        set_value_label(ws, row, col_tier, '합계')
        write_metric_values(ws, row, metric0, agg)
        for tier in range(1, max_tier + 1):
            r = row + tier
            set_value_label(ws, r, col_tier, tier)
            write_metric_values(ws, r, metric0, metrics[tier])
        is_total = label == '전사 계'
        for r in range(start, start + max_tier + 1):
            _metric_border(ws, r, metric0)
            ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
            set_label_border(ws, r, (col_g, col_tier))
            if is_total:
                _metric_fill(ws, r, metric0, TOTAL_FILL)
                ws.cell(row=r, column=col_tier).fill = TOTAL_FILL
        set_group_label(ws, start, col_gw if is_total else col_g, label)
        if is_total:
            ws.merge_cells(start_row=start, start_column=col_gw, end_row=start + max_tier, end_column=col_g)
            # 전사계 블록의 아래쪽 테두리는 없앤다(개인 블록으로 이어지는 스파인이
            # 끊기지 않도록) - 병합범위의 실제 렌더링 기준인 우하단 셀(col_g,
            # 마지막 행)의 bottom만 지우면 된다.
            bottom_cell = ws.cell(row=start + max_tier, column=col_g)
            b = bottom_cell.border
            bottom_cell.border = Border(top=b.top, bottom=None, left=b.left, right=b.right)
        else:
            ws.merge_cells(start_row=start, start_column=col_g, end_row=start + max_tier, end_column=col_g)
        row = start + max_tier + 1

    # B열(전사계+개인+전략+신사업 전체를 잇는 파란 스파인)은 맨 위/맨 아래에만
    # 테두리를 주고 내부는 없앤다.
    for r in range(table_start, row):
        set_label_border(ws, r, (col_gw,), top=(r == table_start), bottom=(r == row - 1))

    for r in range(r0 + 3, row):
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'B': 3.125, 'C': 11.375, 'D': 5.25, 'E': 9.375, 'H': 6.658, 'I': 9.375,
              'J': 6.658, 'K': 10.375, 'L': 6.658, 'M': 8.375, 'N': 6.658, 'T': 8.625,
              'U': 8.375, 'V': 6.658}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.row_dimensions[r0].height = 17.25
    return ws


def _agg_all_tiers(metrics_by_tier):
    a = {
        'target': sum(metrics_by_tier[t]['target'] for t in TIERS),
        'accident': sum(metrics_by_tier[t]['accident'] for t in TIERS),
        'done': sum(metrics_by_tier[t]['done'] for t in TIERS),
        'plan': sum(metrics_by_tier[t]['plan'] for t in TIERS),
        'codes': {cs: sum(metrics_by_tier[t]['codes'][cs] for t in TIERS) for cs in CODE_STRS},
    }
    a['conv_target'] = a['target'] - a['accident']
    a['not_done'] = a['conv_target'] - a['done']
    a['idle'] = a['not_done'] - (a['plan'] + sum(a['codes'].values()))
    return a


# ---------------------------------------------------------------------------
# 3) 월별_부문별 (연차 구분 없는 요약: 전체합산 블록 + 월별 블록)
# ---------------------------------------------------------------------------

def build_month_dept_summary(wb, months, month_dept_acc, all_dept_acc):
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
    write_metric_header(ws, r0, metric0, pct_header_fill=layout.PCT_FILL)

    row = r0 + 3
    tall_rows = [r0]

    from openpyxl.styles import Border, Side
    THIN = Side(style='thin')
    TOTAL_ROW_BORDER = Border(top=THIN, bottom=None, left=THIN, right=THIN)
    DEPT_LABEL_BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    DEPT_SIDE_BORDER = Border(top=None, bottom=None, left=THIN, right=None)

    def write_block(label, total_agg, dept_aggs):
        nonlocal row
        start = row
        set_group_label(ws, row, col_g, label, horizontal='right')
        ws.cell(row=row, column=col_g).border = TOTAL_ROW_BORDER
        ws.cell(row=row, column=col_g + 1).border = TOTAL_ROW_BORDER
        ws.merge_cells(start_row=row, start_column=col_g, end_row=row, end_column=col_g + 1)
        write_metric_values(ws, row, metric0, total_agg)
        _metric_border(ws, row, metric0)
        _pct_tint(ws, row, metric0)
        row += 1
        for i, (d, short) in enumerate(zip(DEPT_ORDER, DEPTS_SHORT)):
            is_last = i == len(DEPT_ORDER) - 1
            side_border = Border(top=None, bottom=THIN, left=THIN, right=None) if is_last else DEPT_SIDE_BORDER
            set_group_label(ws, row, col_g + 1, short, horizontal='right')
            ws.cell(row=row, column=col_g).border = side_border
            ws.cell(row=row, column=col_g + 1).border = DEPT_LABEL_BORDER
            write_metric_values(ws, row, metric0, dept_aggs[d])
            _metric_border(ws, row, metric0)
            _pct_tint(ws, row, metric0)
            row += 1
        return start

    overall_dept = {d: _agg_all_tiers(all_dept_acc.finalize(d)) for d in DEPT_ORDER}
    overall_total = calc.sum_metrics([all_dept_acc.finalize(d) for d in DEPT_ORDER])
    overall_total_agg = _agg_all_tiers(overall_total)
    write_block('전사 계', overall_total_agg, overall_dept)
    tall_rows.append(row)
    row += 1
    for i, month in enumerate(months):
        m_dept = {d: _agg_all_tiers(month_dept_acc.finalize((month, d))) for d in DEPT_ORDER}
        m_total = calc.sum_metrics([month_dept_acc.finalize((month, d)) for d in DEPT_ORDER])
        m_total_agg = _agg_all_tiers(m_total)
        write_block(f'{int(month) if str(month).isdigit() else month}월', m_total_agg, m_dept)
        if i < len(months) - 1:
            tall_rows.append(row)
            row += 1

    widths = {'A': 9.375, 'B': 3.125, 'F': 9.375, 'G': 6.658, 'I': 6.658, 'J': 10.625,
              'K': 6.658, 'M': 6.658, 'N': 8.75, 'O': 8.755, 'U': 6.658}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 4) 상품별_부문별_연차별 (상품명 x 부문 x 연차1~5)
# ---------------------------------------------------------------------------

def build_product_dept_tier(wb, products, product_dept_acc):
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
        per_dept = {d: product_dept_acc.finalize((p, d)) for d in DEPT_ORDER}
        total_metrics = calc.sum_metrics([per_dept[d] for d in DEPT_ORDER])
        for label in ['전사 계'] + DEPTS_SHORT:
            g_start = row
            metrics = total_metrics if label == '전사 계' else per_dept[DEPT_ORDER[DEPTS_SHORT.index(label)]]
            agg = _agg_all_tiers(metrics)
            set_value_label(ws, row, col_tier, '합계')
            write_metric_values(ws, row, metric0, agg)
            for tier in range(1, 6):
                r = row + tier
                set_value_label(ws, r, col_tier, tier)
                write_metric_values(ws, r, metric0, metrics[tier])
            is_total = label == '전사 계'
            for r in range(g_start, g_start + 6):
                _metric_border(ws, r, metric0)
                ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
                set_label_border(ws, r, (col_p, col_g, col_tier))
                if is_total:
                    _metric_fill(ws, r, metric0, TOTAL_FILL)
                    ws.cell(row=r, column=col_tier).fill = TOTAL_FILL
            if is_total:
                ws.merge_cells(start_row=g_start, start_column=col_gw, end_row=g_start + 5, end_column=col_g)
                bottom_cell = ws.cell(row=g_start + 5, column=col_g)
                b = bottom_cell.border
                bottom_cell.border = Border(top=b.top, bottom=None, left=b.left, right=b.right)
                set_group_label(ws, g_start, col_gw, label)
            else:
                ws.merge_cells(start_row=g_start, start_column=col_g, end_row=g_start + 5, end_column=col_g)
                set_group_label(ws, g_start, col_g, label)
            row = g_start + 6
        set_value_label(ws, p_start, col_p, p)
        for r in range(p_start, row):
            set_label_border(ws, r, (col_p,))
            set_label_border(ws, r, (col_gw,), top=(r == p_start), bottom=(r == row - 1))
        ws.merge_cells(start_row=p_start, start_column=col_p, end_row=row - 1, end_column=col_p)
        tall_rows.append(row)
        row += 1

    for r in range(r0 + 3, row):
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'B': 8.375, 'C': 3.125, 'D': 11.375, 'E': 5.25, 'F': 8.375, 'I': 6.658, 'K': 6.658,
              'L': 10.375, 'M': 6.658, 'N': 8.375, 'O': 6.658, 'U': 8.625, 'V': 8.375, 'W': 6.658}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 5) 상품별_부문별_합산 (상품명 x 부문, 연차 구분 없음)
# ---------------------------------------------------------------------------

def build_product_dept_summary(wb, products, product_dept_acc):
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
        per_dept = {d: product_dept_acc.finalize((p, d)) for d in DEPT_ORDER}
        total_metrics = calc.sum_metrics([per_dept[d] for d in DEPT_ORDER])
        for label in ['전사 계'] + DEPTS_SHORT:
            is_total = label == '전사 계'
            label_col = col_gw if is_total else col_g
            metrics = total_metrics if is_total else per_dept[DEPT_ORDER[DEPTS_SHORT.index(label)]]
            agg = _agg_all_tiers(metrics)
            set_group_label(ws, row, label_col, label, horizontal='right')
            if is_total:
                ws.merge_cells(start_row=row, start_column=col_gw, end_row=row, end_column=col_g)
            write_metric_values(ws, row, metric0, agg)
            ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
            set_label_border(ws, row, (col_g,))
            if is_total:
                _metric_fill(ws, row, metric0, TOTAL_FILL)
            _metric_border(ws, row, metric0)
            row += 1
        set_value_label(ws, p_start, col_p, p)
        for r in range(p_start, row):
            set_label_border(ws, r, (col_p,))
            set_label_border(ws, r, (col_gw,), top=(r == p_start), bottom=(r == row - 1))
        ws.merge_cells(start_row=p_start, start_column=col_p, end_row=row - 1, end_column=col_p)
        tall_rows.append(row)
        row += 1

    widths = {'B': 8.375, 'C': 3.125, 'D': 11.375, 'E': 8.375, 'H': 6.658, 'J': 6.658, 'K': 10.375,
              'L': 6.658, 'M': 8.375, 'N': 6.658, 'T': 8.625, 'U': 8.375, 'V': 6.658}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 6) 쳬결기간별_부문별
# ---------------------------------------------------------------------------

def _sum_period_parts(parts):
    target = sum(p['target'] for p in parts)
    accident = sum(p['accident'] for p in parts)
    done = sum(p['done'] for p in parts)
    max_done = sum(p['max_done'] for p in parts)
    plan = sum(p['plan'] for p in parts)
    codes = {cs: sum(p['codes'][cs] for p in parts) for cs in CODE_STRS}
    conv_target = target - accident
    not_done = conv_target - done
    idle = not_done - (plan + sum(codes.values()))
    origin_keys = set()
    for p in parts:
        origin_keys |= set(p.get('not_done_origin', {}).keys())
    not_done_origin = {j: sum(p.get('not_done_origin', {}).get(j, 0) for p in parts) for j in origin_keys}
    return {'target': target, 'accident': accident, 'conv_target': conv_target,
            'done': done, 'max_done': max_done, 'not_done': not_done,
            'plan': plan, 'codes': codes, 'idle': idle, 'not_done_origin': not_done_origin}


def _period_metrics_for(period_acc, dept, period_key, tier):
    if dept == '전사 계':
        return _sum_period_parts([period_acc.get(d, period_key, tier) for d in DEPT_ORDER])
    return period_acc.get(dept, period_key, tier)


def _period_reach(period_acc, dept, period_key):
    if dept == '전사 계':
        s = set()
        for d in DEPT_ORDER:
            s |= set(period_acc.rows_for(d, period_key))
        return sorted(s)
    return period_acc.rows_for(dept, period_key)


def _period_overall(period_acc, dept, period_key, tiers):
    parts = [_period_metrics_for(period_acc, dept, period_key, t) for t in tiers]
    return _sum_period_parts(parts)


def build_period_dept(wb, period_acc):
    """PERIODS 순서(오래된 구간부터)로 실제 도달 연차만큼만 행을 만든다."""
    ws = wb.create_sheet('쳬결기간별_부문별')
    ws.sheet_view.showGridLines = False
    col_gw, col_g, col_period, col_key, col_tier = 2, 3, 4, 5, 6
    metric0 = 7

    r0 = 7
    write_title_and_stamp(ws, '체결기간별_부문별_무사고전환율', 2, metric0 + METRIC_WIDTH_PERIOD - 1, stamp_row=r0 - 1)
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
    write_metric_header_period(ws, r0, metric0, target_label='대상계약\nA')

    row = r0 + 3
    tall_rows = [r0]
    table_start = row
    for label, lookup_key in [('전사 계', '전사 계')] + list(zip(DEPTS_SHORT, DEPT_ORDER)):
        g_start = row
        is_total = label == '전사 계'
        for plabel, key, _start, _end in PERIODS:
            reach_tiers = _period_reach(period_acc, lookup_key, key)
            if not reach_tiers:
                continue
            p_start = row
            set_value_label(ws, row, col_period, plabel)
            set_value_label(ws, row, col_key, '전체')
            overall = _period_overall(period_acc, lookup_key, key, reach_tiers)
            write_period_values(ws, row, metric0, overall)
            _metric_border_period(ws, row, metric0)
            ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
            set_label_border(ws, row, (col_g, col_period, col_tier, col_key))
            if len(reach_tiers) > 1:
                ws.merge_cells(start_row=row, start_column=col_key, end_row=row, end_column=col_tier)
                # E,F 병합 셀 자체의 아래 테두리는 없앤다(병합 셀은 우하단 기준
                # 하나로만 그려지므로 E만 다르게 줄 수 없음) - 대신 바로 아래
                # 행(첫 연차 서브행)의 F열 위 테두리만으로 선을 표현한다.
                set_label_border(ws, row, (col_key, col_tier), top=True, bottom=False)
                row += 1
                for i, t in enumerate(reach_tiers):
                    is_last_tier = i == len(reach_tiers) - 1
                    set_value_label(ws, row, col_tier, t)
                    write_period_values(ws, row, metric0, _period_metrics_for(period_acc, lookup_key, key, t))
                    _metric_border_period(ws, row, metric0)
                    ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
                    set_label_border(ws, row, (col_g, col_period, col_tier))
                    set_label_border(ws, row, (col_key,), top=False, bottom=is_last_tier)
                    row += 1
            else:
                set_value_label(ws, row, col_tier, reach_tiers[0])
                set_label_border(ws, row, (col_g, col_period, col_tier))
                row += 1
            ws.merge_cells(start_row=p_start, start_column=col_period,
                            end_row=row - 1, end_column=col_period)
        if is_total:
            ws.merge_cells(start_row=g_start, start_column=col_gw, end_row=row - 1, end_column=col_g)
            bottom_cell = ws.cell(row=row - 1, column=col_g)
            b = bottom_cell.border
            bottom_cell.border = Border(top=b.top, bottom=None, left=b.left, right=b.right)
            set_group_label(ws, g_start, col_gw, label)
        else:
            ws.merge_cells(start_row=g_start, start_column=col_g, end_row=row - 1, end_column=col_g)
            set_group_label(ws, g_start, col_g, label)

    # B열(전사계+개인+전략+신사업 전체를 잇는 파란 스파인)은 표 전체에 걸쳐
    # 내부 가로줄 없이, 맨 위/맨 아래에만 테두리를 준다.
    for r in range(table_start, row):
        set_label_border(ws, r, (col_gw,), top=(r == table_start), bottom=(r == row - 1))

    for r in range(r0 + 3, row):
        ws.cell(row=r, column=col_key).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'A': 11.5, 'B': 3.125, 'C': 11.375, 'D': 16.25, 'E': 5.25, 'G': 8.375,
              'I': 8.625, 'J': 6.658, 'K': 8.625, 'L': 6.658, 'M': 10.375, 'N': 6.658,
              'O': 10.375, 'P': 6.658, 'R': 6.658, 'X': 8.625, 'Y': 8.375, 'Z': 6.658}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


# ---------------------------------------------------------------------------
# 7) 체결기간별_부문별_잔여미완료 (구간x부문 한 행에 연차1~4 미완료를 나란히 배치
#    - 연차가 올라갈수록 미완료가 늘어나는 추세를 한눈에 보기 위한 보조 시트)
#    사용자가 직접 재구성한 레이아웃: 기존 20칸 지표블록 중 코드별 세부내역
#    (현장활동확인/전환예정/5개코드/현장활동미확인)은 빼고, 대상계약~전환미완료
#    까지 10칸으로 압축한 뒤 바로 뒤에 1~4연차 미완료(건수+%) 4쌍을 붙인다.
# ---------------------------------------------------------------------------

METRIC_WIDTH_PERIOD_COMPACT = 10


def write_metric_header_period_compact(ws, r0, col0, target_label='대상계약\nA',
                                        pct_header_fill=layout.PCT_HEADER_FILL):
    THIN, MEDIUM = layout.THIN, layout.MEDIUM

    def full3(col, label, right_border=THIN):
        ws.merge_cells(start_row=r0, start_column=col, end_row=r0 + 2, end_column=col)
        box = Border(top=THIN, bottom=None, left=THIN, right=right_border)
        for rr in range(r0, r0 + 3):
            _hcell(ws, rr, col, border=box)
        ws.cell(row=r0, column=col, value=label)

    full3(col0, target_label)
    full3(col0 + 1, '사고有\nB')
    full3(col0 + 2, '전환대상\nC (A-B)', right_border=None)
    full3(col0 + 4, '전환완료\nD', right_border=None)
    full3(col0 + 8, '전환미완료\nE (C-D)', right_border=None)

    for off in (3, 5, 9):
        col = col0 + off
        label_border = layout._TYPE_A_LABEL if off != 9 else Border(top=THIN, bottom=None, left=THIN, right=THIN)
        _hcell(ws, r0, col, border=layout._TYPE_A_TOP)
        _hcell(ws, r0 + 1, col, border=layout._TYPE_A_MID)
        _hcell(ws, r0 + 2, col, '%', fill=pct_header_fill, border=label_border)

    m_col = col0 + 6
    _hcell(ws, r0, m_col, border=layout._TYPE_A_TOP)
    ws.merge_cells(start_row=r0 + 1, start_column=m_col, end_row=r0 + 2, end_column=m_col)
    _hcell(ws, r0 + 1, m_col, '최대\n전환완료', border=layout._TYPE_A_LABEL)
    _hcell(ws, r0 + 2, m_col, border=layout._ACT_MERGED_BOTTOM)

    n_col = col0 + 7
    _hcell(ws, r0, n_col, border=layout._TYPE_B_TOP)
    _hcell(ws, r0 + 1, n_col, border=layout._TYPE_B_MID)
    _hcell(ws, r0 + 2, n_col, '%', fill=pct_header_fill, border=layout._TYPE_B_LABEL)

    medium_left = {col0 + 4, col0 + 8}
    medium_right = {col0 + 7}
    top_from = col0 + 4
    for rr in range(r0, r0 + 3):
        for c in range(col0, col0 + METRIC_WIDTH_PERIOD_COMPACT):
            cell = ws.cell(row=rr, column=c)
            b = cell.border
            left = MEDIUM if c in medium_left else b.left
            right = MEDIUM if c in medium_right else b.right
            top = MEDIUM if (rr == r0 and c >= top_from) else b.top
            cell.border = Border(top=top, bottom=b.bottom, left=left, right=right)

    for off in (6, 7):
        col = col0 + off
        cell = ws.cell(row=r0 + 1, column=col)
        b = cell.border
        cell.border = Border(top=MEDIUM, bottom=b.bottom, left=b.left, right=b.right)

    for rr in (r0 + 1, r0 + 2):
        cell = ws.cell(row=rr, column=m_col)
        b = cell.border
        cell.border = Border(top=b.top, bottom=b.bottom, left=MEDIUM, right=b.right)


def _metric_border_period_compact(ws, row, col0, top=True, bottom=True):
    THIN, MEDIUM = layout.THIN, layout.MEDIUM
    medium_left = {col0 + 4, col0 + 6, col0 + 8}
    medium_right = {col0 + 5, col0 + 7}
    for c in range(col0, col0 + METRIC_WIDTH_PERIOD_COMPACT):
        left = MEDIUM if c in medium_left else THIN
        right = MEDIUM if c in medium_right else THIN
        ws.cell(row=row, column=c).border = Border(
            top=THIN if top else None, bottom=THIN if bottom else None, left=left, right=right)


def write_period_values_compact(ws, row, col0, m):
    target, accident, conv_target = m['target'], m['accident'], m['conv_target']
    done, max_done, not_done = m['done'], m['max_done'], m['not_done']

    def setv(off, value, is_pct=False):
        cell = ws.cell(row=row, column=col0 + off, value=value)
        cell.number_format = RATIO_FMT if is_pct else COUNT_FMT
        cell.font = VALUE_FONT_10
        cell.alignment = Alignment(horizontal='right', vertical='center')

    setv(0, target)
    setv(1, accident)
    setv(2, conv_target)
    setv(3, _ratio(conv_target, target), True)
    setv(4, done)
    setv(5, _ratio(done, conv_target), True)
    setv(6, max_done)
    setv(7, _ratio(max_done, conv_target), True)
    setv(8, not_done)
    setv(9, _ratio(not_done, conv_target), True)


def write_remaining_header(ws, r0, col0, n_tiers=4):
    """col0부터 1~n_tiers연차 미완료(건수+%) 쌍을 이어 붙인다."""
    THIN, MEDIUM = layout.THIN, layout.MEDIUM
    for i in range(n_tiers):
        label_col = col0 + i * 2
        pct_col = label_col + 1
        is_last = i == n_tiers - 1
        ws.merge_cells(start_row=r0 + 1, start_column=label_col, end_row=r0 + 2, end_column=label_col)

        _hcell(ws, r0, label_col, border=Border(top=MEDIUM, bottom=None, left=None, right=None))
        _hcell(ws, r0 + 1, label_col, f'{i + 1}연차\n미완료',
               border=Border(top=THIN, bottom=THIN, left=THIN, right=THIN))
        _hcell(ws, r0 + 2, label_col, border=Border(top=None, bottom=THIN, left=THIN, right=THIN))

        edge_right = MEDIUM if is_last else None
        _hcell(ws, r0, pct_col, border=Border(top=MEDIUM, bottom=None, left=None, right=edge_right))
        edge_right = MEDIUM if is_last else THIN
        _hcell(ws, r0 + 1, pct_col, border=Border(top=THIN, bottom=THIN, left=None, right=edge_right))
        _hcell(ws, r0 + 2, pct_col, '%', fill=layout.PCT_HEADER_FILL,
               border=Border(top=THIN, bottom=THIN, left=THIN, right=edge_right))


def _remaining_border(ws, row, col0, n_tiers=4):
    THIN, MEDIUM = layout.THIN, layout.MEDIUM
    width = n_tiers * 2
    for j in range(width):
        c = col0 + j
        right = MEDIUM if j == width - 1 else THIN
        ws.cell(row=row, column=c).border = Border(top=THIN, bottom=THIN, left=THIN, right=right)


def write_remaining_origin_values(ws, row, col0, m, up_to_tier):
    """이 행이 대표하는 연차(up_to_tier)까지, 그 연차 버킷의 미완료를 '언제부터
    뒤처지기 시작했는지'(기원 연차)로 쪼개서 1~up_to_tier열에 채운다. 예를 들어
    3연차 행이면 1~3연차 기원 3칸만 채우고 나머지(4연차 칸)는 비워 둔다."""
    conv_target = m['conv_target']
    origin = m.get('not_done_origin', {})
    for j in range(1, up_to_tier + 1):
        count_col = col0 + (j - 1) * 2
        pct_col = count_col + 1
        value = origin.get(j, 0)
        cc = ws.cell(row=row, column=count_col, value=value)
        cc.number_format = COUNT_FMT
        cc.font = VALUE_FONT_10
        cc.alignment = Alignment(horizontal='right', vertical='center')
        pc = ws.cell(row=row, column=pct_col, value=_ratio(value, conv_target))
        pc.number_format = RATIO_FMT
        pc.font = VALUE_FONT_10
        pc.alignment = Alignment(horizontal='right', vertical='center')


def build_period_dept_remaining(wb, period_acc):
    ws = wb.create_sheet('체결기간별_부문별_잔여미완료')
    ws.sheet_view.showGridLines = False
    col_gw, col_g, col_period, col_key, col_tier = 2, 3, 4, 5, 6
    metric0 = 7
    ext0 = metric0 + METRIC_WIDTH_PERIOD_COMPACT  # 새로 덧붙이는 연차별 미완료 블록의 시작 열(Q)

    r0 = 7
    write_title_and_stamp(ws, '체결기간별_부문별_잔여미완료', 2, ext0 + 8 - 1, stamp_row=r0 - 1)
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
    write_metric_header_period_compact(ws, r0, metric0, target_label='대상계약\nA')

    remaining_tiers = [1, 2, 3, 4]
    write_remaining_header(ws, r0, ext0, n_tiers=len(remaining_tiers))

    row = r0 + 3
    tall_rows = [r0]
    table_start = row
    for label, lookup_key in [('전사 계', '전사 계')] + list(zip(DEPTS_SHORT, DEPT_ORDER)):
        g_start = row
        is_total = label == '전사 계'
        for plabel, key, _start, _end in PERIODS:
            reach_tiers = _period_reach(period_acc, lookup_key, key)
            if not reach_tiers:
                continue
            p_start = row
            set_value_label(ws, row, col_period, plabel)
            set_value_label(ws, row, col_key, '전체')
            overall = _period_overall(period_acc, lookup_key, key, reach_tiers)
            write_period_values_compact(ws, row, metric0, overall)
            _metric_border_period_compact(ws, row, metric0)
            _remaining_border(ws, row, ext0, n_tiers=len(remaining_tiers))
            ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
            set_label_border(ws, row, (col_g, col_period, col_tier, col_key))
            if len(reach_tiers) > 1:
                # '전체' 행은 여러 연차를 합친 값이라 기원 연차별로 다시 쪼개는 게
                # 의미가 없으므로(연차별 미완료가 이미 아래 서브행에 나온다) 여기는
                # 비워 두고, 아래 연차별 서브행에서 각자 자기 연차까지의 기원별
                # 분해(1연차행=1칸, 2연차행=1~2칸, ...)를 채운다.
                ws.merge_cells(start_row=row, start_column=col_key, end_row=row, end_column=col_tier)
                set_label_border(ws, row, (col_key, col_tier), top=True, bottom=False)
                row += 1
                for i, t in enumerate(reach_tiers):
                    is_last_tier = i == len(reach_tiers) - 1
                    set_value_label(ws, row, col_tier, t)
                    tm = _period_metrics_for(period_acc, lookup_key, key, t)
                    write_period_values_compact(ws, row, metric0, tm)
                    _metric_border_period_compact(ws, row, metric0)
                    _remaining_border(ws, row, ext0, n_tiers=len(remaining_tiers))
                    write_remaining_origin_values(ws, row, ext0, tm, up_to_tier=t)
                    ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
                    set_label_border(ws, row, (col_g, col_period, col_tier))
                    set_label_border(ws, row, (col_key,), top=False, bottom=is_last_tier)
                    row += 1
            else:
                # 연차가 하나뿐이면 이 행이 곧 그 연차의 행이므로 여기에 바로
                # 기원별 분해(1칸)를 채운다.
                write_remaining_origin_values(ws, row, ext0, overall, up_to_tier=reach_tiers[0])
                set_value_label(ws, row, col_tier, reach_tiers[0])
                set_label_border(ws, row, (col_g, col_period, col_tier))
                row += 1
            ws.merge_cells(start_row=p_start, start_column=col_period,
                            end_row=row - 1, end_column=col_period)
        if is_total:
            ws.merge_cells(start_row=g_start, start_column=col_gw, end_row=row - 1, end_column=col_g)
            bottom_cell = ws.cell(row=row - 1, column=col_g)
            b = bottom_cell.border
            bottom_cell.border = Border(top=b.top, bottom=None, left=b.left, right=b.right)
            set_group_label(ws, g_start, col_gw, label)
        else:
            ws.merge_cells(start_row=g_start, start_column=col_g, end_row=row - 1, end_column=col_g)
            set_group_label(ws, g_start, col_g, label)

    for r in range(table_start, row):
        set_label_border(ws, r, (col_gw,), top=(r == table_start), bottom=(r == row - 1))

    for r in range(r0 + 3, row):
        ws.cell(row=r, column=col_key).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=r, column=col_tier).alignment = Alignment(horizontal='center', vertical='center')

    widths = {'A': 11.5, 'B': 3.125, 'C': 11.375, 'D': 16.25, 'E': 5.25, 'G': 8.375,
              'I': 8.625, 'J': 6.625, 'K': 8.625, 'L': 6.625, 'M': 10.375, 'N': 6.625,
              'O': 10.375, 'P': 6.625}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for i in range(len(remaining_tiers)):
        count_letter = openpyxl.utils.get_column_letter(ext0 + i * 2)
        pct_letter = openpyxl.utils.get_column_letter(ext0 + i * 2 + 1)
        ws.column_dimensions[count_letter].width = 7.125
        ws.column_dimensions[pct_letter].width = 6.625
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    return ws


def build_glossary_sheet(wb):
    """'항목설명' 시트: 보고서를 처음 보는 사람도 각 항목/시트의 정의를 이해할
    수 있도록 정리한 용어집. 실제 계산 로직에는 관여하지 않는 참고용 시트."""
    from openpyxl.styles import Font, PatternFill, Border, Side
    ws = wb.create_sheet('항목설명')
    ws.sheet_view.showGridLines = False

    TITLE_F = Font(name='맑은 고딕', size=16, bold=True)
    SEC_F = Font(name='맑은 고딕', size=12, bold=True, color='FFFFFF')
    SEC_FILL = PatternFill('solid', fgColor='4472C4')
    HEAD_F = Font(name='맑은 고딕', size=10, bold=True)
    HEAD_FILL = PatternFill('solid', fgColor='D9E1F2')
    BODY_F = Font(name='맑은 고딕', size=10)
    THIN = Side(style='thin', color='BFBFBF')
    BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

    row = [1]  # mutable row cursor

    def title(text):
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c = ws.cell(row=r, column=2, value=text)
        c.font = TITLE_F
        c.alignment = Alignment(horizontal='left', vertical='center')
        row[0] += 2

    def section(text):
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c = ws.cell(row=r, column=2, value=text)
        c.font = SEC_F
        c.fill = SEC_FILL
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 22
        row[0] += 1

    def table_header(cols):
        r = row[0]
        positions = [2, 3, 5]
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        for col, label in zip(positions, cols):
            c = ws.cell(row=r, column=col, value=label)
            c.font = HEAD_F
            c.fill = HEAD_FILL
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center')
        for col in (3, 4):
            ws.cell(row=r, column=col).fill = HEAD_FILL
            ws.cell(row=r, column=col).border = BORDER
        row[0] += 1

    def term_row(term, desc, note=''):
        r = row[0]
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        for col, val in [(2, term), (3, desc), (5, note)]:
            c = ws.cell(row=r, column=col, value=val)
            c.font = BODY_F
            c.border = BORDER
            c.alignment = Alignment(horizontal='left' if col != 2 else 'center',
                                     vertical='center', wrap_text=True)
        d = ws.cell(row=r, column=4)
        d.font = BODY_F
        d.border = BORDER
        ws.row_dimensions[r].height = 30
        row[0] += 1

    def para(text):
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        for col in range(2, 6):
            cc = ws.cell(row=r, column=col)
            cc.font = BODY_F
            cc.border = BORDER
        c = ws.cell(row=r, column=2, value=text)
        c.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        ws.row_dimensions[r].height = 45
        row[0] += 1

    def blank(n=1):
        row[0] += n

    title('무사고전환 보고서 - 항목 및 시트 설명')
    para('이 시트는 보고서에 나오는 지표와 시트별 구성 기준을 설명하는 참고용 안내문입니다. '
         '실제 데이터/계산에는 영향을 주지 않습니다.')
    blank()

    section('1. 공통 지표 정의')
    table_header(['용어', '정의', '비고'])
    term_row('유지계약 (A)', '기준 시점에 유지 중인 전체 계약 건수.')
    term_row('사고有 (B)', '무사고 조건을 충족하지 못해(사고가 발생해) 전환 대상에서 제외되는 계약 건수.')
    term_row('전환대상 (C = A－B)', '사고가 없어 실제로 플랜 전환 대상이 되는 계약 건수.')
    term_row('전환완료 (D)', '전환대상 중 정해진 목표플랜으로 실제 전환이 완료된 계약 건수.')
    term_row('전환미완료 (E = C－D)', '전환대상 중 아직 목표플랜으로 전환되지 않은 계약 건수.')
    term_row('현장활동확인', '전환미완료 건 중, 담당자가 실제로 접촉/처리를 시도한 기록이 있는 건수 (전환예정 + 코드5종 합계).')
    term_row('현장활동미확인', '전환미완료 건 중, 아직 어떤 처리 기록도 없는 건수.', '전환미완료－현장활동확인')
    term_row('전환예정', '특별한 처리 사유 코드 없이 전환을 계속 진행/대기 중인 건수.')
    term_row('코드5종', '연락두절 / 사고있음 / 고객거부 / 압류계약 / ARS거부 - 담당자가 접촉을 시도했지만 특정 사유로 전환이 보류된 건수.')
    blank()

    section('2. 최대전환완료 (체결기간별_부문별 시트 전용)')
    para('상품/플랜마다 "몇 연차까지 전환 목표가 정의되어 있는지"(최대전환년수)가 정해져 있습니다. '
         '전환완료(D) 건 중에서도, 그 연차가 계약의 마지막 전환 단계였던 것 - 즉 더 이상 다음 연차로 '
         '추적할 목표가 없는 건 - 만 따로 뽑은 것이 "최대전환완료"입니다.')
    para('이 값을 그 연차의 전환대상(C)에서 빼면 다음 연차의 대상계약 수와 정확히 같아집니다. '
         '즉 "연차가 올라갈수록 대상계약이 왜 줄어드는지"를 설명해주는 숫자입니다. '
         '(단, 최근 구간은 아직 실제 경과연수가 부족한 계약이 섞여 있어 정확히 일치하지 않을 수 있습니다.)')
    blank()

    section('3. 시트별 보는 법')
    table_header(['시트명', '설명', ''])
    term_row('월별_부문별', '1~12월 각 월 스냅샷 기준, 부문별(개인/전략/신사업) 현황을 연차 구분 없이 합산.')
    term_row('연차별_부문별_상세', '연차(1~4년차)별로 부문별 현황을 12개월 전체 합산.')
    term_row('월별_부문별_연차별', '월별 x 부문별 x 연차별(1~5년차)로 가장 세분화한 현황.')
    term_row('체결기간별_부문별',
             '보험기간 시작일 기준 12개월 단위 "체결 구간"별로, 그 구간에 가입한 계약들이 실제 경과연수에 '
             '따라 연차별로 어떻게 전환 진행 중인지 추적. 다른 시트와 달리 계약 하나가 여러 연차에 동시에 '
             '걸쳐 나타남(사고 발생 또는 최대전환완료 시점까지).')
    term_row('체결기간별_부문별_잔여미완료',
             '체결기간별_부문별과 동일한 구간/연차 표에, "이 미완료 건이 몇 연차부터 뒤처지기 시작했는지"를 '
             '연차별로 쪼개서 보여주는 보조 시트. 아래 5번 참고.')
    term_row('상품별_부문별_연차별', '상품(보험상품)별 x 부문별 x 연차별 현황.')
    term_row('상품별_부문별_합산', '상품별 x 부문별 현황을 연차 구분 없이 합산.')
    blank()

    section('4. 체결기간별_부문별 읽는 법 (가장 복잡한 시트)')
    para('구간이란: 계약체결월이 아니라 실제 보험기간 시작일 기준으로 12개월 단위로 나눈 구간입니다 '
         '(예: 2022.07~2023.06).')
    para('왜 연차가 올라갈수록 대상계약이 줄어드나: 두 가지 이유가 있습니다. '
         '(1) 최대전환완료 - 그 계약의 전환 스케줄이 이미 끝남(위 2번 참고). '
         '(2) 아직 실제 경과연수가 그 연차에 도달하지 않은 계약이 있음(구간 끝자락에 가입한 계약).')
    para('최대전환년수란: 최대전환년수는 "몇 년 뒤에 추적을 끊는다"는 컷오프가 아니라 "총 몇 단계의 전환 '
         '목표가 정의되어 있는가"를 뜻합니다. 전환을 완료하지 못한 계약은 최대전환년수를 넘겨도 실제 '
         '경과연수가 허용하는 한 계속 대상에 남아 추적됩니다. 추적이 끝나는 경우는 오직 (1)사고 발생 '
         '또는 (2)최대전환년수 이상 도달한 연차에서 실제로 목표플랜에 도달(=최대전환완료), 이 두 가지뿐입니다.')
    blank()

    section('5. 체결기간별_부문별_잔여미완료 읽는 법')
    para('이 시트는 체결기간별_부문별의 "전환미완료(E)"가 매 연차 왜 쌓이는지를 "몇 연차부터 뒤처지기 '
         '시작했는지" 기준으로 나눠 보여줍니다. 계약별 목표플랜은 연차가 올라갈수록 더 높은 단계를 '
         '요구하기 때문에, 어느 해에는 목표를 채워서 "완료"였던 계약도 그 다음 해 기준으로는 목표에 '
         '못 미쳐 다시 "미완료"로 잡힐 수 있습니다. 이렇게 한 계약이 최초로 목표에 못 미치기 시작한 '
         '연차를 그 계약의 "기원 연차"라고 부릅니다.')
    para('표를 읽는 법: 각 구간의 "1연차" 행에는 1연차미완료 칸(=1연차부터 계속 미전환인 계약 수) 하나만 '
         '채워지고, "2연차" 행에는 1연차미완료(1연차부터 2연차까지 계속 미전환) + 2연차미완료(1연차엔 '
         '완료였지만 2연차 기준으론 다시 미전환) 두 칸이 채워집니다. 3연차·4연차 행도 같은 방식으로 '
         '한 칸씩 늘어납니다. 아직 도달하지 않은 연차의 칸은 비워 둡니다.')
    para('검산: 한 행에 채워진 기원 연차별 칸을 모두 더하면, 그 행(그 연차)의 체결기간별_부문별 쪽 '
         '"전환미완료(E)" 값과 정확히 같습니다. "전체"(여러 연차를 합친) 행은 이 분해 기준 자체가 '
         '적용되지 않으므로 비워 둡니다.')

    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 22
    return ws


def build(data_paths, mapping_path, output_path):
    print('[1/2] 데이터 계산 중(파이썬 직접 계산, LibreOffice 미사용)...')
    result = calc.compute_all(data_paths, mapping_path)

    print('[2/2] 보고서 시트 작성 중...')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_glossary_sheet(wb)
    build_month_dept_summary(wb, result['months'], result['month_dept'], result['all_dept'])
    build_tier_dept_summary(wb, result['all_dept'])
    build_month_dept_tier(wb, result['months'], result['month_dept'], result['all_dept'])
    build_period_dept_remaining(wb, result['period_acc'])
    build_period_dept(wb, result['period_acc'])
    build_product_dept_tier(wb, result['products'], result['product_dept'])
    build_product_dept_summary(wb, result['products'], result['product_dept'])
    wb.save(output_path)
    print('저장 완료 ->', output_path)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python write_report_v7.py <출력.xlsx> <데이터파일1> [데이터파일2 ...]')
        sys.exit(1)
    output_path = sys.argv[1]
    data_paths = sys.argv[2:]
    build(data_paths, data_paths[0], output_path)
