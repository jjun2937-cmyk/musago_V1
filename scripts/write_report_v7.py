"""
write_report_v7.py — calc_report_v7.py로 계산한 값을, build_report_v7.py가 만든
레이아웃(서식/병합/색상)에 실제로 채워 넣어 최종 보고서 파일을 만든다.
LibreOffice/Excel 수식은 전혀 쓰지 않는다(순수 값만 기록).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.styles import Alignment

import build_report_v7 as layout
from build_report_v7 import (
    METRIC_WIDTH, PCT_OFFSETS, TOTAL_FILL, DEPTS_SHORT,
    write_metric_header, _metric_border, _metric_fill, _pct_tint, write_title_and_stamp, _hcell,
)
import calc_report_v7 as calc
from calc_report_v7 import DEPT_ORDER, DEPT_SHORT, CODE_MAP, TIERS, PERIODS

COUNT_FMT = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
RATIO_FMT = '_-* #,##0.0_-;-* #,##0.0_-;_-* "-"_-;_-@_-'
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
        cell.alignment = Alignment(horizontal='right', vertical='center')

    setv(0, target)
    setv(1, accident)
    setv(2, conv_target)
    setv(3, _ratio(conv_target, target), True)
    setv(4, done)
    setv(5, _ratio(done, conv_target), True)
    setv(6, not_done)
    setv(7, _ratio(not_done, conv_target), True)
    setv(9, _ratio(act_sum, conv_target), True)
    setv(10, plan)
    for i, (_cn, cs) in enumerate(CODE_MAP):
        setv(11 + i, codes.get(cs, 0))
    setv(16, idle)
    setv(17, _ratio(idle, conv_target), True)


def write_period_values(ws, row, col0, m):
    """체결기간별: A/B/C/D/E만 존재(현장활동확인 관련은 이 계산에서 산출하지 않음)."""
    target, accident, conv_target = m['target'], m['accident'], m['conv_target']
    done, not_done = m['done'], m['not_done']

    def setv(off, value, is_pct=False):
        cell = ws.cell(row=row, column=col0 + off, value=value)
        cell.number_format = RATIO_FMT if is_pct else COUNT_FMT
        cell.alignment = Alignment(horizontal='right', vertical='center')

    setv(0, target)
    setv(1, accident)
    setv(2, conv_target)
    setv(3, _ratio(conv_target, target), True)
    setv(4, done)
    setv(5, _ratio(done, conv_target), True)
    setv(6, not_done)
    setv(7, _ratio(not_done, conv_target), True)


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
        ws.cell(row=row, column=col_tier, value='합계')
        write_metric_values(ws, row, metric0, agg)
        for r in range(row, row + 6):
            layout._metric_fill(ws, r, metric0, TOTAL_FILL)
            layout._metric_border(ws, r, metric0)
            ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
            ws.cell(row=r, column=col_tier).fill = TOTAL_FILL
        for t in TIERS:
            ws.cell(row=row + t, column=col_tier, value=t)
            write_metric_values(ws, row + t, metric0, total_metrics[t])
        ws.cell(row=row, column=col_gw, value='전사 계')
        ws.merge_cells(start_row=row, start_column=col_gw, end_row=row + 5, end_column=col_g)
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
            ws.cell(row=row, column=col_tier, value='합계')
            write_metric_values(ws, row, metric0, agg)
            for tier in TIERS:
                r = row + tier
                ws.cell(row=r, column=col_tier, value=tier)
                write_metric_values(ws, r, metric0, dm[tier])
            for r in range(dept_start, dept_start + 6):
                layout._metric_border(ws, r, metric0)
                ws.cell(row=r, column=col_gw).fill = TOTAL_FILL
            ws.cell(row=dept_start, column=col_g, value=d)
            ws.merge_cells(start_row=dept_start, start_column=col_g, end_row=dept_start + 5, end_column=col_g)
            row = dept_start + 6

        ws.cell(row=total_start, column=col_month, value=int(month) if str(month).isdigit() else month)
        ws.merge_cells(start_row=total_start, start_column=col_month, end_row=row - 1, end_column=col_month)
        ws.cell(row=total_start, column=col_month).alignment = Alignment(horizontal='center', vertical='center')
        if i < len(months) - 1:
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
    for label in ['전사 계'] + DEPTS_SHORT:
        start = row
        metrics = total_metrics if label == '전사 계' else per_dept[DEPT_ORDER[DEPTS_SHORT.index(label)]]
        agg = agg_over(metrics, range(1, max_tier + 1))
        ws.cell(row=row, column=col_tier, value='합계')
        write_metric_values(ws, row, metric0, agg)
        for tier in range(1, max_tier + 1):
            r = row + tier
            ws.cell(row=r, column=col_tier, value=tier)
            write_metric_values(ws, r, metric0, metrics[tier])
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

    def write_block(label, total_agg, dept_aggs):
        nonlocal row
        start = row
        ws.cell(row=row, column=col_g, value=label)
        ws.merge_cells(start_row=row, start_column=col_g, end_row=row, end_column=col_g + 1)
        write_metric_values(ws, row, metric0, total_agg)
        _metric_border(ws, row, metric0)
        _pct_tint(ws, row, metric0)
        row += 1
        for d, short in zip(DEPT_ORDER, DEPTS_SHORT):
            ws.cell(row=row, column=col_g + 1, value=short)
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
            label_col = col_gw if label == '전사 계' else col_g
            metrics = total_metrics if label == '전사 계' else per_dept[DEPT_ORDER[DEPTS_SHORT.index(label)]]
            agg = _agg_all_tiers(metrics)
            ws.cell(row=row, column=col_tier, value='합계')
            write_metric_values(ws, row, metric0, agg)
            for tier in range(1, 6):
                r = row + tier
                ws.cell(row=r, column=col_tier, value=tier)
                write_metric_values(ws, r, metric0, metrics[tier])
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
            label_col = col_gw if label == '전사 계' else col_g
            metrics = total_metrics if label == '전사 계' else per_dept[DEPT_ORDER[DEPTS_SHORT.index(label)]]
            agg = _agg_all_tiers(metrics)
            ws.cell(row=row, column=label_col, value=label)
            write_metric_values(ws, row, metric0, agg)
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
# 6) 쳬결기간별_부문별
# ---------------------------------------------------------------------------

def _period_metrics_for(period_acc, dept, period_key, tier):
    if dept == '전사 계':
        parts = [period_acc.get(d, period_key, tier) for d in DEPT_ORDER]
        target = sum(p['target'] for p in parts)
        accident = sum(p['accident'] for p in parts)
        done = sum(p['done'] for p in parts)
        conv_target = target - accident
        not_done = conv_target - done
        return {'target': target, 'accident': accident, 'conv_target': conv_target,
                'done': done, 'not_done': not_done}
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
    target = sum(p['target'] for p in parts)
    accident = sum(p['accident'] for p in parts)
    done = sum(p['done'] for p in parts)
    conv_target = target - accident
    not_done = conv_target - done
    return {'target': target, 'accident': accident, 'conv_target': conv_target,
            'done': done, 'not_done': not_done}


def build_period_dept(wb, period_acc):
    """PERIODS 순서(오래된 구간부터)로 실제 도달 연차만큼만 행을 만든다."""
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
    for label, lookup_key in [('전사 계', '전사 계')] + list(zip(DEPTS_SHORT, DEPT_ORDER)):
        g_start = row
        label_col = col_gw if label == '전사 계' else col_g
        for plabel, key, _start, _end in PERIODS:
            reach_tiers = _period_reach(period_acc, lookup_key, key)
            if not reach_tiers:
                continue
            p_start = row
            ws.cell(row=row, column=col_period, value=plabel)
            ws.cell(row=row, column=col_key, value='전체')
            overall = _period_overall(period_acc, lookup_key, key, reach_tiers)
            write_period_values(ws, row, metric0, overall)
            _metric_border(ws, row, metric0)
            ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
            if len(reach_tiers) > 1:
                row += 1
                for t in reach_tiers:
                    ws.cell(row=row, column=col_key, value=key)
                    ws.cell(row=row, column=col_tier, value=t)
                    write_period_values(ws, row, metric0, _period_metrics_for(period_acc, lookup_key, key, t))
                    _metric_border(ws, row, metric0)
                    ws.cell(row=row, column=col_gw).fill = TOTAL_FILL
                    row += 1
                ws.merge_cells(start_row=p_start + 1, start_column=col_key,
                                end_row=row - 1, end_column=col_key)
            else:
                ws.cell(row=row, column=col_tier, value=reach_tiers[0])
                row += 1
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


def build(data_paths, mapping_path, output_path):
    print('[1/2] 데이터 계산 중(파이썬 직접 계산, LibreOffice 미사용)...')
    result = calc.compute_all(data_paths, mapping_path)

    print('[2/2] 보고서 시트 작성 중...')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_month_dept_summary(wb, result['months'], result['month_dept'], result['all_dept'])
    build_tier_dept_summary(wb, result['all_dept'])
    build_month_dept_tier(wb, result['months'], result['month_dept'], result['all_dept'])
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
