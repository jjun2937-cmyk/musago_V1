"""
write_report_v7.py — calc_report_v7.py로 계산한 값을, build_report_v7.py가 만든
레이아웃(서식/병합/색상)에 실제로 채워 넣어 최종 보고서 파일을 만든다.
LibreOffice/Excel 수식은 전혀 쓰지 않는다(순수 값만 기록).
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill

NO_FILL = PatternFill(fill_type=None)
WHITE_FILL = PatternFill('solid', fgColor='FFFFFF')

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
# 1.5) 월별_연차별 (부문 구분 없이 전사계만, 월별 x 연차별)
# ---------------------------------------------------------------------------

def build_month_tier(wb, months, month_dept_acc):
    ws = wb.create_sheet('월별_연차별')
    ws.sheet_view.showGridLines = False
    col_month, col_tier = 2, 3
    metric0 = 4
    scell = ws.cell(row=1, column=metric0 + METRIC_WIDTH - 1, value=f'({date.today():%Y.%m.%d} 기준)')
    scell.font = layout.STAMP_FONT
    scell.alignment = Alignment(horizontal='right', vertical='bottom')

    r0 = 2
    ws.merge_cells(start_row=r0, start_column=col_month, end_row=r0 + 2, end_column=col_month)
    for rr in range(r0, r0 + 3):
        _hcell(ws, rr, col_month)
    _hcell(ws, r0, col_month, '월')
    ws.merge_cells(start_row=r0, start_column=col_tier, end_row=r0 + 2, end_column=col_tier)
    for rr in range(r0, r0 + 3):
        _hcell(ws, rr, col_tier)
    _hcell(ws, r0, col_tier, '연차')
    write_metric_header(ws, r0, metric0)

    row = r0 + 3
    tall_rows = [r0]
    for month in months:
        month_start = row
        total_metrics = calc.sum_metrics([month_dept_acc.finalize((month, d)) for d in DEPT_ORDER])
        agg = _agg_all_tiers(total_metrics)

        set_value_label(ws, row, col_tier, '합계')
        write_metric_values(ws, row, metric0, agg)
        layout._metric_fill(ws, row, metric0, TOTAL_FILL)
        layout._metric_border(ws, row, metric0)
        ws.cell(row=row, column=col_tier).fill = TOTAL_FILL
        for t in TIERS:
            r = row + t
            set_value_label(ws, r, col_tier, t)
            write_metric_values(ws, r, metric0, total_metrics[t])
            layout._metric_fill(ws, r, metric0, WHITE_FILL)
            layout._metric_border(ws, r, metric0)
            ws.cell(row=r, column=col_tier).fill = WHITE_FILL
        for r in range(row, row + 6):
            set_label_border(ws, r, (col_tier,))

        ws.cell(row=row, column=col_month).fill = TOTAL_FILL
        for r in range(row, row + 6):
            set_label_border(ws, r, (col_month,), top=(r == row), bottom=(r == row + 5))
        set_value_label(ws, row, col_month, int(month) if str(month).isdigit() else month)
        ws.merge_cells(start_row=row, start_column=col_month, end_row=row + 5, end_column=col_month)

        row = month_start + 6

    widths = {'B': 5.625, 'C': 5.25}
    for off, width in ((0, 8.375), (3, 6.625), (5, 6.625), (6, 10.375), (7, 6.625),
                       (8, 8.375), (9, 6.625), (15, 8.625), (16, 8.375), (17, 6.625)):
        widths[openpyxl.utils.get_column_letter(metric0 + off)] = width
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
# 2.5) 월별_부문별_new (달력연도 기준 시계열: calc.compute_all()이 계약별로
#      각 연도 시점 실제 연차를 역산해서 그 연차 기준으로 이미 재판정해 둔
#      result['year_month_new']를 부문/월 합산만 해서 채운다.)
# ---------------------------------------------------------------------------

def _agg_sum(parts):
    """이미 계산된(같은 형태의) 여러 집계 dict를 그대로 다시 더한다(연간 누계용).
    None(그 달엔 데이터 없음)은 건너뛴다."""
    parts = [p for p in parts if p is not None]
    if not parts:
        return None
    target = sum(p['target'] for p in parts)
    accident = sum(p['accident'] for p in parts)
    done = sum(p['done'] for p in parts)
    plan = sum(p['plan'] for p in parts)
    codes = {cs: sum(p['codes'][cs] for p in parts) for cs in CODE_STRS}
    conv_target = target - accident
    not_done = conv_target - done
    idle = not_done - (plan + sum(codes.values()))
    return {'target': target, 'accident': accident, 'conv_target': conv_target,
            'done': done, 'not_done': not_done, 'plan': plan, 'codes': codes, 'idle': idle}


def write_metric_values_transposed(ws, col, row0, m):
    """write_metric_values()와 같은 18개 지표를, 가로(열) 한 줄이 아니라 세로
    (row0부터 아래로 18행)로 채운다. m이 None이면(그 시점엔 데이터가 없음) 전부
    빈 칸으로 남긴다."""
    THIN = layout.THIN

    def setv(off, value, is_pct=False):
        cell = ws.cell(row=row0 + off, column=col)
        if value is not None:
            cell.value = value
        cell.number_format = RATIO_FMT if is_pct else COUNT_FMT
        cell.font = BODY_FONT
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

    if m is None:
        for off in range(18):
            setv(off, None, off in (3, 5, 7, 9, 17))
        return

    target, accident, conv_target = m['target'], m['accident'], m['conv_target']
    done, not_done, plan, idle = m['done'], m['not_done'], m['plan'], m['idle']
    codes = m['codes']
    act_sum = plan + sum(codes.values())

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


def build_month_dept_new(wb, months, year_month_acc, years, reference_year):
    """years: 왼쪽부터 표시할 연도 목록(예: [2025, 2026], 나중에 2024/2027 등을
    앞뒤에 추가하면 자동으로 확장됨). year_month_acc: calc.compute_all()의
    result['year_month_new'] - 계약별로 각 연도 시점 실제 연차를 역산해서 그
    연차 기준 v값/목표플랜으로 이미 다시 판정해 둔 (연도,월,부문) 집계이므로,
    여기서는 그대로 부문 합산/월 합산만 하면 된다. reference_year보다 미래인
    연도는 실제로 계산할 방법이 없으므로 그 연도 블록은 전부 빈칸으로 둔다."""
    ws = wb.create_sheet('월별_부문별_new')
    ws.sheet_view.showGridLines = False

    col_label = 2  # B (B:E 4칸)
    n_label_cols = 4
    row0 = 5       # 유지계약(A) 행
    n_metric_rows = 18
    cols_per_year = 16  # 1~12월 + 누계 + 부문(개인/전략/신사업)

    hdr_r0 = 2
    ws.merge_cells(start_row=hdr_r0, start_column=col_label, end_row=hdr_r0 + 2, end_column=col_label + n_label_cols - 1)
    for rr in range(hdr_r0, hdr_r0 + 3):
        for c in range(col_label, col_label + n_label_cols):
            _hcell(ws, rr, c)
    _hcell(ws, hdr_r0, col_label, '구분')

    col_B, col_C, col_D, col_E = col_label, col_label + 1, col_label + 2, col_label + 3

    label_rows = {
        row0 + 0: ('도래계약(유지중) A', col_B, col_E),
        row0 + 1: ('사고有 B', col_B, col_E),
        row0 + 2: ('전환대상 C (A-B)', col_B, col_E),
        row0 + 4: ('전환완료 D', col_B, col_E),
        row0 + 6: ('전환미완료 E (C-D)', col_B, col_E),
        row0 + 8: ('현장활동 확인', col_C, col_E),
        row0 + 16: ('현장활동 미확인', col_C, col_E),
    }
    pct_only_rows = {row0 + 3, row0 + 5, row0 + 7, row0 + 9, row0 + 17}
    leaf_labels = {
        row0 + 10: '전환예정', row0 + 11: '연락두절', row0 + 12: '사고있음',
        row0 + 13: '고객거부', row0 + 14: '압류계약', row0 + 15: 'ARS거부',
    }
    def _b(spec):
        THIN = layout.THIN
        s = lambda v: THIN if v else None
        return Border(top=s(spec[0]), bottom=s(spec[1]), left=s(spec[2]), right=s(spec[3]))

    # (B, C, D, E) 열 테두리 - 원본 템플릿에서 셀 단위로 그대로 옮긴 값(1=THIN, 0=없음).
    label_borders = {
        row0 + 0: ((1, 1, 1, 1), (1, 1, 0, 0), (1, 1, 0, 0), (1, 1, 0, 1)),
        row0 + 1: ((1, 1, 1, 1), (1, 1, 0, 0), (1, 1, 0, 0), (1, 1, 0, 1)),
        row0 + 2: ((1, 0, 1, 1), (1, 0, 0, 0), (1, 0, 0, 0), (1, 0, 0, 1)),
        row0 + 3: ((0, 1, 1, 0), (0, 1, 0, 0), (0, 1, 0, 0), (1, 1, 1, 1)),
        row0 + 4: ((1, 0, 1, 1), (1, 0, 0, 0), (1, 0, 0, 0), (1, 0, 0, 1)),
        row0 + 5: ((0, 1, 1, 0), (0, 1, 0, 0), (0, 1, 0, 0), (1, 1, 1, 1)),
        row0 + 6: ((0, 0, 1, 1), (0, 0, 0, 0), (0, 0, 0, 0), (0, 0, 0, 1)),
        row0 + 7: ((0, 0, 1, 0), (0, 0, 0, 0), (0, 0, 0, 0), (1, 0, 1, 1)),
        row0 + 8: ((0, 0, 1, 0), (1, 0, 1, 1), (1, 0, 0, 0), (1, 0, 0, 1)),
        row0 + 9: ((0, 0, 1, 0), (0, 0, 1, 0), (0, 0, 0, 0), (1, 1, 1, 1)),
        row0 + 10: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
        row0 + 11: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
        row0 + 12: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
        row0 + 13: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
        row0 + 14: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
        row0 + 15: ((0, 0, 1, 0), (0, 1, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
        row0 + 16: ((0, 0, 1, 0), (1, 0, 1, 1), (1, 0, 0, 0), (1, 0, 0, 1)),
        row0 + 17: ((0, 1, 1, 0), (0, 1, 1, 0), (0, 1, 0, 0), (1, 1, 1, 1)),
    }
    for r in range(row0, row0 + n_metric_rows):
        is_pct_row = r in pct_only_rows
        b_spec, c_spec, d_spec, e_spec = label_borders[r]
        e_fill = layout.PCT_HEADER_FILL if is_pct_row else NO_FILL
        _hcell(ws, r, col_B, fill=layout.HEADER_FILL, border=_b(b_spec))
        _hcell(ws, r, col_C, fill=layout.HEADER_FILL, border=_b(c_spec))
        _hcell(ws, r, col_D, fill=layout.HEADER_FILL, border=_b(d_spec))
        _hcell(ws, r, col_E, fill=e_fill, border=_b(e_spec))
    for r, (text, start_c, end_c) in label_rows.items():
        if start_c != end_c:
            ws.merge_cells(start_row=r, start_column=start_c, end_row=r, end_column=end_c)
        ws.cell(row=r, column=start_c, value=text)
    for r in pct_only_rows | {row0 + 17}:
        ws.cell(row=r, column=col_E, value='%')
    for r, text in leaf_labels.items():
        ws.merge_cells(start_row=r, start_column=col_D, end_row=r, end_column=col_E)
        ws.cell(row=r, column=col_D, value=text)

    col0 = col_label + n_label_cols  # F
    tall_rows = [hdr_r0]

    for yi, year in enumerate(years):
        is_future = year > reference_year
        year_col0 = col0 + yi * cols_per_year
        year_last_col = year_col0 + cols_per_year - 1
        ws.merge_cells(start_row=hdr_r0, start_column=year_col0, end_row=hdr_r0, end_column=year_last_col)
        for c in range(year_col0, year_col0 + cols_per_year):
            _hcell(ws, hdr_r0, c, border=Border(top=layout.THIN, bottom=layout.THIN, left=layout.THIN, right=layout.THIN))
        ws.cell(row=hdr_r0, column=year_col0, value=year)

        month_aggs = []
        for mi, month in enumerate(months):
            mcol = year_col0 + mi
            is_first_month = mi == 0
            is_last_month = mi == len(months) - 1
            right = None if is_last_month else layout.THIN
            left = None if is_first_month else layout.THIN
            agg = None if is_future else _agg_sum([year_month_acc.finalize((year, month, d)) for d in DEPT_ORDER])
            month_aggs.append(agg)
            ws.merge_cells(start_row=hdr_r0 + 1, start_column=mcol, end_row=hdr_r0 + 2, end_column=mcol)
            _hcell(ws, hdr_r0 + 1, mcol, f'{int(month) if str(month).isdigit() else month}월',
                   border=Border(top=layout.THIN, bottom=layout.THIN, left=left, right=right))
            _hcell(ws, hdr_r0 + 2, mcol,
                   border=Border(top=None, bottom=layout.THIN, left=left, right=right))
            write_metric_values_transposed(ws, mcol, row0, agg)

        total_col = year_col0 + len(months)
        ws.merge_cells(start_row=hdr_r0 + 1, start_column=total_col, end_row=hdr_r0 + 2, end_column=total_col)
        _hcell(ws, hdr_r0 + 1, total_col, '누계',
               border=Border(top=layout.THIN, bottom=layout.THIN, left=layout.THIN, right=None))
        _hcell(ws, hdr_r0 + 2, total_col,
               border=Border(top=None, bottom=layout.THIN, left=layout.THIN, right=None))
        write_metric_values_transposed(ws, total_col, row0, _agg_sum(month_aggs))

        for di, (d, short) in enumerate(zip(DEPT_ORDER, DEPTS_SHORT)):
            dcol = total_col + 1 + di
            is_last_dept = di == len(DEPT_ORDER) - 1
            THIN = layout.THIN
            row3_border = Border(top=THIN, bottom=None, left=None, right=(THIN if is_last_dept else None))
            row4_border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
            _hcell(ws, hdr_r0 + 1, dcol, border=row3_border)
            _hcell(ws, hdr_r0 + 2, dcol, short, border=row4_border)
            dept_month_aggs = None if is_future else [year_month_acc.finalize((year, month, d)) for month in months]
            write_metric_values_transposed(ws, dcol, row0, _agg_sum(dept_month_aggs) if dept_month_aggs else None)

    ws.column_dimensions[openpyxl.utils.get_column_letter(col_B)].width = 3.125
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_C)].width = 13.0
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_D)].width = 8.125
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_E)].width = 5.75
    for yi in range(len(years)):
        year_col0 = col0 + yi * cols_per_year
        for c in range(year_col0, year_col0 + cols_per_year):
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13.0
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    for r in (row0 + 0, row0 + 1, row0 + 2, row0 + 4, row0 + 6, row0 + 8, row0 + 16):
        ws.row_dimensions[r].height = 16.5
    return ws


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
    final_accident = sum(p.get('final_accident', 0) for p in parts)
    final_done = sum(p.get('final_done', 0) for p in parts)
    final_not_done = sum(p.get('final_not_done', 0) for p in parts)
    final_plan = sum(p.get('final_plan', 0) for p in parts)
    final_codes = {cs: sum(p.get('final_codes', {}).get(cs, 0) for p in parts) for cs in CODE_STRS}
    return {'target': target, 'accident': accident, 'conv_target': conv_target,
            'done': done, 'max_done': max_done, 'not_done': not_done,
            'plan': plan, 'codes': codes, 'idle': idle, 'not_done_origin': not_done_origin,
            'final_accident': final_accident, 'final_done': final_done, 'final_not_done': final_not_done,
            'final_plan': final_plan, 'final_codes': final_codes}


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
    """'전체' 행 전용: 연차마다 다시 잡히는 target/done/not_done을 그대로 더하면
    같은 계약이 여러 번 중복 집계된다. final_*(계약당 정확히 한 번, 그 계약이
    실제로 도달한 마지막 연차에서만 기록됨)를 합산해서 구간 전체의 '서로 다른
    계약 수' 기준으로 재구성한다. max_done/사고(accident)는 원래도 계약당 한
    번만 기록되므로 그대로 합산한다."""
    parts = [_period_metrics_for(period_acc, dept, period_key, t) for t in tiers]
    s = _sum_period_parts(parts)
    final_accident = s['final_accident']
    final_done = s['final_done']
    final_not_done = s['final_not_done']
    final_plan = s['final_plan']
    final_codes = s['final_codes']
    target = final_accident + final_done + final_not_done
    conv_target = final_done + final_not_done
    idle = final_not_done - (final_plan + sum(final_codes.values()))
    return {'target': target, 'accident': final_accident, 'conv_target': conv_target,
            'done': final_done, 'max_done': s['max_done'], 'not_done': final_not_done,
            'plan': final_plan, 'codes': final_codes, 'idle': idle, 'not_done_origin': {}}


def build_period_dept(wb, period_acc, periods=PERIODS):
    """periods 순서(오래된 구간부터)로 실제 도달 연차만큼만 행을 만든다."""
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
        for plabel, key, _start, _end in periods:
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


def build_period_dept_remaining(wb, period_acc, periods=PERIODS):
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
        for plabel, key, _start, _end in periods:
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
    term_row('월별_부문별_new', '연도(2025/2026)를 나란히 놓고 보는 달력 기준 시계열 시트. 아래 6번 참고.')
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
    blank()

    section('6. 월별_부문별_new 읽는 법')
    para('이 시트는 "월별_부문별_연차별"과 같은 원본 계약 데이터를, 달력 연도 기준 시계열로 다시 '
         '집계한 것입니다. 각 계약마다 "그 연도 그 달 시점엔 실제로 몇 연차였는지"를 역산해서, '
         '그 연차에 해당하는 목표플랜/판정 기준으로 다시 계산합니다 - 지금(오늘) 몇 연차인지를 '
         '그대로 재사용하지 않습니다.')
    para('예를 들어 오늘 기준 4년차인 계약이 있다면, 그 계약은 2025년 시점엔 아직 4년차가 아니라 '
         '3년차였습니다(1년씩 순서대로 올라가므로). 그래서 이 시트의 "2025년" 칸에서는 이 계약을 '
         '3년차 기준(그때의 더 쉬운 목표플랜, 그때의 사고/전환 판정)으로 다시 계산해서 반영합니다. '
         '아직 그 해에 존재하지도 않았던 계약(예: 올해 막 1년차가 된 계약의 2025년 이전 시점)은 '
         '해당 연도 집계에서 완전히 제외됩니다.')
    para('월별 값(1~12월, 각 달의 "누계"칸 포함)은 전사계(부문 구분 없는 합계) 기준입니다. '
         '오른쪽의 개인/전략/신사업 칸은 월별로 나누지 않고, 그 연도 1년 누계(부문별) 값만 보여줍니다.')
    para('주의: 계약은 자기 계약월에 매년 "생일"이 와야 다음 연차로 넘어갑니다. 오늘 기준으로 아직 '
         '생일이 안 지난 달(현재월보다 큰 달)은 "올해" 데이터 자체가 아직 없으므로, 그런 달은 작년 '
         '컬럼과 올해 컬럼이 같은 값을 보여줍니다(현재 시점 기준 가장 최신 값을 그대로 사용).')
    blank()

    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 22
    return ws


def build_reference_sheet(wb, months, month_dept_acc, current_month=None):
    """'참고' 시트: 완료율에 영향을 주는 주요 변수(1년차 진입 후 경과개월수)와
    완료율의 상관관계를 실무자가 참고할 수 있도록 정리. 매 실행마다 실제
    데이터로 다시 계산되며, 보고서의 다른 계산에는 영향을 주지 않는다.
    current_month을 생략하면 실행 시점의 실제 오늘 월을 쓴다."""
    if current_month is None:
        current_month = date.today().month
    ws = wb.create_sheet('참고')
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
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        c = ws.cell(row=r, column=2, value=text)
        c.font = TITLE_F
        c.alignment = Alignment(horizontal='left', vertical='center')
        row[0] += 2

    def section(text):
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        c = ws.cell(row=r, column=2, value=text)
        c.font = SEC_F
        c.fill = SEC_FILL
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 22
        row[0] += 1

    def para(text, height=45):
        r = row[0]
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        for col in range(2, 7):
            cc = ws.cell(row=r, column=col)
            cc.font = BODY_F
            cc.border = BORDER
        c = ws.cell(row=r, column=2, value=text)
        c.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        ws.row_dimensions[r].height = height
        row[0] += 1

    def blank(n=1):
        row[0] += n

    title('참고 - 완료율에 영향을 주는 주요 변수')
    para('이 시트는 "몇 연차가 된 지 얼마나 지났는지"(경과개월)가 전환완료율과 강한 상관관계를 '
         '보인다는 확인 내용을 정리한 참고용 시트입니다. 매 실행 시 실제 데이터로 다시 계산되며, '
         '보고서의 다른 계산에는 전혀 영향을 주지 않습니다.')
    blank()

    section('1년차 진입 후 경과개월 vs 완료율 (전사계, 1년차 기준)')
    para('계약체결월별로 "그 달 계약이 1년차가 된 지 몇 개월이 지났는지"를 계산했습니다. 계약은 '
         '자기 계약월에 매년 "생일"이 와야 다음 연차로 넘어가므로, 현재월(9월) 이전 달은 올해 이미 '
         '생일이 지나 1년차가 된 지 (9－계약월)개월, 현재월 이후 달은 작년에 생일이 지나 1년차가 '
         '된 지 (9＋12－계약월)개월이 지난 상태입니다.')

    r = row[0]
    headers = ['계약체결월', '1년차 진입 후\n경과개월', '1년차 대상', '1년차 완료', '완료율(%)']
    for i, h in enumerate(headers):
        c = ws.cell(row=r, column=2 + i, value=h)
        c.font = HEAD_F
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = 28
    row[0] += 1

    rows_data = []
    for month in months:
        mi = int(month)
        months_since = (current_month - mi) if mi <= current_month else (current_month + 12 - mi)
        metrics_by_tier = calc.sum_metrics([month_dept_acc.finalize((month, d)) for d in DEPT_ORDER])
        t1 = metrics_by_tier[1]
        target, done = t1['conv_target'], t1['done']
        rate = round(done / target * 100, 1) if target else None
        rows_data.append((mi, months_since, target, done, rate))

    for mi, months_since, target, done, rate in rows_data:
        r = row[0]
        for i, v in enumerate([f'{mi}월', months_since, target, done, rate]):
            c = ws.cell(row=r, column=2 + i, value=v)
            c.font = BODY_F
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center')
        row[0] += 1
    blank()

    def _corr(xs, ys):
        n = len(xs)
        mx, my = sum(xs) / n, sum(ys) / n
        cov = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / n
        sx = (sum((x - mx) ** 2 for x in xs) / n) ** 0.5
        sy = (sum((y - my) ** 2 for y in ys) / n) ** 0.5
        return cov / (sx * sy) if sx and sy else None

    valid = [(ms, tg, rt) for _, ms, tg, _, rt in rows_data if rt is not None]
    corr_time = _corr([v[0] for v in valid], [v[2] for v in valid])
    corr_vol = _corr([v[1] for v in valid], [v[2] for v in valid])

    section('상관계수 요약')
    para(f'경과개월수 vs 완료율 상관계수 = {corr_time:.2f}  /  1년차 대상 물량 vs 완료율 상관계수 = {corr_vol:.2f}',
         height=30)
    para('결론: 완료율은 "그 달 물량이 얼마나 많은지"보다 "1년차로 전환된 지 얼마나 시간이 지났는지"에 '
         '훨씬 크게 좌우됩니다. 담당자가 접촉/처리할 시간이 누적될수록 완료율이 올라가는 것으로 보입니다 '
         '(단, 특정 계약월 코호트 고유의 사정으로 예외가 있을 수 있습니다 - 예: 경과개월이 짧은 편임을 '
         '고려해도 완료율이 유독 낮게 나타나는 달이 있을 수 있음).')

    for col, width in zip('BCDEF', [12, 14, 12, 12, 12]):
        ws.column_dimensions[col].width = width
    return ws


def build(data_paths, mapping_path, output_path, reference_year=None, current_month=None,
          new_sheet_years=None, period_pairs=None):
    """reference_year/current_month: "오늘"에 해당하는 연/월(월별_부문별_new,
    참고 시트에서 씀). 둘 다 생략하면 실행 시점의 실제 오늘 날짜를 쓴다. 과거
    특정 시점 기준으로 다시 계산하고 싶을 때만 명시적으로 넘긴다.
    new_sheet_years: 월별_부문별_new에 나란히 보여줄 연도 목록(예: [2025, 2026]).
    생략하면 (reference_year-1, reference_year) 두 해를 자동으로 쓴다.
    period_pairs: 체결기간별_부문별의 "구간" 목록 - (시작YYYYMM, 종료YYYYMM)
    쌍의 리스트(예: [('202207','202306'), ('202307','202406')]). 생략하면
    calc_report_v7.PERIODS(기본 구간)를 그대로 쓴다."""
    periods = calc.make_periods(period_pairs) if period_pairs is not None else None
    print('[1/2] 데이터 계산 중(파이썬 직접 계산, LibreOffice 미사용)...')
    result = calc.compute_all(data_paths, mapping_path,
                               reference_year=reference_year, current_month=current_month,
                               new_sheet_years=new_sheet_years, periods=periods)

    print('[2/2] 보고서 시트 작성 중...')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_glossary_sheet(wb)
    build_month_dept_new(wb, result['months'], result['year_month_new'], years=result['new_sheet_years'],
                          reference_year=result['reference_year'])
    build_month_dept_summary(wb, result['months'], result['month_dept'], result['all_dept'])
    build_tier_dept_summary(wb, result['all_dept'])
    build_month_tier(wb, result['months'], result['month_dept'])
    build_month_dept_tier(wb, result['months'], result['month_dept'], result['all_dept'])
    build_period_dept_remaining(wb, result['period_acc'], periods=result['periods'])
    build_period_dept(wb, result['period_acc'], periods=result['periods'])
    build_product_dept_tier(wb, result['products'], result['product_dept'])
    build_product_dept_summary(wb, result['products'], result['product_dept'])
    build_reference_sheet(wb, result['months'], result['month_dept'], current_month=result['current_month'])
    wb.save(output_path)
    print('저장 완료 ->', output_path)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python write_report_v7.py <출력.xlsx> <데이터파일1> [데이터파일2 ...]')
        sys.exit(1)
    output_path = sys.argv[1]
    data_paths = sys.argv[2:]
    build(data_paths, data_paths[0], output_path)
