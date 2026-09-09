"""
calc_report_v7.py — v7 6시트 보고서를 LibreOffice/Excel 수식 없이 파이썬이 직접
계산해서 값만 채워 만든다(추후 exe 패키징을 위한 순수 계산 엔진의 1차 버전).

입력: 원본 데이터 시트(여러 파일에 걸쳐 있을 수 있음, 각각 Sheet1 + 무사고전환매핑테이블)
      - Sheet1 필수 컬럼: 계약체결월, 수금부문명, 상품명, 경과년수, 판매플랜코드(2개,
        앞=현재/뒤=최초), 보험기간시작일자, 전환사고구분값1~5, 전환처리구분코드
      - 무사고전환매핑테이블 필수 컬럼: *전환전 대표플랜코드, *무사고기간, 전환후 대표플랜코드

핵심 로직(v6 COUNTIFS 수식과 동일한 정의를 파이썬으로 직접 계산):
  대상   = 전체 - (2연차부터, 이미 matured(전환여부TRUE&년수비교TRUE)인 계약 제외)
  사고유 = 이번 연차 값=0(직접) + (이번 연차 공백 & 직전연차=0, 캐스케이드)
           - 각각 이미 matured로 확인된 건 제외
  전환대상 = 대상 - 사고유
  전환완료 = 전환여부(현재플랜==해당연차 목표플랜) TRUE & 년수비교 FALSE
  전환예정/코드별 = 이번연차=1(또는 캐스케이드 공백+직전=0 제외) & 전환여부FALSE & 코드일치
  미활동 = 미전환 - (전환예정+코드5종)

체결기간별_부문별은 완전히 다른 계산(한 계약이 여러 연차에 동시에 기여):
  - 사고발생연차: v[i]=='0'인 첫 tier(없으면 None)
  - 완료연차 k: 최초플랜코드의 연차별 목표플랜 시퀀스에서 현재플랜과 일치하는 가장 큰 연차
  - 최대전환년수 도달 시 그 다음 연차부터는 대상에서 제외
  - 사고발생연차 이전 연차까지만 대상에 포함(사고발생연차 자체는 포함, 그 다음부터 제외)
"""
import os
import sys
from collections import defaultdict
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl

DEPT_ORDER = ['개인사업부문', '전략사업부문', '신사업부문']
DEPT_SHORT = {'개인사업부문': '개인', '전략사업부문': '전략', '신사업부문': '신사업'}
CODE_MAP = [(1, '001'), (2, '002'), (3, '003'), (4, '004'), (7, '007')]
TIERS = [1, 2, 3, 4, 5]

PERIODS = [
    ('2022.07~2023.06', 1, '202207', '202306'),
    ('2023.07~2024.06', 2, '202307', '202406'),
    ('2024.07~2025.06', 3, '202407', '202506'),
    ('2025.07~2026.06', 4, '202507', '202606'),
]


def period_of(ym):
    for label, key, start, end in PERIODS:
        if start <= ym <= end:
            return label, key
    return None, None


def header_index(header):
    idx = {}
    for i, h in enumerate(header):
        if h and h not in idx:
            idx[h] = i
    return idx


def load_mapping(mapping_path):
    """무사고전환매핑테이블 -> (전환전대표플랜코드, 무사고기간)->전환후대표플랜코드,
    그리고 전환전대표플랜코드 -> (최대전환년수, 최대전환플랜코드)."""
    wb = openpyxl.load_workbook(mapping_path, read_only=True, data_only=True)
    ws = wb['무사고전환매핑테이블']
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = header_index(header)
    col_before = idx['*전환전 대표플랜코드']
    col_years = idx['*무사고기간']
    col_after = idx['전환후 대표플랜코드']

    seq = defaultdict(list)   # plan -> [(year, target), ...]
    lookup = {}                # (plan, year) -> target
    for row in rows:
        plan = row[col_before]
        yrs = row[col_years]
        tgt = row[col_after]
        if plan is None or yrs is None:
            continue
        seq[plan].append((yrs, tgt))
        lookup[(plan, yrs)] = tgt
    wb.close()

    maxmap = {}
    for plan, lst in seq.items():
        lst.sort(key=lambda x: x[0])
        final_target = lst[-1][1]
        max_year = min(y for y, t in lst if t == final_target)
        maxmap[plan] = (max_year, final_target)
    return lookup, maxmap, seq


def elapsed_target(lookup, maxmap, plan, year):
    """경과전환판매플랜코드: 매핑 조회 실패 시 최대전환매핑으로 폴백."""
    t = lookup.get((plan, year))
    if t is not None:
        return t
    mm = maxmap.get(plan)
    return mm[1] if mm else None


def completed_tier(seq, plan, current_plan):
    """현재플랜이 해당 최초플랜의 연차별 목표플랜 시퀀스에서 몇 년차까지
    도달했는지(k). 못 찾으면 0."""
    lst = seq.get(plan)
    if not lst:
        return 0
    k = 0
    for y, t in lst:
        if t == current_plan and y > k:
            k = y
    return k


def val(row, col):
    if col is None:
        return ' '
    v = row[col]
    return ' ' if v is None else str(v)


class Cols:
    def __init__(self, header):
        idx = header_index(header)
        self.month = idx['계약체결월']
        self.dept = idx['수금부문명']
        self.product = idx['상품명']
        self.tier = idx['경과년수']
        self.start = idx.get('보험기간시작일자')
        self.v = [idx[f'전환사고구분값{i}'] for i in range(1, 6)]
        self.code = idx['전환처리구분코드']
        plan_positions = [i for i, h in enumerate(header) if h == '판매플랜코드']
        self.current = plan_positions[0]
        self.initial = plan_positions[1]


def get_cols(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Sheet1']
    header = next(ws.iter_rows(max_row=1, values_only=True))
    wb.close()
    return Cols(header)


def iter_all_rows(data_paths, cols):
    """여러 데이터 파일의 Sheet1을 이어서 순회한다."""
    for path in data_paths:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb['Sheet1']
        rows = ws.iter_rows(values_only=True)
        next(rows)
        for row in rows:
            if row[cols.month] is None:
                continue
            yield row
        wb.close()


# ---------------------------------------------------------------------------
# 표준 연차별 지표 누적기(월별_부문별_연차별 / 연차별_부문별_상세 / 상품별_부문별_연차별)
# ---------------------------------------------------------------------------

CODE_STRS = [cs for _, cs in CODE_MAP]


def _new_bucket():
    return {
        'raw': 0, 'matured_excl': 0,
        'acc_t1': 0, 'acc_t1_sub': 0, 'acc_t3': 0, 'acc_t3_sub': 0,
        'done': 0,
        'plan_raw': 0, 'plan_sub': defaultdict(int),
        'code_direct': defaultdict(int), 'code_c2': defaultdict(int), 'code_c3': defaultdict(int),
    }


class TierAccum:
    """key(예: (month,dept) 또는 dept 또는 (product,dept)) x tier(1~5) 별 원시 집계."""

    def __init__(self):
        self.data = defaultdict(lambda: defaultdict(_new_bucket))

    def add(self, key, tier, v_this, v_prev, as_flag, yc, code):
        b = self.data[key][tier]
        b['raw'] += 1
        if tier > 1:
            if v_this in ('1', ' ') and as_flag and yc:
                b['matured_excl'] += 1
        if v_this == '0':
            b['acc_t1'] += 1
            if as_flag and not yc:
                b['acc_t1_sub'] += 1
        if tier > 1 and v_this == ' ' and v_prev == '0':
            b['acc_t3'] += 1
            if as_flag and yc:
                b['acc_t3_sub'] += 1
        if as_flag and not yc:
            b['done'] += 1
        if v_this == '1' and not as_flag:
            b['plan_raw'] += 1
            if code in CODE_STRS:
                b['plan_sub'][code] += 1
            b['code_direct'][code] += 1
        elif tier > 1 and v_this == ' ' and not as_flag:
            b['code_c2'][code] += 1
            if v_prev == '0':
                b['code_c3'][code] += 1

    def finalize_bucket(self, b):
        target = b['raw'] - b['matured_excl']
        accident = (b['acc_t1'] - b['acc_t1_sub']) + (b['acc_t3'] - b['acc_t3_sub'])
        conv_target = target - accident
        done = b['done']
        not_done = conv_target - done
        plan = b['plan_raw'] - sum(b['plan_sub'].get(cs, 0) for cs in CODE_STRS)
        codes = {}
        for cs in CODE_STRS:
            direct = b['code_direct'].get(cs, 0)
            c2 = b['code_c2'].get(cs, 0)
            c3 = b['code_c3'].get(cs, 0)
            codes[cs] = direct + c2 - c3
        idle = not_done - (plan + sum(codes.values()))
        return {
            'target': target, 'accident': accident, 'conv_target': conv_target,
            'done': done, 'not_done': not_done, 'plan': plan, 'codes': codes, 'idle': idle,
        }

    def finalize(self, key):
        out = {}
        for tier in TIERS:
            b = self.data.get(key, {}).get(tier)
            out[tier] = self.finalize_bucket(b) if b else self.finalize_bucket(_new_bucket())
        return out

    def keys(self):
        return list(self.data.keys())


def sum_metrics(metric_list):
    """여러 그룹(예: 개인/전략/신사업)의 finalize() 결과를 더해 '전사 계' 값을 만든다.
    (v6와 동일하게 target/accident/done/plan/code1~5를 더한 뒤 나머지는 재계산)"""
    out = {}
    for tier in TIERS:
        target = sum(m[tier]['target'] for m in metric_list)
        accident = sum(m[tier]['accident'] for m in metric_list)
        done = sum(m[tier]['done'] for m in metric_list)
        plan = sum(m[tier]['plan'] for m in metric_list)
        codes = {cs: sum(m[tier]['codes'][cs] for m in metric_list) for cs in CODE_STRS}
        conv_target = target - accident
        not_done = conv_target - done
        idle = not_done - (plan + sum(codes.values()))
        out[tier] = {
            'target': target, 'accident': accident, 'conv_target': conv_target,
            'done': done, 'not_done': not_done, 'plan': plan, 'codes': codes, 'idle': idle,
        }
    return out


# ---------------------------------------------------------------------------
# 체결기간별_부문별 전용 계산 (사고 캐스케이드 + 플랜매칭 완료판정, 한 계약이
# 여러 연차에 동시 기여)
# ---------------------------------------------------------------------------

class PeriodAccum:
    def __init__(self):
        # data[dept][period_key][tier] = [target, accident, done]  (전환대상=target-accident)
        self.data = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: [0, 0, 0])))
        self.reach = defaultdict(lambda: defaultdict(set))  # dept -> period_key -> {도달한 연차...}

    def add(self, dept, period_key, elapsed, seq_row_plan, current_plan, maxinfo, vlist):
        max_year, _final_target = maxinfo if maxinfo else (None, None)
        max_reach = elapsed if max_year is None else min(elapsed, max_year)
        if max_reach < 1:
            return
        acc_tier = None
        for i in range(5):
            if vlist[i] == '0':
                acc_tier = i + 1
                break
        for n in range(1, max_reach + 1):
            if acc_tier is not None and acc_tier < n:
                break  # 이전 연차에 사고 -> 이후 연차 대상에서 완전히 빠짐
            bucket = self.data[dept][period_key][n]
            bucket[0] += 1  # target(=대상계약A)
            self.reach[dept][period_key].add(n)
            if acc_tier == n:
                bucket[1] += 1  # accident(사고有B)
            else:
                k = seq_row_plan
                if k >= n:
                    bucket[2] += 1  # done(전환완료D)

    def rows_for(self, dept, period_key):
        """해당 부문/구간에서 실제로 존재하는 연차 목록(오름차순)."""
        return sorted(self.reach[dept][period_key])

    def get(self, dept, period_key, tier):
        target, accident, done = self.data[dept][period_key][tier]
        conv_target = target - accident
        not_done = conv_target - done
        return {'target': target, 'accident': accident, 'conv_target': conv_target,
                'done': done, 'not_done': not_done}


# ---------------------------------------------------------------------------
# 메인 계산 루프: 파일들을 한 번 순회하며 필요한 모든 누적기를 채운다.
# ---------------------------------------------------------------------------

def compute_all(data_paths, mapping_path, progress_every=100000):
    lookup, maxmap, seq = load_mapping(mapping_path)
    cols = get_cols(data_paths[0])

    month_dept = TierAccum()      # key=(month,dept)
    all_dept = TierAccum()        # key=dept (전체 기간 합산)
    product_dept = TierAccum()    # key=(product,dept)
    period_acc = PeriodAccum()

    months_seen = set()
    products_seen = set()
    dept_set = set(DEPT_ORDER)

    n = 0
    for row in iter_all_rows(data_paths, cols):
        dept = row[cols.dept]
        if dept not in dept_set:
            continue
        n += 1
        if progress_every and n % progress_every == 0:
            print(f'  ...{n:,}행 처리', flush=True)

        month = row[cols.month]
        product = row[cols.product]
        tier = row[cols.tier]
        if not isinstance(tier, int) or tier < 1 or tier > 5:
            continue
        current_plan = row[cols.current]
        initial_plan = row[cols.initial]
        code = val(row, cols.code)
        vlist = [val(row, c) for c in cols.v]
        v_this = vlist[tier - 1]
        v_prev = vlist[tier - 2] if tier > 1 else None

        target_plan = elapsed_target(lookup, maxmap, initial_plan, tier)
        as_flag = (target_plan is not None and current_plan == target_plan)
        mm = maxmap.get(initial_plan)
        yc = (mm is not None and mm[0] < tier)

        months_seen.add(month)
        products_seen.add(product)

        month_dept.add((month, dept), tier, v_this, v_prev, as_flag, yc, code)
        all_dept.add(dept, tier, v_this, v_prev, as_flag, yc, code)
        product_dept.add((product, dept), tier, v_this, v_prev, as_flag, yc, code)

        if cols.start is not None:
            start = row[cols.start]
            if start:
                ym = str(start)[:6]
                _plabel, pkey = period_of(ym)
                if pkey is not None:
                    k = completed_tier(seq, initial_plan, current_plan)
                    period_acc.add(dept, pkey, tier, k, current_plan, mm, vlist)

    print(f'총 처리 행수: {n:,}')
    return {
        'month_dept': month_dept, 'all_dept': all_dept, 'product_dept': product_dept,
        'period_acc': period_acc,
        'months': sorted(months_seen, key=lambda m: int(m)),
        'products': sorted(products_seen),
    }
