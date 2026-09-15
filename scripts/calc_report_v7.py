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


def make_periods(pairs):
    """(시작YYYYMM, 종료YYYYMM) 쌍 목록 -> PERIODS와 같은 형식의
    (라벨, 구간번호, 시작, 종료) 목록. GUI에서 사용자가 입력한 구간을
    계산에 쓰기 좋은 형태로 변환할 때 쓴다."""
    periods = []
    for i, (start, end) in enumerate(pairs):
        label = f'{start[:4]}.{start[4:]}~{end[:4]}.{end[4:]}'
        periods.append((label, i + 1, start, end))
    return periods


def period_of(ym, periods=None):
    if periods is None:
        periods = PERIODS
    for label, key, start, end in periods:
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
        # 상품별 시트는 요율일련번호(같은 상품코드 안에서 요율만 다른 변형)를
        # 구분하지 않고 상품코드 단위로 합산한다 - 상품명은 상품코드+요율일련번호를
        # 이어붙인 값이라(예: 코드 31038 + 일련번호 3 -> 상품명 310383) 상품코드
        # 자체가 이미 그 변형들의 공통 표시값이 된다.
        self.product = idx['상품코드']
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
        'plan': 0, 'codes': defaultdict(int),
    }


def cascade_counts(tier, v_this, v_prev, as_flag, yc, code):
    """한 행이 특정 연차(tier)에 대해 기여하는 사고유/전환예정/코드별 판정을
    한 곳에서 계산한다(표준 시트와 체결기간별 양쪽에서 재사용).
    반환: (사고유해당여부, matured제외여부, 전환예정/코드해당여부, 어떤코드인지)"""
    accident_hit = False
    accident_matured_excl = False
    if v_this == '0':
        accident_hit = True
        accident_matured_excl = as_flag and not yc
    elif tier > 1 and v_this == ' ' and v_prev == '0':
        accident_hit = True
        accident_matured_excl = as_flag and yc

    matured_excl_target = tier > 1 and v_this in ('1', ' ') and as_flag and yc

    # 코드별(연락두절/사고있음/고객거부/압류계약/ARS거부): 실제 현장에서 코드가
    # 기록됐다는 것 자체가 활동 확인 증거이므로, v_this가 0이 아니면(=1 또는 공란)
    # 카운트. 단 2연차부터는 공란이면서 직전 연차가 사고(0)였던 캐스케이드는 이미
    # 사고유로 처리됐으니 제외.
    # 전환예정(코드 없이 그냥 대기): 아직 어떤 활동 기록도 없는 공란 상태는
    # "대기 중"으로 볼 근거가 없으므로, v_this=1(정상 기록 존재)인 경우만 카운트.
    code_hit = False
    if v_this in ('1', ' ') and not as_flag:
        excluded = tier > 1 and v_this == ' ' and v_prev == '0'
        code_hit = not excluded
    plan_hit = v_this == '1' and not as_flag

    has_code = code in CODE_STRS
    plan_or_code_hit = code_hit if has_code else plan_hit

    return accident_hit, accident_matured_excl, matured_excl_target, plan_or_code_hit


def _apply_row(b, tier, v_this, v_prev, as_flag, yc, code):
    """bucket(dict) 하나에 원시 행 하나를 누적한다. tier는 '이 판정에 쓸 연차'
    (실제 경과년수일 수도, 과거 특정 연도 기준으로 역산한 연차일 수도 있음)."""
    b['raw'] += 1
    accident_hit, accident_matured_excl, matured_excl_target, plan_or_code_hit = \
        cascade_counts(tier, v_this, v_prev, as_flag, yc, code)
    if matured_excl_target:
        b['matured_excl'] += 1
    if v_this == '0':
        b['acc_t1'] += 1
        if accident_matured_excl:
            b['acc_t1_sub'] += 1
    elif tier > 1 and v_this == ' ' and v_prev == '0':
        b['acc_t3'] += 1
        if accident_matured_excl:
            b['acc_t3_sub'] += 1
    if as_flag and not yc:
        b['done'] += 1
    if plan_or_code_hit:
        if code in CODE_STRS:
            b['codes'][code] += 1
        else:
            b['plan'] += 1


def finalize_bucket(b):
    target = b['raw'] - b['matured_excl']
    accident = (b['acc_t1'] - b['acc_t1_sub']) + (b['acc_t3'] - b['acc_t3_sub'])
    conv_target = target - accident
    done = b['done']
    not_done = conv_target - done
    plan = b['plan']
    codes = {cs: b['codes'].get(cs, 0) for cs in CODE_STRS}
    idle = not_done - (plan + sum(codes.values()))
    return {
        'target': target, 'accident': accident, 'conv_target': conv_target,
        'done': done, 'not_done': not_done, 'plan': plan, 'codes': codes, 'idle': idle,
    }


class TierAccum:
    """key(예: (month,dept) 또는 dept 또는 (product,dept)) x tier(1~5) 별 원시 집계."""

    def __init__(self):
        self.data = defaultdict(lambda: defaultdict(_new_bucket))

    def add(self, key, tier, v_this, v_prev, as_flag, yc, code):
        _apply_row(self.data[key][tier], tier, v_this, v_prev, as_flag, yc, code)

    def finalize_bucket(self, b):
        return finalize_bucket(b)

    def finalize(self, key):
        out = {}
        for tier in TIERS:
            b = self.data.get(key, {}).get(tier)
            out[tier] = finalize_bucket(b) if b else finalize_bucket(_new_bucket())
        return out

    def keys(self):
        return list(self.data.keys())


class YearMonthAccum:
    """(target_year, month, dept) 키별 원시 집계 - "월별_부문별_new" 전용.

    TierAccum과 달리 tier(경과년수)를 그대로 쓰지 않고, 계약별로 "그 target_year
    시점엔 실제로 몇 연차였는지"를 역산한 연차(e)를 기준으로 v값/목표플랜을 다시
    판정해서 누적한다. 그래서 예를 들어 지금(오늘 기준) 4년차인 계약도, 2025년
    기준으로는 그 계약이 실제로 3년차였던 시점의 v값/목표플랜으로 판정되어
    2025년 버킷에 들어간다(현재 4년차 값을 그대로 재사용하지 않음)."""

    def __init__(self):
        self.data = defaultdict(_new_bucket)

    def add(self, key, e, v_this, v_prev, as_flag, yc, code):
        _apply_row(self.data[key], e, v_this, v_prev, as_flag, yc, code)

    def finalize(self, key):
        b = self.data.get(key)
        return finalize_bucket(b) if b else finalize_bucket(_new_bucket())

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

def _new_period_bucket():
    return {'target': 0, 'accident': 0, 'done': 0, 'max_done': 0, 'plan': 0, 'codes': defaultdict(int),
            'not_done_origin': defaultdict(int),
            'final_accident': 0, 'final_done': 0, 'final_not_done': 0,
            'final_plan': 0, 'final_codes': defaultdict(int)}


class PeriodAccum:
    """최대전환년수는 '몇 년 뒤에 추적을 끊는다'는 컷오프가 아니라 '총 몇 단계의
    전환목표가 있는가'일 뿐이다. 한 계약의 추적을 실제로 끊는 사유는 (1) 사고
    발생, (2) 최대전환년수 이상 도달한 연차에서 그 연차 목표플랜과 현재플랜이
    일치(진짜 매큐어+완료) 두 가지뿐이다. 미전환건은 최대전환년수를 넘겨도 실제
    경과연수가 허용하는 한 계속 대상에 남는다."""

    def __init__(self):
        self.data = defaultdict(lambda: defaultdict(lambda: defaultdict(_new_period_bucket)))
        self.reach = defaultdict(lambda: defaultdict(set))  # dept -> period_key -> {도달한 연차...}

    def add(self, dept, period_key, elapsed, seq_row_plan, current_plan, maxinfo, vlist, code):
        max_year, _final_target = maxinfo if maxinfo else (None, None)
        max_reach = min(elapsed, 5)
        if max_reach < 1:
            return
        acc_tier = None
        for i in range(5):
            if vlist[i] == '0':
                acc_tier = i + 1
                break
        k = seq_row_plan
        for n in range(1, max_reach + 1):
            if acc_tier is not None and acc_tier < n:
                break  # 이전 연차에 사고 -> 이후 연차 대상에서 완전히 빠짐
            bucket = self.data[dept][period_key][n]
            bucket['target'] += 1
            self.reach[dept][period_key].add(n)
            if acc_tier == n:
                bucket['accident'] += 1
                # 사고는 계약당 한 번만(정확히 그 연차에서) 잡히고 이 조건 자체가
                # 이 계약의 마지막 기여 지점이 되므로(다음 연차부터는 위 break로
                # 완전히 빠짐) 별도 처리 없이 그대로 '최종' 사고 건수로 쓸 수 있다.
                bucket['final_accident'] += 1
                continue
            as_flag_n = k >= n
            matured_n = max_year is not None and max_year <= n
            is_final_n = n == max_reach
            if as_flag_n:
                bucket['done'] += 1
                if matured_n:
                    bucket['max_done'] += 1
                    bucket['final_done'] += 1
                    break  # 최대전환년수 이상 도달 + 완료 = 더 이상 추적할 목표가 없음
                if is_final_n:
                    bucket['final_done'] += 1
                continue
            # 이 계약이 처음 뒤처지기 시작한 연차(기원 연차) = k+1. k는 n과 무관하게
            # 고정값이라 한 번 뒤처지면(k<n) 이후 모든 연차에서도 계속 뒤처진 것으로
            # 잡히므로, 미완료로 잡히는 매 연차마다 기원 연차는 항상 k+1로 동일하다.
            bucket['not_done_origin'][k + 1] += 1
            v_this = vlist[n - 1]
            v_prev = vlist[n - 2] if n > 1 else None
            _, _, _, plan_or_code_hit = cascade_counts(n, v_this, v_prev, as_flag_n, False, code)
            if plan_or_code_hit:
                if code in CODE_STRS:
                    bucket['codes'][code] += 1
                else:
                    bucket['plan'] += 1
            # '최종(final_*)' 계열은 이 계약이 현재 데이터 기준으로 도달한 마지막
            # 연차(max_reach)에서 딱 한 번만 집계한다 - target/done/not_done은
            # 위에서 보듯 살아있는 동안 매 연차 다시 잡히는 누적 구조라 그대로
            # 합산하면 중복 계산되지만, final_*은 계약당 정확히 한 번만 잡혀서
            # 여러 연차(tier)에 걸쳐 합산해도 실제 계약 수와 일치한다.
            if is_final_n:
                bucket['final_not_done'] += 1
                if plan_or_code_hit:
                    if code in CODE_STRS:
                        bucket['final_codes'][code] += 1
                    else:
                        bucket['final_plan'] += 1

    def rows_for(self, dept, period_key):
        """해당 부문/구간에서 실제로 존재하는 연차 목록(오름차순)."""
        return sorted(self.reach[dept][period_key])

    def get(self, dept, period_key, tier):
        b = self.data[dept][period_key][tier]
        target, accident, done = b['target'], b['accident'], b['done']
        conv_target = target - accident
        not_done = conv_target - done
        max_done = b['max_done']
        plan = b['plan']
        codes = {cs: b['codes'].get(cs, 0) for cs in CODE_STRS}
        idle = not_done - (plan + sum(codes.values()))
        not_done_origin = {j: b['not_done_origin'].get(j, 0) for j in range(1, tier + 1)}
        final_codes = {cs: b['final_codes'].get(cs, 0) for cs in CODE_STRS}
        return {'target': target, 'accident': accident, 'conv_target': conv_target,
                'done': done, 'max_done': max_done, 'not_done': not_done,
                'plan': plan, 'codes': codes, 'idle': idle, 'not_done_origin': not_done_origin,
                'final_accident': b['final_accident'], 'final_done': b['final_done'],
                'final_not_done': b['final_not_done'], 'final_plan': b['final_plan'],
                'final_codes': final_codes}


# ---------------------------------------------------------------------------
# 메인 계산 루프: 파일들을 한 번 순회하며 필요한 모든 누적기를 채운다.
# ---------------------------------------------------------------------------

def compute_all(data_paths, mapping_path, progress_every=100000,
                 new_sheet_years=None, reference_year=None, current_month=None, periods=None):
    """new_sheet_years/reference_year/current_month: "월별_부문별_new" 시트 전용
    (year_month_new 결과). reference_year/current_month는 이 데이터를 만든
    "오늘"에 해당하는 연/월 - 계약은 자기 계약월에 매년 생일이 와야 다음 연차로
    넘어가므로, 그 달이 current_month 이전/이후인지에 따라 지금 기록된 경과년수가
    실제로 몇 년도에 갱신된 것인지가 달라진다(아래 vintage_year 계산 참고). 셋 다
    생략하면 실행 시점의 실제 오늘 날짜로 자동 계산된다(수동으로 매번 갱신할
    필요 없음). periods: 체결기간별_부문별의 "구간" 목록(생략하면 모듈 상수
    PERIODS 사용) - make_periods()로 만든 형식."""
    if reference_year is None or current_month is None:
        today = date.today()
        if reference_year is None:
            reference_year = today.year
        if current_month is None:
            current_month = today.month
    if periods is None:
        periods = PERIODS
    if new_sheet_years is None:
        new_sheet_years = (reference_year - 1, reference_year)

    lookup, maxmap, seq = load_mapping(mapping_path)
    cols = get_cols(data_paths[0])

    month_dept = TierAccum()      # key=(month,dept)
    all_dept = TierAccum()        # key=dept (전체 기간 합산)
    product_dept = TierAccum()    # key=(product,dept)
    period_acc = PeriodAccum()
    year_month_new = YearMonthAccum()  # key=(target_year,month,dept) - 월별_부문별_new 전용
    year_month_tier_new = YearMonthAccum()  # key=(target_year,month,dept,e) - 월별_부문별_new_상세 전용

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

        # 월별_부문별_new: 계약체결월 기준 생일이 현재월 이전/이후인지에 따라,
        # 지금 기록된 tier(경과년수)가 실제로 갱신된 연도(vintage + tier)가
        # reference_year(생일이 이미 지났으면) 또는 그 전해(아직 안 지났으면)이다.
        # 여기서 역산한 vintage_year를 기준으로, 목표연도(Y)마다 "그때는 몇
        # 연차였는지(e)"를 다시 구해 그 연차 기준 v값/목표플랜으로 재판정한다.
        m_int = int(month) if str(month).isdigit() else None
        if m_int is not None:
            offset = 1 if m_int > current_month else 0
            vintage_year = reference_year - tier - offset
            # 과거 연도(e < tier)의 "완료" 판정에는 오늘의 current_plan을 그 해의
            # 목표플랜과 직접 비교할 수 없다(현재 정상적으로 매년 갱신 중인 계약은
            # 이미 그 해의 목표플랜을 지나쳐서 오늘 플랜과 더 이상 일치하지 않기
            # 때문). 대신 seq상 current_plan이 몇 년차 목표까지 도달했는지(k_reached)
            # 를 구해서, "그 해(e) 목표 이상을 이미 달성했는가"로 판정한다 - 정확히
            # e년차에 달성했는지는 알 수 없지만(그 사이 언젠가는 달성한 것은 확실),
            # 현재 플랜은 한번 올라가면 내려가지 않으므로 최소한 그 단계는 이미
            # 지났다는 것만은 보장된다.
            k_reached = completed_tier(seq, initial_plan, current_plan)
            for target_year in new_sheet_years:
                if target_year > reference_year:
                    continue  # 아직 오지 않은 미래 연도 - write_report_v7 쪽에서 빈칸으로 표시
                e = target_year - vintage_year
                if e < 1:
                    continue  # 그 계약은 target_year 시점엔 아직 존재하지 않았음
                e = min(e, tier)  # target_year==reference_year인데 아직 올해 생일 전인 달은 현재값으로 대체
                v_this_e = vlist[e - 1]
                v_prev_e = vlist[e - 2] if e > 1 else None
                if e == tier:
                    target_plan_e = elapsed_target(lookup, maxmap, initial_plan, e)
                    as_flag_e = (target_plan_e is not None and current_plan == target_plan_e)
                else:
                    as_flag_e = k_reached >= e
                yc_e = (mm is not None and mm[0] < e)
                year_month_new.add((target_year, month, dept), e, v_this_e, v_prev_e, as_flag_e, yc_e, code)
                year_month_tier_new.add((target_year, month, dept, e), e, v_this_e, v_prev_e, as_flag_e, yc_e, code)

        if cols.start is not None:
            start = row[cols.start]
            if start:
                ym = str(start)[:6]
                _plabel, pkey = period_of(ym, periods)
                if pkey is not None:
                    k = completed_tier(seq, initial_plan, current_plan)
                    period_acc.add(dept, pkey, tier, k, current_plan, mm, vlist, code)

    print(f'총 처리 행수: {n:,}')
    return {
        'month_dept': month_dept, 'all_dept': all_dept, 'product_dept': product_dept,
        'period_acc': period_acc, 'year_month_new': year_month_new,
        'year_month_tier_new': year_month_tier_new,
        'new_sheet_years': list(new_sheet_years), 'reference_year': reference_year,
        'current_month': current_month, 'periods': periods,
        'months': sorted(months_seen, key=lambda m: int(m)),
        'products': sorted(products_seen),
    }
