"""무사고전환 보고서 생성기 - 단일 파일판 (calc/build/write/gui 통합).
회사 보안정책 때문에 여러 파일을 그대로 반입하기 어려운 경우를 위해,
기존 scripts/calc_report_v7.py + build_report_v7.py + write_report_v7.py +
gui_app.py 네 파일을 이 파일 하나로 합쳐 놓았다. 동작은 완전히 동일하다."""

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
BODY_FONT = Font(name=FONT_NAME, size=11)
LABEL_FONT = Font(name=FONT_NAME, size=11, bold=True)

DEPTS_SHORT = ['개인', '전략', '신사업']


def set_group_label(ws, row, col, value, horizontal='center'):
    """전사계/부문(개인·전략·신사업) 등 그룹 라벨 셀: 굵게+가운데(수직) 정렬."""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = LABEL_FONT
    cell.alignment = Alignment(horizontal=horizontal, vertical='center')
    return cell


LABEL_BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def set_label_border(ws, row, cols, top=True, bottom=True):
    """부문별/연차/월/연도별/상품명 등 라벨 칸(지표 블록 바깥쪽)에도 데이터 행마다
    표 전체와 이어지는 테두리를 준다(그렇지 않으면 배경색만 있고 선이 없어 보인다).
    top/bottom=False로 주면 그룹 내부 행 사이의 가로줄을 없앤다."""
    border = Border(top=THIN if top else None, bottom=THIN if bottom else None, left=THIN, right=THIN)
    for c in cols:
        ws.cell(row=row, column=c).border = border


def set_value_label(ws, row, col, value):
    """월/연차/구간 등 값 성격의 라벨 셀: 일반체+가운데(수평/수직) 정렬."""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = BODY_FONT
    cell.alignment = Alignment(horizontal='center', vertical='center')
    return cell

# 지표 블록 리프 정의: (오프셋, 라벨, 병합폭(1이면 단일열), 종류)
# 종류: 'plain'(값만) / 'ratio'(값+바로 다음 열이 %) / 'pct'(비율 표시 전용, ratio가 채움) /
#       'section'(2행 병합 라벨, 리프 자식들이 이어짐)
METRIC_WIDTH = 18


def _hcell(ws, r, c, value=None, fill=HEADER_FILL, border=None):
    cell = ws.cell(row=r, column=c)
    if value is not None:
        cell.value = value
    cell.font = HEADER_FONT
    cell.fill = fill
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = border if border is not None else Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    return cell


# 헤더 3행 블록의 실측(원본 템플릿) 테두리 패턴 - 두 부류로 나뉜다:
#  Type A(전환대상/전환완료/전환미완료 각각에 곧바로 붙는 %열, 오프셋3·5·7):
#    위 2행은 완전히 비워두고(내부 선 없음) 라벨행 윗변에만 얇은 선.
#  Type B(현장활동확인%부터 그 뒤 전환예정·코드5종·현장활동미확인%까지, 오프셋9~17):
#    2번째 행부터 이미 얇은 윗선이 있어 A보다 한 줄 위에서 선이 시작된다.
_TYPE_A_TOP = Border(top=THIN, bottom=None, left=None, right=None)
_TYPE_A_MID = Border(top=None, bottom=None, left=None, right=None)
_TYPE_A_LABEL = Border(top=THIN, bottom=None, left=THIN, right=None)

_TYPE_B_TOP = Border(top=THIN, bottom=None, left=None, right=None)
_TYPE_B_MID = Border(top=THIN, bottom=None, left=None, right=None)
_TYPE_B_LABEL = Border(top=THIN, bottom=None, left=THIN, right=THIN)

_ACT_MERGED_BOTTOM = Border(top=None, bottom=None, left=THIN, right=None)


def write_metric_header(ws, r0, col0, target_label='유지계약\nA', pct_header_fill=PCT_HEADER_FILL):
    """3행 지표 헤더(유지계약~현장활동미확인, 폭 18칸)를 col0부터 그린다."""

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
    full3(col0 + 6, '전환미완료\nE (C-D)', right_border=None)

    for off in (3, 5, 7):
        col = col0 + off
        label_border = _TYPE_A_LABEL if off != 7 else Border(top=THIN, bottom=None, left=THIN, right=THIN)
        _hcell(ws, r0, col, border=_TYPE_A_TOP)
        _hcell(ws, r0 + 1, col, border=_TYPE_A_MID)
        _hcell(ws, r0 + 2, col, '%', fill=pct_header_fill, border=label_border)

    for off, label in ((8, '현장활동\n확인'), (16, '현장활동\n미확인')):
        col = col0 + off
        _hcell(ws, r0, col, border=_TYPE_A_TOP)
        ws.merge_cells(start_row=r0 + 1, start_column=col, end_row=r0 + 2, end_column=col)
        _hcell(ws, r0 + 1, col, label, border=_TYPE_A_LABEL)
        _hcell(ws, r0 + 2, col, border=_ACT_MERGED_BOTTOM)

    leaves = {
        9: '%', 10: '전환예정', 11: '연락두절', 12: '사고있음',
        13: '고객거부', 14: '압류계약', 15: 'ARS거부', 17: '%',
    }
    for off, label in leaves.items():
        col = col0 + off
        _hcell(ws, r0, col, border=_TYPE_B_TOP)
        _hcell(ws, r0 + 1, col, border=_TYPE_B_MID)
        fill = pct_header_fill if off in (9, 17) else HEADER_FILL
        _hcell(ws, r0 + 2, col, label, fill=fill, border=_TYPE_B_LABEL)

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


def _metric_border(ws, row, col0, top=True, bottom=True):
    """top/bottom=False로 주면 그룹 내부 행 사이의 가로줄을 없앤다(그룹 전체를
    하나의 칸처럼 보이게 하고, 그룹의 맨 위/맨 아래 행에서만 top/bottom=True로
    호출해 바깥 테두리만 남긴다)."""
    medium_left = {col0 + 4, col0 + 6}
    medium_right = {col0 + 5, col0 + 17}
    for c in range(col0, col0 + METRIC_WIDTH):
        left = MEDIUM if c in medium_left else THIN
        right = MEDIUM if c in medium_right else THIN
        ws.cell(row=row, column=c).border = Border(
            top=THIN if top else None, bottom=THIN if bottom else None, left=left, right=right)


# ---------------------------------------------------------------------------
# 체결기간별_부문별 전용 지표 블록(폭 20칸): 전환완료(D) 뒤에 '최대전환완료'
# 2칸(값+%)이 추가로 끼어든다 - 최대전환년수 이상 도달 + 완료(더 이상 다음
# 연차 대상에 남지 않는 건)만 별도로 보여주기 위함.
# ---------------------------------------------------------------------------
METRIC_WIDTH_PERIOD = 20


def write_metric_header_period(ws, r0, col0, target_label='대상계약\nA', pct_header_fill=PCT_HEADER_FILL):
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
        label_border = _TYPE_A_LABEL if off != 9 else Border(top=THIN, bottom=None, left=THIN, right=THIN)
        _hcell(ws, r0, col, border=_TYPE_A_TOP)
        _hcell(ws, r0 + 1, col, border=_TYPE_A_MID)
        _hcell(ws, r0 + 2, col, '%', fill=pct_header_fill, border=label_border)

    for off, label in ((6, '최대\n전환완료'), (10, '현장활동\n확인'), (18, '현장활동\n미확인')):
        col = col0 + off
        _hcell(ws, r0, col, border=_TYPE_A_TOP)
        ws.merge_cells(start_row=r0 + 1, start_column=col, end_row=r0 + 2, end_column=col)
        _hcell(ws, r0 + 1, col, label, border=_TYPE_A_LABEL)
        _hcell(ws, r0 + 2, col, border=_ACT_MERGED_BOTTOM)

    leaves = {
        7: '%', 11: '%', 12: '전환예정', 13: '연락두절', 14: '사고있음',
        15: '고객거부', 16: '압류계약', 17: 'ARS거부', 19: '%',
    }
    for off, label in leaves.items():
        col = col0 + off
        _hcell(ws, r0, col, border=_TYPE_B_TOP)
        _hcell(ws, r0 + 1, col, border=_TYPE_B_MID)
        fill = pct_header_fill if off in (7, 11, 19) else HEADER_FILL
        _hcell(ws, r0 + 2, col, label, fill=fill, border=_TYPE_B_LABEL)

    medium_left = {col0 + 4, col0 + 8}
    medium_right = {col0 + 7, col0 + 19}
    top_from = col0 + 4
    for rr in range(r0, r0 + 3):
        for c in range(col0, col0 + METRIC_WIDTH_PERIOD):
            cell = ws.cell(row=rr, column=c)
            b = cell.border
            left = MEDIUM if c in medium_left else b.left
            right = MEDIUM if c in medium_right else b.right
            top = MEDIUM if (rr == r0 and c >= top_from) else b.top
            cell.border = Border(top=top, bottom=b.bottom, left=left, right=right)

    # 최대전환완료(신규 삽입) 구획을 굵은 선으로 한 번 더 강조: M/N열의 위쪽
    # 빈칸(r0+1)과 라벨행 경계를 굵게 바꾼다.
    for off in (6, 7):
        col = col0 + off
        cell = ws.cell(row=r0 + 1, column=col)
        b = cell.border
        cell.border = Border(top=MEDIUM, bottom=b.bottom, left=b.left, right=b.right)

    # M열(최대전환완료) 라벨행의 왼쪽 테두리도 굵게(7행 빈칸은 L-M 사이에
    # 세로선이 없어야 하므로 그대로 두고, 8~9행 라벨 부분만 굵게 바꾼다).
    m_col = col0 + 6
    for rr in (r0 + 1, r0 + 2):
        cell = ws.cell(row=rr, column=m_col)
        b = cell.border
        cell.border = Border(top=b.top, bottom=b.bottom, left=MEDIUM, right=b.right)


def _metric_fill_period(ws, row, col0, fill):
    for c in range(col0, col0 + METRIC_WIDTH_PERIOD):
        ws.cell(row=row, column=c).fill = fill


def _metric_border_period(ws, row, col0, top=True, bottom=True):
    medium_left = {col0 + 4, col0 + 6, col0 + 8}
    medium_right = {col0 + 5, col0 + 7, col0 + 19}
    for c in range(col0, col0 + METRIC_WIDTH_PERIOD):
        left = MEDIUM if c in medium_left else THIN
        right = MEDIUM if c in medium_right else THIN
        ws.cell(row=row, column=c).border = Border(
            top=THIN if top else None, bottom=THIN if bottom else None, left=left, right=right)
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

"""
write_report_v7.py — calc_report_v7.py로 계산한 값을, build_report_v7.py가 만든
레이아웃(서식/병합/색상)에 실제로 채워 넣어 최종 보고서 파일을 만든다.
LibreOffice/Excel 수식은 전혀 쓰지 않는다(순수 값만 기록).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill

NO_FILL = PatternFill(fill_type=None)
WHITE_FILL = PatternFill('solid', fgColor='FFFFFF')

THIN_SIDE = Side(style='thin')

layout = sys.modules[__name__]
calc = sys.modules[__name__]

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


# "월별_부문별_new"/"월별_부문별_new_상세"가 공유하는 B:E 18행 지표 라벨 블록.
# 원본 템플릿에서 셀 단위로 옮긴 테두리 값(1=THIN, 0=없음) - row0(그 연도 블록의
# 유지계약 A 행) 기준 오프셋으로 정의.
_NEW_SHEET_LABEL_BORDERS_BY_OFF = {
    0: ((1, 1, 1, 1), (1, 1, 0, 0), (1, 1, 0, 0), (1, 1, 0, 1)),
    1: ((1, 1, 1, 1), (1, 1, 0, 0), (1, 1, 0, 0), (1, 1, 0, 1)),
    2: ((1, 0, 1, 1), (1, 0, 0, 0), (1, 0, 0, 0), (1, 0, 0, 1)),
    3: ((0, 1, 1, 0), (0, 1, 0, 0), (0, 1, 0, 0), (1, 1, 1, 1)),
    4: ((1, 0, 1, 1), (1, 0, 0, 0), (1, 0, 0, 0), (1, 0, 0, 1)),
    5: ((0, 1, 1, 0), (0, 1, 0, 0), (0, 1, 0, 0), (1, 1, 1, 1)),
    6: ((0, 0, 1, 1), (0, 0, 0, 0), (0, 0, 0, 0), (0, 0, 0, 1)),
    7: ((0, 0, 1, 0), (0, 0, 0, 0), (0, 0, 0, 0), (1, 0, 1, 1)),
    8: ((0, 0, 1, 0), (1, 0, 1, 1), (1, 0, 0, 0), (1, 0, 0, 1)),
    9: ((0, 0, 1, 0), (0, 0, 1, 0), (0, 0, 0, 0), (1, 1, 1, 1)),
    10: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
    11: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
    12: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
    13: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
    14: ((0, 0, 1, 0), (0, 0, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
    15: ((0, 0, 1, 0), (0, 1, 1, 0), (1, 1, 1, 1), (1, 1, 0, 1)),
    16: ((0, 0, 1, 0), (1, 0, 1, 1), (1, 0, 0, 0), (1, 0, 0, 1)),
    17: ((0, 1, 1, 0), (0, 1, 1, 0), (0, 1, 0, 0), (1, 1, 1, 1)),
}
_NEW_SHEET_PCT_ONLY_OFFS = {3, 5, 7, 9, 17}
_NEW_SHEET_LEAF_LABELS_BY_OFF = {
    10: '전환예정', 11: '연락두절', 12: '사고있음',
    13: '고객거부', 14: '압류계약', 15: 'ARS거부',
}
_NEW_SHEET_N_METRIC_ROWS = 18


def _write_new_sheet_label_block(ws, row0, col_label=2):
    col_B, col_C, col_D, col_E = col_label, col_label + 1, col_label + 2, col_label + 3

    def _b(spec):
        THIN = layout.THIN
        s = lambda v: THIN if v else None
        return Border(top=s(spec[0]), bottom=s(spec[1]), left=s(spec[2]), right=s(spec[3]))

    label_rows = {
        row0 + 0: ('도래계약(유지중) A', col_B, col_E),
        row0 + 1: ('사고有 B', col_B, col_E),
        row0 + 2: ('전환대상 C (A-B)', col_B, col_E),
        row0 + 4: ('전환완료 D', col_B, col_E),
        row0 + 6: ('전환미완료 E (C-D)', col_B, col_E),
        row0 + 8: ('현장활동 확인', col_C, col_E),
        row0 + 16: ('현장활동 미확인', col_C, col_E),
    }
    for off in range(_NEW_SHEET_N_METRIC_ROWS):
        r = row0 + off
        is_pct_row = off in _NEW_SHEET_PCT_ONLY_OFFS
        b_spec, c_spec, d_spec, e_spec = _NEW_SHEET_LABEL_BORDERS_BY_OFF[off]
        e_fill = layout.PCT_HEADER_FILL if is_pct_row else NO_FILL
        _hcell(ws, r, col_B, fill=layout.HEADER_FILL, border=_b(b_spec))
        _hcell(ws, r, col_C, fill=layout.HEADER_FILL, border=_b(c_spec))
        _hcell(ws, r, col_D, fill=layout.HEADER_FILL, border=_b(d_spec))
        _hcell(ws, r, col_E, fill=e_fill, border=_b(e_spec))
    for r, (text, start_c, end_c) in label_rows.items():
        if start_c != end_c:
            ws.merge_cells(start_row=r, start_column=start_c, end_row=r, end_column=end_c)
        ws.cell(row=r, column=start_c, value=text)
    for off in _NEW_SHEET_PCT_ONLY_OFFS | {17}:
        ws.cell(row=row0 + off, column=col_E, value='%')
    for off, text in _NEW_SHEET_LEAF_LABELS_BY_OFF.items():
        r = row0 + off
        ws.merge_cells(start_row=r, start_column=col_D, end_row=r, end_column=col_E)
        ws.cell(row=r, column=col_D, value=text)


def build_month_dept_new(wb, months, year_month_acc, years, reference_year):
    """years: 위에서부터 표시할 연도 목록(예: [2025, 2026], 나중에 2024/2027 등을
    앞뒤에 추가하면 자동으로 확장됨) - 연도마다 표 전체를 아래로 쌓는다(같은
    열을 재사용). year_month_acc: calc.compute_all()의 result['year_month_new']
    - 계약별로 각 연도 시점 실제 연차를 역산해서 그 연차 기준 v값/목표플랜으로
    이미 다시 판정해 둔 (연도,월,부문) 집계이므로, 여기서는 그대로 부문 합산/월
    합산만 하면 된다. reference_year보다 미래인 연도는 실제로 계산할 방법이
    없으므로 그 연도 블록은 전부 빈칸으로 둔다."""
    ws = wb.create_sheet('월별_부문별_new')
    ws.sheet_view.showGridLines = False

    col_label = 2  # B (B:E 4칸)
    n_label_cols = 4
    n_header_rows = 3
    n_metric_rows = _NEW_SHEET_N_METRIC_ROWS
    year_block_height = n_header_rows + n_metric_rows + 1  # +1 연도 사이 구분 여백행
    cols_per_year = 16  # 1~12월 + 누계 + 부문(개인/전략/신사업)

    col_B, col_C, col_D, col_E = col_label, col_label + 1, col_label + 2, col_label + 3
    col0 = col_label + n_label_cols  # F

    tall_rows = []
    for yi, year in enumerate(years):
        is_future = year > reference_year
        hdr_r0 = 2 + yi * year_block_height
        row0 = hdr_r0 + n_header_rows
        tall_rows.append(hdr_r0)

        ws.merge_cells(start_row=hdr_r0, start_column=col_label, end_row=hdr_r0 + 2, end_column=col_label + n_label_cols - 1)
        for rr in range(hdr_r0, hdr_r0 + 3):
            for c in range(col_label, col_label + n_label_cols):
                _hcell(ws, rr, c)
        _hcell(ws, hdr_r0, col_label, '구분')
        _write_new_sheet_label_block(ws, row0, col_label)

        year_last_col = col0 + cols_per_year - 1
        ws.merge_cells(start_row=hdr_r0, start_column=col0, end_row=hdr_r0, end_column=year_last_col)
        for c in range(col0, col0 + cols_per_year):
            _hcell(ws, hdr_r0, c, border=Border(top=layout.THIN, bottom=layout.THIN, left=layout.THIN, right=layout.THIN))
        ws.cell(row=hdr_r0, column=col0, value=year)

        month_aggs = []
        for mi, month in enumerate(months):
            mcol = col0 + mi
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

        total_col = col0 + len(months)
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
    for c in range(col0, col0 + cols_per_year):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13.0
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    for yi in range(len(years)):
        hdr_r0 = 2 + yi * year_block_height
        row0 = hdr_r0 + n_header_rows
        for off in (0, 1, 2, 4, 6, 8, 16):
            ws.row_dimensions[row0 + off].height = 16.5
    return ws


def build_month_dept_new_detail(wb, months, year_month_tier_acc, years, reference_year, n_tiers=3):
    """'월별_부문별_new_상세': "월별_부문별_new"와 같은 지표를, 월/누계/부문 각
    칸을 다시 연차(1~n_tiers)별로 쪼개 보여준다. year_month_tier_acc: calc.
    compute_all()의 result['year_month_tier_new'] - (연도,월,부문,연차e)별로
    이미 재판정해 둔 집계."""
    ws = wb.create_sheet('월별_부문별_new_상세')
    ws.sheet_view.showGridLines = False

    col_label = 2  # B (B:E 4칸)
    n_label_cols = 4
    n_header_rows = 4
    n_metric_rows = _NEW_SHEET_N_METRIC_ROWS
    year_block_height = n_header_rows + n_metric_rows + 1  # +1 연도 사이 구분 여백행
    cols_per_group = n_tiers
    col0 = col_label + n_label_cols  # F

    THIN = layout.THIN

    def _write_group_header_2row(gcol0, r1, r2, label):
        ws.merge_cells(start_row=r1, start_column=gcol0, end_row=r2, end_column=gcol0 + cols_per_group - 1)
        for off in range(cols_per_group):
            c = gcol0 + off
            is_first, is_last = off == 0, off == cols_per_group - 1
            fill = layout.HEADER_FILL if is_first else NO_FILL
            b1 = Border(top=None, bottom=None, left=(THIN if is_first else None), right=(THIN if is_last else None))
            b2 = Border(top=None, bottom=None, left=(THIN if is_first else None), right=(THIN if is_last else None))
            _hcell(ws, r1, c, fill=fill, border=b1)
            _hcell(ws, r2, c, fill=NO_FILL, border=b2)
        ws.cell(row=r1, column=gcol0, value=label)

    def _write_group_header_1row(gcol0, r, label):
        ws.merge_cells(start_row=r, start_column=gcol0, end_row=r, end_column=gcol0 + cols_per_group - 1)
        for off in range(cols_per_group):
            c = gcol0 + off
            is_first, is_last = off == 0, off == cols_per_group - 1
            fill = layout.HEADER_FILL if is_first else NO_FILL
            b = Border(top=THIN, bottom=THIN, left=(THIN if is_first else None), right=(THIN if is_last else None))
            _hcell(ws, r, c, fill=fill, border=b)
            # 부문 그룹 위(월 헤더와 같은 줄) 칸도 채우기색만 이어서 준다(테두리 없음).
            _hcell(ws, r - 1, c, fill=layout.HEADER_FILL, border=Border())
        ws.cell(row=r, column=gcol0, value=label)

    def _write_tier_labels(gcol0, r):
        for ti in range(n_tiers):
            _hcell(ws, r, gcol0 + ti, f'{ti + 1}년차',
                   border=Border(top=THIN, bottom=THIN, left=THIN, right=THIN))

    tall_rows = []
    for yi, year in enumerate(years):
        is_future = year > reference_year
        hdr_r0 = 2 + yi * year_block_height
        row0 = hdr_r0 + n_header_rows
        tall_rows.append(hdr_r0)
        cols_per_year = (len(months) + 1 + len(DEPT_ORDER)) * cols_per_group

        ws.merge_cells(start_row=hdr_r0, start_column=col_label, end_row=hdr_r0 + 3, end_column=col_label + n_label_cols - 1)
        for rr in range(hdr_r0, hdr_r0 + 4):
            for c in range(col_label, col_label + n_label_cols):
                _hcell(ws, rr, c)
        _hcell(ws, hdr_r0, col_label, '구분')
        _write_new_sheet_label_block(ws, row0, col_label)

        year_last_col = col0 + cols_per_year - 1
        ws.merge_cells(start_row=hdr_r0, start_column=col0, end_row=hdr_r0, end_column=year_last_col)
        for c in range(col0, col0 + cols_per_year):
            _hcell(ws, hdr_r0, c, border=Border(top=THIN, bottom=THIN, left=THIN, right=THIN))
        ws.cell(row=hdr_r0, column=col0, value=year)

        month_aggs_by_tier = {ti: [] for ti in range(n_tiers)}
        for mi, month in enumerate(months):
            gcol0 = col0 + mi * cols_per_group
            _write_group_header_2row(gcol0, hdr_r0 + 1, hdr_r0 + 2, f'{int(month) if str(month).isdigit() else month}월')
            _write_tier_labels(gcol0, hdr_r0 + 3)
            for ti in range(n_tiers):
                e = ti + 1
                agg = None if is_future else _agg_sum(
                    [year_month_tier_acc.finalize((year, month, d, e)) for d in DEPT_ORDER])
                month_aggs_by_tier[ti].append(agg)
                write_metric_values_transposed(ws, gcol0 + ti, row0, agg)

        total_gcol0 = col0 + len(months) * cols_per_group
        _write_group_header_2row(total_gcol0, hdr_r0 + 1, hdr_r0 + 2, '누계')
        _write_tier_labels(total_gcol0, hdr_r0 + 3)
        for ti in range(n_tiers):
            write_metric_values_transposed(ws, total_gcol0 + ti, row0, _agg_sum(month_aggs_by_tier[ti]))

        dept_gcol0_start = total_gcol0 + cols_per_group
        for di, (d, short) in enumerate(zip(DEPT_ORDER, DEPTS_SHORT)):
            gcol0 = dept_gcol0_start + di * cols_per_group
            _write_group_header_1row(gcol0, hdr_r0 + 2, short)
            _write_tier_labels(gcol0, hdr_r0 + 3)
            for ti in range(n_tiers):
                e = ti + 1
                dept_month_aggs = None if is_future else [
                    year_month_tier_acc.finalize((year, month, d, e)) for month in months]
                write_metric_values_transposed(ws, gcol0 + ti, row0,
                                                _agg_sum(dept_month_aggs) if dept_month_aggs else None)

        # 표 전체 맨 오른쪽 끝(신사업 마지막 칸)은 hdr_r0+1행까지 바깥 테두리가 이어진다.
        _hcell(ws, hdr_r0 + 1, year_last_col, fill=layout.HEADER_FILL, border=Border(right=THIN))

    ws.column_dimensions[openpyxl.utils.get_column_letter(col_label)].width = 3.125
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_label + 1)].width = 13.0
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_label + 2)].width = 8.125
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_label + 3)].width = 5.75
    max_cols_per_year = (len(months) + 1 + len(DEPT_ORDER)) * cols_per_group
    for c in range(col0, col0 + max_cols_per_year):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 6.25
    for r in tall_rows:
        ws.row_dimensions[r].height = 17.25
    for yi in range(len(years)):
        hdr_r0 = 2 + yi * year_block_height
        row0 = hdr_r0 + n_header_rows
        for off in (0, 1, 2, 4, 6, 8, 16):
            ws.row_dimensions[row0 + off].height = 16.5
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
         '3년차 기준(그때의 사고 판정)으로 다시 계산해서 반영합니다. 아직 그 해에 존재하지도 않았던 '
         '계약(예: 올해 막 1년차가 된 계약의 2025년 이전 시점)은 해당 연도 집계에서 완전히 제외됩니다.')
    para('전환완료 판정의 한계: "그 해에 정확히 완료였는지"는 데이터로 알 수 없습니다(과거 시점의 '
         '플랜 이력이 남아있지 않고, 오늘 기준 플랜만 있음). 그래서 "오늘 기준으로 이미 그 연차 '
         '목표 이상을 달성했다면, 그 이전 연도에도 완료였다"고 봅니다 - 플랜은 한번 올라가면 내려가지 '
         '않으므로 이건 확실한 사실이지만, "정확히 그 해에" 달성했는지(예: 2025년에 이미 됐는지, '
         '아니면 2025~2026년 사이 어느 시점에 됐는지)까지는 구분하지 못합니다.')
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


def build_reference_sheet(wb, months, month_dept_acc, current_month=None,
                           year_month_tier_acc=None, new_sheet_years=None, reference_year=None):
    """'참고' 시트: 완료율에 영향을 주는 주요 변수(1년차 진입 후 경과개월수)와
    완료율의 상관관계를 실무자가 참고할 수 있도록 정리. 매 실행마다 실제
    데이터로 다시 계산되며, 보고서의 다른 계산에는 영향을 주지 않는다.
    current_month을 생략하면 실행 시점의 실제 오늘 월을 쓴다. year_month_tier_acc/
    new_sheet_years/reference_year를 주면 "월별_부문별_new_상세" 재구성값 해석
    주의사항 섹션에 실제 1년차 비교표를 함께 넣는다(생략하면 그 섹션은 설명
    문단만 나온다)."""
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
    blank()

    section('월별_부문별_new / _상세 재구성값 해석 시 주의사항')
    para('"월별_부문별_new"와 "_상세" 시트에서, 기준연도(오늘 기준 연도)보다 이전 연도의 값은 '
         '"그 해에 실제로 완료였는지"가 아니라 "지금 시점까지 누적으로 봤을 때 그 단계를 넘었는지"를 '
         '보여줍니다. 계약별로 과거 특정 시점에 어떤 플랜을 갖고 있었는지 이력이 남아있지 않고 현재 '
         '플랜만 알 수 있기 때문에, 과거 시점의 완료 여부는 "현재까지 최소한 그 단계는 도달했는가"로 '
         '근사할 수밖에 없습니다.')
    para('그래서 같은 "N년차" 라벨이라도 연도가 다르면 실제로는 다른 가입월 코호트를 가리킵니다 - '
         '예를 들어 기준연도의 1년차는 그 해 갓 가입해서 지금 딱 1년차인 계약의 진짜 현재 데이터지만, '
         '한 해 전(예: 기준연도－1년)의 1년차는 이미 2년차 이상인 계약을 "최소 1년차는 넘겼는지"로 '
         '되짚어본 값입니다. 후자는 그 뒤로 시간이 더 흐르며 따라잡은 것까지 포함된 누적치라서, 보통 '
         '"진짜 그 시점의 완료 속도"보다 더 높게 나옵니다 - 특히 가입 후 1년 안에는 담당자 컨택이 많아 '
         '전환을 잘 챙기는 경향이 있다는 점을 감안하면, 이 차이는 실제 업무 속도 차이라기보다 재구성 '
         '방식 자체의 한계로 보는 게 맞습니다.')

    if year_month_tier_acc is not None and new_sheet_years and len(new_sheet_years) >= 2 and reference_year is not None:
        year_old, year_new = new_sheet_years[-2], new_sheet_years[-1]
        r = row[0]
        headers = ['계약체결월', f'{year_old}년 1년차 완료율(%)', f'{year_new}년 1년차 완료율(%)', '차이(%p)']
        for i, h in enumerate(headers):
            c = ws.cell(row=r, column=2 + i, value=h)
            c.font = HEAD_F
            c.fill = HEAD_FILL
            c.border = BORDER
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[r].height = 28
        row[0] += 1

        def _rate_e1(year, month):
            agg = _agg_sum([year_month_tier_acc.finalize((year, month, d, 1)) for d in DEPT_ORDER])
            if agg is None or not agg['conv_target']:
                return None
            return round(agg['done'] / agg['conv_target'] * 100, 1)

        for month in months:
            mi = int(month) if str(month).isdigit() else month
            r_old = _rate_e1(year_old, month)
            r_new = _rate_e1(year_new, month)
            gap = round(r_old - r_new, 1) if (r_old is not None and r_new is not None) else None
            r = row[0]
            for i, v in enumerate([f'{mi}월', r_old, r_new, gap]):
                c = ws.cell(row=r, column=2 + i, value=v)
                c.font = BODY_F
                c.border = BORDER
                c.alignment = Alignment(horizontal='center', vertical='center')
            row[0] += 1
        blank()
        para(f'위 표에서 {year_old}년 1년차는 {year_new}년 1년차보다 거의 매달 높게 나옵니다 - 실제로 '
             f'더 잘 전환해서가 아니라, {year_old}년 1년차 값은 지금(현재월보다 이전 달들 기준)까지 '
             f'누적된 값이기 때문입니다(현재월보다 이후 달은 재구성이 적용되지 않아 두 값이 같습니다). '
             '이 표의 차이(%p)를 "그 해의 실제 업무 속도"로 해석하지 마시고, 재구성 방식의 특성으로 '
             '이해해 주세요.', height=45)

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
    build_month_dept_new_detail(wb, result['months'], result['year_month_tier_new'], years=result['new_sheet_years'],
                                 reference_year=result['reference_year'])
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
    build_reference_sheet(wb, result['months'], result['month_dept'], current_month=result['current_month'],
                           year_month_tier_acc=result['year_month_tier_new'], new_sheet_years=result['new_sheet_years'],
                           reference_year=result['reference_year'])
    wb.save(output_path)
    print('저장 완료 ->', output_path)

"""gui_app.py — 무사고전환 보고서 생성기 (데스크톱 GUI).
PyInstaller로 exe를 만들 때 이 파일을 진입점으로 사용한다.
파이썬 표준 라이브러리(tkinter)만 사용하므로 openpyxl 외 추가 의존성이 없다.
"""
import os
import queue
import sys
import threading
import traceback
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = '무사고전환 보고서 생성기'
MAX_PERIOD_ROWS = 7


class TextRedirector:
    """print() 출력을 큐를 통해 GUI 로그창으로 보낸다(스레드 안전)."""

    def __init__(self, msg_queue):
        self.msg_queue = msg_queue

    def write(self, s):
        if s:
            self.msg_queue.put(('log', s))

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('760x900')
        self.minsize(680, 700)

        self.data_files = []
        self.mapping_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.ref_date = tk.StringVar(value=date.today().strftime('%Y-%m'))
        today = date.today()
        self.new_sheet_years = tk.StringVar(value=f'{today.year - 1}~{today.year}')
        self.period_vars = []
        for i in range(MAX_PERIOD_ROWS):
            if i < len(PERIODS):
                _label, _key, start, end = PERIODS[i]
            else:
                start, end = '', ''
            self.period_vars.append((tk.StringVar(value=start), tk.StringVar(value=end)))
        self.msg_queue = queue.Queue()
        self.worker = None

        self._build_widgets()
        self.after(100, self._poll_queue)

    def _build_widgets(self):
        pad = {'padx': 10, 'pady': 6}

        frm_files = ttk.LabelFrame(self, text='1. 분석 데이터 파일 (여러 개 선택 가능)')
        frm_files.pack(fill='x', **pad)

        self.lst_files = tk.Listbox(frm_files, height=6, selectmode='extended')
        self.lst_files.pack(side='left', fill='both', expand=True, padx=(10, 0), pady=10)

        scr = ttk.Scrollbar(frm_files, orient='vertical', command=self.lst_files.yview)
        scr.pack(side='left', fill='y', pady=10)
        self.lst_files.config(yscrollcommand=scr.set)

        btns = ttk.Frame(frm_files)
        btns.pack(side='left', fill='y', padx=10, pady=10)
        ttk.Button(btns, text='파일 추가...', command=self._add_files).pack(fill='x', pady=(0, 4))
        ttk.Button(btns, text='선택 삭제', command=self._remove_selected).pack(fill='x', pady=(0, 4))
        ttk.Button(btns, text='전체 삭제', command=self._clear_files).pack(fill='x')

        frm_map = ttk.LabelFrame(self, text='2. 매핑테이블 파일 ("무사고전환매핑테이블" 시트가 있는 파일)')
        frm_map.pack(fill='x', **pad)
        row_map = ttk.Frame(frm_map)
        row_map.pack(fill='x', padx=10, pady=(10, 2))
        ent_map = ttk.Entry(row_map, textvariable=self.mapping_path)
        ent_map.pack(side='left', fill='x', expand=True)
        ttk.Button(row_map, text='찾아보기...', command=self._choose_mapping).pack(side='left', padx=(6, 0))
        ttk.Label(frm_map, text='비워두면 위 분석 데이터 중 첫 번째 파일을 매핑테이블로도 사용합니다.',
                  foreground='#666666').pack(fill='x', padx=10, pady=(0, 10))

        frm_ref = ttk.LabelFrame(self, text='3. 계산 기준 시점 (YYYY-MM)')
        frm_ref.pack(fill='x', **pad)
        row_ref = ttk.Frame(frm_ref)
        row_ref.pack(fill='x', padx=10, pady=(10, 2))
        ttk.Entry(row_ref, textvariable=self.ref_date, width=12).pack(side='left')
        ttk.Label(frm_ref,
                  text='"월별_부문별_new"/"참고" 시트가 기준으로 삼는 시점입니다. '
                       '기본값은 오늘 날짜이며, 과거 특정 시점 기준으로 계산하고 싶을 때만 바꾸세요.',
                  foreground='#666666').pack(fill='x', padx=10, pady=(0, 10))

        frm_years = ttk.LabelFrame(self, text='4. 월별_부문별_new 표시연도 (YYYY~YYYY)')
        frm_years.pack(fill='x', **pad)
        row_years = ttk.Frame(frm_years)
        row_years.pack(fill='x', padx=10, pady=(10, 2))
        ttk.Entry(row_years, textvariable=self.new_sheet_years, width=14).pack(side='left')
        ttk.Label(frm_years,
                  text='"월별_부문별_new" 시트에 나란히 보여줄 연도 범위입니다(양 끝 연도 포함). '
                       '예: 2025~2026, 2024~2026.',
                  foreground='#666666').pack(fill='x', padx=10, pady=(0, 10))

        frm_periods = ttk.LabelFrame(
            self, text='5. 체결기간별 구간 설정 (최대 7개, 시작~종료 YYYYMM, 비워두면 그 구간은 사용 안 함)')
        frm_periods.pack(fill='x', **pad)
        grid_periods = ttk.Frame(frm_periods)
        grid_periods.pack(fill='x', padx=10, pady=(10, 2))
        ttk.Label(grid_periods, text='구간').grid(row=0, column=0, padx=(0, 6), pady=2)
        ttk.Label(grid_periods, text='시작(YYYYMM)').grid(row=0, column=1, padx=(0, 6), pady=2)
        ttk.Label(grid_periods, text='종료(YYYYMM)').grid(row=0, column=2, pady=2)
        for i, (start_var, end_var) in enumerate(self.period_vars):
            ttk.Label(grid_periods, text=f'{i + 1}구간').grid(row=i + 1, column=0, padx=(0, 6), pady=2, sticky='w')
            ttk.Entry(grid_periods, textvariable=start_var, width=10).grid(row=i + 1, column=1, padx=(0, 6), pady=2)
            ttk.Entry(grid_periods, textvariable=end_var, width=10).grid(row=i + 1, column=2, pady=2)
        ttk.Label(frm_periods, text='보험기간 시작일 기준 12개월 단위 구간입니다. 채워진 구간만 계산에 사용됩니다.',
                  foreground='#666666').pack(fill='x', padx=10, pady=(4, 10))

        frm_out = ttk.LabelFrame(self, text='6. 결과 저장 위치')
        frm_out.pack(fill='x', **pad)
        ent = ttk.Entry(frm_out, textvariable=self.output_path)
        ent.pack(side='left', fill='x', expand=True, padx=(10, 0), pady=10)
        ttk.Button(frm_out, text='찾아보기...', command=self._choose_output).pack(side='left', padx=10, pady=10)

        frm_run = ttk.Frame(self)
        frm_run.pack(fill='x', **pad)
        self.btn_run = ttk.Button(frm_run, text='보고서 생성', command=self._run)
        self.btn_run.pack(side='left')
        self.progress = ttk.Progressbar(frm_run, mode='indeterminate')
        self.progress.pack(side='left', fill='x', expand=True, padx=10)

        frm_log = ttk.LabelFrame(self, text='진행 상황')
        frm_log.pack(fill='both', expand=True, **pad)
        self.txt_log = tk.Text(frm_log, height=12, state='disabled', wrap='word')
        self.txt_log.pack(fill='both', expand=True, padx=10, pady=10)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title='데이터 파일 선택', filetypes=[('Excel 파일', '*.xlsx *.xlsm *.xls')])
        for p in paths:
            if p not in self.data_files:
                self.data_files.append(p)
                self.lst_files.insert('end', p)

    def _remove_selected(self):
        for i in reversed(self.lst_files.curselection()):
            del self.data_files[i]
            self.lst_files.delete(i)

    def _clear_files(self):
        self.data_files.clear()
        self.lst_files.delete(0, 'end')

    def _choose_mapping(self):
        path = filedialog.askopenfilename(
            title='매핑테이블 파일 선택', filetypes=[('Excel 파일', '*.xlsx *.xlsm *.xls')])
        if path:
            self.mapping_path.set(path)

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title='결과 저장 위치', defaultextension='.xlsx',
            filetypes=[('Excel 파일', '*.xlsx')], initialfile='무사고전환_보고서.xlsx')
        if path:
            self.output_path.set(path)

    def _log(self, text):
        self.txt_log.config(state='normal')
        self.txt_log.insert('end', text)
        self.txt_log.see('end')
        self.txt_log.config(state='disabled')

    def _run(self):
        if not self.data_files:
            messagebox.showwarning(APP_TITLE, '데이터 파일을 하나 이상 선택해 주세요.')
            return
        out = self.output_path.get().strip()
        if not out:
            messagebox.showwarning(APP_TITLE, '결과 저장 위치를 지정해 주세요.')
            return

        mapping = self.mapping_path.get().strip() or self.data_files[0]

        ref_date_str = self.ref_date.get().strip()
        try:
            ref_year, ref_month = (int(p) for p in ref_date_str.split('-'))
            if not (1 <= ref_month <= 12):
                raise ValueError
        except ValueError:
            messagebox.showwarning(APP_TITLE, '계산 기준 시점은 "YYYY-MM" 형식으로 입력해 주세요. (예: 2026-09)')
            return

        years_str = self.new_sheet_years.get().strip()
        try:
            y_start, y_end = (int(p) for p in years_str.split('~'))
            if y_start > y_end:
                raise ValueError
        except ValueError:
            messagebox.showwarning(APP_TITLE, '표시연도는 "YYYY~YYYY" 형식으로 입력해 주세요. (예: 2025~2026)')
            return
        new_sheet_years = list(range(y_start, y_end + 1))

        period_pairs = []
        for i, (start_var, end_var) in enumerate(self.period_vars):
            start = start_var.get().strip()
            end = end_var.get().strip()
            if not start and not end:
                continue
            if not (len(start) == 6 and start.isdigit() and len(end) == 6 and end.isdigit() and start <= end):
                messagebox.showwarning(
                    APP_TITLE, f'{i + 1}구간의 시작/종료를 "YYYYMM" 형식으로 올바르게 입력해 주세요. (예: 202207)')
                return
            period_pairs.append((start, end))
        if not period_pairs:
            messagebox.showwarning(APP_TITLE, '체결기간별 구간을 하나 이상 입력해 주세요.')
            return

        self.btn_run.config(state='disabled')
        self.progress.start(12)
        self.txt_log.config(state='normal')
        self.txt_log.delete('1.0', 'end')
        self.txt_log.config(state='disabled')

        self.worker = threading.Thread(
            target=self._run_worker,
            args=(list(self.data_files), mapping, out, ref_year, ref_month, new_sheet_years, period_pairs),
            daemon=True)
        self.worker.start()

    def _run_worker(self, data_paths, mapping_path, out, ref_year, ref_month, new_sheet_years, period_pairs):
        old_stdout = sys.stdout
        sys.stdout = TextRedirector(self.msg_queue)
        try:
            build(data_paths, mapping_path, out, reference_year=ref_year, current_month=ref_month,
                  new_sheet_years=new_sheet_years, period_pairs=period_pairs)
            self.msg_queue.put(('done', out))
        except Exception:
            self.msg_queue.put(('error', traceback.format_exc()))
        finally:
            sys.stdout = old_stdout

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == 'log':
                    self._log(payload)
                elif kind == 'done':
                    self.progress.stop()
                    self.btn_run.config(state='normal')
                    messagebox.showinfo(APP_TITLE, f'보고서 생성이 완료됐습니다.\n\n{payload}')
                elif kind == 'error':
                    self.progress.stop()
                    self.btn_run.config(state='normal')
                    self._log('\n[오류]\n' + payload)
                    messagebox.showerror(APP_TITLE, '보고서 생성 중 오류가 발생했습니다.\n로그 창을 확인해 주세요.')
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
