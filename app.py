# 응급실 2026년 4월 근무표 자동 생성기
# 순수 Python 그리디 스케줄러 (서버 충돌 없음)

from flask import Flask, render_template, request, send_file, redirect, url_for
from datetime import date, timedelta
import calendar as _calendar
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import random

app = Flask(__name__)

# ── 기본 상수 ──────────────────────────────────────────────────────────────────
YEAR, MONTH, NUM_DAYS = 2026, 4, 30


def set_period(year: int, month: int):
    """연도·월을 변경할 때 전역 상수를 갱신합니다."""
    global YEAR, MONTH, NUM_DAYS
    YEAR = year
    MONTH = month
    NUM_DAYS = _calendar.monthrange(year, month)[1]
    # 말일 단독 N(예: 3-3-1·2-3-1의 끝 1): 허용 일자는 항상 NUM_DAYS(당월 마지막 날).
    # 31일로 끝나는 달은 NUM_DAYS==31이므로 31일에 단독 N 1개 적용이 가능하다.

# NO: 야간(N) 누적 20회마다 생기는 휴무(OF 성격). 발생일은 사람마다 다름(대략 3개월에 1회 수준).
#     자동 근무 생성 시 배정하지 않음 — 신청/근무표에서 수기 입력.
SHIFT_NAMES = ['A1', 'D', 'E', 'N', 'OF', 'EDU', '연', '공', '병', '경', 'OH', 'NO']
A1_S, D_S, E_S, N_S, OF_S, EDU_S, YUN_S, GONG_S, BYUNG_S, GYUNG_S, OH_S, NO_S = range(12)

# 연속근무 최대 5일 제한에 포함되는 시프트 (연차·병가·경조 등 휴가는 제외 — D/E/N/공/EDU만)
STREAK_WORK_SHIFTS = frozenset({'D', 'E', 'N', 'EDU', '공'})

# 화면 색상 (연한 파스텔 배경 — 눈 피로 완화, 글자는 진한 톤으로 대비 유지)
SHIFT_COLORS = {
    'A1': '#E3EEF9', 'D': '#FFF9C4', 'E': '#FFE8E0', 'N': '#E8EAF6',
    'OF': '#F5F5F5', 'EDU': '#E8F5E9', '연': '#FCE4EC', '공': '#F3E5F5',
    '병': '#FFEBEE', '경': '#E0F2F1', 'OH': '#FFF3E0',
    'NO': '#ECEFF1',   # N 20회 등 수기 휴무 (OF와 구분)
}
SHIFT_TEXT_COLORS = {
    'A1': '#1565C0', 'D': '#F57F17', 'E': '#E65100', 'N': '#283593',
    'OF': '#78909C', 'EDU': '#2E7D32', '연': '#AD1457', '공': '#6A1B9A',
    '병': '#C62828', '경': '#00695C', 'OH': '#E65100', 'NO': '#37474F',
}

# 마지막 생성 결과 임시 저장 (단일 사용자 로컬 용도)
_last_result = {}


# ── 유틸리티 함수 ──────────────────────────────────────────────────────────────
def get_nurse_names(num_nurses):
    return ['수간호사'] + [f'간호사{i}' for i in range(1, num_nurses)]


def get_april_days(holidays=()):
    holiday_set = set(holidays)
    weekday_names = ['월', '화', '수', '목', '금', '토', '일']
    days = []
    for d in range(1, NUM_DAYS + 1):
        dt = date(YEAR, MONTH, d)
        days.append({
            'day': d,
            'date': dt,
            'weekday': dt.weekday(),
            'weekday_name': weekday_names[dt.weekday()],
            'is_weekend': dt.weekday() >= 5,
            'is_holiday': d in holiday_set,
        })
    return days


def _carry_week_prev_month_off_counts(
    carry: dict, n: int, monday_date: date, month_first: date,
) -> tuple[int, int, int, int]:
    """
    월~일 한 주의 월요일(monday_date) 기준, month_first(당월 1일) **이전**인 날만 carry에서 읽어
    OF/OH/NO 개수와 그 일수를 반환. 전월 말주~당월 첫 주가 이어지는 경우 주간 휴무 판정에 합산한다.
    """
    pre_of = pre_oh = pre_no = 0
    n_prev = 0
    for i in range(7):
        d = monday_date + timedelta(days=i)
        if d >= month_first:
            break
        n_prev += 1
        k = (month_first - d).days
        c = carry.get(n) or ()
        if not (1 <= k <= len(c)):
            continue
        s = c[-k]
        if s == 'OF':
            pre_of += 1
        elif s == 'OH':
            pre_oh += 1
        elif s == 'NO':
            pre_no += 1
    return pre_of, pre_oh, pre_no, n_prev


def _carry_week_next_month_off_counts(
    carry_next: dict, n: int, monday_date: date, month_last: date,
) -> tuple[int, int, int, int]:
    """
    같은 월~일 주 중 month_last(당월 말일) 이후인 날(차월)에 대한 OF/OH/NO 개수와 그 일수 n_next.
    carry_next: {간호사: [차월 1일, 2일, …] 또는 {차월일: 시프트}} — _normalize_carry_in 과 동일 형태.
    """
    post_of = post_oh = post_no = 0
    n_next = 0
    next_first = month_last + timedelta(days=1)
    for i in range(7):
        d = monday_date + timedelta(days=i)
        if d <= month_last:
            continue
        n_next += 1
        if not carry_next:
            continue
        seq = carry_next.get(n)
        day_next = (d - next_first).days + 1
        s = None
        if isinstance(seq, dict):
            s = seq.get(day_next) or seq.get(str(day_next))
        elif isinstance(seq, (list, tuple)) and 1 <= day_next <= len(seq):
            s = seq[day_next - 1]
        if s == 'OF':
            post_of += 1
        elif s == 'OH':
            post_oh += 1
        elif s == 'NO':
            post_no += 1
    return post_of, post_oh, post_no, n_next


def _weekly_off_rule_met(
    of_vis: int,
    oh_vis: int,
    no_vis: int,
    n_prev: int,
    len_wdays: int,
    post_of: int,
    post_oh: int,
    post_no: int,
    n_next: int,
    carry_next_provided: bool,
) -> bool:
    """
    주(월~일) 2 OF 규칙 충족 여부.
    - 마지막주가 당월 안에서 토·일까지 끝나지 않으면(n_next>0) 차월 초와 한 주로 합산.
    - carry_next_month 가 없으면: 당월·전월 구간만으로 주 2 OF를 강제하지 않음(차월에서 맞출 수 있음).
    - 있으면: 전월+당월+차월 합산으로 weekly_of_equiv_satisfied.
    """
    of_t = of_vis + post_of
    oh_t = oh_vis + post_oh
    no_t = no_vis + post_no
    m = n_prev + len_wdays + n_next
    if m <= 0:
        return True
    if n_next > 0 and not carry_next_provided:
        return True
    return weekly_of_equiv_satisfied(of_t, oh_t, no_t, m)


def weekly_of_equiv_satisfied(of_c: int, oh_c: int, no_c: int, m: int) -> bool:
    """
    월~일 한 주 중 평가에 포함된 일수 m일에 대한 '주 2 OF' 충족 여부.
    (당월 일자 + 전월 동주 carry 일자를 합친 m일에 대해 동일 규칙 적용 가능)
    - of_c: OF 개수만 (연차는 주 2 휴무 인정에 포함되지 않음 — OF+연으로 대체 불가).
    - OF가 2개 이상이면 충족.
    - OH만 2개 이상이어도 충족.
    - 또는 OF+OH, OF+NO, OH+NO (각 1개 이상). NO는 주당 최대 1개(별도 검증).
    - m==1: OF/OH/NO 중 하나라도 있으면 충족.
    """
    if m <= 0:
        return True
    if m == 1:
        return of_c + oh_c + no_c >= 1
    return (
        of_c >= 2
        or oh_c >= 2
        or (of_c >= 1 and oh_c >= 1)
        or (of_c >= 1 and no_c >= 1)
        or (oh_c >= 1 and no_c >= 1)
    )


def _weekly_off_ok_after_of_to_yun(
    of_vis: int,
    oh_vis: int,
    no_vis: int,
    n_prev: int,
    wdays: list,
    post_of: int,
    post_oh: int,
    post_no: int,
    n_next: int,
    carry_next_provided: bool,
) -> bool:
    """
    OF 한 칸을 연으로 바꾼 뒤(OF만 -1, 연은 주 2휴무 인정에 불포함)에도
    weekly_of_equiv 가 성립하는지. 말주가 차월로 이어지고 carry_next 가 없을 때는
    당월에 보이는 일수만으로 엄격히 판정(기존 _weekly_off_rule_met 의 무조건 True 회피).
    """
    new_of = max(0, of_vis - 1)
    len_w = len(wdays)
    if n_next > 0 and not carry_next_provided:
        m = n_prev + len_w
        if m <= 0:
            return True
        return weekly_of_equiv_satisfied(new_of, oh_vis, no_vis, m)
    of_t = new_of + post_of
    oh_t = oh_vis + post_oh
    no_t = no_vis + post_no
    m = n_prev + len_w + n_next
    if m <= 0:
        return True
    return weekly_of_equiv_satisfied(of_t, oh_t, no_t, m)


def _weekly_off_strict_satisfied_for_week(
    sched, n: int, wdays: list, carry, carry_next, mon_date: date,
    month_first: date, month_last: date, carry_next_provided: bool,
) -> bool:
    """연차 제외 OF/OH/NO만으로 주 2휴무 충족 여부(차월 미입력 시 당월 구간만 엄격 판정)."""
    pre_of, pre_oh, pre_no, n_prev = _carry_week_prev_month_off_counts(
        carry, n, mon_date, month_first,
    )
    post_of, post_oh, post_no, n_next = _carry_week_next_month_off_counts(
        carry_next, n, mon_date, month_last,
    )
    of_vis = pre_of + sum(1 for d2 in wdays if sched[n].get(d2) == 'OF')
    oh_vis = pre_oh + sum(1 for d2 in wdays if sched[n].get(d2) == 'OH')
    no_vis = pre_no + sum(1 for d2 in wdays if sched[n].get(d2) == 'NO')
    len_w = len(wdays)
    if n_next > 0 and not carry_next_provided:
        m = n_prev + len_w
        if m <= 0:
            return True
        return weekly_of_equiv_satisfied(of_vis, oh_vis, no_vis, m)
    of_t = of_vis + post_of
    oh_t = oh_vis + post_oh
    no_t = no_vis + post_no
    m = n_prev + len_w + n_next
    if m <= 0:
        return True
    return weekly_of_equiv_satisfied(of_t, oh_t, no_t, m)


def _repair_yun_to_of_for_weekly_rule(
    sched, nurses: list, week_days_map: dict, carry, carry_next,
    month_first: date, month_last: date, carry_next_provided: bool, req_locked: set,
):
    """주간 휴무(OF/OH/NO 조합) 미달이면 해당 주의 연→OF (연은 주 2휴무에 불포함)."""
    for _ in range(len(nurses) * len(week_days_map) + 24):
        changed = False
        for n in nurses:
            for _wkey, _wdays in week_days_map.items():
                if not _wdays:
                    continue
                mon_date = date.fromordinal(_wkey)
                if _weekly_off_strict_satisfied_for_week(
                    sched, n, _wdays, carry, carry_next, mon_date,
                    month_first, month_last, carry_next_provided,
                ):
                    continue
                for _d in sorted(_wdays, reverse=True):
                    if sched[n].get(_d) != '연':
                        continue
                    if (n, _d) in req_locked:
                        continue
                    sched[n][_d] = 'OF'
                    changed = True
                    break
        if not changed:
            break


def _request_locked_cells(requests, num_nurses: int) -> set:
    """신청 표에 날짜 키가 있는 (간호사 인덱스, 일) — 생성·후처리·스왑에서 덮어쓰지 않음."""
    locked = set()
    for ni, ds in (requests or {}).items():
        try:
            ni = int(ni)
        except (TypeError, ValueError):
            continue
        if not (0 <= ni < num_nurses):
            continue
        for dn in (ds or {}).keys():
            try:
                locked.add((ni, int(dn)))
            except (TypeError, ValueError):
                continue
    return locked


def _reapply_requests_to_schedule(sched, requests, num_nurses: int) -> None:
    """생성 마지막에 신청 셀을 다시 한 번 씌워 절대 불일치를 제거한다.
    N은 간호사당 N_ABS_MAX를 넘기지 않도록, 새로 N을 늘리는 경우에만 상한을 검사한다."""
    for ni, ds in (requests or {}).items():
        try:
            ni = int(ni)
        except (TypeError, ValueError):
            continue
        if not (0 <= ni < num_nurses):
            continue
        for dn, shift in (ds or {}).items():
            try:
                dni = int(dn)
            except (TypeError, ValueError):
                continue
            if 1 <= dni <= NUM_DAYS:
                old = sched.get(ni, {}).get(dni)
                if shift == 'N' and old != 'N':
                    cur_n = sum(
                        1 for d, s in sched.get(ni, {}).items() if s == 'N'
                    )
                    if cur_n >= N_ABS_MAX:
                        continue
                sched.setdefault(ni, {})[dni] = shift


def _post_reapply_fix_n_cap_and_daily_two(
    sched, num_nurses, holidays, requests, forbidden_pairs, carry_in, shift_bans,
):
    """
    재적용 후 N>7 또는 일일 N<2가 되면 조정한다.
    - 초과 N: 잠기지 않은 날부터 다른 간호사에게 N을 넘기거나 OF로 내린 뒤 긴급 보충.
    """
    nurses = list(range(1, num_nurses))
    carry = _normalize_carry_in(carry_in, num_nurses)
    fp_map = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    den_bans = _normalize_shift_bans(shift_bans, num_nurses)
    locked = _request_locked_cells(requests, num_nurses)
    days = get_april_days(holidays)
    _dn_holiday = {d['day']: bool(d['is_holiday']) for d in days}

    def den_banned(n, sh):
        if sh not in ('D', 'E', 'N'):
            return False
        b = den_bans.get(n)
        return bool(b and sh in b)

    def sk(n, dn, k):
        return _shift_k_days_before(sched, carry, n, dn, k)

    def work_streak_before(n, d):
        count = 0
        for back in range(1, d + 1):
            s = sched[n].get(d - back + 1)
            if s not in STREAK_WORK_SHIFTS:
                break
            count += 1
        if count < d:
            return count
        c = carry.get(n) or ()
        for s in reversed(c):
            if s not in STREAK_WORK_SHIFTS:
                break
            count += 1
        return count

    def fp_same_shift_conflict(n, dn, shift):
        if not fp_map or shift not in ('D', 'E', 'N'):
            return False
        for m in range(num_nurses):
            if m == n:
                continue
            if sched[m].get(dn) != shift:
                continue
            key = (min(n, m), max(n, m))
            allowed = fp_map.get(key)
            if allowed and shift in allowed:
                return True
        return False

    def n_count(n):
        return sum(1 for v in sched.get(n, {}).values() if v == 'N')

    # ── ① N>7: 다른 간호사에게 이전(같은 날 2명 N 유지) 또는 OF
    for _ in range(NUM_DAYS * len(nurses) + 8):
        over = [n for n in nurses if n_count(n) > N_ABS_MAX]
        if not over:
            break
        n = over[0]
        n_days = sorted(
            (d for d in range(1, NUM_DAYS + 1) if sched.get(n, {}).get(d) == 'N'),
            reverse=True,
        )
        moved = False
        for dn in n_days:
            if (n, dn) in locked:
                continue
            for p in nurses:
                if p == n or sched[p].get(dn) == 'N':
                    continue
                if n_count(p) >= N_ABS_MAX:
                    continue
                if den_banned(p, 'N'):
                    continue
                if sk(p, dn, 1) == 'N':
                    continue
                cur = sched[p].get(dn)
                if cur not in (None, 'NO', 'OF', 'OH'):
                    continue
                if (p, dn) in locked:
                    continue
                d0 = dn - 1
                if work_streak_before(p, d0) + 1 > 5:
                    continue
                if fp_same_shift_conflict(p, dn, 'N'):
                    continue
                need = 'OH' if _dn_holiday.get(dn) else 'OF'
                sched[n][dn] = need
                sched[p][dn] = 'N'
                moved = True
                break
            if moved:
                break
        if not moved:
            for dn in n_days:
                if (n, dn) in locked:
                    continue
                need = 'OH' if _dn_holiday.get(dn) else 'OF'
                sched[n][dn] = need
                break
            else:
                break

    # ── ② 일일 N<2 긴급 보충 (상한·금지쌍·전일 N 제외)
    for d in range(NUM_DAYS):
        dn = d + 1
        on_n = [n for n in nurses if sched[n].get(dn) == 'N']
        needed = 2 - len(on_n)
        if needed <= 0:
            continue
        for _ in range(needed * 4):
            on_n = [n for n in nurses if sched[n].get(dn) == 'N']
            if len(on_n) >= 2:
                break
            cands = [
                x for x in nurses
                if x not in on_n
                and not den_banned(x, 'N')
                and sk(x, dn, 1) != 'N'
                and sched[x].get(dn) in (None, 'NO', 'OF', 'OH')
                and (x, dn) not in locked
                and work_streak_before(x, d) + 1 <= 5
                and not fp_same_shift_conflict(x, dn, 'N')
                and n_count(x) < N_ABS_MAX
            ]
            if not cands:
                break
            cands.sort(key=lambda x: (n_count(x), x))
            x = cands[0]
            sched[x][dn] = 'N'


def d_slots_per_day(num_nurses: int, day: dict, head_is_a1: bool) -> int:
    """
    해당 날짜의 D(데이) 배치 상한 인원.
    - 주말/공휴일: 반드시 2명 (절대 규칙, 상한=하한)
    - 평일: 수간 A1 아님 → 2명 / A1·11명 이상 → 3명 / A1·10명 이하 → 2명 (평일 2~3명 운영)
    """
    if day['is_weekend'] or day['is_holiday']:
        return 2
    if not head_is_a1:
        return 2
    if num_nurses >= 11:
        return 3
    return 2


def d_assignment_target(num_nurses: int, day: dict, head_is_a1: bool) -> int:
    """
    해당 날 D 배정 목표 인원(스케줄러·후처리 공통).
    평일·수간 포함 11명 이상: 수간호사 A1 여부와 무관 최소 2명 → max(규칙상 목표, 2).
    """
    t = d_slots_per_day(num_nurses, day, head_is_a1)
    if day['is_weekend'] or day['is_holiday']:
        return t
    if num_nurses >= 11:
        return max(t, 2)
    return t


def _validation_minimum_d(num_nurses: int, day: dict, head_is_a1: bool) -> int:
    """validate_schedule ① 일일 최소 D 인원(주말·공휴 2 절대, 평일은 2명 지향·소인원만 예외)."""
    if day['is_weekend'] or day['is_holiday']:
        return 2
    if num_nurses >= 11:
        return 2
    if not head_is_a1:
        return 2
    # 수간 A1·간호사 적을 때만 평일 D 최소 1 허용
    return 1 if num_nurses <= 7 else 2


def _ensure_validation_d_floor(
    sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
    tie_rng=None, carry_next_month=None,
):
    """
    일일 D가 검증 최소 미달이면 OF(또는 주말·공휴 규칙상 가능한 칸)을 D로 보충.
    말주가 차월로 이어질 때는 동일 주에 OF·OH·NO가 이미 넉넉한 간호사부터 선택해
    '다음 달 첫 주에 OF 2회 맞추기' 여유를 남기도록 한다(carry_next_month 없으면 주 경계 규칙 완화).
    """
    nurses = list(range(1, num_nurses))
    carry = _normalize_carry_in(carry_in, num_nurses)
    carry_next = _normalize_carry_in(carry_next_month, num_nurses) if carry_next_month else {}
    carry_next_provided = carry_next_month is not None
    fp_map = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    den_bans = _normalize_shift_bans(shift_bans, num_nurses)
    locked = _request_locked_cells(requests, num_nurses)
    days = get_april_days(holidays)
    month_first = date(YEAR, MONTH, 1)
    month_last = date(YEAR, MONTH, NUM_DAYS)
    OFF_SET = frozenset({'OF', 'OH', 'NO'})

    def den_banned(n, sh):
        if sh not in ('D', 'E', 'N'):
            return False
        b = den_bans.get(n)
        return bool(b and sh in b)

    def sk(n, dn, k):
        return _shift_k_days_before(sched, carry, n, dn, k)

    def work_streak_before(n, d):
        count = 0
        for back in range(1, d + 1):
            s = sched[n].get(d - back + 1)
            if s not in STREAK_WORK_SHIFTS:
                break
            count += 1
        if count < d:
            return count
        c = carry.get(n) or ()
        for s in reversed(c):
            if s not in STREAK_WORK_SHIFTS:
                break
            count += 1
        return count

    def streak_total(n, d):
        before = work_streak_before(n, d)
        after = 0
        for fwd in range(1, NUM_DAYS - d):
            s = sched[n].get(d + fwd + 1)
            if s not in STREAK_WORK_SHIFTS:
                break
            after += 1
        return before + 1 + after

    def fp_same_shift_conflict(n, dn, shift):
        if not fp_map or shift not in ('D', 'E', 'N'):
            return False
        for m in range(num_nurses):
            if m == n:
                continue
            if sched[m].get(dn) != shift:
                continue
            key = (min(n, m), max(n, m))
            allowed = fp_map.get(key)
            if allowed and shift in allowed:
                return True
        return False

    week_map: dict[int, int] = {}
    week_days_map: dict[int, list] = {}
    for day in days:
        dt = day['date']
        mon = dt - timedelta(days=dt.weekday())
        key = mon.toordinal()
        week_map[day['day']] = key
        week_days_map.setdefault(key, []).append(day['day'])

    def wk_rest_count(n, wkey):
        return sum(
            1 for d2 in week_days_map[wkey]
            if sched[n].get(d2) in ('OF', 'OH', 'NO', '연')
        )

    def _tie(n):
        return tie_rng.random() if tie_rng else 0

    D_ABS_MAX = 6
    for _ in range(NUM_DAYS * len(nurses) + 24):
        d_cnt = {n: sum(1 for v in sched[n].values() if v == 'D') for n in nurses}
        improved = False
        for d, day in enumerate(days):
            dn = d + 1
            head_a1 = sched[0].get(dn) == 'A1'
            req_d = _validation_minimum_d(num_nurses, day, head_a1)
            d_on = sum(1 for n in nurses if sched[n].get(dn) == 'D')
            need = req_d - d_on
            if need <= 0:
                continue

            def can_take_d(n):
                if (n, dn) in locked:
                    return False
                cur = sched[n].get(dn)
                if day['is_holiday']:
                    if cur not in ('OH', '연'):
                        return False
                else:
                    if cur not in ('OF', '연'):
                        return False
                if den_banned(n, 'D'):
                    return False
                if sk(n, dn, 1) == 'E':
                    return False
                if sk(n, dn, 1) == 'N':
                    return False
                if sk(n, dn, 2) == 'N' and sk(n, dn, 1) in OFF_SET:
                    return False
                if streak_total(n, d) > 5:
                    return False
                if fp_same_shift_conflict(n, dn, 'D'):
                    return False
                if d_cnt.get(n, 0) >= D_ABS_MAX:
                    return False
                return True

            cands = [n for n in nurses if can_take_d(n)]
            if not cands:
                continue
            wkey = week_map.get(dn)
            # 말주·차월 이어짐: 주간 휴무가 이미 많은 사람·차월 OF 여유 우선
            if wkey is not None:
                mon_date = date.fromordinal(wkey)
                _, _, _, n_next = _carry_week_next_month_off_counts(
                    carry_next, cands[0], mon_date, month_last,
                )
                has_tail = n_next > 0

                def sort_key(n):
                    pre_of, pre_oh, pre_no, n_prev = _carry_week_prev_month_off_counts(
                        carry, n, mon_date, month_first,
                    )
                    post_of, post_oh, post_no, _npost = _carry_week_next_month_off_counts(
                        carry_next, n, mon_date, month_last,
                    )
                    of_w = sum(1 for d2 in week_days_map[wkey] if sched[n].get(d2) == 'OF')
                    oh_w = sum(1 for d2 in week_days_map[wkey] if sched[n].get(d2) == 'OH')
                    no_w = sum(1 for d2 in week_days_map[wkey] if sched[n].get(d2) == 'NO')
                    of_vis = pre_of + of_w
                    oh_vis = pre_oh + oh_w
                    no_vis = pre_no + no_w
                    len_w = len(week_days_map[wkey])
                    ok_before = _weekly_off_rule_met(
                        of_vis, oh_vis, no_vis, n_prev, len_w,
                        post_of, post_oh, post_no, _npost, carry_next_provided,
                    )
                    if not ok_before:
                        return (999, 999, d_cnt.get(n, 0), _tie(n), n)
                    of_vis2 = of_vis - (1 if sched[n].get(dn) == 'OF' else 0)
                    oh_vis2 = oh_vis - (1 if sched[n].get(dn) == 'OH' else 0)
                    # 연→D: OF/OH 개수 불변(연은 주 2휴무 인정에 미포함)
                    ok_after = _weekly_off_rule_met(
                        of_vis2, oh_vis2, no_vis, n_prev, len_w,
                        post_of, post_oh, post_no, _npost, carry_next_provided,
                    )
                    slack = wk_rest_count(n, wkey)
                    tail_bonus = (0 if has_tail and carry_next_provided else 1)
                    return (
                        0 if ok_after else 1,
                        tail_bonus,
                        -slack,
                        -post_of,
                        d_cnt.get(n, 0),
                        _tie(n),
                        n,
                    )
            else:
                def sort_key(n):
                    return (d_cnt.get(n, 0), _tie(n), n)

            cands.sort(key=sort_key)
            picked = cands[0]
            sched[picked][dn] = 'D'
            improved = True
            break
        if not improved:
            break


def _cap_auto_rest_in_tail_last_week(
    sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
):
    """
    당월 말일이 들어간 주(월~일) 중 **당월에만 놓인 날이 4일 이하**인 말주(예: 27~30일 4칸):
    신청으로 잠기지 않은 쉼(OF/OH/NO/연)을 **간호사당 최대 1일**만 두고,
    나머지 쉼 칸은 D로 바꾼다. 남길 하루는 해당 주·당월 구간에서 **가장 늦은 날** 우선.
    (말주 4일 근무 구간에서 자동 배정 휴무가 한 사람에게 몰리지 않게 함)
    """
    nurses = list(range(1, num_nurses))
    carry = _normalize_carry_in(carry_in, num_nurses)
    fp_map = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    den_bans = _normalize_shift_bans(shift_bans, num_nurses)
    locked = _request_locked_cells(requests, num_nurses)
    days = get_april_days(holidays)
    OFF_SET = frozenset({'OF', 'OH', 'NO'})
    REST_ALL = frozenset({'OF', 'OH', 'NO', '연'})

    def den_banned(n, sh):
        if sh not in ('D', 'E', 'N'):
            return False
        b = den_bans.get(n)
        return bool(b and sh in b)

    def sk(n, dn, k):
        return _shift_k_days_before(sched, carry, n, dn, k)

    def work_streak_before(n, d):
        count = 0
        for back in range(1, d + 1):
            s = sched[n].get(d - back + 1)
            if s not in STREAK_WORK_SHIFTS:
                break
            count += 1
        if count < d:
            return count
        c = carry.get(n) or ()
        for s in reversed(c):
            if s not in STREAK_WORK_SHIFTS:
                break
            count += 1
        return count

    def streak_total(n, d):
        before = work_streak_before(n, d)
        after = 0
        for fwd in range(1, NUM_DAYS - d):
            s = sched[n].get(d + fwd + 1)
            if s not in STREAK_WORK_SHIFTS:
                break
            after += 1
        return before + 1 + after

    def fp_same_shift_conflict(n, dn, shift):
        if not fp_map or shift not in ('D', 'E', 'N'):
            return False
        for m in range(num_nurses):
            if m == n:
                continue
            if sched[m].get(dn) != shift:
                continue
            key = (min(n, m), max(n, m))
            allowed = fp_map.get(key)
            if allowed and shift in allowed:
                return True
        return False

    week_days_map: dict[int, list] = {}
    for day in days:
        dt = day['date']
        mon = dt - timedelta(days=dt.weekday())
        key = mon.toordinal()
        week_days_map.setdefault(key, []).append(day['day'])

    for _wkey, wdays in week_days_map.items():
        if not wdays or NUM_DAYS not in wdays:
            continue
        if len(wdays) > 4:
            continue
        for n in nurses:
            rests = [
                dn for dn in sorted(wdays)
                if sched[n].get(dn) in REST_ALL and (n, dn) not in locked
            ]
            if len(rests) <= 1:
                continue
            keep_dn = max(rests)
            for dn in rests:
                if dn == keep_dn:
                    continue
                d_idx = dn - 1
                day = days[d_idx]
                if sched[n].get(dn) == 'OH' and day['is_holiday']:
                    continue
                if den_banned(n, 'D'):
                    continue
                if sk(n, dn, 1) == 'E':
                    continue
                if sk(n, dn, 1) == 'N':
                    continue
                if sk(n, dn, 2) == 'N' and sk(n, dn, 1) in OFF_SET:
                    continue
                if streak_total(n, d_idx) > 5:
                    continue
                if fp_same_shift_conflict(n, dn, 'D'):
                    continue
                sched[n][dn] = 'D'


def _convert_non_request_yun_to_d(
    sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
):
    """
    신청으로 고정되지 않은 연차 칸은 일일 D 상한·개인 D 상한(6)·연속근무 5일 등을 만족하면 D로 전환.
    (연→D는 해당 날 E·N 인원을 바꾸지 않으므로 일일 E 2·N 2 절대 규칙과 충돌하지 않음.)
    """
    nurses = list(range(1, num_nurses))
    carry = _normalize_carry_in(carry_in, num_nurses)
    fp_map = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    den_bans = _normalize_shift_bans(shift_bans, num_nurses)
    locked = _request_locked_cells(requests, num_nurses)
    days = get_april_days(holidays)
    OFF_SET = frozenset({'OF', 'OH', 'NO'})
    D_ABS_MAX = 6

    def den_banned(n, sh):
        if sh not in ('D', 'E', 'N'):
            return False
        b = den_bans.get(n)
        return bool(b and sh in b)

    def sk(n, dn, k):
        return _shift_k_days_before(sched, carry, n, dn, k)

    def work_streak_before(n, d):
        count = 0
        for back in range(1, d + 1):
            s = sched[n].get(d - back + 1)
            if s not in STREAK_WORK_SHIFTS:
                break
            count += 1
        if count < d:
            return count
        c = carry.get(n) or ()
        for s in reversed(c):
            if s not in STREAK_WORK_SHIFTS:
                break
            count += 1
        return count

    def streak_total(n, d):
        before = work_streak_before(n, d)
        after = 0
        for fwd in range(1, NUM_DAYS - d):
            s = sched[n].get(d + fwd + 1)
            if s not in STREAK_WORK_SHIFTS:
                break
            after += 1
        return before + 1 + after

    def fp_same_shift_conflict(n, dn, shift):
        if not fp_map or shift not in ('D', 'E', 'N'):
            return False
        for m in range(num_nurses):
            if m == n:
                continue
            if sched[m].get(dn) != shift:
                continue
            key = (min(n, m), max(n, m))
            allowed = fp_map.get(key)
            if allowed and shift in allowed:
                return True
        return False

    def can_yun_to_d(n, dn, d):
        if sched[n].get(dn) != '연':
            return False
        if (n, dn) in locked:
            return False
        if den_banned(n, 'D'):
            return False
        if sk(n, dn, 1) == 'E':
            return False
        if sk(n, dn, 1) == 'N':
            return False
        if sk(n, dn, 2) == 'N' and sk(n, dn, 1) in OFF_SET:
            return False
        if streak_total(n, d) > 5:
            return False
        if fp_same_shift_conflict(n, dn, 'D'):
            return False
        if sum(1 for v in sched[n].values() if v == 'D') >= D_ABS_MAX:
            return False
        return True

    for _ in range(NUM_DAYS * len(nurses) * 2 + 48):
        changed = False
        for d, day in enumerate(days):
            dn = d + 1
            head_a1 = sched[0].get(dn) == 'A1'
            d_max = d_slots_per_day(num_nurses, day, head_a1)
            d_on = sum(1 for m in nurses if sched[m].get(dn) == 'D')
            if d_on >= d_max:
                continue
            for n in sorted(nurses):
                if not can_yun_to_d(n, dn, d):
                    continue
                sched[n][dn] = 'D'
                changed = True
                break
            if changed:
                break
        if not changed:
            break


def _auto_yun_to_of_if_quota_room(
    sched, num_nurses, holidays, carry_in, requests,
):
    """
    월 OFF(OF/OH/NO) 쿼터에 여유가 있으면 자동 연차 칸을 OF/OH로 바꿔 연 표기를 줄인다.
    (주말·공휴 D=2 등 다른 규칙은 시프트 종류만 OF/OH로 맞춤)
    """
    nurses = list(range(1, num_nurses))
    carry = _normalize_carry_in(carry_in, num_nurses)
    locked = _request_locked_cells(requests, num_nurses)
    days = get_april_days(holidays)
    of_quota_month = sum(1 for day in days if day['is_weekend'] or day['is_holiday'])

    def sk(n, dn, k):
        return _shift_k_days_before(sched, carry, n, dn, k)

    for _ in range(NUM_DAYS * len(nurses) + 16):
        changed = False
        for n in nurses:
            off_ct = sum(1 for s in sched[n].values() if s in ('OF', 'OH', 'NO'))
            for d, day in enumerate(days):
                dn = d + 1
                if sched[n].get(dn) != '연':
                    continue
                if (n, dn) in locked:
                    continue
                if off_ct >= of_quota_month:
                    break
                # N 다음날·그 외 휴무: 공휴 OH / 평일 OF
                need = 'OH' if day['is_holiday'] else 'OF'
                sched[n][dn] = need
                off_ct += 1
                changed = True
                break
            if changed:
                break
        if not changed:
            break


# ── 스케줄 생성 (순수 Python 그리디) ─────────────────────────────────────────
def solve_schedule(num_nurses, requests, holidays=(), forbidden_pairs=None, carry_in=None,
                   regenerate=False, rng_seed=None, nurse_names=None, carry_next_month=None,
                   shift_bans=None):
    """
    서버 충돌 없는 순수 Python 그리디 스케줄러
    num_nurses : 총 간호사 수 (0번=수간호사, 1..n-1=일반간호사)
    requests   : {nurse_idx: {day_num: shift_name}}
    holidays   : 공휴일 날짜 목록 (1-based)
    forbidden_pairs : (선택) 같은 날 동시 배치 금지
                      [(i,j), ...] 또는 [(i,j,['D','E']), ...] — 적용 시프트만 검사
                      수간호사 인덱스 0 포함. UI에서 2~4명 그룹은 쌍으로 전개되어 전달됨.
    carry_in   : (선택) 전월 말 근무 꼬리 {간호사인덱스: [시프트,...]} 오래된 날→최근 날
    carry_next_month: (선택) 차월 초 근무(마지막주=당월 말~차월 일요일 한 주의 주 2 OF 합산용)
    shift_bans: (선택) dict[int,str] — 간호사 인덱스 → d_only | no_d | no_e | no_n (근무불가)
    regenerate : True면 신청(requests) 셀은 유지한 채 나머지만 다른 타이브레이크·미세조정
    rng_seed   : 재생성 시 그리디·스왑 난수 시드 (None이면 비고정 Random)
    nurse_names: 재생성 미세조정 시 validate_schedule 표시용 (없으면 기본 이름)

    N(야간) 절대 규칙 — 총원 11명 이상이어도 동일: 매일 정확히 2명, 간호사당 월 최대 N_ABS_MAX(7)개,
    월 전체 N슬롯(2×말일)은 일반간호사에게 공평하게 목표 분배(_compute_n_targets_fair).
    """
    try:
        tie_rng = None
        if regenerate:
            tie_rng = random.Random(rng_seed) if rng_seed is not None else random.Random()
        return _greedy_schedule(
            num_nurses, requests, holidays, forbidden_pairs, carry_in,
            tie_rng=tie_rng, nurse_names=nurse_names,
            regenerate=regenerate,
            carry_next_month=carry_next_month,
            shift_bans=shift_bans,
        )
    except Exception as e:
        print(f'[오류] {e}')
        return None, False, str(e)


def _normalize_forbidden_pairs(forbidden_pairs, num_nurses):
    """
    forbidden_pairs: [(i,j), ...] 또는 [(i,j, ['D','E']), ...] 등
    반환: dict (min,max) -> frozenset({'D','E','N'}의 부분집합) — 수간호사(0) 포함.
    """
    out: dict[tuple[int, int], frozenset] = {}
    if not forbidden_pairs:
        return out
    for pair in forbidden_pairs:
        if not pair:
            continue
        shifts = frozenset({'D', 'E', 'N'})
        if len(pair) == 2:
            try:
                a, b = int(pair[0]), int(pair[1])
            except (TypeError, ValueError):
                continue
        elif len(pair) >= 3:
            try:
                a, b = int(pair[0]), int(pair[1])
            except (TypeError, ValueError):
                continue
            raw = pair[2]
            if isinstance(raw, (list, tuple, set, frozenset)):
                shifts = frozenset(x for x in raw if x in ('D', 'E', 'N'))
            elif isinstance(raw, str):
                shifts = frozenset(
                    x.strip() for x in raw.replace(',', ' ').split() if x.strip() in ('D', 'E', 'N')
                )
            if not shifts:
                shifts = frozenset({'D', 'E', 'N'})
        else:
            continue
        if not (0 <= a < num_nurses and 0 <= b < num_nurses and a != b):
            continue
        key = (min(a, b), max(a, b))
        out[key] = out.get(key, frozenset()) | shifts
    return out


def _normalize_shift_bans(shift_bans, num_nurses):
    """
    shift_bans: dict[int, str] — 간호사 인덱스 → 'd_only' | 'no_d' | 'no_e' | 'no_n'
    d_only: D/E/N 중 E·N 자동 배제(데이만 가능).
    반환: dict[int, frozenset] — 간호사별 금지 D/E/N (수간호사 0은 무시).
    """
    mode_to_banned = {
        'd_only': frozenset({'E', 'N'}),
        'no_d': frozenset({'D'}),
        'no_e': frozenset({'E'}),
        'no_n': frozenset({'N'}),
    }
    out: dict[int, frozenset] = {}
    if not shift_bans:
        return out
    for k, mode in shift_bans.items():
        try:
            ni = int(k)
        except (TypeError, ValueError):
            continue
        if not (0 <= ni < num_nurses):
            continue
        m = str(mode).strip() if mode is not None else ''
        if m not in mode_to_banned:
            continue
        out[ni] = mode_to_banned[m]
    return out


# 전월 말 근무 꼬리 (월 경계 규칙용) — 최대 14일, 오래된 날 → 최근 날 순
CARRY_MAX_DAYS = 14

# 간호사당 월 N(야간) 상한 — 수간 포함 총원이 11명 이상이어도 동일(7개까지).
# 일일 N 2명·목표 합=2×말일·공평 분배는 인원 수와 무관하게 동일 규칙(_greedy_schedule·validate_schedule).
N_ABS_MAX = 7


def _compute_n_targets_fair(num_reg: int, total_slots: int, n_max: int = N_ABS_MAX) -> list:
    """
    일반간호사 num_reg명에게 total_slots개의 N 배정 목표(합=total_slots, 각 n_max 이하).
    먼저 균등 분배 후 min(·, n_max) 적용, 부족분은 가장 적게 받은 사람부터 +1(상한 내).
    인원이 늘어도(예: 수간 포함 11명 이상) N 상한·공평 목표 분배 방식은 동일하며, 변하는 것은 num_reg·합계 슬롯뿐이다.
    (예: 8명·60슬롯은 7×8=56으로 상한만으로는 60을 못 채움 — 물리적 한계)
    """
    if num_reg <= 0:
        return []
    if total_slots <= 0:
        return [0] * num_reg
    base = total_slots // num_reg
    rem = total_slots % num_reg
    targets = [min(base + (1 if i < rem else 0), n_max) for i in range(num_reg)]
    s = sum(targets)
    if s < total_slots:
        deficit = total_slots - s
        while deficit > 0:
            best_i = -1
            best_val = None
            for i in range(num_reg):
                if targets[i] >= n_max:
                    continue
                if best_i < 0 or targets[i] < best_val:
                    best_i, best_val = i, targets[i]
            if best_i < 0:
                break
            targets[best_i] += 1
            deficit -= 1
    elif s > total_slots:
        surplus = s - total_slots
        while surplus > 0:
            worst_i = -1
            worst_val = None
            for i in range(num_reg):
                if targets[i] <= 0:
                    continue
                if worst_i < 0 or targets[i] > worst_val:
                    worst_i, worst_val = i, targets[i]
            if worst_i < 0:
                break
            targets[worst_i] -= 1
            surplus -= 1
    return targets


def _n_block_plan_for_target(t: int, pat_idx: int) -> list:
    """
    N 목표 개수별 블록(연속 N 상한) 플랜.
    - 7개: 우선 2-2-3 / 2-3-2 / 3-2-2 순환, 3-3-1은 말일 1개 전제(검증·고립제거 예외와 맞춤).
    - 6개: 3-3, 2-2-2, 2-3-1, 3-2-1 순환 — 끝의 1은 말일만.
    말일은 set_period 기준 NUM_DAYS(28~31). 31일로 끝나는 달은 31일이 말일이며 그날 단독 N 1개가 허용된다.
    """
    _seven = [[2, 2, 3], [2, 3, 2], [3, 2, 2]]
    _six = [[3, 3], [2, 2, 2], [2, 3, 1], [3, 2, 1]]
    if t >= 7:
        if pat_idx % 4 == 3:
            return [3, 3, 1]
        return _seven[pat_idx % 3]
    if t == 6:
        return _six[pat_idx % 4]
    if t == 5:
        return [2, 3]
    if t == 4:
        return [2, 2]
    if t == 3:
        return [3]
    if t == 2:
        return [2]
    if t == 1:
        return [1]
    return [3]


def _normalize_carry_in(carry_in, num_nurses):
    """
    carry_in: { 간호사인덱스: [시프트, ...], ... } 또는 None
    값은 전월 말에서 이어지는 날들(오래된 것부터).
    """
    if not carry_in:
        return {}
    out = {}
    if not isinstance(carry_in, dict):
        return {}
    for k, v in carry_in.items():
        try:
            ni = int(k)
        except (TypeError, ValueError):
            continue
        if not (0 <= ni < num_nurses):
            continue
        if isinstance(v, (list, tuple)):
            seq = list(v)
        elif isinstance(v, str):
            seq = [x.strip() for x in v.replace(',', ' ').split() if x.strip()]
        else:
            continue
        if not seq:
            continue
        out[ni] = tuple(seq[-CARRY_MAX_DAYS:])
    return out


def _shift_k_days_before(sched_map, carry, n, dn, k):
    """
    이번 달 dn일의 k일 전 근무.
    carry[n] = 전월 말쪽 날들(오래된→최근), 마지막 원소가 전월 마지막 날(= 이번 달 1일 직전).
    """
    if k <= 0:
        return None
    td = dn - k
    if td >= 1:
        return sched_map.get(n, {}).get(td)
    c = carry.get(n) or ()
    idx = -td
    if td <= 0 and idx >= 1 and idx <= len(c):
        return c[-idx]
    return None


def _tie_break_map(items, tie_rng):
    """정렬 비교 시 random()을 매번 부르면 순서가 비전이 되므로, 항목별 고정 난수."""
    if not tie_rng:
        return {x: 0 for x in items}
    return {x: tie_rng.random() for x in items}


def _row_pattern_penalty(seq):
    """개인 한 달 시퀀스(문자열 리스트) — 낮을수록 N/OF/휴가가 덜 뭉침."""
    pen = 0
    L = len(seq)
    LEAVE = {'연', '공', '병', '경'}
    for i in range(L - 3):
        a, b, c, d = seq[i], seq[i + 1], seq[i + 2], seq[i + 3]
        if a == b == c == 'N' and d in LEAVE:
            pen += 10
        if a == 'E' and b == c == 'OF' and d == 'N':
            pen += 6
        if a == b == c == 'OF' and d == 'N':
            pen += 4
    for i in range(L - 2):
        if seq[i] == seq[i + 1] == seq[i + 2] == 'OF':
            pen += 2
    return pen


def _rest_gap_local_penalty(seq: list) -> int:
    """쉬는날(OF/OH/NO/연)-단일근무-쉬는날 윈도 패널티 — 정련 스왑이 섬을 줄이도록 유도."""
    REST = frozenset({'OF', 'OH', 'NO', '연'})
    WORK = STREAK_WORK_SHIFTS
    p = 0
    for i in range(len(seq) - 2):
        a, b, c = seq[i], seq[i + 1], seq[i + 2]
        if a in REST and c in REST and b in WORK:
            p += 14
    return p


def _n_block_gap_penalty(sched: dict, n: int) -> int:
    """N 블록 사이 달력 간격이 7일 미만이면 가산 — 'N 블록 간격 부족' 경고 감소 유도."""
    ns = sched.get(n, {})
    n_days = sorted(d for d, s in ns.items() if s == 'N')
    if len(n_days) < 2:
        return 0
    blocks = []
    blk = [n_days[0]]
    for d in n_days[1:]:
        if d == blk[-1] + 1:
            blk.append(d)
        else:
            blocks.append(blk)
            blk = [d]
    blocks.append(blk)
    p = 0
    for i in range(len(blocks) - 1):
        gap = blocks[i + 1][0] - blocks[i][-1] - 1
        if gap < 7:
            p += (7 - gap) * 5
    return p


def _off_quota_soft_penalty(sched: dict, n: int, days: list) -> int:
    """수간 기준 OFF 쿼터 초과분 — 'OFF 초과' 경고와 동조."""
    of_quota = sum(1 for day in days if day['is_weekend'] or day['is_holiday'])
    off_total = sum(1 for v in sched.get(n, {}).values() if v in ('OF', 'OH', 'NO'))
    return max(0, off_total - of_quota) * 8


def _rest_gap_work_excess_penalty(sched: dict, n: int) -> int:
    """쉬는 날 사이 근무가 5일 초과 시 가산 — 검증 절대 오류(D/E/N/공/EDU만 합산)와 동조."""
    REST_GAP = frozenset({'OF', 'OH', 'NO', '연'})
    gap_anchors = sorted(d for d, s in sched.get(n, {}).items() if s in REST_GAP)
    prev_a = None
    p = 0
    for od in gap_anchors:
        if prev_a is not None:
            work_btw = sum(
                1 for d in range(prev_a + 1, od)
                if sched.get(n, {}).get(d) in STREAK_WORK_SHIFTS
            )
            if work_btw > 5:
                p += (work_btw - 5) * 6
        prev_a = od
    return p


def _weekend_rest_balance_penalty(sched: dict, nurses: list, days: list) -> int:
    """토·일·공휴일 쉼(OF/OH/NO/연) 횟수 균등 — 분산이 클수록 패널티."""
    REST = frozenset({'OF', 'OH', 'NO', '연'})
    counts = []
    for n in nurses:
        c = 0
        for d, day in enumerate(days):
            if not (day['is_weekend'] or day['is_holiday']):
                continue
            if sched.get(n, {}).get(d + 1) in REST:
                c += 1
        counts.append(c)
    if len(counts) < 2:
        return 0
    mu = sum(counts) / len(counts)
    return int(sum((x - mu) ** 2 for x in counts) * 4)


def _repair_schedule_validate_errors(
    sched, num_nurses, holidays, forbidden_pairs, carry_in, requests, nurses, tie_rng=None,
    carry_next_month=None, shift_bans=None,
):
    """
    validate_schedule 절대 오류를 줄이기 위해 같은 날 두 간호사 시프트 교환을 시도.
    N-D·E-D·N-OF-D·N-OF-EDU·N블록 직후 휴무(OH/OF)·전월말 N→1일 휴무 등 검증과 동일한 조건.
    """
    carry = _normalize_carry_in(carry_in, num_nurses)
    fp_map = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    days = get_april_days(holidays)
    _dn_holiday = {d['day']: bool(d['is_holiday']) for d in days}
    locked = _request_locked_cells(requests, num_nurses)
    OFF_OK = frozenset({'OF', 'OH', 'NO'})

    def vk(n, dn, k):
        return _shift_k_days_before(sched, carry, n, dn, k)

    def sh(n, dn):
        return sched.get(n, {}).get(dn, '')

    def fp_sc(n, dn, shift):
        if not fp_map or shift not in ('D', 'E', 'N'):
            return False
        for m in nurses:
            if m == n:
                continue
            if sched.get(m, {}).get(dn) != shift:
                continue
            key = (min(n, m), max(n, m))
            if key in fp_map and shift in fp_map[key]:
                return True
        return False

    def err_count():
        iss = validate_schedule(
            sched, num_nurses, holidays,
            forbidden_pairs=forbidden_pairs, carry_in=carry_in,
            requests=requests, carry_next_month=carry_next_month,
            shift_bans=shift_bans,
        )
        return sum(1 for x in iss if x.get('level') == 'error')

    def swap_if_better(n, m, dn):
        if (n, dn) in locked or (m, dn) in locked or n == m:
            return False
        s_n, s_m = sh(n, dn), sh(m, dn)
        if s_n == s_m:
            return False
        ec0 = err_count()
        sched[n][dn], sched[m][dn] = s_m, s_n
        if err_count() < ec0:
            return True
        sched[n][dn], sched[m][dn] = s_n, s_m
        return False

    for _ in range(100):
        if err_count() == 0:
            return
        improved = False
        day_order = list(range(1, NUM_DAYS + 1))
        nurse_order = list(nurses)
        if tie_rng:
            tie_rng.shuffle(day_order)
            tie_rng.shuffle(nurse_order)

        for dn in day_order:
            for n in nurse_order:
                if (n, dn) in locked or sh(n, dn) != 'D':
                    continue
                v1 = vk(n, dn, 1)
                if v1 not in ('N', 'E'):
                    continue
                for m in nurse_order:
                    if sh(m, dn) not in OFF_OK:
                        continue
                    if fp_sc(m, dn, 'D'):
                        continue
                    if swap_if_better(n, m, dn):
                        improved = True
                        break
                if improved:
                    break
            if improved:
                break
        if improved:
            continue

        for dn in day_order:
            for n in nurse_order:
                if (n, dn) in locked:
                    continue
                sc = sh(n, dn)
                if sc not in ('D', 'EDU'):
                    continue
                if vk(n, dn, 2) != 'N' or vk(n, dn, 1) not in ('OF', 'OH'):
                    continue
                for m in nurse_order:
                    if sh(m, dn) not in OFF_OK:
                        continue
                    if sc == 'D' and fp_sc(m, dn, 'D'):
                        continue
                    if swap_if_better(n, m, dn):
                        improved = True
                        break
                if improved:
                    break
            if improved:
                break
        if improved:
            continue

        for n in nurse_order:
            ns = sched.get(n, {})
            n_days = sorted(d for d, s in ns.items() if s == 'N')
            if not n_days:
                continue
            blocks = []
            blk = [n_days[0]]
            for d in n_days[1:]:
                if d == blk[-1] + 1:
                    blk.append(d)
                else:
                    blocks.append(blk)
                    blk = [d]
            blocks.append(blk)
            for blk in blocks:
                end = blk[-1]
                if end >= NUM_DAYS:
                    continue
                dn = end + 1
                need = 'OH' if _dn_holiday.get(dn) else 'OF'
                if (n, dn) in locked or sh(n, dn) == need:
                    continue
                for m in nurse_order:
                    if sh(m, dn) != need:
                        continue
                    if swap_if_better(n, m, dn):
                        improved = True
                        break
                if improved:
                    break
            if improved:
                break
        if improved:
            continue

        dn1 = 1
        for n in nurse_order:
            cseq = list(carry.get(n, ()))
            if not cseq or cseq[-1] != 'N' or sh(n, 1) == 'N':
                continue
            need0 = 'OH' if days[0]['is_holiday'] else 'OF'
            if sh(n, 1) == need0 or (n, dn1) in locked:
                continue
            for m in nurse_order:
                if sh(m, 1) != need0:
                    continue
                if swap_if_better(n, m, dn1):
                    improved = True
                    break
            if improved:
                break
        if improved:
            continue

        break


def _refinement_objective(
    sched, num_nurses, holidays, forbidden_pairs, nurse_names, carry_in, nurses,
    requests=None, carry_next_month=None, shift_bans=None,
):
    """
    재생성 정련용 점수 — (오류 수, 경고 수, 소프트 패널티) 사전순.
    소프트 패널티는 검증 경고·패턴과 연동해 스왑 방향을 유도한다.
    """
    names = nurse_names if nurse_names is not None else get_nurse_names(num_nurses)
    days = get_april_days(holidays)
    issues = validate_schedule(
        sched, num_nurses, holidays,
        forbidden_pairs=forbidden_pairs, nurse_names=names, carry_in=carry_in,
        requests=requests, carry_next_month=carry_next_month,
        shift_bans=shift_bans,
    )
    err_n = sum(1 for x in issues if x.get('level') == 'error')
    warn_n = sum(1 for x in issues if x.get('level') == 'warn')
    pen = 0
    req_l = _request_locked_cells(requests or {}, num_nurses)
    auto_yun = sum(
        1 for n in nurses for d in range(1, NUM_DAYS + 1)
        if sched[n].get(d) == '연' and (n, d) not in req_l
    )
    pen += auto_yun * 28
    for n in nurses:
        seq = [sched[n].get(d, '') for d in range(1, NUM_DAYS + 1)]
        pen += _row_pattern_penalty(seq) + _rest_gap_local_penalty(seq)
        pen += _n_block_gap_penalty(sched, n)
        pen += _off_quota_soft_penalty(sched, n, days)
        pen += _rest_gap_work_excess_penalty(sched, n)
    pen += _weekend_rest_balance_penalty(sched, nurses, days)
    return err_n, warn_n, pen


def _refine_pick_swap_days(tie_rng: random.Random) -> tuple[int, int]:
    """근처 날짜 쌍을 자주 뽑아 N·휴무·섬 패턴을 바꾸기 쉽게 함."""
    if tie_rng.random() < 0.42:
        c = tie_rng.randint(1, NUM_DAYS)
        span = tie_rng.randint(2, 8)
        lo, hi = max(1, c - span), min(NUM_DAYS, c + span)
        pool = list(range(lo, hi + 1))
        if len(pool) >= 2:
            return tie_rng.sample(pool, 2)
    return tie_rng.sample(range(1, NUM_DAYS + 1), 2)


def _refine_schedule_regenerate(
    sched, requests, num_nurses, holidays, forbidden_pairs, carry_in, nurse_names, tie_rng,
    max_tries: int = 180, carry_next_month=None, shift_bans=None,
):
    """
    재생성 전용: 잠기지 않은 칸만 조정.
    - 같은 간호사 두 날짜 스왑: 일별 D/E/N 인원 수는 변할 수 있음(검증으로 필터).
    - 같은 날 두 간호사 스왑: 그 날 시프트 다중집합 불변 → 일일 인력 오류를 늘리지 않음(함께 근무 불가 등은 검증).
    - 목표 (오류, 경고, 소프트패널티) 사전순 감소. 오류 개수는 절대 늘리지 않음.
    """
    if not tie_rng:
        return
    nurses = list(range(1, num_nurses))
    locked = _request_locked_cells(requests, num_nurses)

    score0 = _refinement_objective(
        sched, num_nurses, holidays, forbidden_pairs, nurse_names, carry_in, nurses,
        requests=requests, carry_next_month=carry_next_month,
        shift_bans=shift_bans,
    )
    stale = 0
    err0 = score0[0]
    stale_limit = max(450, max_tries // 2)
    if err0 > 0:
        stale_limit = max(stale_limit, max_tries + max_tries // 2)

    def _accept(score1):
        nonlocal score0, stale
        if score1[0] > score0[0]:
            stale += 1
            return False, True
        improved = score1 < score0
        neut_p = 0.08 if score0[0] == 0 else 0.04
        neutral_ok = score1 == score0 and tie_rng.random() < neut_p
        if improved:
            score0 = score1
            stale = 0
            return True, False
        if neutral_ok:
            stale += 1
            return True, False
        stale += 1
        return False, True

    for _ in range(max_tries):
        if tie_rng.random() < 0.38:
            d = tie_rng.randint(1, NUM_DAYS)
            pool = [x for x in nurses if (x, d) not in locked]
            if len(pool) < 2:
                continue
            n1, n2 = tie_rng.sample(pool, 2)
            s1, s2 = sched[n1].get(d), sched[n2].get(d)
            if s1 == s2:
                continue
            sched[n1][d], sched[n2][d] = s2, s1
            score1 = _refinement_objective(
                sched, num_nurses, holidays, forbidden_pairs, nurse_names, carry_in, nurses,
                requests=requests, carry_next_month=carry_next_month,
                shift_bans=shift_bans,
            )
            ok, revert = _accept(score1)
            if not ok and revert:
                sched[n1][d], sched[n2][d] = s1, s2
            if stale >= stale_limit:
                break
            continue

        n = tie_rng.choice(nurses)
        d1, d2 = _refine_pick_swap_days(tie_rng)
        if (n, d1) in locked or (n, d2) in locked:
            continue
        s1, s2 = sched[n].get(d1), sched[n].get(d2)
        if s1 == s2:
            continue
        sched[n][d1], sched[n][d2] = s2, s1
        score1 = _refinement_objective(
            sched, num_nurses, holidays, forbidden_pairs, nurse_names, carry_in, nurses,
            requests=requests, carry_next_month=carry_next_month,
            shift_bans=shift_bans,
        )
        ok, revert = _accept(score1)
        if not ok and revert:
            sched[n][d1], sched[n][d2] = s1, s2
        if stale >= stale_limit:
            break


def _greedy_schedule(num_nurses, requests, holidays=(), forbidden_pairs=None, carry_in=None,
                     tie_rng=None, nurse_names=None, regenerate: bool = False,
                     carry_next_month=None, shift_bans=None):
    days = get_april_days(holidays)
    nurses = list(range(1, num_nurses))   # 일반간호사 인덱스
    fp_map = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    den_bans = _normalize_shift_bans(shift_bans, num_nurses)

    def den_banned(n, sh):
        if sh not in ('D', 'E', 'N'):
            return False
        b = den_bans.get(n)
        return bool(b and sh in b)

    carry_next = _normalize_carry_in(carry_next_month, num_nurses) if carry_next_month else {}
    carry_next_provided = carry_next_month is not None
    month_last = date(YEAR, MONTH, NUM_DAYS)

    OFF_SET = {'OF', 'OH', 'NO'}   # 오프 계열 (NO: N 누적 20회당·개인별 일자·수기만, 자동 배정 없음)
    # OF/OH/NO/연 사이: 근무일 수 2~5(공·EDU 포함), 5일 초과 절대 불가 — 섬(1일) 후처리 대상
    REST_GAP = frozenset(OFF_SET | {'연'})

    def is_off(shift):
        return shift in OFF_SET or shift == '연' or shift is None

    def is_streak_work(shift):
        """연속근무 5일 제한: D/E/N/공/EDU만 근무일(연차·병가 등은 끊김)."""
        return shift in STREAK_WORK_SHIFTS

    # ── 초기화 (전월 꼬리 carry — 월 경계 규칙) ─────────────────────────────
    sched = {n: {} for n in range(num_nurses)}
    carry = _normalize_carry_in(carry_in, num_nurses)

    def sk(n, dn, k):
        """dn일의 k일 전 근무 (전월 carry 반영)"""
        return _shift_k_days_before(sched, carry, n, dn, k)

    def work_streak_before(n, d):
        """d(0-indexed) 이전 연속 근무일수 (전월 말 carry까지 이어서 계산) — D/E/N/공/EDU만 합산."""
        count = 0
        for back in range(1, d + 1):
            s = sched[n].get(d - back + 1)   # dn = d-back+1
            if not is_streak_work(s):
                break
            count += 1
        if count < d:
            return count
        c = carry.get(n) or ()
        for s in reversed(c):
            if not is_streak_work(s):
                break
            count += 1
        return count

    def fp_same_shift_conflict(n, dn, shift):
        """n을 dn일 shift에 넣을 때, 이미 그 시프트인 동료와 함께 근무 불가 쌍이면 True"""
        if not fp_map:
            return False
        if shift not in ('D', 'E', 'N'):
            return False
        for m in range(num_nurses):
            if m == n:
                continue
            if sched[m].get(dn) != shift:
                continue
            key = (min(n, m), max(n, m))
            allowed = fp_map.get(key)
            if allowed and shift in allowed:
                return True
        return False

    # 토·일·공휴일 쉼(OF/OH/NO/연) 균등 배분 — 근무 후보·빈칸 채우기 정렬에 사용
    REST_WEEKEND = frozenset({'OF', 'OH', 'NO', '연'})

    def weekend_rest_count_before(n, dn):
        """1일~(dn-1)일까지 주말·공휴일에 쉼 시프트인 날 수."""
        t = 0
        for _d, _day in enumerate(days):
            _dn = _d + 1
            if _dn >= dn:
                break
            if not (_day['is_weekend'] or _day['is_holiday']):
                continue
            if sched[n].get(_dn) in REST_WEEKEND:
                t += 1
        return t

    def weekend_work_bias_key(n, d):
        """
        주말·공휴일 근무(N/E/D) 후보: 이미 이번 달에서 토·일 쉼이 많은 사람을 먼저 근무 배정
        (쉼이 적은 사람은 이후 빈칸 OF/OH로 토·일 쉼을 받기 쉽게).
        """
        _day = days[d]
        if not (_day['is_weekend'] or _day['is_holiday']):
            return 0
        dn = d + 1
        return -weekend_rest_count_before(n, dn)

    # 개인 신청 우선 적용
    for n_idx, day_shifts in requests.items():
        for day_num, shift_name in day_shifts.items():
            if 0 <= n_idx < num_nurses and 1 <= day_num <= NUM_DAYS:
                sched[n_idx][day_num] = shift_name

    req_locked = _request_locked_cells(requests, num_nurses)

    # 근무불가: 신청으로 고정되지 않은 칸만 금지 시프트 제거
    for n_idx in range(num_nurses):
        for dn in range(1, NUM_DAYS + 1):
            s = sched[n_idx].get(dn)
            if s not in ('D', 'E', 'N'):
                continue
            if den_banned(n_idx, s) and (n_idx, dn) not in req_locked:
                del sched[n_idx][dn]

    # N 직후 D 절대 불가 — 신청에 N→D가 있으면 해당 D 제거(빈칸으로 두고 이후 OF 등으로 채움)
    # 신청으로 고정된 셀은 삭제·덮어쓰기하지 않음
    for n_idx in range(num_nurses):
        for dn in range(2, NUM_DAYS + 1):
            if sched[n_idx].get(dn) == 'D' and sched[n_idx].get(dn - 1) == 'N':
                if (n_idx, dn) not in req_locked:
                    del sched[n_idx][dn]
        # 1일: 전월 말 N → 당월 1일 D 도 금지
        if sched[n_idx].get(1) == 'D' and sk(n_idx, 1, 1) == 'N':
            if (n_idx, 1) not in req_locked:
                del sched[n_idx][1]

    # N블록 직후(연속 N 제외): 공휴→OH, 그 외→OF 만 허용 — 그 밖 신청 제거
    for n_idx in range(num_nurses):
        for dn in range(1, NUM_DAYS + 1):
            if sk(n_idx, dn, 1) != 'N':
                continue
            cur = sched[n_idx].get(dn)
            if cur == 'N':
                continue
            need = 'OH' if days[dn - 1]['is_holiday'] else 'OF'
            if cur is not None and cur != need:
                if (n_idx, dn) not in req_locked:
                    del sched[n_idx][dn]

    # ── 수간호사 배정 ─────────────────────────────────────────────────────────
    for d, day in enumerate(days):
        dn = d + 1
        if dn not in sched[0]:
            if day['is_holiday']:
                sched[0][dn] = 'OH'
            elif day['is_weekend']:
                sched[0][dn] = 'OF'
            else:
                sched[0][dn] = 'A1'

    def streak_total(n, d):
        """
        d(0-indexed)에 근무를 배정했을 때 예상 연속근무일수.
        앞·뒤 = D/E/N/공/EDU 연속만 합산(연차 등은 끊김).
        """
        before = work_streak_before(n, d)
        after = 0
        for fwd in range(1, NUM_DAYS - d):
            s = sched[n].get(d + fwd + 1)     # dn of (d+fwd)
            if not is_streak_work(s):
                break
            after += 1
        return before + 1 + after

    # ── N 시프트 배정 (공평 분배, 간호사당 최대 N_ABS_MAX야) ───────────────────
    # 총 슬롯 = 2 × 말일. 일반간호사 num_reg명에게 합이 total_n_slots가 되도록 분배(각 ≤ N_ABS_MAX).
    # 수간 포함 11명 이상이어도 N 상한·매일 2명·공평 목표 분배 원칙은 동일(D만 평일 목표가 달라짐).
    num_reg = len(nurses)
    total_n_slots = 2 * NUM_DAYS
    _nt_list = _compute_n_targets_fair(num_reg, total_n_slots, N_ABS_MAX)
    n_target = {}
    n_block_plan = {}
    _pat_idx = 0
    for i, n in enumerate(nurses):
        t = _nt_list[i]
        n_target[n] = t
        n_block_plan[n] = _n_block_plan_for_target(t, _pat_idx)
        if t >= 6:
            _pat_idx += 1

    # 요청으로 이미 배정된 N을 total에 반영
    def get_block_plan_max(n, block_num):
        plan = n_block_plan[n]
        if block_num < len(plan):
            return plan[block_num]
        return plan[-1]   # 마지막 블록 크기 재사용

    # consec: 현재 블록 연속일, ok_from: 다음 N 허용 날짜(0-indexed)
    # block_num: 완료된 블록 수 (0=첫 블록 진행 중)
    ns = {n: {'total': 0, 'consec': 0, 'ok_from': 0, 'block_num': 0} for n in nurses}
    for n in nurses:
        ns[n]['total'] = sum(1 for v in sched[n].values() if v == 'N')
    # 전월 말에서 이어진 N 연속 블록 — 1일차 consec 초기값
    for n in nurses:
        c = carry.get(n)
        if not c:
            continue
        consec = 0
        for s in reversed(c):
            if s == 'N':
                consec += 1
            else:
                break
        ns[n]['consec'] = consec

    def _n_place_is_nonterminal_lone(ni, dayn):
        """dayn에 N을 넣을 때 말일이 아닌 단독 N이 되면 True (배정 전). 말일=NUM_DAYS(31일 말달이면 31)."""
        if dayn == NUM_DAYS:
            return False
        if sk(ni, dayn, 1) == 'N':
            return False
        if dayn < NUM_DAYS and sched[ni].get(dayn + 1) == 'N':
            return False
        return True

    def _n_emergency_can_pair_nonterminal_lone(ni, dayn):
        """
        말일 아닌 단독 N이 될 경우, 다음날·전날 중 하나에 같은 사람 N을
        이어 붙일 수 있어야 함(일 N<2·칸 비움/휴).
        """
        if not _n_place_is_nonterminal_lone(ni, dayn):
            return True
        if dayn < NUM_DAYS:
            n_next = sum(1 for m in nurses if sched[m].get(dayn + 1) == 'N')
            nx = sched[ni].get(dayn + 1)
            if (
                n_next < 2
                and nx in (None, 'OF', 'OH', 'NO')
                and (ni, dayn + 1) not in req_locked
            ):
                return True
        if dayn > 1:
            n_prev = sum(1 for m in nurses if sched[m].get(dayn - 1) == 'N')
            px = sched[ni].get(dayn - 1)
            if (
                n_prev < 2
                and px in (None, 'OF', 'OH', 'NO')
                and (ni, dayn - 1) not in req_locked
            ):
                return True
        return False

    for d in range(NUM_DAYS):
        dn = d + 1
        on_n = [n for n in nurses if sched[n].get(dn) == 'N']
        needed = 2 - len(on_n)

        if needed > 0:
            cands = []
            for n in nurses:
                if n in on_n or dn in sched[n]:
                    continue
                if den_banned(n, 'N'):
                    continue
                if ns[n]['total'] >= N_ABS_MAX:
                    continue
                if ns[n]['total'] >= n_target[n]:      # 개인 목표 초과 금지
                    continue
                if d < ns[n]['ok_from']:
                    # 말일에 N 1개 남은 경우(3,3,1)는 ok_from 완화
                    if not (d == NUM_DAYS - 1 and n_target[n] - ns[n]['total'] == 1):
                        continue
                cur_max = get_block_plan_max(n, ns[n]['block_num'])
                if ns[n]['consec'] >= cur_max:
                    continue    # 현재 블록 최대 야간수 초과
                # 새 블록 시작 조건
                if ns[n]['consec'] == 0:
                    remaining = n_target[n] - ns[n]['total']
                    is_last_day = (d == NUM_DAYS - 1)
                    # 말일에 N 1개만 남은 경우(3,3,1 패턴)는 단독 블록 허용
                    terminal_single = (remaining == 1 and is_last_day)
                    if remaining < 2 and not terminal_single:
                        continue
                    if d >= NUM_DAYS - 1 and not terminal_single:
                        continue
                    next_assigned = sched[n].get(dn + 1)
                    if next_assigned is not None and next_assigned != 'N':
                        continue
                    if (
                        next_assigned is None
                        and not terminal_single
                        and not _n_emergency_can_pair_nonterminal_lone(n, dn)
                    ):
                        continue
                    # 앞 근무 연속 + 최소 야간 수 ≤ 5
                    min_n = 1 if terminal_single else 2
                    if work_streak_before(n, d) + min_n > 5:
                        continue
                else:
                    # 블록 연장: 앞뒤 합산 연속 ≤ 5
                    if work_streak_before(n, d) + 1 > 5:
                        continue
                if fp_same_shift_conflict(n, dn, 'N'):
                    continue
                cands.append(n)

            _ti_n = _tie_break_map(cands, tie_rng)

            def _main_n_merge_key(ni):
                if dn < NUM_DAYS and sched[ni].get(dn + 1) == 'N':
                    return 0
                if sk(ni, dn, 1) == 'N':
                    return 1
                return 2

            def n_score(n, _d=d):
                continuing = ns[n]['consec'] > 0
                tgt = n_target[n]
                expected = (_d / NUM_DAYS) * tgt
                deficit = expected - ns[n]['total']
                return (
                    _main_n_merge_key(n),
                    0 if continuing else 1,
                    weekend_work_bias_key(n, _d),
                    -round(deficit * 4),
                    ns[n]['total'],
                    _ti_n.get(n, 0),
                    n,
                )

            cands.sort(key=n_score)
            placed = 0
            for n in cands:
                if placed >= needed:
                    break
                if den_banned(n, 'N'):
                    continue
                if ns[n]['total'] >= N_ABS_MAX:
                    continue
                if fp_same_shift_conflict(n, dn, 'N'):
                    continue
                sched[n][dn] = 'N'
                ns[n]['total'] += 1
                on_n.append(n)
                placed += 1

        for n in nurses:
            if sched[n].get(dn) == 'N':
                ns[n]['consec'] += 1
            else:
                if ns[n]['consec'] > 0:
                    ns[n]['ok_from'] = d + 6
                    ns[n]['block_num'] += 1   # 블록 종료 → 완료 블록 수 증가
                ns[n]['consec'] = 0

    # ★ N 부족 날짜 구제: N<2 인 날 찾아 인접 블록 연장 또는 목표 1 올려 재시도
    for d in range(NUM_DAYS):
        dn = d + 1
        on_n = [n for n in nurses if sched[n].get(dn) == 'N']
        if len(on_n) >= 2:
            continue
        # 이 날 N이 1명 이하 → 추가 배정 시도 (목표+1 임시 허용)
        def _repair_n_merge_key(ni):
            if dn < NUM_DAYS and sched[ni].get(dn + 1) == 'N':
                return (0, ns[ni]['total'], ni)
            if sk(ni, dn, 1) == 'N':
                return (1, ns[ni]['total'], ni)
            return (2, ns[ni]['total'], ni)

        for n in sorted(nurses, key=_repair_n_merge_key):
            if n in on_n:
                continue
            if (n, dn) in req_locked:
                continue
            curd = sched[n].get(dn)
            if curd is not None and curd not in ('NO', 'OF', 'OH'):
                continue
            if den_banned(n, 'N'):
                continue
            if sk(n, dn, 1) == 'N':
                continue
            if ns[n]['total'] >= N_ABS_MAX:
                continue
            if d < ns[n]['ok_from']:
                # 다음날 이미 N이면 ok_from 무시(OF 위에 N 잇기)
                merge_skip_ok = dn < NUM_DAYS and sched[n].get(dn + 1) == 'N'
                if not merge_skip_ok and not (
                    d == NUM_DAYS - 1 and ns[n]['total'] < n_target[n]
                ):
                    continue
            # 기존 블록 연장 (consec > 0)
            if ns[n]['consec'] > 0:
                cur_max = get_block_plan_max(n, ns[n]['block_num'])
                if ns[n]['consec'] < cur_max + 1:   # 1야 추가 허용
                    if work_streak_before(n, d) + 1 <= 5:
                        if not fp_same_shift_conflict(n, dn, 'N'):
                            sched[n][dn] = 'N'
                            ns[n]['total'] += 1
                            on_n.append(n)
                            if len(on_n) >= 2:
                                break
            # 새 블록 시작 (목표 미달 간호사) – 말일 단독 N 허용
            # 다음날 이미 N이면 OF 등 위에 N을 잇기 위해 목표 도달 후 +1 허용(말일 아닌 단독 N 방지)
            elif (
                (
                    ns[n]['total'] < n_target[n]
                    or (
                        dn < NUM_DAYS
                        and sched[n].get(dn + 1) == 'N'
                        and ns[n]['total'] == n_target[n]
                    )
                )
                and ns[n]['total'] < N_ABS_MAX
                and (
                    d < NUM_DAYS - 1
                    or (d == NUM_DAYS - 1 and n_target[n] - ns[n]['total'] == 1)
                )
            ):
                next_s = sched[n].get(dn + 1)
                is_last = (d == NUM_DAYS - 1)
                min_streak = 1 if is_last else 2
                rem = n_target[n] - ns[n]['total']
                ok_new = is_last or next_s == 'N'
                if not ok_new and next_s is None and rem >= 2:
                    ok_new = _n_emergency_can_pair_nonterminal_lone(n, dn)
                if ok_new:
                    if work_streak_before(n, d) + min_streak <= 5:
                        if not fp_same_shift_conflict(n, dn, 'N'):
                            sched[n][dn] = 'N'
                            ns[n]['total'] += 1
                            on_n.append(n)
                            if len(on_n) >= 2:
                                break

    # ★ 고립 N 제거 (isolated N → 미배정으로 되돌림)
    # 단, 제거 후 당일 N이 2명 미만이 되면 제거하지 않음 (N 매일 2명 절대 규칙)
    for n in nurses:
        for d in range(NUM_DAYS):
            dn = d + 1
            if sched[n].get(dn) != 'N':
                continue
            prev_n = sk(n, dn, 1) == 'N'
            next_n = d < NUM_DAYS - 1 and sched[n].get(dn + 1) == 'N'
            if not prev_n and not next_n:
                # 말일 단독 N: 3-3-1 / 2-3-1 / 3-2-1 등 허용 — 제거하지 않음
                if dn == NUM_DAYS:
                    continue
                day_n_cnt = sum(1 for m in nurses if sched[m].get(dn) == 'N')
                if day_n_cnt >= 3:   # 제거 후에도 ≥2명 유지 가능할 때만 제거
                    if (n, dn) not in req_locked:
                        del sched[n][dn]
                        ns[n]['total'] -= 1
                # 2명 이하이면 고립 N이라도 유지 (매일 2명 우선)

    def _n_block_tail_assign_of():
        """N 연속이 끝난 다음날(달 내) 비었으면 공휴 OH / 아니면 OF.
        말일이 아닌 날의 단독 N은 다음날을 비워 두어 2연속 N으로 보완 가능하게 함."""
        for n in nurses:
            for d in range(NUM_DAYS - 1):
                dn, ndn = d + 1, d + 2
                if sched[n].get(dn) == 'N' and sched[n].get(ndn) != 'N':
                    prev_n = sk(n, dn, 1) == 'N'
                    if (
                        not prev_n
                        and dn < NUM_DAYS
                    ):
                        continue
                    if ndn not in sched[n] and (n, ndn) not in req_locked:
                        sched[n][ndn] = 'OH' if days[ndn - 1]['is_holiday'] else 'OF'

    _n_block_tail_assign_of()

    def _fix_nonterminal_lone_n():
        """말일이 아닌 단독 N → 다음날 또는 전날로 2연속 병합(일일 N=2·상한·금지쌍 준수)."""
        for _ in range(NUM_DAYS * len(nurses) + 5):
            found = False
            for n in nurses:
                for d in range(NUM_DAYS):
                    dn = d + 1
                    if sched[n].get(dn) != 'N' or dn == NUM_DAYS:
                        continue
                    prev_n = sk(n, dn, 1) == 'N'
                    next_n = dn < NUM_DAYS and sched[n].get(dn + 1) == 'N'
                    if prev_n or next_n:
                        continue
                    for ndn, back in ((dn + 1, False), (dn - 1, True)):
                        if ndn < 1 or ndn > NUM_DAYS:
                            continue
                        if (n, ndn) in req_locked or den_banned(n, 'N'):
                            continue
                        cur = sched[n].get(ndn)
                        if cur is not None and cur not in ('OF', 'OH', 'NO'):
                            continue
                        if ns[n]['total'] >= N_ABS_MAX:
                            continue
                        d2 = ndn - 1
                        if work_streak_before(n, d2) + 1 > 5:
                            continue
                        if fp_same_shift_conflict(n, ndn, 'N'):
                            continue
                        on_ndn = sum(1 for m in nurses if sched[m].get(ndn) == 'N')
                        if on_ndn >= 2:
                            continue
                        sched[n][ndn] = 'N'
                        ns[n]['total'] += 1
                        found = True
                        break
                    if found:
                        break
                if found:
                    break
            if not found:
                break

    def _handoff_lone_n_prev_neighbor():
        """전날 N인 다른 간호사가 같은 날 OF/빈칸이면 말일 아닌 단독 N을 넘겨 연속 N으로 만든다."""
        for _ in range(len(nurses) * NUM_DAYS + 5):
            hit = False
            for n in nurses:
                for d in range(NUM_DAYS):
                    dn = d + 1
                    if dn == NUM_DAYS or sched[n].get(dn) != 'N':
                        continue
                    prev_n = sk(n, dn, 1) == 'N'
                    next_n = dn < NUM_DAYS and sched[n].get(dn + 1) == 'N'
                    if prev_n or next_n:
                        continue
                    partners = [m for m in nurses if m != n and sched[m].get(dn) == 'N']
                    if len(partners) != 1:
                        continue
                    for m in sorted(nurses):
                        if m == n or sched[m].get(dn) == 'N':
                            continue
                        if sk(m, dn, 1) != 'N':
                            continue
                        curm = sched[m].get(dn)
                        if curm not in (None, 'NO', 'OF', 'OH'):
                            continue
                        if den_banned(m, 'N'):
                            continue
                        if (n, dn) in req_locked or (m, dn) in req_locked:
                            continue
                        if fp_same_shift_conflict(m, dn, 'N'):
                            continue
                        if work_streak_before(m, d) + 1 > 5:
                            continue
                        if ns[m]['total'] >= N_ABS_MAX:
                            continue
                        del sched[n][dn]
                        ns[n]['total'] -= 1
                        sched[m][dn] = 'N'
                        ns[m]['total'] += 1
                        hit = True
                        break
                    if hit:
                        break
                if hit:
                    break
            if not hit:
                break

    def _n_top_up_under_target():
        """N 목표 미달 간호사에 단독 N 금지 조건을 지키며 1개씩 보충."""
        for _ in range(NUM_DAYS * len(nurses) + 5):
            short = [x for x in nurses if ns[x]['total'] < n_target[x]]
            if not short:
                break
            progressed = False
            for n in sorted(short, key=lambda x: (ns[x]['total'] - n_target[x], x)):
                placed_here = False
                for d in range(NUM_DAYS):
                    dn = d + 1
                    if sched[n].get(dn) == 'N':
                        continue
                    cur = sched[n].get(dn)
                    if cur not in (None, 'NO', 'OF', 'OH'):
                        continue
                    if (n, dn) in req_locked:
                        continue
                    if den_banned(n, 'N'):
                        break
                    if sk(n, dn, 1) == 'N':
                        continue
                    if ns[n]['total'] >= N_ABS_MAX:
                        break
                    if work_streak_before(n, d) + 1 > 5:
                        continue
                    if fp_same_shift_conflict(n, dn, 'N'):
                        continue
                    if sum(1 for m in nurses if sched[m].get(dn) == 'N') >= 2:
                        continue
                    if (
                        dn < NUM_DAYS
                        and _n_place_is_nonterminal_lone(n, dn)
                        and not _n_emergency_can_pair_nonterminal_lone(n, dn)
                    ):
                        continue
                    sched[n][dn] = 'N'
                    ns[n]['total'] += 1
                    placed_here = True
                    progressed = True
                    break
                if not placed_here:
                    continue
            if not progressed:
                break

    # 스케줄과 ns['total'] 불일치 시 긴급 N이 상한을 넘길 수 있음 → 동기화
    for n in nurses:
        ns[n]['total'] = sum(1 for v in sched[n].values() if v == 'N')

    # ★ 긴급 N 보완: 위 모든 단계 후에도 N<2인 날 → 최소 제약으로 강제 배정
    # 절대 규칙 "N 매일 2명" 보장. ok_from·블록 크기 무시, 연속 5일 제한만 유지.
    for d in range(NUM_DAYS):
        dn = d + 1
        on_n = [n for n in nurses if sched[n].get(dn) == 'N']
        if len(on_n) >= 2:
            continue

        # 후보: 전일 N이면 오늘은 의무 휴(OH/OF) 칸 → N 불가. 그 외만 빈/NO·일반 OF/OH 교체
        cands = [
            n for n in nurses
            if n not in on_n
            and not den_banned(n, 'N')
            and sk(n, dn, 1) != 'N'
            and sched[n].get(dn) in (None, 'NO', 'OF', 'OH')
            and work_streak_before(n, d) + 1 <= 5   # 연속 5일 제한
            and (n, dn) not in req_locked
        ]
        # 말일 아닌 단독 N은 이후 이어 붙일 수 없으면 후보에서 제외(다른 후보가 있을 때)
        cands_safe = [n for n in cands if _n_emergency_can_pair_nonterminal_lone(n, dn)]
        cands_use = cands_safe if cands_safe else cands
        # 우선순위: 말일 아닌 단독 N 회피 → 주말·공휴 쉼 균등 → N 총 개수 적은 순
        _ti_em = _tie_break_map(cands_use, tie_rng)

        def _em_merge_key(ni):
            if dn < NUM_DAYS and sched[ni].get(dn + 1) == 'N':
                return 0
            if sk(ni, dn, 1) == 'N':
                return 1
            return 2

        cands_use.sort(
            key=lambda n: (
                _em_merge_key(n),
                1 if _n_place_is_nonterminal_lone(n, dn) else 0,
                weekend_work_bias_key(n, d),
                ns[n]['total'],
                _ti_em.get(n, 0),
                n,
            )
        )

        needed = 2 - len(on_n)
        placed = 0
        for relax_lone in (False, True):
            for n in cands_use:
                if placed >= needed:
                    break
                if sched[n].get(dn) == 'N':
                    continue
                if den_banned(n, 'N'):
                    continue
                if ns[n]['total'] >= N_ABS_MAX:
                    continue
                if fp_same_shift_conflict(n, dn, 'N'):
                    continue
                ndn = dn + 1
                lone_bad = (
                    dn < NUM_DAYS
                    and _n_place_is_nonterminal_lone(n, dn)
                    and not _n_emergency_can_pair_nonterminal_lone(n, dn)
                )
                pair_ok = (
                    lone_bad
                    and ndn <= NUM_DAYS
                    and not den_banned(n, 'N')
                    and sched[n].get(ndn) in (None, 'NO', 'OF', 'OH')
                    and (n, ndn) not in req_locked
                    and not fp_same_shift_conflict(n, ndn, 'N')
                    and sum(1 for m in nurses if sched[m].get(ndn) == 'N') < 2
                    and ns[n]['total'] + 2 <= N_ABS_MAX
                    and work_streak_before(n, d) + 2 <= 5
                )
                if lone_bad and pair_ok:
                    sched[n][dn] = 'N'
                    sched[n][ndn] = 'N'
                    ns[n]['total'] += 2
                    on_n.append(n)
                    placed += 1
                    continue
                if lone_bad and not relax_lone:
                    continue
                sched[n][dn] = 'N'
                ns[n]['total'] += 1
                on_n.append(n)
                placed += 1
            if placed >= needed:
                break

    _fix_nonterminal_lone_n()
    _handoff_lone_n_prev_neighbor()
    _n_top_up_under_target()
    _n_block_tail_assign_of()

    # ── E 시프트 배정 (매일 정확히 2명) ──────────────────────────────────────
    # 근무 연속성 우선: 이미 근무 중(streak 1~4)인 간호사 먼저 배정해
    # OF-단일E-OF 패턴을 자연스럽게 줄임
    e_cnt = {n: sum(1 for v in sched[n].values() if v == 'E') for n in nurses}

    for d, day in enumerate(days):
        dn = d + 1
        on_e = [n for n in nurses if sched[n].get(dn) == 'E']
        needed = 2 - len(on_e)
        if needed <= 0:
            continue

        cands = [
            n for n in nurses
            if n not in on_e and dn not in sched[n]
            and not den_banned(n, 'E')
            and sk(n, dn, 1) != 'N'   # N블록 직후날은 휴무(OF/OH)만 (E 불가)
            and sched[n].get(dn + 1) not in ('D', 'EDU', '공')
            and streak_total(n, d) <= 5   # 연속근무 5일 초과 금지
            and not fp_same_shift_conflict(n, dn, 'E')
        ]
        _ti_e = _tie_break_map(cands, tie_rng)

        def e_score(n):
            streak = work_streak_before(n, d)
            prio = 0 if 1 <= streak <= 4 else (1 if streak == 0 else 2)
            return (prio, weekend_work_bias_key(n, d), e_cnt[n], _ti_e.get(n, 0), n)

        cands.sort(key=e_score)
        placed = 0
        for n in cands:
            if placed >= needed:
                break
            if den_banned(n, 'E'):
                continue
            if fp_same_shift_conflict(n, dn, 'E'):
                continue
            sched[n][dn] = 'E'
            e_cnt[n] += 1
            placed += 1

    # ── D 시프트 배정 (공평 분배: 4~6개/인) ──────────────────────────────────
    d_cnt = {n: sum(1 for v in sched[n].values() if v == 'D') for n in nurses}

    # 총 D 슬롯 계산 → 간호사별 목표·상한 설정
    total_d_slots = sum(
        d_assignment_target(num_nurses, days[_d], sched[0].get(_d + 1) == 'A1')
        for _d in range(NUM_DAYS)
    )
    num_reg = len(nurses)
    D_ABS_MIN, D_ABS_MAX = 4, 6
    _d_base  = total_d_slots // num_reg            # e.g. 38//9 = 4
    _d_extra = total_d_slots % num_reg             # e.g. 38%9  = 2
    # extra명은 (base+1)D, 나머지는 baseD  → 4~5 범위
    n_d_cap = {}
    for _i, _n in enumerate(nurses):
        raw = _d_base + (1 if _i < _d_extra else 0)
        n_d_cap[_n] = max(D_ABS_MIN, min(raw, D_ABS_MAX))

    for d, day in enumerate(days):
        dn = d + 1
        head_a1 = sched[0].get(dn) == 'A1'
        d_target = d_assignment_target(num_nurses, day, head_a1)

        on_d = [n for n in nurses if sched[n].get(dn) == 'D']
        needed = d_target - len(on_d)
        if needed <= 0:
            continue

        def can_d(n):
            if dn in sched[n]:
                return False
            if den_banned(n, 'D'):
                return False
            if sk(n, dn, 1) == 'E':
                return False                                    # E-D 금지
            if sk(n, dn, 1) == 'N':
                return False                                    # N-D 금지 (전날 야간 직후 데이 불가)
            if sk(n, dn, 2) == 'N' and sk(n, dn, 1) in OFF_SET:
                return False                                    # N-OF-D 금지
            return True

        # 절대 상한(6) 미만인 간호사 모두 후보, 스코어로 공평 배정
        cands = [n for n in nurses if n not in on_d and can_d(n)
                 and d_cnt[n] < D_ABS_MAX
                 and streak_total(n, d) <= 5
                 and not fp_same_shift_conflict(n, dn, 'D')]
        _ti_d = _tie_break_map(cands, tie_rng)

        def d_score(n):
            over = max(0, d_cnt[n] - n_d_cap[n])
            p = sk(n, dn, 1)
            nxt = sched[n].get(dn + 1) if d < NUM_DAYS - 1 else None
            isolation = 1 if ((p is None or p in REST_GAP) and
                              (nxt is None or nxt in REST_GAP)) else 0
            streak = work_streak_before(n, d)
            streak_prio = 0 if 1 <= streak <= 4 else 1
            return (
                over,
                weekend_work_bias_key(n, d),
                isolation,
                d_cnt[n],
                streak_prio,
                _ti_d.get(n, 0),
                n,
            )

        cands.sort(key=d_score)
        placed = 0
        for n in cands:
            if placed >= needed:
                break
            if den_banned(n, 'D'):
                continue
            if fp_same_shift_conflict(n, dn, 'D'):
                continue
            sched[n][dn] = 'D'
            d_cnt[n] += 1
            placed += 1

    # ── 나머지 빈칸 → OF / OH ────────────────────────────────────────────────
    # 빈칸: 공휴 OH / 평일 OF. 주말·공휴일은 누적 토·일 쉼이 적은 사람부터 배정해 쉼 균등화
    for d, day in enumerate(days):
        dn = d + 1
        is_we = day['is_weekend'] or day['is_holiday']
        need_ns = [
            n for n in range(num_nurses)
            if dn not in sched[n] and (n, dn) not in req_locked
        ]
        if not need_ns:
            continue
        if is_we:
            need_ns.sort(key=lambda n: (weekend_rest_count_before(n, dn), n))
        for n in need_ns:
            sched[n][dn] = 'OH' if day['is_holiday'] else 'OF'

    # ★ 후처리 ⓪: D 최솟값(4개) 미달 간호사 → OF → D로 보충
    #   연속5일·E-D·N-OF-D 제약 지키면서 OF를 D로 전환
    for n in nurses:
        if d_cnt.get(n, 0) >= D_ABS_MIN:
            continue
        for d, day in enumerate(days):
            if d_cnt.get(n, 0) >= D_ABS_MIN:
                break
            dn = d + 1
            if sched[n].get(dn) != 'OF':
                continue
            # 제약 확인
            if den_banned(n, 'D'):
                continue
            if sk(n, dn, 1) == 'E':
                continue
            if sk(n, dn, 1) == 'N':
                continue
            if sk(n, dn, 2) == 'N' and sk(n, dn, 1) in OFF_SET:
                continue
            if streak_total(n, d) > 5:
                continue
            # 해당 날 D 정원 확인
            head_a1 = sched[0].get(dn) == 'A1'
            d_max_day = d_assignment_target(num_nurses, day, head_a1)
            d_on = sum(1 for m in nurses if sched[m].get(dn) == 'D')
            if d_on >= d_max_day:
                continue
            if fp_same_shift_conflict(n, dn, 'D'):
                continue
            if (n, dn) in req_locked:
                continue
            sched[n][dn] = 'D'
            d_cnt[n] = d_cnt.get(n, 0) + 1

    _convert_non_request_yun_to_d(
        sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
    )
    _auto_yun_to_of_if_quota_room(sched, num_nurses, holidays, carry_in, requests)

    # ★ 후처리 ①: OF-단일근무-OF 섬 → 가능하면 OF로 통합
    #   반복 적용해서 연쇄 섬도 처리
    changed = True
    while changed:
        changed = False
        for n in nurses:
            for d in range(1, NUM_DAYS - 1):
                dn = d + 1
                prev_s = sched[n].get(dn - 1)
                curr_s = sched[n].get(dn)
                next_s = sched[n].get(dn + 1)
                if prev_s not in REST_GAP or next_s not in REST_GAP:
                    continue
                if curr_s in REST_GAP:
                    continue
                # 쉬는날-work-쉬는날 섬 감지 (OF/연/OH/NO)
                if curr_s == 'D':
                    day = days[d]
                    ha1 = sched[0].get(dn) == 'A1'
                    d_min = d_assignment_target(num_nurses, day, ha1)
                    others_d = sum(1 for m in nurses if m != n and sched[m].get(dn) == 'D')
                    if others_d >= d_min and (n, dn) not in req_locked:
                        sched[n][dn] = 'OF'
                        d_cnt[n] -= 1
                        changed = True
                elif curr_s == 'E':
                    others_e = sum(1 for m in nurses if m != n and sched[m].get(dn) == 'E')
                    if others_e >= 2 and (n, dn) not in req_locked:
                        sched[n][dn] = 'OF'
                        e_cnt[n] -= 1
                        changed = True

    # ★ 후처리 ②: OF-단일근무-OF 섬 → 뒤 또는 앞 OF를 근무로 연장
    #   후처리①로 해소 불가(인원 부족)인 섬을 인접 OF를 근무일로 바꿔 2일 연속으로 만듦
    for n in nurses:
        for d in range(1, NUM_DAYS - 1):
            dn = d + 1
            prev_s = sched[n].get(dn - 1)
            curr_s = sched[n].get(dn)
            next_s = sched[n].get(dn + 1)
            if prev_s not in REST_GAP or next_s not in REST_GAP:
                continue
            if curr_s in REST_GAP or curr_s is None:
                continue
            # ── 섬 확인: 앞으로 연장 시도 (next_dn 을 E/D로) ──────────────
            next_dn = dn + 1
            next_d  = d + 1
            if next_d < NUM_DAYS and sched[n].get(next_dn) in OFF_SET:
                after_next = sched[n].get(next_dn + 1) if next_d + 1 < NUM_DAYS else None
                e_on_next  = sum(1 for m in nurses if sched[m].get(next_dn) == 'E')
                # E 연장: E 인원 부족하고 E-D 위반 없고 5일 한도 내
                if (e_on_next < 2 and
                        not den_banned(n, 'E') and
                        after_next not in ('D', 'EDU', '공') and
                        streak_total(n, next_d) <= 5 and (n, next_dn) not in req_locked):
                    sched[n][next_dn] = 'E'
                    e_cnt[n] = e_cnt.get(n, 0) + 1
                    continue
                # D 연장: 정원 미달일 때만 허용 (평일 수간A1 → D=1 유지)
                day_nx = days[next_d]
                ha1_nx = sched[0].get(next_dn) == 'A1'
                d_min_nx = d_assignment_target(num_nurses, day_nx, ha1_nx)
                d_on_nx = sum(1 for m in nurses if sched[m].get(next_dn) == 'D')
                e_prev_nx = sk(n, next_dn, 1) == 'E'  # E-D 금지
                n_prev_nx = sk(n, next_dn, 1) == 'N'  # N-D 금지
                n_of_d_nx = (sk(n, next_dn, 2) == 'N' and sk(n, next_dn, 1) in OFF_SET)
                if (d_on_nx < d_min_nx and
                        not den_banned(n, 'D') and
                        not e_prev_nx and not n_prev_nx and not n_of_d_nx and
                        streak_total(n, next_d) <= 5 and (n, next_dn) not in req_locked):
                    sched[n][next_dn] = 'D'
                    d_cnt[n] = d_cnt.get(n, 0) + 1
                    continue
            # ── 뒤로 연장 시도 (prev_dn 을 E/D로) ────────────────────────────
            prev_dn = dn - 1
            prev_d  = d - 1
            if prev_d >= 0 and sched[n].get(prev_dn) in OFF_SET:
                e_on_prev  = sum(1 for m in nurses if sched[m].get(prev_dn) == 'E')
                after_prev = sched[n].get(prev_dn + 1)  # = curr_s, should not be D for E
                # E 연장 (E-D 금지: E 다음날이 D면 안 됨)
                if (e_on_prev < 2 and
                        not den_banned(n, 'E') and
                        after_prev not in ('D', 'EDU', '공') and
                        streak_total(n, prev_d) <= 5 and (n, prev_dn) not in req_locked):
                    sched[n][prev_dn] = 'E'
                    e_cnt[n] = e_cnt.get(n, 0) + 1
                    continue
                # D 연장 (뒤쪽 OF): 정원 미달일 때만 허용 (평일 수간A1 → D=1 유지)
                day_pv = days[prev_d]
                ha1_pv = sched[0].get(prev_dn) == 'A1'
                d_min_pv = d_assignment_target(num_nurses, day_pv, ha1_pv)
                d_on_pv  = sum(1 for m in nurses if sched[m].get(prev_dn) == 'D')
                e_prev_pv = sk(n, prev_dn, 1) == 'E'
                n_prev_pv = sk(n, prev_dn, 1) == 'N'
                n_of_d_pv = (sk(n, prev_dn, 2) == 'N' and sk(n, prev_dn, 1) in OFF_SET)
                if (d_on_pv < d_min_pv and
                        not den_banned(n, 'D') and
                        not e_prev_pv and not n_prev_pv and not n_of_d_pv and
                        streak_total(n, prev_d) <= 5 and (n, prev_dn) not in req_locked):
                    sched[n][prev_dn] = 'D'
                    d_cnt[n] = d_cnt.get(n, 0) + 1

    # ★ 후처리 ③: 확장도 불가한 섬 → 다른 간호사와 D 스왑
    #   섬 간호사의 D를 같은 날 OF인 다른 간호사에게 넘겨 양쪽 모두 섬이 안 되게 함
    for n in nurses:
        for d in range(1, NUM_DAYS - 1):
            dn = d + 1
            prev_s = sched[n].get(dn - 1)
            curr_s = sched[n].get(dn)
            next_s = sched[n].get(dn + 1)
            if prev_s not in REST_GAP or next_s not in REST_GAP:
                continue
            if curr_s != 'D':   # D 섬만 대상
                continue
            # 이 날 OF인 다른 간호사 중 스왑 적합한 후보 찾기
            swapped = False
            for m in nurses:
                if m == n or sched[m].get(dn) != 'OF':
                    continue
                m_prev = sk(m, dn, 1)
                m_next = sched[m].get(dn + 1)
                # m이 이 날 D를 해도 섬이 안 되는 경우 (앞이나 뒤에 근무가 있을 것)
                m_would_be_island = (m_prev in REST_GAP or m_prev is None) and m_next in REST_GAP
                if m_would_be_island:
                    continue
                # D 배정 제약 확인 (n→OF, m→D)
                m_e_prev = sk(m, dn, 1) == 'E'
                m_n_prev = sk(m, dn, 1) == 'N'
                m_n_of_d = (sk(m, dn, 2) == 'N' and sk(m, dn, 1) in OFF_SET)
                if m_e_prev or m_n_prev or m_n_of_d:
                    continue
                if streak_total(m, d) > 5:
                    continue
                if fp_same_shift_conflict(m, dn, 'D'):
                    continue
                if den_banned(m, 'D'):
                    continue
                # 스왑 실행
                if (n, dn) in req_locked or (m, dn) in req_locked:
                    continue
                sched[n][dn] = 'OF'
                d_cnt[n] = d_cnt.get(n, 0) - 1
                sched[m][dn] = 'D'
                d_cnt[m] = d_cnt.get(m, 0) + 1
                swapped = True
                break

    # ★ 후처리 ④: 주말 D 부족 재시도
    #   섬 제거 후 D를 OF로 돌린 결과 주말 D가 1명으로 줄었을 경우 복구
    for d, day in enumerate(days):
        dn = d + 1
        head_a1 = sched[0].get(dn) == 'A1'
        d_target = d_assignment_target(num_nurses, day, head_a1)
        on_d = [n for n in nurses if sched[n].get(dn) == 'D']
        if len(on_d) >= d_target:
            continue

        def can_d2(n):
            if sched[n].get(dn) != 'OF':
                return False
            if den_banned(n, 'D'):
                return False
            if sk(n, dn, 1) == 'E':
                return False
            if sk(n, dn, 1) == 'N':
                return False
            if sk(n, dn, 2) == 'N' and sk(n, dn, 1) in OFF_SET:
                return False
            return True

        extras = [n for n in nurses if n not in on_d and can_d2(n)
                  and d_cnt.get(n, 0) < D_ABS_MAX
                  and not fp_same_shift_conflict(n, dn, 'D')]     # ← 상한·함께근무불가
        _ti_x = _tie_break_map(extras, tie_rng)
        extras.sort(key=lambda n: (d_cnt[n], _ti_x.get(n, 0), n))
        need_extra = d_target - len(on_d)
        placed = 0
        for n in extras:
            if placed >= need_extra:
                break
            if fp_same_shift_conflict(n, dn, 'D'):
                continue
            if (n, dn) in req_locked:
                continue
            sched[n][dn] = 'D'
            d_cnt[n] += 1
            placed += 1

    # ★ 후처리 ⑤: D 재분배 – 초과(>목표) 간호사 → 미달(<4) 간호사로 교환
    #   n_high의 섬 D를 n_low의 인접 근무일로 이전 (양쪽 섬 방지)
    def is_d_island(n, dn):
        """해당 날이 n의 D 섬(OF-D-OF)인지"""
        p = sk(n, dn, 1)
        if p is None:
            p = 'OF'
        nxt = sched[n].get(dn + 1) if dn < NUM_DAYS else 'OF'
        return (p is None or p in REST_GAP) and (nxt is None or nxt in REST_GAP)

    changed_d = True
    while changed_d:
        changed_d = False
        for n_low in nurses:
            if d_cnt.get(n_low, 0) >= D_ABS_MIN:
                continue
            for d_idx in range(NUM_DAYS):
                if d_cnt.get(n_low, 0) >= D_ABS_MIN:
                    break
                dn = d_idx + 1
                if sched[n_low].get(dn) != 'OF':
                    continue
                # n_low 제약 확인
                if den_banned(n_low, 'D'):
                    continue
                if sk(n_low, dn, 1) == 'E':
                    continue
                if sk(n_low, dn, 1) == 'N':
                    continue
                if sk(n_low, dn, 2) == 'N' and sk(n_low, dn, 1) in OFF_SET:
                    continue
                if streak_total(n_low, d_idx) > 5:
                    continue
                # n_low에게 D를 주면 섬(OF-D-OF)이 되는지 확인
                prev_nl = sk(n_low, dn, 1)
                if prev_nl is None:
                    prev_nl = 'OF'
                next_nl = sched[n_low].get(dn + 1) if d_idx < NUM_DAYS - 1 else 'OF'
                if (prev_nl is None or prev_nl in REST_GAP) and \
                   (next_nl is None or next_nl in REST_GAP):
                    continue  # 섬 생성 → 스킵
                # 이 날 D를 하는 간호사 중 양도 가능자 찾기
                # 우선순위: 섬 D인 n_high → 비섬 n_high
                d_nurses = [m for m in nurses
                            if sched[m].get(dn) == 'D' and m != n_low
                            and d_cnt.get(m, 0) > D_ABS_MIN]
                _ti_dn = _tie_break_map(d_nurses, tie_rng)
                d_nurses.sort(
                    key=lambda m: (0 if is_d_island(m, dn) else 1, -d_cnt.get(m, 0), _ti_dn.get(m, 0), m)
                )
                for n_high in d_nurses:
                    if (n_high, dn) in req_locked or (n_low, dn) in req_locked:
                        continue
                    sched[n_high][dn] = 'OF'
                    d_cnt[n_high] = d_cnt.get(n_high, 0) - 1
                    sched[n_low][dn] = 'D'
                    d_cnt[n_low] = d_cnt.get(n_low, 0) + 1
                    changed_d = True
                    break

    # ★ 후처리 ⑥: 재분배 후 새로 생긴 섬 재정리
    #   ①과 같은 논리를 한 번 더 적용
    changed2 = True
    while changed2:
        changed2 = False
        for n in nurses:
            for d in range(1, NUM_DAYS - 1):
                dn = d + 1
                prev_s = sched[n].get(dn - 1)
                curr_s = sched[n].get(dn)
                next_s = sched[n].get(dn + 1)
                if prev_s not in REST_GAP or next_s not in REST_GAP:
                    continue
                if curr_s in REST_GAP:
                    continue
                if curr_s == 'D':
                    day = days[d]
                    ha1 = sched[0].get(dn) == 'A1'
                    d_min2 = d_assignment_target(num_nurses, day, ha1)
                    others_d = sum(1 for m in nurses if m != n and sched[m].get(dn) == 'D')
                    fixed2 = False
                    if others_d >= d_min2 and d_cnt.get(n, 0) > D_ABS_MIN and (n, dn) not in req_locked:
                        # 인원 충분 + n이 최솟값 초과 → OF로 전환
                        sched[n][dn] = 'OF'
                        d_cnt[n] = d_cnt.get(n, 0) - 1
                        changed2 = True
                        fixed2 = True
                    if not fixed2 and d_cnt.get(n, 0) > D_ABS_MIN:
                        # 스왑으로 섬 해소 (인원 유지하면서 다른 간호사와 교환)
                        for m in nurses:
                            if m == n or sched[m].get(dn) != 'OF':
                                continue
                            m_prev = sk(m, dn, 1)
                            m_next = sched[m].get(dn + 1)
                            if (m_prev is None or m_prev in REST_GAP) and m_next in REST_GAP:
                                continue
                            if sk(m, dn, 1) == 'E':
                                continue
                            if sk(m, dn, 1) == 'N':
                                continue
                            if sk(m, dn, 2) == 'N' and sk(m, dn, 1) in OFF_SET:
                                continue
                            if den_banned(m, 'D'):
                                continue
                            if streak_total(m, d) > 5:
                                continue
                            if (n, dn) in req_locked or (m, dn) in req_locked:
                                continue
                            sched[n][dn] = 'OF'
                            d_cnt[n] = d_cnt.get(n, 0) - 1
                            sched[m][dn] = 'D'
                            d_cnt[m] = d_cnt.get(m, 0) + 1
                            changed2 = True
                            fixed2 = True
                            break
                    if not fixed2:
                        # 섬 유지 불가피 → 인접 날로 D 연장하여 섬 해소
                        for ext_d, ext_dn in [(d + 1, dn + 1), (d - 1, dn - 1)]:
                            if ext_d < 0 or ext_d >= NUM_DAYS:
                                continue
                            if sched[n].get(ext_dn) not in ('OF', 'OH'):
                                continue
                            day_ext = days[ext_d]
                            ha1_ext = sched[0].get(ext_dn) == 'A1'
                            d_min_ext = d_assignment_target(num_nurses, day_ext, ha1_ext)
                            d_on_ext = sum(1 for m in nurses if sched[m].get(ext_dn) == 'D')
                            if d_on_ext >= d_min_ext:
                                continue
                            if sk(n, ext_dn, 1) == 'E':
                                continue
                            if sk(n, ext_dn, 1) == 'N':
                                continue
                            if sk(n, ext_dn, 2) == 'N' and sk(n, ext_dn, 1) in OFF_SET:
                                continue
                            if den_banned(n, 'D'):
                                continue
                            if streak_total(n, ext_d) > 5:
                                continue
                            if (n, ext_dn) in req_locked:
                                continue
                            sched[n][ext_dn] = 'D'
                            d_cnt[n] = d_cnt.get(n, 0) + 1
                            changed2 = True
                            break
                elif curr_s == 'E':
                    others_e = sum(1 for m in nurses if m != n and sched[m].get(dn) == 'E')
                    if others_e >= 2 and (n, dn) not in req_locked:
                        sched[n][dn] = 'OF'
                        e_cnt[n] = e_cnt.get(n, 0) - 1
                        changed2 = True

    # ★ 최종: OF 쿼터 적용 ─────────────────────────────────────────────────────
    # 원칙: 수간호사 OFF(토·일=OF + 공휴일=OH) 합산 수 = 일반 간호사 OFF 쿼터
    # 주(月~日) 주 2 휴무: OF·OH·NO 조합만 인정(연차는 대체 불가). 일일 D/E/N 정원은 연차보다 우선.
    of_quota_month = sum(1 for day in days if day['is_weekend'] or day['is_holiday'])

    # 주 구조 계산 (월~일 기준)
    week_map: dict[int, int]       = {}  # day_num -> week_key
    week_days_map: dict[int, list] = {}  # week_key -> [day_nums]
    for day in days:
        dt  = day['date']
        mon = dt - timedelta(days=dt.weekday())   # 그 주 월요일
        key = mon.toordinal()
        week_map[day['day']] = key
        week_days_map.setdefault(key, []).append(day['day'])

    month_first = date(YEAR, MONTH, 1)

    # ── 주간 OFF(OF+OH+NO) 합산 — 연차 전환 판단용
    def _wk_off(n, key):
        return sum(
            1 for d2 in week_days_map[key]
            if sched[n].get(d2) in ('OF', 'OH', 'NO')
        )

    def _wk_of_only(n, key):
        """주간 'OF' 문자만 카운트 (월~일 범위 중 당월에 포함된 날)."""
        return sum(1 for d2 in week_days_map[key] if sched[n].get(d2) == 'OF')

    def _wk_oh_no(n, key):
        oh = sum(1 for d2 in week_days_map[key] if sched[n].get(d2) == 'OH')
        no = sum(1 for d2 in week_days_map[key] if sched[n].get(d2) == 'NO')
        return oh, no

    # ── ⓪ 월~일: 주 2 OF 인정 — OF×2 | OH×2 | OF+OH | OF+NO | OH+NO 충족까지 보강
    # 전월 동주 carry 합산(전월 말주~당월 첫 주 경계)
    for n in nurses:
        for _wkey, _wdays in week_days_map.items():
            if not _wdays:
                continue
            mon_date = date.fromordinal(_wkey)
            pre_of, pre_oh, pre_no, n_prev = _carry_week_prev_month_off_counts(
                carry, n, mon_date, month_first,
            )
            post_of, post_oh, post_no, n_next = _carry_week_next_month_off_counts(
                carry_next, n, mon_date, month_last,
            )
            _wk_iters = 28 if num_nurses >= 11 else 14
            for _ in range(_wk_iters):
                of_wk = _wk_of_only(n, _wkey)
                oh_wk, no_wk = _wk_oh_no(n, _wkey)
                of_vis = pre_of + of_wk
                oh_vis = pre_oh + oh_wk
                no_vis = pre_no + no_wk
                if _weekly_off_rule_met(
                    of_vis, oh_vis, no_vis, n_prev, len(_wdays),
                    post_of, post_oh, post_no, n_next, carry_next_provided,
                ):
                    break
                _hit = False
                # 주간 2휴무는 OF/OH/NO 조합만 — 연은 인정되지 않으므로 필요 시 연→OF
                for _d in sorted(_wdays, reverse=True):
                    if sched[n].get(_d) != '연':
                        continue
                    if (n, _d) in req_locked:
                        continue
                    sched[n][_d] = 'OF'
                    _hit = True
                    break
                if _hit:
                    continue
                for _lc in ('EDU', '공', '병', '경'):
                    for _d in sorted(_wdays, reverse=True):
                        if sched[n].get(_d) != _lc:
                            continue
                        if (n, _d) in req_locked:
                            continue
                        sched[n][_d] = 'OF'
                        _hit = True
                        break
                    if _hit:
                        break
                if _hit:
                    continue
                for _d in sorted(_wdays, reverse=True):
                    if sched[n].get(_d) != 'OH':
                        continue
                    if (n, _d) in req_locked:
                        continue
                    if sk(n, _d, 1) == 'N':
                        _need_tail = 'OH' if days[_d - 1]['is_holiday'] else 'OF'
                        if _need_tail == 'OH':
                            continue
                    sched[n][_d] = 'OF'
                    _hit = True
                    break
                if not _hit:
                    break

    for n in nurses:
        # ── ①  초과 OFF → 가능하면 D, 아니면 연차 (월 쿼터). 전환 후에도 주 OF 인정 규칙 유지
        nurse_offs_total = sum(1 for s in sched[n].values() if s in ('OF', 'OH', 'NO'))
        surplus = nurse_offs_total - of_quota_month
        nurse_ofs = sorted((dn for dn, s in sched[n].items() if s == 'OF'), reverse=True)
        for dn in nurse_ofs:
            if surplus <= 0:
                break
            if sk(n, dn, 1) == 'N':
                continue
            wkey = week_map.get(dn)
            if not wkey:
                continue
            _wd = week_days_map[wkey]
            mon_date = date.fromordinal(wkey)
            pre_of, pre_oh, pre_no, n_prev = _carry_week_prev_month_off_counts(
                carry, n, mon_date, month_first,
            )
            post_of, post_oh, post_no, n_next = _carry_week_next_month_off_counts(
                carry_next, n, mon_date, month_last,
            )
            of_c_m = _wk_of_only(n, wkey)
            oh_c_m, no_c_m = _wk_oh_no(n, wkey)
            of_vis = pre_of + of_c_m
            oh_vis = pre_oh + oh_c_m
            no_vis = pre_no + no_c_m
            merged_week_off = pre_of + pre_oh + pre_no + sum(
                1 for d2 in _wd if sched[n].get(d2) in ('OF', 'OH', 'NO')
            )
            if merged_week_off <= 2:
                continue
            if not _weekly_off_ok_after_of_to_yun(
                of_vis, oh_vis, no_vis, n_prev, _wd,
                post_of, post_oh, post_no, n_next, carry_next_provided,
            ):
                continue
            if (n, dn) in req_locked:
                continue
            day_obj = days[dn - 1]
            head_a1 = sched[0].get(dn) == 'A1'
            d_max = d_slots_per_day(num_nurses, day_obj, head_a1)
            d_on = sum(1 for m in nurses if sched[m].get(dn) == 'D')
            d_idx = dn - 1
            off_one = frozenset({'OF', 'OH', 'NO'})
            can_surplus_d = (
                d_on < d_max
                and not den_banned(n, 'D')
                and sk(n, dn, 1) != 'E'
                and sk(n, dn, 1) != 'N'
                and not (sk(n, dn, 2) == 'N' and sk(n, dn, 1) in off_one)
                and streak_total(n, d_idx) <= 5
                and not fp_same_shift_conflict(n, dn, 'D')
                and d_cnt.get(n, 0) < D_ABS_MAX
            )
            if can_surplus_d:
                sched[n][dn] = 'D'
                d_cnt[n] = d_cnt.get(n, 0) + 1
            else:
                sched[n][dn] = '연'
            surplus -= 1

    _repair_yun_to_of_for_weekly_rule(
        sched, nurses, week_days_map, carry, carry_next,
        month_first, month_last, carry_next_provided, req_locked,
    )

    _convert_non_request_yun_to_d(
        sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
    )
    _auto_yun_to_of_if_quota_room(sched, num_nurses, holidays, carry_in, requests)

    _cap_auto_rest_in_tail_last_week(
        sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
    )

    if fp_map:
        _repair_fp_same_shift_conflicts(
            sched, nurses, fp_map, days, carry, req_locked=req_locked,
            shift_bans=shift_bans,
        )

    _ensure_validation_d_floor(
        sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
        tie_rng=tie_rng, carry_next_month=carry_next_month,
    )
    _auto_yun_to_of_if_quota_room(sched, num_nurses, holidays, carry_in, requests)

    _repair_schedule_validate_errors(
        sched, num_nurses, holidays, forbidden_pairs, carry_in, requests, nurses, tie_rng,
        carry_next_month=carry_next_month, shift_bans=shift_bans,
    )

    if tie_rng is not None:
        _refine_schedule_regenerate(
            sched, requests, num_nurses, holidays, forbidden_pairs, carry_in, nurse_names, tie_rng,
            max_tries=(1800 if regenerate else 280),
            carry_next_month=carry_next_month,
            shift_bans=shift_bans,
        )
        if regenerate:
            _repair_schedule_validate_errors(
                sched, num_nurses, holidays, forbidden_pairs, carry_in, requests, nurses, tie_rng,
                carry_next_month=carry_next_month, shift_bans=shift_bans,
            )
            _refine_schedule_regenerate(
                sched, requests, num_nurses, holidays, forbidden_pairs, carry_in, nurse_names, tie_rng,
                max_tries=950,
                carry_next_month=carry_next_month,
                shift_bans=shift_bans,
            )

    _reapply_requests_to_schedule(sched, requests, num_nurses)
    _post_reapply_fix_n_cap_and_daily_two(
        sched, num_nurses, holidays, requests, forbidden_pairs, carry_in, shift_bans,
    )
    _convert_non_request_yun_to_d(
        sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
    )
    _auto_yun_to_of_if_quota_room(sched, num_nurses, holidays, carry_in, requests)
    _cap_auto_rest_in_tail_last_week(
        sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
    )
    _ensure_validation_d_floor(
        sched, num_nurses, holidays, carry_in, requests, forbidden_pairs, shift_bans,
        tie_rng=tie_rng, carry_next_month=carry_next_month,
    )
    _auto_yun_to_of_if_quota_room(sched, num_nurses, holidays, carry_in, requests)
    return sched, True, 'FEASIBLE'


def _repair_fp_same_shift_conflicts(sched, nurses, fp_map, days, carry=None,
                                     req_locked=None, shift_bans=None):
    """
    같은 날 D/E/N에 함께 근무 불가 쌍이 남아 있으면,
    해당 날 OF/OH인 다른 간호사와 자리 교환 시도 (당일 시프트 인원 수 유지).
    fp_map: (i,j) -> 적용 시프트 부분집합 (수간호사 0 포함)
    """
    if not fp_map:
        return
    req_locked = req_locked or set()
    carry = _normalize_carry_in(carry, len(sched))
    den_bans = _normalize_shift_bans(shift_bans, len(sched))

    def den_banned(n, sh):
        if sh not in ('D', 'E', 'N'):
            return False
        b = den_bans.get(n)
        return bool(b and sh in b)

    sk = lambda n, dn, k: _shift_k_days_before(sched, carry, n, dn, k)
    OFF_OK = {'OF', 'OH', 'NO'}
    SHIFTS = ('N', 'E', 'D')
    staff_all = list(range(len(sched)))
    for _ in range(150):
        found = None
        for day in days:
            dn = day['day']
            for shift in SHIFTS:
                team = [n for n in staff_all if sched[n].get(dn) == shift]
                for i in range(len(team)):
                    for j in range(i + 1, len(team)):
                        a, b = team[i], team[j]
                        key = (min(a, b), max(a, b))
                        if key in fp_map and shift in fp_map[key]:
                            found = (dn, shift, a, b)
                            break
                    if found:
                        break
                if found:
                    break
            if found:
                break
        if not found:
            return
        dn, shift, a, b = found
        swapped = False
        for victim in (a, b):
            for c in staff_all:
                if c == victim:
                    continue
                cur = sched[c].get(dn)
                if cur not in OFF_OK:
                    continue
                bad = False
                for m in staff_all:
                    if m in (victim, c):
                        continue
                    if sched[m].get(dn) != shift:
                        continue
                    ck = (min(c, m), max(c, m))
                    if ck in fp_map and shift in fp_map[ck]:
                        bad = True
                        break
                if bad:
                    continue
                if shift == 'D' and sk(c, dn, 1) == 'N':
                    continue
                if den_banned(c, shift):
                    continue
                if (victim, dn) in req_locked or (c, dn) in req_locked:
                    continue
                sched[victim][dn] = cur
                sched[c][dn] = shift
                swapped = True
                break
            if swapped:
                break
        if not swapped:
            return


def validate_schedule(schedule, num_nurses, holidays=(), forbidden_pairs=None,
                      nurse_names=None, carry_in=None, requests=None, carry_next_month=None,
                      shift_bans=None):
    """
    생성된 스케줄을 규칙에 따라 검증하고 위반 사항 목록을 반환한다.
    forbidden_pairs: [(i,j), ...] 또는 [(i,j,['D','E']), ...] — 같은 날 동시 배치 금지(수간 0 포함)
    nurse_names: 표시용 이름 (없으면 기본 수간호사/간호사1…)
    carry_in: (선택) 전월 말 근무 꼬리 — 월 경계 규칙 검증용
    carry_next_month: (선택) 차월 초 근무 — 마지막주(당월 말~차월 일요일) 주 2 OF 합산 검증용
    requests: (선택) 생성 시 사용한 신청 — 있으면 스케줄 셀과 반드시 일치해야 함
    shift_bans: (선택) dict[int,str] — 간호사 인덱스별 근무불가(d_only|no_d|no_e|no_n)
    Returns: list of dict  { 'level': 'error'|'warn', 'msg': str }
    """
    issues = []
    days = get_april_days(holidays)
    _dn_holiday = {d['day']: bool(d['is_holiday']) for d in days}
    nurses = list(range(1, num_nurses))
    names = nurse_names if nurse_names is not None else get_nurse_names(num_nurses)
    fp_map = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    den_bans = _normalize_shift_bans(shift_bans, num_nurses)
    carry = _normalize_carry_in(carry_in, num_nurses)
    carry_next = _normalize_carry_in(carry_next_month, num_nurses) if carry_next_month else {}
    carry_next_provided = carry_next_month is not None
    month_last = date(YEAR, MONTH, NUM_DAYS)
    OFF_SET = {'OF', 'OH', 'NO'}
    REST_GAP = frozenset(OFF_SET | {'연'})
    GAP_WORK = frozenset({'D', 'E', 'N', 'EDU', '공', '병', '경', 'A1'})

    def sh(n, dn):
        return schedule.get(n, {}).get(dn, '')

    def vk(n, dn, k):
        return _shift_k_days_before(schedule, carry, n, dn, k)

    def is_off(s):
        return s in OFF_SET or s in ('', None)

    def err(msg):
        issues.append({'level': 'error', 'msg': msg})

    def warn(msg):
        issues.append({'level': 'warn',  'msg': msg})

    if requests:
        for ni, ds in requests.items():
            try:
                ni = int(ni)
            except (TypeError, ValueError):
                continue
            if not (0 <= ni < num_nurses):
                continue
            nm = names[ni]
            for dn, req_shift in (ds or {}).items():
                try:
                    dni = int(dn)
                except (TypeError, ValueError):
                    continue
                if not (1 <= dni <= NUM_DAYS):
                    continue
                cur = schedule.get(ni, {}).get(dni, '')
                if cur != req_shift:
                    err(
                        f"{nm} 신청 불일치(절대): {dni}일 신청={req_shift!r} "
                        f"근무표={cur!r}"
                    )

    if den_bans:
        _ban_label = {
            frozenset({'E', 'N'}): 'D근무만 가능(E·N 불가)',
            frozenset({'D'}): 'D근무 불가',
            frozenset({'E'}): 'E근무 불가',
            frozenset({'N'}): 'N근무 불가',
        }
        for n in range(num_nurses):
            b = den_bans.get(n)
            if not b:
                continue
            nm = names[n]
            blab = _ban_label.get(b, ','.join(sorted(b)) + ' 불가')
            for dn in range(1, NUM_DAYS + 1):
                s = sh(n, dn)
                if s in b:
                    err(f"{nm} 근무불가 위반 ({blab}): {dn}일 {s}")

    # ── ① 일일 인력 요구 ────────────────────────────────────────────────────
    for day in days:
        dn = day['day']
        label = f"{dn}일({day['weekday_name']})"
        head = sh(0, dn)
        is_we = day['is_weekend'] or day['is_holiday']
        d_cnt = sum(1 for n in nurses if sh(n, dn) == 'D')
        e_cnt = sum(1 for n in nurses if sh(n, dn) == 'E')
        n_cnt = sum(1 for n in nurses if sh(n, dn) == 'N')

        # N은 절대 규칙 — 평일/주말/공휴일 구분 없이 반드시 2명
        if n_cnt < 2:
            err(f"{label} 🚨 N 절대 부족: {n_cnt}명 (매일 반드시 2명 필요)")
        if n_cnt > 2:
            warn(f"{label} N 인원 초과: {n_cnt}명 (최대 2명)")

        if is_we:
            tag = '[주말/공휴일]'
            if d_cnt < 2: err(f"{label} {tag} D 인원 부족: {d_cnt}명 → 필요 2명")
            if e_cnt < 2: err(f"{label} {tag} E 인원 부족: {e_cnt}명 → 필요 2명")
        else:
            tag = '[평일]'
            # 주말·공휴 D=2는 위 is_we 분기. 평일: 2~3명 운영 지향(소인원·수간 A1만 최소 1)
            if num_nurses >= 11:
                req_d = 2
            elif head != 'A1':
                req_d = 2
            else:
                req_d = 1 if num_nurses <= 7 else 2
            if d_cnt < req_d: err(f"{label} {tag} D 인원 부족: {d_cnt}명 → 필요 {req_d}명")
            if e_cnt < 2:     err(f"{label} {tag} E 인원 부족: {e_cnt}명 → 필요 2명")
            # 수간 포함 11명 이상·평일·수간 A1: D 상한 3명 — 초과 시 경고
            if head == 'A1' and num_nurses >= 11 and d_cnt > 3:
                warn(f"{label} {tag} D 인원 초과: {d_cnt}명 (11명 이상 평일 최대 3명)")

        if e_cnt > 2: warn(f"{label} E 인원 초과: {e_cnt}명 (최대 2명)")

    # ── ①b 함께 근무 불가 (선택한 시프트에 한해 같은 날 동시 배치 금지, 수간 포함) ─
    if fp_map:
        all_idx = list(range(num_nurses))
        for day in days:
            dn = day['day']
            label = f"{dn}일({day['weekday_name']})"
            for shift in ('D', 'E', 'N'):
                team = [n for n in all_idx if sh(n, dn) == shift]
                for i in range(len(team)):
                    for j in range(i + 1, len(team)):
                        a, b = team[i], team[j]
                        key = (min(a, b), max(a, b))
                        if key in fp_map and shift in fp_map[key]:
                            err(
                                f"{label} 함께 근무 불가: {names[a]} · {names[b]} "
                                f"동시 {shift}"
                            )

    # ── ② 개인별 규칙 ───────────────────────────────────────────────────────
    for n in nurses:
        ns   = schedule.get(n, {})
        nm   = names[n]

        # N 총 개수
        n_total = sum(1 for v in ns.values() if v == 'N')
        if n_total > N_ABS_MAX:
            err(f"{nm} N 초과: {n_total}개 (최대 {N_ABS_MAX}개)")

        # N 블록 분석
        n_days = sorted(d for d, s in ns.items() if s == 'N')
        blocks = []
        if n_days:
            blk = [n_days[0]]
            for d in n_days[1:]:
                if d == blk[-1] + 1:
                    blk.append(d)
                else:
                    blocks.append(blk); blk = [d]
            blocks.append(blk)

        for blk in blocks:
            if len(blk) < 2:
                # 말일 단독 N: 3-3-1 / 2-3-1 / 3-2-1 등 — 당월 마지막 날(NUM_DAYS)만 (31일 말달은 31일)
                if blk[0] != NUM_DAYS:
                    err(
                        f"{nm} N 블록 단독: {blk[0]}일 "
                        f"(1개, 당월 말일({NUM_DAYS}일)만 단독 허용 — 3-3-1·2-3-1·3-2-1)"
                    )
            elif len(blk) > 3:
                err(f"{nm} N 블록 초과: {blk[0]}~{blk[-1]}일 ({len(blk)}개, 최대 3개)")

        for i in range(len(blocks) - 1):
            gap = blocks[i+1][0] - blocks[i][-1] - 1
            if gap < 7:
                warn(f"{nm} N 블록 간격 부족: {blocks[i][-1]}일→{blocks[i+1][0]}일 ({gap}일, 최소 7일)")

        # 전월 말 N → 당월 1일(연속 N 아님): 공휴 OH / 평일 OF
        cseq = list(carry.get(n, ()))
        if cseq and cseq[-1] == 'N':
            s_first = sh(n, 1)
            if s_first != 'N':
                need0 = 'OH' if days[0]['is_holiday'] else 'OF'
                if s_first != need0:
                    err(
                        f"{nm} N블록 직후 휴무 위반: 전월 말 N 이후 1일 "
                        f"({s_first}, 필요 {need0})"
                    )

        # N블록 말 직후(당월): 공휴이면 OH, 아니면 OF
        for blk in blocks:
            end = blk[-1]
            if end >= NUM_DAYS:
                continue
            s1 = sh(n, end + 1)
            need = 'OH' if _dn_holiday.get(end + 1) else 'OF'
            if s1 != need:
                err(
                    f"{nm} N블록 직후 휴무 위반: {end}일 N 다음 {end+1}일 "
                    f"({s1 or '빈칸'}, 필요 {need})"
                )

        for blk in blocks:
            end = blk[-1]
            if end >= NUM_DAYS - 1:
                continue
            s1 = sh(n, end + 1)
            s2 = sh(n, end + 2)
            if s1 in ('OF', 'OH') and s2 == 'D':
                err(f"{nm} N-휴무-D 금지: {end}일N→{end+1}일{s1}→{end+2}일D")
            if s1 in ('OF', 'OH') and s2 == 'EDU':
                err(f"{nm} N-휴무-교육 금지: {end}일N→{end+1}일{s1}→{end+2}일EDU")

        # E-D 금지 (전월 말 E → 당월 1일 D 포함)
        for dn in range(1, NUM_DAYS + 1):
            if vk(n, dn, 1) == 'E' and sh(n, dn) == 'D':
                err(f"{nm} E-D 금지: 전일E→{dn}일D")

        # N-D 금지 (전날 야간 직후 데이 — 절대 불가, 전월 말 N 포함)
        for dn in range(1, NUM_DAYS + 1):
            if vk(n, dn, 1) == 'N' and sh(n, dn) == 'D':
                err(f"{nm} N-D 금지: 전일N→{dn}일D")

        # 연속 근무 최대 5일 (전월 꼬리 + 당월) — D/E/N/공/EDU만 합산(연차 등은 끊김)
        seq = list(carry.get(n, ())) + [sh(n, d) for d in range(1, NUM_DAYS + 1)]
        streak = 0
        for s in seq:
            if s in STREAK_WORK_SHIFTS:
                streak += 1
                if streak > 5:
                    err(f"{nm} 연속근무 초과: 전월이월·당월 합산 {streak}일 (최대 5일)")
            else:
                streak = 0

        # 쉬는 날(OF/OH/NO/연) 사이 근무: 0일(붙은 휴무) OK, 1일은 섬 경고,
        # 2~5일만 허용. D/E/N/공/EDU만 근무일로 합산 — 5일 초과는 절대 오류.
        gap_anchors = sorted(d for d, s in ns.items() if s in REST_GAP)
        prev_a = None
        for od in gap_anchors:
            if prev_a is not None:
                work_btw = sum(
                    1 for d in range(prev_a + 1, od)
                    if sh(n, d) in STREAK_WORK_SHIFTS
                )
                la = sh(n, prev_a) or "?"
                ra = sh(n, od) or "?"
                if work_btw == 1:
                    warn(
                        f"{nm} 쉬는 날 사이 근무 1일(섬): {prev_a}일{la}~{od}일{ra} "
                        f"— 0일 또는 2~5일이어야 함"
                    )
                elif work_btw > 5:
                    err(
                        f"{nm} 쉬는 날 사이 근무 과다(절대): {prev_a}일{la}~{od}일{ra} "
                        f"사이 근무 {work_btw}일 — 최대 5일, 공가·교육 포함"
                    )
            prev_a = od

        # OF 쿼터 검증: 수간호사 기준(토·일 + 공휴일) 합산과 비교
        of_quota = sum(1 for day in days if day['is_weekend'] or day['is_holiday'])
        off_total = sum(1 for v in ns.values() if v in ('OF', 'OH', 'NO'))
        if off_total > of_quota:
            warn(f"{nm} OFF 초과: {off_total}개 (수간호사 기준 최대 {of_quota}개, 초과분은 연차 권장)")

        # 주(월~일): OF×2·OH×2 또는 OF+OH / OF+NO / OH+NO, NO는 주당 최대 1 — 절대
        # 전월 말주~당월 첫 주가 이어지는 경우, 전월 동주 carry의 OF/OH/NO를 합산해 판정
        if days:
            wk_map: dict[int, list] = {}
            for day in days:
                dt  = day['date']
                mon = dt - timedelta(days=dt.weekday())
                wk_map.setdefault(mon.toordinal(), []).append(day['day'])
            month_first = date(YEAR, MONTH, 1)
            for _wk, wdays in wk_map.items():
                if not wdays:
                    continue
                mon_date = date.fromordinal(_wk)
                pre_of, pre_oh, pre_no, n_prev = _carry_week_prev_month_off_counts(
                    carry, n, mon_date, month_first,
                )
                post_of, post_oh, post_no, n_next = _carry_week_next_month_off_counts(
                    carry_next, n, mon_date, month_last,
                )
                of_vis = pre_of + sum(1 for d2 in wdays if sh(n, d2) == 'OF')
                oh_vis = pre_oh + sum(1 for d2 in wdays if sh(n, d2) == 'OH')
                no_vis = pre_no + sum(1 for d2 in wdays if sh(n, d2) == 'NO')
                no_week_total = no_vis + post_no
                m = n_prev + len(wdays) + n_next
                d_range = f"{min(wdays)}~{max(wdays)}일"
                if n_prev:
                    d_range = (
                        f"{d_range} (월~일 주에 전월 {n_prev}일 + 당월 {len(wdays)}일, carry 합산)"
                    )
                if n_next > 0:
                    d_range = f"{d_range} · 말주~차월 일요일 동일 주(차월 {n_next}일)"
                if no_week_total > 1:
                    err(
                        f"{nm} 주간 NO 초과(절대): {d_range} — "
                        f"NO는 같은 주에 최대 1개만 가능 (현재 {no_week_total}개)"
                    )
                if not _weekly_off_rule_met(
                    of_vis, oh_vis, no_vis, n_prev, len(wdays),
                    post_of, post_oh, post_no, n_next, carry_next_provided,
                ):
                    err(
                        f"{nm} 주간 휴무(OF/OH/NO 인정) 부족(절대): {d_range} — "
                        f"OF{of_vis + post_of} OH{oh_vis + post_oh} NO{no_vis + post_no} "
                        f"(평가 {m}일·연차 제외, OF≥2 또는 OH≥2 또는 OF+OH 또는 OF+NO 또는 OH+NO 필요)"
                    )

    return issues


def build_stats(schedule, num_nurses):
    """간호사별 시프트 통계 및 날짜별 인원 통계 계산"""
    nurse_stats = {}
    for n in range(num_nurses):
        shifts = schedule.get(n, {})
        nurse_stats[n] = {
            'N':  sum(1 for v in shifts.values() if v == 'N'),
            'D':  sum(1 for v in shifts.values() if v == 'D'),
            'E':  sum(1 for v in shifts.values() if v == 'E'),
            'OF': sum(1 for v in shifts.values() if v in ('OF', 'OH')),
            'NO': sum(1 for v in shifts.values() if v == 'NO'),
            'A1': sum(1 for v in shifts.values() if v == 'A1'),
        }

    day_stats = {}
    for d in range(1, NUM_DAYS + 1):
        day_stats[d] = {
            'D': sum(1 for n in range(num_nurses) if schedule.get(n, {}).get(d) == 'D'),
            'E': sum(1 for n in range(num_nurses) if schedule.get(n, {}).get(d) == 'E'),
            'N': sum(1 for n in range(num_nurses) if schedule.get(n, {}).get(d) == 'N'),
        }
    return nurse_stats, day_stats


# ── Flask 라우트 ───────────────────────────────────────────────────────────────
@app.route('/', methods=['GET'])
def index():
    days = get_april_days()
    return render_template(
        'index.html',
        days=days,
        shift_names=SHIFT_NAMES,
        shift_colors=SHIFT_COLORS,
        shift_text_colors=SHIFT_TEXT_COLORS,
        num_nurses=10,
        schedule=None,
        nurse_names=get_nurse_names(10),
        holidays=[],
    )


@app.route('/generate', methods=['POST'])
def generate():
    global _last_result

    num_nurses = int(request.form.get('num_nurses', 10))
    num_nurses = max(9, min(25, num_nurses))

    # 공휴일 파싱
    holidays_str = request.form.get('holidays', '').strip()
    holidays = []
    if holidays_str:
        for h in holidays_str.split(','):
            try:
                hday = int(h.strip())
                if 1 <= hday <= NUM_DAYS:
                    holidays.append(hday)
            except ValueError:
                pass

    # 개인 신청 근무 파싱 (숨겨진 입력 필드: req_{n_idx}_{day_num})
    requests = {}
    for key, val in request.form.items():
        if key.startswith('req_') and val.strip() in SHIFT_NAMES:
            parts = key.split('_')
            if len(parts) == 3:
                try:
                    n_idx = int(parts[1])
                    day_num = int(parts[2])
                    if 0 <= n_idx < num_nurses and 1 <= day_num <= NUM_DAYS:
                        requests.setdefault(n_idx, {})[day_num] = val.strip()
                except ValueError:
                    pass

    try:
        schedule, success, status_str = solve_schedule(num_nurses, requests, holidays)
    except Exception as e:
        schedule, success, status_str = None, False, f'예외 발생: {e}'

    days = get_april_days(holidays)
    nurse_names = get_nurse_names(num_nurses)

    if success:
        nurse_stats, day_stats = build_stats(schedule, num_nurses)
        _last_result = {
            'schedule': schedule,
            'num_nurses': num_nurses,
            'holidays': holidays,
            'nurse_names': nurse_names,
        }
        return render_template(
            'index.html',
            days=days,
            shift_names=SHIFT_NAMES,
            shift_colors=SHIFT_COLORS,
            shift_text_colors=SHIFT_TEXT_COLORS,
            num_nurses=num_nurses,
            schedule=schedule,
            nurse_names=nurse_names,
            nurse_stats=nurse_stats,
            day_stats=day_stats,
            status=status_str,
            holidays=holidays,
            error=None,
        )
    else:
        return render_template(
            'index.html',
            days=days,
            shift_names=SHIFT_NAMES,
            shift_colors=SHIFT_COLORS,
            shift_text_colors=SHIFT_TEXT_COLORS,
            num_nurses=num_nurses,
            schedule=None,
            nurse_names=nurse_names,
            holidays=holidays,
            error=(
                f'해결책을 찾지 못했습니다 ({status_str}). '
                '개인 신청 근무 내용을 줄이거나, 간호사 수를 조정 후 다시 시도해주세요.'
            ),
        )


@app.route('/download')
def download():
    global _last_result
    if not _last_result:
        return redirect(url_for('index'))

    schedule = _last_result['schedule']
    num_nurses = _last_result['num_nurses']
    holidays = _last_result['holidays']
    nurse_names = _last_result['nurse_names']
    days = get_april_days(holidays)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '2026년 4월 근무표'

    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    def _xrgb(h: str) -> str:
        return h.replace('#', '').upper()

    EXCEL_BG = {
        sk: (_xrgb(SHIFT_COLORS[sk]), _xrgb(SHIFT_TEXT_COLORS[sk]))
        for sk in SHIFT_COLORS
    }

    N_COL = NUM_DAYS + 2
    OF_COL = NUM_DAYS + 3
    D_COL = NUM_DAYS + 4

    # ─ 타이틀 행
    ws.merge_cells(f'A1:{get_column_letter(D_COL)}1')
    tc = ws['A1']
    tc.value = '응급실 2026년 4월 근무표'
    tc.font = Font(bold=True, size=14, color=_xrgb(SHIFT_TEXT_COLORS['N']))
    tc.fill = PatternFill(
        start_color=_xrgb(SHIFT_COLORS['N']), end_color=_xrgb(SHIFT_COLORS['N']), fill_type='solid',
    )
    tc.alignment = center
    ws.row_dimensions[1].height = 28

    # ─ 헤더 행 (날짜)
    ws.cell(row=2, column=1, value='간호사')
    hdr = ws.cell(row=2, column=1)
    hdr.fill = PatternFill(
        start_color=_xrgb(SHIFT_COLORS['OF']), end_color=_xrgb(SHIFT_COLORS['OF']), fill_type='solid',
    )
    hdr.font = Font(bold=True, color=_xrgb(SHIFT_TEXT_COLORS['NO']), size=10)
    hdr.alignment = center
    hdr.border = thin
    ws.column_dimensions['A'].width = 10

    for d, day in enumerate(days):
        col = d + 2
        cell = ws.cell(row=2, column=col, value=f"{day['day']}\n{day['weekday_name']}")
        cell.alignment = center
        cell.border = thin
        if day['is_holiday']:
            bg, tfg = 'FFEBEE', 'C62828'
        elif day['is_weekend']:
            bg, tfg = 'E3F2FD', '1565C0'
        else:
            bg, tfg = 'F5F5F5', '455A64'
        cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
        cell.font = Font(bold=True, color=tfg, size=9)
        ws.column_dimensions[get_column_letter(col)].width = 4.5

    for col, label, sk in [(N_COL, 'N\n합계', 'N'), (OF_COL, 'OF\n합계', 'OF'), (D_COL, 'D\n합계', 'D')]:
        c = ws.cell(row=2, column=col, value=label)
        c.alignment = center
        c.border = thin
        c.fill = PatternFill(start_color=EXCEL_BG[sk][0], end_color=EXCEL_BG[sk][0], fill_type='solid')
        c.font = Font(bold=True, color=EXCEL_BG[sk][1], size=9)
        ws.column_dimensions[get_column_letter(col)].width = 5.5

    ws.row_dimensions[2].height = 28

    # ─ 간호사 행
    for n_idx, nurse_name in enumerate(nurse_names):
        row = n_idx + 3
        nc = ws.cell(row=row, column=1, value=nurse_name)
        nc.fill = PatternFill(
            start_color=_xrgb(SHIFT_COLORS['OF']), end_color=_xrgb(SHIFT_COLORS['OF']), fill_type='solid',
        )
        nc.font = Font(bold=True, color=_xrgb(SHIFT_TEXT_COLORS['NO']), size=9)
        nc.alignment = center
        nc.border = thin
        ws.row_dimensions[row].height = 18

        nurse_sched = schedule.get(n_idx, {})
        n_cnt = d_cnt = of_cnt = 0

        for d, day in enumerate(days):
            shift = nurse_sched.get(d + 1, '')
            col = d + 2
            cell = ws.cell(row=row, column=col, value=shift)
            cell.alignment = center
            cell.border = thin
            if shift in EXCEL_BG:
                bg, fg = EXCEL_BG[shift]
                cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
                cell.font = Font(color=fg, size=9, bold=True)
            if shift == 'N':
                n_cnt += 1
            if shift == 'D':
                d_cnt += 1
            if shift in ('OF', 'OH', 'NO'):
                of_cnt += 1

        for col, val, sk in [(N_COL, n_cnt, 'N'), (OF_COL, of_cnt, 'OF'), (D_COL, d_cnt, 'D')]:
            bg, fg = EXCEL_BG[sk]
            c = ws.cell(row=row, column=col, value=val)
            c.alignment = center
            c.border = thin
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
            c.font = Font(color=fg, bold=True, size=10)

    # ─ 일별 인원 합계 행
    summary_start = len(nurse_names) + 3
    for idx, (label, shift_key) in enumerate([
        ('D 인원', 'D'),
        ('E 인원', 'E'),
        ('N 인원', 'N'),
    ]):
        row = summary_start + idx
        lb, lf = EXCEL_BG[shift_key]
        lc = ws.cell(row=row, column=1, value=label)
        lc.fill = PatternFill(start_color=lb, end_color=lb, fill_type='solid')
        lc.font = Font(bold=True, color=lf, size=9)
        lc.alignment = center
        lc.border = thin
        ws.row_dimensions[row].height = 16

        for d in range(NUM_DAYS):
            day_num = d + 1
            cnt = sum(1 for n in range(num_nurses) if schedule.get(n, {}).get(day_num) == shift_key)
            cell = ws.cell(row=row, column=d + 2, value=cnt)
            cell.alignment = center
            cell.border = thin
            cell.fill = PatternFill(start_color=lb, end_color=lb, fill_type='solid')
            cell.font = Font(bold=True, color=lf, size=9)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='2026년_4월_근무표.xlsx',
    )


if __name__ == '__main__':
    print('=' * 55)
    print('  응급실 2026년 4월 근무표 생성기 시작!')
    print('  브라우저에서 http://127.0.0.1:5000 을 열어주세요')
    print('=' * 55)
    # threaded=True: 계산 중에도 서버가 다른 요청에 응답할 수 있도록 함
    app.run(debug=False, port=5000, threaded=True)