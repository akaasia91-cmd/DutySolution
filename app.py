# 응급실 2026년 4월 근무표 자동 생성기
# 순수 Python 그리디 스케줄러 (서버 충돌 없음)

from flask import Flask, render_template, request, send_file, redirect, url_for
from datetime import date, timedelta
import calendar as _calendar
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

app = Flask(__name__)

# ── 기본 상수 ──────────────────────────────────────────────────────────────────
YEAR, MONTH, NUM_DAYS = 2026, 4, 30


def set_period(year: int, month: int):
    """연도·월을 변경할 때 전역 상수를 갱신합니다."""
    global YEAR, MONTH, NUM_DAYS
    YEAR = year
    MONTH = month
    NUM_DAYS = _calendar.monthrange(year, month)[1]

SHIFT_NAMES = ['A1', 'D', 'E', 'N', 'OF', 'EDU', '연', '공', '병', '경', 'OH', 'NO']
A1_S, D_S, E_S, N_S, OF_S, EDU_S, YUN_S, GONG_S, BYUNG_S, GYUNG_S, OH_S, NO_S = range(12)

# 근무일수에 포함되는 시프트 (연속근무 5일 제한에 사용)
WORK_SHIFTS = [D_S, E_S, N_S, EDU_S, YUN_S, GONG_S, BYUNG_S, GYUNG_S]

# 화면 색상
SHIFT_COLORS = {
    'A1': '#4A90D9', 'D': '#FDD835', 'E': '#FF7043', 'N': '#283593',
    'OF': '#ECEFF1', 'EDU': '#66BB6A', '연': '#EC407A', '공': '#AB47BC',
    '병': '#EF5350', '경': '#26A69A', 'OH': '#FFA726',
    'NO': '#B0BEC5',   # N 20회 등 수기 휴무 (OF와 구분)
}
SHIFT_TEXT_COLORS = {
    'A1': '#fff', 'D': '#000', 'E': '#fff', 'N': '#fff',
    'OF': '#9E9E9E', 'EDU': '#fff', '연': '#fff', '공': '#fff',
    '병': '#fff', '경': '#fff', 'OH': '#fff', 'NO': '#263238',
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


# ── 스케줄 생성 (순수 Python 그리디) ─────────────────────────────────────────
def solve_schedule(num_nurses, requests, holidays=(), forbidden_pairs=None):
    """
    서버 충돌 없는 순수 Python 그리디 스케줄러
    num_nurses : 총 간호사 수 (0번=수간호사, 1..n-1=일반간호사)
    requests   : {nurse_idx: {day_num: shift_name}}
    holidays   : 공휴일 날짜 목록 (1-based)
    forbidden_pairs : (선택) 같은 날 같은 근무(D/E/N)에 동시 배치하면 안 되는 쌍
                      [(i,j), ...] — 수간호사(0) 제외, 일반간호사 인덱스만
    """
    try:
        return _greedy_schedule(num_nurses, requests, holidays, forbidden_pairs)
    except Exception as e:
        print(f'[오류] {e}')
        return None, False, str(e)


def _normalize_forbidden_pairs(forbidden_pairs, num_nurses):
    """(i,j) 튜플 집합으로 정규화. 수간호사(0) 포함 쌍은 제외."""
    fp_set = set()
    if not forbidden_pairs:
        return fp_set
    for pair in forbidden_pairs:
        if not pair or len(pair) != 2:
            continue
        try:
            a, b = int(pair[0]), int(pair[1])
        except (TypeError, ValueError):
            continue
        if 0 < a < num_nurses and 0 < b < num_nurses and a != b:
            fp_set.add((min(a, b), max(a, b)))
    return fp_set


def _greedy_schedule(num_nurses, requests, holidays=(), forbidden_pairs=None):
    days = get_april_days(holidays)
    nurses = list(range(1, num_nurses))   # 일반간호사 인덱스
    fp_set = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)

    OFF_SET = {'OF', 'OH', 'NO'}   # 오프 계열 (NO: N 20회 시 수기 휴무, 자동 배정 안 함)

    def is_off(shift):
        return shift in OFF_SET or shift is None

    def work_streak_before(n, d):
        """d(0-indexed) 이전 연속 근무일수"""
        count = 0
        for back in range(1, d + 1):
            s = sched[n].get(d - back + 1)   # dn = d-back+1
            if is_off(s):
                break
            count += 1
        return count

    # ── 초기화 ────────────────────────────────────────────────────────────────
    sched = {n: {} for n in range(num_nurses)}

    def fp_same_shift_conflict(n, dn, shift):
        """n을 dn일 shift에 넣을 때, 이미 그 시프트인 동료와 함께 근무 불가 쌍이면 True"""
        if not fp_set:
            return False
        for m in nurses:
            if m == n:
                continue
            if sched[m].get(dn) == shift and (min(n, m), max(n, m)) in fp_set:
                return True
        return False

    # 개인 신청 우선 적용
    for n_idx, day_shifts in requests.items():
        for day_num, shift_name in day_shifts.items():
            if 0 <= n_idx < num_nurses and 1 <= day_num <= NUM_DAYS:
                sched[n_idx][day_num] = shift_name

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
        앞 = 이미 배정된 비-오프 연속 / 뒤 = 이미 배정된 비-오프 연속.
        """
        before = work_streak_before(n, d)
        after = 0
        for fwd in range(1, NUM_DAYS - d):
            s = sched[n].get(d + fwd + 1)     # dn of (d+fwd)
            if is_off(s):
                break
            after += 1
        return before + 1 + after

    # ── N 시프트 배정 (공평 분배) ─────────────────────────────────────────────
    # 총 슬롯 = 2 × 30 = 60. 일반간호사 num_reg명에게 6~7야로 공평 배분.
    # base_n야를 받는 (num_reg - extra_n)명 + (base_n+1)야를 받는 extra_n명 = 60
    num_reg = len(nurses)
    total_n_slots = 2 * NUM_DAYS   # 60
    base_n        = total_n_slots // num_reg   # e.g. 9명이면 6
    extra_n_count = total_n_slots % num_reg    # e.g. 9명이면 6 (→ 6명 7야, 3명 6야)
    num_low       = num_reg - extra_n_count    # 6야 간호사 수 (e.g. 3)

    # 배분: 앞 번호 간호사(extra_n명)가 base_n+1야, 뒤 번호(num_low명)가 base_n야
    # → 전체 합 = extra_n*(base_n+1) + num_low*base_n = total_n_slots
    n_target     = {}   # 간호사별 N 목표 개수
    n_block_plan = {}   # 간호사별 블록 플랜
    # 7야 블록 패턴 3가지를 간호사마다 순환 배정 (2,2,3 / 2,3,2 / 3,2,2)
    _seven_patterns = [[2, 2, 3], [2, 3, 2], [3, 2, 2]]
    _pat_idx = 0
    for i, n in enumerate(nurses):
        if i < extra_n_count:
            n_target[n]     = base_n + 1      # e.g. 7야
            n_block_plan[n] = _seven_patterns[_pat_idx % 3]
            _pat_idx += 1
        else:
            n_target[n]     = base_n          # e.g. 6야
            n_block_plan[n] = [3, 3]          # 6야: 3+3 블록

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

    for d in range(NUM_DAYS):
        dn = d + 1
        on_n = [n for n in nurses if sched[n].get(dn) == 'N']
        needed = 2 - len(on_n)

        if needed > 0:
            def n_score(n, _d=d):
                continuing = ns[n]['consec'] > 0
                tgt = n_target[n]
                # 비례 기대치 대비 부족분 → 클수록 우선 배정
                expected = (_d / NUM_DAYS) * tgt
                deficit  = expected - ns[n]['total']
                return (0 if continuing else 1, -round(deficit * 4), ns[n]['total'], n)

            cands = []
            for n in nurses:
                if n in on_n or dn in sched[n]:
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

            cands.sort(key=n_score)
            placed = 0
            for n in cands:
                if placed >= needed:
                    break
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
        for n in nurses:
            if n in on_n or dn in sched[n]:
                continue
            if d < ns[n]['ok_from']:
                # 말일에 N 미달 간호사: ok_from 완화 (3,3,1 패턴 허용)
                if not (d == NUM_DAYS - 1 and ns[n]['total'] < n_target[n]):
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
            elif ns[n]['total'] < n_target[n] and (
                d < NUM_DAYS - 1 or
                (d == NUM_DAYS - 1 and n_target[n] - ns[n]['total'] == 1)
            ):
                next_s = sched[n].get(dn + 1)
                is_last = (d == NUM_DAYS - 1)
                min_streak = 1 if is_last else 2
                if is_last or next_s is None or next_s == 'N':
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
            prev_n = d > 0 and sched[n].get(d) == 'N'
            next_n = d < NUM_DAYS - 1 and sched[n].get(dn + 1) == 'N'
            if not prev_n and not next_n:
                day_n_cnt = sum(1 for m in nurses if sched[m].get(dn) == 'N')
                if day_n_cnt >= 3:   # 제거 후에도 ≥2명 유지 가능할 때만 제거
                    del sched[n][dn]
                    ns[n]['total'] -= 1
                # 2명 이하이면 고립 N이라도 유지 (매일 2명 우선)

    # N 블록 끝 다음날 자동 OF
    for n in nurses:
        for d in range(NUM_DAYS - 1):
            dn, ndn = d + 1, d + 2
            if sched[n].get(dn) == 'N' and sched[n].get(ndn) != 'N':
                if ndn not in sched[n]:
                    sched[n][ndn] = 'OF'

    # ★ 긴급 N 보완: 위 모든 단계 후에도 N<2인 날 → 최소 제약으로 강제 배정
    # 절대 규칙 "N 매일 2명" 보장. ok_from·블록 크기 무시, 연속 5일 제한만 유지.
    for d in range(NUM_DAYS):
        dn = d + 1
        on_n = [n for n in nurses if sched[n].get(dn) == 'N']
        if len(on_n) >= 2:
            continue

        # 후보: 오늘 미배정이거나 OF인 간호사 (다른 근무가 있으면 제외)
        cands = [
            n for n in nurses
            if n not in on_n
            and sched[n].get(dn) in (None, 'OF', 'NO')
            and work_streak_before(n, d) + 1 <= 5   # 연속 5일 제한
        ]
        # 우선순위: N 총 개수 적은 순 → 인덱스 순
        cands.sort(key=lambda n: (ns[n]['total'], n))

        needed = 2 - len(on_n)
        placed = 0
        for n in cands:
            if placed >= needed:
                break
            if fp_same_shift_conflict(n, dn, 'N'):
                continue
            sched[n][dn] = 'N'
            ns[n]['total'] += 1
            on_n.append(n)
            placed += 1

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

        def e_score(n):
            streak = work_streak_before(n, d)
            prio = 0 if 1 <= streak <= 4 else (1 if streak == 0 else 2)
            return (prio, e_cnt[n], n)

        cands = [
            n for n in nurses
            if n not in on_e and dn not in sched[n]
            and sched[n].get(dn + 1) not in ('D', 'EDU', '공')
            and streak_total(n, d) <= 5   # 연속근무 5일 초과 금지
            and not fp_same_shift_conflict(n, dn, 'E')
        ]
        cands.sort(key=e_score)
        placed = 0
        for n in cands:
            if placed >= needed:
                break
            if fp_same_shift_conflict(n, dn, 'E'):
                continue
            sched[n][dn] = 'E'
            e_cnt[n] += 1
            placed += 1

    # ── D 시프트 배정 (공평 분배: 4~6개/인) ──────────────────────────────────
    d_cnt = {n: sum(1 for v in sched[n].values() if v == 'D') for n in nurses}

    # 총 D 슬롯 계산 → 간호사별 목표·상한 설정
    total_d_slots = sum(
        2 if (days[_d]['is_weekend'] or days[_d]['is_holiday']
              or sched[0].get(_d + 1) != 'A1' or num_nurses >= 11) else 1
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
        is_off_day = day['is_weekend'] or day['is_holiday']
        head_a1 = sched[0].get(dn) == 'A1'
        # 주말/공휴일 수간호사 부재 → D=2 / 평일 A1 → D=1 / 11명↑ → D=2
        d_target = 2 if (is_off_day or not head_a1 or num_nurses >= 11) else 1

        on_d = [n for n in nurses if sched[n].get(dn) == 'D']
        needed = d_target - len(on_d)
        if needed <= 0:
            continue

        def can_d(n):
            if dn in sched[n]:
                return False
            if d > 0 and sched[n].get(dn - 1) == 'E':
                return False                                    # E-D 금지
            if d >= 2 and sched[n].get(dn - 2) == 'N' \
                    and sched[n].get(dn - 1) in ('OF', 'OH'):
                return False                                    # N-OF-D 금지
            return True

        def d_score(n):
            # 1순위: 개인 목표 초과분 (0이면 목표 이하, 양수면 초과)
            over = max(0, d_cnt[n] - n_d_cap[n])
            # 2순위: 섬 방지 – 앞뒤 모두 OF/미배정이면 패널티
            p = sched[n].get(dn - 1) if d > 0 else None
            nxt = sched[n].get(dn + 1) if d < NUM_DAYS - 1 else None
            isolation = 1 if ((p is None or p in OFF_SET) and
                              (nxt is None or nxt in OFF_SET)) else 0
            # 3순위: D 총 개수 적은 간호사 우선
            # 4순위: 연속 근무 중인 간호사
            streak = work_streak_before(n, d)
            streak_prio = 0 if 1 <= streak <= 4 else 1
            return (over, isolation, d_cnt[n], streak_prio, n)

        # 절대 상한(6) 미만인 간호사 모두 후보, 스코어로 공평 배정
        cands = [n for n in nurses if n not in on_d and can_d(n)
                 and d_cnt[n] < D_ABS_MAX
                 and streak_total(n, d) <= 5
                 and not fp_same_shift_conflict(n, dn, 'D')]
        cands.sort(key=d_score)
        placed = 0
        for n in cands:
            if placed >= needed:
                break
            if fp_same_shift_conflict(n, dn, 'D'):
                continue
            sched[n][dn] = 'D'
            d_cnt[n] += 1
            placed += 1

    # ── 나머지 빈칸 → OF / OH ────────────────────────────────────────────────
    # 공휴일 빈칸은 OH(공휴일 휴무), 일반 빈칸은 OF
    for n in range(num_nurses):
        for d, day in enumerate(days):
            dn = d + 1
            if dn not in sched[n]:
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
            if d > 0 and sched[n].get(dn - 1) == 'E':
                continue
            if d >= 2 and sched[n].get(dn - 2) == 'N' \
                    and sched[n].get(dn - 1) in OFF_SET:
                continue
            if streak_total(n, d) > 5:
                continue
            # 해당 날 D 정원 확인
            is_off_day = day['is_weekend'] or day['is_holiday']
            head_a1 = sched[0].get(dn) == 'A1'
            d_max_day = 2 if (is_off_day or not head_a1 or num_nurses >= 11) else 1
            d_on = sum(1 for m in nurses if sched[m].get(dn) == 'D')
            if d_on >= d_max_day:
                continue
            if fp_same_shift_conflict(n, dn, 'D'):
                continue
            sched[n][dn] = 'D'
            d_cnt[n] = d_cnt.get(n, 0) + 1

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
                if prev_s not in OFF_SET or next_s not in OFF_SET:
                    continue
                if curr_s in OFF_SET:
                    continue
                # OF-work-OF 섬 감지
                if curr_s == 'D':
                    day = days[d]
                    iod = day['is_weekend'] or day['is_holiday']
                    ha1 = sched[0].get(dn) == 'A1'
                    d_min = 2 if (iod or not ha1 or num_nurses >= 11) else 1
                    others_d = sum(1 for m in nurses if m != n and sched[m].get(dn) == 'D')
                    if others_d >= d_min:
                        sched[n][dn] = 'OF'
                        d_cnt[n] -= 1
                        changed = True
                elif curr_s == 'E':
                    others_e = sum(1 for m in nurses if m != n and sched[m].get(dn) == 'E')
                    if others_e >= 2:
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
            if prev_s not in OFF_SET or next_s not in OFF_SET:
                continue
            if curr_s in OFF_SET or curr_s is None:
                continue
            # ── 섬 확인: 앞으로 연장 시도 (next_dn 을 E/D로) ──────────────
            next_dn = dn + 1
            next_d  = d + 1
            if next_d < NUM_DAYS and sched[n].get(next_dn) in OFF_SET:
                after_next = sched[n].get(next_dn + 1) if next_d + 1 < NUM_DAYS else None
                e_on_next  = sum(1 for m in nurses if sched[m].get(next_dn) == 'E')
                # E 연장: E 인원 부족하고 E-D 위반 없고 5일 한도 내
                if (e_on_next < 2 and
                        after_next not in ('D', 'EDU', '공') and
                        streak_total(n, next_d) <= 5):
                    sched[n][next_dn] = 'E'
                    e_cnt[n] = e_cnt.get(n, 0) + 1
                    continue
                # D 연장: 정원 미달일 때만 허용 (평일 수간A1 → D=1 유지)
                day_nx = days[next_d]
                iod_nx = day_nx['is_weekend'] or day_nx['is_holiday']
                ha1_nx = sched[0].get(next_dn) == 'A1'
                d_min_nx = 2 if (iod_nx or not ha1_nx or num_nurses >= 11) else 1
                d_on_nx = sum(1 for m in nurses if sched[m].get(next_dn) == 'D')
                e_prev_nx = sched[n].get(next_dn - 1) == 'E'  # E-D 금지
                n_of_d_nx = (next_d >= 2 and
                             sched[n].get(next_dn - 2) == 'N' and
                             sched[n].get(next_dn - 1) in OFF_SET)
                if (d_on_nx < d_min_nx and
                        not e_prev_nx and not n_of_d_nx and
                        streak_total(n, next_d) <= 5):
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
                        after_prev not in ('D', 'EDU', '공') and
                        streak_total(n, prev_d) <= 5):
                    sched[n][prev_dn] = 'E'
                    e_cnt[n] = e_cnt.get(n, 0) + 1
                    continue
                # D 연장 (뒤쪽 OF): 정원 미달일 때만 허용 (평일 수간A1 → D=1 유지)
                day_pv = days[prev_d]
                iod_pv = day_pv['is_weekend'] or day_pv['is_holiday']
                ha1_pv = sched[0].get(prev_dn) == 'A1'
                d_min_pv = 2 if (iod_pv or not ha1_pv or num_nurses >= 11) else 1
                d_on_pv  = sum(1 for m in nurses if sched[m].get(prev_dn) == 'D')
                e_prev_pv = (prev_d > 0 and sched[n].get(prev_dn - 1) == 'E')
                n_of_d_pv = (prev_d >= 2 and
                             sched[n].get(prev_dn - 2) == 'N' and
                             sched[n].get(prev_dn - 1) in OFF_SET)
                if (d_on_pv < d_min_pv and
                        not e_prev_pv and not n_of_d_pv and
                        streak_total(n, prev_d) <= 5):
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
            if prev_s not in OFF_SET or next_s not in OFF_SET:
                continue
            if curr_s != 'D':   # D 섬만 대상
                continue
            # 이 날 OF인 다른 간호사 중 스왑 적합한 후보 찾기
            swapped = False
            for m in nurses:
                if m == n or sched[m].get(dn) != 'OF':
                    continue
                m_prev = sched[m].get(dn - 1)
                m_next = sched[m].get(dn + 1)
                # m이 이 날 D를 해도 섬이 안 되는 경우 (앞이나 뒤에 근무가 있을 것)
                m_would_be_island = m_prev in OFF_SET and m_next in OFF_SET
                if m_would_be_island:
                    continue
                # D 배정 제약 확인 (n→OF, m→D)
                m_e_prev = (d > 0 and sched[m].get(dn - 1) == 'E')
                m_n_of_d = (d >= 2 and sched[m].get(dn - 2) == 'N'
                            and sched[m].get(dn - 1) in OFF_SET)
                if m_e_prev or m_n_of_d:
                    continue
                if streak_total(m, d) > 5:
                    continue
                if fp_same_shift_conflict(m, dn, 'D'):
                    continue
                # 스왑 실행
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
        is_off_day = day['is_weekend'] or day['is_holiday']
        head_a1 = sched[0].get(dn) == 'A1'
        d_target = 2 if (is_off_day or not head_a1 or num_nurses >= 11) else 1
        on_d = [n for n in nurses if sched[n].get(dn) == 'D']
        if len(on_d) >= d_target:
            continue

        def can_d2(n):
            if sched[n].get(dn) != 'OF':
                return False
            if d > 0 and sched[n].get(dn - 1) == 'E':
                return False
            if d >= 2 and sched[n].get(dn - 2) == 'N' \
                    and sched[n].get(dn - 1) in ('OF', 'OH'):
                return False
            return True

        extras = [n for n in nurses if n not in on_d and can_d2(n)
                  and d_cnt.get(n, 0) < D_ABS_MAX
                  and not fp_same_shift_conflict(n, dn, 'D')]     # ← 상한·함께근무불가
        extras.sort(key=lambda n: (d_cnt[n], n))
        need_extra = d_target - len(on_d)
        placed = 0
        for n in extras:
            if placed >= need_extra:
                break
            if fp_same_shift_conflict(n, dn, 'D'):
                continue
            sched[n][dn] = 'D'
            d_cnt[n] += 1
            placed += 1

    # ★ 후처리 ⑤: D 재분배 – 초과(>목표) 간호사 → 미달(<4) 간호사로 교환
    #   n_high의 섬 D를 n_low의 인접 근무일로 이전 (양쪽 섬 방지)
    def is_d_island(n, dn):
        """해당 날이 n의 D 섬(OF-D-OF)인지"""
        p = sched[n].get(dn - 1) if dn > 1 else 'OF'
        nxt = sched[n].get(dn + 1) if dn < NUM_DAYS else 'OF'
        return (p is None or p in OFF_SET) and (nxt is None or nxt in OFF_SET)

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
                if d_idx > 0 and sched[n_low].get(dn - 1) == 'E':
                    continue
                if d_idx >= 2 and sched[n_low].get(dn - 2) == 'N' \
                        and sched[n_low].get(dn - 1) in OFF_SET:
                    continue
                if streak_total(n_low, d_idx) > 5:
                    continue
                # n_low에게 D를 주면 섬(OF-D-OF)이 되는지 확인
                prev_nl = sched[n_low].get(dn - 1) if d_idx > 0 else 'OF'
                next_nl = sched[n_low].get(dn + 1) if d_idx < NUM_DAYS - 1 else 'OF'
                if (prev_nl is None or prev_nl in OFF_SET) and \
                   (next_nl is None or next_nl in OFF_SET):
                    continue  # 섬 생성 → 스킵
                # 이 날 D를 하는 간호사 중 양도 가능자 찾기
                # 우선순위: 섬 D인 n_high → 비섬 n_high
                d_nurses = [m for m in nurses
                            if sched[m].get(dn) == 'D' and m != n_low
                            and d_cnt.get(m, 0) > D_ABS_MIN]
                d_nurses.sort(key=lambda m: (0 if is_d_island(m, dn) else 1, -d_cnt.get(m, 0)))
                for n_high in d_nurses:
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
                if prev_s not in OFF_SET or next_s not in OFF_SET:
                    continue
                if curr_s in OFF_SET:
                    continue
                if curr_s == 'D':
                    day = days[d]
                    iod = day['is_weekend'] or day['is_holiday']
                    ha1 = sched[0].get(dn) == 'A1'
                    d_min2 = 2 if (iod or not ha1 or num_nurses >= 11) else 1
                    others_d = sum(1 for m in nurses if m != n and sched[m].get(dn) == 'D')
                    fixed2 = False
                    if others_d >= d_min2 and d_cnt.get(n, 0) > D_ABS_MIN:
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
                            m_prev = sched[m].get(dn - 1)
                            m_next = sched[m].get(dn + 1)
                            if m_prev in OFF_SET and m_next in OFF_SET:
                                continue
                            if d > 0 and sched[m].get(dn - 1) == 'E':
                                continue
                            if d >= 2 and sched[m].get(dn - 2) == 'N' \
                                    and sched[m].get(dn - 1) in OFF_SET:
                                continue
                            if streak_total(m, d) > 5:
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
                            if sched[n].get(ext_dn) not in OFF_SET:
                                continue
                            day_ext = days[ext_d]
                            iod_ext = day_ext['is_weekend'] or day_ext['is_holiday']
                            ha1_ext = sched[0].get(ext_dn) == 'A1'
                            d_min_ext = 2 if (iod_ext or not ha1_ext or num_nurses >= 11) else 1
                            d_on_ext = sum(1 for m in nurses if sched[m].get(ext_dn) == 'D')
                            if d_on_ext >= d_min_ext:
                                continue
                            if ext_dn > 1 and sched[n].get(ext_dn - 1) == 'E':
                                continue
                            if ext_d >= 2 and sched[n].get(ext_dn - 2) == 'N' \
                                    and sched[n].get(ext_dn - 1) in OFF_SET:
                                continue
                            if streak_total(n, ext_d) > 5:
                                continue
                            sched[n][ext_dn] = 'D'
                            d_cnt[n] = d_cnt.get(n, 0) + 1
                            changed2 = True
                            break
                elif curr_s == 'E':
                    others_e = sum(1 for m in nurses if m != n and sched[m].get(dn) == 'E')
                    if others_e >= 2:
                        sched[n][dn] = 'OF'
                        e_cnt[n] = e_cnt.get(n, 0) - 1
                        changed2 = True

    # ★ 최종: OF 쿼터 적용 ─────────────────────────────────────────────────────
    # 원칙: 수간호사 OFF(토·일=OF + 공휴일=OH) 합산 수 = 일반 간호사 OFF 쿼터
    # 주(月~日)별 최소 2 OFF(OF+OH+연차) 유지, 초과분은 연차로 전환
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

    # ── 주간 OFF(OF+OH) 수를 동적으로 반환하는 헬퍼 ──────────────────────────
    # OH + OF 가 한 주에 있으면 주간 최소 2 OFF 충족으로 간주
    def _wk_off(n, key):
        """주(key) 내 OF + OH + NO 합산 수 (연차 제외 — 연차는 보조 역할)"""
        return sum(
            1 for d2 in week_days_map[key]
            if sched[n].get(d2) in ('OF', 'OH', 'NO')
        )

    for n in nurses:
        # ── ①  초과 OFF → 연차 (주간 최소 2 OFF 유지)
        # 쿼터 = 토·일 + 공휴일 합산, OH는 법정 휴일이므로 변환 대상에서 제외
        # NO는 수기 휴무이므로 OF→연차 전환 대상에서 제외(쿼터 합산에는 포함)
        nurse_offs_total = sum(1 for s in sched[n].values() if s in ('OF', 'OH', 'NO'))
        surplus = nurse_offs_total - of_quota_month
        nurse_ofs = sorted((dn for dn, s in sched[n].items() if s == 'OF'), reverse=True)
        for dn in nurse_ofs:
            if surplus <= 0:
                break
            wkey = week_map.get(dn)
            # OH + OF 합산이 2 초과인 경우에만 OF → 연차 (2개 유지 보장)
            if wkey and _wk_off(n, wkey) > 2:
                sched[n][dn] = '연'
                surplus -= 1

        # ── ②  주간 최소 2 OFF(OF+OH) 충족 여부 확인
        # OH 1개 + OF 1개 = 2개 → 충족으로 간주 (연차는 보조)
        for wkey, wdays in week_days_map.items():
            if _wk_off(n, wkey) >= 2:
                continue
            # 부족해도 강제 전환 금지 (인력 부족 유발) → 검증기에서 경고 처리

    if fp_set:
        _repair_fp_same_shift_conflicts(sched, nurses, fp_set, days)

    return sched, True, 'FEASIBLE'


def _repair_fp_same_shift_conflicts(sched, nurses, fp_set, days):
    """
    같은 날 D/E/N에 함께 근무 불가 쌍이 남아 있으면,
    해당 날 OF/OH인 다른 간호사와 자리 교환 시도 (당일 시프트 인원 수 유지).
    """
    if not fp_set:
        return
    OFF_OK = {'OF', 'OH', 'NO'}
    SHIFTS = ('N', 'E', 'D')
    for _ in range(150):
        found = None
        for day in days:
            dn = day['day']
            for shift in SHIFTS:
                team = [n for n in nurses if sched[n].get(dn) == shift]
                for i in range(len(team)):
                    for j in range(i + 1, len(team)):
                        a, b = team[i], team[j]
                        if (min(a, b), max(a, b)) in fp_set:
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
            for c in nurses:
                if c == victim:
                    continue
                cur = sched[c].get(dn)
                if cur not in OFF_OK:
                    continue
                bad = False
                for m in nurses:
                    if m in (victim, c):
                        continue
                    if sched[m].get(dn) != shift:
                        continue
                    if (min(c, m), max(c, m)) in fp_set:
                        bad = True
                        break
                if bad:
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
                      nurse_names=None):
    """
    생성된 스케줄을 규칙에 따라 검증하고 위반 사항 목록을 반환한다.
    forbidden_pairs: [(i,j), ...] 일반간호사 인덱스 쌍 — 같은 날 D/E/N 동시 배치 금지
    nurse_names: 표시용 이름 (없으면 기본 수간호사/간호사1…)
    Returns: list of dict  { 'level': 'error'|'warn', 'msg': str }
    """
    issues = []
    days = get_april_days(holidays)
    nurses = list(range(1, num_nurses))
    names = nurse_names if nurse_names is not None else get_nurse_names(num_nurses)
    fp_set = _normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    OFF_SET = {'OF', 'OH', 'NO'}
    WORK_SHIFTS = {'D', 'E', 'N', 'EDU', '연', '공', '병', '경', 'A1'}

    def sh(n, dn):
        return schedule.get(n, {}).get(dn, '')

    def is_off(s):
        return s in OFF_SET or s in ('', None)

    def err(msg):
        issues.append({'level': 'error', 'msg': msg})

    def warn(msg):
        issues.append({'level': 'warn',  'msg': msg})

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
            req_d = 2 if (head != 'A1' or num_nurses >= 11) else 1
            if d_cnt < req_d: err(f"{label} {tag} D 인원 부족: {d_cnt}명 → 필요 {req_d}명")
            if e_cnt < 2:     err(f"{label} {tag} E 인원 부족: {e_cnt}명 → 필요 2명")

        if e_cnt > 2: warn(f"{label} E 인원 초과: {e_cnt}명 (최대 2명)")

    # ── ①b 함께 근무 불가 (같은 날 같은 D/E/N에 둘 다 배치 금지) ────────────
    if fp_set:
        for day in days:
            dn = day['day']
            label = f"{dn}일({day['weekday_name']})"
            for shift in ('D', 'E', 'N'):
                team = [n for n in nurses if sh(n, dn) == shift]
                for i in range(len(team)):
                    for j in range(i + 1, len(team)):
                        a, b = team[i], team[j]
                        if (min(a, b), max(a, b)) in fp_set:
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
        if n_total > 7:
            err(f"{nm} N 초과: {n_total}개 (최대 7개)")

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
                # 말일(NUM_DAYS) 단독 N은 3,3,1 패턴으로 허용
                if blk[0] != NUM_DAYS:
                    err(f"{nm} N 블록 단독: {blk[0]}일 (1개, 최소 2개 연속 — 말일 제외)")
            elif len(blk) > 3:
                err(f"{nm} N 블록 초과: {blk[0]}~{blk[-1]}일 ({len(blk)}개, 최대 3개)")

        for i in range(len(blocks) - 1):
            gap = blocks[i+1][0] - blocks[i][-1] - 1
            if gap < 7:
                warn(f"{nm} N 블록 간격 부족: {blocks[i][-1]}일→{blocks[i+1][0]}일 ({gap}일, 최소 7일)")

        # N 직후 금지 패턴
        for blk in blocks:
            end = blk[-1]
            s1 = sh(n, end + 1)
            s2 = sh(n, end + 2)
            if s1 in OFF_SET and s2 == 'D':
                err(f"{nm} N-휴무-D 금지: {end}일N→{end+1}일{s1}→{end+2}일D")
            if s1 == '연':
                err(f"{nm} N-연 금지: {end}일N→{end+1}일연")
            if s1 in OFF_SET and s2 == 'EDU':
                err(f"{nm} N-휴무-교육 금지: {end}일N→{end+1}일{s1}→{end+2}일EDU")

        # E-D 금지
        for dn in range(1, NUM_DAYS):
            if sh(n, dn) == 'E' and sh(n, dn + 1) == 'D':
                err(f"{nm} E-D 금지: {dn}일E→{dn+1}일D")

        # 연속 근무 최대 5일
        streak = streak_start = 0
        for dn in range(1, NUM_DAYS + 1):
            if sh(n, dn) in WORK_SHIFTS:
                if streak == 0:
                    streak_start = dn
                streak += 1
                if streak > 5:
                    err(f"{nm} 연속근무 초과: {streak_start}~{dn}일 ({streak}일, 최대 5일)")
            else:
                streak = 0

        # OF 간격: 연속 OF(0일)는 허용, 근무가 있다면 2~5일이어야 함 (1일 섬 금지)
        of_days_list = sorted(d for d, s in ns.items() if s in OFF_SET)
        prev_of = None
        for od in of_days_list:
            if prev_of is not None:
                work_btw = sum(
                    1 for d in range(prev_of + 1, od)
                    if sh(n, d) in WORK_SHIFTS
                )
                if work_btw == 1:
                    warn(f"{nm} OF 사이 근무 1일(섬): {prev_of}일OF~{od}일OF — 0일 또는 2~5일이어야 함")
                elif work_btw > 5:
                    warn(f"{nm} OF 간격 초과: {prev_of}일OF~{od}일OF 사이 근무 {work_btw}일 (최대 5일)")
            prev_of = od

        # OF 쿼터 검증: 수간호사 기준(토·일 + 공휴일) 합산과 비교
        of_quota = sum(1 for day in days if day['is_weekend'] or day['is_holiday'])
        off_total = sum(1 for v in ns.values() if v in ('OF', 'OH', 'NO'))
        if off_total > of_quota:
            warn(f"{nm} OFF 초과: {off_total}개 (수간호사 기준 최대 {of_quota}개, 초과분은 연차 권장)")

        # 주(週)별 최소 2 off(OF+OH+연차) 검증
        if days:
            wk_map: dict[int, list] = {}
            for day in days:
                dt  = day['date']
                mon = dt - timedelta(days=dt.weekday())
                wk_map.setdefault(mon.toordinal(), []).append(day['day'])
            for wkey, wdays in wk_map.items():
                # OH + OF 합산으로 판단 (OH 1개 + OF 1개 = 충족)
                off_cnt = sum(1 for d2 in wdays if sh(n, d2) in OFF_SET)
                if off_cnt < 2:
                    d_range = f"{min(wdays)}~{max(wdays)}일"
                    warn(f"{nm} 주간 OFF 부족: {d_range} — {off_cnt}개 (OF+OH+NO 합산 최소 2개)")

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

    EXCEL_BG = {
        'A1': ('4A90D9', 'FFFFFF'), 'D': ('FDD835', '000000'),
        'E': ('FF7043', 'FFFFFF'), 'N': ('283593', 'FFFFFF'),
        'OF': ('EEEEEE', '888888'), 'EDU': ('66BB6A', 'FFFFFF'),
        '연': ('EC407A', 'FFFFFF'), '공': ('AB47BC', 'FFFFFF'),
        '병': ('EF5350', 'FFFFFF'), '경': ('26A69A', 'FFFFFF'),
        'OH': ('FFA726', 'FFFFFF'), 'NO': ('B0BEC5', '263238'),
    }

    N_COL = NUM_DAYS + 2
    OF_COL = NUM_DAYS + 3
    D_COL = NUM_DAYS + 4

    # ─ 타이틀 행
    ws.merge_cells(f'A1:{get_column_letter(D_COL)}1')
    tc = ws['A1']
    tc.value = '응급실 2026년 4월 근무표'
    tc.font = Font(bold=True, size=14, color='FFFFFF')
    tc.fill = PatternFill(start_color='1A237E', end_color='1A237E', fill_type='solid')
    tc.alignment = center
    ws.row_dimensions[1].height = 28

    # ─ 헤더 행 (날짜)
    ws.cell(row=2, column=1, value='간호사')
    hdr = ws.cell(row=2, column=1)
    hdr.fill = PatternFill(start_color='37474F', end_color='37474F', fill_type='solid')
    hdr.font = Font(bold=True, color='FFFFFF', size=10)
    hdr.alignment = center
    hdr.border = thin
    ws.column_dimensions['A'].width = 10

    for d, day in enumerate(days):
        col = d + 2
        cell = ws.cell(row=2, column=col, value=f"{day['day']}\n{day['weekday_name']}")
        cell.alignment = center
        cell.border = thin
        bg = 'C62828' if day['is_holiday'] else ('1565C0' if day['is_weekend'] else '37474F')
        cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
        cell.font = Font(bold=True, color='FFFFFF', size=9)
        ws.column_dimensions[get_column_letter(col)].width = 4.5

    for col, label, bg in [(N_COL, 'N\n합계', '283593'), (OF_COL, 'OF\n합계', '455A64'), (D_COL, 'D\n합계', 'E65100')]:
        c = ws.cell(row=2, column=col, value=label)
        c.alignment = center
        c.border = thin
        c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
        c.font = Font(bold=True, color='FFFFFF', size=9)
        ws.column_dimensions[get_column_letter(col)].width = 5.5

    ws.row_dimensions[2].height = 28

    # ─ 간호사 행
    for n_idx, nurse_name in enumerate(nurse_names):
        row = n_idx + 3
        nc = ws.cell(row=row, column=1, value=nurse_name)
        nc.fill = PatternFill(start_color='455A64', end_color='455A64', fill_type='solid')
        nc.font = Font(bold=True, color='FFFFFF', size=9)
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

        for col, val, bg, fg in [
            (N_COL, n_cnt, '3949AB', 'FFFFFF'),
            (OF_COL, of_cnt, '546E7A', 'FFFFFF'),
            (D_COL, d_cnt, 'F57F17', '000000'),
        ]:
            c = ws.cell(row=row, column=col, value=val)
            c.alignment = center
            c.border = thin
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
            c.font = Font(color=fg, bold=True, size=10)

    # ─ 일별 인원 합계 행
    summary_start = len(nurse_names) + 3
    for idx, (label, shift_key, bg) in enumerate([
        ('D 인원', 'D', 'F57F17'),
        ('E 인원', 'E', 'BF360C'),
        ('N 인원', 'N', '1A237E'),
    ]):
        row = summary_start + idx
        lc = ws.cell(row=row, column=1, value=label)
        lc.fill = PatternFill(start_color=bg, end_color=bg, fill_type='solid')
        lc.font = Font(bold=True, color='FFFFFF', size=9)
        lc.alignment = center
        lc.border = thin
        ws.row_dimensions[row].height = 16

        for d in range(NUM_DAYS):
            day_num = d + 1
            cnt = sum(1 for n in range(num_nurses) if schedule.get(n, {}).get(day_num) == shift_key)
            cell = ws.cell(row=row, column=d + 2, value=cnt)
            cell.alignment = center
            cell.border = thin
            cell.font = Font(bold=True, size=9)

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