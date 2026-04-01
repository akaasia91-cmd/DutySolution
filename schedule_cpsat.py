# -*- coding: utf-8 -*-
"""
OR-Tools CP-SAT (`from ortools.sat.python import cp_model` → `model = cp_model.CpModel()`)
기반 근무표 솔버. 절대 규칙은 `model.Add(...)` 로 선언한다.

※ 근무표 생성은 본 CP-SAT 모듈만 사용한다.
model.Add 로: 일별 E·N=2명, D는 총원·수간 A1 여부(app.d_regular_d_bounds: 11명 주말/비A1평일=2, A1평일=1~2 · 12~13명=2~3)·월간 N·N-D/E-D·
근무 블록 최소 2일·주간 OF≤3·NO≤1·(선택)weekly_of_equiv 선형화·N블록 직후 OF/OH·N-휴무-D/교육 등을 넣는다.
수간 월간 OF 상한은 CP 하드 시 INFEASIBLE 이 잦아 제외하고 `validate_schedule` 경고로만 본다.
OH는 공휴일(holidays로 넘긴 일자)에만 배정 가능: 비공휴일은 model.Add(x[..., OH]==0), 공휴일 없으면 월간 OH 합 0.
나이트 블록 사이 비N 일수는 하드 하한 6일(7일 목표는 이전 버전에서 2단 시도 — 현재는 6만 하드).
퐁당·1일 섬·최소2연속 근무는 소프트(벌점); N-OF-D·전월N-1OF-2D는 하드, N-휴무-휴무·N-휴무-E는 목적 보상(OH≡OF).
공평성·주말 오프 균등은 가벼운 목적(낮은 벌점)만 — 하드 아님. 핵심 인원·N-OF-D 등은 `model.Add` 하드 유지.
D/E 월합 편차·|D−E|·토·일 휴무 편차는 작은 가중으로만 권장; 유연 D·13명 D 선호는 기존 가중 유지.
CP-SAT 는 최대 `_CP_SAT_MAX_TIME_SECONDS`(60초)·`num_search_workers=cpu_count` 로 실행하며, 시간 내
찾은 OPTIMAL/FEASIBLE 해를 사용한다(시간 한도로 인한 미증명 FEASIBLE 포함). 해를 찾은 뒤 `validate_schedule` 결과와 관계없이 success=True.
"""
from __future__ import annotations

import os
from datetime import date, timedelta
from typing import Any

import app

STREAK_WORK_SHIFTS = app.STREAK_WORK_SHIFTS

# 검증 정합용 확장 제약(필요 시 진단만 끄기; 기본은 모두 True)
_CP_STREAK = True
_CP_WEEKLY = True
# 주 2휴무 선형화(weekly_of_equiv) — True면 5월 등 주말·야간 구조에서 INFEASIBLE이 잦음. False면 CP는 OF≤3·NO≤1 등만 하드, 2휴무는 validate 경고로 정합.
_CP_WEEKLY_EQUIV = False
# 수간 OF/OH 월간 상한은 검증 경고(validate)로만 본다. CP 하드 제약은 총 칸·주간 규칙과 충돌해 INFEASIBLE 이 잦다.
_CP_HEAD_QUOTA = False
# N블록 직후 OF/OH·N-휴무-D 는 검증과의 동치 선형화가 어렵거나 과제약(INFEASIBLE)을 유발할 수 있어
# 기본은 검증(validate_schedule)으로만 본다. 필요 시 True.
_CP_N_REST = True
_CP_N_OFF_D = True

# 목적: D/E 월합 max−min 편차 — 낮은 가중(빠른 해 우선, 공평성은 차순위 권장)
_OBJ_FAIRNESS_PRIORITY = 55
# 토·일 OF+OH: False=소프트(편차 벌점만), True=하드 max−min≤1(느리고 INFEASIBLE 위험)
_CP_WEEKEND_OFF_EVEN_HARD = False
_WEEKEND_OFF_EVEN_SOFT_WEIGHT = 32
# 간호사별 월간 |D−E| 합 — 가벼운 권장(값은 공평성 스케일에 맞춤)
_OBJ_DE_MONTH_ABS_WEIGHT = 14
# 퐁당 소프트 벌점
_POND_SOFT_WEIGHT = 80
# 13명: 일별 D 합 최대화(3명 선호) — 음의 가중으로 Minimize 안에서 우선(공평성보다 작게)
_PREFER_D3_WEIGHT_13 = 40
# 나이트 블록 사이 최소 간격(일) — 하드 하한만, 5일 이하는 제약으로 차단
_N_BLOCK_MIN_GAP_HARD = 6
# 야간 후 복귀: N-OF-OF 선호(목적에서 가중 최대화), N-OF-E 차선(더 낮은 가중)
_REWARD_N_OFOF = 3500
_REWARD_N_OFE = 650
# CP-SAT 실행: 최대 시간(초) — 시간 안의 FEASIBLE·OPTIMAL 해를 쓰고 최적 증명에 매달리지 않음
_CP_SAT_MAX_TIME_SECONDS = 60

# 퐁당퐁당·연속근무(사용자 정의); 공(公)·병·경 등은 휴게로 본다(STREAK_WORK의 공와 무관)
_POND_WORK = frozenset({'D', 'E', 'N', 'EDU'})
_POND_REST = frozenset({'OF', 'OH', 'NO', '연', '공', '병', '경'})


def _pond_wr_cell(
    n: int,
    dn: int,
    x: dict,
    req_norm: dict,
    locked: set[tuple[int, int]],
) -> tuple[Any, Any] | None:
    """당일이 근무·휴게인지 0/1 선형식 또는 상수. 분류 불가면 해당 간호사 제약 생략용 None."""
    if (n, dn) in locked:
        sh = req_norm.get(n, {}).get(dn, '') or ''
        if sh in _POND_WORK:
            return 1, 0
        if sh in _POND_REST:
            return 0, 1
        return None
    wt = [x[n, dn, s] for s in _POND_WORK if (n, dn, s) in x]
    rt = [x[n, dn, s] for s in _POND_REST if (n, dn, s) in x]
    if not wt and not rt:
        return None
    wv = sum(wt) if wt else 0
    rv = sum(rt) if rt else 0
    return wv, rv


def _carry_pond_r_last(carry_norm: dict, n: int) -> int | None:
    """전월 마지막 날(당월 1일 직전)이 휴게면 1, 근무면 0, 없으면 None."""
    seq = list(carry_norm.get(n) or ())
    if not seq:
        return None
    sh = seq[-1]
    if sh in _POND_REST:
        return 1
    if sh in _POND_WORK:
        return 0
    return None


def _carry_pond_wr_first_next(
    carry_next: dict, num_nurses: int, n: int,
) -> tuple[int | None, int | None]:
    """
    차월 1일: (w, r) 각 0/1 또는 불명 None.
    월말 연속근무(최소 2일)·퐁당 경계에만 사용.
    """
    seq = list(app._normalize_carry_in(carry_next or {}, num_nurses).get(n) or ())
    if not seq:
        return None, None
    sh = seq[0]
    if sh in _POND_WORK:
        return 1, 0
    if sh in _POND_REST:
        return 0, 1
    return None, None


def _add_ponddang_streak_soft(
    model: Any,
    x: dict,
    regular: list[int],
    num_days: int,
    carry_norm: dict,
    carry_next: dict,
    carry_next_provided: bool,
    req_norm: dict,
    locked: set[tuple[int, int]],
    num_nurses: int,
    penalty_terms: list[Any],
) -> None:
    """
    퐁당·최소 2일 연속 근무: 하드 대신 slack Bool + 목적 벌점.
    w+r_prev+r_next<=2+s, r_prev+w<=1+w_next+s — 평소 s=0 선호.
    """
    for n in regular:
        wr_cache: dict[int, tuple[Any, Any] | None] = {}
        for d in range(1, num_days + 1):
            wr_cache[d] = _pond_wr_cell(n, d, x, req_norm, locked)

        r_before = _carry_pond_r_last(carry_norm, n)
        w_next1, r_next1 = (None, None)
        if carry_next_provided:
            w_next1, r_next1 = _carry_pond_wr_first_next(carry_next, num_nurses, n)

        for d in range(2, num_days):
            mid = wr_cache.get(d)
            pr = wr_cache.get(d - 1)
            nx = wr_cache.get(d + 1)
            if mid is None or pr is None or nx is None:
                continue
            w_m, _ = mid
            _, r_p = pr
            _, r_n = nx
            s1 = model.NewBoolVar('')
            model.Add(w_m + r_p + r_n <= 2 + s1)
            penalty_terms.append(s1)
            r_prev, w_n = pr[1], nx[0]
            w_m2, _ = mid
            s2 = model.NewBoolVar('')
            model.Add(r_prev + w_m2 <= 1 + w_n + s2)
            penalty_terms.append(s2)

        if num_days >= 2:
            w1 = wr_cache.get(1)
            w2 = wr_cache.get(2)
            if w1 is not None and w2 is not None and r_before is not None:
                w_m, _ = w1
                _, r_n = w2
                w_n2, _ = w2
                s3 = model.NewBoolVar('')
                model.Add(w_m + r_before + r_n <= 2 + s3)
                penalty_terms.append(s3)
                s4 = model.NewBoolVar('')
                model.Add(r_before + w_m <= 1 + w_n2 + s4)
                penalty_terms.append(s4)

        if num_days >= 2:
            wlm = wr_cache.get(num_days - 1)
            wlast = wr_cache.get(num_days)
            if wlm is not None and wlast is not None:
                _, r_p = wlm
                w_lv, _ = wlast
                if r_next1 is not None:
                    s5 = model.NewBoolVar('')
                    model.Add(w_lv + r_p + r_next1 <= 2 + s5)
                    penalty_terms.append(s5)
                if w_next1 is not None:
                    s6 = model.NewBoolVar('')
                    model.Add(r_p + w_lv <= 1 + w_next1 + s6)
                    penalty_terms.append(s6)

        if num_days == 1:
            w1 = wr_cache.get(1)
            if w1 is not None and r_before is not None and r_next1 is not None:
                w_m, _ = w1
                s7 = model.NewBoolVar('')
                model.Add(w_m + r_before + r_next1 <= 2 + s7)
                penalty_terms.append(s7)
            if w1 is not None and r_before is not None and w_next1 is not None:
                w_m, _ = w1
                s8 = model.NewBoolVar('')
                model.Add(r_before + w_m <= 1 + w_next1 + s8)
                penalty_terms.append(s8)


def _add_weekly_equiv_linear(
    model: Any,
    of_lin: Any,
    oh_lin: Any,
    no_lin: Any,
    m: int,
) -> None:
    """
    app.weekly_of_equiv_satisfied(of, oh, no, m) 과 동치인 선형 제약
    (BoolOr + OnlyEnforceIf 로 OR 조건 구현).
    """
    if m <= 0:
        return
    if m == 1:
        model.Add(of_lin + oh_lin + no_lin >= 1)
        return
    b1 = model.NewBoolVar('')
    b2 = model.NewBoolVar('')
    b3 = model.NewBoolVar('')
    b4 = model.NewBoolVar('')
    b5 = model.NewBoolVar('')
    model.AddBoolOr([b1, b2, b3, b4, b5])
    model.Add(of_lin >= 2).OnlyEnforceIf(b1)
    model.Add(oh_lin >= 2).OnlyEnforceIf(b2)
    model.Add(of_lin >= 1).OnlyEnforceIf(b3)
    model.Add(oh_lin >= 1).OnlyEnforceIf(b3)
    model.Add(of_lin >= 1).OnlyEnforceIf(b4)
    model.Add(no_lin >= 1).OnlyEnforceIf(b4)
    model.Add(oh_lin >= 1).OnlyEnforceIf(b5)
    model.Add(no_lin >= 1).OnlyEnforceIf(b5)


def _add_n_block_min_gap_hard(
    model: Any,
    x: dict,
    regular: list,
    num_days: int,
    min_gap_days: int,
    carry_in: Any,
    num_nurses: int,
) -> None:
    """
    연속 N 블록 사이 '쉬는 날(비N)' 개수 최소 min_gap_days — app.validate gap 와 동일 계산.
    gap = 다음 블록 첫 N일 - 이전 블록 마지막 N일 - 1 >= min_gap_days.
    전월 말 N 이고 당월 1일이 연속이 아니면, 첫 당월 N은 (min_gap+1)일 이전에는 배치 불가.
    """
    for n in regular:
        end_blk: dict[int, Any] = {}
        start_blk: dict[int, Any] = {}
        carry_tail = _carry_prev_is_n(carry_in, n, num_nurses)

        for d in range(1, num_days + 1):
            if (n, d, 'N') not in x:
                continue
            xd = x[n, d, 'N']
            if d < num_days and (n, d + 1, 'N') in x:
                eb = model.NewBoolVar(f'nbend_{n}_{d}')
                model.Add(eb <= xd)
                model.Add(eb + x[n, d + 1, 'N'] <= 1)
                model.Add(eb >= xd - x[n, d + 1, 'N'])
                end_blk[d] = eb
            else:
                end_blk[d] = xd
            if d == 1:
                if carry_tail:
                    start_blk[d] = 0
                else:
                    start_blk[d] = xd
            else:
                if (n, d - 1, 'N') in x:
                    sb = model.NewBoolVar(f'nbst_{n}_{d}')
                    model.Add(sb <= xd)
                    model.Add(sb + x[n, d - 1, 'N'] <= 1)
                    model.Add(sb >= xd - x[n, d - 1, 'N'])
                    start_blk[d] = sb
                else:
                    start_blk[d] = xd

        for d1 in range(1, num_days + 1):
            e1 = end_blk.get(d1)
            if e1 is None or (isinstance(e1, int) and e1 == 0):
                continue
            d2_max = min(num_days, d1 + min_gap_days)
            for d2 in range(d1 + 2, d2_max + 1):
                e2 = start_blk.get(d2)
                if e2 is None or (isinstance(e2, int) and e2 == 0):
                    continue
                mid = [x[n, t, 'N'] for t in range(d1 + 1, d2) if (n, t, 'N') in x]
                sm = sum(mid) if mid else 0
                model.Add(e1 + e2 <= 1 + sm)

        if carry_tail:
            for s in range(2, min(num_days, min_gap_days) + 1):
                if (n, s, 'N') not in x:
                    continue
                preds = [x[n, t, 'N'] for t in range(1, s) if (n, t, 'N') in x]
                if preds:
                    model.Add(x[n, s, 'N'] <= sum(preds))
                else:
                    model.Add(x[n, s, 'N'] == 0)


def _normalize_requests(requests: dict | None) -> dict[int, dict[int, str]]:
    out: dict[int, dict[int, str]] = {}
    if not requests:
        return out
    for k, v in requests.items():
        try:
            ni = int(k)
        except (TypeError, ValueError):
            continue
        inner: dict[int, str] = {}
        for dk, sv in (v or {}).items():
            try:
                dni = int(dk)
            except (TypeError, ValueError):
                continue
            inner[dni] = str(sv).strip()
        if inner:
            out[ni] = inner
    return out


def _locked_cells(requests: dict | None, num_nurses: int) -> set[tuple[int, int]]:
    r = _normalize_requests(requests)
    locked: set[tuple[int, int]] = set()
    for ni, ds in r.items():
        if not (0 <= ni < num_nurses):
            continue
        for dn in ds:
            locked.add((ni, dn))
    return locked


def _build_head_schedule(
    days: list, requests: dict | None, num_nurses: int,
) -> dict[int, str]:
    """수간호사(0) 기본 패턴 + 신청 덮어쓰기."""
    rq0 = _normalize_requests(requests).get(0, {})
    sched0: dict[int, str] = {}
    for day in days:
        dn = day['day']
        if dn in rq0:
            sched0[dn] = rq0[dn]
            continue
        if day['is_holiday']:
            sched0[dn] = 'OH'
        elif day['is_weekend']:
            sched0[dn] = 'OF'
        else:
            sched0[dn] = 'A1'
    return sched0


def _carry_prev_is_n(carry_in: dict | None, n: int, num_nurses: int) -> bool:
    c = app._normalize_carry_in(carry_in or {}, num_nurses).get(n) or ()
    return bool(c) and c[-1] == 'N'


def _carry_prev_is_e(carry_in: dict | None, n: int, num_nurses: int) -> bool:
    c = app._normalize_carry_in(carry_in or {}, num_nurses).get(n) or ()
    return bool(c) and c[-1] == 'E'


def _bool_and3(model: Any, a: Any, b: Any, c: Any, name: str) -> Any:
    """a∧b∧c (a,b,c가 0/1 선형식·상수)."""
    r = model.NewBoolVar(name)
    model.Add(r <= a)
    model.Add(r <= b)
    model.Add(r <= c)
    model.Add(r >= a + b + c - 2)
    return r


def _off_of_or_oh(model: Any, x: dict, n: int, d: int, name: str) -> Any:
    """해당 일이 OF 또는 OH(휴무 동등)인지 0/1. 둘 다 변수 없으면 정수 0만 반환(BoolVar 아님)."""
    has_of = (n, d, 'OF') in x
    has_oh = (n, d, 'OH') in x
    if not has_of and not has_oh:
        return 0
    if has_of and not has_oh:
        return x[n, d, 'OF']
    if has_oh and not has_of:
        return x[n, d, 'OH']
    v = model.NewBoolVar(name)
    a, b = x[n, d, 'OF'], x[n, d, 'OH']
    model.Add(v >= a)
    model.Add(v >= b)
    model.Add(v <= a + b)
    return v


def _is_absent_off_term(term: Any) -> bool:
    """_off_of_or_oh가 돌려준 '해당 일 휴무 변수 없음'인 정수 0만 True. BoolVar는 절대 True가 되지 않음."""
    return isinstance(term, int) and term == 0


def _add_n_of_d_carry_hard(
    model: Any,
    x: dict,
    regular: list[int],
    num_days: int,
    carry_in: Any,
    num_nurses: int,
) -> None:
    """전월 말 N 직후 당월 1일 휴무(OF/OH)·2일 D/EDU 복귀 금지 (월내 N-휴무-D와 동일)."""
    if num_days < 2:
        return
    for n in regular:
        if not _carry_prev_is_n(carry_in, n, num_nurses):
            continue
        for off_s, bad_s in (('OF', 'D'), ('OF', 'EDU'), ('OH', 'D'), ('OH', 'EDU')):
            if (n, 1, off_s) in x and (n, 2, bad_s) in x:
                model.Add(1 + x[n, 1, off_s] + x[n, 2, bad_s] <= 2)


def _collect_n_recovery_reward_vars(
    model: Any,
    x: dict,
    regular: list[int],
    num_days: int,
    carry_in: Any,
    num_nurses: int,
) -> tuple[list[Any], list[Any]]:
    """
    N 후 이틀 휴무(OF/OH)·차선 N-휴무-E 지표. 목적 Minimize에서 빼면 최대화.
    """
    r_ofof: list[Any] = []
    r_ofe: list[Any] = []
    for n in regular:
        if num_days >= 2 and _carry_prev_is_n(carry_in, n, num_nurses):
            o1 = _off_of_or_oh(model, x, n, 1, f'nrr_o1_{n}_c')
            o2 = _off_of_or_oh(model, x, n, 2, f'nrr_o2_{n}_c')
            if not _is_absent_off_term(o1) and not _is_absent_off_term(o2):
                r_ofof.append(_bool_and3(model, 1, o1, o2, f'nrr_of_{n}_c'))
            if not _is_absent_off_term(o1) and (n, 2, 'E') in x:
                r_ofe.append(_bool_and3(model, 1, o1, x[n, 2, 'E'], f'nrr_e_{n}_c'))
        for d in range(1, num_days - 1):
            if (n, d, 'N') not in x:
                continue
            o1 = _off_of_or_oh(model, x, n, d + 1, f'nrr_o1_{n}_{d}')
            o2 = _off_of_or_oh(model, x, n, d + 2, f'nrr_o2_{n}_{d}')
            if _is_absent_off_term(o1):
                continue
            a_n = x[n, d, 'N']
            if not _is_absent_off_term(o2):
                r_ofof.append(_bool_and3(model, a_n, o1, o2, f'nrr_of_{n}_{d}'))
            if (n, d + 2, 'E') in x:
                r_ofe.append(_bool_and3(model, a_n, o1, x[n, d + 2, 'E'], f'nrr_e_{n}_{d}'))
    return r_ofof, r_ofe


def _streak_terms_for_month_day(
    n: int, dn: int,
    x: dict,
    req_norm: dict,
    locked: set[tuple[int, int]],
) -> tuple[int, list]:
    """연속근무(검증 STREAK_WORK_SHIFTS)에 대한 상수항 + BoolVar 리스트."""
    if (n, dn) in locked:
        sh = req_norm.get(n, {}).get(dn, '')
        return (1 if sh in STREAK_WORK_SHIFTS else 0, [])
    terms = []
    for s in ('D', 'E', 'N', 'EDU', '공'):
        key = (n, dn, s)
        if key in x:
            terms.append(x[key])
    return (0, terms)


def solve_schedule_cpsat(
    num_nurses: int,
    requests: dict | None,
    holidays: tuple | list = (),
    forbidden_pairs: Any = None,
    carry_in: dict | None = None,
    carry_next_month: Any = None,
    shift_bans: dict | None = None,
) -> tuple[dict | None, bool, str]:
    """
    CP-SAT로 스케줄 생성. 반환: (sched_dict, success, status_str)
    sched: { nurse_idx: { day: shift_str } }

    success: CP-SAT 가 OPTIMAL/FEASIBLE 해를 찾은 경우에만 True(INFEASIBLE 등은 False).
    해를 찾은 뒤 검증 오류·경고는 `validate_schedule`·UI 팝업으로 안내하며 success 는 막지 않는다.
    """
    from ortools.sat.python import cp_model

    days = app.get_april_days(holidays)
    num_days = app.NUM_DAYS
    n_abs_max = app.N_ABS_MAX
    holiday_days = frozenset(d['day'] for d in days if d['is_holiday'])

    if num_nurses < 2:
        return None, False, 'CP-SAT: 간호사 인원이 부족합니다.'

    fp_map = app._normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    den_bans = app._normalize_shift_bans(shift_bans, num_nurses)
    req_norm = _normalize_requests(requests)
    locked = _locked_cells(requests, num_nurses)
    for ni, ds in req_norm.items():
        if not (0 <= ni < num_nurses):
            continue
        for dn, shv in ds.items():
            try:
                dni = int(dn)
            except (TypeError, ValueError):
                continue
            if shv == 'OH' and dni not in holiday_days:
                return None, False, (
                    'CP-SAT: OH는 「공휴일 날짜」에 포함된 일에만 신청·배정할 수 있습니다. '
                    f'{dni}일은 목록에 없습니다.'
                )
    carry_next_provided = carry_next_month is not None
    carry_next = app._normalize_carry_in(carry_next_month, num_nurses) if carry_next_month else {}
    month_first = date(app.YEAR, app.MONTH, 1)
    month_last = date(app.YEAR, app.MONTH, num_days)

    head = _build_head_schedule(days, requests, num_nurses)
    sched: dict[int, dict[int, str]] = {0: dict(head)}
    regular = list(range(1, num_nurses))
    num_reg = len(regular)

    # 잠금 셀은 신청 시프트만 허용(검증 1:1 일치)
    def banned_shifts(ni: int) -> frozenset:
        return den_bans.get(ni, frozenset())

    def allowed_shifts_for_cell(ni: int, dn: int) -> list[str]:
        if (ni, dn) in locked:
            sh = req_norm.get(ni, {}).get(dn, 'OF')
            return [sh]
        core = ('D', 'E', 'N', 'OF', 'OH')
        return [s for s in core if s not in banned_shifts(ni)] or ['OF']

    total_n_slots = 2 * num_days
    n_targets = app._compute_n_targets_fair(num_reg, total_n_slots, n_abs_max)
    tgt_map = {regular[i]: n_targets[i] for i in range(num_reg)}

    head_of_q = app._monthly_head_nurse_of_count({0: head}, days)
    head_oh_q = app._monthly_head_nurse_oh_count({0: head}, days)

    n_gap_suffix = ''

    model = cp_model.CpModel()
    x: dict[tuple[int, int, str], Any] = {}
    
    for n in regular:
        for d in range(1, num_days + 1):
            allowed = allowed_shifts_for_cell(n, d)
            for s in allowed:
                x[n, d, s] = model.NewBoolVar(f'x_{n}_{d}_{s}')
            model.Add(sum(x[n, d, s] for s in allowed) == 1)

    # OH는 holidays에 명시된 공휴일에만 가능(그 외 날짜·공휴일 목록 비어 있음 → OH 불가)
    for n in regular:
        for d in range(1, num_days + 1):
            if d not in holiday_days and (n, d, 'OH') in x:
                model.Add(x[n, d, 'OH'] == 0)

    # 일별 D / E / N — E·N 항상 2명; D는 수간 해당일 시프트·주말·공휴·총원에 따른 하한·상한
    _flex_d_days: list[int] = []
    for d in range(1, num_days + 1):
        day = days[d - 1]
        hsh = head.get(d) or ''
        lo, hi = app.d_regular_d_bounds(num_nurses, day, hsh)
        d_sum = sum(x[n, d, 'D'] for n in regular if (n, d, 'D') in x)
        model.Add(d_sum >= lo)
        model.Add(d_sum <= hi)
        if lo < hi:
            _flex_d_days.append(d)
        model.Add(sum(x[n, d, 'E'] for n in regular if (n, d, 'E') in x) == 2)
        model.Add(sum(x[n, d, 'N'] for n in regular if (n, d, 'N') in x) == 2)

    # 간호사별 월간 N 합
    for n in regular:
        n_vars = [x[n, d, 'N'] for d in range(1, num_days + 1) if (n, d, 'N') in x]
        if n_vars:
            model.Add(sum(n_vars) == tgt_map.get(n, 0))

    # OF/OH 월간 상한: 수간호사와 동일(검증 경고 조건을 선제적으로 만족)
    if _CP_HEAD_QUOTA:
        for n in regular:
            of_vars = [x[n, d, 'OF'] for d in range(1, num_days + 1) if (n, d, 'OF') in x]
            oh_vars = [x[n, d, 'OH'] for d in range(1, num_days + 1) if (n, d, 'OH') in x]
            if of_vars:
                model.Add(sum(of_vars) <= head_of_q)
            if oh_vars:
                model.Add(sum(oh_vars) <= head_oh_q)

    # N-D 금지
    for n in regular:
        for d in range(2, num_days + 1):
            if (n, d, 'D') in x and (n, d - 1, 'N') in x:
                model.Add(x[n, d, 'D'] + x[n, d - 1, 'N'] <= 1)
        if (n, 1, 'D') in x and _carry_prev_is_n(carry_in, n, num_nurses):
            model.Add(x[n, 1, 'D'] == 0)

    # E-D 금지
    for n in regular:
        for d in range(2, num_days + 1):
            if (n, d, 'D') in x and (n, d - 1, 'E') in x:
                model.Add(x[n, d, 'D'] + x[n, d - 1, 'E'] <= 1)
        if (n, 1, 'D') in x and _carry_prev_is_e(carry_in, n, num_nurses):
            model.Add(x[n, 1, 'D'] == 0)

    if _CP_N_REST:
        # 전월 말 N → 1일: N이 아니면 공휴 OH / 평일 OF (검증: N블록 직후 휴무)
        day1_hol = bool(days[0]['is_holiday'])
        for n in regular:
            if not _carry_prev_is_n(carry_in, n, num_nurses):
                continue
            if (n, 1, 'N') not in x:
                continue
            need = 'OH' if day1_hol else 'OF'
            if (n, 1, need) in x:
                model.Add(x[n, 1, need] >= 1 - x[n, 1, 'N'])

        # N 다음날: N이 아니면 말일이 아닌 경우 공휴 OH / 평일 OF
        _dn_holiday = {d['day']: bool(d['is_holiday']) for d in days}
        for n in regular:
            for d in range(1, num_days):
                if (n, d, 'N') not in x:
                    continue
                need = 'OH' if _dn_holiday.get(d + 1) else 'OF'
                if (n, d + 1, need) not in x:
                    continue
                xn1 = x[n, d + 1, 'N'] if (n, d + 1, 'N') in x else 0
                model.Add(x[n, d + 1, need] >= x[n, d, 'N'] - xn1)

    if _CP_N_OFF_D:
        # N + (OF|OH) + (D|EDU) 동시 3은 불가 → 항상 선형식: a+b+c<=2 (잠금 셀이 셋째칸만 고정일 때는 앞 두 항<=1)
        for n in regular:
            for end in range(1, num_days - 1):
                if (n, end, 'N') not in x:
                    continue
                v_n = x[n, end, 'N']
                for off_s, bad_s in (('OF', 'D'), ('OF', 'EDU'), ('OH', 'D'), ('OH', 'EDU')):
                    if (n, end + 1, off_s) not in x:
                        continue
                    v_mid = x[n, end + 1, off_s]
                    k_bad = (n, end + 2, bad_s)
                    if k_bad in x:
                        model.Add(v_n + v_mid + x[k_bad] <= 2)
                    elif (n, end + 2) in locked:
                        if req_norm.get(n, {}).get(end + 2, '') == bad_s:
                            model.Add(v_n + v_mid <= 1)

        _add_n_of_d_carry_hard(model, x, regular, num_days, carry_in, num_nurses)

    # 단독 나이트 금지 — 말일 NUM_DAYS 예외
    for n in regular:
        for d in range(1, num_days):
            if (n, d, 'N') not in x:
                continue
            rhs_terms = []
            if d > 1 and (n, d - 1, 'N') in x:
                rhs_terms.append(x[n, d - 1, 'N'])
            elif d == 1 and _carry_prev_is_n(carry_in, n, num_nurses):
                rhs_terms.append(1)
            if (n, d + 1, 'N') in x:
                rhs_terms.append(x[n, d + 1, 'N'])
            if not rhs_terms:
                model.Add(x[n, d, 'N'] == 0)
            else:
                model.Add(x[n, d, 'N'] <= sum(rhs_terms))

    # N 연속 최대 3일 → 4일 합 <= 3
    for n in regular:
        for d in range(1, num_days - 2):
            quad = [x[n, d + k, 'N'] for k in range(4) if (n, d + k, 'N') in x]
            if len(quad) == 4:
                model.Add(sum(quad) <= 3)

    _add_n_block_min_gap_hard(
        model, x, regular, num_days, _N_BLOCK_MIN_GAP_HARD,
        carry_in, num_nurses,
    )

    carry_norm = app._normalize_carry_in(carry_in, num_nurses)

    # 연속 근무 최대 5일: (전월 carry + 당월) 6일 창마다 근무 시프트 합 <= 5
    if _CP_STREAK:
        for n in regular:
            carry_seq = list(carry_norm.get(n) or ())
            lcarry = len(carry_seq)
            span = lcarry + num_days
            if span < 6:
                continue
            for w in range(span - 5):
                const = 0
                terms: list = []
                for k in range(6):
                    idx = w + k
                    if idx < lcarry:
                        const += 1 if carry_seq[idx] in STREAK_WORK_SHIFTS else 0
                    else:
                        dn = idx - lcarry + 1
                        c0, tlist = _streak_terms_for_month_day(n, dn, x, req_norm, locked)
                        const += c0
                        terms.extend(tlist)
                if terms:
                    model.Add(const + sum(terms) <= 5)
                else:
                    model.Add(const <= 5)

    pond_penalty_terms: list[Any] = []
    _add_ponddang_streak_soft(
        model, x, regular, num_days,
        carry_norm, carry_next, carry_next_provided,
        req_norm, locked, num_nurses,
        pond_penalty_terms,
    )

    # 주간 OF 상한·NO 상한·주 2 휴무(weekly_of_equiv) — validate 와 동일한 주 경계
    if _CP_WEEKLY:
        wk_map: dict[int, list] = {}
        for day in days:
            dt = day['date']
            mon = dt - timedelta(days=dt.weekday())
            wk_map.setdefault(mon.toordinal(), []).append(day['day'])

        for _wk, wdays in wk_map.items():
            if not wdays:
                continue
            mon_date = date.fromordinal(_wk)
            for n in regular:
                pre_of, pre_oh, pre_no, n_prev = app._carry_week_prev_month_off_counts(
                    carry_norm, n, mon_date, month_first,
                )
                post_of, post_oh, post_no, n_next = app._carry_week_next_month_off_counts(
                    carry_next, n, mon_date, month_last,
                )
                len_w = len(wdays)
                m_week = n_prev + len_w + n_next

                of_month = sum(x[n, d, 'OF'] for d in wdays if (n, d, 'OF') in x)
                oh_month = sum(x[n, d, 'OH'] for d in wdays if (n, d, 'OH') in x)
                no_month = sum(x[n, d, 'NO'] for d in wdays if (n, d, 'NO') in x)

                of_vis_expr = pre_of + of_month
                tot_of_add = post_of if carry_next_provided else 0
                model.Add(of_vis_expr + tot_of_add <= 3)

                no_week_expr = pre_no + no_month + post_no
                model.Add(no_week_expr <= 1)

                if m_week <= 0:
                    continue
                if n_next > 0 and not carry_next_provided:
                    continue

                if _CP_WEEKLY_EQUIV:
                    of_t_expr = pre_of + of_month + post_of
                    oh_t_expr = pre_oh + oh_month + post_oh
                    no_t_expr = pre_no + no_month + post_no
                    _add_weekly_equiv_linear(model, of_t_expr, oh_t_expr, no_t_expr, m_week)

    # 함께 근무 불가
    for dn in range(1, num_days + 1):
        for shift in ('D', 'E', 'N'):
            for a in range(num_nurses):
                for b in range(a + 1, num_nurses):
                    key = (a, b)
                    if key not in fp_map or shift not in fp_map[key]:
                        continue
                    if a == 0:
                        if head.get(dn) != shift:
                            continue
                        if (b, dn, shift) in x:
                            model.Add(x[b, dn, shift] == 0)
                    elif b == 0:
                        if head.get(dn) != shift:
                            continue
                        if (a, dn, shift) in x:
                            model.Add(x[a, dn, shift] == 0)
                    else:
                        if (a, dn, shift) in x and (b, dn, shift) in x:
                            model.Add(x[a, dn, shift] + x[b, dn, shift] <= 1)

    # ── 권장(저가중): D/E 편차·주말 휴무 편차 — 하드 아님 ---------------------------------
    d_tot_exprs: list[Any] = []
    e_tot_exprs: list[Any] = []
    for n in regular:
        dv = [x[n, d, 'D'] for d in range(1, num_days + 1) if (n, d, 'D') in x]
        ev = [x[n, d, 'E'] for d in range(1, num_days + 1) if (n, d, 'E') in x]
        d_tot_n = sum(dv) if dv else 0
        e_tot_n = sum(ev) if ev else 0
        d_tot_exprs.append(d_tot_n)
        e_tot_exprs.append(e_tot_n)

    wk_nums = [d for d in range(1, num_days + 1) if days[d - 1]['is_weekend']]
    wk_even_spread: Any | None = None
    if wk_nums and regular:
        wk_off_exprs: list[Any] = []
        for n in regular:
            wof = []
            for d in wk_nums:
                if (n, d, 'OF') in x:
                    wof.append(x[n, d, 'OF'])
                if (n, d, 'OH') in x:
                    wof.append(x[n, d, 'OH'])
            wk_off_exprs.append(sum(wof) if wof else 0)
        wk_ub = len(wk_nums)
        max_wk = model.NewIntVar(0, wk_ub, '')
        min_wk = model.NewIntVar(0, wk_ub, '')
        for wo in wk_off_exprs:
            model.Add(max_wk >= wo)
            model.Add(min_wk <= wo)
        if _CP_WEEKEND_OFF_EVEN_HARD:
            model.Add(max_wk <= min_wk + 1)
        else:
            wk_even_spread = model.NewIntVar(0, wk_ub, '')
            model.Add(wk_even_spread == max_wk - min_wk)

    max_d_tot = model.NewIntVar(0, num_days, '')
    min_d_tot = model.NewIntVar(0, num_days, '')
    max_e_tot = model.NewIntVar(0, num_days, '')
    min_e_tot = model.NewIntVar(0, num_days, '')
    for d_tot_n, e_tot_n in zip(d_tot_exprs, e_tot_exprs):
        model.Add(max_d_tot >= d_tot_n)
        model.Add(min_d_tot <= d_tot_n)
        model.Add(max_e_tot >= e_tot_n)
        model.Add(min_e_tot <= e_tot_n)
    diff_d = model.NewIntVar(0, num_days, '')
    diff_e = model.NewIntVar(0, num_days, '')
    model.Add(diff_d == max_d_tot - min_d_tot)
    model.Add(diff_e == max_e_tot - min_e_tot)

    de_abs_terms: list[Any] = []
    for d_tot_n, e_tot_n in zip(d_tot_exprs, e_tot_exprs):
        delta = model.NewIntVar(-num_days, num_days, '')
        model.Add(delta == d_tot_n - e_tot_n)
        ab = model.NewIntVar(0, num_days, '')
        model.AddAbsEquality(ab, delta)
        de_abs_terms.append(ab)

    obj = _OBJ_FAIRNESS_PRIORITY * (diff_d + diff_e)
    if de_abs_terms:
        obj += _OBJ_DE_MONTH_ABS_WEIGHT * sum(de_abs_terms)
    r_ofof, r_ofe = _collect_n_recovery_reward_vars(
        model, x, regular, num_days, carry_in, num_nurses,
    )
    if r_ofof:
        obj -= _REWARD_N_OFOF * sum(r_ofof)
    if r_ofe:
        obj -= _REWARD_N_OFE * sum(r_ofe)
    if wk_even_spread is not None:
        obj += _WEEKEND_OFF_EVEN_SOFT_WEIGHT * wk_even_spread
    if _flex_d_days:
        _obj_flex_d = [
            x[n, d, 'D']
            for d in _flex_d_days
            for n in regular
            if (n, d, 'D') in x
        ]
        if _obj_flex_d:
            if num_nurses == 13:
                obj -= _PREFER_D3_WEIGHT_13 * sum(_obj_flex_d)
            else:
                obj += sum(_obj_flex_d)
    if pond_penalty_terms:
        obj += _POND_SOFT_WEIGHT * sum(pond_penalty_terms)
    model.Minimize(obj)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(_CP_SAT_MAX_TIME_SECONDS)
    solver.parameters.num_search_workers = max(1, os.cpu_count() or 1)
    status = solver.Solve(model)
    if status == cp_model.UNKNOWN:
        return None, False, (
            f'CP-SAT: UNKNOWN — 최대 {_CP_SAT_MAX_TIME_SECONDS}초 내 실행 가능한 해를 찾지 못했습니다.'
        )
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return None, False, (
            'CP-SAT: INFEASIBLE — N-D·주간 2OFF·나이트 간격(6일 이상) 등 하드 제약을 만족하는 배정이 없습니다. '
            '신청·인원을 조정해 주세요.'
        )

    for n in regular:
        sched[n] = {}
        for d in range(1, num_days + 1):
            allowed = allowed_shifts_for_cell(n, d)
            chosen = None
            for s in allowed:
                if (n, d, s) in x and solver.Value(x[n, d, s]) == 1:
                    chosen = s
                    break
            if chosen is None:
                return None, False, 'CP-SAT: 내부 오류(배정 복원 실패).'
            sched[n][d] = chosen

    issues = app.validate_schedule(
        sched, num_nurses, holidays,
        forbidden_pairs=forbidden_pairs, carry_in=carry_in,
        requests=requests, carry_next_month=carry_next_month,
        shift_bans=shift_bans,
    )
    err_n = sum(1 for z in issues if z.get('level') == 'error')
    warn_n = sum(1 for z in issues if z.get('level') == 'warn')

    # CP-SAT 가 해를 찾았다면 검증 결과와 관계없이 근무표를 화면에 반영(success=True)
    if err_n == 0 and warn_n == 0:
        return sched, True, 'FEASIBLE_CP_SAT (검증 통과)' + n_gap_suffix
    preview = '; '.join(z.get('msg', '') for z in issues[:3])
    bits = []
    if err_n:
        bits.append(f'검토 권고(오류 표기) {err_n}건')
    if warn_n:
        bits.append(f'경고 {warn_n}건')
    tail = f' {"; ".join(bits)}. 예: {preview}' if preview else f' {"; ".join(bits)}.'
    return sched, True, f'FEASIBLE_CP_SAT — 주의사항과 함께 생성됨.{tail}' + n_gap_suffix
