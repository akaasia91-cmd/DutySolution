# -*- coding: utf-8 -*-
"""
OR-Tools CP-SAT 근무표 솔버.

1차: 일별 E/N 최소·D범위·신청 고정 하드 + 삶의 질 소프트. 재생성 시 시드를 달리해 다른 해 탐색.
CP-SAT가 해를 못 내면 신청 유지·빈칸 OF 폴백으로도 표를 반환(검증 이슈는 별도). 솔버 상한 약 10초.
"""
from __future__ import annotations

import os
from datetime import date, timedelta
from typing import Any

import app

STREAK_WORK_SHIFTS = app.STREAK_WORK_SHIFTS

# hospital_config.json 시드·Streamlit DEFAULT_DEPT_TOTAL_HEADCOUNT와 동기(총원=수간+일반)
DEFAULT_DEPT_TOTAL_HEADCOUNT: dict[str, int] = {
    "응급실": 10,
    "신관 3병동": 12,
    "본관 5병동": 12,
    "본관 6병동": 12,
    "본관 7병동": 12,
    "본관 8병동": 11,
    "중환자실": 22,
}

# 근무표 위반 셀 시각화 — app.VIOLATION_CELL_* / Streamlit Styler와 동일 색상
SCHEDULE_VIOLATION_HIGHLIGHT_ERROR = "#87CEFA"  # LightSkyBlue (오류)
SCHEDULE_VIOLATION_HIGHLIGHT_WARN = "#98FB98"  # PaleGreen (경고)

# 자동 배정 가능: D·E·N·OF·OH만. 경·공·EDU·연·병·NO는 신청 칸에만 허용.
_AUTO_ASSIGN_SHIFTS = ('D', 'E', 'N', 'OF', 'OH')
# 소프트 가중 (절대 규칙 외)
_W_TIER2 = 100_000
_W_TIER3 = 10_000
# 보조 형평(티어 대비 미미하게 유지)
_W_N_SUM_SQUARES = 400
_W_N_SPREAD_EXCESS = 8_000
_W_N_TGT_DEV = 8_000
_W_LOW_FAIR = 20
_OBJ_FAIRNESS_PRIORITY = 12
_CP_WEEKEND_OFF_EVEN_HARD = False
_WEEKEND_OFF_EVEN_SOFT_WEIGHT = 18
_OBJ_DE_MONTH_ABS_WEIGHT = 4
_POND_SOFT_WEIGHT = 25
_POND_SOFT_WEIGHT_N10 = 15
_PREFER_D3_WEIGHT_13 = 40
_N_BLOCK_GAP_MIN_DAYS = 5
# N블록 6일·7일 간격 선호 — 티어2 하위
_PENALTY_N_GAP_TIGHT_A = 25_000
_PENALTY_N_GAP_TIGHT_B = 20_000
_REWARD_N_OFOF = 3_800
_REWARD_N_OFE = 650
_CP_SAT_MAX_TIME_SECONDS = 10.0
# 총원(수간 포함) 기준 — 레거시(솔버 상한은 전 부서 동일 10초)
_LARGE_STAFF_MIN_TOTAL_NURSES = 13


def _cpsat_solver_max_seconds(unit_profile: str, num_nurses: int) -> float:
    """응답 속도 우선: 모든 부서 동일 상한(초)."""
    _ = unit_profile
    _ = num_nurses
    return float(_CP_SAT_MAX_TIME_SECONDS)


_W_SAFE_N_D = 620_000
_W_SAFE_E_AFTER = 620_000
_W_SAFE_N_REST = 580_000
_W_SAFE_NOFD = 560_000
_W_STREAK_EXCESS = 200_000
_W_N_QUAD = 400_000
_W_N_GAP_MIN = 450_000
_W_N_BLK_ATTACH = 420_000
# 함께 근무 불가: 모두 소프트(구 하드 구간 포함)
_W_FORBIDDEN_PAIR_SOFT = 5_000_000
_FORBIDDEN_PAIR_HARD_MIN_NURSES = 12  # 레거시·미사용(항상 소프트)
# 일일 인원(E/N/D) 하드 이후 부가 패턴(공가 연속·고립 N 등) 소프트
_W_AUX_PATTERN_SOFT = 95_000
_W_MIN_WORK_DAYS = 8
_W_MIN_WORK_SOFT = 35_000
_W_WEEKLY_REST_SOFT = 4_000
# 전날 N + 당일 OF (N-OF) 지양 — 소프트(일일 인원 하드보다 우선순위 낮음).
_W_NOF_BEFORE_OF_SOFT = 110_000
_W_NOF_BEFORE_REQ_OF_SOFT = 210_000

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


def _build_n_block_boundary_maps(
    model: Any,
    x: dict,
    regular: list,
    num_days: int,
    carry_in: Any,
    num_nurses: int,
) -> dict[int, tuple[dict[int, Any], dict[int, Any]]]:
    """각 간호사별 N 블록 '말일 종료'·'첫일 시작' 0/1 표현 (변수 한 벌만 생성)."""
    out: dict[int, tuple[dict[int, Any], dict[int, Any]]] = {}
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
                start_blk[d] = 0 if carry_tail else xd
            else:
                if (n, d - 1, 'N') in x:
                    sb = model.NewBoolVar(f'nbst_{n}_{d}')
                    model.Add(sb <= xd)
                    model.Add(sb + x[n, d - 1, 'N'] <= 1)
                    model.Add(sb >= xd - x[n, d - 1, 'N'])
                    start_blk[d] = sb
                else:
                    start_blk[d] = xd
        out[n] = (end_blk, start_blk)
    return out


def _add_n_block_min_gap_hard(
    model: Any,
    x: dict,
    regular: list,
    num_days: int,
    min_gap_days: int,
    carry_in: Any,
    num_nurses: int,
    boundary_maps: dict[int, tuple[dict[int, Any], dict[int, Any]]],
) -> None:
    """
    연속 N 블록 사이 비N 일수 ≥ min_gap_days (gap = 다음첫N − 이전말N − 1).
    전월 말 N 이고 당월 1일이 연속이 아니면, 첫 당월 N은 (min_gap+1)일 이전 금지.
    """
    for n in regular:
        end_blk, start_blk = boundary_maps[n]
        carry_tail = _carry_prev_is_n(carry_in, n, num_nurses)
        for d1 in range(1, num_days + 1):
            e1 = end_blk.get(d1)
            if e1 is None or (type(e1) is int and e1 == 0):
                continue
            d2_max = min(num_days, d1 + min_gap_days)
            for d2 in range(d1 + 2, d2_max + 1):
                e2 = start_blk.get(d2)
                if e2 is None or (type(e2) is int and e2 == 0):
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


def _add_n_block_gap_spread_soft(
    model: Any,
    regular: list,
    num_days: int,
    boundary_maps: dict[int, tuple[dict[int, Any], dict[int, Any]]],
    tight_a: list[Any],
    tight_b: list[Any],
) -> None:
    """
    비N 간격이 하드(5일)만·6일만인 연속 블록 쌍에 벌점 Bool(말일 e1 ∧ 첫일 e2).
    하드가 아닌 소프트로만 7일 이상 간격을 선호.
    """
    off_a, off_b = 6, 7
    for n in regular:
        end_blk, start_blk = boundary_maps[n]
        for d1 in range(1, num_days + 1):
            e1 = end_blk.get(d1)
            if e1 is None or (type(e1) is int and e1 == 0):
                continue
            for off, bucket in ((off_a, tight_a), (off_b, tight_b)):
                d2 = d1 + off
                if d2 > num_days:
                    continue
                e2 = start_blk.get(d2)
                if e2 is None or (type(e2) is int and e2 == 0):
                    continue
                bucket.append(_bool_and2(model, e1, e2, f'ngapsp_{n}_{d1}_{off}'))


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


def _requests_clamped_to_nurses(requests: dict | None, num_nurses: int) -> dict | None:
    """부서 단위 인원(num_nurses) 밖의 간호사 키는 제외(타 부서 혼선 방지)."""
    if not requests or not isinstance(requests, dict):
        return requests
    out: dict[int, Any] = {}
    for k, v in requests.items():
        if isinstance(k, bool):
            continue
        try:
            ni = int(k)
        except (TypeError, ValueError):
            continue
        if 0 <= ni < num_nurses:
            out[ni] = v
    return out or None


def _req_shift_at(req_norm: dict[int, dict[int, str]], ni: int, dni: int) -> str | None:
    ds = req_norm.get(ni)
    if not ds:
        return None
    v = ds.get(dni)
    if v is None:
        v = ds.get(str(dni))
    if v is None:
        return None
    return str(v).strip()


def _allowed_shifts_cell(
    ni: int,
    dni: int,
    req_norm: dict[int, dict[int, str]],
    holiday_days: frozenset[int],
    preg_set: frozenset[int],
) -> list[str]:
    """진단·모델 공통: 해당 셀 허용 시프트 목록(자동 배정 시 임산부는 N 제외)."""
    s_req = _req_shift_at(req_norm, ni, dni)
    if s_req:
        if s_req == 'OH' and dni not in holiday_days:
            return ['OF']
        return [s_req]
    out: list[str] = []
    for s in _AUTO_ASSIGN_SHIFTS:
        if s == 'OH' and dni not in holiday_days:
            continue
        if s == 'N' and ni in preg_set:
            continue
        out.append(s)
    return out or ['OF']


def _diagnose_hard_infeasibility(
    days: list,
    num_nurses: int,
    regular: list[int],
    head: dict[int, str],
    req_norm: dict[int, dict[int, str]],
    holiday_days: frozenset[int],
    preg_set: frozenset[int],
    names: list[str],
    unit_profile: str = 'ward',
) -> str:
    """신청·임산부 반영 후 일일 E/N 최소·D 하한을 채울 수 있는지 사전 점검."""
    up = (unit_profile or 'ward').strip().lower()
    if up not in ('icu', 'er', 'ward'):
        up = 'ward'
    for day in days:
        d = int(day['day'])
        hsh = head.get(d) or ''
        need_e, need_n, (lo, hi) = app.daily_regular_staff_targets(
            num_nurses, day, hsh, up,
        )
        doms = []
        for n in regular:
            doms.append(
                frozenset(_allowed_shifts_cell(n, d, req_norm, holiday_days, preg_set)),
            )
        max_e = sum(1 for dom in doms if 'E' in dom)
        max_n = sum(1 for dom in doms if 'N' in dom)
        max_d = sum(1 for dom in doms if 'D' in dom)
        if max_e < need_e:
            return (
                f'{d}일({day["weekday_name"]}): 이브닝(E) 최소 {need_e}명인데, '
                f'신청 반영 후 E 가능 인원은 {max_e}명뿐입니다.'
            )
        if max_n < need_n:
            return (
                f'{d}일({day["weekday_name"]}): 나이트(N) 최소 {need_n}명인데, '
                f'신청 반영 후 N 가능 인원은 {max_n}명뿐입니다.'
            )
        if max_d < lo:
            return (
                f'{d}일({day["weekday_name"]}): 데이(D) 최소 {lo}명(상한 {hi}, 수간 {hsh or "—"})인데, '
                f'신청 반영 후 D 가능 인원은 {max_d}명뿐입니다.'
            )
    return (
        f'{names[0] if names else "팀"} 기준으로 일일 인원·신청 조합은 가능하나, '
        'N-OF-D·연속 근무 등 소프트 제약과 동시 만족되는 해가 없습니다.'
    )


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


def _bool_and2(model: Any, a: Any, b: Any, name: str) -> Any:
    """a∧b (0/1 선형식·상수)."""
    r = model.NewBoolVar(name)
    model.Add(r <= a)
    model.Add(r <= b)
    model.Add(r >= a + b - 1)
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


def _x_has_off_shift_keys(x: dict, n: int, d: int) -> bool:
    """해당 일에 OF/OH 결정변수가 있는지(파이썬 dict만 검사; OR-Tools 식과 비교하지 않음)."""
    return (n, d, 'OF') in x or (n, d, 'OH') in x


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
        for off_s, bad_s in (
            ('OF', 'D'), ('OF', 'EDU'), ('OF', '공'),
            ('OH', 'D'), ('OH', 'EDU'), ('OH', '공'),
            ('NO', 'D'), ('NO', 'EDU'), ('NO', '공'),
        ):
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
            if _x_has_off_shift_keys(x, n, 1) and _x_has_off_shift_keys(x, n, 2):
                o1c = _off_of_or_oh(model, x, n, 1, f'nrr_o1_{n}_c')
                o2c = _off_of_or_oh(model, x, n, 2, f'nrr_o2_{n}_c')
                r_ofof.append(_bool_and3(model, 1, o1c, o2c, f'nrr_of_{n}_c'))
            if _x_has_off_shift_keys(x, n, 1) and (n, 2, 'E') in x:
                o1e = _off_of_or_oh(model, x, n, 1, f'nrr_o1e_{n}_c')
                r_ofe.append(_bool_and3(model, 1, o1e, x[n, 2, 'E'], f'nrr_e_{n}_c'))
        for d in range(1, num_days - 1):
            if (n, d, 'N') not in x or not _x_has_off_shift_keys(x, n, d + 1):
                continue
            o1 = _off_of_or_oh(model, x, n, d + 1, f'nrr_o1_{n}_{d}')
            a_n = x[n, d, 'N']
            if _x_has_off_shift_keys(x, n, d + 2):
                o2 = _off_of_or_oh(model, x, n, d + 2, f'nrr_o2_{n}_{d}')
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


def _streak_terms_from_x(n: int, dn: int, x: dict) -> list:
    return [x[n, dn, s] for s in ('D', 'E', 'N', 'EDU', '공') if (n, dn, s) in x]


def _restore_sched_from_values(
    x: dict,
    val_map: dict,
    regular: list[int],
    num_days: int,
    allowed_for_cell_fn: Any,
) -> dict[int, dict[int, str]] | None:
    """BoolVar 할당 맵에서 일반 간호사 일별 시프트 복원."""
    sched: dict[int, dict[int, str]] = {}
    for n in regular:
        sched[n] = {}
        for d in range(1, num_days + 1):
            chosen = None
            for s in allowed_for_cell_fn(n, d):
                if (n, d, s) in x and val_map.get((n, d, s), 0) == 1:
                    chosen = s
                    break
            if chosen is None:
                return None
            sched[n][d] = chosen
    return sched


def _sched_fallback_requests_of(
    head: dict[int, str],
    req_norm: dict,
    regular: list[int],
    num_days: int,
) -> dict[int, dict[int, str]]:
    """솔버가 해를 못 내면: 신청 유지, 빈 칸은 OF."""
    sched: dict[int, dict[int, str]] = {0: dict(head)}
    for n in regular:
        sched[n] = {}
        for d in range(1, num_days + 1):
            rs = _req_shift_at(req_norm, n, d)
            sched[n][d] = rs if rs else 'OF'
    return sched


def solve_schedule_cpsat(
    num_nurses: int,
    requests: dict | None,
    holidays: tuple | list = (),
    forbidden_pairs: Any = None,
    carry_in: dict | None = None,
    carry_next_month: Any = None,
    shift_bans: dict | None = None,
    not_available: Any = None,
    pregnant_nurses: Any = None,
    nurse_names: Any = None,
    regenerate: bool = False,
    rng_seed: Any = None,
    unit_profile: str = 'ward',
) -> tuple[dict | None, bool, str, list[dict]]:
    """
    일당 1시프트·신청·일별 인원 하드 + N-OF 등 소프트. 해 없으면 폴백 표. 반환 4번째: `validate_schedule` 위반 목록.

    regenerate=True이면 rng_seed를 스크램블해 매번 다른 탐색. not_available 등은 API 동기화.
    """
    from ortools.sat.python import cp_model

    class _BestSolutionCollector(cp_model.CpSolverSolutionCallback):
        def __init__(self, x_map: dict):
            cp_model.CpSolverSolutionCallback.__init__(self)
            self._items = list(x_map.items())
            self.best_obj: float | None = None
            self.best_values: dict | None = None

        def on_solution_callback(self):
            obj = self.ObjectiveValue()
            if self.best_obj is None or obj < self.best_obj:
                self.best_obj = obj
                self.best_values = {k: self.Value(v) for k, v in self._items}

    days = app.get_april_days(holidays)
    num_days = app.NUM_DAYS
    n_abs_max = app.N_ABS_MAX
    holiday_days = frozenset(d['day'] for d in days if d['is_holiday'])

    if num_nurses < 2:
        return None, False, 'CP-SAT: 간호사 인원이 부족합니다.', []

    requests = _requests_clamped_to_nurses(requests, num_nurses)

    _solver_seed: int | None = None
    try:
        _raw = int(rng_seed) if rng_seed is not None else 0
    except (TypeError, ValueError, OverflowError):
        _raw = 0
    if regenerate:
        # 재생성마다 다른 분기 탐색(동일 입력이라도 패턴 변화)
        _solver_seed = (_raw * 1_103_515_245 + 12_345 + (num_nurses * 7919) + app.MONTH * 193) & 0x7FFFFFFF
        if _solver_seed == 0:
            _solver_seed = 1
    elif rng_seed is not None:
        _solver_seed = _raw & 0x7FFFFFFF

    fp_map = app._normalize_forbidden_pairs(forbidden_pairs, num_nurses)
    den_bans = app._normalize_shift_bans(shift_bans, num_nurses)
    if nurse_names is not None and len(nurse_names) == num_nurses:
        _names = [str(nm) for nm in nurse_names]
    else:
        _names = app.get_nurse_names(num_nurses)
    na_frozen = app._normalize_not_available(not_available, num_nurses, nurse_names=_names)
    preg_set = app._normalize_pregnant_nurses(pregnant_nurses, num_nurses, nurse_names=_names)
    req_norm = _normalize_requests(requests)
    # 단일 부서 단위 호출: 인덱스가 num_nurses 밖이면 다른 부서에서 섞인 값으로 간주하고 제외
    req_norm = {ni: ds for ni, ds in req_norm.items() if 0 <= ni < num_nurses}
    for ni in preg_set:
        ds = req_norm.get(ni, {})
        for dnk, shv in ds.items():
            if str(shv).strip() != 'N':
                continue
            try:
                dni = int(dnk)
            except (TypeError, ValueError):
                continue
            return None, False, (
                f'【절대 규칙】{_names[ni]}님은 임산부로 나이트(N) 신청·배정이 불가합니다.'
            ), []
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
                ), []
    carry_next_provided = carry_next_month is not None
    carry_next = app._normalize_carry_in(carry_next_month, num_nurses) if carry_next_month else {}
    month_first = date(app.YEAR, app.MONTH, 1)
    month_last = date(app.YEAR, app.MONTH, num_days)

    _uprof = (unit_profile or 'ward').strip().lower()
    if _uprof not in ('icu', 'er', 'ward'):
        _uprof = 'ward'
    _solve_time_sec = _cpsat_solver_max_seconds(_uprof, num_nurses)

    head = _build_head_schedule(days, requests, num_nurses)
    sched: dict[int, dict[int, str]] = {0: dict(head)}
    regular = list(range(1, num_nurses))
    num_reg = len(regular)
    n_gap_suffix = f' (야간 블록 간격·{_N_BLOCK_GAP_MIN_DAYS}일 목표·벌점 완화)'

    def allowed_for_cell(ni: int, dn: int) -> list[str]:
        return _allowed_shifts_cell(ni, dn, req_norm, holiday_days, preg_set)

    total_n_slots = 2 * num_days
    n_targets = app._compute_n_targets_fair(num_reg, total_n_slots, n_abs_max)
    tgt_map = {regular[i]: n_targets[i] for i in range(num_reg)}
    floor_n_avg = total_n_slots // num_reg if num_reg > 0 else 0
    n_cap_hard = min(app.N_ABS_MAX, floor_n_avg + 1) if num_reg > 0 else app.N_ABS_MAX

    model = cp_model.CpModel()
    x: dict[tuple[int, int, str], Any] = {}
    for n in regular:
        for d in range(1, num_days + 1):
            allowed = allowed_for_cell(n, d)
            for s in allowed:
                x[n, d, s] = model.NewBoolVar(f'x_{n}_{d}_{s}')
            model.Add(sum(x[n, d, s] for s in allowed) == 1)

    # 신청 칸 100% 고정(연·공·D 등 표에 넣은 값 그대로).
    for n in regular:
        for d in range(1, num_days + 1):
            rs = _req_shift_at(req_norm, n, d)
            if rs and (n, d, rs) in x:
                model.Add(x[n, d, rs] == 1)

    obj_terms: list[Any] = []

    _w_uv = _W_TIER2
    for n in regular:
        banned = den_bans.get(n, frozenset())
        for s in banned:
            for d in range(1, num_days + 1):
                if (n, d, s) not in x:
                    continue
                if _req_shift_at(req_norm, n, d) == s:
                    continue
                obj_terms.append(_w_uv * x[n, d, s])
    for (n, d, s) in na_frozen:
        if n not in regular:
            continue
        if (n, d, s) not in x:
            continue
        if _req_shift_at(req_norm, n, d) == s:
            continue
        obj_terms.append(_w_uv * x[n, d, s])

    # N-(다음날)OF 패턴 지양: d-1일 N 이고 d일 OF이면 소프트 위반(신청 OF인 d일이면 벌점↑, D/E 다음 OF보다 덜 선호되게)
    for n in regular:
        if _carry_prev_is_n(carry_in, n, num_nurses) and (n, 1, 'OF') in x:
            _w_c = (
                _W_NOF_BEFORE_REQ_OF_SOFT
                if _req_shift_at(req_norm, n, 1) == 'OF'
                else _W_NOF_BEFORE_OF_SOFT
            )
            obj_terms.append(_w_c * x[n, 1, 'OF'])
        for d in range(2, num_days + 1):
            if (n, d - 1, 'N') not in x or (n, d, 'OF') not in x:
                continue
            _req_of_d = _req_shift_at(req_norm, n, d) == 'OF'
            s_nof = model.NewIntVar(0, 1, f'nof_soft_{n}_{d}')
            model.Add(x[n, d - 1, 'N'] + x[n, d, 'OF'] <= 1 + s_nof)
            _w_n = _W_NOF_BEFORE_REQ_OF_SOFT if _req_of_d else _W_NOF_BEFORE_OF_SOFT
            obj_terms.append(_w_n * s_nof)

    # 일별 인원 — E/N은 부서 최소(이상 허용), D는 하한·상한(응급실 A1 평일 D=1 등).
    for d in range(1, num_days + 1):
        day = days[d - 1]
        hsh = head.get(d) or ''
        need_e, need_n, (lo, hi) = app.daily_regular_staff_targets(
            num_nurses, day, hsh, _uprof,
        )
        ev = [x[n, d, 'E'] for n in regular if (n, d, 'E') in x]
        nv = [x[n, d, 'N'] for n in regular if (n, d, 'N') in x]
        dv = [x[n, d, 'D'] for n in regular if (n, d, 'D') in x]
        e_sum = sum(ev) if ev else 0
        n_sum = sum(nv) if nv else 0
        d_sum = sum(dv) if dv else 0
        model.Add(e_sum >= need_e)
        model.Add(n_sum >= need_n)
        model.Add(d_sum >= lo)
        model.Add(d_sum <= hi)

    _mw_floor = min(_W_MIN_WORK_DAYS, num_days)
    for n in regular:
        wvars = [
            x[n, d, s]
            for d in range(1, num_days + 1)
            for s in ('D', 'E', 'N')
            if (n, d, s) in x
        ]
        if not wvars:
            continue
        ws = sum(wvars)
        slack_w = model.NewIntVar(0, num_days, f'minwork_{n}')
        model.Add(ws + slack_w >= _mw_floor)
        obj_terms.append(_W_MIN_WORK_SOFT * slack_w)

    # 월간 N: 상한 소프트(가변) + 형평
    n_tot_exprs: list[Any] = []
    for n in regular:
        nv = [x[n, d, 'N'] for d in range(1, num_days + 1) if (n, d, 'N') in x]
        if not nv:
            continue
        nt = sum(nv)
        n_tot_exprs.append(nt)
        over_cap = model.NewIntVar(0, num_days, f'nocap_{n}')
        model.Add(nt <= n_cap_hard + over_cap)
        obj_terms.append(_W_TIER2 * over_cap)
        over_abs = model.NewIntVar(0, num_days, f'noabs_{n}')
        model.Add(nt <= app.N_ABS_MAX + over_abs)
        obj_terms.append(_W_TIER2 * over_abs)
        nsq = model.NewIntVar(0, app.N_ABS_MAX * app.N_ABS_MAX, f'nsq_{n}')
        model.AddMultiplicationEquality(nsq, [nt, nt])
        obj_terms.append(_W_N_SUM_SQUARES * nsq)
        tgt = tgt_map.get(n, 0)
        dlt = model.NewIntVar(-num_days, num_days, f'ndev_{n}')
        model.Add(dlt == nt - tgt)
        nab = model.NewIntVar(0, num_days, f'nab_{n}')
        model.AddAbsEquality(nab, dlt)
        obj_terms.append(_W_N_TGT_DEV * nab)

    if n_tot_exprs:
        max_nt = model.NewIntVar(0, num_days, 'n_tot_max')
        min_nt = model.NewIntVar(0, num_days, 'n_tot_min')
        for nt in n_tot_exprs:
            model.Add(max_nt >= nt)
            model.Add(min_nt <= nt)
        n_spread = model.NewIntVar(0, num_days, 'n_spread')
        model.Add(n_spread == max_nt - min_nt)
        n_spread_excess = model.NewIntVar(0, num_days, 'n_spread_xs')
        model.Add(n_spread <= 1 + n_spread_excess)
        obj_terms.append(_W_N_SPREAD_EXCESS * n_spread_excess)

    # N-D·E직후금지·N-OF-D·말단 N 등 — 신청과 충돌 가능하므로 벌점 완화
    for n in regular:
        for d in range(2, num_days + 1):
            if (n, d, 'D') in x and (n, d - 1, 'N') in x:
                s_nd = model.NewIntVar(0, 1, f'nd_{n}_{d}')
                model.Add(x[n, d, 'D'] + x[n, d - 1, 'N'] <= 1 + s_nd)
                obj_terms.append(_W_SAFE_N_D * s_nd)
        if (n, 1, 'D') in x and _carry_prev_is_n(carry_in, n, num_nurses):
            s_nd1 = model.NewIntVar(0, 1, f'nd1_{n}')
            model.Add(x[n, 1, 'D'] <= s_nd1)
            obj_terms.append(_W_SAFE_N_D * s_nd1)

    # E 직후: D·EDU·公 은 소프트(인력 하드와 충돌 시 완화).
    _bad_after_e_soft = ('D', 'EDU')
    for n in regular:
        for d in range(2, num_days + 1):
            if (n, d - 1, 'E') not in x:
                continue
            ep = x[n, d - 1, 'E']
            for bad in _bad_after_e_soft:
                if (n, d, bad) in x:
                    s_e = model.NewIntVar(0, 1, f'eaf_{n}_{d}_{bad}')
                    model.Add(ep + x[n, d, bad] <= 1 + s_e)
                    obj_terms.append(_W_SAFE_E_AFTER * s_e)
            if (n, d, '공') in x:
                s_eg = model.NewIntVar(0, 1, f'e_공_{n}_{d}')
                model.Add(ep + x[n, d, '공'] <= 1 + s_eg)
                obj_terms.append(_W_AUX_PATTERN_SOFT * s_eg)
        if _carry_prev_is_e(carry_in, n, num_nurses):
            for bad in _bad_after_e_soft:
                if (n, 1, bad) in x:
                    s_e1 = model.NewIntVar(0, 1, f'eaf1_{n}_{bad}')
                    model.Add(x[n, 1, bad] <= s_e1)
                    obj_terms.append(_W_SAFE_E_AFTER * s_e1)
            if (n, 1, '공') in x:
                obj_terms.append(_W_AUX_PATTERN_SOFT * x[n, 1, '공'])

    day1_hol = bool(days[0]['is_holiday'])
    _dn_holiday = {d['day']: bool(d['is_holiday']) for d in days}
    for n in regular:
        if _carry_prev_is_n(carry_in, n, num_nurses) and (n, 1, 'N') in x:
            need = 'OH' if day1_hol else 'OF'
            if (n, 1, need) in x:
                s_nr = model.NewIntVar(0, 1, f'nrest1_{n}')
                model.Add(x[n, 1, need] + s_nr >= x[n, 1, 'N'])
                obj_terms.append(_W_SAFE_N_REST * s_nr)
        for d in range(1, num_days):
            if (n, d, 'N') not in x:
                continue
            need = 'OH' if _dn_holiday.get(d + 1) else 'OF'
            if (n, d + 1, need) not in x:
                continue
            xn1 = x[n, d + 1, 'N'] if (n, d + 1, 'N') in x else 0
            s_nr = model.NewIntVar(0, 1, f'nrest_{n}_{d}')
            model.Add(x[n, d + 1, need] + s_nr >= x[n, d, 'N'] - xn1)
            obj_terms.append(_W_SAFE_N_REST * s_nr)

    # N 직후 공가 — 고액 벌점 (N-OF-OF-공 허용)
    for n in regular:
        for d in range(1, num_days):
            if (n, d, 'N') not in x or (n, d + 1, '공') not in x:
                continue
            s_npg = model.NewIntVar(0, 1, f'npg_{n}_{d}')
            model.Add(x[n, d, 'N'] + x[n, d + 1, '공'] <= 1 + s_npg)
            obj_terms.append(_W_AUX_PATTERN_SOFT * s_npg)
        if _carry_prev_is_n(carry_in, n, num_nurses) and (n, 1, '공') in x:
            obj_terms.append(_W_AUX_PATTERN_SOFT * x[n, 1, '공'])

    for n in regular:
        for end in range(1, num_days - 1):
            if (n, end, 'N') not in x:
                continue
            v_n = x[n, end, 'N']
            for off_s, bad_s in (
                ('OF', 'D'), ('OF', 'EDU'), ('OF', '공'),
                ('OH', 'D'), ('OH', 'EDU'), ('OH', '공'),
                ('NO', 'D'), ('NO', 'EDU'), ('NO', '공'),
            ):
                if (n, end + 1, off_s) not in x:
                    continue
                v_mid = x[n, end + 1, off_s]
                k_bad = (n, end + 2, bad_s)
                if k_bad in x:
                    s_trip = model.NewIntVar(0, 1, f'ntrip_{n}_{end}_{off_s}_{bad_s}')
                    model.Add(v_n + v_mid + x[k_bad] <= 2 + s_trip)
                    w_t = _W_AUX_PATTERN_SOFT if bad_s == '공' else _W_SAFE_NOFD
                    obj_terms.append(w_t * s_trip)

    if num_days >= 2:
        for n in regular:
            if not _carry_prev_is_n(carry_in, n, num_nurses):
                continue
            for off_s, bad_s in (
                ('OF', 'D'), ('OF', 'EDU'), ('OF', '공'),
                ('OH', 'D'), ('OH', 'EDU'), ('OH', '공'),
                ('NO', 'D'), ('NO', 'EDU'), ('NO', '공'),
            ):
                if (n, 1, off_s) in x and (n, 2, bad_s) in x:
                    s_ct = model.NewIntVar(0, 1, f'ntrc_{n}_{off_s}_{bad_s}')
                    model.Add(1 + x[n, 1, off_s] + x[n, 2, bad_s] <= 2 + s_ct)
                    w_c = _W_AUX_PATTERN_SOFT if bad_s == '공' else _W_SAFE_NOFD
                    obj_terms.append(w_c * s_ct)

    for n in regular:
        for d in range(1, num_days):
            if (n, d, 'N') not in x:
                continue
            if _req_shift_at(req_norm, n, d) == 'N':
                continue
            rhs_terms = []
            if d > 1 and (n, d - 1, 'N') in x:
                rhs_terms.append(x[n, d - 1, 'N'])
            elif d == 1 and _carry_prev_is_n(carry_in, n, num_nurses):
                rhs_terms.append(1)
            if (n, d + 1, 'N') in x:
                rhs_terms.append(x[n, d + 1, 'N'])
            if not rhs_terms:
                obj_terms.append(_W_AUX_PATTERN_SOFT * x[n, d, 'N'])
            else:
                s_blk = model.NewIntVar(0, 1, f'nblk_{n}_{d}')
                model.Add(x[n, d, 'N'] <= sum(rhs_terms) + s_blk)
                obj_terms.append(_W_N_BLK_ATTACH * s_blk)

    for n in regular:
        for d in range(1, num_days - 2):
            quad = [x[n, d + k, 'N'] for k in range(4) if (n, d + k, 'N') in x]
            if len(quad) == 4:
                s4 = model.NewIntVar(0, 1, f'n4_{n}_{d}')
                model.Add(sum(quad) <= 3 + s4)
                obj_terms.append(_W_N_QUAD * s4)

    n_gap_tight_a: list[Any] = []
    n_gap_tight_b: list[Any] = []
    nb_bounds = _build_n_block_boundary_maps(
        model, x, regular, num_days, carry_in, num_nurses,
    )
    min_gap = _N_BLOCK_GAP_MIN_DAYS
    for n in regular:
        end_blk, start_blk = nb_bounds[n]
        carry_tail = _carry_prev_is_n(carry_in, n, num_nurses)
        for d1 in range(1, num_days + 1):
            e1 = end_blk.get(d1)
            if e1 is None or (type(e1) is int and e1 == 0):
                continue
            d2_max = min(num_days, d1 + min_gap)
            for d2 in range(d1 + 2, d2_max + 1):
                e2 = start_blk.get(d2)
                if e2 is None or (type(e2) is int and e2 == 0):
                    continue
                mid = [x[n, t, 'N'] for t in range(d1 + 1, d2) if (n, t, 'N') in x]
                sm = sum(mid) if mid else 0
                s_gap = model.NewIntVar(0, 1, f'ng_{n}_{d1}_{d2}')
                model.Add(e1 + e2 <= 1 + sm + s_gap)
                obj_terms.append(_W_N_GAP_MIN * s_gap)
        if carry_tail:
            for s in range(2, min(num_days, min_gap) + 1):
                if (n, s, 'N') not in x:
                    continue
                if _req_shift_at(req_norm, n, s) == 'N':
                    continue
                preds = [x[n, t, 'N'] for t in range(1, s) if (n, t, 'N') in x]
                if preds:
                    s_tl = model.NewIntVar(0, 1, f'ntl_{n}_{s}')
                    model.Add(x[n, s, 'N'] <= sum(preds) + s_tl)
                    obj_terms.append(_W_N_GAP_MIN * s_tl)
                else:
                    obj_terms.append(_W_N_GAP_MIN * x[n, s, 'N'])

    _add_n_block_gap_spread_soft(
        model, regular, num_days, nb_bounds, n_gap_tight_a, n_gap_tight_b,
    )

    carry_norm = app._normalize_carry_in(carry_in, num_nurses)

    streak_win_days = 6
    streak_cap = 5
    for n in regular:
        carry_seq = list(carry_norm.get(n) or ())
        lcarry = len(carry_seq)
        span = lcarry + num_days
        if span < streak_win_days:
            continue
        for w in range(span - (streak_win_days - 1)):
            const = 0
            terms: list = []
            for k in range(streak_win_days):
                idx = w + k
                if idx < lcarry:
                    const += 1 if carry_seq[idx] in STREAK_WORK_SHIFTS else 0
                else:
                    dn = idx - lcarry + 1
                    terms.extend(_streak_terms_from_x(n, dn, x))
            if terms:
                s_st = model.NewIntVar(0, streak_win_days, f'str6_{n}_{w}')
                model.Add(const + sum(terms) <= streak_cap + s_st)
                obj_terms.append(_W_STREAK_EXCESS * s_st)
            elif const > streak_cap:
                s_st = model.NewIntVar(0, streak_win_days, f'str6c_{n}_{w}')
                model.Add(const <= streak_cap + s_st)
                obj_terms.append(_W_STREAK_EXCESS * s_st)

    pond_penalty_terms: list[Any] = []
    _add_ponddang_streak_soft(
        model, x, regular, num_days,
        carry_norm, carry_next, carry_next_provided,
        req_norm, set(), num_nurses,
        pond_penalty_terms,
    )

    low_weekly: list[Any] = []
    wk_map: dict[int, list] = {}
    for day in days:
        sun = app._week_sunday(day['date'])
        wk_map.setdefault(sun.toordinal(), []).append(day['day'])

    for _wk, wdays in wk_map.items():
        if not wdays:
            continue
        sun_date = date.fromordinal(_wk)
        for n in regular:
            pre_of, pre_oh, pre_no, n_prev = app._carry_week_prev_month_off_counts(
                carry_norm, n, sun_date, month_first,
            )
            post_of, post_oh, post_no, n_next = app._carry_week_next_month_off_counts(
                carry_next, n, sun_date, month_last,
            )
            pre_rest, _ = app._carry_week_prev_rest_total(
                carry_norm, n, sun_date, month_first,
            )
            post_rest, _ = app._carry_week_next_rest_total(
                carry_next, n, sun_date, month_last,
            )
            len_w = len(wdays)
            m_week = n_prev + len_w + n_next

            of_month = sum(x[n, d, 'OF'] for d in wdays if (n, d, 'OF') in x)
            oh_month = sum(x[n, d, 'OH'] for d in wdays if (n, d, 'OH') in x)
            no_month = sum(x[n, d, 'NO'] for d in wdays if (n, d, 'NO') in x)

            of_vis_expr = pre_of + of_month
            tot_of_add = post_of if carry_next_provided else 0
            of_sl = model.NewIntVar(0, 4, '')
            model.Add(of_vis_expr + tot_of_add <= 3 + of_sl)
            low_weekly.append(of_sl)

            no_week_expr = pre_no + no_month + post_no
            no_sl = model.NewIntVar(0, 3, '')
            model.Add(no_week_expr <= 1 + no_sl)
            low_weekly.append(no_sl)

            if m_week <= 0:
                continue
            if n_next > 0 and not carry_next_provided:
                continue
            cseq = list(carry_norm.get(n) or ())
            if n_prev > 0 and len(cseq) < n_prev:
                continue

            month_rest: list[Any] = []
            for d in wdays:
                for s in ('OF', 'OH', 'NO', '연', '공', '병', '경'):
                    if (n, d, s) in x:
                        month_rest.append(x[n, d, s])
            post_part = post_rest if carry_next_provided else 0
            rest_lin = pre_rest + post_part
            if month_rest:
                rest_lin = rest_lin + sum(month_rest)
            wk_sl = model.NewIntVar(0, 7, '')
            model.Add(rest_lin + wk_sl >= 2)
            low_weekly.append(wk_sl)

    if low_weekly:
        obj_terms.append(_W_WEEKLY_REST_SOFT * sum(low_weekly))

    # 함께 근무 불가 — 전량 고액 소프트(위반 시에도 해 유지)
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
                        if (b, dn, shift) not in x:
                            continue
                        v = model.NewBoolVar(f'fph_{b}_{dn}_{shift}')
                        model.Add(x[b, dn, shift] <= v)
                        obj_terms.append(_W_FORBIDDEN_PAIR_SOFT * v)
                    elif b == 0:
                        if head.get(dn) != shift:
                            continue
                        if (a, dn, shift) not in x:
                            continue
                        v = model.NewBoolVar(f'fph_{a}_{dn}_{shift}')
                        model.Add(x[a, dn, shift] <= v)
                        obj_terms.append(_W_FORBIDDEN_PAIR_SOFT * v)
                    else:
                        if (a, dn, shift) not in x or (b, dn, shift) not in x:
                            continue
                        v = model.NewBoolVar(f'fp_{a}_{b}_{dn}_{shift}')
                        model.Add(x[a, dn, shift] + x[b, dn, shift] <= 1 + v)
                        obj_terms.append(_W_FORBIDDEN_PAIR_SOFT * v)

    d_tot_exprs: list[Any] = []
    e_tot_exprs: list[Any] = []
    for n in regular:
        dv = [x[n, d, 'D'] for d in range(1, num_days + 1) if (n, d, 'D') in x]
        ev = [x[n, d, 'E'] for d in range(1, num_days + 1) if (n, d, 'E') in x]
        d_tot_exprs.append(sum(dv) if dv else 0)
        e_tot_exprs.append(sum(ev) if ev else 0)

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
        if not _CP_WEEKEND_OFF_EVEN_HARD:
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

    _flex_d_days: list[int] = []
    for d in range(1, num_days + 1):
        day = days[d - 1]
        hsh = head.get(d) or ''
        lo, hi = app.d_regular_d_bounds(num_nurses, day, hsh, _uprof)
        if lo < hi:
            _flex_d_days.append(d)

    obj_terms.append(_OBJ_FAIRNESS_PRIORITY * _W_LOW_FAIR * (diff_d + diff_e))
    if n_gap_tight_a:
        obj_terms.append(_PENALTY_N_GAP_TIGHT_A * sum(n_gap_tight_a))
    if n_gap_tight_b:
        obj_terms.append(_PENALTY_N_GAP_TIGHT_B * sum(n_gap_tight_b))
    if de_abs_terms:
        obj_terms.append(_OBJ_DE_MONTH_ABS_WEIGHT * _W_LOW_FAIR * sum(de_abs_terms))
    r_ofof, r_ofe = _collect_n_recovery_reward_vars(
        model, x, regular, num_days, carry_in, num_nurses,
    )
    if r_ofof:
        obj_terms.append(-_REWARD_N_OFOF * sum(r_ofof))
    if r_ofe:
        obj_terms.append(-_REWARD_N_OFE * sum(r_ofe))
    if wk_even_spread is not None:
        obj_terms.append(_WEEKEND_OFF_EVEN_SOFT_WEIGHT * _W_LOW_FAIR * wk_even_spread)
    if _flex_d_days:
        _obj_flex_d = [
            x[n, d, 'D']
            for d in _flex_d_days
            for n in regular
            if (n, d, 'D') in x
        ]
        if _obj_flex_d:
            if num_nurses == 13:
                obj_terms.append(-_PREFER_D3_WEIGHT_13 * _W_LOW_FAIR * sum(_obj_flex_d))
            else:
                obj_terms.append(_W_LOW_FAIR * sum(_obj_flex_d))
    if pond_penalty_terms:
        _pond_w = _POND_SOFT_WEIGHT_N10 if num_nurses == 10 else _POND_SOFT_WEIGHT
        obj_terms.append(_pond_w * sum(pond_penalty_terms))

    model.Minimize(sum(obj_terms) if obj_terms else 0)

    cb = _BestSolutionCollector(x)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(_solve_time_sec)
    solver.parameters.num_search_workers = max(1, os.cpu_count() or 1)
    if _solver_seed is not None:
        solver.parameters.random_seed = int(_solver_seed)
    status = solver.Solve(model, cb)

    val_map: dict | None = cb.best_values
    obj_val: float = float(cb.best_obj) if cb.best_obj is not None else 0.0
    if val_map is None and status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        try:
            val_map = {k: solver.Value(v) for k, v in x.items()}
            obj_val = float(solver.ObjectiveValue())
        except Exception:
            val_map = None

    _status_names = {
        cp_model.OPTIMAL: 'OPTIMAL',
        cp_model.FEASIBLE: 'FEASIBLE',
        cp_model.INFEASIBLE: 'INFEASIBLE',
        cp_model.UNKNOWN: 'UNKNOWN',
        cp_model.MODEL_INVALID: 'MODEL_INVALID',
    }
    status_name = _status_names.get(status, str(status))

    if val_map is not None:
        restored = _restore_sched_from_values(
            x, val_map, regular, num_days, allowed_for_cell,
        )
        if restored is not None:
            sched.update(restored)
        else:
            val_map = None

    _fb_used = False
    if val_map is None:
        sched = _sched_fallback_requests_of(head, req_norm, regular, num_days)
        _fb_used = True
        obj_val = 0.0
        if status == cp_model.INFEASIBLE:
            _hint = _diagnose_hard_infeasibility(
                days, num_nurses, regular, head, req_norm,
                holiday_days, preg_set, _names, _uprof,
            )
            status_name = f'{status_name}·완화폴백(신청+빈칸OF)·참고:{_hint}'
        elif status == cp_model.UNKNOWN:
            status_name = (
                f'{status_name}·완화폴백(신청+빈칸OF)·'
                f'{_solve_time_sec:g}초 내 최적해 없음'
            )
        else:
            status_name = f'{status_name}·완화폴백(신청+빈칸OF)'

    issues = app.validate_schedule(
        sched, num_nurses, holidays,
        forbidden_pairs=forbidden_pairs, carry_in=carry_in,
        requests=requests, carry_next_month=carry_next_month,
        shift_bans=shift_bans,
        not_available=not_available,
        nurse_names=_names,
        engine_soft_report=False,
        unit_profile=_uprof,
    )
    warn_n = sum(1 for z in issues if z.get('level') == 'warn')
    err_n = sum(1 for z in issues if z.get('level') == 'error')
    tail = '폴백·목적N/A' if _fb_used else f'목적≈{obj_val:.0f}'
    status_str = (
        f'CP-SAT [{status_name}] {tail}·한도 {_solve_time_sec:g}s·'
        f'검토 오류 {err_n}·경고 {warn_n}건{n_gap_suffix}'
    )
    return sched, True, status_str, issues


def solve_schedule(
    num_nurses: int,
    requests: dict | None,
    holidays: tuple | list = (),
    forbidden_pairs: Any = None,
    carry_in: dict | None = None,
    regenerate: bool = False,
    rng_seed: Any = None,
    nurse_names: Any = None,
    carry_next_month: Any = None,
    shift_bans: dict | None = None,
    not_available: Any = None,
    pregnant_nurses: Any = None,
    unit_profile: str = 'ward',
) -> tuple[dict | None, bool, str, list[dict]]:
    """
    `app.solve_schedule`·Streamlit 호출부와 동일한 인자 이름·순서.
    내부적으로 `solve_schedule_cpsat`로 위임한다.

    호출 시 requests·carry_in 등은 **한 부서(인덱스 0…num_nurses-1)** 기준이어야 한다.
    인덱스 범위 밖 키는 솔버에서 자동으로 제외한다.
    """
    return solve_schedule_cpsat(
        num_nurses,
        requests,
        holidays,
        forbidden_pairs=forbidden_pairs,
        carry_in=carry_in,
        carry_next_month=carry_next_month,
        shift_bans=shift_bans,
        not_available=not_available,
        pregnant_nurses=pregnant_nurses,
        nurse_names=nurse_names,
        regenerate=regenerate,
        rng_seed=rng_seed,
        unit_profile=unit_profile,
    )
