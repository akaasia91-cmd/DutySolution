"""
응급실 근무표 생성기 – Streamlit UI v2
- 연도·월 선택 가능
- 부서(Department) CRUD
- 간호사(Staff) CRUD: 추가 / 이름 수정 / 삭제
- 부서별 신청 근무 입력 달력 (data_editor)
- 함께 근무 불가 쌍 (같은 날 D/E/N 동시 배치)
- 부서별 근무표 생성 + 컬러 테이블 + 엑셀 다운로드
- st.session_state 영속 저장
"""

import streamlit as st
import pandas as pd
import io
import json
import calendar as _calendar
from pathlib import Path

import app as _app                          # 전역 상수(YEAR/MONTH/NUM_DAYS) 동적 갱신
from app import (
    solve_schedule, get_april_days, build_stats, validate_schedule,
    SHIFT_NAMES, SHIFT_COLORS, SHIFT_TEXT_COLORS,
)
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ════════════════════════════════════════════════════════════════════════════════
#  페이지 설정
# ════════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="응급실 근무표 생성기",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ════════════════════════════════════════════════════════════════════════════════
#  전역 CSS
# ════════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
/* 브라우저가 폼 컨트롤을 다크모드로 그리지 않게 */
.stApp, section[data-testid="stSidebar"], section[data-testid="stSidebar"] input {
    color-scheme: light !important;
}

.stApp { background-color: #F0F2F6; }

/* 사이드바 — Streamlit CSS 변수(다크 텍스트 색이 입력에 전달되도록) */
section[data-testid="stSidebar"] {
    --text-color: #262730 !important;
    --stTextColor: #262730 !important;
    --widget-text-color: #000000 !important;
}

/* 사이드바 — 흰색 배경 + 선명한 검정 계열 글자 */
section[data-testid="stSidebar"] > div:first-child {
    background: #ffffff !important;
    border-right: 1px solid #e0e0e0;
}
section[data-testid="stSidebar"] {
    color: #212121 !important;
}
section[data-testid="stSidebar"] .stMarkdown,
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown li,
section[data-testid="stSidebar"] .stMarkdown h1,
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3,
section[data-testid="stSidebar"] .stMarkdown h4 {
    color: #212121 !important;
}
section[data-testid="stSidebar"] hr { border-color: #e0e0e0 !important; margin:0.6rem 0; }

/* 연도·월 2열 — 한 겹 겹친 느낌 제거(그림자·반투명·블러 없음) */
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"],
section[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] {
    opacity: 1 !important;
    filter: none !important;
    box-shadow: none !important;
}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] > [data-testid="stElementContainer"],
section[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] > [data-testid="stElementContainer"] {
    opacity: 1 !important;
}
section[data-testid="stSidebar"] [data-testid="column"] {
    background: transparent !important;
    box-shadow: none !important;
}
section[data-testid="stSidebar"] [data-testid="stElementContainer"] {
    box-shadow: none !important;
}

section[data-testid="stSidebar"] [data-testid="stSelectbox"] label p,
section[data-testid="stSidebar"] [data-testid="stSelectbox"] label span,
section[data-testid="stSidebar"] .stTextInput label p,
section[data-testid="stSidebar"] .stTextInput label span {
    color: #000000 !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    opacity: 1 !important;
    -webkit-font-smoothing: antialiased;
}

section[data-testid="stSidebar"] [data-testid="stTextInput"] input,
section[data-testid="stSidebar"] [data-testid="stTextInput"] textarea,
section[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stTextInput"] input,
section[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stTextInput"] textarea {
    background: #fafafa !important;
    border: 1px solid #bdbdbd !important;
    color: #0d1117 !important;
    -webkit-text-fill-color: #0d1117 !important;
    caret-color: #0d1117 !important;
    opacity: 1 !important;
    border-radius: 8px;
}
section[data-testid="stSidebar"] [data-testid="stTextInput"] input::placeholder,
section[data-testid="stSidebar"] [data-testid="stTextInput"] textarea::placeholder,
section[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stTextInput"] input::placeholder {
    color: #616161 !important;
    -webkit-text-fill-color: #616161 !important;
    opacity: 1 !important;
}

/* 모든 사이드바 텍스트 입력 — Streamlit은 .stTextInput + data-baseweb="input" 조합 사용 */
section[data-testid="stSidebar"] .stTextInput input,
section[data-testid="stSidebar"] .stTextInput textarea,
section[data-testid="stSidebar"] [data-baseweb="input"] input,
section[data-testid="stSidebar"] [data-baseweb="textarea"] textarea {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    caret-color: #000000 !important;
    background-color: #ffffff !important;
    opacity: 1 !important;
}
section[data-testid="stSidebar"] .stTextInput input:-webkit-autofill,
section[data-testid="stSidebar"] [data-baseweb="input"] input:-webkit-autofill {
    -webkit-text-fill-color: #000000 !important;
    caret-color: #000000 !important;
    box-shadow: 0 0 0px 1000px #ffffff inset !important;
}

/* Expander 안 입력(새 부서 추가 등) — 추가 보강 */
section[data-testid="stSidebar"] [data-testid="stExpander"] input:not([role="combobox"]),
section[data-testid="stSidebar"] [data-testid="stExpanderDetails"] input:not([role="combobox"]),
section[data-testid="stSidebar"] [data-testid="stExpander"] textarea,
section[data-testid="stSidebar"] [data-testid="stExpanderDetails"] textarea {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    caret-color: #000000 !important;
    background-color: #ffffff !important;
    opacity: 1 !important;
}

/* 사이드바 내 모든 텍스트형 input (타입 미지정 포함) — 콤보박스 제외 */
section[data-testid="stSidebar"] input[type="text"],
section[data-testid="stSidebar"] input[type="search"],
section[data-testid="stSidebar"] input:not([type]):not([role="combobox"]) {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    caret-color: #000000 !important;
    background-color: #ffffff !important;
}

/* 앱 전역: text_input 위젯은 항상 검정 글자 (사이드바 DOM 분리 대비) */
div[data-testid="stTextInput"] input,
.stTextInput input {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    caret-color: #000000 !important;
}

/* select — 순배경(회색끔 제거) + 순검정 글자, 덧씌운 느낌 나는 그림자 제거 */
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    background-color: #ffffff !important;
    background-image: none !important;
    border: 1.5px solid #616161 !important;
    border-radius: 6px !important;
    box-shadow: none !important;
    opacity: 1 !important;
    color: #000000 !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background-color: #ffffff !important;
    background-image: none !important;
    border: 1.5px solid #616161 !important;
    border-radius: 6px !important;
    box-shadow: none !important;
    opacity: 1 !important;
    color: #000000 !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] {
    color: #000000 !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    -webkit-font-smoothing: antialiased;
}
section[data-testid="stSidebar"] [data-baseweb="select"] > div > div {
    color: #000000 !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] span,
section[data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] div,
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] p {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
}
/* 사이드바 셀렉트: 좁은 열에서 '간…' 말줄임 방지 — 전체 이름 표시 */
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    min-width: 0 !important;
    width: 100% !important;
    max-width: 100% !important;
}
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"] p {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
}
/* 드롭다운 목록 항목 */
div[data-baseweb="popover"] li[role="option"] {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
}

section[data-testid="stSidebar"] [data-testid="stExpander"] summary,
section[data-testid="stSidebar"] [data-testid="stExpander"] summary span,
section[data-testid="stSidebar"] [data-testid="stExpander"] summary p {
    color: #212121 !important;
}
section[data-testid="stSidebar"] .stCaption { color: #616161 !important; }

/* 펼친 드롭다운 목록은 밝은 배경 + 검정 글자 (포털로 body에 그려질 수 있음) */
div[data-baseweb="popover"] ul[role="listbox"] li,
div[data-baseweb="popover"] li[role="option"] {
    color: #111111 !important;
    -webkit-text-fill-color: #111111 !important;
    background-color: #ffffff !important;
}
div[data-baseweb="popover"] ul[role="listbox"] {
    background-color: #ffffff !important;
}

/* 버튼 */
div.stButton > button[kind="primary"] {
    background:#2E7D32; border:none; border-radius:10px;
    font-size:15px; font-weight:700; height:52px;
}
div.stButton > button[kind="primary"]:hover { background:#1B5E20; }

/* 작은 보조 버튼 */
div.stButton > button[kind="secondary"] {
    border-radius:6px; font-size:12px; padding:2px 8px;
}

/* 섹션 카드 */
.card {
    background:white; border-radius:12px; padding:16px 20px;
    box-shadow:0 2px 10px rgba(0,0,0,0.07); margin-bottom:14px;
}
.card-title { font-size:17px; font-weight:800; color:#1A237E; margin-bottom:6px; }
.card-sub   { font-size:12px; color:#546E7A; margin-top:2px; }

/* 배지 */
.dept-badge {
    display:inline-block; background:#E8EAF6; color:#1A237E;
    border-radius:20px; padding:3px 12px; font-size:12px; font-weight:600; margin-right:4px;
}

/* 신청 근무·수정 모드 data_editor — 셀·선택창 소형화 */
div[data-testid="stDataFrame"] td,
div[data-testid="stDataFrame"] th {
    font-size: 11px !important;
    padding: 2px 4px !important;
}
div[data-testid="stDataFrame"] [data-baseweb="select"] > div {
    font-size: 11px !important;
    min-height: 26px !important;
}
div[data-testid="stDataFrame"] [data-baseweb="popover"] li,
div[data-testid="stDataFrame"] ul[role="listbox"] li {
    font-size: 11px !important;
    min-height: 26px !important;
    padding: 2px 8px !important;
}

/* 신청 근무 확정 박스 — 본문 글자 검정 (테마 간섭 방지) */
.req-save-panel, .req-save-panel h4, .req-save-panel p,
.req-save-status { color: #111111 !important; }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════════
#  Session State 초기화
# ════════════════════════════════════════════════════════════════════════════════
def _default_nurses(n: int = 9) -> list[str]:
    return ["수간호사"] + [f"간호사{i}" for i in range(1, n + 1)]


# 부서·간호사 명단 영속 저장 (앱 폴더의 JSON — 다시 실행·새로고침 후에도 복원)
_DEPT_SAVE_PATH = Path(__file__).resolve().parent / "user_departments.json"


def _departments_payload_ok(dep: dict) -> bool:
    if not dep or not isinstance(dep, dict):
        return False
    for name, nurses in dep.items():
        if not isinstance(name, str) or not str(name).strip():
            return False
        if not isinstance(nurses, list) or len(nurses) < 1:
            return False
    return True


def _load_departments_from_disk() -> dict | None:
    if not _DEPT_SAVE_PATH.is_file():
        return None
    try:
        with open(_DEPT_SAVE_PATH, encoding="utf-8") as f:
            data = json.load(f)
        dep = data.get("departments")
        if not _departments_payload_ok(dep):
            return None
        out = {str(k): [str(x) for x in v] for k, v in dep.items()}
        raw_fp = data.get("forbidden_pairs")
        fp_out = {}
        if isinstance(raw_fp, dict):
            for dk, rows in raw_fp.items():
                if not isinstance(rows, list):
                    continue
                clean = []
                for row in rows:
                    if isinstance(row, (list, tuple)) and len(row) == 2:
                        a, b = str(row[0]).strip(), str(row[1]).strip()
                        if a and b and a != b:
                            clean.append(sorted([a, b]))
                fp_out[str(dk)] = clean
        return {"departments": out, "active_dept": data.get("active_dept"), "forbidden_pairs": fp_out}
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return None


def _save_departments_to_disk() -> None:
    if "departments" not in st.session_state:
        return
    try:
        payload = {
            "departments": dict(st.session_state.departments),
            "active_dept": st.session_state.get("active_dept", ""),
            "forbidden_pairs": dict(st.session_state.get("dept_forbidden_pairs", {})),
        }
        with open(_DEPT_SAVE_PATH, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def _fp_pairs_to_indices(nurse_names: list[str], pairs: list) -> list[tuple[int, int]]:
    """이름 쌍 → 간호사 인덱스 쌍 (수간호사 0번 제외)."""
    idx = {name: i for i, name in enumerate(nurse_names)}
    out: list[tuple[int, int]] = []
    seen: set[tuple[int, int]] = set()
    for row in pairs or []:
        if not isinstance(row, (list, tuple)) or len(row) < 2:
            continue
        a, b = str(row[0]).strip(), str(row[1]).strip()
        if a not in idx or b not in idx or a == b:
            continue
        i, j = idx[a], idx[b]
        if i == 0 or j == 0:
            continue
        key = (min(i, j), max(i, j))
        if key not in seen:
            seen.add(key)
            out.append(key)
    return out


def _init_state():
    if "departments" not in st.session_state:
        loaded = _load_departments_from_disk()
        if loaded:
            st.session_state.departments = loaded["departments"]
            ad = loaded.get("active_dept") or ""
            keys = list(st.session_state.departments.keys())
            st.session_state.active_dept = ad if ad in st.session_state.departments else keys[0]
        else:
            st.session_state.departments = {"응급실": _default_nurses(9)}
        loaded_fp = loaded.get("forbidden_pairs") if loaded else None
        if isinstance(loaded_fp, dict):
            st.session_state.dept_forbidden_pairs = {
                str(k): v for k, v in loaded_fp.items() if isinstance(v, list)
            }
        else:
            st.session_state.dept_forbidden_pairs = {}
    if "dept_forbidden_pairs" not in st.session_state:
        _ld = _load_departments_from_disk()
        if _ld and isinstance(_ld.get("forbidden_pairs"), dict):
            st.session_state.dept_forbidden_pairs = {
                str(k): v for k, v in _ld["forbidden_pairs"].items() if isinstance(v, list)
            }
        else:
            st.session_state.dept_forbidden_pairs = {}
    if "active_dept" not in st.session_state:
        st.session_state.active_dept = list(st.session_state.departments.keys())[0]
    # 연도·월
    if "sel_year" not in st.session_state:
        st.session_state.sel_year = 2026
    if "sel_month" not in st.session_state:
        st.session_state.sel_month = 4
    # 부서별 데이터 (dict of dict)
    for key in ("dept_schedules", "dept_requests", "dept_holidays", "nurse_gen", "edit_mode"):
        if key not in st.session_state:
            st.session_state[key] = {}
    # 규칙 위반 팝업 제어
    if "show_violations" not in st.session_state:
        st.session_state.show_violations = False
    if "violations" not in st.session_state:
        st.session_state.violations = []

_init_state()

# ── 연도·월 전역 상수 동기화 (렌더링마다 app 모듈 갱신)
_app.set_period(st.session_state.sel_year, st.session_state.sel_month)


# ════════════════════════════════════════════════════════════════════════════════
#  규칙 위반 팝업 다이얼로그
# ════════════════════════════════════════════════════════════════════════════════
@st.dialog("⚠️ 규칙 검증 결과", width="small")
def _show_violations_dialog():
    issues = st.session_state.violations
    errors = [v for v in issues if v["level"] == "error"]
    warns  = [v for v in issues if v["level"] == "warn"]

    if not issues:
        st.success("✅ 모든 규칙을 만족합니다!")
        if st.button("닫기", type="primary", use_container_width=True):
            st.session_state.show_violations = False
            st.rerun()
        return

    st.caption(f"🔴 오류 {len(errors)}건 &nbsp;|&nbsp; 🟡 경고 {len(warns)}건")
    st.markdown("---")

    if errors:
        st.markdown("**🔴 오류**")
        for v in errors:
            st.error(v["msg"], icon="🚨")

    if warns:
        st.markdown("**🟡 경고**")
        for v in warns:
            st.warning(v["msg"], icon="⚠️")

    st.markdown("---")
    if st.button("닫기", type="primary", use_container_width=True):
        st.session_state.show_violations = False
        st.rerun()


# 팝업 자동 열기
if st.session_state.show_violations:
    _show_violations_dialog()

# ── 안전하게 active_dept 보정 (부서 삭제 후 남은 부서로 자동 전환)
if st.session_state.active_dept not in st.session_state.departments:
    st.session_state.active_dept = list(st.session_state.departments.keys())[0]

# ════════════════════════════════════════════════════════════════════════════════
#  헬퍼 함수
# ════════════════════════════════════════════════════════════════════════════════
def _parse_holidays(text: str) -> list[int]:
    result = []
    for tok in text.replace("，", ",").split(","):
        tok = tok.strip()
        if tok.isdigit():
            d = int(tok)
            if 1 <= d <= _app.NUM_DAYS:
                result.append(d)
    return sorted(set(result))


def _day_label(day: dict) -> str:
    mark = "🔴" if day["is_holiday"] else ("🔵" if day["is_weekend"] else "")
    return f"{day['day']}({day['weekday_name']}){mark}"


def _make_requests_df(nurses: list[str], days: list) -> pd.DataFrame:
    num_days = _app.NUM_DAYS
    return pd.DataFrame(
        [[""] * num_days for _ in range(len(nurses))],
        index=nurses,
        columns=[_day_label(d) for d in days],
    )


def _df_to_requests(df: pd.DataFrame, days: list) -> dict:
    result = {}
    for i in range(len(df)):
        for j, day in enumerate(days):
            val = str(df.iloc[i, j]).strip()
            if val and val not in ("", "None", "nan"):
                result.setdefault(i, {})[day["day"]] = val
    return result


def _render_schedule_html(schedule: dict, nurse_names: list, days: list,
                          requests: dict | None = None) -> str:
    num = len(nurse_names)
    requests = requests or {}
    th = lambda txt, bg, extra="", fg="#37474F": (
        f'<th style="background:{bg};color:{fg};padding:4px 2px;'
        f'text-align:center;white-space:nowrap;{extra}">{txt}</th>'
    )
    rows = ["<tr>"]
    rows.append(th("간호사", "#ECEFF1",
                   "min-width:80px;padding:5px 8px;position:sticky;left:0;z-index:5;", "#263238"))
    for day in days:
        if day["is_holiday"]:
            dbg, dfg = "#FFEBEE", "#C62828"
        elif day["is_weekend"]:
            dbg, dfg = "#E3F2FD", "#1565C0"
        else:
            dbg, dfg = "#F5F5F5", "#455A64"
        rows.append(th(
            f"{day['day']}<br><span style='font-size:9px'>{day['weekday_name']}</span>",
            dbg, "min-width:30px;", dfg,
        ))
    for lbl in ["N", "D", "E", "OF", "NO"]:
        bg = SHIFT_COLORS.get(lbl, "#ECEFF1")
        fg = SHIFT_TEXT_COLORS.get(lbl, "#37474F")
        rows.append(th(
            f"{lbl}<br><span style='font-size:9px'>합계</span>",
            bg, "min-width:36px;", fg,
        ))
    rows.append("</tr>")

    for n_idx, name in enumerate(nurse_names):
        ns = schedule.get(n_idx, {})
        counts = {"N": 0, "D": 0, "E": 0, "OF": 0, "NO": 0}
        row_bg = "#FAFAFA" if n_idx % 2 == 0 else "#F5F5F5"
        cells = [
            f'<td style="background:#ECEFF1;color:#263238;font-weight:700;'
            f'padding:4px 8px;white-space:nowrap;position:sticky;left:0;'
            f'border-right:2px solid #CFD8DC;">{name}</td>'
        ]
        nurse_req = requests.get(n_idx, {})
        for day in days:
            d_num = day["day"]
            shift = ns.get(d_num, "")
            bg  = SHIFT_COLORS.get(shift, "#ECEFF1")
            fg  = SHIFT_TEXT_COLORS.get(shift, "#9E9E9E")
            # 신청 근무이면 밑줄 표시
            is_requested = nurse_req.get(d_num) == shift and shift != ""
            underline = "text-decoration:underline;text-underline-offset:3px;" if is_requested else ""
            cells.append(
                f'<td style="background:{bg};color:{fg};font-weight:700;{underline}'
                f'padding:3px 1px;text-align:center;border:1px solid #E0E0E0;">{shift}</td>'
            )
            if shift == "N":          counts["N"] += 1
            elif shift == "D":        counts["D"] += 1
            elif shift == "E":        counts["E"] += 1
            elif shift in ("OF", "OH"):
                counts["OF"] += 1
            elif shift == "NO":
                counts["NO"] += 1

        for key in ["N", "D", "E", "OF", "NO"]:
            bg = SHIFT_COLORS.get(key, "#ECEFF1")
            fg = SHIFT_TEXT_COLORS.get(key, "#37474F")
            cells.append(
                f'<td style="background:{bg};color:{fg};font-weight:700;'
                f'text-align:center;padding:3px;">{counts[key]}</td>'
            )
        rows.append(f'<tr style="background:{row_bg};">' + "".join(cells) + "</tr>")

    for lbl, sk in [("D인원", "D"), ("E인원", "E"), ("N인원", "N")]:
        hbg = SHIFT_COLORS.get(sk, "#ECEFF1")
        hfg = SHIFT_TEXT_COLORS.get(sk, "#37474F")
        bg = SHIFT_COLORS.get(sk, "#FAFAFA")
        cells = [
            f'<td style="background:{hbg};color:{hfg};font-weight:700;'
            f'padding:4px 8px;white-space:nowrap;position:sticky;left:0;'
            f'border-right:2px solid #CFD8DC;">{lbl}</td>'
        ]
        for day in days:
            cnt = sum(1 for n in range(num) if schedule.get(n, {}).get(day["day"]) == sk)
            cells.append(
                f'<td style="background:{bg};color:{hfg};font-weight:700;text-align:center;'
                f'padding:3px;border:1px solid #E0E0E0;">{cnt}</td>'
            )
        cells += ["<td></td>"] * 5
        rows.append("<tr>" + "".join(cells) + "</tr>")

    return (
        '<div style="overflow-x:auto;border-radius:10px;'
        'box-shadow:0 2px 12px rgba(0,0,0,0.09);">'
        '<table style="border-collapse:collapse;font-size:12px;width:100%;">'
        "<thead>" + rows[0] + "</thead>"
        "<tbody>" + "".join(rows[1:]) + "</tbody>"
        "</table></div>"
    )


def _schedule_to_edit_df(schedule: dict, nurse_names: list, days: list) -> pd.DataFrame:
    """schedule dict → data_editor용 DataFrame (행=간호사, 열=날짜)"""
    rows = []
    for n_idx, name in enumerate(nurse_names):
        row = {}
        for day in days:
            row[_day_label(day)] = schedule.get(n_idx, {}).get(day["day"], "")
        rows.append(row)
    return pd.DataFrame(rows, index=nurse_names)


def _edit_df_to_schedule(df: pd.DataFrame, days: list) -> dict:
    """data_editor DataFrame → schedule dict"""
    schedule = {}
    for n_idx in range(len(df)):
        schedule[n_idx] = {}
        for j, day in enumerate(days):
            val = str(df.iloc[n_idx, j]).strip()
            if val and val not in ("", "None", "nan"):
                schedule[n_idx][day["day"]] = val
    return schedule


def _generate_excel(schedule, num_nurses, nurse_names, days) -> bytes:
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = f"{_app.YEAR}년 {_app.MONTH}월 근무표"
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"),  bottom=Side(style="thin"))
    def _xrgb(h: str) -> str:
        return h.replace("#", "").upper()

    BG = {
        sk: (_xrgb(SHIFT_COLORS[sk]), _xrgb(SHIFT_TEXT_COLORS[sk]))
        for sk in SHIFT_COLORS
    }
    num_days = _app.NUM_DAYS
    NC, OC, DC = num_days + 2, num_days + 3, num_days + 4

    year_label = _app.YEAR
    month_label = _app.MONTH
    ws.merge_cells(f"A1:{get_column_letter(DC)}1")
    c = ws["A1"]; c.value = f"{year_label}년 {month_label}월 근무표"
    c.fill = PatternFill("solid", fgColor=_xrgb(SHIFT_COLORS["N"])); c.alignment = ctr
    c.font = Font(bold=True, size=14, color=_xrgb(SHIFT_TEXT_COLORS["N"]))
    ws.row_dimensions[1].height = 28

    h = ws.cell(2, 1, "간호사")
    h.fill = PatternFill("solid", fgColor=_xrgb(SHIFT_COLORS["OF"])); h.font = Font(
        bold=True, color=_xrgb(SHIFT_TEXT_COLORS["NO"]), size=10,
    )
    h.alignment = ctr; h.border = thin; ws.column_dimensions["A"].width = 11

    for d, day in enumerate(days):
        col = d + 2
        cell = ws.cell(2, col, f"{day['day']}\n{day['weekday_name']}")
        cell.alignment = ctr; cell.border = thin
        if day["is_holiday"]:
            bg, tfg = "FFEBEE", "C62828"
        elif day["is_weekend"]:
            bg, tfg = "E3F2FD", "1565C0"
        else:
            bg, tfg = "F5F5F5", "455A64"
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(bold=True, color=tfg, size=9)
        ws.column_dimensions[get_column_letter(col)].width = 4.5

    for col, lbl, sk in [(NC, "N\n합계", "N"), (OC, "OF\n합계", "OF"), (DC, "D\n합계", "D")]:
        c = ws.cell(2, col, lbl); c.alignment = ctr; c.border = thin
        c.fill = PatternFill("solid", fgColor=BG[sk][0])
        c.font = Font(bold=True, color=BG[sk][1], size=9)
        ws.column_dimensions[get_column_letter(col)].width = 5.5
    ws.row_dimensions[2].height = 28

    for n_idx, name in enumerate(nurse_names):
        row = n_idx + 3
        nc = ws.cell(row, 1, name)
        nc.fill = PatternFill("solid", fgColor=_xrgb(SHIFT_COLORS["OF"]))
        nc.font = Font(bold=True, color=_xrgb(SHIFT_TEXT_COLORS["NO"]), size=9)
        nc.alignment = ctr; nc.border = thin; ws.row_dimensions[row].height = 18
        ns = schedule.get(n_idx, {}); n_c = d_c = of_c = 0
        for d, day in enumerate(days):
            shift = ns.get(d + 1, ""); col = d + 2
            cell = ws.cell(row, col, shift); cell.alignment = ctr; cell.border = thin
            if shift in BG:
                bg, fg = BG[shift]
                cell.fill = PatternFill("solid", fgColor=bg); cell.font = Font(color=fg, size=9, bold=True)
            if shift == "N": n_c += 1
            elif shift == "D": d_c += 1
            elif shift in ("OF","OH","NO"): of_c += 1
        for col, val, sk in [(NC, n_c, "N"), (OC, of_c, "OF"), (DC, d_c, "D")]:
            bg, fg = BG[sk]
            c = ws.cell(row, col, val); c.alignment = ctr; c.border = thin
            c.fill = PatternFill("solid", fgColor=bg); c.font = Font(color=fg, bold=True, size=10)

    sr = len(nurse_names) + 3
    for idx, (lbl, sk) in enumerate([("D 인원", "D"), ("E 인원", "E"), ("N 인원", "N")]):
        row = sr + idx; lc = ws.cell(row, 1, lbl)
        lb, lf = BG[sk]
        lc.fill = PatternFill("solid", fgColor=lb); lc.font = Font(bold=True, color=lf, size=9)
        lc.alignment = ctr; lc.border = thin; ws.row_dimensions[row].height = 16
        for d in range(num_days):
            cnt = sum(1 for n in range(num_nurses) if schedule.get(n, {}).get(d + 1) == sk)
            cell = ws.cell(row, d + 2, cnt); cell.alignment = ctr; cell.border = thin
            cell.fill = PatternFill("solid", fgColor=lb)
            cell.font = Font(bold=True, color=lf, size=9)

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ════════════════════════════════════════════════════════════════════════════════
generate_btn = False
with st.sidebar:

    # ── 로고 ─────────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="text-align:center;padding:10px 0 10px;">
        <div style="font-size:34px;">🏥</div>
        <div style="font-size:17px;font-weight:800;color:#1A237E;letter-spacing:.4px;">근무표 생성기</div>
    </div>""", unsafe_allow_html=True)

    # ── 연도·월 선택 ─────────────────────────────────────────────────────────
    col_y, col_m = st.columns(2)
    with col_y:
        sel_year = st.selectbox(
            "연도",
            list(range(2024, 2032)),
            index=list(range(2024, 2032)).index(st.session_state.sel_year),
            key="year_selectbox",
        )
    with col_m:
        month_names = ["1월","2월","3월","4월","5월","6월",
                       "7월","8월","9월","10월","11월","12월"]
        sel_month = st.selectbox(
            "월",
            list(range(1, 13)),
            index=st.session_state.sel_month - 1,
            format_func=lambda m: month_names[m - 1],
            key="month_selectbox",
        )

    # 연·월이 바뀌면 모든 부서 데이터 초기화
    if sel_year != st.session_state.sel_year or sel_month != st.session_state.sel_month:
        st.session_state.sel_year  = sel_year
        st.session_state.sel_month = sel_month
        st.session_state.dept_schedules = {}
        st.session_state.dept_requests  = {}
        _app.set_period(sel_year, sel_month)
        st.rerun()

    # 현재 선택 배지
    st.markdown(
        f'<div style="text-align:center;font-size:15px;font-weight:800;'
        f'color:#000000;margin-bottom:6px;letter-spacing:0.02em;'
        f'-webkit-font-smoothing:antialiased;">'
        f'📅 {sel_year}년 {month_names[sel_month-1]} '
        f'({_calendar.monthrange(sel_year, sel_month)[1]}일)</div>',
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════════
    #  ① 부서 관리
    # ════════════════════════════════════════════════════════════════════════
    st.markdown("#### 🏢 부서 관리")

    dept_list = list(st.session_state.departments.keys())

    # 부서 선택 selectbox
    try:
        active_idx = dept_list.index(st.session_state.active_dept)
    except ValueError:
        active_idx = 0

    active_dept = st.selectbox(
        "현재 부서",
        dept_list,
        index=active_idx,
        key="dept_selectbox",
        label_visibility="collapsed",
    )
    st.session_state.active_dept = active_dept

    # 선택된 부서 배지
    st.markdown(
        f'<div class="dept-badge">📂 {active_dept}</div>'
        f'<span style="font-size:11px;color:#546E7A;">'
        f'{len(st.session_state.departments[active_dept])}명</span>',
        unsafe_allow_html=True,
    )
    st.markdown("")

    # ── 새 부서 추가
    with st.expander("➕  새 부서 추가"):
        new_dept_input = st.text_input(
            "부서 이름", key="new_dept_input",
            placeholder="예: 본관 5병동",
            label_visibility="collapsed",
        )
        if st.button("부서 추가", key="btn_add_dept", use_container_width=True):
            name = new_dept_input.strip()
            if not name:
                st.warning("부서 이름을 입력하세요.")
            elif name in st.session_state.departments:
                st.error("이미 존재하는 부서입니다.")
            else:
                st.session_state.departments[name] = _default_nurses(9)
                st.session_state.active_dept = name
                if "new_dept_input" in st.session_state:
                    st.session_state.new_dept_input = ""
                _save_departments_to_disk()
                st.rerun()

    # ── 현재 부서 삭제
    if len(dept_list) > 1:
        if st.button(f"🗑️  '{active_dept}' 부서 삭제", key="btn_del_dept",
                     help="현재 부서와 모든 데이터를 삭제합니다"):
            for store in ("dept_schedules", "dept_requests", "dept_holidays", "nurse_gen"):
                st.session_state[store].pop(active_dept, None)
            st.session_state.dept_forbidden_pairs.pop(active_dept, None)
            del st.session_state.departments[active_dept]
            st.session_state.active_dept = list(st.session_state.departments.keys())[0]
            _save_departments_to_disk()
            st.rerun()
    else:
        st.caption("(부서가 1개일 때는 삭제 불가)")

    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════════
    #  ② 간호사 명단 (선택된 부서) – 드롭다운 형태
    # ════════════════════════════════════════════════════════════════════════
    nurses = st.session_state.departments[active_dept]
    gen    = st.session_state.nurse_gen.get(active_dept, 0)

    with st.expander(f"👩‍⚕️ 간호사 명단  ({len(nurses)}명)", expanded=False):
        updated_nurses: list[str] = []
        nurse_to_delete: int | None = None

        for i, name in enumerate(nurses):
            col_name, col_del = st.columns([5, 1])
            with col_name:
                icon = "👑" if i == 0 else "👤"
                new_name = st.text_input(
                    icon,
                    value=name,
                    key=f"nname_{active_dept}_{i}_g{gen}",
                    label_visibility="collapsed",
                    placeholder=f"{'수간호사 이름' if i == 0 else f'간호사{i} 이름'}",
                )
                updated_nurses.append(new_name if new_name.strip() else name)
            with col_del:
                st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
                if i == 0:
                    st.markdown("👑", help="수간호사는 삭제할 수 없습니다")
                else:
                    if st.button("✕", key=f"del_nurse_{active_dept}_{i}_g{gen}",
                                 help=f"'{name}' 삭제"):
                        nurse_to_delete = i

        # 삭제 처리
        if nurse_to_delete is not None:
            updated_nurses.pop(nurse_to_delete)
            st.session_state.departments[active_dept] = updated_nurses
            _fp = st.session_state.dept_forbidden_pairs.get(active_dept, [])
            st.session_state.dept_forbidden_pairs[active_dept] = [
                p for p in _fp
                if isinstance(p, (list, tuple)) and len(p) >= 2
                and str(p[0]).strip() in updated_nurses and str(p[1]).strip() in updated_nurses
            ]
            st.session_state.dept_requests[active_dept]  = None
            st.session_state.dept_schedules[active_dept] = None
            st.session_state.nurse_gen[active_dept]      = gen + 1
            st.rerun()

        # 이름 변경 동기화
        st.session_state.departments[active_dept] = updated_nurses
        df_existing = st.session_state.dept_requests.get(active_dept)
        if df_existing is not None and len(df_existing) == len(updated_nurses):
            df_existing.index = updated_nurses

        # 간호사 추가 버튼
        st.markdown("<div style='margin-top:4px'></div>", unsafe_allow_html=True)
        if st.button("➕  간호사 추가", key="btn_add_nurse", use_container_width=True):
            new_idx = len(st.session_state.departments[active_dept])
            st.session_state.departments[active_dept].append(f"간호사{new_idx}")
            st.session_state.dept_requests[active_dept]  = None
            st.session_state.dept_schedules[active_dept] = None
            st.session_state.nurse_gen[active_dept]      = gen + 1
            st.rerun()

    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════════
    #  ③ 공휴일
    # ════════════════════════════════════════════════════════════════════════
    st.markdown("#### 📅 공휴일 설정")
    default_hols = st.session_state.dept_holidays.get(active_dept, "")
    holidays_raw = st.text_input(
        "공휴일 날짜 (쉼표로 구분)",
        value=default_hols,
        key=f"holidays_{active_dept}",
        placeholder="예: 5, 15, 25",
        label_visibility="collapsed",
    )
    st.session_state.dept_holidays[active_dept] = holidays_raw
    holidays = _parse_holidays(holidays_raw)

    if holidays:
        badge = " · ".join(f"{h}일" for h in holidays)
        st.markdown(
            f'<div style="background:#E3F2FD;border:1px solid #90CAF9;border-radius:6px;'
            f'padding:5px 10px;font-size:11px;color:#1565C0;margin-top:2px;">📌 {badge}</div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════════
    #  ③b 함께 근무 불가 (같은 날 D/E/N 동시 배치 금지)
    # ════════════════════════════════════════════════════════════════════════
    st.markdown(
        '<p style="font-size:11px;font-weight:600;margin:0 0 2px 0;color:#212121;">'
        "🙅 함께 근무 불가</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p style="font-size:10px;line-height:1.35;color:#616161;margin:0 0 6px 0;">'
        "미숙련 간호사 등 <strong>같은 날·같은 근무(D / E / N)</strong>에 함께 서면 안 되는 두 명을 등록합니다. "
        "(수간호사는 여기서 선택하지 않습니다)</p>",
        unsafe_allow_html=True,
    )
    _fp_list = st.session_state.dept_forbidden_pairs.setdefault(active_dept, [])
    _staff = [n for n in nurses if n != nurses[0]]
    # 가로 2열이면 너비 부족으로 '간…' 말줄임됨 → 세로 풀너비로 전체 이름 표시
    _sel_a = st.selectbox(
        "간호사 A",
        _staff if _staff else [""],
        key=f"fp_a_{active_dept}",
        label_visibility="collapsed",
    )
    _b_opts = [x for x in _staff if x != _sel_a] or _staff
    _sel_b = st.selectbox(
        "간호사 B",
        _b_opts,
        key=f"fp_b_{active_dept}",
        label_visibility="collapsed",
    )
    if st.button("➕ 추가", key=f"fp_add_{active_dept}", use_container_width=True):
        if _sel_a and _sel_b and _sel_a != _sel_b:
            _pair = sorted([_sel_a, _sel_b])
            _have = {tuple(sorted([str(x[0]), str(x[1])])) for x in _fp_list if len(x) >= 2}
            if tuple(_pair) not in _have:
                _fp_list.append(_pair)
                _save_departments_to_disk()
                st.rerun()
    if _fp_list:
        for _i, _pr in enumerate(list(_fp_list)):
            _r1, _r2 = st.columns([5, 1])
            with _r1:
                st.markdown(
                    f'<div style="font-size:10px;color:#37474F;padding:1px 0;line-height:1.3;">'
                    f"🔗 {_pr[0]} · {_pr[1]}</div>",
                    unsafe_allow_html=True,
                )
            with _r2:
                if st.button("삭제", key=f"fp_rm_{active_dept}_{_i}", use_container_width=True):
                    _fp_list.pop(_i)
                    _save_departments_to_disk()
                    st.rerun()
    else:
        st.markdown(
            '<p style="font-size:10px;color:#9E9E9E;margin:0;">등록된 쌍이 없습니다.</p>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════════
    #  ④ 근무표 생성 버튼
    # ════════════════════════════════════════════════════════════════════════
    generate_btn = st.button(
        "🗓️  근무표 생성", type="primary", use_container_width=True,
        key="btn_generate",
    )

    if st.session_state.dept_schedules.get(active_dept):
        st.markdown(
            '<div style="text-align:center;font-size:11px;'
            'color:#546E7A;margin-top:8px;">'
            '✅ 근무표가 생성되어 있습니다</div>',
            unsafe_allow_html=True,
        )


# ════════════════════════════════════════════════════════════════════════════════
#  MAIN – 변수 준비
# ════════════════════════════════════════════════════════════════════════════════
nurses      = st.session_state.departments[active_dept]   # 최신 명단
num_nurses  = len(nurses)
days        = get_april_days(holidays)
col_labels  = [_day_label(d) for d in days]
gen         = st.session_state.nurse_gen.get(active_dept, 0)
editor_key  = f"req_editor_{active_dept}_n{num_nurses}_g{gen}"

# requests_df 준비 (없거나 행 수 불일치 시 새로 생성)
df_req = st.session_state.dept_requests.get(active_dept)
if df_req is None or df_req.shape[0] != num_nurses:
    df_req = _make_requests_df(nurses, days)
    st.session_state.dept_requests[active_dept] = df_req
else:
    df_req.index   = nurses
    df_req.columns = col_labels

# None / nan → 공백으로 정규화
df_req = df_req.apply(lambda col: col.map(
    lambda x: "" if (x is None or str(x).strip() in ("None", "nan")) else str(x).strip()
))
st.session_state.dept_requests[active_dept] = df_req

# ════════════════════════════════════════════════════════════════════════════════
#  MAIN – 신청 근무 입력 달력
# ════════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="card">
  <div class="card-title">📝 신청 근무 입력 &nbsp;
    <span class="dept-badge">{active_dept}</span>
  </div>
  <div class="card-sub">
    {_app.YEAR}년 {_app.MONTH}월 &nbsp;·&nbsp;
    셀을 클릭해 원하는 근무를 선택하세요. 빈칸은 자동 배정됩니다.
    &nbsp;🔵 토·일 &nbsp;|&nbsp; 🔴 공휴일
  </div>
</div>
""", unsafe_allow_html=True)

# 범례 (작은 칩 형태)
legend_items = [
    ("A1","수간호사"), ("D","데이"), ("E","이브닝"), ("N","나이트"),
    ("OF","휴무"), ("OH","휴일"), ("NO","N 20회 휴무(수기)"),
    ("연","연차"), ("병","병가"), ("공","공가"), ("경","경조"), ("EDU","교육"),
]
_leg_chips = []
for shift, tip in legend_items:
    bg = SHIFT_COLORS.get(shift, "#ECEFF1")
    fg = SHIFT_TEXT_COLORS.get(shift, "#000")
    _leg_chips.append(
        f'<span title="{tip}" style="display:inline-block;background:{bg};color:{fg};'
        f'text-align:center;padding:1px 5px;margin:0 2px 3px 0;border-radius:3px;'
        f'font-size:9px;font-weight:700;line-height:1.35;">{shift}</span>'
    )
st.markdown(
    f'<div style="display:flex;flex-wrap:wrap;align-items:center;gap:0;margin:0 0 6px 0;">'
    f'{"".join(_leg_chips)}</div>',
    unsafe_allow_html=True,
)

# data_editor
shift_options = [""] + SHIFT_NAMES
col_config = {
    lbl: st.column_config.SelectboxColumn(
        lbl, options=shift_options, width="small", required=False,
    )
    for lbl in col_labels
}
edited_df = st.data_editor(
    df_req,
    column_config=col_config,
    use_container_width=True,
    height=min(40 * num_nurses + 90, 620),
    key=editor_key,
    num_rows="fixed",
)

# 저장 영역 (전체 너비 — 좁은 열에 넣으면 버튼이 안 보이는 경우가 있음)
req_saved_key = f"req_saved_{active_dept}_g{gen}"

def _clean_req_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(
        lambda col: col.map(
            lambda x: ""
            if (x is None or str(x).strip() in ("None", "nan"))
            else str(x).strip()
        )
    )

with st.container(border=True):
    # Streamlit 알림/캡션은 테마에 따라 흰색으로 보일 수 있어 명시적으로 검정 처리
    st.markdown(
        '<div class="req-save-panel">'
        '<h4 style="margin:0 0 8px 0;font-size:1.1rem;color:#111111;font-weight:700;">💾 신청 근무 확정</h4>'
        '<p style="margin:0 0 12px 0;font-size:13px;color:#222222;line-height:1.5;">'
        '표에서 근무를 고른 다음 <strong>저장하기</strong>를 눌러야 반영됩니다. '
        '(스크롤을 내려 이 영역이 보이는지 확인하세요.)</p></div>',
        unsafe_allow_html=True,
    )

    if st.button(
        "💾 저장하기",
        type="primary",
        use_container_width=True,
        key=f"btn_save_requests_{active_dept}_g{gen}",
    ):
        cleaned = _clean_req_df(edited_df)
        st.session_state.dept_requests[active_dept] = cleaned
        st.session_state[req_saved_key] = True
        st.rerun()

    if st.session_state.get(req_saved_key):
        st.markdown(
            '<div class="req-save-status req-save-ok" style="background:#E8F5E9;border:1px solid #A5D6A7;'
            'border-radius:8px;padding:10px 14px;color:#111111;font-size:14px;margin:8px 0;line-height:1.45;">'
            "✅ 신청 근무가 저장되었습니다.</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="req-save-status req-save-warn" style="background:#FFF8E1;border:1px solid #FFE082;'
            'border-radius:8px;padding:10px 14px;color:#111111;font-size:14px;margin:8px 0;line-height:1.45;">'
            "⚠️ 아직 저장되지 않았습니다. 근무표 생성 전에 <strong>저장하기</strong>를 눌러 주세요.</div>",
            unsafe_allow_html=True,
        )

    c_clear, _ = st.columns([1, 3])
    with c_clear:
        if st.button(
            "🗑️ 신청 전체 지우기",
            use_container_width=True,
            key=f"btn_clear_requests_{active_dept}_g{gen}",
        ):
            st.session_state.dept_requests[active_dept] = _make_requests_df(nurses, days)
            st.session_state[req_saved_key] = False
            st.rerun()

# ════════════════════════════════════════════════════════════════════════════════
#  근무표 생성 처리
# ════════════════════════════════════════════════════════════════════════════════
if generate_btn:
    # 저장된 신청 근무 사용 (버튼으로 확정된 데이터)
    saved_df = st.session_state.dept_requests.get(active_dept, edited_df)
    requests = _df_to_requests(saved_df, days)
    _fp_idx = _fp_pairs_to_indices(
        nurses,
        st.session_state.dept_forbidden_pairs.get(active_dept, []),
    )
    with st.spinner("⏳ 근무표를 계산하는 중입니다…"):
        schedule, success, status = solve_schedule(
            num_nurses, requests, holidays, forbidden_pairs=_fp_idx or None,
        )
    if success:
        st.session_state.dept_schedules[active_dept] = {
            "schedule":    schedule,
            "nurse_names": nurses.copy(),
            "holidays":    holidays,
            "requests":    requests,
        }
        # 규칙 검증
        issues = validate_schedule(
            schedule, num_nurses, holidays,
            forbidden_pairs=_fp_idx or None,
            nurse_names=nurses,
        )
        st.session_state.violations     = issues
        st.session_state.show_violations = True   # 팝업 자동 열기
        if not issues:
            st.toast("✅ 근무표 생성 완료! 모든 규칙 통과", icon="🎉")
        else:
            errors = sum(1 for v in issues if v["level"] == "error")
            warns  = sum(1 for v in issues if v["level"] == "warn")
            st.toast(f"⚠️ 규칙 위반 {errors}건 오류 / {warns}건 경고 발견", icon="⚠️")
    else:
        st.error(
            f"❌ 근무표 생성 실패: {status}\n\n"
            "신청 근무를 줄이거나 간호사 수를 조정 후 다시 시도해 주세요."
        )

# ════════════════════════════════════════════════════════════════════════════════
#  MAIN – 생성된 근무표
# ════════════════════════════════════════════════════════════════════════════════
sched_data = st.session_state.dept_schedules.get(active_dept)

if sched_data:
    schedule    = sched_data["schedule"]
    sched_names = sched_data["nurse_names"]
    sched_hols  = sched_data["holidays"]
    sched_reqs  = sched_data.get("requests", {})
    sched_days  = get_april_days(sched_hols)
    sched_n     = len(sched_names)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── 수정 모드 상태
    is_edit = st.session_state.edit_mode.get(active_dept, False)

    # ── 헤더 버튼 행 ───────────────────────────────────────────────────────────
    col_t, col_edit, col_vld, col_dl = st.columns([3, 1, 1, 1])
    with col_t:
        edit_badge = ' <span style="color:#E65100;font-size:12px;">✏️ 수정 중</span>' if is_edit else ""
        st.markdown(f"""
        <div class="card" style="padding:12px 20px;margin-bottom:8px;">
          <div class="card-title">📋 생성된 근무표 &nbsp;
            <span class="dept-badge">{active_dept}</span>{edit_badge}
          </div>
          <div class="card-sub">{_app.YEAR}년 {_app.MONTH}월 · {sched_n}명</div>
        </div>
        """, unsafe_allow_html=True)

    with col_edit:
        st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)
        if not is_edit:
            if st.button("✏️ 수정", use_container_width=True, key="btn_edit_on"):
                st.session_state.edit_mode[active_dept] = True
                st.rerun()
        else:
            if st.button("❌ 취소", use_container_width=True, key="btn_edit_off"):
                st.session_state.edit_mode[active_dept] = False
                st.rerun()

    with col_vld:
        st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)
        vld_issues = st.session_state.get("violations", [])
        err_cnt  = sum(1 for v in vld_issues if v["level"] == "error")
        warn_cnt = sum(1 for v in vld_issues if v["level"] == "warn")
        btn_label = (
            "✅ 규칙 통과" if not vld_issues
            else f"⚠️ {err_cnt}오류/{warn_cnt}경고"
        )
        if st.button(btn_label, use_container_width=True, key="btn_violations"):
            st.session_state.show_violations = True
            st.rerun()

    with col_dl:
        st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)
        excel_bytes = _generate_excel(schedule, sched_n, sched_names, sched_days)
        st.download_button(
            "📥 엑셀 다운로드", data=excel_bytes,
            file_name=f"{_app.YEAR}년_{_app.MONTH}월_근무표_{active_dept}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ── 수정 모드: data_editor ─────────────────────────────────────────────────
    if is_edit:
        st.info("셀을 클릭하면 근무를 변경할 수 있습니다. 수정 후 **💾 저장** 버튼을 눌러주세요.", icon="✏️")

        # 근무 범례 (수정 모드 안내)
        shift_options = [""] + SHIFT_NAMES
        edit_df = _schedule_to_edit_df(schedule, sched_names, sched_days)
        col_cfg = {
            lbl: st.column_config.SelectboxColumn(lbl, options=shift_options, width="small")
            for lbl in edit_df.columns
        }
        edited = st.data_editor(
            edit_df,
            column_config=col_cfg,
            use_container_width=True,
            height=min(42 * sched_n + 90, 700),
            key=f"sched_editor_{active_dept}",
            num_rows="fixed",
        )

        # 저장 버튼
        save_col, _ = st.columns([1, 3])
        with save_col:
            if st.button("💾 저장", type="primary", use_container_width=True, key="btn_save_edit"):
                new_schedule = _edit_df_to_schedule(edited, sched_days)
                st.session_state.dept_schedules[active_dept]["schedule"] = new_schedule
                # 재검증
                _fp_ed = _fp_pairs_to_indices(
                    sched_names,
                    st.session_state.dept_forbidden_pairs.get(active_dept, []),
                )
                issues = validate_schedule(
                    new_schedule, sched_n, sched_hols,
                    forbidden_pairs=_fp_ed or None,
                    nurse_names=sched_names,
                )
                st.session_state.violations     = issues
                st.session_state.show_violations = bool(issues)
                st.session_state.edit_mode[active_dept] = False
                if issues:
                    err_c = sum(1 for v in issues if v["level"] == "error")
                    war_c = sum(1 for v in issues if v["level"] == "warn")
                    st.toast(f"💾 저장 완료 — 위반 {err_c}오류/{war_c}경고 발견", icon="⚠️")
                else:
                    st.toast("💾 저장 완료! 모든 규칙 통과", icon="✅")
                st.rerun()

    # ── 뷰 모드: 컬러 HTML 근무표 ─────────────────────────────────────────────
    else:
        st.markdown(
            _render_schedule_html(schedule, sched_names, sched_days, sched_reqs),
            unsafe_allow_html=True,
        )

    # ── 통계 ──────────────────────────────────────────────────────────────────
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("""
    <div style="font-size:16px;font-weight:700;color:#ffffff;margin-bottom:8px;">
        📊 간호사별 근무 통계
    </div>""", unsafe_allow_html=True)

    nurse_stats, _ = build_stats(schedule, sched_n)
    rows = []
    for n_idx, name in enumerate(sched_names):
        s = nurse_stats.get(n_idx, {})
        rows.append({
            "간호사":       name,
            "🌙 N (야간)":  s.get("N", 0),
            "☀️ D (낮)":   s.get("D", 0),
            "🌆 E (저녁)":  s.get("E", 0),
            "🏖️ OF (휴무)": s.get("OF", 0),
            "📌 NO":       s.get("NO", 0),
            "⭐ A1":        s.get("A1", 0),
        })
    stat_df = pd.DataFrame(rows).set_index("간호사")

    def _hl(col):
        def _cell(sk):
            b = SHIFT_COLORS.get(sk, "#ECEFF1")
            t = SHIFT_TEXT_COLORS.get(sk, "#37474F")
            return f"background:{b};color:{t};font-weight:700;"
        colors = {
            "🌙 N (야간)": _cell("N"),
            "☀️ D (낮)":  _cell("D"),
            "🌆 E (저녁)": _cell("E"),
            "🏖️ OF (휴무)": _cell("OF"),
            "📌 NO":      _cell("NO"),
            "⭐ A1":       _cell("A1"),
        }
        return [colors.get(col.name, "")] * len(col)

    st.dataframe(
        stat_df.style.apply(_hl, axis=0),
        use_container_width=True,
        height=42 * sched_n + 60,
    )

# 테마·위젯 CSS보다 나중에 적용 — text_input 글자색 최종 고정(검정)
st.markdown(
    """
    <style>
    .stApp, section[data-testid="stSidebar"] { color-scheme: light !important; }
    section[data-testid="stSidebar"] > div:first-child { background: #ffffff !important; }
    div[data-testid="stTextInput"] input,
    .stTextInput input,
    section[data-testid="stSidebar"] [data-baseweb="input"] input {
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
        caret-color: #000000 !important;
        background-color: #ffffff !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div {
        background: #ffffff !important;
        box-shadow: none !important;
        border: 1.5px solid #616161 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"],
    section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] p {
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
        font-weight: 600 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 부서·간호사 명단을 JSON에 저장 (페이지 로드마다 최신 상태 유지)
_save_departments_to_disk()
