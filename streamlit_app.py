"""
응급실 근무표 생성기 – Streamlit UI v2
- 연도·월 선택 가능
- 부서(Department) CRUD
- 간호사(Staff) CRUD: 추가 / 이름 수정 / 삭제
- 부서별 신청 근무 입력 달력 (data_editor)
- 함께 근무 불가 쌍 (선택한 D/E/N 근무에 한해 같은 날 동시 배치 금지)
- 부서별 근무표 생성 + 컬러 테이블 + 엑셀 다운로드
- st.session_state 영속 저장
- 전월 말 근무 이월(JSON) — 월 경계 N-D·연속근무 등
"""

import streamlit as st
import pandas as pd
import io
import json
import calendar as _calendar
from pathlib import Path

import app as _app                          # 전역 상수(YEAR/MONTH/NUM_DAYS) 동적 갱신
from app import (
    solve_schedule, get_april_days, validate_schedule,
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
    initial_sidebar_state="collapsed",
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

/* 메인 영역 가로 넓게 (근무 스케줄 표 시야 확보) */
section[data-testid="stMain"] .block-container {
    max-width: min(1920px, 100%) !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
}
section[data-testid="stMain"] [data-testid="stDataFrame"],
section[data-testid="stMain"] [data-testid="stDataEditor"] {
    width: 100% !important;
}

/* 메인 상단 패널 — select·체크 가독성 (연도·월·부서명 검정) */
section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border: 1.5px solid #616161 !important;
    border-radius: 6px !important;
    box-shadow: none !important;
    color: #000000 !important;
}
section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"] {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
}
section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"] p,
section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"] span,
section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"] div {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 600 !important;
}
section[data-testid="stMain"] [data-baseweb="select"] [role="combobox"] p,
section[data-testid="stMain"] [data-baseweb="select"] [role="combobox"] span,
section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] p {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 600 !important;
}
section[data-testid="stMain"] [data-testid="stCheckbox"] label span,
section[data-testid="stMain"] [data-testid="stCheckbox"] label p {
    color: #111111 !important;
}
section[data-testid="stMain"] [data-testid="stExpander"] summary,
section[data-testid="stMain"] [data-testid="stExpander"] summary span {
    color: #212121 !important;
}

/* 상단 설정 패널 — 작은 expander·버튼·한 줄 배치 */
section[data-testid="stMain"] [data-testid="stExpander"] details > summary {
    font-size: 12px !important;
    font-weight: 600 !important;
    padding: 0.2rem 0.45rem !important;
    min-height: 2rem !important;
    list-style: none;
}
section[data-testid="stMain"] [data-testid="stExpander"] [data-testid="stVerticalBlock"] {
    gap: 0.35rem !important;
}
section[data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"] {
    padding: 0.45rem 0.65rem !important;
}
section[data-testid="stMain"] [data-testid="stHorizontalBlock"] > div [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    min-height: 32px !important;
}
section[data-testid="stMain"] [data-testid="stHorizontalBlock"] > div div.stButton > button {
    min-height: 32px !important;
    font-size: 12px !important;
    padding: 4px 8px !important;
}

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

/* 사이드바 체크박스 라벨 — 테마에 밝은 글자가 묻는 경우 방지 */
section[data-testid="stSidebar"] [data-testid="stCheckbox"] label span,
section[data-testid="stSidebar"] [data-testid="stCheckbox"] label p {
    color: #111111 !important;
}

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

/* 함께 근무 불가 — selectbox(간호사 A/B) 보강 (플레이스홀더·화살표) */
section[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"] svg {
    fill: #111111 !important;
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
_SCHEDULE_ARCHIVE_PATH = Path(__file__).resolve().parent / "schedule_month_archive.json"
CARRY_AUTO_DAYS = 7


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
                    if not isinstance(row, (list, tuple)) or len(row) < 2:
                        continue
                    a, b = str(row[0]).strip(), str(row[1]).strip()
                    if not a or not b or a == b:
                        continue
                    names = sorted([a, b])
                    if len(row) >= 3 and isinstance(row[2], (list, tuple)):
                        sh = [x for x in row[2] if x in ("D", "E", "N")]
                        if not sh:
                            sh = ["D", "E", "N"]
                    else:
                        sh = ["D", "E", "N"]
                    sh = sorted(sh, key=lambda x: "DEN".index(x))
                    clean.append([names[0], names[1], sh])
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


def _fp_pairs_to_indices(nurse_names: list[str], pairs: list) -> list[tuple[int, int, frozenset]]:
    """이름 쌍(+적용 시프트) → (i, j, frozenset('D','E','N')) (수간호사 0번 제외)."""
    idx = {name: i for i, name in enumerate(nurse_names)}
    merged: dict[tuple[int, int], frozenset] = {}
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
        if len(row) >= 3 and isinstance(row[2], (list, tuple, set, frozenset)):
            sh = frozenset(x for x in row[2] if x in ("D", "E", "N"))
        else:
            sh = frozenset({"D", "E", "N"})
        if not sh:
            sh = frozenset({"D", "E", "N"})
        merged[key] = merged.get(key, frozenset()) | sh
    return [(i, j, s) for (i, j), s in merged.items()]


def _period_storage_key(year: int, month: int) -> str:
    """신청·생성 근무를 연·월마다 따로 보관할 때 사용 (월 바꿔도 다른 달 데이터 유지)."""
    return f"{int(year)}|{int(month)}"


def _migrate_period_stores_if_needed() -> None:
    """기존 세션: 부서→표 단일 저장 → 부서→연월→표."""
    y = st.session_state.sel_year
    m = st.session_state.sel_month
    pk = _period_storage_key(y, m)

    def _first_nonempty(d: dict):
        for v in d.values():
            if v is not None:
                return v
        return None

    dr = st.session_state.get("dept_requests")
    if isinstance(dr, dict) and dr:
        fn = _first_nonempty(dr)
        if fn is not None and not isinstance(fn, dict):
            new_dr = {}
            for dept, val in dr.items():
                new_dr[dept] = {}
                if val is not None and hasattr(val, "shape"):
                    new_dr[dept][pk] = val
            st.session_state.dept_requests = new_dr
        elif fn is None:
            st.session_state.dept_requests = {str(d): {} for d in dr}

    ds = st.session_state.get("dept_schedules")
    if isinstance(ds, dict) and ds:
        fn = _first_nonempty(ds)
        inner_is_bundle = (
            isinstance(fn, dict)
            and fn
            and isinstance(next(iter(fn.values())), dict)
            and "schedule" in next(iter(fn.values()))
        )
        if inner_is_bundle:
            pass
        elif fn is not None and isinstance(fn, dict) and "schedule" in fn:
            new_ds = {}
            for dept, val in ds.items():
                new_ds[dept] = {}
                if val is not None and isinstance(val, dict) and "schedule" in val:
                    new_ds[dept][pk] = val
            st.session_state.dept_schedules = new_ds
        elif fn is None:
            st.session_state.dept_schedules = {str(d): {} for d in ds}

    em = st.session_state.get("edit_mode")
    if isinstance(em, dict) and em:
        v0 = next(iter(em.values()))
        if not isinstance(v0, dict):
            st.session_state.edit_mode = {
                d: ({pk: bool(v)} if v else {}) for d, v in em.items()
            }


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
    _migrate_period_stores_if_needed()

_init_state()


def _parse_carry_in_text(raw: str, nurse_names: list[str]):
    """
    전월 말 근무 JSON → {간호사인덱스: [시프트,...]} 또는 None.
    파싱 실패 시 False.
    """
    if raw is None or not str(raw).strip():
        return None
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return False
    if not isinstance(data, dict):
        return False
    idx = {name: i for i, name in enumerate(nurse_names)}
    out = {}
    for k, v in data.items():
        if isinstance(k, bool):
            continue
        if isinstance(k, int):
            i = k
        else:
            ks = str(k).strip()
            if ks.isdigit():
                i = int(ks)
            elif ks in idx:
                i = idx[ks]
            else:
                continue
        if not isinstance(v, list):
            continue
        seq = [str(x).strip() for x in v if str(x).strip()]
        if seq:
            out[i] = seq
    return out if out else None


def _month_archive_key(year: int, month: int) -> str:
    return f"{int(year)}|{int(month)}"


def _prev_year_month(year: int, month: int) -> tuple[int, int]:
    if month <= 1:
        return year - 1, 12
    return year, month - 1


def _schedule_to_jsonable(sched: dict) -> dict:
    out = {}
    for n, row in sched.items():
        nk = str(int(n) if not isinstance(n, str) else n)
        out[nk] = {str(int(d) if not isinstance(d, str) else d): v for d, v in row.items()}
    return out


def _schedule_from_jsonable(data: dict) -> dict:
    out = {}
    for nk, row in (data or {}).items():
        n = int(nk)
        out[n] = {}
        for dk, v in (row or {}).items():
            out[n][int(dk)] = v
    return out


def _load_schedule_archive() -> dict:
    if not _SCHEDULE_ARCHIVE_PATH.is_file():
        return {}
    try:
        with open(_SCHEDULE_ARCHIVE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError, TypeError):
        return {}


def _save_schedule_archive(archive: dict) -> None:
    try:
        with open(_SCHEDULE_ARCHIVE_PATH, "w", encoding="utf-8") as f:
            json.dump(archive, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def _archive_put_month(dept: str, year: int, month: int, nurse_names: list[str], schedule: dict) -> None:
    """해당 연·월 근무표를 디스크 아카이브에 저장 (자동 이월용)."""
    if not dept or not nurse_names or not schedule:
        return
    arch = _load_schedule_archive()
    arch.setdefault(str(dept), {})[_month_archive_key(year, month)] = {
        "nurse_names": [str(x) for x in nurse_names],
        "schedule": _schedule_to_jsonable(schedule),
    }
    _save_schedule_archive(arch)


def _build_carry_from_prev_month(
    dept: str,
    year: int,
    month: int,
    nurse_names: list[str],
    n_days: int = CARRY_AUTO_DAYS,
) -> tuple[dict[int, list[str]] | None, str | None]:
    """
    직전 달 아카이브에서 마지막 n_days일 근무를 추출 → carry_in 형식.
    성공 시 (dict, None), 실패 시 (None, 메시지).
    """
    py, pm = _prev_year_month(year, month)
    arch = _load_schedule_archive()
    entry = arch.get(str(dept), {}).get(_month_archive_key(py, pm))
    if not entry:
        return None, f"{py}년 {pm}월에 저장된 근무표가 없습니다. 먼저 그 달에 근무표를 생성·저장하세요."
    old_names = [str(x) for x in (entry.get("nurse_names") or [])]
    sched = _schedule_from_jsonable(entry.get("schedule") or {})
    last_day = _calendar.monthrange(py, pm)[1]
    start_d = max(1, last_day - int(n_days) + 1)
    day_list = list(range(start_d, last_day + 1))
    name_to_si = {n: i for i, n in enumerate(old_names)}
    carry: dict[int, list[str]] = {}
    for i, nm in enumerate(nurse_names):
        si = name_to_si.get(str(nm).strip())
        seq = []
        for d in day_list:
            if si is None:
                seq.append("OF")
            else:
                v = sched.get(si, {}).get(d, "")
                seq.append(str(v).strip() if v not in (None, "") else "OF")
        carry[i] = seq
    return carry, None


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
            dbg, "min-width:36px;", dfg,
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
        '<div style="overflow-x:auto;width:100%;border-radius:10px;'
        'box-shadow:0 2px 12px rgba(0,0,0,0.09);-webkit-overflow-scrolling:touch;">'
        '<table style="border-collapse:collapse;font-size:12px;width:max-content;min-width:100%;">'
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
#  상단 설정 패널 (근무표·신청 표 바로 위, 가로 전체)
# ════════════════════════════════════════════════════════════════════════════════
generate_btn = False
_MONTH_NAMES = [
    "1월", "2월", "3월", "4월", "5월", "6월",
    "7월", "8월", "9월", "10월", "11월", "12월",
]

with st.container(border=True):
    # ── 제목 + 연·월 (한 줄·컴팩트) ─────────────────────────────────────────
    _h1, _h2, _h3, _h4 = st.columns([1.75, 0.55, 0.55, 1.15])
    with _h1:
        st.markdown(
            '<p style="margin:0;padding-top:6px;font-size:1.05rem;font-weight:800;color:#1A237E;">'
            "🏥 근무표 생성기</p>",
            unsafe_allow_html=True,
        )
    with _h2:
        sel_year = st.selectbox(
            "연도",
            list(range(2024, 2032)),
            index=list(range(2024, 2032)).index(st.session_state.sel_year),
            key="year_selectbox",
        )
    with _h3:
        sel_month = st.selectbox(
            "월",
            list(range(1, 13)),
            index=st.session_state.sel_month - 1,
            format_func=lambda m: _MONTH_NAMES[m - 1],
            key="month_selectbox",
        )
    with _h4:
        st.markdown(
            f'<p style="margin:0;padding-top:0.45rem;font-size:12px;font-weight:700;color:#333;">'
            f"📅 {sel_year}년 {_MONTH_NAMES[sel_month - 1]} "
            f"· {_calendar.monthrange(sel_year, sel_month)[1]}일</p>",
            unsafe_allow_html=True,
        )

    # 연·월 변경 시 앱 기간만 갱신 (신청·생성 근무는 부서×연월별로 유지)
    if sel_year != st.session_state.sel_year or sel_month != st.session_state.sel_month:
        st.session_state.sel_year  = sel_year
        st.session_state.sel_month = sel_month
        _app.set_period(sel_year, sel_month)
        st.rerun()

    st.divider()

    dept_list = list(st.session_state.departments.keys())
    try:
        active_idx = dept_list.index(st.session_state.active_dept)
    except ValueError:
        active_idx = 0

    # 가로 1행: 부서 선택 + 부서추가 + 명단 + 공휴일
    _r0a, _r0b, _r0c, _r0d = st.columns([2.0, 1.0, 1.05, 1.2])
    with _r0a:
        active_dept = st.selectbox(
            "현재 부서",
            dept_list,
            index=active_idx,
            key="dept_selectbox",
            label_visibility="collapsed",
        )
        st.session_state.active_dept = active_dept
        st.markdown(
            f'<p style="margin:2px 0 0 0;font-size:11px;color:#546E7A;">📂 {active_dept} · '
            f'{len(st.session_state.departments[active_dept])}명</p>',
            unsafe_allow_html=True,
        )
    nurses = st.session_state.departments[active_dept]
    gen = st.session_state.nurse_gen.get(active_dept, 0)

    with _r0b:
        with st.expander("➕ 부서", expanded=False):
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

    with _r0c:
        with st.expander(f"👩 명단({len(nurses)})", expanded=False):
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
                st.session_state.dept_requests[active_dept]  = {}
                st.session_state.dept_schedules[active_dept] = {}
                st.session_state.nurse_gen[active_dept]      = gen + 1
                st.rerun()

            # 이름 변경 동기화
            st.session_state.departments[active_dept] = updated_nurses
            _rq_pk = _period_storage_key(sel_year, sel_month)
            _rq_sub = st.session_state.dept_requests.setdefault(active_dept, {})
            if not isinstance(_rq_sub, dict):
                _rq_sub = {}
                st.session_state.dept_requests[active_dept] = _rq_sub
            df_existing = _rq_sub.get(_rq_pk)
            if df_existing is not None and len(df_existing) == len(updated_nurses):
                df_existing.index = updated_nurses

            # 간호사 추가 버튼
            st.markdown("<div style='margin-top:4px'></div>", unsafe_allow_html=True)
            if st.button("➕  간호사 추가", key="btn_add_nurse", use_container_width=True):
                new_idx = len(st.session_state.departments[active_dept])
                st.session_state.departments[active_dept].append(f"간호사{new_idx}")
                st.session_state.dept_requests[active_dept]  = {}
                st.session_state.dept_schedules[active_dept] = {}
                st.session_state.nurse_gen[active_dept]      = gen + 1
                st.rerun()

    with _r0d:
        with st.expander("📅 휴일", expanded=False):
            default_hols = st.session_state.dept_holidays.get(active_dept, "")
            holidays_raw = st.text_input(
                "공휴일",
                value=default_hols,
                key=f"holidays_{active_dept}",
                placeholder="5, 15, 25",
                label_visibility="collapsed",
            )
            st.session_state.dept_holidays[active_dept] = holidays_raw
            _hol_parsed = _parse_holidays(holidays_raw)
            if _hol_parsed:
                badge = " · ".join(f"{h}일" for h in _hol_parsed)
                st.markdown(
                    f'<div style="background:#E3F2FD;border:1px solid #90CAF9;border-radius:4px;'
                    f'padding:4px 8px;font-size:10px;color:#1565C0;">📌 {badge}</div>',
                    unsafe_allow_html=True,
                )

    # 가로 2행: 함께 근무 불가 | 전월 이월 | 부서 삭제 | 근무표 생성
    _r1a, _r1b, _r1c, _r1d = st.columns([2.5, 2.2, 0.45, 1.2])
    with _r1a:
        with st.expander("🙅 불가", expanded=False):
            # ════════════════════════════════════════════════════════════════════
            #  ③b 함께 근무 불가 (선택한 D/E/N에 한해 같은 날 동시 배치 금지)
            # ════════════════════════════════════════════════════════════════════
            st.markdown(
                '<p style="font-size:11px;font-weight:600;margin:0 0 2px 0;color:#212121;">'
                "🙅 함께 근무 불가</p>",
                unsafe_allow_html=True,
            )
            st.markdown(
                '<p style="font-size:10px;line-height:1.35;color:#616161;margin:0 0 6px 0;">'
                "미숙련 간호사 등 <strong>같은 날·같은 근무</strong>에 함께 서면 안 되는 두 명을 등록합니다. "
                "아래에서 <strong>D / E / N</strong> 중 적용할 근무를 고릅니다. "
                "(수간호사는 여기서 선택하지 않습니다)</p>",
                unsafe_allow_html=True,
            )
            _fp_list = st.session_state.dept_forbidden_pairs.setdefault(active_dept, [])
            _staff = [n for n in nurses if n != nurses[0]]
            # 첫 항목을 빈칸으로 두어, 드롭다운에서 고른 뒤에만 이름이 보이게 함
            _a_opts = [""] + _staff if _staff else [""]
            _sel_a = st.selectbox(
                "간호사 A",
                _a_opts,
                key=f"fp_a_{active_dept}",
                label_visibility="collapsed",
            )
            if _sel_a:
                _b_candidates = [x for x in _staff if x != _sel_a]
            else:
                _b_candidates = list(_staff)
            _b_opts = [""] + _b_candidates if _b_candidates else [""]
            _sel_b = st.selectbox(
                "간호사 B",
                _b_opts,
                key=f"fp_b_{active_dept}",
                label_visibility="collapsed",
            )
            st.markdown(
                '<p style="font-size:11px;font-weight:600;color:#111111;margin:8px 0 4px 0;">적용 근무</p>',
                unsafe_allow_html=True,
            )
            _fc1, _fc2, _fc3 = st.columns(3)
            with _fc1:
                _chk_d = st.checkbox("D 근무 불가", value=True, key=f"fp_shift_d_{active_dept}")
            with _fc2:
                _chk_e = st.checkbox("E 근무 불가", value=True, key=f"fp_shift_e_{active_dept}")
            with _fc3:
                _chk_n = st.checkbox("N 근무 불가", value=True, key=f"fp_shift_n_{active_dept}")
            _fp_shift_sel = [s for s, ok in (("D", _chk_d), ("E", _chk_e), ("N", _chk_n)) if ok]
            if st.button("➕ 추가", key=f"fp_add_{active_dept}", use_container_width=True):
                if _sel_a and _sel_b and _sel_a != _sel_b:
                    if not _fp_shift_sel:
                        st.warning("적용할 근무(D/E/N)를 하나 이상 선택해 주세요.")
                    else:
                        _pair_names = sorted([_sel_a, _sel_b])
                        _key = tuple(_pair_names)
                        _shifts = sorted(_fp_shift_sel, key=lambda x: "DEN".index(x))
                        _found_i = None
                        for _ix, _row in enumerate(_fp_list):
                            if (
                                isinstance(_row, (list, tuple)) and len(_row) >= 2
                                and tuple(sorted([str(_row[0]), str(_row[1])])) == _key
                            ):
                                _found_i = _ix
                                break
                        if _found_i is not None:
                            _old = _fp_list[_found_i]
                            _prev = (
                                set(_old[2]) if len(_old) >= 3 and isinstance(_old[2], list) else {"D", "E", "N"}
                            )
                            _merged = sorted(_prev | set(_shifts), key=lambda x: "DEN".index(x))
                            _fp_list[_found_i] = [_pair_names[0], _pair_names[1], _merged]
                        else:
                            _fp_list.append([_pair_names[0], _pair_names[1], _shifts])
                        _save_departments_to_disk()
                        st.rerun()
            if _fp_list:
                for _i, _pr in enumerate(list(_fp_list)):
                    _r1, _r2 = st.columns([5, 1])
                    with _r1:
                        _sh_disp = (
                            _pr[2]
                            if len(_pr) >= 3 and isinstance(_pr[2], list)
                            else ["D", "E", "N"]
                        )
                        _tags = "".join(
                            f'<span style="display:inline-block;background:#ECEFF1;padding:1px 6px;'
                            f'border-radius:4px;margin:2px 4px 0 0;font-size:9px;">{s} 불가</span>'
                            for s in _sh_disp
                        )
                        st.markdown(
                            f'<div style="font-size:10px;color:#37474F;padding:1px 0;line-height:1.35;">'
                            f"🔗 {_pr[0]} · {_pr[1]}<br/>{_tags}</div>",
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

    with _r1b:
        with st.expander("📎 이월", expanded=False):
            st.markdown(
                '<p style="font-size:10px;line-height:1.4;color:#616161;margin:0 0 6px 0;">'
                "직전 달 <strong>마지막 며칠</strong> 근무를 넣으면 이번 달 <strong>1일</strong>부터 "
                "N→D 금지·연속근무 등이 맞게 적용됩니다.</p>",
                unsafe_allow_html=True,
            )
            if st.button(
                f"📥 직전 달 마지막 {CARRY_AUTO_DAYS}일 자동",
                key=f"btn_carry_auto_{active_dept}",
                use_container_width=True,
            ):
                _co, _em = _build_carry_from_prev_month(
                    active_dept, sel_year, sel_month, nurses, CARRY_AUTO_DAYS,
                )
                if _em:
                    st.warning(_em)
                else:
                    st.session_state[f"carry_txt_{active_dept}"] = json.dumps(
                        {str(k): v for k, v in _co.items()},
                        ensure_ascii=False,
                    )
                    st.toast(
                        f"✅ 직전 달 마지막 {CARRY_AUTO_DAYS}일을 반영했습니다.",
                        icon="📎",
                    )
                    st.rerun()
            _cpy, _cpm = _prev_year_month(sel_year, sel_month)
            st.caption(f"저장분: **{_cpy}년 {_cpm}월**")
            st.text_area(
                "전월 말 JSON",
                height=90,
                key=f"carry_txt_{active_dept}",
                placeholder=(
                    '{"0": ["OF"], "1": ["N", "N", "OF"], "2": ["D", "E"]}  ← 수간=0, 간호사=1…'
                ),
                label_visibility="collapsed",
            )

    with _r1c:
        st.markdown("<div style='min-height:2.2rem'></div>", unsafe_allow_html=True)
        if len(dept_list) > 1:
            if st.button(
                "🗑️",
                key="btn_del_dept",
                help=f"'{active_dept}' 부서 삭제 (데이터 포함)",
            ):
                for store in ("dept_schedules", "dept_requests", "dept_holidays", "nurse_gen"):
                    st.session_state[store].pop(active_dept, None)
                st.session_state.dept_forbidden_pairs.pop(active_dept, None)
                del st.session_state.departments[active_dept]
                st.session_state.active_dept = list(st.session_state.departments.keys())[0]
                _save_departments_to_disk()
                st.rerun()
        else:
            st.caption("—")

    with _r1d:
        generate_btn = st.button(
            "🗓️ 생성",
            type="primary",
            use_container_width=True,
            key="btn_generate",
        )
        _hint_pk = _period_storage_key(st.session_state.sel_year, st.session_state.sel_month)
        _hint_sub = st.session_state.dept_schedules.get(active_dept, {})
        if isinstance(_hint_sub, dict) and _hint_sub.get(_hint_pk):
            st.caption("✅ 생성됨")

    holidays = _parse_holidays(st.session_state.dept_holidays.get(active_dept, ""))


# ════════════════════════════════════════════════════════════════════════════════
#  MAIN – 변수 준비
# ════════════════════════════════════════════════════════════════════════════════
nurses      = st.session_state.departments[active_dept]   # 최신 명단
num_nurses  = len(nurses)
days        = get_april_days(holidays)
col_labels  = [_day_label(d) for d in days]
gen         = st.session_state.nurse_gen.get(active_dept, 0)
_period_pk  = _period_storage_key(st.session_state.sel_year, st.session_state.sel_month)
editor_key  = f"req_editor_{active_dept}_n{num_nurses}_g{gen}_{_period_pk}"

# requests_df 준비 (없거나 행 수 불일치 시 새로 생성) — 부서×연월별 유지
_rq_sub = st.session_state.dept_requests.setdefault(active_dept, {})
if not isinstance(_rq_sub, dict):
    _rq_sub = {}
    st.session_state.dept_requests[active_dept] = _rq_sub
df_req = _rq_sub.get(_period_pk)
if df_req is None or df_req.shape[0] != num_nurses:
    df_req = _make_requests_df(nurses, days)
    _rq_sub[_period_pk] = df_req
else:
    df_req.index   = nurses
    df_req.columns = col_labels

# None / nan → 공백으로 정규화
df_req = df_req.apply(lambda col: col.map(
    lambda x: "" if (x is None or str(x).strip() in ("None", "nan")) else str(x).strip()
))
_rq_sub[_period_pk] = df_req

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
    height=min(40 * num_nurses + 90, 720),
    key=editor_key,
    num_rows="fixed",
)

# 저장 영역 (전체 너비 — 좁은 열에 넣으면 버튼이 안 보이는 경우가 있음)
req_saved_key = f"req_saved_{active_dept}_{_period_pk}_g{gen}"

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
        _rq_sub[_period_pk] = cleaned
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
            _rq_sub[_period_pk] = _make_requests_df(nurses, days)
            st.session_state[req_saved_key] = False
            st.rerun()

# ════════════════════════════════════════════════════════════════════════════════
#  근무표 생성 처리
# ════════════════════════════════════════════════════════════════════════════════
if generate_btn:
    # 저장된 신청 근무 사용 (버튼으로 확정된 데이터)
    _saved_req = _rq_sub.get(_period_pk)
    saved_df = edited_df if _saved_req is None else _saved_req
    requests = _df_to_requests(saved_df, days)
    _fp_idx = _fp_pairs_to_indices(
        nurses,
        st.session_state.dept_forbidden_pairs.get(active_dept, []),
    )
    _carry_raw = st.session_state.get(f"carry_txt_{active_dept}", "") or ""
    _carry_in = _parse_carry_in_text(_carry_raw, nurses)
    if _carry_in is False:
        st.error("전월 말 근무(JSON) 형식이 올바르지 않습니다. 중괄호·쉼표·따옴표를 확인해 주세요.")
    else:
        with st.spinner("⏳ 근무표를 계산하는 중입니다…"):
            schedule, success, status = solve_schedule(
                num_nurses, requests, holidays,
                forbidden_pairs=_fp_idx or None,
                carry_in=_carry_in,
            )
        if success:
            st.session_state.dept_schedules.setdefault(active_dept, {})[_period_pk] = {
                "schedule":    schedule,
                "nurse_names": nurses.copy(),
                "holidays":    holidays,
                "requests":    requests,
            }
            _archive_put_month(
                active_dept,
                st.session_state.sel_year,
                st.session_state.sel_month,
                nurses,
                schedule,
            )
            issues = validate_schedule(
                schedule, num_nurses, holidays,
                forbidden_pairs=_fp_idx or None,
                nurse_names=nurses,
                carry_in=_carry_in,
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
_sched_sub = st.session_state.dept_schedules.get(active_dept, {})
sched_data = _sched_sub.get(_period_pk) if isinstance(_sched_sub, dict) else None

if sched_data:
    schedule    = sched_data["schedule"]
    sched_names = sched_data["nurse_names"]
    sched_hols  = sched_data["holidays"]
    sched_reqs  = sched_data.get("requests", {})
    sched_days  = get_april_days(sched_hols)
    sched_n     = len(sched_names)

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── 수정 모드 상태 (부서×연월)
    _em_sub = st.session_state.edit_mode.setdefault(active_dept, {})
    if not isinstance(_em_sub, dict):
        _em_sub = {}
        st.session_state.edit_mode[active_dept] = _em_sub
    is_edit = _em_sub.get(_period_pk, False)

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
                _em_sub[_period_pk] = True
                st.rerun()
        else:
            if st.button("❌ 취소", use_container_width=True, key="btn_edit_off"):
                _em_sub[_period_pk] = False
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
            key=f"sched_editor_{active_dept}_{_period_pk}",
            num_rows="fixed",
        )

        # 저장 버튼
        save_col, _ = st.columns([1, 3])
        with save_col:
            if st.button("💾 저장", type="primary", use_container_width=True, key="btn_save_edit"):
                new_schedule = _edit_df_to_schedule(edited, sched_days)
                st.session_state.dept_schedules.setdefault(active_dept, {})[_period_pk]["schedule"] = new_schedule
                _archive_put_month(
                    active_dept,
                    st.session_state.sel_year,
                    st.session_state.sel_month,
                    sched_names,
                    new_schedule,
                )
                # 재검증
                _fp_ed = _fp_pairs_to_indices(
                    sched_names,
                    st.session_state.dept_forbidden_pairs.get(active_dept, []),
                )
                _carry_ed = _parse_carry_in_text(
                    st.session_state.get(f"carry_txt_{active_dept}", "") or "",
                    sched_names,
                )
                _carry_for_v = None if _carry_ed is False else _carry_ed
                issues = validate_schedule(
                    new_schedule, sched_n, sched_hols,
                    forbidden_pairs=_fp_ed or None,
                    nurse_names=sched_names,
                    carry_in=_carry_for_v,
                )
                st.session_state.violations     = issues
                st.session_state.show_violations = bool(issues)
                _em_sub[_period_pk] = False
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
    /* 메인 — 연도·월·부서 select 표시 글자 검정 (테마 덮어쓰기) */
    section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"],
    section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"] p,
    section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] [role="combobox"] span {
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
        opacity: 1 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# 부서·간호사 명단을 JSON에 저장 (페이지 로드마다 최신 상태 유지)
_save_departments_to_disk()
