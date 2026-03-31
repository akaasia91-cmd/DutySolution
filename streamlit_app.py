"""
응급실 근무표 생성기 – Streamlit UI v2
- 연도·월 선택 가능
- 부서(Department) CRUD
- 간호사(Staff) CRUD: 추가 / 이름 수정 / 삭제
- 부서별 신청 근무 입력 달력 (data_editor)
- 함께 근무 불가 그룹(2~4명, 수간 포함 / 선택한 D/E/N에 한해 같은 날 동시 배치 금지)
- 부서별 근무표 생성 + 컬러 테이블 + 엑셀 다운로드
- st.session_state 영속 저장
- 전월 말 근무 이월(JSON) — 월 경계 N-D·연속근무 등
"""

import streamlit as st
import streamlit.components.v1 as components
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

/* 빈 사이드바 숨김 + 메인 가로 전체 사용 (접기 버튼·열 제거) */
section[data-testid="stSidebar"] {
    display: none !important;
    width: 0 !important;
    min-width: 0 !important;
}
[data-testid="stSidebarCollapsedControl"],
[data-testid="collapsedControl"] {
    display: none !important;
}
section[data-testid="stMain"] {
    width: 100% !important;
    max-width: 100% !important;
    margin-left: 0 !important;
}
[data-testid="stAppViewContainer"] {
    padding-left: 0 !important;
    padding-right: 0 !important;
}
header[data-testid="stHeader"] {
    padding-left: 0.35rem !important;
    padding-right: 0.35rem !important;
    padding-top: 0.25rem !important;
    padding-bottom: 0.25rem !important;
}

/* 메인 영역 — 상하좌우 여백 최소화 */
section[data-testid="stMain"] .block-container {
    max-width: 100% !important;
    padding: 0.12rem 0.2rem 0.25rem 0.2rem !important;
}
section[data-testid="stMain"] [data-testid="stVerticalBlock"] {
    gap: 0.12rem !important;
    row-gap: 0.12rem !important;
}
/* 상단 테두리 패널 안만 위젯 세로 간격 축소 */
section[data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stElementContainer"] {
    margin-bottom: 0 !important;
    margin-top: 0 !important;
}
section[data-testid="stMain"] hr,
section[data-testid="stMain"] [data-testid="stHorizontalRule"] {
    margin: 0.12rem 0 !important;
}
section[data-testid="stMain"] [data-testid="stDataFrame"],
section[data-testid="stMain"] [data-testid="stDataEditor"] {
    width: 100% !important;
}

/* 메인 상단 패널 — select·체크 가독성 (연도·월·부서명 검정) */
section[data-testid="stMain"] [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border: 1px solid #757575 !important;
    border-radius: 4px !important;
    box-shadow: none !important;
    color: #000000 !important;
    min-height: 24px !important;
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

/* 상단 설정 패널(부서·연월) — 최대 압축 */
section[data-testid="stMain"] [data-testid="stExpander"] details > summary {
    font-size: 10px !important;
    font-weight: 600 !important;
    padding: 0.02rem 0.22rem !important;
    min-height: 1.35rem !important;
    list-style: none;
}
section[data-testid="stMain"] [data-testid="stExpander"] [data-testid="stVerticalBlock"] {
    gap: 0.12rem !important;
}
section[data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"] {
    padding: 0.08rem 0.18rem !important;
    margin-bottom: 0.06rem !important;
}
section[data-testid="stMain"] [data-testid="stHorizontalBlock"] {
    gap: 0.1rem !important;
    row-gap: 0.1rem !important;
}
section[data-testid="stMain"] [data-testid="stHorizontalBlock"] > div [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    min-height: 24px !important;
}
section[data-testid="stMain"] [data-testid="stHorizontalBlock"] > div div.stButton > button {
    min-height: 24px !important;
    font-size: 10px !important;
    padding: 1px 4px !important;
}

/* 사이드바 — Streamlit CSS 변수(다크 텍스트 색이 입력에 전달되도록) */
section[data-testid="stSidebar"] {
    --text-color: #262730 !important;
    --stTextColor: #262730 !important;
    --widget-text-color: #000000 !important;
}

/* 사이드바 — 흰색 배경 + 선명한 검정 계열 글자 */
section[data-testid="stSidebar"] > div:first-child,
section[data-testid="stSidebar"] [data-testid="stSidebarContent"],
section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"] {
    background: #ffffff !important;
    border-right: 1px solid #e0e0e0;
}
section[data-testid="stSidebar"] {
    color: #212121 !important;
    background-color: #ffffff !important;
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
/* multiselect(Choose options) 패널 전체 흰 배경 — 메인·사이드바 공통 */
div[data-baseweb="popover"] {
    background-color: #ffffff !important;
    box-shadow: 0 4px 16px rgba(0,0,0,0.12) !important;
}
div[data-baseweb="popover"] [data-baseweb="menu"],
div[data-baseweb="popover"] ul {
    background-color: #ffffff !important;
}
/* 함께 근무 불가: 안내 문단과 multiselect 겹침 방지 */
section[data-testid="stMain"] [data-testid="stExpanderDetails"] {
    overflow: visible !important;
}
.fp-multiselect-anchor {
    height: 6px;
    min-height: 6px;
    display: block;
}
section[data-testid="stMain"] [data-testid="stMultiSelect"] {
    margin-top: 2px !important;
    margin-bottom: 8px !important;
    position: relative;
    z-index: 1;
}
section[data-testid="stMain"] [data-testid="stMultiSelect"] [data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border: 1px solid #757575 !important;
    border-radius: 6px !important;
    box-shadow: none !important;
    min-height: 38px !important;
}
section[data-testid="stMain"] [data-testid="stMultiSelect"] [data-baseweb="select"] [role="combobox"] {
    color: #111111 !important;
    -webkit-text-fill-color: #111111 !important;
}
section[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-baseweb="select"] > div {
    background-color: #ffffff !important;
    border: 1.5px solid #616161 !important;
    border-radius: 6px !important;
    box-shadow: none !important;
}
section[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-baseweb="select"] [role="combobox"] {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
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

/* 신청·생성 근무 data_editor — 이름 열 + 1~말일 한 화면 (가로 균등·날짜 헤더 세로) */
section[data-testid="stMain"] div[data-testid="stDataFrame"] table {
    table-layout: fixed !important;
    width: 100% !important;
    max-width: 100% !important;
}
section[data-testid="stMain"] div[data-testid="stDataFrame"] thead th:not(:first-child) {
    writing-mode: vertical-rl !important;
    text-orientation: mixed !important;
    transform: rotate(180deg);
    font-size: 6px !important;
    font-weight: 700 !important;
    padding: 0 1px !important;
    height: 2.65em !important;
    line-height: 1 !important;
    vertical-align: middle !important;
}
section[data-testid="stMain"] div[data-testid="stDataFrame"] thead th:first-child {
    width: 13% !important;
    min-width: 7.25rem !important;
    max-width: none !important;
    font-size: 8px !important;
    padding: 1px 4px !important;
    line-height: 1.15 !important;
    vertical-align: bottom !important;
    overflow: visible !important;
    white-space: normal !important;
    word-break: keep-all !important;
}
section[data-testid="stMain"] div[data-testid="stDataFrame"] td,
section[data-testid="stMain"] div[data-testid="stDataFrame"] tbody th {
    font-size: 7px !important;
    padding: 0 !important;
    line-height: 1.05 !important;
}
section[data-testid="stMain"] div[data-testid="stDataFrame"] td:first-child,
section[data-testid="stMain"] div[data-testid="stDataFrame"] tbody th:first-child {
    width: 13% !important;
    min-width: 7.25rem !important;
    max-width: none !important;
    font-size: 8px !important;
    padding: 1px 4px !important;
    white-space: normal !important;
    word-break: keep-all !important;
    overflow: visible !important;
    overflow-x: visible !important;
    overflow-y: visible !important;
    text-overflow: unset !important;
}
/* 이름 셀 내부 래퍼 — 가로·세로 스크롤바 없이 전체 표시 */
section[data-testid="stMain"] div[data-testid="stDataFrame"] td:first-child > div,
section[data-testid="stMain"] div[data-testid="stDataFrame"] tr td:first-child [data-testid="cell"] {
    overflow: visible !important;
    overflow-x: visible !important;
    overflow-y: visible !important;
    max-height: none !important;
    white-space: normal !important;
    word-break: keep-all !important;
}
section[data-testid="stMain"] div[data-testid="stDataFrame"] [data-baseweb="select"] > div {
    font-size: 7px !important;
    min-height: 15px !important;
    padding: 0 1px !important;
}
section[data-testid="stMain"] div[data-testid="stDataFrame"] [data-baseweb="select"] [role="combobox"],
section[data-testid="stMain"] div[data-testid="stDataFrame"] [data-baseweb="select"] p {
    font-size: 7px !important;
    line-height: 1.05 !important;
}
section[data-testid="stMain"] div[data-testid="stDataFrame"] [data-baseweb="popover"] li,
section[data-testid="stMain"] div[data-testid="stDataFrame"] ul[role="listbox"] li {
    font-size: 8px !important;
    min-height: 18px !important;
    padding: 1px 5px !important;
}
/* 내부 가로 스크롤 최소화(균등 분배 우선) */
section[data-testid="stMain"] [data-testid="stDataEditor"] > div [data-testid="stHorizontalBlock"],
section[data-testid="stMain"] [data-testid="stDataFrame"] {
    max-width: 100% !important;
}

/* 신청 근무 확정 박스 — 본문 글자 검정 (테마 간섭 방지) */
.req-save-panel, .req-save-panel h4, .req-save-panel p,
.req-save-status { color: #111111 !important; }

/* 생성된 근무표(HTML 미리보기) — 가로 스크롤(말일·합계 열까지) */
section[data-testid="stMain"] .duty-generated-schedule-wrap {
    overflow-x: scroll !important;
    overflow-y: hidden;
    width: 100% !important;
    max-width: 100% !important;
    min-width: 0 !important;
    box-sizing: border-box !important;
    -webkit-overflow-scrolling: touch;
    scrollbar-gutter: stable;
}
section[data-testid="stMain"] .duty-generated-schedule-wrap table {
    width: max-content !important;
    min-width: unset !important;
    max-width: none !important;
    table-layout: auto !important;
}
section[data-testid="stMain"] [data-testid="stMarkdownContainer"]:has(.duty-generated-schedule-wrap),
section[data-testid="stMain"] [data-testid="stElementContainer"]:has(.duty-generated-schedule-wrap) {
    overflow-x: visible !important;
    max-width: 100% !important;
}
/* 생성 근무표 편집 data_editor — 바로 다음 블록에 가로 스크롤 허용 */
section[data-testid="stMain"] [data-testid="stElementContainer"]:has(.duty-schedule-editor-hscroll)
    + [data-testid="stElementContainer"] [data-testid="stDataFrame"] {
    overflow-x: auto !important;
    max-width: 100% !important;
}
section[data-testid="stMain"] [data-testid="stElementContainer"]:has(.duty-schedule-editor-hscroll)
    + [data-testid="stElementContainer"] [data-testid="stDataFrame"] > div {
    overflow-x: auto !important;
    max-width: 100% !important;
}
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
                    if isinstance(row[0], list):
                        names = sorted({str(x).strip() for x in row[0] if str(x).strip()})
                        sh_src = row[1] if len(row) > 1 else None
                    else:
                        a, b = str(row[0]).strip(), str(row[1]).strip()
                        if not a or not b or a == b:
                            continue
                        names = sorted({a, b})
                        sh_src = row[2] if len(row) > 2 else None
                    if len(names) < 2 or len(names) > 4:
                        continue
                    if isinstance(sh_src, (list, tuple)):
                        sh = [x for x in sh_src if x in ("D", "E", "N")]
                        if not sh:
                            sh = ["D", "E", "N"]
                    else:
                        sh = ["D", "E", "N"]
                    sh = sorted(sh, key=lambda x: "DEN".index(x))
                    clean.append([names, sh])
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


def _fp_row_names_from_entry(row) -> list[str] | None:
    """저장 행 → 정렬된 고유 이름 2~4명 또는 None."""
    if not row or not isinstance(row, (list, tuple)) or len(row) < 2:
        return None
    if isinstance(row[0], list):
        names = sorted({str(x).strip() for x in row[0] if str(x).strip()})
    else:
        a, b = str(row[0]).strip(), str(row[1]).strip()
        if not a or not b:
            return None
        names = sorted({a, b})
    if len(names) < 2 or len(names) > 4:
        return None
    return names


def _fp_pairs_to_indices(nurse_names: list[str], pairs: list) -> list[tuple[int, int, frozenset]]:
    """이름 그룹 2~4명(+적용 시프트) → 쌍 전개 (i, j, frozenset('D','E','N')). 수간호사 포함."""
    idx = {name: i for i, name in enumerate(nurse_names)}
    merged: dict[tuple[int, int], frozenset] = {}
    for row in pairs or []:
        names = _fp_row_names_from_entry(row)
        if not names:
            continue
        inds: list[int] = []
        bad = False
        for nm in names:
            if nm not in idx:
                bad = True
                break
            inds.append(idx[nm])
        if bad:
            continue
        inds = sorted(set(inds))
        if len(inds) < 2:
            continue
        if isinstance(row[0], list):
            sh_raw = row[1] if len(row) > 1 else None
        else:
            sh_raw = row[2] if len(row) > 2 else None
        if isinstance(sh_raw, (list, tuple, set, frozenset)):
            sh = frozenset(x for x in sh_raw if x in ("D", "E", "N"))
        else:
            sh = frozenset({"D", "E", "N"})
        if not sh:
            sh = frozenset({"D", "E", "N"})
        for ia in range(len(inds)):
            for ib in range(ia + 1, len(inds)):
                i, j = inds[ia], inds[ib]
                key = (min(i, j), max(i, j))
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


def _day_label_compact(day: dict) -> str:
    """신청 근무 표 헤더용 — 가로 폭 최소 (일만 + 표시)."""
    d = day["day"]
    if day["is_holiday"]:
        return f"{d}♦"
    if day["is_weekend"]:
        return f"{d}·"
    return str(d)


def _monday_week_split_style(day: dict) -> str:
    """월요일(weekday==0) 칸 왼쪽 — 일요일·월요일 사이 주 구분 빨간 굵은 세로선."""
    if day.get("weekday") == 0:
        # border shorthand 이후에 left만 덮어씀 (일·월 사이 한 줄)
        return "border-left:6px solid #B71C1C;"
    return ""


def _inject_week_split_css(days: list) -> None:
    """st.data_editor: 월요일 열 왼쪽 굵은 빨간 선 (인덱스 열 다음 = nth-child 2가 1일)."""
    indices = [j for j, d in enumerate(days) if d.get("weekday") == 0]
    if not indices:
        return
    parts: list[str] = []
    for j in indices:
        n = j + 2  # 1=행 이름, 2=1일, …
        parts.extend(
            [
                f'section[data-testid="stMain"] [data-testid="stDataFrame"] thead th:nth-child({n})',
                f'section[data-testid="stMain"] [data-testid="stDataFrame"] tbody td:nth-child({n})',
                f'section[data-testid="stMain"] [data-testid="stDataFrame"] tr > th:nth-child({n})',
                f'section[data-testid="stMain"] [data-testid="stDataFrame"] tr > td:nth-child({n})',
                f'[data-testid="stDataFrame"] thead th:nth-child({n})',
                f'[data-testid="stDataFrame"] tbody td:nth-child({n})',
            ]
        )
    st.markdown(
        "<style>"
        + ",\n".join(parts)
        + " {\n  border-left: 6px solid #B71C1C !important;\n"
        "  box-shadow: none !important;\n}\n</style>",
        unsafe_allow_html=True,
    )


def _make_requests_df(nurses: list[str], days: list) -> pd.DataFrame:
    num_days = _app.NUM_DAYS
    return pd.DataFrame(
        [[""] * num_days for _ in range(len(nurses))],
        index=nurses,
        columns=[_day_label_compact(d) for d in days],
    )


def _df_to_requests(df: pd.DataFrame, days: list) -> dict:
    result = {}
    for i in range(len(df)):
        for j, day in enumerate(days):
            val = str(df.iloc[i, j]).strip()
            if val and val not in ("", "None", "nan"):
                result.setdefault(i, {})[day["day"]] = val
    return result


# 근무표 HTML 미리보기·엑셀 다운로드 공통 셀 색
_PREVIEW_FG_BLACK = "#000000"
_PREVIEW_BG_DE = "#FFFFFF"
_PREVIEW_BG_OF_PINK = "#F8BBD0"
_PREVIEW_BG_LEAVE_YELLOW = "#FFF59D"


def _preview_shift_bg_fg(shift: str) -> tuple[str, str]:
    """미리보기 셀 (배경, 글자). N만 기존 앱 색 유지, 나머지 글자는 검정."""
    if not shift:
        return "#FFFFFF", "#BDBDBD"
    if shift == "N":
        return (
            SHIFT_COLORS.get("N", "#E8EAF6"),
            SHIFT_TEXT_COLORS.get("N", "#283593"),
        )
    if shift in ("D", "E"):
        return _PREVIEW_BG_DE, _PREVIEW_FG_BLACK
    if shift in ("OF", "NO"):
        return _PREVIEW_BG_OF_PINK, _PREVIEW_FG_BLACK
    if shift in ("연", "공", "EDU", "경"):
        return _PREVIEW_BG_LEAVE_YELLOW, _PREVIEW_FG_BLACK
    bg = SHIFT_COLORS.get(shift, "#FFFFFF")
    return bg, _PREVIEW_FG_BLACK


def _render_schedule_html(schedule: dict, nurse_names: list, days: list,
                          requests: dict | None = None) -> str:
    num = len(nurse_names)
    requests = requests or {}
    th = lambda txt, bg, extra="", fg="#37474F": (
        f'<th style="background:{bg};color:{fg};padding:4px 2px;'
        f'text-align:center;white-space:nowrap;{extra}">{txt}</th>'
    )
    _hdr: list[str] = ["<tr>"]
    _hdr.append(th("간호사", "#ECEFF1",
                   "min-width:80px;padding:5px 8px;position:sticky;left:0;z-index:5;", "#263238"))
    for day in days:
        if day["is_holiday"]:
            dbg, dfg = "#FFEBEE", "#C62828"
        elif day["is_weekend"]:
            dbg, dfg = "#E3F2FD", "#1565C0"
        else:
            dbg, dfg = "#F5F5F5", "#455A64"
        _wsp = _monday_week_split_style(day)
        _hdr.append(th(
            f"{day['day']}<br><span style='font-size:9px'>{day['weekday_name']}</span>",
            dbg, f"min-width:36px;{_wsp}", dfg,
        ))
    for lbl in ["N", "D", "E", "OF", "OH", "연"]:
        bg, fg = _preview_shift_bg_fg(lbl)
        _hdr.append(th(
            f"{lbl}<br><span style='font-size:9px'>합계</span>",
            bg, "min-width:36px;", fg,
        ))
    _hdr.append("</tr>")
    _header_html = "".join(_hdr)
    _body: list[str] = []

    for n_idx, name in enumerate(nurse_names):
        ns = schedule.get(n_idx, {})
        counts = {"N": 0, "D": 0, "E": 0, "OF": 0, "OH": 0, "연": 0}
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
            bg, fg = _preview_shift_bg_fg(shift)
            # 신청 근무이면 밑줄 표시
            is_requested = nurse_req.get(d_num) == shift and shift != ""
            underline = "text-decoration:underline;text-underline-offset:3px;" if is_requested else ""
            _wsp = _monday_week_split_style(day)
            cells.append(
                f'<td style="background:{bg};color:{fg};font-weight:700;{underline}'
                f'padding:3px 1px;text-align:center;border:1px solid #E0E0E0;{_wsp}">{shift}</td>'
            )
            if shift == "N":          counts["N"] += 1
            elif shift == "D":        counts["D"] += 1
            elif shift == "E":        counts["E"] += 1
            elif shift in ("OF", "NO"):
                counts["OF"] += 1
            elif shift == "OH":
                counts["OH"] += 1
            elif shift == "연":
                counts["연"] += 1

        for key in ["N", "D", "E", "OF", "OH", "연"]:
            bg, fg = _preview_shift_bg_fg(key)
            cells.append(
                f'<td style="background:{bg};color:{fg};font-weight:700;'
                f'text-align:center;padding:3px;">{counts[key]}</td>'
            )
        _body.append(f'<tr style="background:{row_bg};">' + "".join(cells) + "</tr>")

    for lbl, sk in [("D인원", "D"), ("E인원", "E"), ("N인원", "N")]:
        hbg, hfg = _preview_shift_bg_fg(sk)
        bg, data_fg = hbg, hfg
        if sk in ("D", "E"):
            hbg = _PREVIEW_BG_DE
            hfg = _PREVIEW_FG_BLACK
            bg = _PREVIEW_BG_DE
            data_fg = _PREVIEW_FG_BLACK
        cells = [
            f'<td style="background:{hbg};color:{hfg};font-weight:700;'
            f'padding:4px 8px;white-space:nowrap;position:sticky;left:0;'
            f'border-right:2px solid #CFD8DC;">{lbl}</td>'
        ]
        for day in days:
            cnt = sum(1 for n in range(num) if schedule.get(n, {}).get(day["day"]) == sk)
            _wsp = _monday_week_split_style(day)
            cells.append(
                f'<td style="background:{bg};color:{data_fg};font-weight:700;text-align:center;'
                f'padding:3px;border:1px solid #E0E0E0;{_wsp}">{cnt}</td>'
            )
        cells += ["<td></td>"] * 6
        _body.append("<tr>" + "".join(cells) + "</tr>")

    return (
        '<div class="duty-generated-schedule-wrap" style="overflow-x:scroll;width:100%;max-width:100%;'
        'min-width:0;box-sizing:border-box;border-radius:10px;'
        'box-shadow:0 2px 12px rgba(0,0,0,0.09);-webkit-overflow-scrolling:touch;">'
        '<table style="border-collapse:collapse;font-size:12px;width:max-content;">'
        "<thead>" + _header_html + "</thead>"
        "<tbody>" + "".join(_body) + "</tbody>"
        "</table></div>"
    )


def _render_requests_preview_html(df: pd.DataFrame, nurse_names: list, days: list) -> str:
    """신청 근무 data_editor 내용 — 생성 근무표와 동일한 헤더·색상(합계 열 없음)."""
    col_labels = [_day_label_compact(d) for d in days]
    dfn = df.reindex(index=list(nurse_names), columns=col_labels, fill_value="")
    th = lambda txt, bg, extra="", fg="#37474F": (
        f'<th style="background:{bg};color:{fg};padding:4px 2px;'
        f'text-align:center;white-space:nowrap;{extra}">{txt}</th>'
    )
    _hdr: list[str] = ["<tr>"]
    _hdr.append(th("간호사", "#ECEFF1",
                   "min-width:80px;padding:5px 8px;position:sticky;left:0;z-index:5;", "#263238"))
    for day in days:
        if day["is_holiday"]:
            dbg, dfg = "#FFEBEE", "#C62828"
        elif day["is_weekend"]:
            dbg, dfg = "#E3F2FD", "#1565C0"
        else:
            dbg, dfg = "#F5F5F5", "#455A64"
        _wsp = _monday_week_split_style(day)
        _hdr.append(th(
            f"{day['day']}<br><span style='font-size:9px'>{day['weekday_name']}</span>",
            dbg, f"min-width:36px;{_wsp}", dfg,
        ))
    _hdr.append("</tr>")
    _header_html = "".join(_hdr)
    _body: list[str] = []
    for n_idx, name in enumerate(nurse_names):
        row_bg = "#FAFAFA" if n_idx % 2 == 0 else "#F5F5F5"
        cells = [
            f'<td style="background:#ECEFF1;color:#263238;font-weight:700;'
            f'padding:4px 8px;white-space:nowrap;position:sticky;left:0;'
            f'border-right:2px solid #CFD8DC;">{name}</td>'
        ]
        for j, day in enumerate(days):
            raw = dfn.iat[n_idx, j]
            shift = "" if (raw is None or str(raw).strip() in ("", "None", "nan")) else str(raw).strip()
            if shift:
                bg, fg = _preview_shift_bg_fg(shift)
            else:
                bg, fg = "#FFFFFF", "#BDBDBD"
            _wsp = _monday_week_split_style(day)
            cells.append(
                f'<td style="background:{bg};color:{fg};font-weight:700;'
                f'padding:3px 1px;text-align:center;border:1px solid #E0E0E0;{_wsp}">{shift}</td>'
            )
        _body.append(f'<tr style="background:{row_bg};">' + "".join(cells) + "</tr>")
    return (
        '<div class="duty-generated-schedule-wrap" style="overflow-x:scroll;width:100%;max-width:100%;'
        'min-width:0;box-sizing:border-box;border-radius:10px;'
        'box-shadow:0 2px 12px rgba(0,0,0,0.09);-webkit-overflow-scrolling:touch;">'
        '<table style="border-collapse:collapse;font-size:12px;width:max-content;">'
        "<thead>" + _header_html + "</thead>"
        "<tbody>" + "".join(_body) + "</tbody>"
        "</table></div>"
    )


def _show_schedule_preview_iframe(
    html_fragment: str, num_nurses: int, *, extra_rows: int = 5,
) -> None:
    """Streamlit 본문이 표 너비까지 늘어나 `st.markdown` 가로 스크롤이 안 생기는 경우 — iframe에서 스크롤."""
    # srcdoc 안에서 스크립트 태그로 잘못 해석되는 경우 방지
    safe = html_fragment.replace("</script>", "<\\/script>")
    # 간호사 행 + (생성표: 합계·요약 행 / 신청표: 헤더만) — extra_rows로 여유 조절
    _h = min(72 + max(num_nurses + extra_rows, 8) * 28, 1400)
    doc = (
        "<!DOCTYPE html><html><head><meta charset=\"utf-8\"/>"
        "<style>"
        "html,body{margin:0;padding:6px;background:#fafafa;}"
        "body{overflow:auto;overflow-x:auto;overflow-y:auto;-webkit-overflow-scrolling:touch;}"
        ".duty-generated-schedule-wrap{overflow:visible!important;width:max-content!important;"
        "max-width:none!important;min-width:0;}"
        "</style></head><body>"
        f"{safe}</body></html>"
    )
    components.html(doc, width=None, height=_h, scrolling=True)


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

    def _px(sk: str) -> tuple[str, str]:
        """미리보기와 동일 (배경/글자 HEX, 알파벳 대문자 6자리)."""
        bg, fg = _preview_shift_bg_fg(sk)
        return _xrgb(bg), _xrgb(fg)

    _hdr_name_bg, _hdr_name_fg = _xrgb("#ECEFF1"), _xrgb("#263238")
    num_days = _app.NUM_DAYS
    NC, OC, OHC, DC = num_days + 2, num_days + 3, num_days + 4, num_days + 5

    year_label = _app.YEAR
    month_label = _app.MONTH
    ws.merge_cells(f"A1:{get_column_letter(DC)}1")
    c = ws["A1"]; c.value = f"{year_label}년 {month_label}월 근무표"
    c.fill = PatternFill("solid", fgColor=_hdr_name_bg); c.alignment = ctr
    c.font = Font(bold=True, size=14, color=_hdr_name_fg)
    ws.row_dimensions[1].height = 28

    h = ws.cell(2, 1, "간호사")
    h.fill = PatternFill("solid", fgColor=_hdr_name_bg)
    h.font = Font(bold=True, color=_hdr_name_fg, size=10)
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

    for col, lbl, sk in [
        (NC, "N\n합계", "N"),
        (OC, "OF\n합계", "OF"),
        (OHC, "OH\n합계", "OH"),
        (DC, "D\n합계", "D"),
    ]:
        c = ws.cell(2, col, lbl); c.alignment = ctr; c.border = thin
        _bg, _fg = _px(sk)
        c.fill = PatternFill("solid", fgColor=_bg)
        c.font = Font(bold=True, color=_fg, size=9)
        ws.column_dimensions[get_column_letter(col)].width = 5.5
    ws.row_dimensions[2].height = 28

    for n_idx, name in enumerate(nurse_names):
        row = n_idx + 3
        nc = ws.cell(row, 1, name)
        nc.fill = PatternFill("solid", fgColor=_hdr_name_bg)
        nc.font = Font(bold=True, color=_hdr_name_fg, size=9)
        nc.alignment = ctr; nc.border = thin; ws.row_dimensions[row].height = 18
        ns = schedule.get(n_idx, {}); n_c = d_c = of_c = oh_c = 0
        for d, day in enumerate(days):
            shift = ns.get(d + 1, ""); col = d + 2
            cell = ws.cell(row, col, shift); cell.alignment = ctr; cell.border = thin
            bg, fg = _px(shift)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = Font(color=fg, size=9, bold=True)
            if shift == "N": n_c += 1
            elif shift == "D": d_c += 1
            elif shift in ("OF", "NO"): of_c += 1
            elif shift == "OH": oh_c += 1
        for col, val, sk in [
            (NC, n_c, "N"),
            (OC, of_c, "OF"),
            (OHC, oh_c, "OH"),
            (DC, d_c, "D"),
        ]:
            bg, fg = _px(sk)
            c = ws.cell(row, col, val); c.alignment = ctr; c.border = thin
            c.fill = PatternFill("solid", fgColor=bg); c.font = Font(color=fg, bold=True, size=10)

    sr = len(nurse_names) + 3
    for idx, (lbl, sk) in enumerate([("D 인원", "D"), ("E 인원", "E"), ("N 인원", "N")]):
        row = sr + idx; lc = ws.cell(row, 1, lbl)
        lb, lf = _px(sk)
        if sk in ("D", "E"):
            lb, lf = _xrgb(_PREVIEW_BG_DE), _xrgb(_PREVIEW_FG_BLACK)
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
_MONTH_NAMES = [
    "1월", "2월", "3월", "4월", "5월", "6월",
    "7월", "8월", "9월", "10월", "11월", "12월",
]

with st.container(border=True):
    # ── 제목 + 연·월 (한 줄·최소 높이) ───────────────────────────────────────
    _h1, _h2, _h3, _h4 = st.columns([1.15, 0.4, 0.4, 0.82], gap="small")
    with _h1:
        st.markdown(
            '<p style="margin:0;padding:0;font-size:0.88rem;font-weight:800;color:#1A237E;line-height:1.15;">'
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
            f'<p style="margin:0;padding:0;font-size:10px;font-weight:700;color:#333;line-height:1.15;">'
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

    st.markdown(
        '<hr style="margin:0.2rem 0;border:none;border-top:1px solid #e0e0e0;">',
        unsafe_allow_html=True,
    )

    dept_list = list(st.session_state.departments.keys())
    try:
        active_idx = dept_list.index(st.session_state.active_dept)
    except ValueError:
        active_idx = 0

    # 가로 1행: 부서 선택 + 부서추가 + 명단 + 공휴일 (열 간격 최소)
    _r0a, _r0b, _r0c, _r0d = st.columns([1.55, 0.72, 0.75, 0.82], gap="small")
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
            f'<p style="margin:0;font-size:9px;color:#546E7A;line-height:1.1;">📂 {active_dept} · '
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
                def _fp_all_names_ok(p):
                    ns = _fp_row_names_from_entry(p)
                    return bool(ns) and all(n in updated_nurses for n in ns)

                st.session_state.dept_forbidden_pairs[active_dept] = [
                    p for p in _fp if _fp_all_names_ok(p)
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
    _r1a, _r1b, _r1c, _r1d = st.columns([2.15, 1.95, 0.38, 1.05], gap="small")
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
                '<p class="fp-forbidden-help" style="font-size:10px;line-height:1.45;color:#616161;'
                'margin:0 0 14px 0;padding-bottom:2px;">'
                "<strong>수간호사 포함</strong> <strong>2~4명</strong>을 고릅니다. 선택한 사람들은 같은 날·같은 근무에 "
                "동시에 배치되지 않습니다. 아래에서 <strong>D / E / N</strong> 중 적용할 근무를 고릅니다.</p>",
                unsafe_allow_html=True,
            )
            _fp_list = st.session_state.dept_forbidden_pairs.setdefault(active_dept, [])
            st.markdown('<div class="fp-multiselect-anchor"></div>', unsafe_allow_html=True)
            _fp_pick = st.multiselect(
                "함께 근무 불가 인원 (2~4명)",
                nurses,
                key=f"fp_multi_{active_dept}",
                max_selections=4,
                label_visibility="collapsed",
                placeholder="간호사 선택",
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
                _nuniq = sorted(set(_fp_pick))
                if len(_nuniq) < 2:
                    st.warning("2명 이상(최대 4명) 선택해 주세요.")
                elif len(_nuniq) > 4:
                    st.warning("최대 4명까지 선택할 수 있습니다.")
                elif not _fp_shift_sel:
                    st.warning("적용할 근무(D/E/N)를 하나 이상 선택해 주세요.")
                else:
                    _gkey = tuple(_nuniq)
                    _shifts = sorted(_fp_shift_sel, key=lambda x: "DEN".index(x))
                    _found_i = None
                    for _ix, _row in enumerate(_fp_list):
                        _ex = _fp_row_names_from_entry(_row)
                        if _ex and tuple(_ex) == _gkey:
                            _found_i = _ix
                            break
                    if _found_i is not None:
                        _old = _fp_list[_found_i]
                        if isinstance(_old[0], list):
                            _prev = set(_old[1]) if len(_old) > 1 and isinstance(_old[1], list) else {"D", "E", "N"}
                        else:
                            _prev = (
                                set(_old[2]) if len(_old) >= 3 and isinstance(_old[2], list) else {"D", "E", "N"}
                            )
                        _merged = sorted(_prev | set(_shifts), key=lambda x: "DEN".index(x))
                        _fp_list[_found_i] = [list(_nuniq), _merged]
                    else:
                        _fp_list.append([list(_nuniq), _shifts])
                    _save_departments_to_disk()
                    st.rerun()
            if _fp_list:
                for _i, _pr in enumerate(list(_fp_list)):
                    _r1, _r2 = st.columns([5, 1])
                    with _r1:
                        _nm_disp = _fp_row_names_from_entry(_pr)
                        _lbl = " · ".join(_nm_disp) if _nm_disp else "?"
                        if isinstance(_pr[0], list):
                            _sh_disp = _pr[1] if len(_pr) > 1 and isinstance(_pr[1], list) else ["D", "E", "N"]
                        else:
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
                            f"🔗 {_lbl}<br/>{_tags}</div>",
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
        st.markdown("<div style='min-height:1rem'></div>", unsafe_allow_html=True)
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
        _hint_pk = _period_storage_key(st.session_state.sel_year, st.session_state.sel_month)
        _hint_sub = st.session_state.dept_schedules.get(active_dept, {})
        _has_sched = isinstance(_hint_sub, dict) and bool(_hint_sub.get(_hint_pk))
        _gen_lbl = "🗓️ 재생성" if _has_sched else "🗓️ 생성"
        if st.button(
            _gen_lbl,
            type="primary",
            use_container_width=True,
            key="btn_generate",
            help="신청으로 적은 칸은 유지하고, 자동 배정 칸만 규칙에 맞게 다시 짭니다. 재생성마다 패턴이 달라질 수 있습니다.",
        ):
            st.session_state["_pending_schedule_generate"] = True
        if _has_sched:
            st.caption("✅ 생성됨 · 재생성 가능")

    holidays = _parse_holidays(st.session_state.dept_holidays.get(active_dept, ""))


# ════════════════════════════════════════════════════════════════════════════════
#  MAIN – 변수 준비
# ════════════════════════════════════════════════════════════════════════════════
nurses      = st.session_state.departments[active_dept]   # 최신 명단
num_nurses  = len(nurses)
days        = get_april_days(holidays)
# 신청 근무 표는 짧은 열 제목(한 화면에 한 달)
req_col_labels = [_day_label_compact(d) for d in days]
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
    df_req.columns = req_col_labels

# None / nan → 공백으로 정규화
df_req = df_req.apply(lambda col: col.map(
    lambda x: "" if (x is None or str(x).strip() in ("None", "nan")) else str(x).strip()
))
_rq_sub[_period_pk] = df_req

_inject_week_split_css(days)

# ════════════════════════════════════════════════════════════════════════════════
#  MAIN – 생성된 근무표
# ════════════════════════════════════════════════════════════════════════════════
_sched_sub = st.session_state.dept_schedules.get(active_dept, {})
sched_data = _sched_sub.get(_period_pk) if isinstance(_sched_sub, dict) else None

if sched_data:
    schedule    = sched_data["schedule"]
    sched_names = sched_data["nurse_names"]
    sched_hols  = sched_data["holidays"]
    sched_days  = get_april_days(sched_hols)
    sched_n     = len(sched_names)
    sched_reqs  = sched_data.get("requests", {})

    st.markdown("<hr>", unsafe_allow_html=True)

    # ── 수정 모드 (✏️ 눌렀을 때만 편집 표 — 평소는 컬러 미리보기만)
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

    if is_edit:
        st.info("셀을 클릭하면 근무를 변경할 수 있습니다. 수정 후 **💾 저장**을 눌러 확정하세요.", icon="✏️")
        st.markdown(
            '<div class="duty-schedule-editor-hscroll" aria-hidden="true"></div>',
            unsafe_allow_html=True,
        )
        _sched_shift_options = [""] + SHIFT_NAMES
        edit_df = _schedule_to_edit_df(schedule, sched_names, sched_days)
        col_cfg = {
            lbl: st.column_config.SelectboxColumn(lbl, options=_sched_shift_options, width="small")
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
                    requests=sched_reqs or None,
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
    else:
        _show_schedule_preview_iframe(
            _render_schedule_html(schedule, sched_names, sched_days, sched_reqs),
            sched_n,
        )

# ════════════════════════════════════════════════════════════════════════════════
#  MAIN – 신청 근무 입력 달력
# ════════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="card" style="padding:10px 14px;margin-bottom:8px;">
  <div class="card-title" style="font-size:15px;margin-bottom:3px;line-height:1.2;">📝 신청 근무 입력 &nbsp;
    <span class="dept-badge" style="font-size:10px;padding:2px 8px;">{active_dept}</span>
  </div>
  <div class="card-sub" style="font-size:10px;line-height:1.3;margin:0;">
    {_app.YEAR}년 {_app.MONTH}월 · 날짜는 세로 헤더(1~말일) · 왼쪽 이름 · 클릭 선택 · 빈칸 자동 · <strong>·</strong>토일 <strong>♦</strong>공휴일
    · <strong>NO</strong>는 N 누적 20회 휴무(개인별 날짜, 자동배정 없음) → 직접 선택
  </div>
</div>
""", unsafe_allow_html=True)

# 범례 (작은 칩 형태)
legend_items = [
    ("A1","수간호사"), ("D","데이"), ("E","이브닝"), ("N","나이트"),
    ("OF","휴무"), ("OH","휴일"),
    ("NO","N 누적 20회마다 생기는 휴무. 발생일은 사람마다 다름(약 3개월에 1회 수준). 자동배정 없음·수기 입력."),
    ("연","연차"), ("병","병가"), ("공","공가"), ("경","경조"), ("EDU","교육"),
]
_leg_chips = []
for shift, tip in legend_items:
    bg, fg = _preview_shift_bg_fg(shift)
    _leg_chips.append(
        f'<span title="{tip}" style="display:inline-block;background:{bg};color:{fg};'
        f'text-align:center;padding:0 3px;margin:0 1px 1px 0;border-radius:2px;'
        f'font-size:8px;font-weight:700;line-height:1.2;">{shift}</span>'
    )
st.markdown(
    f'<div style="display:flex;flex-wrap:wrap;align-items:center;gap:0;margin:0 0 4px 0;">'
    f'{"".join(_leg_chips)}</div>',
    unsafe_allow_html=True,
)

# data_editor (행高·헤더 최소화로 한 달 컬럼 한 화면에 가깝게)
shift_options = [""] + SHIFT_NAMES
col_config = {
    lbl: st.column_config.SelectboxColumn(
        lbl, options=shift_options, width="small", required=False,
    )
    for lbl in req_col_labels
}
# 행高约 16px 목표 → 세로로 간호사 전원 한 화면에 가깝게
_req_table_h = min(16 * num_nurses + 44, 580)


def _clean_req_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(
        lambda col: col.map(
            lambda x: ""
            if (x is None or str(x).strip() in ("None", "nan"))
            else str(x).strip()
        )
    )


edited_df = st.data_editor(
    df_req,
    column_config=col_config,
    use_container_width=True,
    height=_req_table_h,
    key=editor_key,
    num_rows="fixed",
)

st.markdown(
    '<p style="margin:10px 0 2px 0;font-size:14px;font-weight:700;color:#1A237E;">👁️ 신청 근무 미리보기</p>'
    '<p style="margin:0 0 6px 0;font-size:11px;color:#546E7A;line-height:1.35;">'
    "위 편집 표와 동일한 내용이며, 생성된 근무표와 같은 색으로 표시됩니다. 빈 칸은 흰색입니다.</p>",
    unsafe_allow_html=True,
)
_show_schedule_preview_iframe(
    _render_requests_preview_html(_clean_req_df(edited_df), nurses, days),
    num_nurses,
    extra_rows=2,
)

# 저장 영역 (전체 너비 — 좁은 열에 넣으면 버튼이 안 보이는 경우가 있음)
req_saved_key = f"req_saved_{active_dept}_{_period_pk}_g{gen}"

with st.container(border=True):
    # Streamlit 알림/캡션은 테마에 따라 흰색으로 보일 수 있어 명시적으로 검정 처리
    st.markdown(
        '<div class="req-save-panel">'
        '<h4 style="margin:0 0 8px 0;font-size:1.1rem;color:#111111;font-weight:700;">💾 신청 근무 확정</h4>'
        '<p style="margin:0 0 12px 0;font-size:13px;color:#222222;line-height:1.5;">'
        '<strong>🗓️ 생성</strong>은 항상 위 표의 <strong>현재 내용</strong>을 사용합니다. '
        '<strong>저장하기</strong>는 브라우저를 닫아도 신청을 유지하려면 누르세요. '
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
            "⚠️ 아직 저장되지 않았습니다. (생성은 위 표 내용으로 가능 · 저장은 새로고침 후 유지용)</div>",
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
#  근무표 생성 처리 (수회: session 플래그 + 항상 현재 편집 표 기준)
# ════════════════════════════════════════════════════════════════════════════════
if st.session_state.pop("_pending_schedule_generate", False):
    req_df = _clean_req_df(edited_df)
    requests = _df_to_requests(req_df, days)
    _fp_idx = _fp_pairs_to_indices(
        nurses,
        st.session_state.dept_forbidden_pairs.get(active_dept, []),
    )
    _carry_raw = st.session_state.get(f"carry_txt_{active_dept}", "") or ""
    _carry_in = _parse_carry_in_text(_carry_raw, nurses)
    if _carry_in is False:
        st.error("전월 말 근무(JSON) 형식이 올바르지 않습니다. 중괄호·쉼표·따옴표를 확인해 주세요.")
    else:
        _sched_ex = st.session_state.dept_schedules.get(active_dept, {})
        _regen = isinstance(_sched_ex, dict) and bool(_sched_ex.get(_period_pk))
        if _regen:
            st.session_state["_schedule_regen_ctr"] = int(st.session_state.get("_schedule_regen_ctr", 0)) + 1
        _seed = (
            (int(st.session_state.get("_schedule_regen_ctr", 0)) * 1_000_003)
            ^ hash(_period_pk)
            ^ hash(active_dept)
        ) & 0x7FFFFFFF
        with st.spinner(
            "⏳ 근무표를 다시 짜는 중입니다… (신청 셀 유지·자동 칸만 조정)"
            if _regen
            else "⏳ 근무표를 계산하는 중입니다…"
        ):
            schedule, success, status = solve_schedule(
                num_nurses, requests, holidays,
                forbidden_pairs=_fp_idx or None,
                carry_in=_carry_in,
                regenerate=_regen,
                rng_seed=_seed if _regen else None,
                nurse_names=nurses,
            )
        if success:
            _rq_sub[_period_pk] = req_df
            st.session_state[req_saved_key] = True
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
                requests=requests,
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
