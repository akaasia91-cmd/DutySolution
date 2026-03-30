@echo off
chcp 65001 > nul
echo =====================================================
echo   응급실 2026년 4월 근무표 생성기 (Streamlit)
echo =====================================================
echo.

REM 파이썬이 설치되어 있는지 확인
python --version > nul 2>&1
IF ERRORLEVEL 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo Python 공식 사이트(https://www.python.org)에서 설치해주세요.
    pause
    exit /b
)

REM 필요한 패키지 설치
echo 필요한 패키지를 설치합니다 (처음 한 번만 필요합니다)...
pip install -q -r requirements.txt

echo.
echo 프로그램을 시작합니다...
echo 잠시 후 브라우저가 자동으로 열립니다.
echo (열리지 않으면 브라우저에서 http://localhost:8501 을 입력하세요)
echo.
echo 종료하려면 이 창을 닫거나 Ctrl+C 를 누르세요.
echo =====================================================

REM Streamlit 앱 실행 (자동으로 브라우저 열림)
REM PATH에 streamlit이 없어도 동작하도록 python -m 사용
python -m streamlit run streamlit_app.py --server.port 8501

pause
