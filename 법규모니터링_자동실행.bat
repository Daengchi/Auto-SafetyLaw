@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set ROOT=%~dp0
if exist "%ROOT%python\python.exe" (set PY=%ROOT%python\python.exe) else (set PY=python)
set GIT=%ROOT%git-portable\cmd\git.exe

if not exist "%ROOT%logs" mkdir "%ROOT%logs"
set LOG=%ROOT%logs\실행로그.log

echo ============================================ >> "%LOG%"
echo [%date% %time%] 실행 시작 >> "%LOG%"

:: 코드 자동 업데이트 (GitHub 최신 버전 반영)
"%GIT%" -C "%ROOT%" pull --quiet >> "%LOG%" 2>&1

:: 02. 제개정사항 모니터링
cd /d "%ROOT%02. 제개정사항 모니터링"
"%PY%" main.py --no-report >> "%LOG%" 2>&1

:: 03. 법규 제개정 이메일 알림
cd /d "%ROOT%03. 법규 제개정 이메일 알림"
"%PY%" main.py >> "%LOG%" 2>&1

echo [%date% %time%] 실행 종료 >> "%LOG%"
