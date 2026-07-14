@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set APP=%~dp0
pushd "%~dp0.."
set ROOT=%CD%
popd
if exist "%APP%python\python.exe" (set PY=%APP%python\python.exe) else if exist "%ROOT%\python\python.exe" (set PY=%ROOT%\python\python.exe) else (set PY=python)
if exist "%APP%git-portable\cmd\git.exe" (set GIT=%APP%git-portable\cmd\git.exe) else (set GIT=%ROOT%\git-portable\cmd\git.exe)

if not exist "%APP%logs" mkdir "%APP%logs"
set LOG=%APP%logs\실행로그.log

echo ============================================ >> "%LOG%"
echo [%date% %time%] 실행 시작 >> "%LOG%"

:: 코드 자동 업데이트 (GitHub 최신 버전 반영)
"%GIT%" config --global --add safe.directory "%ROOT:\=/%" >> "%LOG%" 2>&1
"%GIT%" -C "%ROOT%" pull --quiet >> "%LOG%" 2>&1

:: 02. 제개정사항 모니터링
cd /d "%ROOT%\02. 제개정사항 모니터링"
set PYTHONPATH=%CD%
"%PY%" main.py --no-report >> "%LOG%" 2>&1

:: 03. 법규 제개정 이메일 알림
cd /d "%ROOT%\03. 법규 제개정 이메일 알림"
set PYTHONPATH=%CD%
"%PY%" main.py >> "%LOG%" 2>&1

echo [%date% %time%] 실행 종료 >> "%LOG%"
