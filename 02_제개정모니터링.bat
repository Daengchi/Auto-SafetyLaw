@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set ROOT=%~dp0
if exist "%ROOT%프로그램\python\python.exe" (set PY=%ROOT%프로그램\python\python.exe) else if exist "%ROOT%python\python.exe" (set PY=%ROOT%python\python.exe) else (set PY=python)
if exist "%ROOT%프로그램\git-portable\cmd\git.exe" (set GIT=%ROOT%프로그램\git-portable\cmd\git.exe) else (set GIT=%ROOT%git-portable\cmd\git.exe)

echo 코드 업데이트 확인 중...
set SAFE_DIR=%ROOT:~0,-1%
"%GIT%" config --global --add safe.directory "%SAFE_DIR:\=/%"
"%GIT%" -C "%ROOT:~0,-1%" pull --quiet

echo.
echo ══════════════════════════════════════════════
echo   법규 제개정사항 모니터링
echo   laws.json 법령의 제·개정 변경 확인 후 리포트 생성
echo ══════════════════════════════════════════════
echo.

cd /d "%ROOT%02. 제개정사항 모니터링"
set PYTHONPATH=%CD%
"%PY%" main.py --output output

echo.
pause
