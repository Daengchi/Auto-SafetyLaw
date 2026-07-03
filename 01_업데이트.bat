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
echo   기존 법규준수평가표 업데이트
echo   파일 선택 창이 열립니다
echo ══════════════════════════════════════════════
echo.

cd /d "%ROOT%01. 법규 데이터 관리"
set PYTHONPATH=%CD%
powershell -NoProfile -ExecutionPolicy Bypass -Sta -File "%CD%\select_and_update.ps1" -Py "%PY%"

echo.
pause
