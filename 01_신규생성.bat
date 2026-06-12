@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set ROOT=%~dp0
if exist "%ROOT%python\python.exe" (set PY=%ROOT%python\python.exe) else (set PY=python)
set GIT=%ROOT%git-portable\cmd\git.exe

echo 코드 업데이트 확인 중...
set SAFE_DIR=%ROOT:~0,-1%
"%GIT%" config --global --add safe.directory "%SAFE_DIR:\=/%"
"%GIT%" -C "%ROOT:~0,-1%" pull --quiet

echo.
echo ══════════════════════════════════════════════
echo   법규준수평가표 신규 생성
echo   laws.json 에 정의된 법령 전체 처리
echo ══════════════════════════════════════════════
echo.

cd /d "%ROOT%01. 법규 데이터 관리"
set PYTHONPATH=%CD%
"%PY%" main.py --output output

echo.
pause
