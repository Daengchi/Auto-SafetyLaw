@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set ROOT=%~dp0
if exist "%ROOT%python\python.exe" (set PY=%ROOT%python\python.exe) else (set PY=python)
set GIT=%ROOT%git-portable\cmd\git.exe

echo 코드 업데이트 확인 중...
"%GIT%" -C "%ROOT:~0,-1%" pull --quiet

echo.
echo ══════════════════════════════════════════════
echo   기존 법규준수평가표 업데이트
echo   파일 선택 창이 열립니다
echo ══════════════════════════════════════════════
echo.

cd /d "%ROOT%01. 법규 데이터 관리"
set PYTHONPATH=%ROOT%01. 법규 데이터 관리
"%PY%" main.py --update

echo.
pause
