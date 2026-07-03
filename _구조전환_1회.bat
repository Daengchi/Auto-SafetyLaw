@echo off
chcp 65001 >nul
setlocal
set ROOT=%~dp0
if exist "%ROOT%git-portable\cmd\git.exe" (set GIT=%ROOT%git-portable\cmd\git.exe) else (set GIT=%ROOT%프로그램\git-portable\cmd\git.exe)

echo ══════════════════════════════════════════════
echo   구조 전환 (기존 사업장 1회 실행)
echo   최신 코드를 받고 python / git-portable / logs 를
echo   프로그램\ 폴더로 정리한 뒤 예약작업을 재등록합니다.
echo   설정(.env, laws.json, 데이터)은 그대로 보존됩니다.
echo ══════════════════════════════════════════════
echo.

echo [1/3] 최신 코드 받는 중...
"%GIT%" -C "%ROOT:~0,-1%" pull --quiet

echo [2/3] 런타임 폴더 정리 중...
if not exist "%ROOT%프로그램" mkdir "%ROOT%프로그램"
for %%D in (python git-portable logs) do (
  if exist "%ROOT%%%D\" if not exist "%ROOT%프로그램\%%D\" move "%ROOT%%%D" "%ROOT%프로그램\" >nul
)
if exist "%ROOT%amendments.json" if not exist "%ROOT%프로그램\amendments.json" move "%ROOT%amendments.json" "%ROOT%프로그램\" >nul

echo [3/3] 예약작업 재등록 중... (실패 시 이 파일을 관리자 권한으로 다시 실행)
if exist "%ROOT%프로그램\작업등록.ps1" powershell -NoProfile -ExecutionPolicy Bypass -File "%ROOT%프로그램\작업등록.ps1"

echo.
echo 완료! 이제 이 파일(_구조전환_1회.bat)은 삭제하셔도 됩니다.
pause
endlocal
