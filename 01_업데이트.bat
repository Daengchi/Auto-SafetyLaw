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
echo   기존 법규준수평가표 업데이트
echo   파일 선택 창이 열립니다
echo ══════════════════════════════════════════════
echo.

cd /d "%ROOT%01. 법규 데이터 관리"
set PYTHONPATH=%CD%
powershell -NoProfile -STA -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Title = '업데이트할 법규준수평가표를 선택하세요'; $d.Filter = 'Excel 파일 (*.xlsm;*.xlsx)|*.xlsm;*.xlsx|모든 파일 (*.*)|*.*'; if ($d.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host '파일을 선택하지 않았습니다.'; exit 0 }; & '%PY%' main.py --update --file $d.FileName; exit $LASTEXITCODE"

echo.
pause
