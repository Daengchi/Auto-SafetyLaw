# setup.ps1 — Auto-SafetyLaw 설치 스크립트
# 설치.bat 에서 자동으로 실행됩니다. 직접 실행하지 마세요.

$INSTALL_DIR = "C:\SafetyLaw"
$REPO_URL    = "https://github.com/Daengchi/Auto-SafetyLaw.git"
$PY_VER      = "3.12.9"
$PY_URL      = "https://www.python.org/ftp/python/$PY_VER/python-$PY_VER-embed-amd64.zip"
$GIT_URL     = "https://github.com/git-for-windows/git/releases/download/v2.47.1.windows.2/MinGit-2.47.1.2-64-bit.zip"

$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Step { param($n, $msg) Write-Host "[$n] $msg" -ForegroundColor Cyan }
function OK   { Write-Host "    완료." -ForegroundColor Green }
function Fail {
    param($msg)
    Write-Host "    [오류] $msg" -ForegroundColor Red
    Read-Host "Enter 를 눌러 종료"
    exit 1
}

if (Test-Path $INSTALL_DIR) {
    Write-Host "[오류] $INSTALL_DIR 이(가) 이미 존재합니다." -ForegroundColor Red
    Write-Host "       폴더를 삭제하거나 INSTALL_DIR 경로를 변경한 후 다시 실행하세요." -ForegroundColor Yellow
    Read-Host "Enter 를 눌러 종료"
    exit 1
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  안전보건 법규 자동화 -- 설치 시작"
Write-Host "  설치 경로: $INSTALL_DIR"
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""

# ── 1/5: Portable Git ─────────────────────────────────────────────────────────
Step "1/5" "Portable Git 다운로드 중..."
$TEMPGIT = "$env:TEMP\mingit_setup"
if (Test-Path $TEMPGIT) { Remove-Item $TEMPGIT -Recurse -Force }
try {
    Invoke-WebRequest $GIT_URL -OutFile "$env:TEMP\mingit.zip" -UseBasicParsing
    Expand-Archive "$env:TEMP\mingit.zip" -DestinationPath $TEMPGIT -Force
    Remove-Item "$env:TEMP\mingit.zip"
} catch { Fail "Git 다운로드 실패: $_" }
OK

# ── 2/5: 저장소 클론 ───────────────────────────────────────────────────────────
Step "2/5" "저장소 클론 중..."
& "$TEMPGIT\cmd\git.exe" clone $REPO_URL $INSTALL_DIR
if ($LASTEXITCODE -ne 0) { Fail "클론 실패. 네트워크 연결 및 방화벽을 확인하세요." }
OK

# ── 3/5: Portable Git 이동 ────────────────────────────────────────────────────
Step "3/5" "Portable Git 배치 중..."
Move-Item $TEMPGIT "$INSTALL_DIR\git-portable"
OK

# ── 4/5: Python 환경 구성 ─────────────────────────────────────────────────────
Step "4/5" "Python 환경 구성 중 (수 분 소요)..."
$PYDIR = "$INSTALL_DIR\python"
New-Item -ItemType Directory -Path $PYDIR -Force | Out-Null

try {
    Invoke-WebRequest $PY_URL -OutFile "$env:TEMP\py-embed.zip" -UseBasicParsing
    Expand-Archive "$env:TEMP\py-embed.zip" -DestinationPath $PYDIR -Force
    Remove-Item "$env:TEMP\py-embed.zip"
} catch { Fail "Python 다운로드 실패: $_" }

if (-not (Test-Path "$PYDIR\python.exe")) { Fail "Python 압축 해제 실패." }

# site.py 활성화 (pip 필수)
$pth = (Get-ChildItem $PYDIR -Filter "python3*._pth").FullName
(Get-Content $pth) -replace '#import site', 'import site' | Set-Content $pth

# pip 설치
Invoke-WebRequest "https://bootstrap.pypa.io/get-pip.py" -OutFile "$INSTALL_DIR\get-pip.py" -UseBasicParsing
& "$PYDIR\python.exe" "$INSTALL_DIR\get-pip.py" --quiet
if ($LASTEXITCODE -ne 0) { Fail "pip 설치 실패." }
Remove-Item "$INSTALL_DIR\get-pip.py"

# 패키지 설치
& "$PYDIR\python.exe" -m pip install -q -r "$INSTALL_DIR\01. 법규 데이터 관리\requirements.txt"
& "$PYDIR\python.exe" -m pip install -q -r "$INSTALL_DIR\02. 제개정사항 모니터링\requirements.txt"
& "$PYDIR\python.exe" -m pip install -q -r "$INSTALL_DIR\03. 법규 제개정 이메일 알림\requirements.txt"
OK

# ── 5/5: 설정 파일 초기화 ─────────────────────────────────────────────────────
Step "5/5" "설정 파일 생성 중..."

if (-not (Test-Path "$INSTALL_DIR\laws.json")) {
    Copy-Item "$INSTALL_DIR\laws.json.example" "$INSTALL_DIR\laws.json"
}
foreach ($proj in @("01. 법규 데이터 관리", "02. 제개정사항 모니터링", "03. 법규 제개정 이메일 알림")) {
    if (-not (Test-Path "$INSTALL_DIR\$proj\.env")) {
        Copy-Item "$INSTALL_DIR\$proj\.env.example" "$INSTALL_DIR\$proj\.env"
    }
}
if (-not (Test-Path "$INSTALL_DIR\03. 법규 제개정 이메일 알림\recipients.json")) {
    Copy-Item "$INSTALL_DIR\03. 법규 제개정 이메일 알림\recipients.json.example" `
              "$INSTALL_DIR\03. 법규 제개정 이메일 알림\recipients.json"
}

foreach ($dir in @("logs", "01. 법규 데이터 관리\output", "02. 제개정사항 모니터링\data", "02. 제개정사항 모니터링\output")) {
    New-Item -ItemType Directory -Path "$INSTALL_DIR\$dir" -Force | Out-Null
}
OK

# ── 완료 안내 ──────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "======================================================" -ForegroundColor Green
Write-Host "  설치 완료! 아래 파일을 설정한 후 사용하세요." -ForegroundColor Green
Write-Host ""
Write-Host "  [1] API 키  (https://www.law.go.kr 에서 무료 발급)"
Write-Host "      $INSTALL_DIR\01. 법규 데이터 관리\.env"
Write-Host "      $INSTALL_DIR\02. 제개정사항 모니터링\.env"
Write-Host ""
Write-Host "  [2] 이메일 / SMTP 설정"
Write-Host "      $INSTALL_DIR\03. 법규 제개정 이메일 알림\.env"
Write-Host ""
Write-Host "  [3] 수신자 이메일 목록"
Write-Host "      $INSTALL_DIR\03. 법규 제개정 이메일 알림\recipients.json"
Write-Host ""
Write-Host "  [4] 관리 법령 목록  (불필요한 항목 삭제)"
Write-Host "      $INSTALL_DIR\laws.json"
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""

$reg = Read-Host "작업 스케줄러 등록할까요? 평일 09:00 자동 실행 [Y/N]"
if ($reg -ieq 'Y') {
    try {
        Register-ScheduledTask `
            -TaskName "SafetyLaw-Monitor" `
            -Trigger  (New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At '09:00') `
            -Action   (New-ScheduledTaskAction -Execute "$INSTALL_DIR\법규모니터링_자동실행.bat") `
            -RunLevel Highest -Force | Out-Null
        Write-Host "  작업 스케줄러 등록 완료." -ForegroundColor Green
    } catch {
        Write-Host "  [주의] 등록 실패. 관리자 권한으로 다시 실행하세요." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "  설정 완료 후 01_신규생성.bat 을 실행하세요." -ForegroundColor Cyan
Read-Host "Enter 를 눌러 종료"
