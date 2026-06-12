# setup.ps1 - Auto-SafetyLaw installer
# Run via install.bat - do not run directly.

$INSTALL_DIR = "C:\SafetyLaw"
$REPO_URL    = "https://github.com/Daengchi/Auto-SafetyLaw.git"
$PY_VER      = "3.12.9"
$PY_URL      = "https://www.python.org/ftp/python/$PY_VER/python-$PY_VER-embed-amd64.zip"
$GIT_URL     = "https://github.com/git-for-windows/git/releases/download/v2.47.1.windows.2/MinGit-2.47.1.2-64-bit.zip"

function Step { param($n, $msg) Write-Host "[$n] $msg" -ForegroundColor Cyan }
function OK   { Write-Host "    Done." -ForegroundColor Green }
function Fail {
    param($msg)
    Write-Host "[ERROR] $msg" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

if (Test-Path $INSTALL_DIR) {
    Write-Host "[ERROR] $INSTALL_DIR already exists." -ForegroundColor Red
    Write-Host "        Delete the folder and run again." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "======================================================"
Write-Host "  Auto-SafetyLaw - Installation"
Write-Host "  Target: $INSTALL_DIR"
Write-Host "======================================================"
Write-Host ""

# 1/5: Portable Git
Step "1/5" "Downloading Portable Git..."
$TEMPGIT = "$env:TEMP\mingit_setup"
if (Test-Path $TEMPGIT) { Remove-Item $TEMPGIT -Recurse -Force }
try {
    Invoke-WebRequest $GIT_URL -OutFile "$env:TEMP\mingit.zip" -UseBasicParsing
    Expand-Archive "$env:TEMP\mingit.zip" -DestinationPath $TEMPGIT -Force
    Remove-Item "$env:TEMP\mingit.zip"
} catch { Fail "Git download failed: $_" }
OK

# 2/5: Clone repository
Step "2/5" "Cloning repository..."
& "$TEMPGIT\cmd\git.exe" clone $REPO_URL $INSTALL_DIR
if ($LASTEXITCODE -ne 0) { Fail "Clone failed. Check network or firewall." }
OK

# 3/5: Move Portable Git
Step "3/5" "Moving Portable Git..."
Move-Item $TEMPGIT "$INSTALL_DIR\git-portable"
$safeDir = $INSTALL_DIR.Replace('\', '/')
& "$INSTALL_DIR\git-portable\cmd\git.exe" config --system safe.directory $safeDir
OK

# 4/5: Python setup
Step "4/5" "Setting up Python (may take a few minutes)..."
$PYDIR = "$INSTALL_DIR\python"
New-Item -ItemType Directory -Path $PYDIR -Force | Out-Null
try {
    Invoke-WebRequest $PY_URL -OutFile "$env:TEMP\py-embed.zip" -UseBasicParsing
    Expand-Archive "$env:TEMP\py-embed.zip" -DestinationPath $PYDIR -Force
    Remove-Item "$env:TEMP\py-embed.zip"
} catch { Fail "Python download failed: $_" }

if (-not (Test-Path "$PYDIR\python.exe")) { Fail "Python extraction failed." }

$pth = (Get-ChildItem $PYDIR -Filter "python3*._pth").FullName
(Get-Content $pth) -replace '#import site', 'import site' | Set-Content $pth

Invoke-WebRequest "https://bootstrap.pypa.io/get-pip.py" -OutFile "$INSTALL_DIR\get-pip.py" -UseBasicParsing
& "$PYDIR\python.exe" "$INSTALL_DIR\get-pip.py" --quiet
if ($LASTEXITCODE -ne 0) { Fail "pip installation failed." }
Remove-Item "$INSTALL_DIR\get-pip.py"

$reqFiles = Get-ChildItem $INSTALL_DIR -Recurse -Filter "requirements.txt"
foreach ($req in $reqFiles) {
    & "$PYDIR\python.exe" -m pip install -q -r $req.FullName
}
OK

# 5/5: Config files
Step "5/5" "Creating config files..."

if (-not (Test-Path "$INSTALL_DIR\laws.json")) {
    $ex = Get-Item "$INSTALL_DIR\laws.json.example" -ErrorAction SilentlyContinue
    if ($ex) { Copy-Item $ex.FullName "$INSTALL_DIR\laws.json" }
}

$projFolders = Get-ChildItem $INSTALL_DIR -Directory | Where-Object { $_.Name -match "^\d+\." }
foreach ($proj in $projFolders) {
    $envFile = Join-Path $proj.FullName ".env"
    $envEx   = Join-Path $proj.FullName ".env.example"
    if (-not (Test-Path $envFile) -and (Test-Path $envEx)) {
        Copy-Item $envEx $envFile
    }
    $rec   = Join-Path $proj.FullName "recipients.json"
    $recEx = Join-Path $proj.FullName "recipients.json.example"
    if (-not (Test-Path $rec) -and (Test-Path $recEx)) {
        Copy-Item $recEx $rec
    }
    foreach ($sub in @("output", "data")) {
        New-Item -ItemType Directory -Path (Join-Path $proj.FullName $sub) -Force | Out-Null
    }
}

New-Item -ItemType Directory -Path "$INSTALL_DIR\logs" -Force | Out-Null
OK

# Done
Write-Host ""
Write-Host "======================================================"
Write-Host "  Installation complete!" -ForegroundColor Green
Write-Host ""
Write-Host "  Configure these files before running:"
Write-Host "  [1] API key  (get free key at https://www.law.go.kr)"
$projFolders | Where-Object { $_.Name -match "01\.|02\." } | ForEach-Object {
    Write-Host "      $($_.FullName)\.env"
}
Write-Host "  [2] Email/SMTP"
$projFolders | Where-Object { $_.Name -match "03\." } | ForEach-Object {
    Write-Host "      $($_.FullName)\.env"
    Write-Host "      $($_.FullName)\recipients.json"
}
Write-Host "  [3] Law list (remove unneeded laws)"
Write-Host "      $INSTALL_DIR\laws.json"
Write-Host "======================================================"
Write-Host ""

$reg = Read-Host "Register Task Scheduler? (Weekdays 09:00 auto-run) [Y/N]"
if ($reg -ieq 'Y') {
    $bat = Get-ChildItem $INSTALL_DIR -Filter "*자동실행*" | Select-Object -First 1
    if ($bat) {
        try {
            Register-ScheduledTask `
                -TaskName "SafetyLaw-Monitor" `
                -Trigger  (New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At '09:00') `
                -Action   (New-ScheduledTaskAction -Execute $bat.FullName) `
                -RunLevel Highest -Force | Out-Null
            Write-Host "  Task Scheduler registered." -ForegroundColor Green
        } catch {
            Write-Host "  [WARNING] Registration failed. Run as Administrator." -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "  Run 01_신규생성.bat to generate your first report." -ForegroundColor Cyan
Read-Host "Press Enter to exit"
