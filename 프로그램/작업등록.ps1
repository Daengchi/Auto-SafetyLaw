# 법규 제개정 점검·이메일 알림 — Windows 작업 스케줄러 등록
# PowerShell에서 1회 실행하면 평일(월~금) 09시 자동 실행 작업이 등록된다.

$bat = Join-Path $PSScriptRoot "법규모니터링_자동실행.bat"

$action  = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$bat`""
$trigger = New-ScheduledTaskTrigger -Weekly `
             -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 9:00am
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
             -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask -TaskName "법규제개정_이메일알림" `
  -Action $action -Trigger $trigger -Settings $settings `
  -Description "평일 09시 법규 개정 점검 후 변동 시 이메일 발송" -Force

Write-Host "등록 완료: 법규제개정_이메일알림 (평일 09:00)"
