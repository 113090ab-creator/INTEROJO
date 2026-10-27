[CmdletBinding()]
param([string]$PythonPath = (Get-Command python.exe -ErrorAction Stop).Source)

$ErrorActionPreference = 'Stop'
$taskName = 'INTEROJO APS Snapshot Dispatch'
$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $PSScriptRoot 'dispatch_snapshot_refresh.py'
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
if ([TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).TotalHours -ne 9) {
  throw 'This schedule expects the Windows timezone to be Korea Standard Time (UTC+09:00).'
}

& $PythonPath $scriptPath --check
if ($LASTEXITCODE -ne 0) { throw 'GitHub access check failed; task was not changed.' }

$triggers = @()
foreach ($hour in @(8, 9, 16, 17)) {
  foreach ($minute in @(0, 10, 20, 30, 40, 50)) {
    if (($hour -eq 8 -and $minute -lt 20) -or
        ($hour -eq 9 -and $minute -gt 20) -or
        ($hour -eq 16 -and $minute -lt 30) -or
        ($hour -eq 17 -and $minute -gt 30)) { continue }
    $at = (Get-Date).Date.AddHours($hour).AddMinutes($minute)
    $triggers += New-ScheduledTaskTrigger -Daily -At $at
  }
}
$backgroundPython = Join-Path (Split-Path -Parent $PythonPath) 'pythonw.exe'
if (-not (Test-Path -LiteralPath $backgroundPython)) { throw 'pythonw.exe is required for hidden scheduled execution.' }
$action = New-ScheduledTaskAction -Execute $backgroundPython -Argument "`"$scriptPath`"" -WorkingDirectory $repoRoot
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -AllowStartIfOnBatteries `
  -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
  -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)
$principal = New-ScheduledTaskPrincipal -UserId $identity -LogonType Interactive -RunLevel Limited
$task = New-ScheduledTask -Action $action -Trigger $triggers -Settings $settings -Principal $principal `
  -Description 'Start the official GitHub APS snapshot workflow from this PC. Requires Windows login and network; locked screen is supported.'

$backupDir = Join-Path $repoRoot 'outputs\snapshot_scheduler'
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
foreach ($oldName in @($taskName, 'INTEROJO APS Snapshot Refresh')) {
  $existing = Get-ScheduledTask -TaskName $oldName -ErrorAction SilentlyContinue
  if ($existing) {
    Export-ScheduledTask -TaskName $oldName | Set-Content -LiteralPath (Join-Path $backupDir "$oldName-$(Get-Date -Format yyyyMMdd_HHmmss).xml") -Encoding Unicode
  }
}
Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null
$legacy = Get-ScheduledTask -TaskName 'INTEROJO APS Snapshot Refresh' -ErrorAction SilentlyContinue
if ($legacy) { Disable-ScheduledTask -TaskName $legacy.TaskName | Out-Null }

Write-Output "Registered: $taskName ($($triggers.Count) daily triggers)"
Write-Output 'KST 08:20-09:20 and 16:30-17:30, every 10 minutes. Keep this PC on and signed in.'
Get-ScheduledTaskInfo -TaskName $taskName | Select-Object LastRunTime,LastTaskResult,NextRunTime
