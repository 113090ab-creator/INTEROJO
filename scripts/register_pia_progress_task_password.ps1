[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$taskName = "PIA 정기오더 진도현황 자동생성"
$repoRoot = "C:\Users\유현아\Documents\GitHub\INTEROJO"
$scriptPath = Join-Path $repoRoot "scripts\run_pia_progress_daily.ps1"
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

if (-not (Test-Path -LiteralPath $scriptPath)) {
  throw "자동 실행 스크립트를 찾지 못했습니다: $scriptPath"
}

Write-Host ""
Write-Host "작업 스케줄러를 '로그온 여부와 관계없이 실행' 모드로 등록합니다."
Write-Host "Windows 계정 암호는 채팅이나 파일에 저장하지 않고, Windows 작업 스케줄러에만 전달됩니다."
Write-Host "사용자 계정: $identity"
Write-Host ""

Write-Host "아래 입력란에는 Windows 로그인 암호를 입력하세요. PIN/지문/얼굴인식 암호가 아니라 실제 계정 암호입니다."
$securePassword = Read-Host -AsSecureString "PIA 자동화 실행용 Windows 계정 암호"
$credential = [pscredential]::new($identity, $securePassword)

$passwordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
try {
  $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)
}
finally {
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)
}

if ([string]::IsNullOrWhiteSpace($plainPassword)) {
  throw "암호가 비어 있어 등록을 중단했습니다."
}

$action = New-ScheduledTaskAction `
  -Execute "powershell.exe" `
  -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -SkipIfExists" `
  -WorkingDirectory $repoRoot

$trigger = New-ScheduledTaskTrigger -Daily -At 8:50am

$settings = New-ScheduledTaskSettingsSet `
  -WakeToRun `
  -StartWhenAvailable `
  -AllowStartIfOnBatteries `
  -DontStopIfGoingOnBatteries `
  -MultipleInstances IgnoreNew `
  -ExecutionTimeLimit (New-TimeSpan -Hours 2)

$principal = New-ScheduledTaskPrincipal `
  -UserId $credential.UserName `
  -LogonType Password `
  -RunLevel Limited

$task = New-ScheduledTask `
  -Action $action `
  -Trigger $trigger `
  -Settings $settings `
  -Principal $principal

Register-ScheduledTask `
  -TaskName $taskName `
  -InputObject $task `
  -User $credential.UserName `
  -Password $plainPassword `
  -Force | Out-Null

Write-Host ""
Write-Host "등록 완료. 스케줄러 경로로 테스트 실행합니다."
Start-ScheduledTask -TaskName $taskName

do {
  Start-Sleep -Seconds 2
  $state = (Get-ScheduledTask -TaskName $taskName).State
} while ($state -eq "Running")

$info = Get-ScheduledTaskInfo -TaskName $taskName
$principalInfo = (Get-ScheduledTask -TaskName $taskName).Principal

Write-Host ""
Write-Host "작업 이름: $taskName"
Write-Host "로그온 유형: $($principalInfo.LogonType)"
Write-Host "마지막 실행 결과: $($info.LastTaskResult)"
Write-Host "다음 실행 시간: $($info.NextRunTime)"

if ($info.LastTaskResult -ne 0) {
  throw "테스트 실행이 실패했습니다. LastTaskResult=$($info.LastTaskResult)"
}

Write-Host ""
Write-Host "완료: 이제 PC가 켜져 있고 Windows가 로그인 화면까지 올라온 상태라면 로그인 전에도 실행됩니다."
