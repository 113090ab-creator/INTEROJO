[CmdletBinding()]
param(
  [string]$RunDate = (Get-Date -Format "yyyy-MM-dd"),
  [switch]$SkipIfExists,
  [string]$Packing1DayPath = "",
  [string]$PackingFrpPath = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$runDateCompact = $RunDate -replace "-", ""
$outputDir = Join-Path $repoRoot "outputs\pia_progress_$runDateCompact"
$outputFileName = "PIA 정기오더 진행 현황_$runDateCompact.xlsx"
$outputPath = Join-Path $outputDir $outputFileName
$summaryPath = Join-Path $outputDir "PIA 정기오더 진행 현황_$runDateCompact.summary.txt"
$logDir = Join-Path $repoRoot "outputs\pia_progress_logs"
$logPath = Join-Path $logDir "pia_progress_${runDateCompact}_$(Get-Date -Format 'HHmmss').log"

$python = "C:\Users\유현아\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$node = "C:\Users\유현아\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
$nodeModules = "C:\Users\유현아\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules"
$generateScript = Join-Path $repoRoot "outputs\pia_progress_20260819\generate_progress_data.py"
$buildScript = Join-Path $repoRoot "outputs\pia_progress_20260819\build_progress.mjs"
$downloads = "C:\Users\유현아\Downloads"

function Resolve-LatestPackingFile {
  param([Parameter(Mandatory = $true)][string]$Pattern)
  $file = Get-ChildItem -LiteralPath $downloads -File -Filter $Pattern |
    Where-Object { $_.Name -notlike "~$*" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if (-not $file) {
    throw "포장 파일을 찾지 못했습니다: $Pattern"
  }
  return $file.FullName
}

function Format-Count {
  param([object]$Value)
  return ([int64][math]::Round([double]$Value)).ToString("N0")
}

function Format-DeltaText {
  param([object]$Value, [string]$Suffix = "")
  $number = [int64][math]::Round([double]$Value)
  if ($number -gt 0) {
    return "``$(Format-Count $number)$Suffix`` 감소"
  }
  if ($number -lt 0) {
    return "``$(Format-Count (-1 * $number))$Suffix`` 증가"
  }
  return "``0$Suffix`` 정체"
}

function Assert-Path {
  param([Parameter(Mandatory = $true)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    throw "필수 파일이 없습니다: $Path"
  }
}

New-Item -ItemType Directory -Force -Path $outputDir, $logDir | Out-Null
Start-Transcript -LiteralPath $logPath -Append | Out-Null

try {
  if (-not $Packing1DayPath) {
    $Packing1DayPath = Resolve-LatestPackingFile "1Day_2026Y,해외수주,포장,출고관리_*.xlsx"
  }
  if (-not $PackingFrpPath) {
    $PackingFrpPath = Resolve-LatestPackingFile "FRP_2026Y,해외수주,포장,출고관리_*.xlsx"
  }

  Assert-Path $python
  Assert-Path $node
  Assert-Path $generateScript
  Assert-Path $buildScript
  Assert-Path $Packing1DayPath
  Assert-Path $PackingFrpPath

  Write-Host "PIA progress automation started: $RunDate"
  Write-Host "1-Day packing: $Packing1DayPath"
  Write-Host "FRP packing: $PackingFrpPath"
  Write-Host "Output: $outputPath"

  if ($SkipIfExists -and (Test-Path -LiteralPath $outputPath)) {
    Write-Host "Existing output found. Skipping data/API refresh and verifying only."
  } else {
    $env:PIA_PROGRESS_DATE = $RunDate
    $env:PIA_PROGRESS_OUTPUT_DIR = $outputDir
    $env:PIA_PACKING_1DAY = $Packing1DayPath
    $env:PIA_PACKING_FRP = $PackingFrpPath
    $env:NODE_PATH = $nodeModules

    & $python $generateScript
    if ($LASTEXITCODE -ne 0) {
      throw "generate_progress_data.py failed with exit code $LASTEXITCODE"
    }

    $scriptNodeModules = Join-Path (Split-Path -Parent $buildScript) "node_modules"
    if (-not (Test-Path -LiteralPath $scriptNodeModules)) {
      New-Item -ItemType Junction -Path $scriptNodeModules -Target $nodeModules | Out-Null
    }

    $markScript = Get-ChildItem -LiteralPath "C:\Users\유현아\.codex\plugins\cache\openai-primary-runtime\spreadsheets" -Recurse -File -Filter "mark_artifact_operation_started.mjs" -ErrorAction SilentlyContinue |
      Sort-Object FullName -Descending |
      Select-Object -First 1
    if ($markScript) {
      & $node $markScript.FullName --operation-kind create --expected-output-count 1 --output-format xlsx
      if ($LASTEXITCODE -ne 0) {
        Write-Warning "mark_artifact_operation_started.mjs failed with exit code $LASTEXITCODE"
      }
    }

    $env:PIA_PROGRESS_OUTPUT_FILE = $outputFileName
    & $node $buildScript
    $buildExitCode = $LASTEXITCODE
    if ($buildExitCode -ne 0 -and -not (Test-Path -LiteralPath $outputPath)) {
      throw "build_progress.mjs failed with exit code $buildExitCode and output was not created"
    }
    if ($buildExitCode -ne 0) {
      Write-Warning "build_progress.mjs returned exit code $buildExitCode after saving the workbook"
    }
  }

  $verifyCode = @'
import json
import sys
from pathlib import Path

import openpyxl

path = Path(sys.argv[1])
if not path.exists():
    raise SystemExit(f"missing output: {path}")

workbook = openpyxl.load_workbook(path, data_only=True)
required = [
    "전체현황",
    "일별추세",
    "이니셜별추세",
    "진도현황",
    "요약",
    "사출현황",
    "사출일별진도",
    "유효생산관리",
    "검증",
    "포장상세",
]
missing = [sheet for sheet in required if sheet not in workbook.sheetnames]
if missing:
    raise SystemExit(f"missing sheets: {missing}")

canada_count = 0
error_strings = []
error_prefixes = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")
for sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]
    for row in sheet.iter_rows(values_only=True):
        for value in row:
            if isinstance(value, str):
                if "캐나다향" in value:
                    canada_count += 1
                if value.startswith(error_prefixes):
                    error_strings.append((sheet_name, value))
if canada_count:
    raise SystemExit(f"캐나다향 포함 셀: {canada_count}")
if error_strings:
    raise SystemExit(f"error strings: {error_strings[:5]}")

overview = workbook["전체현황"]
data_path = path.with_name("progress_data.json")
data = json.loads(data_path.read_text(encoding="utf-8"))
visible_rows = [
    row for row in data.get("rows", [])
    if "캐나다향" not in str(row.get("브랜드명", ""))
]
packing_shortage = int(round(sum(float(row.get("포장부족수량") or 0) for row in visible_rows)))

metrics = {
    "output": str(path),
    "order_qty": int(round(float(overview.cell(8, 3).value or 0))),
    "production_shortage": int(round(float(overview.cell(8, 4).value or 0))),
    "production_rate": float(overview.cell(8, 6).value or 0),
    "production_rate_text": f"{float(overview.cell(8, 6).value or 0) * 100:.2f}%",
    "production_decrease": int(round(float(overview.cell(8, 7).value or 0))),
    "production_color_change": int(round(float(overview.cell(8, 8).value or 0))),
    "injection_shortage": int(round(float(overview.cell(8, 9).value or 0))),
    "injection_decrease": int(round(float(overview.cell(8, 11).value or 0))),
    "injection_color_change": int(round(float(overview.cell(8, 12).value or 0))),
    "packing_rate": float(overview.cell(8, 13).value or 0),
    "packing_rate_text": f"{float(overview.cell(8, 13).value or 0) * 100:.2f}%",
    "packing_shortage": packing_shortage,
    "packing_1day": data.get("sources", {}).get("packing_1day", ""),
    "packing_frp": data.get("sources", {}).get("packing_frp", ""),
}
Path(sys.argv[2]).write_text(json.dumps(metrics, ensure_ascii=False), encoding="utf-8")
'@

  $verifyScriptPath = Join-Path $logDir "verify_pia_progress_$runDateCompact.py"
  $metricsPath = Join-Path $logDir "metrics_pia_progress_$runDateCompact.json"
  $verifyCode | Set-Content -LiteralPath $verifyScriptPath -Encoding UTF8
  $verifyOutput = & $python $verifyScriptPath $outputPath $metricsPath
  if ($LASTEXITCODE -ne 0) {
    throw "verification failed with exit code $LASTEXITCODE"
  }

  if (-not (Test-Path -LiteralPath $metricsPath)) {
    throw "verification did not write metrics JSON"
  }
  $metricsJson = Get-Content -LiteralPath $metricsPath -Raw -Encoding UTF8
  $metrics = $metricsJson | ConvertFrom-Json

  $summaryLines = @(
    "파일: $outputPath",
    "",
    "전체 기준:",
    "전체 생산진도율 ``$($metrics.production_rate_text)``, 포장진도율 ``$($metrics.packing_rate_text)``입니다.",
    "생산부족 ``$(Format-Count $metrics.production_shortage)`` / 전일대비 $(Format-DeltaText $metrics.production_decrease) / 컬러종수 $(Format-DeltaText $metrics.production_color_change '종')",
    "사출부족 ``$(Format-Count $metrics.injection_shortage)`` / 전일대비 $(Format-DeltaText $metrics.injection_decrease) / 컬러종수 $(Format-DeltaText $metrics.injection_color_change '종')",
    "포장부족 ``$(Format-Count $metrics.packing_shortage)``",
    "",
    "포장 원본 1-Day: $($metrics.packing_1day)",
    "포장 원본 FRP: $($metrics.packing_frp)",
    "검증: 캐나다향 0건, 오류 문자열 0건"
  )
  $summaryLines | Set-Content -LiteralPath $summaryPath -Encoding UTF8
  $summaryLines | ForEach-Object { Write-Host $_ }
}
finally {
  Stop-Transcript | Out-Null
}
