# 由 H1/H3 详细对比 Markdown 与 data/*_2026_data_summer.xls(实为 xlsx) 生成页面读取的 CSV
# 用法：powershell -ExecutionPolicy Bypass -File tools\build_web_data.ps1
# 输出 web_data\observed_daily.csv 与 web_data\forecast_rsi.csv；web_data\stations.csv 为手工维护，不会被覆盖
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = Split-Path -Parent $PSScriptRoot
$out = Join-Path $root 'web_data'
$inv = [Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object System.Text.UTF8Encoding $false
$year = '2026'

# cyano_ugl_max 的原始列：将军岩/严陵坞只有“蓝绿藻”探头(F 列)，梅城码头/严子陵取“蓝藻 叶绿素a”(I 列)
$sources = [ordered]@{
  jiangjunyan   = @{ file = 'jiangjunyan_2026_data_summer.xls';   col = 'F' }
  meichengmatou = @{ file = 'meichengmatou_2026_data_summer.xls'; col = 'I' }
  yanlingwu     = @{ file = 'yanlingwu_2026_data_summer.xls';     col = 'F' }
  yanziling     = @{ file = 'yanziling_2026_data_summer.xls';     col = 'I' }
}

function Get-DailyMax($id) {
  $src = $sources[$id]
  $zip = [IO.Compression.ZipFile]::OpenRead((Join-Path $root "data\$($src.file)"))
  try {
    $reader = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open(), $utf8)
    $xml = $reader.ReadToEnd(); $reader.Close()
  } finally { $zip.Dispose() }
  $days = @{}
  foreach ($row in [regex]::Matches($xml, '<row[^>]*>(.*?)</row>', 'Singleline')) {
    $cells = @{}
    foreach ($c in [regex]::Matches($row.Groups[1].Value, '<c r="([A-Z]+)\d+"[^>]*>(.*?)</c>', 'Singleline')) {
      $cells[$c.Groups[1].Value] = ($c.Groups[2].Value -replace '<[^>]+>', '').Trim()
    }
    $t = [string]$cells['B']
    if ($t -notmatch '^\d{4}-\d{2}-\d{2}') { continue }
    $v = 0.0
    if (-not [double]::TryParse([string]$cells[$src.col], [Globalization.NumberStyles]::Float, $inv, [ref]$v)) { continue }
    $date = $t.Substring(0, 10)
    if ($days.ContainsKey($date)) {
      $days[$date].n++
      if ($v -gt $days[$date].max) { $days[$date].max = $v }
    } else { $days[$date] = @{ max = $v; n = 1 } }
  }
  return $days
}

function Read-Md($name) { [IO.File]::ReadAllLines((Join-Path $root $name), $utf8) }
function Split-Row($line) { $p = $line.Split('|'); $p[1..($p.Length - 2)] | ForEach-Object { ($_ -replace '\*', '').Trim() } }
function Shift-Date($iso, $d) { ([datetime]::ParseExact($iso, 'yyyy-MM-dd', $inv)).AddDays($d).ToString('yyyy-MM-dd') }

$forecasts = New-Object System.Collections.Generic.List[object]

# H1：第 4 节逐日表，行 = 目标日 t+1；起报日 origin = 目标日 - 1
$st = $null
foreach ($line in (Read-Md '2026_h1_detailed_comparison.md')) {
  if ($line -match '^### 4\.\d+ .*?([a-z]+)\W') { $st = $Matches[1]; continue }
  if ($line -match '^## 5') { $st = $null }
  if (-not $st -or $line -notmatch '^\| \d{2}-\d{2} \|') { continue }
  $c = @(Split-Row $line); $target = "$year-$($c[0])"
  $forecasts.Add([pscustomobject]@{ s = $st; m = 'RSI-H1'; o = (Shift-Date $target -1); t = $target; h = 1; p = $c[2]; q10 = ''; q90 = ''; a = $c[1] })
}

# H3：第 5 节各站“RSI-H3 每日 point / q10 / q90”表
$st = $null; $inTable = $false
foreach ($line in (Read-Md '2026_h3_detailed_comparison.md')) {
  if ($line -match '^### 5\.\d+ .*?([a-z]+)\W*$') { $st = $Matches[1]; $inTable = $false; continue }
  if ($line -match '^#### .*RSI-H3 .*point') { $inTable = $true; continue }
  if ($line -match '^#') { $inTable = $false }
  if (-not $st -or -not $inTable -or $line -notmatch '^\| \d{2}-\d{2} \|') { continue }
  $c = @(Split-Row $line)
  $forecasts.Add([pscustomobject]@{ s = $st; m = 'RSI-H3'; o = "$year-$($c[0])"; t = "$year-$($c[1])"; h = [int]$c[2]; p = $c[3]; q10 = $c[4]; q90 = $c[5]; a = $c[6] })
}

New-Item -ItemType Directory -Force $out | Out-Null
$obs = New-Object System.Collections.Generic.List[string]
$obs.Add('station_id,date,cyano_ugl_max,n_samples')
$mismatch = 0
foreach ($id in $sources.Keys) {
  $days = Get-DailyMax $id
  foreach ($date in ($days.Keys | Sort-Object)) {
    if ($date -lt "$year-08-01" -or $date -gt "$year-09-30") { continue }
    $obs.Add(('{0},{1},{2},{3}' -f $id, $date, $days[$date].max.ToString('0.000', $inv), $days[$date].n))
  }
  foreach ($f in ($forecasts | Where-Object { $_.s -eq $id })) {
    $o = $days[$f.t]
    if (-not $o -or [math]::Abs($o.max - [double]::Parse($f.a, $inv)) -gt 0.006) {
      $mismatch++; Write-Warning "actual mismatch $id $($f.m) $($f.t) md=$($f.a) xlsx=$($o.max)"
    }
  }
}
[IO.File]::WriteAllText((Join-Path $out 'observed_daily.csv'), (($obs -join "`n") + "`n"), $utf8)

$fc = New-Object System.Collections.Generic.List[string]
$fc.Add('station_id,model,origin_date,target_date,horizon,point,q10,q90,actual')
foreach ($f in $forecasts) { $fc.Add(('{0},{1},{2},{3},{4},{5},{6},{7},{8}' -f $f.s, $f.m, $f.o, $f.t, $f.h, $f.p, $f.q10, $f.q90, $f.a)) }
[IO.File]::WriteAllText((Join-Path $out 'forecast_rsi.csv'), (($fc -join "`n") + "`n"), $utf8)

Write-Host "observed_daily.csv: $($obs.Count - 1) rows; forecast_rsi.csv: $($fc.Count - 1) rows"
if ($mismatch) { Write-Warning "$mismatch actual mismatches" } else { Write-Host 'Markdown actuals match xlsx daily max' }
