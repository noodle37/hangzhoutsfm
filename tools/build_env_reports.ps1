# 生成 web_data/drivers_daily.csv 与 web_data/reports_daily.json
# 驱动因子：data/weather/*_daily.csv（气温/风速/降水）+ data/*_2026_data_summer.xls（水温 C 列 / 溶解氧 E 列 / 透光率 S 列，逐日均值）
# 风险预警：日报/日报/*.docx 中 表1 未来三天天气预报、表2 蓝藻水华预报、五、建议
# 用法：powershell -ExecutionPolicy Bypass -File tools\build_env_reports.ps1
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = Split-Path -Parent $PSScriptRoot
$out = Join-Path $root 'web_data'
$inv = [Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object System.Text.UTF8Encoding $false
$ids = @('jiangjunyan', 'meichengmatou', 'yanlingwu', 'yanziling')

function Read-ZipEntry($path, $entry) {
  $zip = [IO.Compression.ZipFile]::OpenRead($path)
  try {
    $reader = New-Object IO.StreamReader($zip.GetEntry($entry).Open(), $utf8)
    $t = $reader.ReadToEnd(); $reader.Close(); return $t
  } finally { $zip.Dispose() }
}
function Parse-Num($s) {
  $v = 0.0
  if ([double]::TryParse([string]$s, [Globalization.NumberStyles]::Float, $inv, [ref]$v)) { return $v }
  return $null
}

# ---------- drivers_daily.csv ----------
$drivers = @{}   # key "id|date" -> hashtable of sums
foreach ($id in $ids) {
  $wx = Import-Csv (Join-Path $root "data\weather\${id}_daily.csv")
  foreach ($r in $wx) {
    if ($r.date -lt '2026-08-01' -or $r.date -gt '2026-09-30') { continue }
    $drivers["$id|$($r.date)"] = @{
      air_temp = Parse-Num $r.temperature_2m_mean
      wind     = Parse-Num $r.wind_speed_10m_max
      precip   = Parse-Num $r.precipitation_sum
    }
  }
  $xml = Read-ZipEntry (Join-Path $root "data\${id}_2026_data_summer.xls") 'xl/worksheets/sheet1.xml'
  $acc = @{}
  foreach ($row in [regex]::Matches($xml, '<row[^>]*>(.*?)</row>', 'Singleline')) {
    $cells = @{}
    foreach ($c in [regex]::Matches($row.Groups[1].Value, '<c r="([A-Z]+)\d+"[^>]*>(.*?)</c>', 'Singleline')) {
      $cells[$c.Groups[1].Value] = ($c.Groups[2].Value -replace '<[^>]+>', '').Trim()
    }
    $t = [string]$cells['B']
    if ($t -notmatch '^\d{4}-\d{2}-\d{2}') { continue }
    $date = $t.Substring(0, 10)
    if ($date -lt '2026-08-01' -or $date -gt '2026-09-30') { continue }
    if (-not $acc.ContainsKey($date)) { $acc[$date] = @{ C = @(0.0, 0); E = @(0.0, 0); S = @(0.0, 0) } }
    foreach ($col in 'C', 'E', 'S') {
      $v = Parse-Num $cells[$col]
      if ($null -ne $v) { $acc[$date][$col][0] += $v; $acc[$date][$col][1]++ }
    }
  }
  foreach ($date in $acc.Keys) {
    $key = "$id|$date"
    if (-not $drivers.ContainsKey($key)) { $drivers[$key] = @{} }
    foreach ($pair in @(@('C', 'water_temp'), @('E', 'do'), @('S', 'light'))) {
      $a = $acc[$date][$pair[0]]
      $drivers[$key][$pair[1]] = if ($a[1] -gt 0) { [math]::Round($a[0] / $a[1], 3) } else { $null }
    }
  }
}
$cols = 'air_temp', 'wind', 'precip', 'water_temp', 'do', 'light'
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('station_id,date,' + ($cols -join ','))
foreach ($key in ($drivers.Keys | Sort-Object)) {
  $p = $key.Split('|'); $d = $drivers[$key]
  $vals = foreach ($c in $cols) { if ($null -ne $d[$c]) { ([double]$d[$c]).ToString('0.###', $inv) } else { '' } }
  $lines.Add("$($p[0]),$($p[1])," + ($vals -join ','))
}
[IO.File]::WriteAllText((Join-Path $out 'drivers_daily.csv'), (($lines -join "`n") + "`n"), $utf8)
Write-Host "drivers_daily.csv: $($lines.Count - 1) rows"

# ---------- reports_daily.json ----------
$nameMap = [ordered]@{ '严子陵' = 'yanziling'; '严陵坞' = 'yanlingwu'; '梅城码头' = 'meichengmatou'; '将军岩' = 'jiangjunyan' }
$reports = [ordered]@{}
$files = Get-ChildItem (Join-Path $root '日报\日报') -Filter '*202609*.docx' | Sort-Object Name
foreach ($f in $files) {
  if ($f.Name -notmatch '(2026\d{4})') { continue }
  $raw = $Matches[1]
  $iso = $raw.Substring(0, 4) + '-' + $raw.Substring(4, 2) + '-' + $raw.Substring(6, 2)
  $xml = Read-ZipEntry $f.FullName 'word/document.xml'
  $text = $xml -replace '</w:p>', "`n" -replace '<[^>]+>', ''
  $text = [System.Net.WebUtility]::HtmlDecode($text)
  $ls = @($text -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })

  $issue = ''
  foreach ($l in $ls[0..([math]::Min(6, $ls.Count - 1))]) { if ($l -match '第\s*(\d+)\s*期') { $issue = $Matches[1]; break } }

  # 表1：未来三天天气预报；风向/风力可能合并为一个单元格，行宽 6 或 7
  $weather = New-Object System.Collections.Generic.List[object]
  for ($i = 0; $i -lt $ls.Count; $i++) {
    if ($ls[$i] -notmatch '^(\d{1,2})月\s*(\d{1,2})\s*日$') { continue }
    $mo = [int]$Matches[1]; $dy = [int]$Matches[2]
    $t1 = -1
    for ($j = $i + 2; $j -le [math]::Min($i + 4, $ls.Count - 1); $j++) { if ($ls[$j] -match '^\d+(\.\d+)?℃$') { $t1 = $j; break } }
    if ($t1 -lt 0 -or $t1 + 2 -ge $ls.Count) { continue }
    if ($ls[$t1 + 1] -notmatch '℃$') { continue }
    if ($ls[$t1 + 2] -notmatch '有利$|不利$|^一般$') { continue }
    $weather.Add([ordered]@{
      date      = ('2026-{0:00}-{1:00}' -f $mo, $dy)
      cond      = $ls[$i + 1]
      wind      = (($ls[($i + 2)..($t1 - 1)]) -join ' ').Trim()
      tmin      = Parse-Num ($ls[$t1] -replace '℃', '')
      tmax      = Parse-Num ($ls[$t1 + 1] -replace '℃', '')
      favorable = $ls[$t1 + 2]
    })
    $i = $t1 + 2
  }

  # 表2：蓝藻浓度预测及水华风险预报（日期 + 4 站 × 浓度/概率/等级）
  $risk = New-Object System.Collections.Generic.List[object]
  $inRisk = $false; $curDate = ''
  for ($i = 0; $i -lt $ls.Count; $i++) {
    $l = $ls[$i]
    if ($l -match '蓝藻浓度预测及水华风险预报') { $inRisk = $true; continue }
    if (-not $inRisk) { continue }
    if ($l -match '^表3' -or $l -match '依据《湖库') { break }
    if ($l -match '^(\d{1,2})月\s*(\d{1,2})\s*日$') { $curDate = ('2026-{0:00}-{1:00}' -f [int]$Matches[1], [int]$Matches[2]); continue }
    $sid = $null
    foreach ($n in $nameMap.Keys) { if ($l -eq $n) { $sid = $nameMap[$n]; break } }
    if (-not $sid -or -not $curDate -or $i + 3 -ge $ls.Count) { continue }
    $conc = Parse-Num $ls[$i + 1]
    $prob = $null
    if ($ls[$i + 2] -match '^([\d.]+)\s*%$') { $prob = Parse-Num $Matches[1] }
    $level = if ($ls[$i + 3] -match '风险$') { $ls[$i + 3] } else { $null }
    $risk.Add([ordered]@{ date = $curDate; station = $sid; conc = $conc; prob = $prob; level = $level })
    $i += 3
  }

  # 五、建议
  $advice = ''
  for ($i = 0; $i -lt $ls.Count; $i++) {
    if ($ls[$i] -match '^[一二三四五六七八九十]+、\s*建议' -or $ls[$i] -eq '五、建议') {
      $buf = @()
      for ($j = $i + 1; $j -lt $ls.Count; $j++) {
        if ($ls[$j] -match '^附图' -or $ls[$j] -match '^[一二三四五六七八九十]+、') { break }
        if ($ls[$j] -match '^(left)?\d+$') { continue }
        $buf += $ls[$j]
      }
      $advice = ($buf -join '')
      break
    }
  }
  if (-not $advice -and -not $risk.Count) { Write-Warning "$($f.Name): no risk/advice extracted"; continue }

  # 表3：预测与实测对比（实测区间/日报原预测/备注）；HTML 版页面会将预测列替换为 RSI
  $compare = New-Object System.Collections.Generic.List[object]
  $inCmp = $false; $curDate = ''
  for ($i = 0; $i -lt $ls.Count; $i++) {
    $l = $ls[$i]
    if ($l -match '^表3') { $inCmp = $true; continue }
    if (-not $inCmp) { continue }
    if ($l -match '^附图') { break }
    if ($l -match '^(\d{1,2})月\s*(\d{1,2})\s*日$') { $curDate = ('2026-{0:00}-{1:00}' -f [int]$Matches[1], [int]$Matches[2]); continue }
    $sid = $null
    foreach ($n in $nameMap.Keys) { if ($l -eq $n) { $sid = $nameMap[$n]; break } }
    if (-not $sid -or -not $curDate -or $i + 3 -ge $ls.Count) { continue }
    $compare.Add([ordered]@{ date = $curDate; station = $sid; observed = $ls[$i + 1]; report_pred = (Parse-Num $ls[$i + 2]); note = $ls[$i + 3] })
    $i += 3
  }

  # 章节原文（一~五）：保留原始行，供日报 HTML 版按结构渲染
  $heads = @()
  for ($i = 0; $i -lt $ls.Count; $i++) { if ($ls[$i] -match '^([一二三四五])、') { $heads += , @($i, $Matches[1]) } }
  $secs = [ordered]@{}
  for ($h = 0; $h -lt $heads.Count; $h++) {
    $start = $heads[$h][0]
    $end = if ($h + 1 -lt $heads.Count) { $heads[$h + 1][0] } else { $ls.Count }
    $buf = @()
    for ($j = $start + 1; $j -lt $end; $j++) {
      if ($ls[$j] -match '^附图') { break }
      if ($ls[$j] -match '^(left|center)?\d+$') { continue }
      $buf += $ls[$j]
    }
    $secs[$heads[$h][1]] = @{ title = $ls[$start]; lines = $buf }
  }
  $reports[$iso] = [ordered]@{ issue = $issue; file = $f.Name; weather = $weather; risk = $risk; compare = $compare; advice = $advice; sections = $secs }
}
$json = ConvertTo-Json $reports -Depth 8
[IO.File]::WriteAllText((Join-Path $out 'reports_daily.json'), $json, $utf8)
Write-Host "reports_daily.json: $($reports.Count) days"
foreach ($k in $reports.Keys) {
  $r = $reports[$k]
  Write-Host ("  {0}: risk={1} weather={2} compare={3} secs={4} advice={5}chars" -f $k, $r.risk.Count, $r.weather.Count, $r.compare.Count, $r.sections.Count, $r.advice.Length)
}