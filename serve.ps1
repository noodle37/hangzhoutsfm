$root = "C:\Users\Admin\hangzhou"
$prefix = "http://localhost:8000/"

# 从本地 .env 读取天地图 Key（.env 已被 gitignore，不会上传）
$tdtKey = ""
$envFile = Join-Path $root ".env"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        if ($_ -match '^\s*TDT_KEY\s*=\s*(.+?)\s*$') { $tdtKey = $Matches[1] }
    }
}
if (-not $tdtKey) { Write-Host "WARNING: .env 中未找到 TDT_KEY，天地图将无法加载" }

# 日报附图：按需从固定位置（日报\日报\*.docx）读取，不预先解压、不预加载
Add-Type -AssemblyName System.IO.Compression.FileSystem
$figDir = Join-Path $root '日报\日报'
$figCache = @{}
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Get-DocxPath([string]$iso) {
    if ($iso -notmatch '^\d{4}-\d{2}-\d{2}$') { return $null }
    $compact = $iso.Replace('-', '')
    $f = Get-ChildItem $figDir -Filter "*$compact*.docx" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($f) { return $f.FullName }
    return $null
}

function Read-ZipText($zip, [string]$name) {
    $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
    if (-not $e) { return '' }
    $s = $e.Open()
    $sr = New-Object IO.StreamReader($s, [System.Text.Encoding]::UTF8)
    $t = $sr.ReadToEnd(); $sr.Close()
    return $t
}

# 返回 { source, captions, matched, images: [{ i, entry, caption }] }；图片顺序 = 正文出现顺序
function Get-FigManifest([string]$iso) {
    if ($figCache.ContainsKey($iso)) { return $figCache[$iso] }
    $path = Get-DocxPath $iso
    if (-not $path) { return $null }
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $xml = Read-ZipText $zip 'word/document.xml'
        $relsXml = Read-ZipText $zip 'word/_rels/document.xml.rels'
        if (-not $xml) { return $null }
        $rels = @{}
        foreach ($m in [regex]::Matches($relsXml, '<Relationship[^>]*>')) {
            $v = $m.Value
            if ($v -notmatch '/image"') { continue }
            $idm = [regex]::Match($v, 'Id="([^"]+)"')
            $tm = [regex]::Match($v, 'Target="([^"]+)"')
            if ($idm.Success -and $tm.Success) { $rels[$idm.Groups[1].Value] = $tm.Groups[1].Value }
        }
        $entries = @()
        foreach ($m in [regex]::Matches($xml, 'r:embed="([^"]+)"')) {
            $id = $m.Groups[1].Value
            if ($rels.ContainsKey($id)) { $entries += ('word/' + ($rels[$id] -replace '^\.\./', '')) }
        }
        $text = [System.Net.WebUtility]::HtmlDecode(($xml -replace '</w:p>', "`n" -replace '<[^>]+>', ''))
        $lines = @($text -split "`n" | ForEach-Object { ($_.Trim() -replace '^(left|center|right)?\d+', '').Trim() } | Where-Object { $_ -ne '' })
        $caps = @($lines | Where-Object { $_ -match '^附图\s*\d+\s*（' })
        # 「多光谱图及可见光图」为双图，其余单图；总数吻合才按顺序套标题，否则只给影像不硬套
        $want = @($caps | ForEach-Object { if ($_ -match '及可见光') { 2 } else { 1 } })
        $need = ($want | Measure-Object -Sum).Sum
        $matched = ($caps.Count -gt 0 -and $need -eq $entries.Count)
        $images = @()
        $acc = 0; $ci = 0
        for ($i = 0; $i -lt $entries.Count; $i++) {
            $cap = ''
            if ($matched) {
                while ($ci -lt $caps.Count -and $i -ge ($acc + $want[$ci])) { $acc += $want[$ci]; $ci++ }
                if ($ci -lt $caps.Count) { $cap = $caps[$ci] }
            }
            $images += [ordered]@{ i = $i; entry = $entries[$i]; caption = $cap }
        }
        $manifest = [ordered]@{ source = [IO.Path]::GetFileName($path); captions = $caps; matched = $matched; images = $images }
        $figCache[$iso] = $manifest
        return $manifest
    } finally { $zip.Dispose() }
}

function Send-Bytes($ctx, [byte[]]$bytes, [string]$mime, [int]$status = 200, [string]$cache = 'no-store') {
    $ctx.Response.StatusCode = $status
    $ctx.Response.ContentType = $mime
    $ctx.Response.Headers.Set('Cache-Control', $cache)
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
}

function Serve-ReportFig($ctx, [string]$path) {
    $iso = [string]$ctx.Request.QueryString['date']
    $manifest = Get-FigManifest $iso
    if (-not $manifest) {
        Send-Bytes $ctx ($utf8NoBom.GetBytes('{"error":"no report for date"}')) 'application/json; charset=utf-8' 404
        return
    }
    if ($path -eq '/report_figs') {
        $pub = [ordered]@{
            source   = $manifest.source
            matched  = $manifest.matched
            captions = $manifest.captions
            images   = @($manifest.images | ForEach-Object { [ordered]@{ i = $_.i; caption = $_.caption; url = ('/report_fig?date=' + $iso + '&i=' + $_.i) } })
        }
        Send-Bytes $ctx ($utf8NoBom.GetBytes((ConvertTo-Json $pub -Depth 6))) 'application/json; charset=utf-8'
        return
    }
    $i = 0
    if (-not [int]::TryParse([string]$ctx.Request.QueryString['i'], [ref]$i)) { $i = -1 }
    $img = $manifest.images | Where-Object { $_.i -eq $i } | Select-Object -First 1
    if (-not $img) {
        Send-Bytes $ctx ($utf8NoBom.GetBytes('404 Not Found')) 'text/plain; charset=utf-8' 404
        return
    }
    $docx = Get-DocxPath $iso
    $zip = [IO.Compression.ZipFile]::OpenRead($docx)
    try {
        $e = $zip.Entries | Where-Object { $_.FullName -eq $img.entry }
        if (-not $e) {
            Send-Bytes $ctx ($utf8NoBom.GetBytes('404 Not Found')) 'text/plain; charset=utf-8' 404
            return
        }
        $ms = New-Object IO.MemoryStream
        $s = $e.Open(); $s.CopyTo($ms); $s.Close()
        $ext = [IO.Path]::GetExtension($e.FullName).ToLower()
        $mime = if ($ext -eq '.png') { 'image/png' } elseif ($ext -eq '.gif') { 'image/gif' } elseif ($ext -eq '.bmp') { 'image/bmp' } else { 'image/jpeg' }
        Send-Bytes $ctx $ms.ToArray() $mime 200 'public, max-age=86400'
        $ms.Dispose()
    } finally { $zip.Dispose() }
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "Serving $root at $prefix"
while ($listener.IsListening) {
    try {
        $ctx = $listener.GetContext()
        $path = $ctx.Request.Url.LocalPath
        if ($path -eq "/report_figs" -or $path -eq "/report_fig") {
            Serve-ReportFig $ctx $path
            $ctx.Response.OutputStream.Close()
            continue
        }
        if ($path -eq "/compare") {
            $file = (Get-ChildItem $root -Filter *benchmark*.html | Select-Object -First 1).FullName
        } elseif ($path -eq "/report") {
            $file = (Get-ChildItem $root -Filter *技术报告*.html | Select-Object -First 1).FullName
        } elseif ($path -eq "/" -or $path -eq "/index.html") {
            # 主页固定为平台交互原型，排除对比页与技术报告
            $file = (Get-ChildItem $root -Filter *.html | Where-Object { $_.Name -NotLike '*benchmark*' -and $_.Name -NotLike '*技术报告*' } | Select-Object -First 1).FullName
        } else {
            $file = Join-Path $root ([uri]::UnescapeDataString($path).TrimStart("/"))
            # 禁止通过 HTTP 访问 .env 等密钥文件
            if ($file -and ([System.IO.Path]::GetFileName($file) -eq ".env")) { $file = $null }
        }
        if ($file -and (Test-Path $file -PathType Leaf)) {
            $ext = [System.IO.Path]::GetExtension($file).ToLower()
            if ($ext -eq ".html" -and $tdtKey) {
                # 注入天地图 Key：仅替换占位符，源文件里不保存真实 Key
                $html = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($html.Replace("__TDT_KEY__", $tdtKey))
            } else {
                $bytes = [System.IO.File]::ReadAllBytes($file)
            }
            $mime = switch ($ext) {
                ".html" { "text/html; charset=utf-8" }
                ".css"  { "text/css; charset=utf-8" }
                ".js"   { "application/javascript; charset=utf-8" }
                ".png"  { "image/png" }
                ".jpg"  { "image/jpeg" }
                ".jpeg" { "image/jpeg" }
                ".svg"  { "image/svg+xml" }
                ".csv"  { "text/csv; charset=utf-8" }
                ".json" { "application/json; charset=utf-8" }
                default { "application/octet-stream" }
            }
            $ctx.Response.ContentType = $mime
            $ctx.Response.ContentLength64 = $bytes.Length
            $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $ctx.Response.StatusCode = 404
            $msg = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
            $ctx.Response.OutputStream.Write($msg, 0, $msg.Length)
        }
        $ctx.Response.OutputStream.Close()
    } catch {
        Write-Host "Error: $_"
    }
}
