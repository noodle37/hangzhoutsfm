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
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "Serving $root at $prefix"
while ($listener.IsListening) {
    try {
        $ctx = $listener.GetContext()
        $path = $ctx.Request.Url.LocalPath
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
                ".svg"  { "image/svg+xml" }
                ".csv"  { "text/csv; charset=utf-8" }
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
