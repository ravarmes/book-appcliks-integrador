Add-Type -AssemblyName System.Net.HttpListener
$root = (Get-Location).Path
$prefix = "http://localhost:8000/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "Serving $root at $prefix"

while ($true) {
    $context = $listener.GetContext()
    $request = $context.Request
    $rel = [System.Uri]::UnescapeDataString($request.Url.AbsolutePath.TrimStart('/'))
    if ([string]::IsNullOrEmpty($rel)) { $rel = "index.html" }
    $file = Join-Path $root $rel

    if (Test-Path $file) {
        $ext = [System.IO.Path]::GetExtension($file).ToLower()
        switch ($ext) {
            ".html" { $context.Response.ContentType = "text/html" }
            ".css"  { $context.Response.ContentType = "text/css" }
            ".js"   { $context.Response.ContentType = "application/javascript" }
            ".svg"  { $context.Response.ContentType = "image/svg+xml" }
            ".png"  { $context.Response.ContentType = "image/png" }
            ".jpg"  { $context.Response.ContentType = "image/jpeg" }
            ".jpeg" { $context.Response.ContentType = "image/jpeg" }
            default  { $context.Response.ContentType = "application/octet-stream" }
        }
        $bytes = [System.IO.File]::ReadAllBytes($file)
        $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        $context.Response.StatusCode = 404
        $writer = New-Object System.IO.StreamWriter($context.Response.OutputStream)
        $writer.Write("404 Not Found")
        $writer.Flush()
    }

    $context.Response.Close()
}