param(
  [int]$Port = 8000
)

$Base = Get-Location
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Preview server listening on http://localhost:$Port/"

try {
  while ($true) {
    $ctx = $listener.GetContext()
    $rel = $ctx.Request.Url.LocalPath.TrimStart('/')
    if ([string]::IsNullOrWhiteSpace($rel)) { $rel = 'index.html' }
    $path = Join-Path $Base $rel

    if (Test-Path $path) {
      $bytes = [System.IO.File]::ReadAllBytes($path)
      $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
      switch ($ext) {
        '.html' { $ctx.Response.ContentType = 'text/html' }
        '.css'  { $ctx.Response.ContentType = 'text/css' }
        '.js'   { $ctx.Response.ContentType = 'application/javascript' }
        '.png'  { $ctx.Response.ContentType = 'image/png' }
        '.svg'  { $ctx.Response.ContentType = 'image/svg+xml' }
        default { $ctx.Response.ContentType = 'application/octet-stream' }
      }
      $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
      $ctx.Response.StatusCode = 404
      $msg = "Not Found: $rel"
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($msg)
      $ctx.Response.OutputStream.Write($bytes,0,$bytes.Length)
    }
    $ctx.Response.Close()
  }
}
finally {
  $listener.Stop()
  $listener.Close()
}