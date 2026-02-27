Param(
  [string]$Root = "c:\Projetos\_Livros\APOSTILA_DB_EPUB",
  [int]$Width = 1600,
  [switch]$Overwrite
)

$ErrorActionPreference = 'Stop'

$imagesDir = Join-Path $Root 'OEBPS\images'
if (-not (Test-Path -LiteralPath $imagesDir)) {
  Write-Error "Diretório de imagens não encontrado: $imagesDir"
}

function Find-Inkscape {
  $cmd = Get-Command inkscape -ErrorAction SilentlyContinue
  $maybe = $null; if ($cmd) { $maybe = $cmd.Path }
  $paths = @(
    $maybe,
    'C:\\Program Files\\Inkscape\\bin\\inkscape.exe',
    'C:\\Program Files\\Inkscape\\inkscape.exe',
    'C:\\Program Files\\Inkscape\\bin\\inkscape.com',
    'C:\\Program Files (x86)\\Inkscape\\inkscape.exe'
  ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
  if ($paths -and $paths[0]) { return $paths[0] } else { return $null }
}

$inkPath = Find-Inkscape
if (-not $inkPath) {
  Write-Warning "Inkscape não encontrado. Instale-o em 'C:\\Program Files\\Inkscape' ou adicione ao PATH. Vou continuar, mas nenhuma conversão será efetuada."
}

$svgs = Get-ChildItem -LiteralPath $imagesDir -Filter *.svg -File -ErrorAction Stop
foreach ($svg in $svgs) {
  $pngPath = [System.IO.Path]::ChangeExtension($svg.FullName, '.png')
  if (-not $Overwrite -and (Test-Path -LiteralPath $pngPath)) {
    Write-Host "PNG já existe, pulando: " $svg.Name
    continue
  }
  if (-not $inkPath) { continue }
  Write-Host "Convertendo: " $svg.Name " -> " ([System.IO.Path]::GetFileName($pngPath))
  & $inkPath --export-type=png --export-filename="$pngPath" --export-width=$Width "$($svg.FullName)"
}

Write-Host "Conversão finalizada."