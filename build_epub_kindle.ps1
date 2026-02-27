param(
  [string]$Root = "c:\GitHub\book-poo2",
  [string]$Output = "Livro_POO2_kindle.epub"
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$outPath = Join-Path $Root $Output
$oebpsBase = Join-Path $Root 'OEBPS'
$kindleImgDir = Join-Path $oebpsBase 'images\kindle'

if (Test-Path $outPath) { Remove-Item $outPath -Force }

$zipStream = [System.IO.File]::Create($outPath)
$zip = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)

function Add-ZipEntry {
  param($Archive, $FullPath, $EntryPath, $NoCompress = $false)
  $level = if ($NoCompress) { [System.IO.Compression.CompressionLevel]::NoCompression } else { [System.IO.Compression.CompressionLevel]::Optimal }
  $entry = $Archive.CreateEntry($EntryPath, $level)
  $inStream = [System.IO.File]::OpenRead($FullPath)
  $outStream = $entry.Open()
  try { $inStream.CopyTo($outStream) } finally { $inStream.Dispose(); $outStream.Dispose() }
}

function Add-ZipEntryContent {
  param($Archive, $EntryPath, $Content)
  $entry = $Archive.CreateEntry($EntryPath, [System.IO.Compression.CompressionLevel]::Optimal)
  $outStream = $entry.Open()
  try { $bytes = [System.Text.Encoding]::UTF8.GetBytes($Content); $outStream.Write($bytes, 0, $bytes.Length) } finally { $outStream.Dispose() }
}

function New-KindleImage {
  param([string]$SrcPath, [string]$DestDir, [int]$Quality = 55, [int]$MaxWidth = 900)
  $name = [System.IO.Path]::GetFileNameWithoutExtension($SrcPath)
  $dest = Join-Path $DestDir "$name.jpg"
  if (Test-Path $dest) { return $dest }
  $img = [System.Drawing.Image]::FromFile($SrcPath)
  try {
    $w = $img.Width; $h = $img.Height
    if ($w -gt $MaxWidth) {
      $ratio = $MaxWidth / $w
      $nw = [int]([math]::Round($w * $ratio))
      $nh = [int]([math]::Round($h * $ratio))
    } else {
      $nw = $w; $nh = $h
    }
    $bmp = New-Object System.Drawing.Bitmap $nw, $nh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    try {
      $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
      $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
      $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
      $g.DrawImage($img, 0, 0, $nw, $nh)
    } finally { $g.Dispose() }
    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
    $ep = New-Object System.Drawing.Imaging.EncoderParameters 1
    $enc = [System.Drawing.Imaging.Encoder]::Quality
    $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ($enc, [int]$Quality)
    if (!(Test-Path $DestDir)) { New-Item -ItemType Directory -Path $DestDir | Out-Null }
    $bmp.Save($dest, $codec, $ep)
    $bmp.Dispose()
  } finally { $img.Dispose() }
  return $dest
}

# 1) mimetype (sem compressao, primeira entrada - CRITICO para EPUB)
Add-ZipEntry -Archive $zip -FullPath (Join-Path $Root 'mimetype') -EntryPath 'mimetype' -NoCompress $true

# 2) META-INF
Add-ZipEntry -Archive $zip -FullPath (Join-Path $Root 'META-INF/container.xml') -EntryPath 'META-INF/container.xml'

# 3) Gerar content.opf limpo para Kindle (com caminhos corretos)
$opfPath = Join-Path $oebpsBase 'content.opf'
$opfXml = [xml](Get-Content -LiteralPath $opfPath -Raw -Encoding UTF8)
$ns = New-Object System.Xml.XmlNamespaceManager($opfXml.NameTable)
$ns.AddNamespace('opf', 'http://www.idpf.org/2007/opf')

# Pré-gerar versões compactas para PNG/JPG
$imgBase = Join-Path $oebpsBase 'images'
if (!(Test-Path $kindleImgDir)) { New-Item -ItemType Directory -Path $kindleImgDir | Out-Null }
$rasters = Get-ChildItem $imgBase -File | Where-Object { $_.Extension -in @('.png', '.jpg', '.jpeg') }
foreach ($rf in $rasters) { New-KindleImage -SrcPath $rf.FullName -DestDir $kindleImgDir | Out-Null }

# Atualizar hrefs no manifest
$items = $opfXml.SelectNodes('//opf:manifest/opf:item', $ns)
foreach ($item in $items) {
  $href = $item.GetAttribute('href')
  if ($href -match '^images\/(.+)\.(png|jpg|jpeg)$') {
    $name = $Matches[1]
    $item.SetAttribute('href', "images/kindle/$name.jpg")
    $item.SetAttribute('media-type', 'image/jpeg')
  }
}

# Remover cap00_modelo do EPUB (apenas para Kindle)
$cap00Item = $opfXml.SelectSingleNode('//opf:manifest/opf:item[@id="cap00_modelo" or @href="text/cap00_modelo.html"]', $ns)
if ($cap00Item -ne $null) { $cap00Item.ParentNode.RemoveChild($cap00Item) | Out-Null }
$cap00Ref = $opfXml.SelectSingleNode('//opf:spine/opf:itemref[@idref="cap00_modelo"]', $ns)
if ($cap00Ref -ne $null) { $cap00Ref.ParentNode.RemoveChild($cap00Ref) | Out-Null }

Add-ZipEntryContent -Archive $zip -EntryPath 'OEBPS/content.opf' -Content $opfXml.OuterXml

# 4) NCX e NAV
Add-ZipEntry -Archive $zip -FullPath (Join-Path $oebpsBase 'toc.ncx') -EntryPath 'OEBPS/toc.ncx'
Add-ZipEntry -Archive $zip -FullPath (Join-Path $oebpsBase 'nav.xhtml') -EntryPath 'OEBPS/nav.xhtml'

# 5) CSS
Add-ZipEntry -Archive $zip -FullPath (Join-Path $oebpsBase 'css/style.css') -EntryPath 'OEBPS/css/style.css'

# 6) HTML files - substituir caminhos de imagens
$htmlFiles = Get-ChildItem (Join-Path $oebpsBase 'text') -Filter '*.html'
$referencedImages = New-Object System.Collections.Generic.HashSet[string]
foreach ($hf in $htmlFiles) {
  $content = Get-Content -LiteralPath $hf.FullName -Raw -Encoding UTF8
  
  $content = [regex]::Replace($content, "src=['""]\.\.\/images\/([^'""]+)\.(png|jpg|jpeg)['""]", "src='../images/kindle/`$1.jpg'")
  $matches = [regex]::Matches($content, "src=['""]\.\.\/(images\/[^'""]+)['""]")
  foreach ($m in $matches) { $referencedImages.Add($m.Groups[1].Value) | Out-Null }
  
  $content = [regex]::Replace($content, '<img[^>]*class\s*=\s*["'']icon-img["''][^>]*>', '', 'Singleline')
  
  if ($content -notmatch 'data-noicons') {
    $css = '<style data-noicons>.box-title .icon-img{display:none !important;}.box-title::before{content:"" !important;}</style>'
    $content = $content -replace '</head>', "$css</head>"
  }
  
  if ($content -notmatch 'data-kindle') {
    $kindleCss = @"
<style data-kindle>
/* Texto comum: figuras ocupam no máximo 96% da largura */
figure img, figure svg { max-width: 96% !important; height: auto !important; margin: 0 auto !important; display: block !important; }
img, svg { max-width: 98% !important; height: auto !important; }
/* Blocos coloridos: figuras mais estreitas para acomodar bordas/padding */
.box-content img, .box-content svg, .box-table img, .box-content figure img, .box-content figure svg { 
  max-width: 88% !important; 
  height: auto !important; 
  margin: 0.25rem auto !important; 
  display: block !important;
  box-sizing: border-box !important;
}
.box-table { width: 100% !important; table-layout: fixed !important; }
.box-content { word-wrap: break-word !important; overflow-wrap: anywhere !important; hyphens: auto !important; }
figcaption { max-width: 100% !important; }
/* Listagens e códigos: garantir visibilidade com rolagem horizontal */
figure.listing pre.code, pre.code { 
  white-space: pre !important; 
  overflow-x: auto !important; 
  display: block !important; 
  box-sizing: border-box !important;
  background: #f7f7f7 !important;
  border: 1px solid #ddd !important;
  padding: 0.6rem !important;
  font-family: Consolas, "Courier New", monospace !important;
  font-size: 0.90rem !important;
  color: #000 !important;
}
code { white-space: pre-wrap; color: #000 !important; }
</style>
"@
    $content = $content -replace '</head>', "$kindleCss</head>"
  }
  
  Add-ZipEntryContent -Archive $zip -EntryPath "OEBPS/text/$($hf.Name)" -Content $content
}

# 7) Referências do manifest
$manifestRefs = $opfXml.SelectNodes('//opf:manifest/opf:item', $ns) | ForEach-Object { $_.GetAttribute('href') } | Where-Object { $_ -like 'images/*' }
foreach ($href in $manifestRefs) { $referencedImages.Add($href) | Out-Null }

# 8) Incluir apenas imagens referenciadas
$included = 0
foreach ($path in $referencedImages) {
  if ($path -like 'images/kindle/*') {
    $fn = [System.IO.Path]::GetFileName($path)
    $full = Join-Path $kindleImgDir $fn
    if (Test-Path $full) { Add-ZipEntry -Archive $zip -FullPath $full -EntryPath "OEBPS/$path"; $included++ }
  } elseif ($path -like 'images/*.svg') {
    $full = Join-Path $imgBase ([System.IO.Path]::GetFileName($path))
    if (Test-Path $full) { Add-ZipEntry -Archive $zip -FullPath $full -EntryPath "OEBPS/$path"; $included++ }
  } else {
    $full = Join-Path $imgBase ([System.IO.Path]::GetFileName($path))
    if (Test-Path $full) {
      $jpg = New-KindleImage -SrcPath $full -DestDir $kindleImgDir
      $fn = [System.IO.Path]::GetFileName($jpg)
      Add-ZipEntry -Archive $zip -FullPath $jpg -EntryPath "OEBPS/images/kindle/$fn"
      $included++
    }
  }
}
Write-Host "Incluidas $included imagens referenciadas"

Remove-Variable referencedImages -ErrorAction Ignore

# Fechar ZIP
$zip.Dispose()
$zipStream.Dispose()

$size = [math]::Round((Get-Item $outPath).Length / 1MB, 2)
Write-Host "EPUB Kindle gerado: $outPath ($size MB)"
