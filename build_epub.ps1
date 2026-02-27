param(
  [string]$Root = "c:\\GitHub\\book-poo2",
  [string]$Output = "Livro_POO2.epub",
  [switch]$NoIcons,
  [switch]$PreferPng,
  [switch]$SanitizeManifest,
  [switch]$NoCover,
  [switch]$RemoveSvg,
  [switch]$UseCompressedJpg
)

$ErrorActionPreference = 'Stop'

# Carregar APIs de compressÃ£o
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$rootPath = $Root
$outPath = Join-Path $rootPath $Output

if (Test-Path $outPath) {
  Remove-Item $outPath -Force
}

# Abrir arquivo ZIP (EPUB) para criaÃ§Ã£o
$zipStream = [System.IO.File]::Create($outPath)
$zip = New-Object System.IO.Compression.ZipArchive($zipStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)

function Add-ZipEntry {
  param(
    [System.IO.Compression.ZipArchive]$Archive,
    [string]$FullPath,
    [string]$EntryPath,
    [bool]$NoCompress = $false
  )
  $level = [System.IO.Compression.CompressionLevel]::Optimal
  if ($NoCompress) { $level = [System.IO.Compression.CompressionLevel]::NoCompression }
  $entry = $Archive.CreateEntry($EntryPath, $level)
  $inStream = [System.IO.File]::OpenRead($FullPath)
  $outStream = $entry.Open()
  try {
    $inStream.CopyTo($outStream)
  } finally {
    $inStream.Dispose(); $outStream.Dispose()
  }
}

function Add-ZipEntryContent {
  param(
    [System.IO.Compression.ZipArchive]$Archive,
    [string]$EntryPath,
    [string]$Content,
    [bool]$NoCompress = $false
  )
  $level = [System.IO.Compression.CompressionLevel]::Optimal
  if ($NoCompress) { $level = [System.IO.Compression.CompressionLevel]::NoCompression }
  $entry = $Archive.CreateEntry($EntryPath, $level)
  $outStream = $entry.Open()
  try {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Content)
    $outStream.Write($bytes, 0, $bytes.Length)
  } finally {
    $outStream.Dispose()
  }
}

# 1) mimetype (sem compressÃ£o e primeira entrada)
Add-ZipEntry -Archive $zip -FullPath (Join-Path $rootPath 'mimetype') -EntryPath 'mimetype' -NoCompress $true

# 2) META-INF/container.xml
Add-ZipEntry -Archive $zip -FullPath (Join-Path $rootPath 'META-INF/container.xml') -EntryPath 'META-INF/container.xml'

# 3) OEBPS: incluir SOMENTE arquivos do manifest + content.opf
$opfPath = Join-Path $rootPath 'OEBPS\content.opf'
$xml = New-Object System.Xml.XmlDocument
$xml.Load($opfPath)
$nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsmgr.AddNamespace('opf','http://www.idpf.org/2007/opf')
$items = $xml.SelectNodes('//opf:manifest/opf:item', $nsmgr)

# Base OEBPS para verificaÃ§Ãµes
$oebpsBase = Join-Path $rootPath 'OEBPS'

# SanitizaÃ§Ã£o do manifest: remover ausentes e preferir PNG quando possÃ­vel
if ($SanitizeManifest -or $PreferPng -or $NoCover -or $RemoveSvg) {
  $manifestNode = $xml.SelectSingleNode('//opf:manifest', $nsmgr)
  # remover duplicidades por href
  $seen = New-Object 'System.Collections.Generic.HashSet[string]'
  foreach ($it in @($items)) {
    $href = $it.GetAttribute('href')
    $full = Join-Path $oebpsBase $href
    $isSvg = $href -match '\.svg$'
    if (-not $seen.Add($href)) {
      Write-Warning "Item duplicado no OPF; removendo: $href"
      $null = $manifestNode.RemoveChild($it)
      continue
    }
    if (-not (Test-Path -LiteralPath $full)) {
      Write-Warning "Item ausente no disco; removendo do OPF: $href"
      $null = $manifestNode.RemoveChild($it)
      continue
    }
    if ($RemoveSvg -and $isSvg) {
      Write-Warning "Removendo SVG do OPF por -RemoveSvg: $href"
      $null = $manifestNode.RemoveChild($it)
      continue
    }
    if ($PreferPng -and $isSvg) {
      $pngHref = ($href -replace '\.svg$', '.png')
      $pngFull = Join-Path $oebpsBase $pngHref
      if (Test-Path -LiteralPath $pngFull) {
        $it.SetAttribute('href', $pngHref)
        $it.SetAttribute('media-type', 'image/png')
      }
    }
  }
  if ($NoCover) {
    # Remover capa (html e imagem) e meta cover
    $coverItems = @($xml.SelectNodes("//opf:manifest/opf:item[contains(@href,'text/capa') or contains(@href,'images/capa') ]", $nsmgr))
    foreach ($ci in $coverItems) {
      $id = $ci.GetAttribute('id')
      foreach ($ref in @($xml.SelectNodes("//opf:spine/opf:itemref[@idref='$id']", $nsmgr))) { $null = $ref.ParentNode.RemoveChild($ref) }
      $null = $ci.ParentNode.RemoveChild($ci)
    }
    foreach ($meta in @($xml.SelectNodes('//opf:metadata/opf:meta', $nsmgr))) {
      if ($meta.GetAttribute('name') -eq 'cover') { $null = $meta.ParentNode.RemoveChild($meta) }
    }
  }
  # Garantir atributo toc="ncx" no spine
  $spine = $xml.SelectSingleNode('//opf:spine', $nsmgr)
  if ($spine -and -not $spine.Attributes['toc']) {
    $spine.SetAttribute('toc', 'ncx')
  }
}

# Recoletar itens e hrefs apÃ³s limpeza
$items = $xml.SelectNodes('//opf:manifest/opf:item', $nsmgr)
$hrefs = @()
foreach ($it in $items) { $hrefs += $it.GetAttribute('href') }
# Deduplicar hrefs por seguranÃ§a
$hrefs = $hrefs | Select-Object -Unique
# Garantir inclusÃ£o de content.opf (serÃ¡ escrito a partir do XML atualizado)
$hrefs += 'content.opf'

foreach ($href in $hrefs) {
  $fullPath = Join-Path $oebpsBase $href
  if (-not (Test-Path -LiteralPath $fullPath)) {
    Write-Warning "Recurso listado no manifest nÃ£o encontrado: $href"
    continue
  }
  $entryPath = 'OEBPS/' + ($href -replace '\\','/')
  # content.opf: escrever versÃ£o saneada do XML
  if ($href -ieq 'content.opf') {
    Add-ZipEntryContent -Archive $zip -EntryPath $entryPath -Content $xml.OuterXml
    continue
  }
  if (($NoIcons -or $PreferPng) -and ($href -match '\.(x?html)$')) {
    # Leitura robusta do arquivo como string
    try {
      $content = Get-Content -LiteralPath $fullPath -Raw -Encoding UTF8
    } catch {
      try { $content = [System.IO.File]::ReadAllText($fullPath) } catch { $content = '' }
    }
    if ($null -eq $content) { $content = '' }

    # Preferir PNGs no Kindle: trocar src de imagens SVG por PNG quando existir
    if ($PreferPng) {
      $imagesDir = Join-Path $oebpsBase 'images'
      $imgMatches = [System.Text.RegularExpressions.Regex]::Matches(
        $content,
        'src=["\'']\.\./images/(?<name>[^"\'']+?)\.svg["\'']',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
      )
      foreach ($m in $imgMatches) {
        $name = $m.Groups['name'].Value
        $pngPath = Join-Path $imagesDir ($name + '.png')
        if (Test-Path -LiteralPath $pngPath) {
          $old = $m.Value
          $new = ($m.Value -replace '\.svg', '.png')
          $content = $content.Replace($old, $new)
        }
      }
    }

    # Usar imagens JPG comprimidas da pasta kindle (substituir PNG por JPG)
    if ($UseCompressedJpg) {
      $kindleDir = Join-Path $oebpsBase 'images\kindle'
      if (Test-Path $kindleDir) {
        $pngMatches = [regex]::Matches($content, "src=[`"']\.\.\/images\/(?<name>[^`"']+?)\.png[`"']", 'IgnoreCase')
        foreach ($m in $pngMatches) {
          $name = $m.Groups['name'].Value
          $jpgPath = Join-Path $kindleDir ($name + '.jpg')
          if (Test-Path -LiteralPath $jpgPath) {
            $old = $m.Value
            $new = $old -replace '\.png', '.jpg' -replace '/images/', '/images/kindle/'
            $content = $content.Replace($old, $new)
          }
        }
      }
    }
    # Remover imagens SVG caso solicitado
    if ($RemoveSvg) {
      $content = [System.Text.RegularExpressions.Regex]::Replace(
        $content,
        '<img[^>]*src\s*=\s*["\'']\.\./images/[^"\'']+?\.svg["\''][^>]*>',
        '',
        [System.Text.RegularExpressions.RegexOptions]::Singleline
      )
      $content = [System.Text.RegularExpressions.Regex]::Replace(
        $content,
        '<object[^>]*data\s*=\s*["\'']\.\./images/[^"\'']+?\.svg["\''][^>]*>.*?</object>',
        '',
        [System.Text.RegularExpressions.RegexOptions]::Singleline
      )
    }

    # Remover <img ... class="icon-img" ...>
    $content = [System.Text.RegularExpressions.Regex]::Replace(
      $content,
      '<img[^>]*class\s*=\s*["\'']icon-img["\''][^>]*>',
      '',
      [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    # Remover <svg ... class="icon-img" ...>...</svg>
    $content = [System.Text.RegularExpressions.Regex]::Replace(
      $content,
      '<svg[^>]*class\s*=\s*["\'']icon-img["\''][^>]*>.*?</svg>',
      '',
      [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    # Remover qualquer conteÃºdo textual/sÃ­mbolos antes de <strong> dentro de .box-title
    $content = [System.Text.RegularExpressions.Regex]::Replace(
      $content,
      '(<div[^>]*class\s*=\s*["\'']box-title["\''][^>]*>)\s*.*?<strong>',
      '$1<strong>',
      [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    # Ocultar pseudo-elementos ::before por seguranÃ§a
    if ($content -notmatch 'data-noicons') {
      $noIconsCss = '<style data-noicons> .box-title .icon-img{display:none !important;} .box-title::before{content:"" !important;} pre.code, pre.code.sql { background:#fff !important; color:#000 !important; border-color:#cfcfcf !important; } pre.code.sql .kw, pre.code.sql .type, pre.code.sql .number, pre.code.sql .string, pre.code.sql .comment, pre.code.sql .op { color: inherit !important; font-weight: 400 !important; font-style: normal !important; } </style>'
      $content = [System.Text.RegularExpressions.Regex]::Replace(
        $content,
        '</head>',
        $noIconsCss + '</head>'
      )
    }
    Add-ZipEntryContent -Archive $zip -EntryPath $entryPath -Content $content
  } elseif ($NoIcons -and ($href -match '\.css$')) {
    # Sanitizar CSS para remover qualquer conteÃºdo inserido via ::before em tÃ­tulos
    try {
      $css = Get-Content -LiteralPath $fullPath -Raw -Encoding UTF8
    } catch {
      try { $css = [System.IO.File]::ReadAllText($fullPath) } catch { $css = '' }
    }
    if ($null -eq $css) { $css = '' }

    # Substituir a propriedade content dentro de blocos de .<classe> .box-title::before
    $css = [System.Text.RegularExpressions.Regex]::Replace(
      $css,
      '(\.[A-Za-z-]+\s+\.box-title::before\s*\{[^}]*?)content\s*:\s*[^;]+;?([^}]*\})',
      '$1content: none !important;$2',
      [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    # Fallback para seletores que usam apenas .box-title::before
    $css = [System.Text.RegularExpressions.Regex]::Replace(
      $css,
      '(\.box-title::before\s*\{[^}]*?)content\s*:\s*[^;]+;?([^}]*\})',
      '$1content: none !important;$2',
      [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    # Garantir regra global de seguranÃ§a
    if ($css -notmatch '\.box-title::before\s*\{\s*content\s*:\s*none') {
      $css += "`n.box-title::before{content: none !important;}" + "`n.box-title .icon-img{display:none !important;}" + "`n"
    }

    Add-ZipEntryContent -Archive $zip -EntryPath $entryPath -Content $css
  } else {
    Add-ZipEntry -Archive $zip -FullPath $fullPath -EntryPath $entryPath
  }
}


# Adicionar imagens comprimidas da pasta kindle ao EPUB (se UseCompressedJpg)
if ($UseCompressedJpg) {
  $kindleDir = Join-Path $oebpsBase 'images\kindle'
  if (Test-Path $kindleDir) {
    $kindleFiles = Get-ChildItem $kindleDir -Filter "*.jpg" -ErrorAction SilentlyContinue
    foreach ($kf in $kindleFiles) {
      $entryPath = "OEBPS/images/kindle/$($kf.Name)"
      Add-ZipEntry -Archive $zip -FullPath $kf.FullName -EntryPath $entryPath
    }
    Write-Host "Adicionadas $($kindleFiles.Count) imagens comprimidas da pasta kindle"
  }
}
# Fechar ZIP
$zip.Dispose()
$zipStream.Dispose()

Write-Host "EPUB gerado:" $outPath

