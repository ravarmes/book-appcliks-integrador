param(
  [string]$OEBPS = "c:\GitHub\book-bd\OEBPS"
)

$ErrorActionPreference = 'Stop'

$imgDir = Join-Path $OEBPS 'images'
$backupDir = Join-Path $imgDir '_backup'
if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }

# Coletar referências a imagens (relativas à pasta images)
$referenced = New-Object System.Collections.Generic.HashSet[string]

$filesToScan = @()
$filesToScan += Get-ChildItem (Join-Path $OEBPS 'text') -Filter '*.html'
$filesToScan += Get-ChildItem $OEBPS -Filter 'content.opf'
$filesToScan += Get-ChildItem $OEBPS -Filter 'nav.xhtml'
$filesToScan += Get-ChildItem $OEBPS -Filter 'toc.ncx'
$filesToScan += Get-ChildItem (Join-Path $OEBPS 'css') -Filter '*.css'

foreach ($f in $filesToScan) {
  $content = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
  $matches = [regex]::Matches($content, "images\/([A-Za-z0-9_\-\.\/]+)")
  foreach ($m in $matches) {
    $pathRel = $m.Groups[1].Value
    # Ignorar subpastas kindle/ e _backup/
    if ($pathRel -like "kindle/*" -or $pathRel -like "_backup/*") { continue }
    # Se houver subpastas adicionais, considerar apenas nome do arquivo para a raiz
    $fileName = [System.IO.Path]::GetFileName($pathRel)
    if (![string]::IsNullOrWhiteSpace($fileName)) { $referenced.Add($fileName) | Out-Null }
  }
}

# Mover apenas arquivos na raiz de images (não recursivo)
$rootFiles = Get-ChildItem $imgDir -File
$moved = 0
$kept = 0
foreach ($rf in $rootFiles) {
  if ($referenced.Contains($rf.Name)) {
    $kept++
    continue
  }
  $dest = Join-Path $backupDir $rf.Name
  Move-Item -LiteralPath $rf.FullName -Destination $dest -Force
  $moved++
}

Write-Host "Imagens na raiz de 'images': $($rootFiles.Count)"
Write-Host "Referenciadas e mantidas: $kept"
Write-Host "Movidas para _backup: $moved"
