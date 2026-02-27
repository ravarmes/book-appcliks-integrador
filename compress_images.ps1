param(
  [string]$ImagesDir = "c:\GitHub\book-bd\OEBPS\images",
  [string]$OutputDir = "c:\GitHub\book-bd\OEBPS\images\kindle",
  [int]$MaxWidth = 800,
  [int]$Quality = 75
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# Criar diretorio de saida
if (-not (Test-Path $OutputDir)) {
  New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# Processar apenas arquivos PNG do Capitulo 1 e icones
$pngFiles = Get-ChildItem -Path $ImagesDir -Filter "*.png" -File | Where-Object {
  $_.Name -match "^1-" -or $_.Name -match "^icon-" -or $_.Name -eq "book-bd-capa.png"
}

foreach ($file in $pngFiles) {
  $outPath = Join-Path $OutputDir $file.Name
  
  try {
    $img = [System.Drawing.Image]::FromFile($file.FullName)
    
    # Calcular nova dimensao mantendo proporcao
    $ratio = 1.0
    if ($img.Width -gt $MaxWidth) {
      $ratio = $MaxWidth / $img.Width
    }
    $newWidth = [int]($img.Width * $ratio)
    $newHeight = [int]($img.Height * $ratio)
    
    # Criar bitmap redimensionado
    $newImg = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
    $graphics = [System.Drawing.Graphics]::FromImage($newImg)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.DrawImage($img, 0, 0, $newWidth, $newHeight)
    
    # Configurar qualidade JPEG para compressao
    $jpegCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]$Quality)
    
    # Salvar como JPEG comprimido (menor que PNG)
    $jpgPath = $outPath -replace '\.png$', '.jpg'
    $newImg.Save($jpgPath, $jpegCodec, $encoderParams)
    
    # Tambem salvar como PNG se preferir manter formato
    $newImg.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    
    $originalSize = [math]::Round($file.Length / 1KB, 1)
    $newSize = [math]::Round((Get-Item $outPath).Length / 1KB, 1)
    $jpgSize = [math]::Round((Get-Item $jpgPath).Length / 1KB, 1)
    Write-Host "OK $($file.Name): $originalSize KB -> PNG: $newSize KB | JPG: $jpgSize KB"
    
    $graphics.Dispose()
    $newImg.Dispose()
    $img.Dispose()
  } catch {
    Write-Warning "Erro ao processar $($file.Name): $_"
  }
}

Write-Host ""
Write-Host "Imagens comprimidas salvas em: $OutputDir"
