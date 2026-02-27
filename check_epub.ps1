Param(
    [string]$EpubPath = "c:\GitHub\book-poo2\Livro_POO2_kindle.epub"
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not (Test-Path -LiteralPath $EpubPath)) {
    Write-Error "Arquivo EPUB não encontrado: $EpubPath"
    exit 1
}

$fs = [System.IO.File]::OpenRead($EpubPath)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)

    $first = $zip.Entries[0].FullName
    $entriesCount = $zip.Entries.Count
    Write-Output ("FirstEntry: " + $first)
    Write-Output ("EntriesCount: " + $entriesCount)
    # Verificar conteúdo do mimetype
    $mimeEntry = $zip.Entries | Where-Object { $_.FullName -eq 'mimetype' }
    if ($mimeEntry) {
        $ms = $mimeEntry.Open(); try {
            $sr = New-Object System.IO.StreamReader($ms)
            $mimeText = $sr.ReadToEnd()
            Write-Output ("MimetypeCorrect: " + ($mimeText -eq 'application/epub+zip'))
        } finally { $ms.Dispose() }
    }

    $container = $zip.Entries | Where-Object { $_.FullName -eq "META-INF/container.xml" }
    Write-Output ("ContainerXMLFound: " + ([bool]$container))

    $nav = $zip.Entries | Where-Object { $_.FullName -eq "OEBPS/nav.xhtml" }
    Write-Output ("NavXHTMLFound: " + ([bool]$nav))

    $opf = $zip.Entries | Where-Object { $_.FullName -eq "OEBPS/content.opf" }
    Write-Output ("OPFFound: " + ([bool]$opf))
    
    # Analisar OPF para capa e SVGs
    if ($opf) {
        $tmp = New-Object System.IO.StreamReader($opf.Open())
        $opfXmlText = $tmp.ReadToEnd(); $tmp.Dispose()
        $x = New-Object System.Xml.XmlDocument
        try { $x.LoadXml($opfXmlText) } catch {}
        if ($x.DocumentElement) {
            $n = New-Object System.Xml.XmlNamespaceManager($x.NameTable)
            $n.AddNamespace('opf','http://www.idpf.org/2007/opf')
            $items = $x.SelectNodes('//opf:manifest/opf:item', $n)
            $hrefs = @(); foreach ($it in $items) { $hrefs += $it.GetAttribute('href') }
            $svgCount = ($hrefs | Where-Object { $_ -match '\.svg$' }).Count
            Write-Output ("ManifestItems: " + $items.Count)
            Write-Output ("ManifestSVGs: " + $svgCount)
            $coverHtml = $items | Where-Object { $_.GetAttribute('href') -match 'text/capa.*\.(x?html)$' }
            $coverImg  = $items | Where-Object { $_.GetAttribute('href') -match 'images/.*capa.*\.(png|jpg|jpeg)$' }
            Write-Output ("CoverHTMLInManifest: " + ([bool]$coverHtml))
            Write-Output ("CoverIMGInManifest: " + ([bool]$coverImg))
            $spine = $x.SelectNodes('//opf:spine/opf:itemref', $n)
            $spineHasCover = $false
            foreach ($ref in $spine) {
                if ($coverHtml) { if ($ref.GetAttribute('idref') -eq $coverHtml.GetAttribute('id')) { $spineHasCover = $true } }
            }
            Write-Output ("SpineHasCover: " + $spineHasCover)
            $metaCover = $x.SelectSingleNode('//opf:metadata/opf:meta[@name="cover"]', $n)
            Write-Output ("MetaCoverPresent: " + ([bool]$metaCover))
        }
    }

    Write-Output "Entries:"
    foreach ($e in $zip.Entries) {
        Write-Output (" - " + $e.FullName)
    }
    # Detectar entradas duplicadas por caminho
    $dups = $zip.Entries | Group-Object FullName | Where-Object { $_.Count -gt 1 }
    if ($dups) {
        Write-Output "DuplicateEntriesDetected: True"
        foreach ($d in $dups) { Write-Output (" Duplicated: " + $d.Name + " x" + $d.Count) }
    } else {
        Write-Output "DuplicateEntriesDetected: False"
    }

} finally {
    if ($zip) { $zip.Dispose() }
    $fs.Dispose()
}
