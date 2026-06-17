$ErrorActionPreference = 'Stop'

$path = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v9.docx'
$backup = $path -replace '\.docx$', '.v9-backup.docx'

if (-not (Test-Path -LiteralPath $path)) {
    Write-Output 'File not found'
    exit
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$tempDir = [System.IO.Path]::GetTempPath() + [System.Guid]::NewGuid().ToString()
[System.IO.Directory]::CreateDirectory($tempDir) | Out-Null

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($path, $tempDir)

    $docXmlPath = Join-Path $tempDir 'word\document.xml'
    [xml]$docXml = Get-Content -LiteralPath $docXmlPath -Encoding UTF8

    $nsManager = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
    $nsManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

    $tNodes = $docXml.SelectNodes('//w:t', $nsManager)
    $coloredCount = 0

    foreach ($tNode in $tNodes) {
        $text = $tNode.InnerText
        if ($text -match '^\{[^}]+\}$') {
            $pNode = $tNode.SelectSingleNode('ancestor::w:p', $nsManager)
            if ($pNode) {
                $rPrNode = $pNode.SelectSingleNode('w:pPr/w:rPr', $nsManager)
                if (-not $rPrNode) {
                    $pPrNode = $pNode.SelectSingleNode('w:pPr', $nsManager)
                    if (-not $pPrNode) {
                        $pPrNode = $docXml.CreateElement('w:pPr', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                        $pNode.InsertBefore($pPrNode, $pNode.FirstChild) | Out-Null
                    }
                    $rPrNode = $docXml.CreateElement('w:rPr', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                    $pPrNode.AppendChild($rPrNode) | Out-Null
                }

                $colorNode = $rPrNode.SelectSingleNode('w:color', $nsManager)
                if (-not $colorNode) {
                    $colorNode = $docXml.CreateElement('w:color', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                    $rPrNode.AppendChild($colorNode) | Out-Null
                }
                $colorAttr = $docXml.CreateAttribute('w', 'val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                $colorAttr.Value = 'FF0000'
                $colorNode.SetAttributeNode($colorAttr) | Out-Null
                $coloredCount++
            }
        }
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = [System.Text.Encoding]::UTF8
    $settings.Indent = $false
    $writer = [System.Xml.XmlWriter]::Create($docXmlPath, $settings)
    $docXml.WriteTo($writer)
    $writer.Close()

    $tempZip = $path -replace '\.docx$', '.temp.docx'
    if (Test-Path -LiteralPath $tempZip) {
        Remove-Item -LiteralPath $tempZip -Force
    }

    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $tempZip)

    if (Test-Path -LiteralPath $backup) {
        Remove-Item -LiteralPath $backup -Force
    }
    Rename-Item -LiteralPath $path -NewName $backup

    Rename-Item -LiteralPath $tempZip -NewName $path

    Write-Output ('TagsColorized=' + $coloredCount)
}
finally {
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}
