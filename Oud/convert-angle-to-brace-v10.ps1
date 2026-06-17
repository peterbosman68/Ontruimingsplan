$ErrorActionPreference = 'Stop'

$docPath = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v11.docx'

if (-not (Test-Path -LiteralPath $docPath)) {
    throw 'Document not found.'
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Normalize-TagName {
    param([string]$raw)

    $s = $raw.ToLowerInvariant()
    $s = [System.Text.RegularExpressions.Regex]::Replace($s, '[^a-z0-9]+', '_')
    $s = [System.Text.RegularExpressions.Regex]::Replace($s, '_+', '_')
    $s = $s.Trim('_')

    if ([string]::IsNullOrWhiteSpace($s)) {
        return 'veld'
    }

    return $s
}

$tmpDir = Join-Path $env:TEMP ('docx-edit-' + [guid]::NewGuid().ToString())
[System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($docPath, $tmpDir)

    $wordDir = Join-Path $tmpDir 'word'
    $targets = Get-ChildItem -LiteralPath $wordDir -File -Filter '*.xml' | Where-Object {
        $_.Name -match '^(document|header\d+|footer\d+)\.xml$'
    }

    $totalReplacements = 0

    foreach ($file in $targets) {
        [xml]$xml = Get-Content -LiteralPath $file.FullName -Encoding UTF8
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

        $textNodes = $xml.SelectNodes('//w:t', $ns)
        foreach ($node in $textNodes) {
            $original = $node.InnerText
            if ([string]::IsNullOrEmpty($original)) {
                continue
            }

            $updated = [System.Text.RegularExpressions.Regex]::Replace(
                $original,
                '<([^<>]{1,140})>',
                {
                    param($m)
                    $script:totalReplacements++
                    '{' + (Normalize-TagName $m.Groups[1].Value) + '}'
                }
            )

            if ($updated -ne $original) {
                $node.InnerText = $updated
            }
        }

        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $settings.Indent = $false
        $writer = [System.Xml.XmlWriter]::Create($file.FullName, $settings)
        $xml.WriteTo($writer)
        $writer.Close()
    }

    $tmpDoc = $docPath -replace '\.docx$', '.tmp.docx'
    if (Test-Path -LiteralPath $tmpDoc) {
        Remove-Item -LiteralPath $tmpDoc -Force
    }

    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmpDir, $tmpDoc)

    Move-Item -LiteralPath $tmpDoc -Destination $docPath -Force

    Write-Output ('Replacements=' + $totalReplacements)
}
finally {
    if (Test-Path -LiteralPath $tmpDir) {
        Remove-Item -LiteralPath $tmpDir -Recurse -Force
    }
}
