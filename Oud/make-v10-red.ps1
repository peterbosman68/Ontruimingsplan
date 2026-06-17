$ErrorActionPreference = 'Stop'

$source = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v9.docx'
$dest = $source -replace 'v9\.docx', 'v10.docx'

if (-not (Test-Path -LiteralPath $source)) {
    Write-Output 'Source file not found'
    exit
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$tempDir = [System.IO.Path]::GetTempPath() + [System.Guid]::NewGuid().ToString()
[System.IO.Directory]::CreateDirectory($tempDir) | Out-Null

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($source, $tempDir)

    $docXmlPath = Join-Path $tempDir 'word\document.xml'
    $docContent = [System.IO.File]::ReadAllText($docXmlPath, [System.Text.Encoding]::UTF8)

    $coloredCount = 0

    $pattern = '<w:color w:val="([A-Fa-f0-9]{6})"\s*/>'
    $docContent = [System.Text.RegularExpressions.Regex]::Replace($docContent, $pattern, {
        param($m)
        if ($m.Groups[1].Value.ToUpper() -ne 'FF0000') {
            $coloredCount++
            return '<w:color w:val="FF0000"/>'
        }
        return $m.Groups[0].Value
    })

    [System.IO.File]::WriteAllText($docXmlPath, $docContent, [System.Text.Encoding]::UTF8)

    if (Test-Path -LiteralPath $dest) {
        Remove-Item -LiteralPath $dest -Force
    }

    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $dest)

    Write-Output ('ColorUpdated=' + $coloredCount + ', CreatedV10')
}
finally {
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}
