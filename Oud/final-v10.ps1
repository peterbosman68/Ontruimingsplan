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

    $before = [System.Text.RegularExpressions.Regex]::Matches($docContent, 'EE0000').Count
    
    $docContent = $docContent -replace 'EE0000', 'FF0000'

    $after = [System.Text.RegularExpressions.Regex]::Matches($docContent, 'FF0000').Count

    [System.IO.File]::WriteAllText($docXmlPath, $docContent, [System.Text.Encoding]::UTF8)

    if (Test-Path -LiteralPath $dest) {
        Remove-Item -LiteralPath $dest -Force
    }

    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $dest)

    Write-Output ('DarkRedReplaced=' + $before + ', CreatedV10')
}
finally {
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}
