$ErrorActionPreference = 'Stop'

$source = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v10.docx'
$dest = $source

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

    $before = [System.Text.RegularExpressions.Regex]::Matches($docContent, 'FF0000').Count
    
    $docContent = $docContent -replace 'FF0000', '000000'

    [System.IO.File]::WriteAllText($docXmlPath, $docContent, [System.Text.Encoding]::UTF8)

    $tempZip = $dest -replace '\.docx$', '.temp.docx'
    if (Test-Path -LiteralPath $tempZip) {
        Remove-Item -LiteralPath $tempZip -Force
    }

    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $tempZip)

    Remove-Item -LiteralPath $dest -Force
    Rename-Item -LiteralPath $tempZip -NewName $dest

    Write-Output ('BlackColorApplied=' + $before)
}
finally {
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}
