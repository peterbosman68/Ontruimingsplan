$ErrorActionPreference = 'Stop'

$path = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v9.docx'

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($path, $false, $false)

    $range = $doc.Content
    $find = $range.Find

    $find.ClearFormatting() | Out-Null
    $find.Text = '{*}'
    $find.MatchWildcards = $true
    $find.Forward = $true
    $find.Wrap = 2

    $coloredCount = 0

    while ($find.Execute()) {
        $range.Font.Color = 255
        $coloredCount++
        $range.Start = $range.End
        $range.End = $doc.Content.End
    }

    $doc.Save()

    Write-Output ('TagsColored=' + $coloredCount)
}
finally {
    if ($doc) { $doc.Close([ref]0) }
    if ($word) { $word.Quit() }
    if ($doc) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null }
    if ($word) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
