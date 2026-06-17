$ErrorActionPreference = 'Stop'

$path = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v11.docx'

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force

function Replace-InStoryRange {
    param($range, [string]$findText, [string]$replaceText)

    $r = $range.Duplicate
    $f = $r.Find
    $f.ClearFormatting() | Out-Null
    $f.Replacement.ClearFormatting() | Out-Null
    $f.Text = $findText
    $f.Replacement.Text = $replaceText
    $f.MatchWildcards = $false
    $f.MatchCase = $false
    $f.MatchWholeWord = $false
    $f.Forward = $true
    $f.Wrap = 1
    $null = $f.Execute($findText, $false, $false, $false, $false, $false, $true, 1, $false, $replaceText, 2)
}

function Count-InStoryRange {
    param($range, [string]$text)

    $count = 0
    $r = $range.Duplicate
    $f = $r.Find
    $f.ClearFormatting() | Out-Null
    $f.Text = $text
    $f.MatchWildcards = $false
    $f.MatchCase = $false
    $f.MatchWholeWord = $false
    $f.Forward = $true
    $f.Wrap = 0

    while ($f.Execute()) {
        $count++
        $r.Start = $r.End
        $r.End = $range.End
    }

    return $count
}

$medicalLine = 'Werkdagen 08.00-17.00: bel {naam_huisartsenpraktijk} en {telefoon_huisartsenpraktijk}. 17.00-08.00, weekend en feestdagen: bel Huisartsenspoedpost {telefoon_huisartsenspoedpost}. '

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    $doc = $word.Documents.Open($path, $false, $false)

    foreach ($story in $doc.StoryRanges) {
        Replace-InStoryRange $story 'ga bij ontruimingsalarm naar de verzamel{plaats}' ('ga bij ontruimingsalarm naar {naam_verzamelplaats}')
        Replace-InStoryRange $story 'ga bij ontruimingsalarm naar de verzamelplaats' ('ga bij ontruimingsalarm naar {naam_verzamelplaats}')
        Replace-InStoryRange $story 'ga bij ontruimingsalarm naar {naam_verzamelplaats}' ($medicalLine + 'ga bij ontruimingsalarm naar {naam_verzamelplaats}')
        Replace-InStoryRange $story 'verzamel{plaats}' '{naam_verzamelplaats}'
    }

    $doc.Save()

    $c1 = 0
    $c2 = 0
    $c3 = 0
    $c4 = 0

    foreach ($story in $doc.StoryRanges) {
        $c1 += Count-InStoryRange $story '{naam_huisartsenpraktijk}'
        $c2 += Count-InStoryRange $story '{telefoon_huisartsenpraktijk}'
        $c3 += Count-InStoryRange $story '{telefoon_huisartsenspoedpost}'
        $c4 += Count-InStoryRange $story '{naam_verzamelplaats}'
    }

    Write-Output ('naam_huisartsenpraktijk=' + $c1)
    Write-Output ('telefoon_huisartsenpraktijk=' + $c2)
    Write-Output ('telefoon_huisartsenspoedpost=' + $c3)
    Write-Output ('naam_verzamelplaats=' + $c4)
    Write-Output 'Saved v11'
}
finally {
    if ($doc) { $doc.Close([ref]0) | Out-Null }
    if ($word) { $word.Quit() }
    if ($doc) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
    if ($word) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
