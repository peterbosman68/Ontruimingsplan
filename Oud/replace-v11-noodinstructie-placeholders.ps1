$ErrorActionPreference = 'Stop'

$path = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v11.docx'

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force

function Get-AllStoryRanges {
    param($docObj)

    $list = New-Object System.Collections.ArrayList
    foreach ($sr in $docObj.StoryRanges) {
        $cur = $sr
        while ($null -ne $cur) {
            [void]$list.Add($cur)
            try {
                $cur = $cur.NextStoryRange
            }
            catch {
                $cur = $null
            }
        }
    }

    return $list
}

function Count-InStories {
    param($docObj, [string]$text)

    $count = 0
    $stories = Get-AllStoryRanges $docObj

    foreach ($story in $stories) {
        $r = $story.Duplicate
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
            $r.End = $story.End
        }
    }

    return $count
}

function Replace-InStories {
    param($docObj, [string]$findText, [string]$replaceText)

    $stories = Get-AllStoryRanges $docObj

    foreach ($story in $stories) {
        $r = $story.Duplicate
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
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    $doc = $word.Documents.Open($path, $false, $false)

    $pairs = @(
        @('<Naam huisartsenpraktijk>', '{naam_huisartsenpraktijk}'),
        @('<Telefoon huisartsenpraktijk>', '{telefoon_huisartsenpraktijk}'),
        @('<Telefoon huisartsenspoedpost>', '{telefoon_huisartsenspoedpost}'),
        @('<Naam verzamelplaats>', '{naam_verzamelplaats}')
    )

    foreach ($p in $pairs) {
        $before = Count-InStories $doc $p[0]
        Replace-InStories $doc $p[0] $p[1]
        $afterOld = Count-InStories $doc $p[0]
        $afterNew = Count-InStories $doc $p[1]

        Write-Output ('{0} | before={1} old_after={2} new_after={3}' -f $p[0], $before, $afterOld, $afterNew)
    }

    $doc.Save()
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
