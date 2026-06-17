$ErrorActionPreference = 'Stop'

$path = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v11.docx'
$terms = @('Werkdagen', 'huisartsenpraktijk', 'Huisartsenspoedpost', 'ontruimingsalarm', '{naam_verzamelplaats}', 'Bij ontruimingsalarm')

Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$d = $null

try {
    $d = $w.Documents.Open($path, $false, $true)

    $i = 0
    foreach ($sr in $d.StoryRanges) {
        $i++
        $txt = $sr.Text
        foreach ($term in $terms) {
            $idx = $txt.IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase)
            if ($idx -ge 0) {
                $start = [Math]::Max(0, $idx - 140)
                $len = [Math]::Min(360, $txt.Length - $start)
                $snippet = $txt.Substring($start, $len)
                Write-Output ('--- STORY ' + $i + ' TERM ' + $term + ' ---')
                Write-Output $snippet
            }
        }
    }
}
finally {
    if ($d) { $d.Close([ref]0) | Out-Null }
    $w.Quit()
    if ($d) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($d) | Out-Null }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
}
