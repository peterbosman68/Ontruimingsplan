$ErrorActionPreference = 'Stop'

$path = 'c:\Users\peter\OneDrive\5. NieuweBedrijven\WebsitesBouwen\GBBCtolkamer\Ontruimingsplan\OntruimingsplanConformNEN8112 260610.v11.docx'

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
        if ($txt -match 'Noodinstructie|ontruimingsalarm|Werkdagen|huisarts|Huisartsenspoedpost|bel') {
            Write-Output ('--- STORY ' + $i + ' ---')
            if ($txt.Length -gt 5000) {
                $txt = $txt.Substring(0, 5000)
            }
            Write-Output $txt
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
