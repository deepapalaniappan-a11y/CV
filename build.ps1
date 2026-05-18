# One command. Run this after editing MASTER-CV.md.
#   Right-click build.ps1 > Run with PowerShell, or in a terminal:  ./build.ps1
# It regenerates the Word CV, the PDF, the JSON Resume file and the web CV.

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

Write-Host 'Building CV from MASTER-CV.md ...' -ForegroundColor Cyan
python build_cv.py
if ($LASTEXITCODE -ne 0) { throw 'build_cv.py failed. Fix the error above and re-run.' }

$docx = Join-Path $here 'Deepa_Palaniappan_CV_2026.docx'
$pdf  = Join-Path $here 'Deepa_Palaniappan_CV_2026.pdf'

Write-Host 'Exporting PDF ...' -ForegroundColor Cyan
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $d = $word.Documents.Open($docx, $false, $true)
    $pages = $d.ComputeStatistics(2)
    $d.ExportAsFixedFormat($pdf, 17, $false, 0)
    $d.Close($false)
    Copy-Item $pdf (Join-Path $here 'docs\Deepa_Palaniappan_CV_2026.pdf') -Force
    Write-Host ("Done. CV is {0} pages." -f $pages) -ForegroundColor Green
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

Write-Host ''
Write-Host 'Updated files:' -ForegroundColor Green
Write-Host '  Deepa_Palaniappan_CV_2026.docx        (Word, editable backup)'
Write-Host '  Deepa_Palaniappan_CV_2026.pdf         (send this to applications)'
Write-Host '  docs\index.html + docs\resume.json    (public web CV)'
Write-Host '  docs\Deepa_Palaniappan_CV_2026.pdf    (PDF served by the web CV)'
Write-Host ''
Write-Host 'To publish: open GitHub Desktop, write a short summary, Commit, then Push.' -ForegroundColor Cyan
