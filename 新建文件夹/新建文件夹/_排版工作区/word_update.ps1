param(
  [Parameter(Mandatory=$true)][string]$Docx,
  [Parameter(Mandatory=$true)][string]$Pdf
)
$ErrorActionPreference = "Stop"
$word = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.Options.UpdateFieldsAtPrint = $true
  $doc = $word.Documents.Open($Docx, $false, $false)
  # 全量更新域（含 TOC / PAGE / NUMPAGES）
  $doc.Fields.Update() | Out-Null
  foreach ($f in $doc.Fields) { try { $f.Update() | Out-Null } catch {} }
  foreach ($toc in $doc.TablesOfContents) { try { $toc.Update() | Out-Null } catch {} }
  $doc.Repaginate()
  Start-Sleep -Milliseconds 500
  $pages = $doc.ComputeStatistics(2)   # wdStatisticPages
  $words = $doc.ComputeStatistics(0)   # wdStatisticWords
  $doc.Save()
  $doc.ExportAsFixedFormat($Pdf, 17)   # wdFormatPDF
  $doc.Close($true)
  Set-Content -Path "$Pdf.result.txt" -Value "OK pages=$pages words=$words" -Encoding UTF8
} finally {
  if ($word -ne $null) { $word.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
}
