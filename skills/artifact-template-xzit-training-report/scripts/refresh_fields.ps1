param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputDocx,
    [Parameter(Mandatory=$true)][string]$OutputPdf
)
$ErrorActionPreference='Stop'
$sourceFile=(Resolve-Path -LiteralPath $InputDocx).Path
$targetDoc=[IO.Path]::GetFullPath($OutputDocx)
$targetPdf=[IO.Path]::GetFullPath($OutputPdf)
if ($sourceFile -eq $targetDoc -or $sourceFile -eq $targetPdf -or $targetDoc -eq $targetPdf -or (Test-Path -LiteralPath $targetDoc) -or (Test-Path -LiteralPath $targetPdf)) {throw 'Use distinct, new output paths; overwriting is disabled.'}
$wordApp=New-Object -ComObject Word.Application
$wordApp.Visible=$false
$wordApp.DisplayAlerts=0
$reportDoc=$null
try {
    $reportDoc=$wordApp.Documents.Open($sourceFile,$false,$true)
    $reportDoc.Repaginate()
    foreach ($story in $reportDoc.StoryRanges) {$null=$story.Fields.Update()}
    foreach ($toc in $reportDoc.TablesOfContents) {$toc.Update()}
    $reportDoc.Repaginate()
    foreach ($toc in $reportDoc.TablesOfContents) {$toc.UpdatePageNumbers()}
    $reportDoc.SaveAs2($targetDoc,16)
    $reportDoc.ExportAsFixedFormat($targetPdf,17)
    Write-Output "Saved updated DOCX and PDF; inspect all rendered pages."
} finally {
    if ($null -ne $reportDoc) {$reportDoc.Close(0)}
    $wordApp.Quit()
}
