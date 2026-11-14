param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputDocx,
    [Parameter(Mandatory=$true)][string]$OutputPdf,
    [string]$OutputLayout
)
$ErrorActionPreference='Stop'
$sourceFile=(Resolve-Path -LiteralPath $InputDocx).Path
$targetDoc=[IO.Path]::GetFullPath($OutputDocx)
$targetPdf=[IO.Path]::GetFullPath($OutputPdf)
if ($sourceFile -eq $targetDoc -or $sourceFile -eq $targetPdf -or $targetDoc -eq $targetPdf -or (Test-Path -LiteralPath $targetDoc) -or (Test-Path -LiteralPath $targetPdf)) {throw 'Use distinct, new output paths; overwriting is disabled.'}
if ($OutputLayout) {
    $layoutPath=[IO.Path]::GetFullPath($OutputLayout)
    if ($layoutPath -in @($sourceFile,$targetDoc,$targetPdf) -or (Test-Path -LiteralPath $layoutPath)) {throw 'Layout output must be a distinct new path.'}
}
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
    if ($OutputLayout) {
        $layoutTables=@()
        for ($tableIndex=1; $tableIndex -le $reportDoc.Tables.Count; $tableIndex++) {
            $layoutTable=$reportDoc.Tables.Item($tableIndex)
            # Only generated/mapped body tables are candidates for automatic
            # continuation. Cover or legacy merged tables may not expose Rows.
            if ($layoutTable.Cell(1,1).Range.Paragraphs.Item(1).Style.NameLocal -ne 'XZITTable') {continue}
            $layoutRows=@()
            for ($rowIndex=1; $rowIndex -le $layoutTable.Rows.Count; $rowIndex++) {
                $rowRange=$layoutTable.Rows.Item($rowIndex).Range.Duplicate
                $rowRange.Collapse(1)
                $startPage=$rowRange.Information(3)
                $rowRange=$layoutTable.Rows.Item($rowIndex).Range.Duplicate
                $rowRange.MoveEnd(1,-1) | Out-Null
                $rowRange.Collapse(0)
                $layoutRows+=@{row=$rowIndex;start_page=$startPage;end_page=$rowRange.Information(3)}
            }
            $layoutTables+=@{table_index=$tableIndex;rows=$layoutRows}
        }
        @{page_count=$reportDoc.ComputeStatistics(2);tables=$layoutTables} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $layoutPath -Encoding UTF8
    }
    $reportDoc.SaveAs2($targetDoc,16)
    $reportDoc.ExportAsFixedFormat($targetPdf,17)
    Write-Output "Saved updated DOCX and PDF; inspect all rendered pages."
} finally {
    if ($null -ne $reportDoc) {$reportDoc.Close(0)}
    $wordApp.Quit()
}
