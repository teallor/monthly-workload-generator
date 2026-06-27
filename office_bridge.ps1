param(
    [Parameter(Mandatory = $true)][ValidateSet('inspect', 'convert-doc', 'extract-excel', 'write')][string]$Action,
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Output,
    [string]$Payload
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new()
$OutputEncoding = [Console]::OutputEncoding
$labelDate = -join [char[]](0x65E5, 0x671F)
$labelCourse = -join [char[]](0x8BFE, 0x7A0B, 0x540D, 0x79F0)
$labelClass = -join [char[]](0x73ED, 0x7EA7)
$titleMarker = -join [char[]](0x6708, 0x5EA6, 0x5DE5, 0x4F5C, 0x91CF, 0x7EDF, 0x8BA1)

if ($Action -eq 'convert-doc') {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        $doc = $word.Documents.Open([IO.Path]::GetFullPath($Path), $false, $true)
        try { $doc.SaveAs2([IO.Path]::GetFullPath($Output), 16) }
        finally { $doc.Close(0) }
    }
    finally { try { $word.Quit() } catch { } }
    exit 0
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    if ($Action -eq 'extract-excel') {
        $workbook = $excel.Workbooks.Open([IO.Path]::GetFullPath($Path), 0, $true)
        try {
            $sheets = @()
            foreach ($sheet in $workbook.Worksheets) {
                $used = $sheet.UsedRange
                $rowCount = [Math]::Min([int]$used.Rows.Count, 500)
                $colCount = [Math]::Min([int]$used.Columns.Count, 80)
                $rows = @()
                for ($r = 0; $r -lt $rowCount; $r++) {
                    $values = @()
                    for ($c = 0; $c -lt $colCount; $c++) {
                        $values += [string]$sheet.Cells.Item($used.Row + $r, $used.Column + $c).Text
                    }
                    $rows += ,$values
                }
                $sheets += [ordered]@{ name = [string]$sheet.Name; rows = $rows }
            }
            $json = ConvertTo-Json $sheets -Depth 8 -Compress
            [IO.File]::WriteAllText([IO.Path]::GetFullPath($Output), $json, [Text.UTF8Encoding]::new($true))
        }
        finally { $workbook.Close($false) }
        exit 0
    }

    if ($Action -eq 'inspect') {
        $workbook = $excel.Workbooks.Open([IO.Path]::GetFullPath($Path), 0, $true)
        try {
            $sheet = $null
            $headerRow = 0
            $dateCol = 0
            $courseCol = 0
            $assessmentCol = 0
            $assessmentEndCol = 0
            foreach ($candidate in $workbook.Worksheets) {
                $used = $candidate.UsedRange
                $maxRows = [Math]::Min([int]($used.Row + $used.Rows.Count - 1), 100)
                $maxCols = [Math]::Min([int]($used.Column + $used.Columns.Count - 1), 60)
                for ($r = [int]$used.Row; $r -le $maxRows -and $headerRow -eq 0; $r++) {
                    $foundDateCol = 0
                    $foundCourseCol = 0
                    $foundAssessmentCol = 0
                    for ($c = [int]$used.Column; $c -le $maxCols; $c++) {
                        $value = ([string]$candidate.Cells.Item($r, $c).Text).Trim()
                        if ($value -eq $labelDate) { $foundDateCol = $c }
                        if ($value -eq $labelCourse) { $foundCourseCol = $c }
                        if ($value -eq $labelClass -and $foundCourseCol -gt 0 -and $c -gt $foundCourseCol) { $foundAssessmentCol = $c }
                    }
                    if ($foundDateCol -gt 0 -and $foundCourseCol -gt 0 -and $foundAssessmentCol -gt 0) {
                        $sheet = $candidate
                        $headerRow = $r
                        $dateCol = $foundDateCol
                        $courseCol = $foundCourseCol
                        $assessmentCol = $foundAssessmentCol
                        $assessmentEndCol = $assessmentCol + 2
                    }
                }
                if ($headerRow -gt 0) { break }
            }
            if ($null -eq $sheet) { throw 'Teaching header not found' }
            $titleCell = $null
            $used = $sheet.UsedRange
            for ($r = [int]$used.Row; $r -le [Math]::Min([int]($used.Row + $used.Rows.Count - 1), $headerRow); $r++) {
                for ($c = [int]$used.Column; $c -le [Math]::Min([int]($used.Column + $used.Columns.Count - 1), 60); $c++) {
                    $cell = $sheet.Cells.Item($r, $c)
                    if ([string]$cell.Text -like "*${titleMarker}*") { $titleCell = $cell; break }
                }
                if ($null -ne $titleCell) { break }
            }
            if ($null -eq $titleCell) { throw 'Month title cell not found' }
            $dataStart = $headerRow + 1
            $nameCell = $sheet.Cells.Item($dataStart, 1)
            $dataEnd = $dataStart
            if ($nameCell.MergeCells) { $dataEnd = [int]($nameCell.MergeArea.Row + $nameCell.MergeArea.Rows.Count - 1) }
            $trainingEndCol = $assessmentCol - 1
            $assessmentSlots = @()
            $seenSlotRows = @{}
            for ($slotRow = $dataStart; $slotRow -le $dataEnd; $slotRow++) {
                $slotCell = $sheet.Cells.Item($slotRow, $assessmentCol)
                $slotStart = $slotRow
                $slotEnd = $slotRow
                if ($slotCell.MergeCells) {
                    $slotStart = [int]$slotCell.MergeArea.Row
                    $slotEnd = [int]($slotCell.MergeArea.Row + $slotCell.MergeArea.Rows.Count - 1)
                }
                if (-not $seenSlotRows.ContainsKey($slotStart)) {
                    $seenSlotRows[$slotStart] = $true
                    $assessmentSlots += [ordered]@{ start_row = $slotStart; end_row = $slotEnd }
                }
            }
            $result = [ordered]@{
                sheet = [string]$sheet.Name
                title = [string]$titleCell.Value2
                title_cell = [string]$titleCell.Address($false, $false)
                header_row = $headerRow
                data_start_row = $dataStart
                data_end_row = $dataEnd
                training_start_col = $dateCol
                training_end_col = $trainingEndCol
                assessment_start_col = $assessmentCol
                assessment_end_col = $assessmentEndCol
                assessment_slots = $assessmentSlots
            }
            $result | ConvertTo-Json -Compress
        }
        finally { $workbook.Close($false) }
        exit 0
    }

    if ($Action -eq 'write') {
        Copy-Item -LiteralPath $Path -Destination $Output -Force
        $data = Get-Content -Raw -LiteralPath $Payload | ConvertFrom-Json
        $workbook = $excel.Workbooks.Open([IO.Path]::GetFullPath($Output))
        try {
            $layout = $data.layout
            $sheet = $workbook.Worksheets.Item([string]$layout.sheet)
            $sheet.Range([string]$layout.title_cell).Value2 = [string]$data.title
            $startRow = [int]$layout.data_start_row
            $endRow = [int]$layout.data_end_row
            $trainingStart = [int]$layout.training_start_col
            $trainingEnd = [int]$layout.training_end_col
            $assessmentStart = [int]$layout.assessment_start_col
            $assessmentEnd = [int]$layout.assessment_end_col
            $trainingRecords = @($data.records | Where-Object { [bool]$_.is_training })
            $assessmentRecords = @($data.records | Where-Object { -not [bool]$_.is_training })
            $trainingRange = $sheet.Range($sheet.Cells.Item($startRow, $trainingStart), $sheet.Cells.Item($endRow, $trainingEnd))
            $assessmentRange = $sheet.Range($sheet.Cells.Item($startRow, $assessmentStart), $sheet.Cells.Item($endRow, $assessmentEnd))
            $trainingRange.ClearContents()
            $assessmentRange.ClearContents()
            if ($trainingRecords.Count -gt 1) {
                $trainingRange.UnMerge()
                for ($row = $startRow + 1; $row -le $endRow; $row++) {
                    $sheet.Range($sheet.Cells.Item($startRow, $trainingStart), $sheet.Cells.Item($startRow, $trainingEnd)).Copy()
                    $sheet.Range($sheet.Cells.Item($row, $trainingStart), $sheet.Cells.Item($row, $trainingEnd)).PasteSpecial(-4122)
                }
            }
            try { $excel.CutCopyMode = $false } catch { }
            $trainingRow = $startRow
            $assessmentIndex = 0
            foreach ($record in $data.records) {
                if ([bool]$record.is_training) {
                    $sheet.Cells.Item($trainingRow, $trainingStart).Value2 = ([string]$record.date).Replace('-', '/')
                    $sheet.Cells.Item($trainingRow, $trainingStart + 1).Value2 = [string]$record.sheet_class_name
                    $sheet.Cells.Item($trainingRow, $trainingStart + 2).Value2 = [string]$record.audience
                    $sheet.Cells.Item($trainingRow, $trainingStart + 3).Value2 = [string]$record.course_name
                    $sheet.Cells.Item($trainingRow, $trainingStart + 4).FormulaR1C1 = [string]$record.hours
                    $sheet.Cells.Item($trainingRow, $trainingStart + 5).FormulaR1C1 = $(if ([bool]$record.is_live) {'1'} else {'0'})
                    $sheet.Cells.Item($trainingRow, $trainingStart + 6).FormulaR1C1 = $(if ([bool]$record.is_delivery) {'1'} else {'0'})
                    $sheet.Cells.Item($trainingRow, $trainingStart + 7).FormulaR1C1 = $(if ([bool]$record.is_delivery) {[string]$record.hours} else {'0'})
                    $trainingRow++
                }
                else {
                    $assessmentRow = [int]$layout.assessment_slots[$assessmentIndex].start_row
                    $sheet.Cells.Item($assessmentRow, $assessmentStart).Value2 = [string]$record.sheet_class_name
                    $sheet.Cells.Item($assessmentRow, $assessmentStart + 1).FormulaR1C1 = [string]$record.hours
                    $sheet.Cells.Item($assessmentRow, $assessmentStart + 2).FormulaR1C1 = $(if ([bool]$record.is_live) {'1'} else {'0'})
                    $assessmentIndex++
                }
            }
            $workbook.Save()
        }
        finally { $workbook.Close($true) }
    }
}
finally { try { $excel.Quit() } catch { } }
