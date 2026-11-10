param([switch]$WriteFormattedSamples)

$ErrorActionPreference = 'Stop'
$excel = $null
$testFolder = Join-Path ([IO.Path]::GetTempPath()) ('comos-formatter-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testFolder)

function Assert-True($Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Get-WorkbookHash([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { [BitConverter]::ToString($hasher.ComputeHash($stream)) }
    finally { $stream.Dispose(); $hasher.Dispose() }
}

function Invoke-Formatter($Book) {
    $excel.Run("'$($hostBook.Name)'!ComosExportFormatter.FormatComosWorkbook", $Book)
}

function Assert-Layout($Book, [int]$Rows, [int]$Columns) {
    $sheet = $Book.Worksheets.Item('Query')
    Assert-True ($sheet.Range('A1').Value2 -ceq 'Line No.') 'Header is not A1.'
    Assert-True (-not $sheet.Rows.Item(1).Hidden) 'Header is hidden.'
    Assert-True ($sheet.Range('A1').Font.Bold) 'Header is not bold.'
    Assert-True ($sheet.AutoFilter.Range.Rows.Count -eq $Rows) 'Filter row count is incorrect.'
    Assert-True ($sheet.AutoFilter.Range.Columns.Count -eq $Columns) 'Filter column count is incorrect.'
    Assert-True ($sheet.Range('A1').WrapText) 'Wrap Text is missing.'
    Assert-True ($sheet.Range('A1').VerticalAlignment -eq -4108) 'Vertical center is missing.'
    Assert-True ($sheet.Range('A1').Interior.Pattern -eq -4142) 'Header still has a fill.'
    Assert-True ($sheet.Cells.Item($Rows, $Columns).Borders.Item(10).Weight -eq 2) 'Final-cell border is not thin.'
    Assert-True ($sheet.Cells.Item($Rows, $Columns).Borders.Item(10).Color -eq 0) 'Border is not black.'
    Assert-True ($sheet.Cells.Item($Rows + 2, $Columns + 2).Borders.Item(7).LineStyle -eq -4142) 'Outside border remains.'
    Assert-True ($sheet.Range('A1').Borders.Item(5).LineStyle -eq -4142) 'Diagonal border remains.'
    Assert-True ($Book.Windows.Item(1).View -eq 1) 'Workbook did not switch to Normal view.'
    Assert-True ($excel.ActiveWindow.FreezePanes) 'Freeze Panes is missing.'
    Assert-True ($excel.ActiveWindow.SplitRow -eq 1) 'Wrong frozen row.'
    Assert-True ($excel.ActiveWindow.SplitColumn -eq 0) 'An unwanted column is frozen.'
    foreach ($worksheet in $Book.Worksheets) {
        Assert-True (-not $worksheet.ProtectContents) 'A worksheet is still protected.'
    }
}

function Assert-Rejected($Book, [string]$Expected) {
    $rejected = $false
    try { [void](Invoke-Formatter $Book) }
    catch {
        $rejected = $true
        Assert-True ($_.Exception.ToString().Contains($Expected)) "Unexpected error: $($_.Exception.Message)"
    }
    Assert-True $rejected "Expected rejection: $Expected"
    Assert-True ($excel.EnableEvents -and $excel.ScreenUpdating) 'Excel state was not restored after an error.'
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $hostBook = $excel.Workbooks.Add()
    [void]$hostBook.VBProject.VBComponents.Import((Join-Path $PSScriptRoot 'ComosExportFormatter.bas'))
    $hostBook.SaveAs((Join-Path $testFolder 'FormatterHarness.xlsm'), 52)

    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Query'
    $sheet.Range('A1:J4').NumberFormat = '@'
    $headers = @('SystemUID', 'Line No.', 'Decimal', 'Internal', 'PID No.', 'Multiple PIDs', 'Formula', 'Precision', 'Other', 'Metadata')
    $values = @('internal', '00123', '12.345', 'KEEP', '90222243', '90222243 / 90222272', 'formula', '1234567890123456', ' -0.50 ', 'DELETE')
    for ($index = 0; $index -lt $headers.Count; $index++) {
        $sheet.Cells.Item(2, $index + 1).Value2 = $headers[$index]
        $sheet.Cells.Item(3, $index + 1).Value2 = $values[$index]
    }
    $sheet.Range('B4').Value2 = '000'
    $sheet.Range('C4').Value2 = '12.340'
    $sheet.Range('I4').Value2 = '1,23'
    $sheet.Range('G3').NumberFormat = 'General'
    $sheet.Range('G3').Formula = '=1+2'
    $sheet.Rows.Item(1).Hidden = $true
    $sheet.Columns.Item(4).Hidden = $true
    $sheet.Columns.Item(10).Hidden = $true
    $sheet.Range('A1:M12').Borders.LineStyle = 1
    $sheet.Range('A1:M12').Interior.Color = 255
    $sheet.Protect()
    $cover = $book.Worksheets.Add()
    $cover.Name = 'Query_CSI'
    $cover.Range('A1').Value2 = 'Unchanged cover'
    $cover.Range('A1').Interior.Color = 255
    $cover.Protect()
    $source = Join-Path $testFolder 'synthetic.xlsx'
    $book.SaveAs($source, 51)
    $originalHash = Get-WorkbookHash $source
    $originalCalculation = $excel.Calculation
    $sheet.Activate()
    $book.Windows.Item(1).View = 3
    Write-Output (Invoke-Formatter $book)
    Assert-Layout $book 3 8
    Assert-True ($sheet.Range('A2').Value2 -eq 123 -and $sheet.Range('A2').Value2 -isnot [string]) 'Leading-zero conversion failed.'
    Assert-True ($sheet.Range('A3').Value2 -eq 0 -and $sheet.Range('A3').Value2 -isnot [string]) 'All-zero conversion failed.'
    Assert-True ($sheet.Range('B2').Value2 -eq 12.345 -and $sheet.Range('B2').NumberFormat -eq '0.000') 'Three-decimal conversion failed.'
    Assert-True ($sheet.Range('B3').Value2 -eq 12.34 -and $sheet.Range('B3').NumberFormat -eq '0.000') 'Trailing decimal zero was lost.'
    Assert-True ($sheet.Range('C2').Value2 -eq 'KEEP' -and $sheet.Columns.Item(3).Hidden) 'Interior hidden column was deleted.'
    Assert-True ($sheet.Range('D2').Value2 -eq 90222243 -and $sheet.Range('D2').Value2 -isnot [string]) 'Single PID conversion failed.'
    Assert-True ($sheet.Range('E2').Value2 -ceq '90222243 / 90222272') 'Slash-separated PIDs changed.'
    Assert-True ($sheet.Range('F2').Formula -eq '=1+2') 'Formula changed.'
    Assert-True ($sheet.Range('G2').Value2 -ceq '1234567890123456') 'Long numeric ID lost precision.'
    Assert-True ($sheet.Range('H2').Value2 -eq -0.5 -and $sheet.Range('H2').NumberFormat -eq '0.00') 'Signed/trimmed decimal failed.'
    Assert-True ($sheet.Range('H3').Value2 -ceq '1,23') 'Comma-separated text changed.'
    Assert-True ($sheet.Range('J8').Borders.Item(7).LineStyle -eq -4142) 'Old border outside the data remains.'
    Assert-True ($sheet.Range('J8').Interior.Pattern -eq -4142) 'Old fill outside the data remains.'
    Assert-True ($cover.Range('A1').Value2 -ceq 'Unchanged cover' -and $cover.Range('A1').Interior.Color -eq 255) 'Cover was reformatted.'
    Assert-True ($excel.Calculation -eq $originalCalculation -and $excel.EnableEvents -and $excel.ScreenUpdating) 'Excel state was not restored.'
    Assert-True ((Get-WorkbookHash $source) -eq $originalHash) 'Original synthetic file changed.'
    $formatted = $book.FullName
    $book.Close($false)
    $book = $excel.Workbooks.Open($formatted)
    Assert-Layout $book 3 8
    Assert-True ($book.Worksheets.Item('Query').Range('B3').NumberFormat -eq '0.000') 'Decimal format did not survive reopening.'
    Write-Output 'PASS: synthetic conversion, hidden columns, formulas, layout, cover, save/reopen, source hash and application settings.'
    Write-Output (Invoke-Formatter $book)
    Assert-Layout $book 3 8
    $book.Close($false)
    Write-Output 'PASS: second formatting run retains the header and data.'

    $book = $excel.Workbooks.Open($source)
    Assert-Rejected $book 'Output already exists'
    Assert-True ($book.Worksheets.Item('Query').Range('B2').Value2 -ceq 'Line No.') 'Rejected run deleted data.'
    $book.Close($false)

    $book = $excel.Workbooks.Add()
    $book.SaveAs((Join-Path $testFolder 'no-query.xlsx'), 51)
    Assert-Rejected $book 'no Query worksheet'
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Query'
    $sheet.Range('A1').Value2 = 'Wrong header'
    Assert-Rejected $book 'Line No. was not found'
    $sheet.Range('A1').Value2 = 'Line No.'
    $sheet.Protect('test-only-password')
    Assert-Rejected $book 'password'
    Assert-True ($sheet.ProtectContents) 'Password-protected sheet changed.'
    $book.Close($false)
    Write-Output 'PASS: existing output, missing Query, missing header and password-protection failures.'

    foreach ($sample in @(
        @{ Name = 'Pipe_list_MED.xlsx'; Rows = 1543; Columns = 28 },
        @{ Name = 'Valve_List_MED.xlsx'; Rows = 1665; Columns = 44 }
    )) {
        $original = Join-Path $PSScriptRoot $sample.Name
        $originalHash = Get-WorkbookHash $original
        $copy = Join-Path $testFolder $sample.Name
        Copy-Item -LiteralPath $original -Destination $copy
        $book = $excel.Workbooks.Open($copy)
        $coverText = $book.Worksheets.Item('Query_CSI').Range('A1').Value2
        $summary = Invoke-Formatter $book
        Write-Output $summary
        Assert-Layout $book $sample.Rows $sample.Columns
        Assert-True ($book.Worksheets.Item('Query_CSI').Range('A1').Value2 -ceq $coverText) 'Sample cover changed.'
        $formatted = $book.FullName
        $book.Close($false)
        $book = $excel.Workbooks.Open($formatted)
        Assert-Layout $book $sample.Rows $sample.Columns
        $book.Close($false)
        Assert-True ((Get-WorkbookHash $original) -eq $originalHash) 'Repository sample changed.'
        if ($WriteFormattedSamples) {
            $destination = Join-Path $PSScriptRoot ([IO.Path]::GetFileName($formatted))
            Assert-True (-not (Test-Path -LiteralPath $destination)) "Refusing to overwrite: $destination"
            Copy-Item -LiteralPath $formatted -Destination $destination
        }
        Write-Output "PASS: $($sample.Name), expected dimensions, save/reopen and unchanged original."
    }
    Write-Output "ALL TESTS PASSED. Test files: $testFolder"
}
finally {
    if ($null -ne $excel) {
        $excel.DisplayAlerts = $false
        $excel.Workbooks.Close()
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}