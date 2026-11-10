$ErrorActionPreference = 'Stop'
$excel = $null

function Assert-True($Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Invoke-Paste($Selection, [string]$ModuleName = 'FilteredPaste') {
    $result = $excel.Run("'$($hostBook.Name)'!$ModuleName.RunForTest", $Selection)
    if ($result -is [string] -and $result.StartsWith('ERROR: ')) { throw $result.Substring(7) }
    $result
}

function Assert-Rejected($Selection, [string]$Expected) {
    $rejected = $false
    try { [void](Invoke-Paste $Selection) }
    catch {
        $rejected = $true
        Assert-True ($_.Exception.ToString().Contains($Expected)) "Unexpected error: $($_.Exception.Message)"
    }
    Assert-True $rejected "Expected rejection: $Expected"
    Assert-True ($excel.EnableEvents -and $excel.ScreenUpdating -and $excel.Workbooks.Count -eq 2) 'State leaked after rejection.'
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $hostBook = $excel.Workbooks.Add()
    $module = $hostBook.VBProject.VBComponents.Add(1)
    $module.Name = 'FilteredPaste'
    $macroCode = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'excel_paste_to_filtered_rows.vba'))
    $testWrapper = @'

Public Function RunForTest(ByVal target As Range) As Variant
    On Error GoTo Failed
    target.Parent.Parent.Activate
    target.Parent.Activate
    target.Select
    RunForTest = PasteVisibleValues(target)
    Exit Function
Failed:
    RunForTest = "ERROR: " & Err.Description
End Function

Public Sub ClearCopyForTest()
    Application.CutCopyMode = False
End Sub
'@
    $module.CodeModule.AddFromString($macroCode + $testWrapper)
    $book = $excel.Workbooks.Add()
    $source = $book.Worksheets.Item(1)
    $source.Name = 'Source'
    $target = $book.Worksheets.Add()
    $target.Name = 'Target'
    $source.Range('A1:B3').NumberFormat = '@'
    $source.Range('A1').Value2 = '00123'
    $source.Range('B1').Value2 = "line1`nline2"
    $source.Range('A2').Value2 = "left`tright"
    $source.Range('B2').Value2 = '=1+2'
    $source.Range('A3').Value2 = 'last'
    $target.Range('A1').Value2 = 'Filter'
    foreach ($row in 2..6) {
        $target.Cells.Item($row, 1).Value2 = $(if ($row % 2 -eq 0) { 'keep' } else { 'hide' })
    }
    $target.Range('B2:C6').Value2 = 'untouched'
    [void]$target.Range('A1:C6').AutoFilter(1, 'keep')
    $originalCriteria = $target.AutoFilter.Filters.Item(1).Criteria1
    $source.Range('A1:B3').Copy()
    $target.Activate()
    $selection = $target.Range('B2:C6')
    $selection.Select()
    $count = Invoke-Paste $selection
    Assert-True ($count -eq 6) 'Wrong pasted cell count (including trailing blank).'
    Assert-True ($target.Range('B2').Value2 -ceq '00123') 'Leading zeros were lost.'
    Assert-True ($target.Range('C2').Value2 -ceq "line1`nline2") 'Embedded newline was lost.'
    Assert-True ($target.Range('B4').Value2 -ceq "left`tright") 'Embedded tab was lost.'
    Assert-True ($target.Range('C4').Value2 -ceq '=1+2' -and -not $target.Range('C4').HasFormula) 'Text was interpreted as a formula.'
    Assert-True ($target.Range('B6').Value2 -ceq 'last' -and $null -eq $target.Range('C6').Value2) 'Trailing blank was lost.'
    Assert-True ($target.Range('B3').Value2 -ceq 'untouched' -and $target.Range('B5').Value2 -ceq 'untouched') 'Hidden rows were overwritten.'
    Assert-True ($target.AutoFilter.Filters.Item(1).Criteria1 -ceq $originalCriteria -and $target.Rows.Item(3).Hidden) 'Filter changed.'
    Assert-True ($target.Range('B2').Font.Color -eq 255 -and $target.Range('B2').NumberFormat -eq 'General') 'Target formatting changed unexpectedly.'
    Assert-True ($excel.EnableEvents -and $excel.ScreenUpdating -and $excel.Workbooks.Count -eq 2) 'Application state or temporary workbook leaked.'
    Write-Output 'PASS: filtered multi-column row order, text types, embedded delimiters, trailing blank, filters and cleanup.'

    foreach ($row in 2..6) { $target.Cells.Item($row, 7).Value2 = 'anchor-untouched' }
    $target.Range('H2').Value2 = 7.2
    $target.Range('H4').Value2 = 8.3
    $target.Range('H2:H4').Copy()
    $count = Invoke-Paste $target.Range('G2')
    Assert-True ($count -eq 2 -and $target.Range('G2').Value2 -eq 7.2 -and $target.Range('G4').Value2 -eq 8.3) 'Two copied filtered cells did not expand from one starting cell.'
    Assert-True ($target.Range('G3').Value2 -ceq 'anchor-untouched' -and $target.Range('G6').Value2 -ceq 'anchor-untouched') "Anchor expansion changed hidden or excess rows: G3=[$($target.Range('G3').Value2)], G6=[$($target.Range('G6').Value2)]."
    Assert-True ($excel.Selection.Address() -ceq '$G$2') 'Anchor selection was not restored.'
    $source.Range('A1:B3').Copy()
    [void](Invoke-Paste $target.Range('B2'))
    Assert-True ($target.Range('B2').Value2 -ceq '00123' -and $target.Range('C2').Value2 -ceq "line1`nline2" -and $target.Range('B6').Value2 -ceq 'last') 'Multi-column anchor expansion lost shape or values.'
    $source.Range('A1:A2').Copy()
    Assert-Rejected $target.Range('B6') 'Not enough visible target rows'
    $source.Range('A1:B1').Copy()
    Assert-Rejected $target.Range('C2') 'Not enough visible target columns'
    Assert-True ($target.Range('B6').Value2 -ceq 'last') 'Rejected anchor expansion wrote a partial target.'
    Assert-True ($target.AutoFilter.Filters.Item(1).Criteria1 -ceq $originalCriteria -and $target.Rows.Item(3).Hidden) 'Anchor expansion changed the filter.'
    Write-Output 'PASS: screenshot case H2/H4 to G2 anchor, multi-column anchor shape and filter-boundary rejection.'

    $target.ShowAllData()
    $source.Range('D1').Value2 = 9.125
    $source.Range('E1').Formula = '=D1*2'
    $source.Range('F1').Value2 = $true
    $source.Range('G1').Formula = '=NA()'
    $source.Range('D1:G1').Copy()
    [void](Invoke-Paste $target.Range('E2:H2'))
    Assert-True ($target.Range('E2').Value2 -eq 9.125 -and $target.Range('F2').Value2 -eq 18.25) 'Number or formula result changed.'
    Assert-True (-not $target.Range('F2').HasFormula -and $target.Range('G2').Value2 -eq $true) 'Formula or Boolean type changed.'
    Assert-True ($target.Range('H2').Text -eq '#N/A') 'Excel error value was lost.'
    $source.Range('E1').Copy()
    [void](Invoke-Paste $target.Range('J2'))
    Assert-True ($target.Range('J2').Value2 -eq 18.25 -and $target.Range('J3').Value2 -eq $null) 'Single-cell target expanded or formula result changed.'
    $source.Range('J1:J3').Clear()
    $source.Range('J1:J3').Copy()
    [void](Invoke-Paste $target.Range('J4:J6'))
    Assert-True ($target.Range('J4:J6').Value2[1, 1] -eq $null -and $target.Range('J6').Value2 -eq $null) 'All-blank source failed.'
    Write-Output 'PASS: numbers, formula results, Boolean and error types, single-cell target and all-blank source.'

    $source.Range('J1').Value2 = 'first'
    $source.Range('J2').Value2 = 'second'
    $source.Range('J3').Value2 = 'third'
    $source.Range('J4').Value2 = 'fourth'
    $source.Range('J1:J4').Copy()
    $reverseSelection = $target.Range('M5:N5,M3:N3')
    [void](Invoke-Paste $reverseSelection)
    Assert-True ($target.Range('M3').Value2 -ceq 'first' -and $target.Range('N3').Value2 -ceq 'second' -and $target.Range('M5').Value2 -ceq 'third' -and $target.Range('N5').Value2 -ceq 'fourth') 'Disjoint target order is not row-major.'
    $target.Range('P2:R4').Value2 = 'hidden-test'
    $target.Columns.Item(17).Hidden = $true
    $target.Rows.Item(3).Hidden = $true
    $source.Range('J1:J4').Copy()
    [void](Invoke-Paste $target.Range('P2:R4'))
    Assert-True ($target.Range('P2').Value2 -ceq 'first' -and $target.Range('R2').Value2 -ceq 'second' -and $target.Range('P4').Value2 -ceq 'third') 'Hidden row/column target mapping failed.'
    Assert-True ($target.Range('Q2').Value2 -ceq 'hidden-test' -and $target.Range('P3').Value2 -ceq 'hidden-test') 'Manually hidden cells changed.'
    $target.Columns.Item(17).Hidden = $false
    $target.Rows.Item(3).Hidden = $false
    $target.Range('T2').Value2 = 'overlap-one'
    $target.Range('T3').Value2 = 'overlap-two'
    $target.Range('T2:T3').Copy()
    [void](Invoke-Paste $target.Range('T3:T4'))
    Assert-True ($target.Range('T3').Value2 -ceq 'overlap-one' -and $target.Range('T4').Value2 -ceq 'overlap-two') 'Overlapping source and target failed.'
    Write-Output 'PASS: disjoint selections, hidden rows/columns and overlapping source/target.'

    $target.Range('J10:J11').Value2 = 'must-stay'
    $source.Range('J1').Copy()
    Assert-Rejected $target.Range('J10:J11') 'count mismatch'
    Assert-True ($target.Range('J10').Value2 -ceq 'must-stay' -and $target.Range('J11').Value2 -ceq 'must-stay') 'Mismatch changed data.'
    $target.Protect()
    $source.Range('J1').Copy()
    Assert-Rejected $target.Range('J10') 'protected'
    $target.Unprotect()
    $target.Range('L10:M10').Merge()
    $source.Range('J1').Copy()
    Assert-Rejected $target.Range('L10:M10') 'Merged cells'
    $target.Range('L10:M10').UnMerge()
    $target.Range('L10:L11').FormulaArray = '=ROW(L10:L11)'
    $source.Range('J1:J2').Copy()
    Assert-Rejected $target.Range('L10:L11') 'array formulas'
    $target.Range('L10:L11').ClearContents()
    $target.Rows.Item(12).Hidden = $true
    $source.Range('J1').Copy()
    Assert-Rejected $target.Range('L12:M12') 'No visible'
    $target.Rows.Item(12).Hidden = $false
    $source.Range('J1').Copy()
    Assert-Rejected $target.Range('Z1:Z100001') 'smaller target'
    $excel.Run("'$($hostBook.Name)'!FilteredPaste.ClearCopyForTest")
    Assert-Rejected $target.Range('J10') 'Copy source cells'
    Assert-True ($target.Range('J10').Value2 -ceq 'must-stay') 'Rejected operations changed data.'
    Write-Output 'PASS: mismatch, protected sheet, merged/array targets, invisible target, size limit and missing copy rejected without writes.'

    $source.Range('L1').Value2 = 'Filter'
    $source.Range('L2').Value2 = 'yes'
    $source.Range('L3').Value2 = 'no'
    $source.Range('L4').Value2 = 'yes'
    $source.Range('M2').Value2 = 'visible-one'
    $source.Range('M3').Value2 = 'hidden-source'
    $source.Range('M4').Value2 = 'visible-two'
    [void]$source.Range('L1:M4').AutoFilter(1, 'yes')
    $source.Range('M2:M4').Copy()
    [void](Invoke-Paste $target.Range('V2:V3'))
    Assert-True ($target.Range('V2').Value2 -ceq 'visible-one' -and $target.Range('V3').Value2 -ceq 'visible-two') 'Filtered source rows were not respected.'
    $source.ShowAllData()
    $target.Range('A20').Value2 = 'Group'
    $target.Range('B20').Value2 = 'Data'
    foreach ($row in 21..24) {
        $target.Cells.Item($row, 1).Value2 = $(if ($row % 2 -eq 1) { 'keep' } else { 'hide' })
        $target.Cells.Item($row, 2).Value2 = 'table-original'
    }
    $target.AutoFilterMode = $false
    $table = $target.ListObjects.Add(1, $target.Range('A20:B24'), $null, 1)
    [void]$table.Range.AutoFilter(1, 'keep')
    $criteria = $table.AutoFilter.Filters.Item(1).Criteria1
    $source.Range('J1:J2').Copy()
    [void](Invoke-Paste $target.Range('B21:B24'))
    Assert-True ($target.Range('B21').Value2 -ceq 'first' -and $target.Range('B23').Value2 -ceq 'second' -and $target.Range('B22').Value2 -ceq 'table-original') 'Table target mapping failed.'
    Assert-True ($table.AutoFilter.Filters.Item(1).Criteria1 -ceq $criteria -and $target.Rows.Item(22).Hidden) 'Table filter changed.'
    $table.AutoFilter.ShowAllData()
    Write-Output 'PASS: filtered source and Excel Table target/filter preservation.'

    $otherBook = $excel.Workbooks.Add()
    $otherBook.Worksheets.Item(1).Range('A1').Value2 = 'other-workbook'
    $otherBook.Worksheets.Item(1).Range('A1').Copy()
    [void](Invoke-Paste $target.Range('V5'))
    Assert-True ($target.Range('V5').Value2 -ceq 'other-workbook') 'Cross-workbook source failed.'
    $otherBook.Close($false)
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $source.Range('J1').Copy()
    [void](Invoke-Paste $target.Range('V6'))
    Assert-True (-not $excel.EnableEvents -and -not $excel.ScreenUpdating) 'Pre-disabled application settings were not preserved.'
    $excel.EnableEvents = $true
    $excel.ScreenUpdating = $true
    Write-Output 'PASS: cross-workbook source and pre-disabled application settings.'

    $faultModule = $hostBook.VBProject.VBComponents.Add(1)
    $faultModule.Name = 'FaultPaste'
    $faultCode = $macroCode.Replace('targetCells(cellIndex).Font.Color = vbRed', "targetCells(cellIndex).Font.Color = vbRed`r`n        If cellIndex = 2 Then Err.Raise vbObjectError + 2199, , ""Injected write failure""")
    Assert-True ($faultCode -cne $macroCode) 'Fault injection did not match the write path.'
    $faultModule.CodeModule.AddFromString($faultCode + $testWrapper)
    $target.Range('X2').Formula = '=2+3'
    $target.Range('X2').NumberFormat = '0.00'
    $target.Range('X2').Font.ThemeColor = 5
    $target.Range('X2').Font.TintAndShade = 0.25
    $target.Range('X3').NumberFormat = '@'
    $target.Range('X3').Value2 = '=literal'
    $target.Range('X3').Font.Color = 32768
    $originalTint = $target.Range('X2').Font.TintAndShade
    $source.Range('J1:J2').Copy()
    $rejected = $false
    try { [void](Invoke-Paste $target.Range('X2:X3') 'FaultPaste') }
    catch {
        $rejected = $true
        Assert-True ($_.Exception.ToString().Contains('Injected write failure')) 'Wrong injected failure.'
    }
    Assert-True $rejected 'Injected failure did not run.'
    Assert-True ($target.Range('X2').Formula -ceq '=2+3' -and $target.Range('X2').NumberFormat -ceq '0.00') 'Rollback did not restore formula/format.'
    Assert-True ($target.Range('X2').Font.ThemeColor -eq 5 -and $target.Range('X2').Font.TintAndShade -eq $originalTint) 'Rollback did not restore theme font.'
    Assert-True ($target.Range('X3').Value2 -ceq '=literal' -and -not $target.Range('X3').HasFormula -and $target.Range('X3').Font.Color -eq 32768) 'Rollback did not restore literal text/color.'
    Assert-True ($excel.EnableEvents -and $excel.ScreenUpdating -and $excel.Workbooks.Count -eq 2) 'Injected failure leaked application state.'
    Write-Output 'PASS: injected mid-write failure restores formulas, literal text, formats, font colors and application state.'

    $target.Range('B21:B24').Formula = '=ROW()'
    $target.Range('B22').Value2 = 'calculated-column-exception'
    $target.Range('B24').Value2 = 500
    [void]$table.Range.AutoFilter(1, 'keep')
    $source.Range('J1:J2').Copy()
    $rejected = $false
    try { [void](Invoke-Paste $target.Range('B21:B24') 'FaultPaste') }
    catch {
        $rejected = $true
        Assert-True ($_.Exception.ToString().Contains('Injected write failure')) 'Wrong table rollback failure.'
    }
    Assert-True $rejected 'Table fault injection did not run.'
    Assert-True ($target.Range('B21').Formula -ceq '=ROW()' -and $target.Range('B23').Formula -ceq '=ROW()') 'Table formula rollback failed.'
    Assert-True ($target.Range('B22').Value2 -ceq 'calculated-column-exception' -and $target.Range('B24').Value2 -eq 500) 'Table rollback overwrote hidden exceptions.'
    Assert-True ($table.AutoFilter.Filters.Item(1).Criteria1 -ceq $criteria -and $target.Rows.Item(22).Hidden) 'Table rollback changed the filter.'
    $table.AutoFilter.ShowAllData()
    $target.Range('Z10').Formula2 = '=SEQUENCE(2)'
    $source.Range('J1:J2').Copy()
    Assert-Rejected $target.Range('Z10:Z11') 'spill ranges'
    Assert-True ($target.Range('Z10').Formula2 -ceq '=SEQUENCE(2)' -and $target.Range('Z11').Value2 -eq 2) 'Spill rejection changed the formula.'
    $book.Protect('test-only', $true, $false)
    $source.Range('J1').Copy()
    [void](Invoke-Paste $target.Range('V8'))
    Assert-True ($target.Range('V8').Value2 -ceq 'first' -and $book.ProtectStructure) 'Protected workbook structure prevented a values-only paste.'
    $book.Unprotect('test-only')
    Write-Output 'PASS: calculated-column rollback preserves hidden exceptions, spill targets rejected and protected workbook structure supported.'
} finally {
    if ($null -ne $excel) {
        $excel.DisplayAlerts = $false
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}