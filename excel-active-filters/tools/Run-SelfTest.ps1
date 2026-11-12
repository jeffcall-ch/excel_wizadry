[CmdletBinding()]
param([switch]$TraceSteps)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbooks = $null
$workbook = $null
$project = $null
$components = $null
$component = $null
$codeModule = $null
$ownedProcess = $null
$excelProcessId = 0
$ownsInstance = $false
$exitCode = 1
$expectedChecks = 26
$reportPath = Join-Path $PSScriptRoot 'selftest-last.txt'
$existingIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

try {
    if ($null -eq ('ActiveFiltersRunner.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ActiveFiltersRunner {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    }
}
'@
    }
    $excel = New-Object -ComObject Excel.Application
    [uint32]$instanceId = 0
    $threadId = [ActiveFiltersRunner.NativeMethods]::GetWindowThreadProcessId(
        [IntPtr]([long]$excel.Hwnd), [ref]$instanceId)
    if ($threadId -eq 0 -or $instanceId -eq 0 -or $instanceId -in $existingIds) {
        throw 'Could not verify a NEW Excel process. Refusing to change or quit an existing instance.'
    }
    $excelProcessId = [int]$instanceId
    $ownedProcess = [System.Diagnostics.Process]::GetProcessById($excelProcessId)
    if ($ownedProcess.ProcessName -ne 'EXCEL') {
        $ownedProcess.Dispose()
        $ownedProcess = $null
        throw 'Window PID is not an Excel process; refusing to interact with it.'
    }
    $ownsInstance = $true
    Write-Output "Runner-owned Excel PID: $excelProcessId"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbooks = $excel.Workbooks
    if ($workbooks.Count -ne 0) { throw 'New automation instance unexpectedly contains startup workbooks; no files will be opened or saved.' }
    $workbook = $workbooks.Add()
    try {
        $project = $workbook.VBProject
        $components = $project.VBComponents
        [void]$components.Count
    }
    catch {
        throw ('VBA project access is blocked. Check Excel File > Options > Trust Center > ' +
            'Trust Center Settings > Macro Settings > Trust access to the VBA project object model. ' +
            'Organization policy may block access. The runner does not change this setting. ' + $_.Exception.Message)
    }
    $sourceRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'src'
    foreach ($name in @('AF_FilterTarget.cls', 'ActiveFilters.bas')) {
        $component = $components.Import((Join-Path $sourceRoot $name))
        if ($TraceSteps -and $name -eq 'ActiveFilters.bas') {
            $tracePath = Join-Path $PSScriptRoot 'selftest-trace.txt'
            [IO.File]::WriteAllText($tracePath, '', [Text.Encoding]::ASCII)
            $codeModule = $component.CodeModule
            $code = $codeModule.Lines(1, $codeModule.CountOfLines)
            $code = [regex]::Replace($code, '(?m)^(\s*mCurrentStep = "[^"\r\n]*")\r?$', '$1' + "`r`n    AF_RunnerTrace mCurrentStep")
            $code += @"

Private Sub AF_RunnerTrace(stepText As String)
    Dim fileNumber As Integer
    On Error Resume Next
    fileNumber = FreeFile
    Open "$tracePath" For Append As #fileNumber
    Print #fileNumber, stepText
    Close #fileNumber
End Sub
"@
            $codeModule.DeleteLines(1, $codeModule.CountOfLines)
            $codeModule.AddFromString($code)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($codeModule)
            $codeModule = $null
        }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
        $component = $null
    }
    $macro = "'$($workbook.Name)'!ActiveFilters_SelfTestReport"
    $report = [string]$excel.Run($macro)
    Write-Output $report
    [System.IO.File]::WriteAllText($reportPath, $report, [System.Text.Encoding]::UTF8)
    $summaries = [regex]::Matches($report, '(?m)^Summary: (\d+) PASS, (\d+) FAIL\r?$')
    $passLines = [regex]::Matches($report, '(?m)^PASS: ').Count
    if ($summaries.Count -eq 1 -and $summaries[0].Groups[2].Value -eq '0' -and
        [int]$summaries[0].Groups[1].Value -eq $expectedChecks -and
        [int]$summaries[0].Groups[1].Value -eq $passLines -and $report -notmatch '(?m)^FAIL:') {
        $exitCode = 0
    }
}
catch {
    $report = "FAIL: runner error - $($_.Exception.Message)`r`nSummary: 0 PASS, 1 FAIL`r`n"
    Write-Output $report
    [System.IO.File]::WriteAllText($reportPath, $report, [System.Text.Encoding]::UTF8)
    $exitCode = 1
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) }
        catch { Write-Warning "Temporary workbook cleanup failed: $($_.Exception.Message)"; $exitCode = 1 }
    }
    if ($ownsInstance -and $null -ne $excel) {
        try { $excel.Quit() }
        catch { Write-Warning "Runner Excel Quit failed: $($_.Exception.Message)"; $exitCode = 1 }
    }
    foreach ($reference in @($codeModule, $component, $components, $project, $workbook, $workbooks, $excel)) {
        if ($null -ne $reference -and [System.Runtime.InteropServices.Marshal]::IsComObject($reference)) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference) }
            catch { Write-Warning "COM release failed: $($_.Exception.Message)"; $exitCode = 1 }
        }
    }
    $codeModule = $component = $components = $project = $workbook = $workbooks = $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($null -ne $ownedProcess) {
        try {
            if (-not $ownedProcess.WaitForExit(10000)) {
                Write-Warning "Stopping ONLY runner-owned Excel PID $excelProcessId after the 10-second exit timeout."
                $ownedProcess.Kill()
                $ownedProcess.WaitForExit()
            }
            Write-Output "Runner-owned Excel PID $excelProcessId has exited."
        }
        catch { Write-Warning "Owned-process cleanup failed: $($_.Exception.Message)"; $exitCode = 1 }
        finally { $ownedProcess.Dispose() }
    }
}
exit $exitCode