[CmdletBinding()]
param(
    [ValidateSet('Both', 'Unsaved', 'Saved')]
    [string]$Mode = 'Both'
)

$ErrorActionPreference = 'Stop'
try {
    $vbeKey = 'HKCU:\Software\Microsoft\VBA\7.1\Common'
    if (Test-Path -LiteralPath $vbeKey) {
        $vbeSettings = Get-ItemProperty -LiteralPath $vbeKey
        if ($null -eq $vbeSettings.PSObject.Properties['BreakOnAllErrors']) {
            Write-Warning 'VBE BreakOnAllErrors value is absent; Error Trapping could not be verified.'
        } elseif ($vbeSettings.BreakOnAllErrors -ne 0) {
            throw "VBE Error Trapping is 'Break on All Errors' - set it to 'Break on Unhandled Errors' (VBE > Tools > Options > General)"
        }
    } else {
        Write-Warning 'VBE settings key is absent; Error Trapping could not be verified.'
    }
} catch {
    Write-Error $_ -ErrorAction Continue
    exit 1
}

function Invoke-AFSelfTest([string]$RunnerMode) {
$excel = $null
$workbooks = $null
$workbook = $null
$project = $null
$components = $null
$component = $null
$watchdog = $null
$ownedProcess = $null
$excelProcessId = 0
$ownsInstance = $false
$savedPath = Join-Path $env:TEMP 'AF_SelfTest.xlsb'
$savedByRunner = $false
$exitCode = 1
$reportPath = Join-Path $PSScriptRoot "selftest-last-$($RunnerMode.ToLower()).txt"
$tracePath = Join-Path $PSScriptRoot 'selftest-trace.txt'
$existingIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

try {
    if ($script:hangCount -ge 2) {
        throw 'Excel hung twice without a source change; refusing another attempt. Fix the cause before rerunning.'
    }
    if (Test-Path -LiteralPath $tracePath) { Remove-Item -LiteralPath $tracePath }
    $sourceRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'src'
    foreach ($name in @('AF_FilterTarget.cls', 'ActiveFilters.bas')) {
        $sourcePath = Join-Path $sourceRoot $name
        $bytes = [IO.File]::ReadAllBytes($sourcePath)
        if (($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) -or
            ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or
                ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF))) -or
            ($bytes.Length -ge 4 -and $bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 0xFE -and $bytes[3] -eq 0xFF)) {
            throw "Invalid VBA source $sourcePath`: BOM found; save as ASCII without a BOM, with CRLF line endings."
        }
        for ($index = 0; $index -lt $bytes.Length; $index++) {
            if ($bytes[$index] -gt 127) {
                throw "Invalid VBA source $sourcePath`: non-ASCII byte at offset $index."
            }
            if (($bytes[$index] -eq 10 -and ($index -eq 0 -or $bytes[$index - 1] -ne 13)) -or
                ($bytes[$index] -eq 13 -and ($index + 1 -eq $bytes.Length -or $bytes[$index + 1] -ne 10))) {
                throw "Invalid VBA source $sourcePath`: bare LF or CR at offset $index; use CRLF line endings."
            }
        }
    }
    $moduleSource = [IO.File]::ReadAllText((Join-Path $sourceRoot 'ActiveFilters.bas'), [Text.Encoding]::ASCII)
    foreach ($procedure in @('ActiveFilters_OnClick', 'ActiveFilters_SelfTest')) {
        if ($moduleSource -notmatch "(?m)^Public Sub $procedure\(Optional ByVal unused As Variant\)\r?$") {
            throw "Macro dialog hiding: expected Optional parameter on $procedure; actual declaration differs."
        }
    }
    if ($RunnerMode -eq 'Saved' -and (Test-Path -LiteralPath $savedPath)) {
        throw "Refusing to overwrite existing file: $savedPath"
    }
    if ($null -eq ('ActiveFiltersRunner.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
namespace ActiveFiltersRunner {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    }
    public sealed class Watchdog : IDisposable {
        private readonly object gate = new object();
        private readonly Process process;
        private readonly Timer timer;
        private bool finished;
        public bool TimedOut { get; private set; }
        public string KillError { get; private set; }

        public Watchdog(Process ownedProcess, int milliseconds) {
            process = ownedProcess;
            // Pin the process handle now; never resolve this PID again after a timeout.
            IntPtr handle = process.Handle;
            timer = new Timer(OnTimeout, null, milliseconds, Timeout.Infinite);
        }
        private void OnTimeout(object state) {
            lock (gate) {
                if (finished) return;
                TimedOut = true;
                try {
                    if (!process.HasExited) process.Kill();
                } catch (Exception error) {
                    KillError = error.Message;
                }
            }
        }
        public void Dispose() {
            lock (gate) { finished = true; }
            timer.Dispose();
        }
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
    [void]$ownedProcess.Handle
    $ownsInstance = $true
    Write-Host "Runner-owned Excel PID: $excelProcessId"
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
    foreach ($name in @('AF_FilterTarget.cls', 'ActiveFilters.bas')) {
        $component = $components.Import((Join-Path $sourceRoot $name))
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)
        $component = $null
    }
    if ($RunnerMode -eq 'Saved') {
        $workbook.SaveAs($savedPath, 50)
        $savedByRunner = $true
    }
    Write-Output "Mode: $RunnerMode; macro host: $($workbook.Name)"
    $macro = "'$($workbook.Name)'!ActiveFilters_SelfTestReport"
    $watchdog = New-Object ActiveFiltersRunner.Watchdog($ownedProcess, 300000)
    try { $report = [string]$excel.Run($macro, $tracePath) }
    finally { $watchdog.Dispose() }
    if ($watchdog.TimedOut) { throw 'Excel Run exceeded the five-minute timeout.' }
    Write-Output $report
    [System.IO.File]::WriteAllText($reportPath, $report, [System.Text.Encoding]::UTF8)
    $summaries = [regex]::Matches($report, '(?m)^Summary: (\d+) PASS, (\d+) FAIL\r?$')
    $passLines = [regex]::Matches($report, '(?m)^PASS: ').Count
    if ($summaries.Count -eq 1 -and $summaries[0].Groups[2].Value -eq '0' -and
        $passLines -gt 0 -and
        [int]$summaries[0].Groups[1].Value -eq $passLines -and $report -notmatch '(?m)^FAIL:') {
        $exitCode = 0
    } else {
        Write-Warning 'Self-test failed or report has a missing/inconsistent summary.'
    }
}
catch {
    $message = $_.Exception.Message
    if ($null -ne $watchdog -and $watchdog.TimedOut) {
        $script:hangCount++
        $lastStep = 'unavailable (no trace file)'
        if (Test-Path -LiteralPath $tracePath) {
            $steps = @(Get-Content -LiteralPath $tracePath | Where-Object { $_.Trim().Length -gt 0 })
            if ($steps.Count -gt 0) { $lastStep = $steps[-1] }
        }
        $message = "Excel hung - likely a VBA break or message dialog at step: $lastStep"
        if ($watchdog.KillError) { $message += "; owned-process termination failed: $($watchdog.KillError)" }
        $message += "`r`nHang fingerprint: $script:sourceFingerprint`r`nHang count for unchanged sources: $script:hangCount"
    }
    $report = "FAIL: runner error - $message`r`nSummary: 0 PASS, 1 FAIL`r`n"
    Write-Output $report
    [System.IO.File]::WriteAllText($reportPath, $report, [System.Text.Encoding]::UTF8)
    $exitCode = 1
}
finally {
    if ($null -ne $watchdog) { $watchdog.Dispose() }
    $processExited = $ownsInstance -and $ownedProcess.HasExited
    if ($null -ne $workbook -and -not $processExited) {
        try { $workbook.Close($false) }
        catch { Write-Warning "Temporary workbook cleanup failed: $($_.Exception.Message)"; $exitCode = 1 }
    }
    if ($ownsInstance -and $null -ne $excel -and -not $processExited) {
        try { $excel.Quit() }
        catch { Write-Warning "Runner Excel Quit failed: $($_.Exception.Message)"; $exitCode = 1 }
    }
    foreach ($reference in @($component, $components, $project, $workbook, $workbooks, $excel)) {
        if ($null -ne $reference -and [System.Runtime.InteropServices.Marshal]::IsComObject($reference)) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference) }
            catch { Write-Warning "COM release failed: $($_.Exception.Message)"; $exitCode = 1 }
        }
    }
    $component = $components = $project = $workbook = $workbooks = $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($null -ne $ownedProcess) {
        try {
            if (-not $ownedProcess.WaitForExit(10000)) {
                Write-Warning "Stopping ONLY runner-owned Excel PID $excelProcessId after the 10-second exit timeout."
                $ownedProcess.Kill()
                if (-not $ownedProcess.WaitForExit(10000)) { throw 'Runner-owned Excel did not exit after termination.' }
            }
            Write-Output "Runner-owned Excel PID $excelProcessId has exited."
        }
        catch { Write-Warning "Owned-process cleanup failed: $($_.Exception.Message)"; $exitCode = 1 }
        finally { $ownedProcess.Dispose() }
    }
    if ($savedByRunner) {
        try { Remove-Item -LiteralPath $savedPath }
        catch { Write-Warning "Temporary .xlsb cleanup failed: $($_.Exception.Message)"; $exitCode = 1 }
    }
}
$script:runnerExitCode = [Math]::Max($script:runnerExitCode, $exitCode)
}

$script:runnerExitCode = 0
$sourceRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'src'
$script:sourceFingerprint = (@(
    (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
    (Get-FileHash -LiteralPath (Join-Path $sourceRoot 'ActiveFilters.bas') -Algorithm SHA256).Hash
    (Get-FileHash -LiteralPath (Join-Path $sourceRoot 'AF_FilterTarget.cls') -Algorithm SHA256).Hash
) -join '-')
$script:hangCount = 0
$combinedReportPath = Join-Path $PSScriptRoot 'selftest-last.txt'
if (Test-Path -LiteralPath $combinedReportPath) {
    $previousReport = [IO.File]::ReadAllText($combinedReportPath)
    $hangs = [regex]::Matches($previousReport,
        'Hang fingerprint: ' + [regex]::Escape($script:sourceFingerprint) +
        '\r?\nHang count for unchanged sources: (\d+)')
    foreach ($hang in $hangs) { $script:hangCount = [Math]::Max($script:hangCount, [int]$hang.Groups[1].Value) }
    if ($script:hangCount -ge 2) {
        Write-Error 'Excel hung twice without a source change; refusing another attempt. Fix the cause before rerunning.' -ErrorAction Continue
        exit 1
    }
}
$modes = @($Mode)
if ($Mode -eq 'Both') { $modes = @('Unsaved', 'Saved') }
$output = foreach ($runnerMode in $modes) {
    Invoke-AFSelfTest $runnerMode
    if ($script:hangCount -ge 2) { break }
}
$output | Write-Output
[IO.File]::WriteAllText($combinedReportPath,
    ($output -join "`r`n"), [Text.Encoding]::UTF8)
exit $script:runnerExitCode