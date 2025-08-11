# Navisworks XML Filter Generator - PowerShell Version
# This script activates the virtual environment and runs the XML generator

Write-Host "========================================" -ForegroundColor Green
Write-Host "Navisworks XML Filter Generator" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Change to script directory
Set-Location $PSScriptRoot

# Check if virtual environment exists
if (-not (Test-Path "..\.venv\Scripts\Activate.ps1")) {
    Write-Host "ERROR: Virtual environment not found at ..\.venv\" -ForegroundColor Red
    Write-Host "Please ensure the virtual environment is set up correctly." -ForegroundColor Red
    exit 1
}

# Check if Excel file exists
if (-not (Test-Path "Navis_KKS_List.xlsx")) {
    Write-Host "ERROR: Navis_KKS_List.xlsx not found in current directory" -ForegroundColor Red
    Write-Host "Please ensure the Excel file is present." -ForegroundColor Red
    exit 1
}

Write-Host "Activating virtual environment..." -ForegroundColor Yellow
Write-Host ""

# Activate virtual environment and run script
& "..\.venv\Scripts\Activate.ps1"
python navisworks_xml_filter_gen.py

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Script execution completed" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Check if XML file was generated
$xmlFiles = Get-ChildItem -Name "*_navis_filter.xml"
if ($xmlFiles) {
    Write-Host "SUCCESS: XML file generated - $($xmlFiles[0])" -ForegroundColor Green
} else {
    Write-Host "WARNING: XML file may not have been generated" -ForegroundColor Yellow
}
