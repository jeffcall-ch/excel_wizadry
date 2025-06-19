# PowerShell script to run tests with coverage report
Write-Host "Running tests with coverage report..." -ForegroundColor Green

# Get the directory where this script is located
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Change to the script directory
Set-Location $scriptDir

# Run pytest with coverage
python -m pytest -xvs --cov=.. --cov-report=term --cov-report=html

Write-Host "Test execution completed." -ForegroundColor Green

