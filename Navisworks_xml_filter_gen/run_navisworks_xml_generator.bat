@echo off
REM Navisworks XML Filter Generator
REM This script activates the virtual environment and runs the XML generator

echo ========================================
echo Navisworks XML Filter Generator
echo ========================================
echo.

REM Change to the script directory
cd /d "%~dp0"

REM Check if virtual environment exists
if not exist "..\.venv\Scripts\Activate.ps1" (
    echo ERROR: Virtual environment not found at ..\.venv\
    echo Please ensure the virtual environment is set up correctly.
    echo.
    exit /b 1
)

REM Check if Excel file exists
if not exist "Navis_KKS_List.xlsx" (
    echo ERROR: Navis_KKS_List.xlsx not found in current directory
    echo Please ensure the Excel file is present.
    echo.
    exit /b 1
)

echo Activating virtual environment...
echo.

REM Run PowerShell command to activate venv and run the script
powershell -ExecutionPolicy Bypass -Command "Set-Location '%~dp0'; & '..\\.venv\\Scripts\\Activate.ps1'; python navisworks_xml_filter_gen.py"

echo.
echo ========================================
echo Script execution completed
echo ========================================
echo.

REM Check if XML file was generated (look for any *_navis_filter.xml file)
for %%f in (*_navis_filter.xml) do (
    echo SUCCESS: XML file generated - %%f
    goto :found
)
echo WARNING: XML file may not have been generated
:found
