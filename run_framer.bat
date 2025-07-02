@echo off
:: This batch file launches the Window Framer C# application.

:: Change the current directory to the directory where this batch file is located.
:: The /d switch is important in case the script is run from a different drive.
cd /d "%~dp0"

:: Define the paths to the local dotnet executable and the project
set DOTNET_EXE=C:\Users\szil\AppData\Local\Microsoft\dotnet\dotnet.exe
set PROJECT_PATH=.\window_frame_csharp

:: Check if the local .NET SDK exists
if not exist "%DOTNET_EXE%" (
    echo Local .NET SDK not found at '%DOTNET_EXE%'.
    echo Please run the installation commands first.
    pause
    exit /b
)

:: Run the application
echo Starting Window Framer...
call %DOTNET_EXE% run --project "%PROJECT_PATH%"
