@echo off
REM Build and publish the Shortcut Launcher application
REM Run this from the ShortcutLauncher directory

echo ========================================
echo   Shortcut Launcher - Build Script
echo ========================================
echo.

REM Check if in correct directory
if not exist "ShortcutLauncher.csproj" (
    echo ERROR: ShortcutLauncher.csproj not found!
    echo Please run this script from the ShortcutLauncher directory.
    pause
    exit /b 1
)

echo [1/4] Cleaning previous builds...
if exist "bin\" rmdir /s /q "bin"
if exist "obj\" rmdir /s /q "obj"

echo [2/4] Restoring NuGet packages...
dotnet restore

echo [3/4] Building Release configuration...
dotnet build -c Release

echo [4/4] Publishing single-file executable...
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true --self-contained true

echo.
echo ========================================
echo   Build Complete!
echo ========================================
echo.
echo Executable location:
echo bin\Release\net8.0-windows\win-x64\publish\ShortcutLauncher.exe
echo.
echo Next steps:
echo 1. Navigate to the publish folder
echo 2. Right-click ShortcutLauncher.exe
echo 3. Select "Pin to taskbar"
echo.

set /p OPEN="Open publish folder? (Y/N): "
if /i "%OPEN%"=="Y" explorer "bin\Release\net8.0-windows\win-x64\publish"

pause
