param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

Write-Host "Restoring for win-x64..."
dotnet restore -r win-x64

Write-Host "Publishing WeekDateCalc as single-file self-contained win-x64..."
dotnet publish -c $Configuration -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --no-restore

Write-Host "Done. EXE path:"
Write-Host "bin\\$Configuration\\net8.0-windows\\win-x64\\publish\\WeekDateCalc.exe"
