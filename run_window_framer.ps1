# This script launches the Window Framer C# application.
# It is designed to be run from the project's root directory.

# Get the directory where this script is located.
$scriptDir = $PSScriptRoot

# Define the full paths for the local .NET executable and the C# project.
$dotnetExe = Join-Path -Path $scriptDir -ChildPath ".dotnet/dotnet.exe"
$projectPath = Join-Path -Path $scriptDir -ChildPath "window_frame_csharp"

# Check if the local .NET SDK exists. If not, guide the user.
if (-not (Test-Path $dotnetExe)) {
    Write-Warning "Local .NET SDK not found at '$dotnetExe'."
    Write-Host "Please run the installation commands from the README or previous instructions first."
    Read-Host "Press Enter to exit."
    exit
}

# If the SDK is found, run the application.
Write-Host "Starting Window Framer..."
& $dotnetExe run --project $projectPath

