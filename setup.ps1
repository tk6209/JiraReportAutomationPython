# setup.ps1 - Organizes project structure for JiraReportAutomationPython

Write-Host "`n🔧 Setting up JiraReportAutomationPython project structure..." -ForegroundColor Cyan

# Define base paths
$root     = Split-Path -Parent $MyInvocation.MyCommand.Definition
$assets   = Join-Path $root "assets"
$macros   = Join-Path $root "macros"
$archive  = Join-Path $root "archive"

# Create directories if not exist
$folders = @($assets, $macros, $archive)
foreach ($folder in $folders) {
    if (-Not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory | Out-Null
        Write-Host "📁 Created: $folder" -ForegroundColor Green
    } else {
        Write-Host "📁 Exists:  $folder" -ForegroundColor DarkGray
    }
}

# Move files if they exist
$filesToMove = @(
    @{ Source = "python.ico"; Target = $assets },
    @{ Source = "Dashboard of Operations.xlsm"; Target = $macros },
    @{ Source = "Dashboard of Operations 2019-12-19.xlsx"; Target = $archive },
    @{ Source = "OnePage.py.bak"; Target = $archive }
)

foreach ($item in $filesToMove) {
    $sourcePath = Join-Path $root $item.Source
    $targetPath = Join-Path $item.Target $item.Source

    if (Test-Path $sourcePath) {
        Move-Item -Path $sourcePath -Destination $targe
