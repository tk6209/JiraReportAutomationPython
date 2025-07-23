# setup.ps1 - Organizes project structure for JiraReportAutomationPython

Write-Host "`n[INFO] Setting up JiraReportAutomationPython project structure..."

# Define base paths
$root     = Split-Path -Parent $MyInvocation.MyCommand.Definition
$assets   = Join-Path $root "assets"
$macros   = Join-Path $root "macros"
$archive  = Join-Path $root "archive"

# Create directories if they don't exist
$folders = @($assets, $macros, $archive)
foreach ($folder in $folders) {
    if (-Not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory | Out-Null
        Write-Host "[CREATED] $folder"
    }
    else {
        Write-Host "[EXISTS ] $folder"
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
        Move-Item -Path $sourcePath -Destination $targetPath -Force
        Write-Host "[MOVED  ] $($item.Source) to $($item.Target)"
    }
    else {
        Write-Host "[SKIPPED] $($item.Source) not found"
    }
}

Write-Host "`n[DONE] Project setup complete."