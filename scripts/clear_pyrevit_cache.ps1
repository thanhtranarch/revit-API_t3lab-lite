# Clear pyRevit Assembly Cache
# Run this before reloading pyRevit to avoid file locking issues
# NOTE: Revit must be closed for this to work properly

param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  pyRevit Cache Cleaner" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# Check if Revit is running
$revitProcesses = Get-Process | Where-Object { $_.ProcessName -like "*Revit*" }
if ($revitProcesses -and -not $Force) {
    Write-Host "WARNING: Revit is currently running!" -ForegroundColor Red
    Write-Host "Please close Revit before running this script." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Or use -Force to attempt cleanup anyway (not recommended)" -ForegroundColor Gray
    exit 1
}

# Paths to clean
$pyRevitBase = "$env:APPDATA\pyRevit-Master"
$cachePaths = @(
    "$pyRevitBase\pyrevit\*cache*",
    "$pyRevitBase\pyrevit\*.dll",
    "$pyRevitBase\Extensions\*.dll",
    "$env:TEMP\pyRevit*"
)

$totalCleaned = 0
$totalSize = 0

foreach ($pattern in $cachePaths) {
    Write-Host "Searching: $pattern" -ForegroundColor Gray

    $items = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue -Recurse

    foreach ($item in $items) {
        try {
            $size = 0
            if ($item.PSIsContainer) {
                $size = (Get-ChildItem -Path $item.FullName -Recurse -File -ErrorAction SilentlyContinue |
                         Measure-Object -Property Length -Sum).Sum
                Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
                Write-Host "  Removed folder: $($item.Name) ($([math]::Round($size/1MB, 2)) MB)" -ForegroundColor Yellow
            } else {
                $size = $item.Length
                Remove-Item -Path $item.FullName -Force -ErrorAction Stop
                Write-Host "  Removed file: $($item.Name) ($([math]::Round($size/1KB, 2)) KB)" -ForegroundColor Yellow
            }

            $totalCleaned++
            $totalSize += $size
        } catch {
            Write-Host "  Could not remove: $($item.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Cleanup Complete!" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Items cleaned: $totalCleaned" -ForegroundColor Green
Write-Host "Space freed: $([math]::Round($totalSize/1MB, 2)) MB" -ForegroundColor Green
Write-Host ""
Write-Host "You can now restart Revit." -ForegroundColor Cyan
