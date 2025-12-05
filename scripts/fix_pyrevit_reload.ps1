# pyRevit Reload Fix - Automated Patcher
# This script patches the pyRevit core to fix the assembly file locking issue
# Run as Administrator

param(
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  pyRevit Reload Fix - Assembly Lock Patcher" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# Locate pyRevit installation
$pyRevitPath = "$env:APPDATA\pyRevit-Master\pyrevitlib\pyrevit\loader"
$asmmakerPath = Join-Path $pyRevitPath "asmmaker.py"

if (-not (Test-Path $asmmakerPath)) {
    Write-Host "ERROR: Could not find pyRevit installation at:" -ForegroundColor Red
    Write-Host $pyRevitPath -ForegroundColor Red
    Write-Host ""
    Write-Host "Please ensure pyRevit is installed." -ForegroundColor Yellow
    exit 1
}

Write-Host "Found pyRevit installation:" -ForegroundColor Green
Write-Host $pyRevitPath -ForegroundColor Gray
Write-Host ""

# Create backup
$backupPath = "$asmmakerPath.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
Write-Host "Creating backup..." -ForegroundColor Yellow
if (-not $WhatIf) {
    Copy-Item $asmmakerPath $backupPath
    Write-Host "Backup created: $backupPath" -ForegroundColor Green
} else {
    Write-Host "[WHATIF] Would create backup: $backupPath" -ForegroundColor Gray
}
Write-Host ""

# Read the current file
$content = Get-Content $asmmakerPath -Raw

# Check if already patched
if ($content -match "# pyRevit Reload Fix") {
    Write-Host "File appears to already be patched!" -ForegroundColor Yellow
    Write-Host "No changes needed." -ForegroundColor Green
    exit 0
}

# Define the patch
$oldCode = @'
        asm_builder.Save(
            asm_file_path,
            peke,
            imachine
        )
'@

$newCode = @'
        # pyRevit Reload Fix - Assembly lock handling with retry logic
        import time
        import uuid

        max_retries = 5
        retry_delay = 0.5  # seconds

        for attempt in range(max_retries):
            try:
                asm_builder.Save(
                    asm_file_path,
                    peke,
                    imachine
                )
                break  # Success, exit retry loop
            except IOError as e:
                if attempt < max_retries - 1:
                    # Wait before retrying with exponential backoff
                    time.sleep(retry_delay)
                    retry_delay *= 2
                else:
                    # Last attempt failed, try with unique filename
                    try:
                        unique_suffix = str(uuid.uuid4())[:8]
                        base_name = op.splitext(asm_file_path)[0]
                        ext = op.splitext(asm_file_path)[1]
                        new_path = "{}_{}{}".format(base_name, unique_suffix, ext)
                        asm_builder.Save(new_path, peke, imachine)
                        asm_file_path = new_path
                        logger.warning("Assembly saved with unique name due to lock: {}".format(new_path))
                        break
                    except IOError:
                        # If all retries fail, raise the original error
                        raise e
'@

# Apply the patch
if ($content -match [regex]::Escape($oldCode)) {
    Write-Host "Applying patch..." -ForegroundColor Yellow
    $patchedContent = $content -replace [regex]::Escape($oldCode), $newCode

    if (-not $WhatIf) {
        Set-Content -Path $asmmakerPath -Value $patchedContent -NoNewline
        Write-Host "Patch applied successfully!" -ForegroundColor Green
    } else {
        Write-Host "[WHATIF] Would apply patch to: $asmmakerPath" -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "IMPORTANT: Please restart Revit for changes to take effect." -ForegroundColor Cyan
} else {
    Write-Host "WARNING: Could not find the expected code pattern to patch." -ForegroundColor Yellow
    Write-Host "The file may have been modified or is from a different pyRevit version." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please apply the manual fix as described in PYREVIT_RELOAD_FIX.md" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Patch Complete!" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Backup location: $backupPath" -ForegroundColor Gray
Write-Host ""
Write-Host "To revert the patch, run:" -ForegroundColor Yellow
Write-Host "  Copy-Item '$backupPath' '$asmmakerPath' -Force" -ForegroundColor Gray
