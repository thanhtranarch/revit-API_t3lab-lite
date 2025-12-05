# revit-API_t3lab-lite

T3Lab Lite - IronPython Scripts for Autodesk Revit

## Installation

1. Install [pyRevit](https://github.com/eirannejad/pyRevit)
2. Clone or download this repository
3. Copy the `T3Lab_Lite.extension` folder to your pyRevit extensions directory
4. Reload pyRevit

## Known Issues

### pyRevit Reload Error

If you encounter an `IOError` when trying to reload pyRevit (file locking issue), please see:
📄 [PYREVIT_RELOAD_FIX.md](PYREVIT_RELOAD_FIX.md)

**Quick Fix:**
```powershell
# Run as Administrator
PowerShell -ExecutionPolicy Bypass -File scripts/fix_pyrevit_reload.ps1
```

## Troubleshooting

For other issues, please check the documentation or open an issue on GitHub.