# pyRevit Reload Error Fix

## Problem Description

When clicking the pyRevit Reload button, you may encounter this error:

```
IOError: System.IO.IOException: The process cannot access the file because it is being used by another process. (Exception from HRESULT: 0x80070020)
```

This occurs in `asmmaker.py` when pyRevit tries to save dynamically compiled assemblies, but the files are locked by the current Revit process.

## Root Cause

The issue is in the pyRevit core file:
`C:\Users\<username>\AppData\Roaming\pyRevit-Master\pyrevitlib\pyrevit\loader\asmmaker.py`

When pyRevit reloads, it tries to save new assembly files, but the old files are still locked by the running Revit process, preventing the write operation.

## Quick Workarounds

### Option 1: Restart Revit (Recommended)
Simply close and reopen Revit. This clears all file locks.

### Option 2: Use Rocket Mode
Rocket mode avoids some assembly caching issues:
1. Open pyRevit settings
2. Enable "Rocket Mode"
3. Restart Revit

### Option 3: Clear Assembly Cache
Manually delete assembly cache files before reloading:
1. Close Revit
2. Navigate to: `%APPDATA%\pyRevit-Master\pyrevit\`
3. Delete any `.dll` or assembly cache folders
4. Restart Revit

## Permanent Fix

To permanently fix this in your pyRevit installation, you need to patch the core `asmmaker.py` file with retry logic and better file handling.

### Automated Fix Script

A PowerShell script has been provided in this repository to automatically patch your pyRevit installation:
`scripts/fix_pyrevit_reload.ps1`

**Usage:**
```powershell
# Run as Administrator
PowerShell -ExecutionPolicy Bypass -File scripts/fix_pyrevit_reload.ps1
```

### Manual Fix

If you prefer to manually patch the file:

1. Navigate to: `C:\Users\<username>\AppData\Roaming\pyRevit-Master\pyrevitlib\pyrevit\loader\`
2. Backup `asmmaker.py`
3. Replace the problematic section (see below)

**Original code** (around line 129 in `_create_asm_file`):
```python
asm_builder.Save(
    asm_file_path,
    peke,
    imachine
)
```

**Replace with:**
```python
import time

# Retry logic to handle file locking issues
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
            # Wait before retrying
            time.sleep(retry_delay)
            retry_delay *= 2  # Exponential backoff
        else:
            # Last attempt failed, try with unique filename
            import uuid
            unique_suffix = str(uuid.uuid4())[:8]
            base_name = os.path.splitext(asm_file_path)[0]
            ext = os.path.splitext(asm_file_path)[1]
            new_path = "{}_{}{}".format(base_name, unique_suffix, ext)

            try:
                asm_builder.Save(new_path, peke, imachine)
                # Update the path for subsequent operations
                asm_file_path = new_path
            except IOError:
                # If all retries fail, raise the original error
                raise e
```

## Prevention

To minimize the occurrence of this issue:

1. **Avoid frequent reloads** - Only reload when necessary
2. **Close unnecessary Revit sessions** - Multiple instances can increase file locking
3. **Update pyRevit** - Newer versions may have better file handling
4. **Use Rocket Mode** - This mode has optimizations that reduce file locking issues

## Additional Resources

- [pyRevit Documentation](https://pyrevit.readthedocs.io/)
- [pyRevit GitHub Issues](https://github.com/eirannejad/pyRevit/issues)

---

**Last Updated:** 2025-12-05
**Author:** T3Lab Development Team
