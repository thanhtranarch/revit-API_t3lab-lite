"""
pyRevit Reload Fix - Automated Patcher (Python version)
This script patches the pyRevit core to fix the assembly file locking issue

Usage:
    python fix_pyrevit_reload.py
    python fix_pyrevit_reload.py --whatif
"""

import os
import sys
import shutil
import re
from datetime import datetime

def main():
    whatif = "--whatif" in sys.argv or "-n" in sys.argv

    print("=" * 60)
    print("  pyRevit Reload Fix - Assembly Lock Patcher")
    print("=" * 60)
    print()

    # Locate pyRevit installation
    appdata = os.environ.get('APPDATA')
    if not appdata:
        print("ERROR: Could not determine APPDATA directory")
        return 1

    pyrevit_path = os.path.join(appdata, 'pyRevit-Master', 'pyrevitlib', 'pyrevit', 'loader')
    asmmaker_path = os.path.join(pyrevit_path, 'asmmaker.py')

    if not os.path.exists(asmmaker_path):
        print(f"ERROR: Could not find pyRevit installation at:")
        print(f"  {pyrevit_path}")
        print()
        print("Please ensure pyRevit is installed.")
        return 1

    print(f"Found pyRevit installation:")
    print(f"  {pyrevit_path}")
    print()

    # Create backup
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = f"{asmmaker_path}.backup_{timestamp}"

    print("Creating backup...")
    if not whatif:
        shutil.copy2(asmmaker_path, backup_path)
        print(f"Backup created: {backup_path}")
    else:
        print(f"[WHATIF] Would create backup: {backup_path}")
    print()

    # Read the current file
    with open(asmmaker_path, 'r') as f:
        content = f.read()

    # Check if already patched
    if "# pyRevit Reload Fix" in content:
        print("File appears to already be patched!")
        print("No changes needed.")
        return 0

    # Define the patch
    old_code = """        asm_builder.Save(
            asm_file_path,
            peke,
            imachine
        )"""

    new_code = """        # pyRevit Reload Fix - Assembly lock handling with retry logic
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
                        raise e"""

    # Apply the patch
    if old_code in content:
        print("Applying patch...")
        patched_content = content.replace(old_code, new_code)

        if not whatif:
            with open(asmmaker_path, 'w') as f:
                f.write(patched_content)
            print("Patch applied successfully!")
        else:
            print(f"[WHATIF] Would apply patch to: {asmmaker_path}")

        print()
        print("IMPORTANT: Please restart Revit for changes to take effect.")
    else:
        print("WARNING: Could not find the expected code pattern to patch.")
        print("The file may have been modified or is from a different pyRevit version.")
        print()
        print("Please apply the manual fix as described in PYREVIT_RELOAD_FIX.md")
        return 1

    print()
    print("=" * 60)
    print("  Patch Complete!")
    print("=" * 60)
    print()
    print(f"Backup location: {backup_path}")
    print()
    print("To revert the patch, run:")
    print(f"  copy \"{backup_path}\" \"{asmmaker_path}\"")

    return 0

if __name__ == "__main__":
    sys.exit(main())
