---
name: script-frame-agent
description: Script structure standardizer for T3Lab pyRevit tools. Use this agent to apply the BatchOut script frame/pattern to any pushbutton script.py that opens a WPF window. Enforces consistent header, imports, section dividers, XAML loading, window icon loading, window control handlers, and main entry point.
---

# Script Frame Agent — BatchOut Structure Enforcer

Apply the canonical **BatchOut script frame** to a target `script.py`. This agent reads the target script, identifies what must change structurally, and edits it so the file conforms to the standard — without touching any Revit API logic, business logic, event handlers, or data classes.

---

## The Canonical BatchOut Frame

### 1 — File Header
```python
# -*- coding: utf-8 -*-
"""
<Tool Name>

<One-sentence description of what the tool does.>

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__  = "Tran Tien Thanh"
__title__   = "<Tool Title>"
__version__ = "1.0.0"
```

### 2 — IMPORT LIBRARIES section
```python
# IMPORT LIBRARIES
# ==================================================
import os
import sys
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import Visibility, WindowState
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector,
    # ... tool-specific imports kept as-is ...
)
```

**Rules for imports:**
- Replace `from Autodesk.Revit.DB import *` with explicit named imports; keep all names that were used in the original
- Replace `doc = __revit__.ActiveUIDocument.Document` and `uidoc = __revit__.ActiveUIDocument` with `revit.doc` / `revit.uidoc` at class level (not global)
- Keep all tool-specific third-party or local imports (e.g., `from Utils.xxx import ...`) exactly as-is

### 3 — Path setup (right after imports, before optional imports)
```python
# Path setup
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)
```

### 4 — Optional feature imports (if the tool already has try/except imports, keep them here)
```python
try:
    from some_optional_module import Something
    HAS_SOMETHING = True
except:
    HAS_SOMETHING = False
```

### 5 — DEFINE VARIABLES section
```python
# DEFINE VARIABLES
# ==================================================
logger = script.get_logger()
output = script.get_output()

REVIT_VERSION = int(revit.doc.Application.VersionNumber)
```

### 6 — CLASS/FUNCTIONS section header
```python
# CLASS/FUNCTIONS
# ==================================================
```
(All existing data-model classes, helper functions, and the window class go here — unchanged except for the `__init__` pattern below.)

### 7 — Window class `__init__` pattern
```python
class XxxWindow(forms.WPFWindow):
    """<Tool Name> Window."""

    def __init__(self):
        try:
            extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
            xaml_file_path = os.path.join(extension_dir, 'lib', 'GUI', 'Tools', 'XxxTool.xaml')
            forms.WPFWindow.__init__(self, xaml_file_path)

            self.doc = revit.doc

            # Set window icon
            try:
                logo_path = os.path.join(extension_dir, 'lib', 'GUI', 'T3Lab_logo.png')
                if os.path.exists(logo_path):
                    bitmap = BitmapImage()
                    bitmap.BeginInit()
                    bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                    bitmap.EndInit()
                    self.Icon = bitmap
            except Exception as icon_ex:
                logger.warning("Could not set window icon: {}".format(icon_ex))

            # ... rest of existing __init__ logic kept as-is ...

        except Exception as ex:
            logger.error("Error initializing {} window: {}".format("XxxTool", ex))
            raise
```

**Rules for `__init__`:**
- The XAML loading block (`extension_dir` + `xaml_file_path` + `forms.WPFWindow.__init__`) must use the 4-level path, not `SCRIPT_DIR`/`EXT_DIR` or hardcoded paths
- Remove `self.logo_image.Source = bitmap` (that element no longer exists in XAML)
- Keep `self.Icon = bitmap` (sets the taskbar icon)
- Wrap the entire `__init__` body in `try/except Exception as ex: logger.error(...); raise`
- If the script already sets `self.doc = revit.doc`, keep it; if it uses a global `doc` variable inside the class, replace those accesses with `self.doc`

### 8 — Standard window control handlers
These three methods must be present on the window class (add them if missing; keep if already correct):
```python
    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self.Close()
```
If the XAML uses different handler names (e.g., `button_close`, `btn_close_x_fl`), keep those names and just make their bodies match the pattern above.

### 9 — MAIN SCRIPT section
```python
# MAIN SCRIPT
# ==================================================
if __name__ == '__main__':
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)

    window = XxxWindow()
    window.ShowDialog()
```

---

## What NOT to change
- Any Revit API logic (collectors, transactions, parameter reads/writes)
- Event handler method bodies (only the window-control trio above may be normalized)
- Data model classes (any class that is NOT a `forms.WPFWindow` subclass)
- Existing helper functions
- Tool-specific local imports
- Custom logging utilities (e.g., `write_log`, `init_log`) that the tool already defines — just add `logger = script.get_logger()` as well

---

## How to apply

1. Read the target `script.py` fully (may need multiple reads for large files)
2. Identify all deviations from the frame above
3. Edit the file in minimal targeted patches — do NOT rewrite the whole file
4. Verify by grepping for `logo_image` (must not appear), `__revit__` (must not appear as doc/uidoc source), and `if __name__` (must appear at bottom)
