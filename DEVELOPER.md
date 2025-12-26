# Developer Documentation

## T3Lab Lite Extension - Development Guide

**Author:** Tran Tien Thanh
**Email:** trantienthanh909@gmail.com
**LinkedIn:** [linkedin.com/in/sunarch7899/](https://linkedin.com/in/sunarch7899/)

---

## 📐 Code Standards

### File Format Template

All Python files must follow this standard format:

```python
# -*- coding: utf-8 -*-
"""
Tool Name

Description of what the tool does.
Additional details about functionality.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__  = "Tool Name"

# IMPORT LIBRARIES
# ==================================================
import os
import sys

# DEFINE VARIABLES
# ==================================================
logger = script.get_logger()

# HELPER FUNCTIONS
# ==================================================

def helper_function():
    """Function description"""
    pass

# CLASSES
# ==================================================

class MyClass:
    """Class description"""
    pass

# MAIN SCRIPT
# ==================================================

if __name__ == '__main__':
    # Main execution
    pass
```

### Coding Conventions

#### 1. String Formatting
**Always use `.format()` method** - IronPython doesn't support f-strings

```python
# ✅ GOOD
message = "Loading {} families from {}".format(count, folder)

# ❌ BAD
message = f"Loading {count} families from {folder}"
```

#### 2. Comments
**All comments must be in English**

```python
# ✅ GOOD
# Load family into the current document

# ❌ BAD
# Tải family vào tài liệu hiện tại
```

#### 3. Logging
**Use proper log levels without DEBUG prefix**

```python
# ✅ GOOD
logger.info("Starting family scan")
logger.debug("Scanned {} files".format(count))
logger.error("Failed to load family: {}".format(ex))

# ❌ BAD
logger.info("DEBUG: Starting family scan")
```

#### 4. Error Handling
**Always include comprehensive error handling**

```python
# ✅ GOOD
try:
    result = risky_operation()
    logger.info("Operation succeeded")
except SpecificException as ex:
    logger.error("Specific error: {}".format(ex))
    logger.error(traceback.format_exc())
except Exception as ex:
    logger.error("Unexpected error: {}".format(ex))
    logger.error(traceback.format_exc())
    forms.alert("Error: {}".format(ex))
```

---

## 🏗️ Architecture

### Extension Structure

```
T3Lab_Lite.extension/
├── extension.json                 # Extension metadata
├── T3Lab_Lite.tab/               # Ribbon tab definition
│   ├── bundle.yaml               # Tab configuration
│   ├── Project.panel/            # Panel for project tools
│   │   ├── bundle.yaml
│   │   ├── LoadFamily.pushbutton/
│   │   │   ├── script.py         # Button script
│   │   │   └── icon.png          # Button icon (32x32)
│   │   └── ...
│   └── ...
├── lib/                          # Shared libraries
│   ├── GUI/                      # WPF dialogs
│   │   ├── FamilyLoaderDialog.py
│   │   ├── FamilyLoader.xaml
│   │   └── ...
│   ├── Create/                   # Creation utilities
│   ├── Selection/                # Selection helpers
│   ├── Renaming/                 # Renaming utilities
│   └── Snippets/                 # Reusable code snippets
├── checks/                       # Model checker plugins
└── commands/                     # Custom commands
```

### Library Organization

#### GUI Module (`lib/GUI/`)
WPF-based user interfaces for complex tools.

**Key files:**
- `FamilyLoaderDialog.py` - Local family browser
- `FamilyLoaderCloudDialog.py` - Cloud family browser
- `ParameterSelectorDialog.py` - Parameter selection UI
- `SelectFromDict.py` - Dictionary selection dialog

**Pattern:**
```python
# dialog_name.py
class MyWindow(Window):
    def __init__(self):
        # Load XAML
        xaml_path = os.path.join(os.path.dirname(__file__), 'MyWindow.xaml')
        with open(xaml_path, 'r') as f:
            self.ui = XamlReader.Parse(f.read())
        self.Content = self.ui

        # Get controls
        self.btn_ok = self.ui.FindName('btn_ok')

        # Wire events
        self.btn_ok.Click += self.on_ok_clicked

    def on_ok_clicked(self, sender, e):
        # Handle click
        pass

def show_dialog():
    window = MyWindow()
    window.ShowDialog()
```

#### Selection Module (`lib/Selection/`)
Helpers for element selection and filtering.

#### Snippets Module (`lib/Snippets/`)
Reusable code patterns for common Revit operations.

---

## 🔧 Development Workflow

### 1. Setting Up Development Environment

```bash
# Clone repository
git clone https://github.com/thanhtranarch/revit-API_t3lab-lite.git

# Link to pyRevit extensions folder (Windows)
mklink /D "%APPDATA%\pyRevit\Extensions\T3Lab_Lite.extension" "path\to\revit-API_t3lab-lite\T3Lab_Lite.extension"
```

### 2. Creating a New Tool

#### Step 1: Create Button Folder
```
T3Lab_Lite.tab/Panel.panel/MyTool.pushbutton/
├── script.py
├── icon.png (32x32 pixels)
└── tooltip.png (optional, for extended tooltip)
```

#### Step 2: Write Script
```python
# -*- coding: utf-8 -*-
"""
My Tool

Description of what this tool does.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__  = "My Tool"

# IMPORT LIBRARIES
# ==================================================
from pyrevit import revit, DB, forms, script

# DEFINE VARIABLES
# ==================================================
logger = script.get_logger()
doc = revit.doc

# MAIN SCRIPT
# ==================================================

if __name__ == '__main__':
    try:
        logger.info("Running My Tool")

        # Your code here

        logger.info("Tool completed successfully")
    except Exception as ex:
        logger.error("Error: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
```

#### Step 3: Test in Revit
1. Reload pyRevit (pyRevit → Reload)
2. Test the button
3. Check output window for errors

### 3. Creating WPF Dialogs

#### Create XAML File
```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="My Dialog" Height="400" Width="600">
    <Grid>
        <Button x:Name="btn_ok" Content="OK" />
    </Grid>
</Window>
```

#### Create Python Backend
```python
from System.Windows import Window
from System.Windows.Markup import XamlReader

class MyDialog(Window):
    def __init__(self):
        # Load XAML
        xaml_path = os.path.join(os.path.dirname(__file__), 'MyDialog.xaml')
        with open(xaml_path, 'r') as f:
            self.ui = XamlReader.Parse(f.read())

        self.Content = self.ui
        self.btn_ok = self.ui.FindName('btn_ok')
        self.btn_ok.Click += self.on_ok

    def on_ok(self, sender, e):
        self.DialogResult = True
        self.Close()
```

---

## 🧪 Testing

### Manual Testing Checklist

For each tool:
- [ ] Test in empty project
- [ ] Test in project with data
- [ ] Test error cases (no selection, invalid input, etc.)
- [ ] Check log output for warnings/errors
- [ ] Verify UI responsiveness
- [ ] Test with different Revit versions (if applicable)

### Common Test Scenarios

**Selection-based tools:**
- No selection
- Single element
- Multiple elements
- Mixed element types

**Document modification tools:**
- Read-only document
- Workshared document
- Project vs Family document

**UI dialogs:**
- Cancel button
- Empty inputs
- Invalid inputs
- Large datasets

---

## 🐛 Debugging

### Using pyRevit Output Window

```python
logger.debug("Debug message - only in debug mode")
logger.info("Informational message")
logger.warning("Warning message")
logger.error("Error message")

# Print to output
print("Direct output")
```

### Using Visual Studio for Debugging

1. Add this to your script:
```python
import ptvsd
ptvsd.enable_attach(address=('localhost', 5678))
ptvsd.wait_for_attach()
```

2. Attach Visual Studio debugger to port 5678

### Common Issues

**Issue: Import errors**
```python
# Add lib to path
import sys
import os
lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)
```

**Issue: XAML not loading**
```python
# Use absolute path
xaml_path = os.path.join(os.path.dirname(__file__), 'MyDialog.xaml')
if not os.path.exists(xaml_path):
    logger.error("XAML not found: {}".format(xaml_path))
```

---

## 📦 Building Releases

### Version Numbering

Follow [Semantic Versioning](https://semver.org/):
- **MAJOR:** Breaking changes
- **MINOR:** New features (backward compatible)
- **PATCH:** Bug fixes

### Release Checklist

- [ ] Update version in `extension.json`
- [ ] Update `CHANGELOG.md`
- [ ] Test all tools
- [ ] Update documentation
- [ ] Create git tag
- [ ] Push to GitHub

```bash
# Tag release
git tag -a v1.0.0 -m "Release version 1.0.0"
git push origin v1.0.0
```

---

## 🔐 Security Best Practices

### Never Hardcode Credentials

```python
# ❌ BAD
API_TOKEN = "secret-token-12345"

# ✅ GOOD
def get_api_token():
    config = load_config()
    return config.get('api_token', '')
```

### Use Configuration Files

Store sensitive data in `~/.t3lab/family_loader_config.json`:
```json
{
  "api_token": "your-token-here",
  "api_url": "https://your-api.com"
}
```

### Add to .gitignore
```gitignore
# Configuration with credentials
*.config.json
.env
*.token
```

---

## 📚 Resources

### Revit API Documentation
- [Revit API Docs](https://www.revitapidocs.com/)
- [The Building Coder](https://thebuildingcoder.typepad.com/)
- [Boost Your BIM](https://boostyourbim.wordpress.com/)

### pyRevit
- [pyRevit Documentation](https://pyrevitlabs.notion.site/pyRevit-bd907d6292ed4ce997c46e84b6ef67a0)
- [pyRevit GitHub](https://github.com/eirannejad/pyRevit)

### IronPython
- [IronPython Documentation](https://ironpython.net/)
- [Python 2.7 Docs](https://docs.python.org/2.7/) (IronPython is based on Python 2.7)

---

## 🤝 Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for detailed contribution guidelines.

**Quick Start:**
1. Fork the repository
2. Create a feature branch
3. Follow code standards
4. Test thoroughly
5. Submit pull request

---

## 📞 Contact

**Tran Tien Thanh**
- Email: trantienthanh909@gmail.com
- LinkedIn: [linkedin.com/in/sunarch7899/](https://linkedin.com/in/sunarch7899/)
- GitHub: [@thanhtranarch](https://github.com/thanhtranarch)

---

**Last Updated:** December 2024
