# T3Lab Lite - Revit Productivity Extension

<div align="center">

**Professional IronPython Scripts for Autodesk Revit**

[![pyRevit](https://img.shields.io/badge/pyRevit-4.8+-blue.svg)](https://github.com/eirannejad/pyRevit)
[![Revit](https://img.shields.io/badge/Revit-2019--2024-orange.svg)](https://www.autodesk.com/products/revit)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

</div>

## 📋 Overview

T3Lab Lite is a comprehensive pyRevit extension providing essential productivity tools for Revit users. Built with clean, maintainable code following industry best practices.

**Author:** Tran Tien Thanh
**Email:** trantienthanh909@gmail.com
**LinkedIn:** [linkedin.com/in/sunarch7899/](https://linkedin.com/in/sunarch7899/)

---

## ✨ Features

### 📦 Project Panel
- **Load Family** - Browse and batch load families from local folders
- **Load Family (Cloud)** - Access cloud-based family library via Vercel API
- **Workset Manager** - Efficient workset creation and management
- **Dimension Text Tools** - Uppercase and formatting utilities for dimensions

### 🧹 TidyUp Panel
- **Dimension Tools**
  - Find Dimensions - Locate all dimensions in project
  - Remove Dimensions - Batch deletion with filters
  - Rename Dimensions - Bulk renaming operations

- **Text Note Tools**
  - Find Text Notes - Search text notes by criteria
  - Remove Text Notes - Batch deletion with filters
  - Rename Text Notes - Bulk renaming operations

### 📤 Export Panel
- **Batch Out** - Export multiple sheets to PDF/DWF with automation

### 🎨 Graphic Panel
- **Reset Overrides** - Clear graphic overrides in views

### 🔧 Temporary Tools Panel
- **Drilling Point Coordination** - MEP coordination utilities
- (Development/Testing tools)

---

## 🚀 Installation

### Prerequisites
- Autodesk Revit (2019 or later)
- [pyRevit](https://github.com/eirannejad/pyRevit) installed

### Steps

1. **Install pyRevit** (if not already installed)
   ```bash
   # Download from: https://github.com/eirannejad/pyRevit/releases
   # Run the installer
   ```

2. **Clone or Download this Repository**
   ```bash
   git clone https://github.com/thanhtranarch/revit-API_t3lab-lite.git
   ```

3. **Copy Extension to pyRevit**
   - Copy the `T3Lab_Lite.extension` folder to your pyRevit extensions directory:
     - **Default location:** `%APPDATA%\pyRevit\Extensions\`
     - **Or use:** `pyRevit → Extensions → Look up extension paths`

4. **Reload pyRevit**
   - In Revit: pyRevit → Reload
   - The "T3Lab Lite" tab should appear in the Revit ribbon

---

## ⚙️ Configuration

### Family Loader Cloud Setup

For cloud family loading, create a configuration file:

**Location:** `~/.t3lab/family_loader_config.json`

**Content:**
```json
{
  "cloud_api_base": "https://your-deployment.vercel.app",
  "cloud_api_endpoint": "/api/families",
  "vercel_bypass_token": "your-bypass-token-here"
}
```

**Security Note:** Never commit credentials to version control. See `.env.example` for template.

---

## 📖 Usage Examples

### Loading Families from Cloud

1. Click **Load Family (Cloud)** button
2. Browse cloud library (automatically loads on startup)
3. Search and filter families by category
4. Select families and click "Load"
5. Families are downloaded and loaded automatically

### Batch Dimension Cleanup

1. Click **TidyUp → Dimension → Remove Dim**
2. Set filter criteria (view, type, etc.)
3. Preview dimensions to be removed
4. Confirm deletion

### Export Sheets to PDF

1. Click **Export → Batch Out**
2. Select sheets to export
3. Configure PDF settings
4. Choose output folder
5. Click "Export" for batch processing

---

## 🛠️ Development

### Project Structure

```
T3Lab_Lite.extension/
├── T3Lab_Lite.tab/           # Main ribbon tab
│   ├── Project.panel/        # Project tools
│   ├── FJX TidyUp.panel/    # Cleanup utilities
│   ├── Export.panel/         # Export tools
│   └── Graphic.panel/        # Graphics tools
├── lib/                      # Shared libraries
│   ├── GUI/                  # WPF dialogs
│   ├── Create/               # Creation utilities
│   ├── Selection/            # Selection helpers
│   └── Snippets/             # Reusable code
├── checks/                   # Model checker plugins
└── extension.json            # Extension metadata
```

### Code Standards

- **Format:** All code follows standardized format (see templates)
- **Strings:** Use `.format()` method (no f-strings for IronPython compatibility)
- **Comments:** English only
- **Logging:** Use proper log levels without DEBUG prefix
- **Documentation:** Comprehensive docstrings

### Running Tests

```bash
# (Tests to be added in future release)
```

---

## 🐛 Known Issues

### pyRevit Reload Error

If you encounter an `IOError` when trying to reload pyRevit:

**Quick Fix:**
```powershell
# Run as Administrator
PowerShell -ExecutionPolicy Bypass -File scripts/fix_pyrevit_reload.ps1
```

**Details:** See [PYREVIT_RELOAD_FIX.md](PYREVIT_RELOAD_FIX.md)

---

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Follow code standards (see Development section)
4. Test your changes in Revit
5. Submit a pull request

---

## 📝 Changelog

### Version 1.0.0 (Current)
- ✅ Complete standardization of code format
- ✅ Security improvements (credential management)
- ✅ Enhanced error handling and logging
- ✅ Comprehensive documentation
- ✅ Code cleanup and optimization

See full changelog: [CHANGELOG.md](CHANGELOG.md) _(to be created)_

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 👤 Author

**Tran Tien Thanh**
- Email: trantienthanh909@gmail.com
- LinkedIn: [linkedin.com/in/sunarch7899/](https://linkedin.com/in/sunarch7899/)
- GitHub: [@thanhtranarch](https://github.com/thanhtranarch)

---

## 🙏 Acknowledgments

- [pyRevit](https://github.com/eirannejad/pyRevit) - Amazing Revit scripting framework
- Autodesk Revit API - Comprehensive building design platform
- All contributors and users

---

## 📞 Support

- **Issues:** [GitHub Issues](https://github.com/thanhtranarch/revit-API_t3lab-lite/issues)
- **Discussions:** [GitHub Discussions](https://github.com/thanhtranarch/revit-API_t3lab-lite/discussions)
- **Email:** trantienthanh909@gmail.com

---

<div align="center">

**Made with ❤️ for the Revit community**

⭐ Star this repo if you find it helpful!

</div>
