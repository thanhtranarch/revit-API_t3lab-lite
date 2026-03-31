# Python WPF Window Pattern

Every tool window class must follow this pattern:

```python
import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import WindowState
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind
from pyrevit import forms

class MyToolWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, "MyTool.xaml")
        self._load_logo()
        # ... init logic ...

    def _load_logo(self):
        """Load T3Lab logo into the title bar and window icon."""
        try:
            ext_dir   = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
            logo_path = os.path.join(ext_dir, 'lib', 'GUI', 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                self.logo_image.Source = bitmap
                self.Icon = bitmap
        except Exception:
            pass

    # Required window chrome handlers
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
