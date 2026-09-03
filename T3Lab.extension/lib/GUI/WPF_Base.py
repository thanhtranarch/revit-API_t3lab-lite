# -*- coding: utf-8 -*-
"""
WPF Base
========
Universal WPF Window base class and loader for pyRevit under CPython 3 and IronPython.

Features:
  1. Full CPython 3 (Python.NET) compatibility via System.Windows.Markup.XamlReader.
  2. Automatic extraction and dynamic binding of XAML event handlers (Click, SelectionChanged, etc.).
  3. Automatic binding of all named elements (x:Name / Name) to instance attributes.
  4. Automatic Revit window ownership and host styling.
  5. Backward compatibility with IronPython 2.7/3.x.
  6. Auto-monkeypatch for pyrevit.forms.WPFWindow in CPython so existing dialogs run seamlessly.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__  = "Tran Tien Thanh"
__title__   = "WPF Base"

import os
import io
import re
import sys

# ─── CLR References ────────────────────────────────────────────────────────────
try:
    import clr
    for _ref in (
        "System",
        "System.IO",
        "System.Xml",
        "WindowsBase",
        "PresentationCore",
        "PresentationFramework",
        "System.Xaml",
        "System.Diagnostics.Process",
    ):
        try:
            clr.AddReference(_ref)
        except Exception:
            pass
except Exception:
    clr = None

try:
    from System.Windows import Window, WindowState, Visibility
except Exception:
    class Window(object):
        pass
    class WindowState(object):
        Normal = 0
        Minimized = 1
        Maximized = 2
    class Visibility(object):
        Visible = 0
        Hidden = 1
        Collapsed = 2

try:
    from System.Windows.Input import Key
except Exception:
    Key = None

try:
    from System.Windows.Interop import WindowInteropHelper
except Exception:
    WindowInteropHelper = None

try:
    from System.Windows.Markup import XamlReader
except Exception:
    try:
        import System.Windows.Markup
        XamlReader = System.Windows.Markup.XamlReader
    except Exception:
        XamlReader = None

try:
    from System.IO import StringReader
except Exception:
    try:
        import System.IO
        StringReader = System.IO.StringReader
    except Exception:
        StringReader = None

try:
    from System.Xml import XmlReader
except Exception:
    try:
        import System.Xml
        XmlReader = System.Xml.XmlReader
    except Exception:
        XmlReader = None

try:
    from System.Diagnostics.Process import Start
except Exception:
    def Start(url):
        import os
        try:
            os.startfile(url)
        except Exception:
            pass

# Check if running in IronPython
try:
    from pyrevit.compat import IRONPY
except Exception:
    IRONPY = 'IronPython' in sys.version

# Set of standard WPF event attribute names extracted from XAML
_EVENT_NAMES = {
    'Click', 'Checked', 'Unchecked', 'SelectionChanged', 'TextChanged',
    'MouseLeftButtonDown', 'MouseLeftButtonUp', 'MouseRightButtonDown', 'MouseRightButtonUp',
    'MouseDoubleClick', 'MouseEnter', 'MouseLeave', 'MouseMove', 'MouseWheel',
    'Loaded', 'Unloaded', 'Closing', 'Closed', 'Initialized',
    'ContentRendered', 'Activated', 'Deactivated', 'StateChanged', 'LocationChanged', 'SourceInitialized',
    'KeyDown', 'KeyUp', 'PreviewKeyDown', 'PreviewKeyUp', 'PreviewTextInput',
    'DropDownClosed', 'DropDownOpened', 'ValueChanged', 'Drop', 'DragOver',
    'PreviewMouseLeftButtonDown', 'PreviewMouseLeftButtonUp', 'MouseDown', 'MouseUp',
    'LostFocus', 'GotFocus', 'ScrollChanged', 'SizeChanged',
    'CellEditEnding', 'RowEditEnding', 'BeginningEdit', 'SelectedDateChanged'
}


def _sanitize_xaml(xaml_content):
    """
    Sanitize XAML for CPython XamlReader:
    Strips event handler attributes and records them with target element names.
    Returns (clean_xaml_string, event_bindings, named_elements).
    """
    event_bindings = []
    named_elements = set()
    counter = [0]

    for m in re.finditer(r'(?:x:)?Name\s*=\s*"([^"]+)"', xaml_content):
        named_elements.add(m.group(1))

    def process_tag(match):
        tag = match.group(0)
        has_event = False
        for evt in _EVENT_NAMES:
            if re.search(r'\b' + evt + r'\s*=\s*"[^"]*"', tag):
                has_event = True
                break
        if not has_event:
            return tag

        is_root_window = tag.startswith('<Window')
        name_match = re.search(r'(?:x:)?Name\s*=\s*"([^"]+)"', tag)
        if name_match:
            elem_name = name_match.group(1)
        elif is_root_window:
            elem_name = "__root__"
        else:
            counter[0] += 1
            elem_name = "__t3_dyn_" + str(counter[0])
            named_elements.add(elem_name)
            tag = re.sub(r'^(<[A-Za-z0-9_.:]+)', r'\1 x:Name="' + elem_name + '"', tag)

        for evt in _EVENT_NAMES:
            evt_match = re.search(r'\b(' + evt + r')\s*=\s*"([^"]+)"', tag)
            if evt_match:
                event_name = evt_match.group(1)
                handler_name = evt_match.group(2)
                event_bindings.append((elem_name, event_name, handler_name))
                tag = re.sub(r'\s*\b' + evt + r'\s*=\s*"[^"]*"', '', tag)

        return tag

    clean_xaml = re.sub(r'<[A-Za-z0-9_.:]+(?:\s+[^>]*?)?/?>', process_tag, xaml_content)
    return clean_xaml, event_bindings, named_elements


class T3WPFWindow(Window):
    """
    Universal WPF Window base class supporting both CPython 3 and IronPython.
    """

    def __init__(self, xaml_source, literal_string=False, handle_esc=True, set_owner=True):
        try:
            Window.__init__(self)
        except Exception:
            pass
        self._xaml_source = xaml_source
        self.load_xaml(xaml_source, literal_string=literal_string,
                       handle_esc=handle_esc, set_owner=set_owner)

    def load_xaml(self, xaml_source, literal_string=False, handle_esc=True, set_owner=True):
        """Loads XAML and wires named elements + event handlers."""
        # Read XAML content
        if literal_string:
            xaml_content = xaml_source
        else:
            if not os.path.isabs(xaml_source):
                # Search relative to calling module or GUI/Tools
                _here = os.path.dirname(os.path.abspath(__file__))
                _tools = os.path.join(_here, "Tools")
                cand1 = os.path.join(_tools, os.path.basename(xaml_source))
                cand2 = os.path.abspath(xaml_source)
                xaml_path = cand1 if os.path.exists(cand1) else cand2
            else:
                xaml_path = xaml_source

            with io.open(xaml_path, 'r', encoding='utf-8') as f:
                xaml_content = f.read()

        if IRONPY:
            # IronPython engine: use wpf.LoadComponent if available
            try:
                import wpf
                if literal_string:
                    wpf.LoadComponent(self, StringReader(xaml_content))
                else:
                    wpf.LoadComponent(self, xaml_path)
            except Exception:
                self._load_via_xaml_reader(xaml_content)
        else:
            # CPython engine: use sanitized XamlReader
            self._load_via_xaml_reader(xaml_content)

        if set_owner:
            self.setup_owner()
        if handle_esc:
            self.setup_default_handlers()

    def _load_via_xaml_reader(self, xaml_content):
        """Hydrates Window via System.Windows.Markup.XamlReader."""
        clean_xaml, event_bindings, named_elements = _sanitize_xaml(xaml_content)

        loaded_win = None
        xr = XamlReader
        if xr is None:
            try:
                import clr
                clr.AddReference("PresentationFramework")
                from System.Windows.Markup import XamlReader as xr
            except Exception:
                xr = None

        # The real XAML error surfaces here. Swallowing it sent every failure
        # down the XmlReader fallback below, which then died with a misleading
        # "XmlReader has no attribute Create" and hid the actual cause.
        parse_error = None
        if xr is not None and hasattr(xr, 'Parse'):
            try:
                loaded_win = xr.Parse(clean_xaml)
            except Exception as ex:
                parse_error = ex
                loaded_win = None

        if loaded_win is None:
            sr = StringReader
            if sr is None:
                try:
                    from System.IO import StringReader as sr
                except Exception:
                    sr = None

            # On .NET Core / Revit 2026 the System.Xml facade can resolve to an
            # XmlReader without the static Create overloads, so test for the
            # member rather than for None.
            xml_r = XmlReader
            if xml_r is None or not hasattr(xml_r, 'Create'):
                try:
                    import clr
                    clr.AddReference("System.Xml")
                    clr.AddReference("System.Private.Xml")
                    from System.Xml import XmlReader as xml_r
                except Exception:
                    pass
                if xml_r is not None and not hasattr(xml_r, 'Create'):
                    xml_r = None

            if xr is not None and sr is not None and xml_r is not None:
                try:
                    string_reader = sr(clean_xaml)
                    xml_reader = xml_r.Create(string_reader)
                    loaded_win = xr.Load(xml_reader)
                except Exception as ex:
                    if parse_error is None:
                        parse_error = ex
                    loaded_win = None

        if loaded_win is None:
            if parse_error is not None:
                raise RuntimeError(
                    "Failed to parse XAML for {}:\n{}".format(
                        getattr(self, '_xaml_source', '<inline XAML>'), parse_error))
            raise RuntimeError("Failed to parse XAML via XamlReader. "
                               "Ensure PresentationFramework is loaded.")

        # Copy essential window properties
        for prop in ('Title', 'Width', 'Height', 'MinWidth', 'MinHeight',
                     'MaxWidth', 'MaxHeight', 'WindowStartupLocation',
                     'WindowStyle', 'AllowsTransparency', 'Background',
                     'ResizeMode', 'ShowInTaskbar', 'FontFamily', 'FontSize'):
            try:
                setattr(self, prop, getattr(loaded_win, prop))
            except Exception:
                pass

        # Copy WindowChrome if present
        try:
            from System.Windows.Shell import WindowChrome
            chrome = WindowChrome.GetWindowChrome(loaded_win)
            if chrome is not None:
                WindowChrome.SetWindowChrome(self, chrome)
        except Exception:
            pass

        try:
            self.Resources = loaded_win.Resources
        except Exception:
            pass

        # Transfer NameScope so FindName works on self
        try:
            from System.Windows.Markup import NameScope
            scope = NameScope.GetNameScope(loaded_win)
            if scope is not None:
                NameScope.SetNameScope(self, scope)
        except Exception:
            pass

        # Move content to self
        try:
            content = loaded_win.Content
            loaded_win.Content = None
            self.Content = content
        except Exception:
            pass

        # Bind all named elements as instance attributes
        for name in named_elements:
            try:
                elem = loaded_win.FindName(name)
                if elem is None:
                    elem = self.FindName(name)
                if elem is not None:
                    setattr(self, name, elem)
                    try:
                        self.RegisterName(name, elem)
                    except Exception:
                        pass
            except Exception:
                pass

        # Bind event handlers dynamically
        for elem_name, event_name, handler_name in event_bindings:
            try:
                if elem_name == '__root__':
                    ctrl = self
                else:
                    ctrl = getattr(self, elem_name, None)
                    if ctrl is None:
                        ctrl = loaded_win.FindName(elem_name) or self.FindName(elem_name)
                if ctrl is None:
                    continue
                handler = getattr(self, handler_name, None)
                if handler is not None and callable(handler):
                    evt = getattr(ctrl, event_name, None)
                    if evt is not None:
                        evt += handler
            except Exception:
                pass

        # Wire title bar drag if an element named 'title_bar' or similar exists
        for tb_name in ('title_bar', 'titlebar', 'border_titlebar', 'TitleBar'):
            tb = getattr(self, tb_name, None)
            if tb is not None:
                try:
                    tb.MouseLeftButtonDown += self._on_title_bar_mouse_down
                except Exception:
                    pass

    def setup_owner(self):
        """Sets host Revit application window as owner."""
        try:
            from pyrevit.api import AdWindows
            wih = WindowInteropHelper(self)
            wih.Owner = AdWindows.ComponentManager.ApplicationWindow
        except Exception:
            pass

    def setup_default_handlers(self):
        """Binds Escape key to close window."""
        try:
            self.PreviewKeyDown += self._handle_esc_key
        except Exception:
            pass

    def _handle_esc_key(self, sender, args):
        try:
            if args.Key == Key.Escape:
                self.Close()
        except Exception:
            pass

    def _on_title_bar_mouse_down(self, sender, e):
        try:
            from System.Windows.Input import MouseButtonState
            if e.LeftButton == MouseButtonState.Pressed:
                self.DragMove()
        except Exception:
            pass

    # ── Standard Window Operations ─────────────────────────────────────────

    def show(self, modal=False):
        """Show window."""
        if modal:
            return self.ShowDialog()
        self.Show()

    def show_dialog(self):
        """Show modal window."""
        return self.ShowDialog()

    def hide(self):
        """Hide window."""
        self.Hide()

    def close(self):
        """Close window."""
        self.Close()

    # ── Common Event Handlers ──────────────────────────────────────────────

    def button_close(self, sender, e):
        self.Close()

    def close_button_clicked(self, sender, e):
        self.Close()

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def Hyperlink_RequestNavigate(self, sender, e):
        try:
            Start(e.Uri.AbsoluteUri)
        except Exception:
            pass

    def _adopt_host_font(self):
        pass

    def _apply_theme(self):
        pass


class my_WPF(T3WPFWindow):
    """Legacy alias for backward compatibility."""
    pass


WPFWindow = T3WPFWindow


# ── Monkeypatch pyrevit.forms.WPFWindow for CPython ─────────────────────────
try:
    from pyrevit import forms
    if not getattr(forms, 'IRONPY', False):
        forms.WPFWindow = T3WPFWindow
except Exception:
    pass
