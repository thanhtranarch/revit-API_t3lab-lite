# -*- coding: utf-8 -*-
"""
T3Lab UI Standard Showcase Dialog
GUI classes for reviewing standard UI components and styling.

The showcase is the reference implementation of the **Revit-native** tool
window: Segoe UI 12, square hairline chrome, and every colour resolved from
``GUI.RevitTheme`` rather than typed into the XAML. That means this class has
one job beyond loading the file — call ``RevitTheme.apply()`` so the
``{DynamicResource T3Theme*}`` keys the XAML binds to actually exist. Without
it the window falls back to the literal light-theme brushes declared at the
top of ``Window.Resources`` and never follows Revit into dark mode.
"""

import os
import sys
import clr
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows import (Thickness, GridLength, GridUnitType,
                            HorizontalAlignment, VerticalAlignment, FontWeights,
                            MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, WindowState)
from System.Windows.Controls import (RowDefinition, ColumnDefinition, Border,
                                      StackPanel, TextBlock, TextBox, Button,
                                      ComboBox, ComboBoxItem, DataGrid, Orientation,
                                      DataGridTextColumn, DataGridCheckBoxColumn,
                                      ScrollViewer, TabControl, TabItem, CheckBox)
from System.Windows.Media import SolidColorBrush, BrushConverter
from System.Windows.Data import Binding
from System.Collections.ObjectModel import ObservableCollection

from pyrevit import revit, DB, forms

# Revit's own light/dark palette. Guarded: the showcase must still open (in
# light values) on a host where the theme bridge cannot import, rather than
# taking the whole window down over a colour lookup.
try:
    from GUI import RevitTheme as _theme
except Exception:
    try:
        import RevitTheme as _theme          # flat sys.path (deployed extension)
    except Exception:
        _theme = None

# XAML Path
# Single copy, in the extension — so this works when the extension is deployed
# on its own (e.g. %APPDATA%\pyRevit\Extensions) without the repo checkout.
# The repo-side duplicate under .claude/standard/ was removed on 2026-08-28
# together with the retired Lumina standard it demonstrated; the UI standard is
# now `pyRevit UI Design System/`, not this showcase.
GUI_DIR = os.path.dirname(__file__)  # [repo]/T3Lab.extension/lib/GUI
XAML_FILE = os.path.join(GUI_DIR, 'Tools', 'UIStandardShowcase.xaml')


class ShowcaseItem(object):
    """Một hàng của bảng P4 trong showcase.

    `T3.StatusPill` cần đúng 2 property: StatusText (chữ) và Severity
    ("Success" / "Warning" / "Danger"). Trạng thái KHÔNG BAO GIỜ chỉ bằng màu —
    chấm màu luôn đi kèm chữ.
    """

    def __init__(self, item_id, name, category, template, status_text,
                 severity, issues):
        self._id = item_id
        self._name = name
        self._category = category
        self._template = template
        self._status_text = status_text
        self._severity = severity
        self._issues = issues

    @property
    def id(self):
        return self._id

    @property
    def name(self):
        return self._name

    @property
    def category(self):
        return self._category

    @property
    def template(self):
        return self._template

    @property
    def issues(self):
        return self._issues

    # T3.StatusPill bind vào 2 tên này
    @property
    def StatusText(self):
        return self._status_text

    @property
    def Severity(self):
        return self._severity

    # Giữ lại cho code cũ còn đọc .status
    @property
    def status(self):
        return self._status_text


class UIShowcaseWindow(forms.WPFWindow):
    """Window class for T3Lab UI Standard Showcase.

    The window has no caption of its own: the XAML declares
    ``WindowStyle="SingleBorderWindow"`` and lets Windows draw the title
    bar, the way ``Autodesk.UI.Windows.ChildWindow`` does for Revit's own
    dialogs. There are therefore no minimise/maximise/close handlers here
    — the OS owns those buttons. ``close_button_clicked`` stays, for the
    OK and Cancel buttons.
    """

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._skin = None
        self._pinned = False        # True while previewing the non-host skin
        self._adopt_host_font()
        self._apply_theme()
        self._load_sample_data()
        # Revit 2024+ raises UIApplication.ThemeChanged, but that event does not
        # exist on 2023 and needs a UIApplication we do not hold here. Re-syncing
        # when the window is activated covers the same case with no version
        # dependency: flip Revit's theme, click back into the tool, it follows.
        try:
            self.Activated += self._window_activated
        except Exception:
            pass

    # ── Theme ────────────────────────────────────────────────────────────────

    def _adopt_host_font(self):
        """Take the OS message font instead of the hardcoded Segoe UI 12.

        The XAML declares Segoe UI 12 so it still renders standalone, but that
        is only correct on a default Windows. Revit's dialogs use the shell's
        message font, which changes with the user's text-scaling setting and
        with some locales — so read it and let the whole window inherit.
        """
        if _theme is None:
            return
        family, size = _theme.host_font()
        if family is None:
            return
        try:
            self.FontFamily = family
            if size and size > 0:
                self.FontSize = size
        except Exception:
            pass

    def _apply_theme(self, theme=None):
        """Write the T3Theme* brushes into Window.Resources.

        Every colour in the XAML is a DynamicResource against these keys, so a
        second call with a different theme re-skins the live window — no
        rebuild, no reload of the XAML. Where Revit exposes its own brush for a
        token, that is what lands here; see the host bridge in RevitTheme.
        """
        if _theme is None:
            return
        try:
            self._skin = _theme.apply(self, theme)
        except Exception as ex:
            print("Warning: theme not applied: {}".format(ex))
            return
        try:
            self.chk_dark_preview.IsChecked = (self._skin == 'dark')
        except Exception:
            pass
        self._report_palette_source()

    def _report_palette_source(self):
        """Show how much of the window is Revit's own palette, not our copy.

        The whole point of the host bridge is that these colours are the host's,
        so the count belongs where it can be checked at a glance rather than in
        a log nobody opens. On a version that exposes nothing it reads 0/N,
        which is the honest answer, not a failure.
        """
        try:
            report = _theme.host_report(self._skin)
        except Exception:
            return
        live = sum(1 for _v, src in report.values() if src == 'revit')
        state = "Previewing {}".format(self._skin) if self._pinned else "Ready"

        if live:
            text = "{} · palette {}/{} live from Revit".format(
                state, live, len(report))
        else:
            # Not a failure. The curated values ARE Revit's, decoded from
            # UIFramework.dll, so this is the same palette by a different route.
            # What matters is that it is not a MIXTURE: a few live values among
            # many copied ones is what produces a colour that does not belong.
            try:
                usable, resolved, total = _theme.host_coverage(self._skin)
                why = _theme.host_error()
            except Exception:
                usable, resolved, total, why = True, 0, 0, None
            text = "{} · palette from T3Lab copy of Revit".format(state)
            if total and not usable:
                text += " (host answered {}/{}, below threshold)".format(
                    resolved, total)
            elif why:
                text += " ({})".format(why[:90])
        try:
            self.status_text.Text = text
        except Exception:
            pass

    def _window_activated(self, sender, e):
        """Follow Revit if its theme changed while the tool was in the back."""
        if _theme is None or self._pinned:
            return
        try:
            if _theme.current_theme() != self._skin:
                self._apply_theme()
        except Exception:
            pass

    def theme_preview_clicked(self, sender, e):
        """Preview the other skin without restarting Revit.

        Revit stays the real source of truth (``UIThemeManager.CurrentTheme``);
        this pins the palette so both skins can be reviewed side by side while
        designing a tool, and unpins as soon as the choice lands back on the
        host's own theme. It is an ordinary check box in an ordinary group box —
        it used to be a glyph button in a custom caption bar, which is not a
        control Revit has anywhere.
        """
        if _theme is None:
            return
        try:
            other = 'dark' if bool(self.chk_dark_preview.IsChecked) else 'light'
        except Exception:
            return
        # Unpin BEFORE asking, or current_theme() hands back the pin we set on
        # the previous click instead of what Revit is actually running.
        try:
            _theme.force_theme(None)
            host = _theme.current_theme()
        except Exception:
            host = None
        self._pinned = (other != host)
        try:
            _theme.force_theme(other if self._pinned else None)
        except Exception:
            pass
        self._apply_theme(other)

    def setup_icon(self):
        """No icon. Revit dialogs do not put one in the caption.

        pyRevit's own setup_icon() assigns the pushbutton PNG to Window.Icon,
        which is where the T3Lab logo in the corner came from. Revit's
        ChildWindow strips the caption icon instead (RemoveWindowIcon /
        WS_EX_DLGMODALFRAME), so overriding this to do nothing — together
        with WindowStyle="ToolWindow" in the XAML — is what matches it.
        """
        return

    def _unused_setup_icon(self):
        """Kept for reference: the previous icon-loading behaviour."""
        try:
            # Resolve the pushbutton's icon.png path relative to this file
            current_dir = os.path.dirname(__file__)  # lib/GUI
            extension_dir = os.path.dirname(os.path.dirname(current_dir))  # T3Lab.extension
            icon_path = os.path.join(extension_dir, "T3Lab.tab", "Standard.panel", "UIStandard.pushbutton", "icon.png")
            if os.path.exists(icon_path):
                self.set_icon(icon_path)
            else:
                # Fallback to default pyRevit icon
                super(UIShowcaseWindow, self).setup_icon()
        except Exception as e:
            # Silence icon loading errors to prevent window initialization crash
            print("Warning: Failed to load window icon: {}".format(e))

    # ── Window chrome ────────────────────────────────────────────────────────
    # The caption is ours again. Not by preference — a Revit dialog uses the
    # native one — but because the native caption follows the WINDOWS theme, so
    # Revit-dark on a light Windows renders a white title bar above a dark
    # window. Painting it ourselves is the only pure-WPF way to make it follow
    # Revit. See docs/revit-native-ui.md.

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
        """Caption close, and the OK / Cancel buttons."""
        self.Close()

    def _load_sample_data(self):
        """Dữ liệu mẫu cho bảng P4. Chữ hiển thị phải là TIẾNG ANH (luật 12)."""
        items = [
            ShowcaseItem("104231", "01_Plan_Column setting-out", "Floor Plan",
                         "ARC_Plan_1-100", "Compliant", "Success", "0"),
            ShowcaseItem("104232", "02_Plan_Column dimensions", "Floor Plan",
                         "ARC_Plan_1-100", "Compliant", "Success", "0"),
            ShowcaseItem("104235", "Detail_Column rebar splice", "Detail View",
                         "\u2014 none \u2014", "Template missing", "Danger", "1"),
            ShowcaseItem("104240", "3D_Overall perspective", "3D View",
                         "\u2014 none \u2014", "Naming standard", "Danger", "2"),
            ShowcaseItem("104245", "Section_Longitudinal section", "Section",
                         "ARC_Sect_1-50", "Needs review", "Warning", "1"),
            ShowcaseItem("104250", "Elevation_Grid A-D", "Elevation",
                         "ARC_Elev_1-100", "Compliant", "Success", "0"),
            ShowcaseItem("104255", "Schedule_Beam rebar take-off", "Schedule",
                         "n/a", "Compliant", "Success", "0"),
            ShowcaseItem("104261", "Detail_Pad footing M1", "Detail View",
                         "STR_Det_1-25", "Naming standard", "Danger", "1"),
            ShowcaseItem("104266", "Plan_Services ceiling", "Ceiling Plan",
                         "ARC_Plan_1-100", "Compliant", "Success", "0"),
            ShowcaseItem("104270", "Elevation_Grid 1-9", "Elevation",
                         "ARC_Elev_1-100", "Needs review", "Warning", "3"),
        ]
        self.sample_grid.ItemsSource = ObservableCollection[object](items)


def show_ui_standard_showcase():
    """Launch the UI Standard Showcase Dialog"""
    try:
        window = UIShowcaseWindow()
        window.ShowDialog()
    except Exception as e:
        print("\nFATAL ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        
        MessageBox.Show(
            "Error starting UI Standard Showcase:\n\n{}".format(str(e)),
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        )