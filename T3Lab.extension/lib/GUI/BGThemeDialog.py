# -*- coding: utf-8 -*-
"""Background Theme Dialog class."""

import os
from pyrevit import forms
from System.Windows import WindowState
from System.Windows.Media import BrushConverter

# Absolute path to XAML
_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'BGTheme.xaml')

def brush(hex_string):
    return BrushConverter().ConvertFromString(hex_string)

class BackgroundThemeWindow(forms.WPFWindow):
    def __init__(self, r, g, b, presets, on_apply_callback):
        # WPFWindow.__init__ loads the XAML and hooks up handlers
        forms.WPFWindow.__init__(self, _XAML)
        
        self.presets = presets
        self.on_apply_callback = on_apply_callback
        self.applied = False
        self._suspend = False

        # Build preset buttons
        self._build_presets()

        # Wire events
        self.SliderR.ValueChanged += self._on_slider
        self.SliderG.ValueChanged += self._on_slider
        self.SliderB.ValueChanged += self._on_slider
        self.HexApply.Click += self._on_hex_apply
        self.BtnApply.Click += self._on_apply
        self.BtnClose.Click += self._on_close

        # Initial state
        self.set_rgb(r, g, b)

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

    def _build_presets(self):
        import System.Windows.Controls as WC
        from System.Windows import Thickness
        preset_style = self.Resources["TertiaryButton"]
        for name, rgb in self.presets:
            btn = WC.Button()
            btn.Content = name
            btn.Style = preset_style
            btn.Margin = Thickness(4)
            r_val, g_val, b_val = rgb
            btn.Tag = str(r_val) + "," + str(g_val) + "," + str(b_val)
            btn.Click += self._on_preset_click
            self.PresetGrid.Children.Add(btn)

    def _on_preset_click(self, sender, args):
        try:
            parts = sender.Tag.split(",")
            r, g, b = int(parts[0]), int(parts[1]), int(parts[2])
            self.set_rgb(r, g, b)
        except Exception:
            pass

    def set_rgb(self, r, g, b):
        r = self._clamp(r)
        g = self._clamp(g)
        b = self._clamp(b)
        self._suspend = True
        try:
            self.SliderR.Value = r
            self.SliderG.Value = g
            self.SliderB.Value = b
            self.LblR.Text = str(r)
            self.LblG.Text = str(g)
            self.LblB.Text = str(b)
            self.HexBox.Text = self._to_hex(r, g, b)
            self._update_preview(r, g, b)
        finally:
            self._suspend = False

    def _update_preview(self, r, g, b):
        hex_str = self._to_hex(r, g, b)
        self.PreviewBox.Background = brush(hex_str)
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        self.PreviewLbl.Foreground = brush("#FFFFFF") if luminance < 140 else brush("#0F172A")
        self.PreviewLbl.Text = hex_str

    def _on_slider(self, sender, args):
        if self._suspend:
            return
        r = int(self.SliderR.Value)
        g = int(self.SliderG.Value)
        b = int(self.SliderB.Value)
        self.LblR.Text = str(r)
        self.LblG.Text = str(g)
        self.LblB.Text = str(b)
        self._suspend = True
        try:
            self.HexBox.Text = self._to_hex(r, g, b)
        finally:
            self._suspend = False
        self._update_preview(r, g, b)

    def _on_hex_apply(self, sender, args):
        rgb = self._from_hex(self.HexBox.Text)
        if rgb is None:
            self.PreviewLbl.Text = "Invalid HEX"
            return
        self.set_rgb(*rgb)

    def current_rgb(self):
        return (int(self.SliderR.Value),
                int(self.SliderG.Value),
                int(self.SliderB.Value))

    def _on_apply(self, sender, args):
        r, g, b = self.current_rgb()
        if self.on_apply_callback:
            self.on_apply_callback(r, g, b)
        self.applied = True

    def _on_close(self, sender, args):
        self.Close()

    @staticmethod
    def _clamp(v):
        v = int(round(v))
        return max(0, min(v, 255))

    def _to_hex(self, r, g, b):
        return "#" + self._hex2(r) + self._hex2(g) + self._hex2(b)

    def _hex2(self, v):
        v = self._clamp(v)
        digits = "0123456789ABCDEF"
        return digits[v // 16] + digits[v % 16]

    @staticmethod
    def _from_hex(hex_string):
        if not hex_string:
            return None
        s = hex_string.strip().lstrip("#")
        if len(s) != 6:
            return None
        try:
            return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
        except Exception:
            return None

def show_bg_theme_dialog(r, g, b, presets, on_apply_callback):
    """Factory function to create and show the BG Theme dialog."""
    dlg = BackgroundThemeWindow(r, g, b, presets, on_apply_callback)
    dlg.ShowDialog()
    return dlg
