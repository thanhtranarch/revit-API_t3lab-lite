# -*- coding: utf-8 -*-
"""Visual & Style Manager — event handling for the Visual & Style Manager launcher window."""

import os
import sys
import imp
import __builtin__

from pyrevit import forms
from System.Windows import WindowState

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'VisualSettings.xaml')


class VisualSettingsWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit
        self.doc = revit.ActiveUIDocument.Document

        # Dynamic imports and panel instantiation
        self._init_sub_panels()

        # Connect tab navigation events
        self.btn_tab_styles.Checked += self._on_tab_changed
        self.btn_tab_splasher.Checked += self._on_tab_changed

        # Chrome actions
        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close_chrome.Click += self._close_chrome

    def _init_sub_panels(self):
        """Import sub-tools and load their grids into the tabs."""
        # Find Settings.panel directory path
        gui_dir = os.path.dirname(__file__)
        lib_dir = os.path.dirname(gui_dir)
        ext_dir = os.path.dirname(lib_dir)
        panel_dir = os.path.join(ext_dir, "T3Lab.tab", "Settings.panel")

        # 1. Load Style Manager
        try:
            styles_dir = os.path.join(panel_dir, "ManaStyles")
            styles_mod = imp.load_source("mana_styles", os.path.join(styles_dir, "script.py"))
            self._styles_win = styles_mod.StyleManagerWindow()
            
            # Extract content and host it
            grid_content = self._styles_win.Content
            self._styles_win.Content = None
            self.grid_style_manager.Children.Add(grid_content)
            
            # Override Close method
            self._styles_win.Close = self.Close
        except Exception as ex:
            print("Error loading Style Manager Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

        # 2. Load Color Splasher
        try:
            splasher_dir = os.path.join(panel_dir, "ColorSplasher")
            splasher_mod = imp.load_source("color_splasher", os.path.join(splasher_dir, "script.py"))
            self._splasher_win = splasher_mod.ColorSplasherWindow()
            
            # Extract content and host it
            grid_content = self._splasher_win.Content
            self._splasher_win.Content = None
            self.grid_color_splasher.Children.Add(grid_content)
            
            # Override Close method
            self._splasher_win.Close = self.Close
        except Exception as ex:
            print("Error loading Color Splasher Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

    def _on_tab_changed(self, sender, e):
        """Switch active tab in TabControl when a RadioButton is checked."""
        if sender == self.btn_tab_styles:
            self.main_tab_control.SelectedIndex = 0
            self.status_text.Text = "Style Manager — Rename and manage Line Styles, Patterns, and Fill Patterns"
        elif sender == self.btn_tab_splasher:
            self.main_tab_control.SelectedIndex = 1
            self.status_text.Text = "Color Splasher — Dynamic parameter-based element color override"

    def _minimize(self, sender, e):
        self.WindowState = WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()


def show_visual_settings(script_dir, revit):
    VisualSettingsWindow(script_dir, revit).ShowDialog()
