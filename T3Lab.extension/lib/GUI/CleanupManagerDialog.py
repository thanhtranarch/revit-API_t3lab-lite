# -*- coding: utf-8 -*-
"""Cleanup Manager — event handling for the Cleanup Manager launcher window."""

import os
import sys
import imp
import __builtin__

from pyrevit import forms
from System.Windows import WindowState

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'CleanupManager.xaml')


class CleanupManagerWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit
        self.doc = revit.ActiveUIDocument.Document

        # Dynamic imports and panel instantiation
        self._init_sub_panels()

        # Connect tab navigation events
        self.btn_tab_smart_purge.Checked += self._on_tab_changed
        self.btn_tab_advanced_purge.Checked += self._on_tab_changed
        self.btn_tab_smart_delete.Checked += self._on_tab_changed

        # Chrome actions
        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close_chrome.Click += self._close_chrome

    def _init_sub_panels(self):
        """Import sub-tools and load their grids into the tabs."""
        # Find Settings.panel directory path
        # CleanupManagerDialog.py is in [repo]/T3Lab.extension/lib/GUI/
        # Settings.panel is in [repo]/T3Lab.extension/T3Lab.tab/Settings.panel/
        gui_dir = os.path.dirname(__file__)
        lib_dir = os.path.dirname(gui_dir)
        ext_dir = os.path.dirname(lib_dir)
        panel_dir = os.path.join(ext_dir, "T3Lab.tab", "Settings.panel")

        # 1. Load Smart Purge
        try:
            smart_purge_lib = os.path.join(panel_dir, "SmartPurge", "lib")
            if smart_purge_lib not in sys.path:
                sys.path.insert(0, smart_purge_lib)

            sp_mod = imp.load_source("smart_purge_window_v2", os.path.join(smart_purge_lib, "smart_purge_window_v2.py"))
            self._smart_purge_win = sp_mod.SmartPurgeWindowV2(self.doc)
            
            # Extract content and host it
            grid_content = self._smart_purge_win.Content
            self._smart_purge_win.Content = None
            self.grid_smart_purge.Children.Add(grid_content)
            
            # Override Close method
            self._smart_purge_win.Close = self.Close
        except Exception as ex:
            print("Error loading Smart Purge Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

        # 2. Load Advanced Purge
        try:
            adv_purge_lib = os.path.join(panel_dir, "AdvancedPurge", "lib")
            if adv_purge_lib not in sys.path:
                sys.path.insert(0, adv_purge_lib)

            ap_mod = imp.load_source("advanced_purge_window", os.path.join(adv_purge_lib, "advanced_purge_window.py"))
            self._adv_purge_win = ap_mod.AdvancedPurgeWindow(self.doc)
            
            # Extract content and host it
            grid_content = self._adv_purge_win.Content
            self._adv_purge_win.Content = None
            self.grid_advanced_purge.Children.Add(grid_content)
            
            # Override Close method
            self._adv_purge_win.Close = self.Close
        except Exception as ex:
            print("Error loading Advanced Purge Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

        # 3. Load Smart Delete
        try:
            sd_dir = os.path.join(panel_dir, "SmartDelete")
            sd_mod = imp.load_source("smart_delete", os.path.join(sd_dir, "script.py"))
            self._smart_delete_win = sd_mod.SmartDeleteWindow()
            
            # Extract content and host it
            grid_content = self._smart_delete_win.Content
            self._smart_delete_win.Content = None
            self.grid_smart_delete.Children.Add(grid_content)
            
            # Override Close method
            self._smart_delete_win.Close = self.Close
        except Exception as ex:
            print("Error loading Smart Delete Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

    def _on_tab_changed(self, sender, e):
        """Switch active tab in TabControl when a RadioButton is checked."""
        if sender == self.btn_tab_smart_purge:
            self.main_tab_control.SelectedIndex = 0
            self.status_text.Text = "Smart Purge — Safely purge unused elements"
        elif sender == self.btn_tab_advanced_purge:
            self.main_tab_control.SelectedIndex = 1
            self.status_text.Text = "Advanced Purge — Deep purging (requires caution)"
        elif sender == self.btn_tab_smart_delete:
            self.main_tab_control.SelectedIndex = 2
            self.status_text.Text = "Smart Delete — Dependency-checked deletion"

    def _minimize(self, sender, e):
        self.WindowState = WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()


def show_cleanup_manager(script_dir, revit):
    CleanupManagerWindow(script_dir, revit).ShowDialog()
