# -*- coding: utf-8 -*-
"""Model Auditor — event handling for the Model Auditor launcher window."""

import os
import sys
import imp
import __builtin__

from pyrevit import forms
from System.Windows import WindowState

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ModelAuditor.xaml')


class ModelAuditorWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit
        self.doc = revit.ActiveUIDocument.Document
        self.uidoc = revit.ActiveUIDocument

        # Dynamic imports and panel instantiation
        self._init_sub_panels()

        # Connect top tab navigation events
        self.btn_tab_health.Checked += self._on_tab_changed
        self.btn_tab_compliance.Checked += self._on_tab_changed
        self.btn_tab_warning.Checked += self._on_tab_changed
        self.btn_tab_elements.Checked += self._on_tab_changed

        # Connect sub-tab navigation events
        self.btn_sub_inplace.Checked += self._on_sub_tab_changed
        self.btn_sub_material.Checked += self._on_sub_tab_changed

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

        # 1. Load Health Check
        try:
            hc_dir = os.path.join(panel_dir, "HealthCheck")
            hc_mod = imp.load_source("health_check", os.path.join(hc_dir, "script.py"))
            self._health_check_win = hc_mod.ModelHealthWindow(self.doc, self.uidoc)
            
            # Extract content and host it
            grid_content = self._health_check_win.Content
            self._health_check_win.Content = None
            self.grid_health_check.Children.Add(grid_content)
            
            # Override Close method
            self._health_check_win.Close = self.Close
        except Exception as ex:
            print("Error loading Health Check Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

        # 2. Load Model Checker
        try:
            mc_dir = os.path.join(panel_dir, "ModelChecker")
            mc_mod = imp.load_source("model_checker", os.path.join(mc_dir, "script.py"))
            self._model_checker_win = mc_mod.ModelCheckerWindow()
            
            # Extract content and host it (ModelCheckerWindow wraps window in self.window)
            grid_content = self._model_checker_win.window.Content
            self._model_checker_win.window.Content = None
            self.grid_model_checker.Children.Add(grid_content)
            
            # Override Close method
            self._model_checker_win.window.Close = self.Close
        except Exception as ex:
            print("Error loading Model Checker Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

        # 3. Load Warning Manager
        try:
            warning_dir = os.path.join(panel_dir, "Warning")
            warning_mod = imp.load_source("warning_manager", os.path.join(warning_dir, "script.py"))
            self._warning_win = warning_mod.WarningManager()
            
            # Extract content and host it
            grid_content = self._warning_win.Content
            self._warning_win.Content = None
            self.grid_warning_manager.Children.Add(grid_content)
            
            # Override Close method
            self._warning_win.Close = self.Close
        except Exception as ex:
            print("Error loading Warning Manager Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

        # 4. Load In-Place Models
        try:
            inplace_dir = os.path.join(panel_dir, "InPlaceModel")
            inplace_mod = imp.load_source("inplace_model", os.path.join(inplace_dir, "script.py"))
            self._inplace_win = inplace_mod.InPlaceManagerWindow()
            
            # Extract content and host it
            grid_content = self._inplace_win.Content
            self._inplace_win.Content = None
            self.grid_inplace_model.Children.Add(grid_content)
            
            # Override Close method
            self._inplace_win.Close = self.Close
        except Exception as ex:
            print("Error loading In-Place Model Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

        # 5. Load Material List
        try:
            material_dir = os.path.join(panel_dir, "MaterialList")
            material_mod = imp.load_source("material_list", os.path.join(material_dir, "script.py"))
            self._material_win = material_mod.MaterialManagerWindow()
            
            # Extract content and host it
            grid_content = self._material_win.Content
            self._material_win.Content = None
            self.grid_material_list.Children.Add(grid_content)
            
            # Override Close method
            self._material_win.Close = self.Close
        except Exception as ex:
            print("Error loading Material List Panel: {}".format(ex))
            import traceback
            traceback.print_exc()

    def _on_tab_changed(self, sender, e):
        """Switch active tab in TabControl when a RadioButton is checked."""
        if sender == self.btn_tab_health:
            self.main_tab_control.SelectedIndex = 0
            self.status_text.Text = "Model Health Dashboard — Health score and summary metrics"
        elif sender == self.btn_tab_compliance:
            self.main_tab_control.SelectedIndex = 1
            self.status_text.Text = "Compliance Checker — JSON rule-based BEP auditing"
        elif sender == self.btn_tab_warning:
            self.main_tab_control.SelectedIndex = 2
            self.status_text.Text = "Warning Manager — Warnings impact analysis and auto-fixing"
        elif sender == self.btn_tab_elements:
            self.main_tab_control.SelectedIndex = 3
            # Set sub-tab status text
            if self.btn_sub_inplace.IsChecked:
                self.status_text.Text = "In-Place Model Auditor — Manage in-place family instances"
            else:
                self.status_text.Text = "Material List Auditor — Export and analyze model material usage data"

    def _on_sub_tab_changed(self, sender, e):
        """Switch sub-tab for Special Audits (In-Place vs Material List)."""
        if not hasattr(self, 'sub_tab_control'):
            return
        if sender == self.btn_sub_inplace:
            self.sub_tab_control.SelectedIndex = 0
            self.status_text.Text = "In-Place Model Auditor — Manage in-place family instances"
        elif sender == self.btn_sub_material:
            self.sub_tab_control.SelectedIndex = 1
            self.status_text.Text = "Material List Auditor — Export and analyze model material usage data"

    def _minimize(self, sender, e):
        self.WindowState = WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()


def show_model_auditor(script_dir, revit):
    ModelAuditorWindow(script_dir, revit).ShowDialog()
