# -*- coding: utf-8 -*-
"""ManaAlign — unified controller for alignment and dimensioning tools."""

import os
import sys
from pyrevit import forms, revit, DB, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    OverrideGraphicSettings,
    Transaction,
    ElementId,
    ElementTransformUtils,
    Options,
    GeometryInstance,
    Solid,
)

# UI element imports
import System.Windows
from System.Windows import WindowState, Visibility

# Path setup
lib_dir = os.path.dirname(os.path.dirname(__file__))
if lib_dir not in sys.path:
    sys.path.append(lib_dir)
if os.path.dirname(__file__) not in sys.path:
    sys.path.append(os.path.dirname(__file__))

# Import sub-dialog classes
import SmartAlignDialog
import AutoDimensionDialog
import SnapDimensionDialog
import AlignPositionsDialog

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ManaAlign.xaml')


def get_element_edges(element, view):
    """Get all edges from an element's geometry for linework override resets."""
    edges = []
    try:
        options = Options()
        options.View = view
        options.ComputeReferences = True
        geom = element.get_Geometry(options)

        if geom:
            for geom_obj in geom:
                if isinstance(geom_obj, GeometryInstance):
                    inst_geom = geom_obj.GetInstanceGeometry()
                    if inst_geom:
                        for inst_obj in inst_geom:
                            if isinstance(inst_obj, Solid):
                                for edge in inst_obj.Edges:
                                    edges.append(edge)
                elif isinstance(geom_obj, Solid):
                    for edge in geom_obj.Edges:
                        edges.append(edge)
    except Exception:
        pass
    return edges


class ManaAlignWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)
        self.uidoc = revit.uidoc
        self.doc = revit.doc

        # Initialize and nest sub-panels
        self._init_sub_panels()

        # Connect navigation
        self.btn_tab_smart_align.Checked += self._on_tab_changed
        self.btn_tab_auto_dimension.Checked += self._on_tab_changed
        self.btn_tab_snap_dimension.Checked += self._on_tab_changed
        self.btn_tab_align_positions.Checked += self._on_tab_changed

        # Reset overrides action
        self.btn_reset_overrides.Click += self._on_reset_overrides

        # Chrome actions
        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close_chrome.Click += self._close_chrome

    def _init_sub_panels(self):
        """Loads sub-tool grids and collapses their headers to avoid duplicates."""
        # 1. Smart Align
        try:
            self._smart_align_win = SmartAlignDialog.SmartAlignWindow()
            smart_align_grid = self._smart_align_win.Content
            self._smart_align_win.Content = None
            self.grid_smart_align.Children.Add(smart_align_grid)
            
            # Collapse sub-tool header and footer to unify status bar
            smart_align_grid.RowDefinitions[0].Height = System.Windows.GridLength(0)
            smart_align_grid.RowDefinitions[2].Height = System.Windows.GridLength(0)
            
            # Re-wire close
            self._smart_align_win.Close = self.Close
        except Exception as ex:
            print("Error loading Smart Align panel: {}".format(ex))

        # 2. Auto Dimension
        try:
            self._auto_dim_win = AutoDimensionDialog.AutoDimensionWindow(self.uidoc, self.doc)
            auto_dim_grid = self._auto_dim_win.Content
            self._auto_dim_win.Content = None
            self.grid_auto_dimension.Children.Add(auto_dim_grid)
            
            # Collapse sub-tool header and footer status
            auto_dim_grid.RowDefinitions[0].Height = System.Windows.GridLength(0)
            auto_dim_grid.RowDefinitions[3].Height = System.Windows.GridLength(0)
            
            # Re-wire close
            self._auto_dim_win.Close = self.Close
        except Exception as ex:
            print("Error loading Auto Dimension panel: {}".format(ex))

        # 3. Snap Dimension
        try:
            self._snap_dim_win = SnapDimensionDialog.MainWin()
            snap_dim_border = self._snap_dim_win.w.Content
            self._snap_dim_win.w.Content = None
            self.grid_snap_dimension.Children.Add(snap_dim_border)
            
            # Border child is the layout Grid
            snap_dim_grid = snap_dim_border.Child
            snap_dim_grid.RowDefinitions[0].Height = System.Windows.GridLength(0)
            
            # Re-wire close
            self._snap_dim_win.w.Close = self.Close
        except Exception as ex:
            print("Error loading Snap Dimension panel: {}".format(ex))

        # 4. Align Positions
        try:
            self._align_positions_win = AlignPositionsDialog.AlignPositionsWindow(self.uidoc, self.doc, self)
            align_positions_border = self._align_positions_win.Content
            self._align_positions_win.Content = None
            self.grid_align_positions.Children.Add(align_positions_border)
            
            # Border child is the layout Grid
            align_positions_grid = align_positions_border.Child
            align_positions_grid.RowDefinitions[0].Height = System.Windows.GridLength(0)
            align_positions_grid.RowDefinitions[4].Height = System.Windows.GridLength(0)
            
            # Re-wire close
            self._align_positions_win.Close = self.Close
        except Exception as ex:
            print("Error loading Align Positions panel: {}".format(ex))

    def _on_tab_changed(self, sender, e):
        """Switch active TabControl index and update status bar text."""
        if sender == self.btn_tab_smart_align:
            self.main_tab_control.SelectedIndex = 0
            self.status_text.Text = "Smart Align — Graphical alignment and distribution"
        elif sender == self.btn_tab_auto_dimension:
            self.main_tab_control.SelectedIndex = 1
            self.status_text.Text = "Auto Dimension — Place automatic dimension chains"
        elif sender == self.btn_tab_snap_dimension:
            self.main_tab_control.SelectedIndex = 2
            self.status_text.Text = "Snap Dimension — Align dimension lines to grid references"
        elif sender == self.btn_tab_align_positions:
            self.main_tab_control.SelectedIndex = 3
            self.status_text.Text = "Align Positions — Snap element offsets to clean grid multiples"

    def _on_reset_overrides(self, sender, e):
        """Executes graphic override resets on selected elements in the active view."""
        view = self.uidoc.ActiveGraphicsView
        if not view:
            forms.alert("No active graphics view found.")
            return

        # Fetch selection or entire view if nothing selected
        sel_ids = self.uidoc.Selection.GetElementIds()
        if not sel_ids or sel_ids.Count == 0:
            # Confirm resetting whole view
            if not forms.alert("No elements selected. Reset graphic overrides for ALL elements in the active view?", yes=True, no=True):
                return
            collector = FilteredElementCollector(self.doc, view.Id).WhereElementIsNotElementType().ToElementIds()
        else:
            collector = sel_ids

        override = OverrideGraphicSettings()
        reset_count = 0

        with Transaction(self.doc, "Reset Overrides") as t:
            t.Start()
            for el_id in collector:
                try:
                    view.SetElementOverrides(el_id, override)
                    element = self.doc.GetElement(el_id)
                    if element:
                        edges = get_element_edges(element, view)
                        for edge in edges:
                            try:
                                if edge.Reference:
                                    view.RemoveLinePatternOverride(edge.Reference)
                                    view.SetLineworkGraphicsStyle(edge.Reference, ElementId.InvalidElementId)
                            except Exception:
                                pass
                    reset_count += 1
                except Exception:
                    pass
            t.Commit()

        self.status_text.Text = "Successfully reset overrides on {} elements.".format(reset_count)
        script.get_output().print_md("### Reset Overrides: Complete\nReset overrides on **{}** elements in view **{}**.".format(reset_count, view.Name))

    def _minimize(self, sender, e):
        self.WindowState = WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()


def show_dialog():
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)
    window = ManaAlignWindow()
    # Align & Dimension tools are modeless to allow viewport interaction
    window.show(modal=False)
