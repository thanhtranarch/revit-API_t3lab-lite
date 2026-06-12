# -*- coding: utf-8 -*-
"""
Grid Swap v1.0 - DQT
Swap gridlines between 3D and 2D extents in Revit views.
Supports batch operations across multiple views.

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Grid\nSwap"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Swap gridlines between 3D and 2D extents. Batch swap across views."

# =====================================================
# IMPORTS
# =====================================================
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, View, ViewType, \
    DatumEnds, DatumExtentType, ElementId
from Autodesk.Revit.UI import *
from pyrevit import revit, HOST_APP

# Alias for Revit Grid to avoid conflict with WPF Grid
RevitGrid = DB.Grid

import sys
import os
import codecs
import System
from System.Windows import Window, Thickness, HorizontalAlignment, VerticalAlignment, TextWrapping, Visibility
from System.Windows import RoutedEventArgs, MessageBox as WPFMessageBox, MessageBoxButton, MessageBoxResult, MessageBoxImage
from System.Windows.Controls import *
from System.Windows.Media import SolidColorBrush, BrushConverter
from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Xml import XmlReader

# =====================================================
# REVIT CONTEXT
# =====================================================
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# =====================================================
# XAML UI DEFINITION
# =====================================================


# =====================================================
# GRID ITEM DATA CLASS
# =====================================================
class GridItem(System.Object):
    """Data class representing a grid's state in a specific view"""
    
    def __init__(self, grid, view):
        self._grid = grid
        self._view = view
        self._is_checked = False
        self._grid_name = grid.Name if grid.Name else "Unnamed"
        self._element_id = str(grid.Id.IntegerValue)
        
        # Determine 3D or 2D extent
        try:
            datum_extent_type = grid.GetDatumExtentTypeInView(DatumEnds.End0, view)
            self._is_3d = (datum_extent_type == DatumExtentType.Model)
        except:
            self._is_3d = True  # Default to 3D if cannot determine
        
        # Get bubble visibility
        try:
            self._bubble_start = "Visible" if grid.IsBubbleVisibleInView(DatumEnds.End0, view) else "Hidden"
        except:
            self._bubble_start = "N/A"
        
        try:
            self._bubble_end = "Visible" if grid.IsBubbleVisibleInView(DatumEnds.End1, view) else "Hidden"
        except:
            self._bubble_end = "N/A"
        
        # Get curve endpoints
        try:
            curve = grid.Curve
            if curve:
                sp = curve.GetEndPoint(0)
                ep = curve.GetEndPoint(1)
                self._start_point = "{:.1f}, {:.1f}".format(sp.X, sp.Y)
                self._end_point = "{:.1f}, {:.1f}".format(ep.X, ep.Y)
            else:
                self._start_point = "N/A"
                self._end_point = "N/A"
        except:
            self._start_point = "N/A"
            self._end_point = "N/A"
    
    # --- Properties for WPF Binding ---
    @property
    def IsChecked(self):
        return self._is_checked
    @IsChecked.setter
    def IsChecked(self, value):
        self._is_checked = value
    
    @property
    def GridName(self):
        return self._grid_name
    
    @property
    def ExtentStatus(self):
        return "3D" if self._is_3d else "2D"
    
    @property
    def BubbleStart(self):
        return self._bubble_start
    
    @property
    def BubbleEnd(self):
        return self._bubble_end
    
    @property
    def StartPoint(self):
        return self._start_point
    
    @property
    def EndPoint(self):
        return self._end_point
    
    @property
    def ElementId(self):
        return self._element_id
    
    @property
    def Is3D(self):
        return self._is_3d
    
    @property
    def Grid(self):
        return self._grid
    
    @property
    def View(self):
        return self._view


# =====================================================
# MAIN WINDOW CLASS
# =====================================================
class GridSwapWindow(object):
    """Main window for Grid Swap tool"""
    
    def __init__(self):
        # Parse XAML
        # Find T3Lab.extension parent folder dynamically
        current_dir = os.path.dirname(__file__)
        while current_dir and not current_dir.endswith('T3Lab.extension'):
            parent = os.path.dirname(current_dir)
            if parent == current_dir:
                break
            current_dir = parent
        xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "ConvertGridline.xaml")

        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml_content = f.read()
        xml_reader = XmlReader.Create(StringReader(xaml_content))
        self.window = XamlReader.Load(xml_reader)
        
        # Get UI elements
        self.txt_total_grids = self.window.FindName("txtTotalGrids")
        self.txt_is_3d = self.window.FindName("txtIs3D")
        self.txt_is_2d = self.window.FindName("txtIs2D")
        self.txt_selected_grids = self.window.FindName("txtSelectedGrids")
        self.txt_view_count = self.window.FindName("txtViewCount")
        
        self.cmb_view = self.window.FindName("cmbView")
        self.cmb_filter = self.window.FindName("cmbFilter")
        self.txt_search = self.window.FindName("txtSearch")
        
        self.dg_grids = self.window.FindName("dgGrids")
        
        self.btn_select_all = self.window.FindName("btnSelectAll")
        self.btn_select_none = self.window.FindName("btnSelectNone")
        self.btn_select_3d = self.window.FindName("btnSelect3D")
        self.btn_select_2d = self.window.FindName("btnSelect2D")
        self.btn_swap_to_2d = self.window.FindName("btnSwapTo2D")
        self.btn_swap_to_3d = self.window.FindName("btnSwapTo3D")
        self.btn_toggle = self.window.FindName("btnToggle")
        self.btn_refresh = self.window.FindName("btnRefresh")
        self.btn_close = self.window.FindName("btnClose")
        
        # Bubble control buttons
        self.btn_bubble_start_on = self.window.FindName("btnBubbleStartOn")
        self.btn_bubble_start_off = self.window.FindName("btnBubbleStartOff")
        self.btn_bubble_end_on = self.window.FindName("btnBubbleEndOn")
        self.btn_bubble_end_off = self.window.FindName("btnBubbleEndOff")
        self.btn_bubble_all_on = self.window.FindName("btnBubbleAllOn")
        self.btn_bubble_all_off = self.window.FindName("btnBubbleAllOff")
        
        # Data storage
        self.all_grid_items = []
        self.views_dict = {}  # name -> view
        
        # Wire events
        self._wire_events()
        
        # Initialize data
        self._load_views()
        self._load_grids()
    
    def _wire_events(self):
        """Connect UI events to handlers"""
        self.btn_select_all.Click += self._on_select_all
        self.btn_select_none.Click += self._on_select_none
        self.btn_select_3d.Click += self._on_select_3d
        self.btn_select_2d.Click += self._on_select_2d
        self.btn_swap_to_2d.Click += self._on_swap_to_2d
        self.btn_swap_to_3d.Click += self._on_swap_to_3d
        self.btn_toggle.Click += self._on_toggle
        self.btn_refresh.Click += self._on_refresh
        self.btn_close.Click += self._on_close
        self.cmb_view.SelectionChanged += self._on_view_changed
        self.cmb_filter.SelectionChanged += self._on_filter_changed
        self.txt_search.TextChanged += self._on_search_changed
        self.dg_grids.MouseDoubleClick += self._on_row_double_click
        
        # Bubble control events
        self.btn_bubble_start_on.Click += self._on_bubble_start_on
        self.btn_bubble_start_off.Click += self._on_bubble_start_off
        self.btn_bubble_end_on.Click += self._on_bubble_end_on
        self.btn_bubble_end_off.Click += self._on_bubble_end_off
        self.btn_bubble_all_on.Click += self._on_bubble_all_on
        self.btn_bubble_all_off.Click += self._on_bubble_all_off
    
    # =====================================================
    # DATA LOADING
    # =====================================================
    def _load_views(self):
        """Load all plan/elevation/section views into ComboBox"""
        self.views_dict.clear()
        self.cmb_view.Items.Clear()
        
        # Get active view first
        active_view = doc.ActiveView
        
        # Collect valid views (plans, elevations, sections, 3D)
        collector = FilteredElementCollector(doc).OfClass(View)
        valid_view_types = [
            ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation,
            ViewType.Section, ViewType.ThreeD, ViewType.AreaPlan,
            ViewType.EngineeringPlan, ViewType.Detail
        ]
        
        views = []
        for v in collector:
            if v.IsTemplate:
                continue
            if v.ViewType in valid_view_types:
                views.append(v)
        
        # Sort by name
        views.sort(key=lambda v: v.Name)
        
        # Add active view first
        active_name = ""
        if active_view and not active_view.IsTemplate:
            active_name = "[Active] " + active_view.Name
            self.cmb_view.Items.Add(active_name)
            self.views_dict[active_name] = active_view
        
        # Add all other views
        for v in views:
            if v.Id != active_view.Id:
                name = v.Name
                # Avoid duplicate names
                if name in self.views_dict:
                    name = "{} (ID:{})".format(name, v.Id.IntegerValue)
                self.cmb_view.Items.Add(name)
                self.views_dict[name] = v
        
        # Select active view
        if self.cmb_view.Items.Count > 0:
            self.cmb_view.SelectedIndex = 0
        
        # Update view count
        self.txt_view_count.Text = str(len(views))
    
    def _load_grids(self):
        """Load all grids for the selected view"""
        self.all_grid_items = []
        
        # Get selected view
        view = self._get_selected_view()
        if view is None:
            self._update_display()
            return
        
        # Collect all grids
        collector = FilteredElementCollector(doc).OfClass(RevitGrid)
        grids = list(collector)
        
        for grid in grids:
            try:
                item = GridItem(grid, view)
                self.all_grid_items.append(item)
            except Exception as e:
                print("Error loading grid {}: {}".format(grid.Name, str(e)))
        
        # Sort by name
        self.all_grid_items.sort(key=lambda x: x.GridName)
        
        self._update_display()
    
    def _get_selected_view(self):
        """Get the currently selected view from ComboBox"""
        if self.cmb_view.SelectedItem is None:
            return None
        selected_name = str(self.cmb_view.SelectedItem)
        return self.views_dict.get(selected_name, None)
    
    def _update_display(self):
        """Apply filters and update DataGrid and summary cards"""
        search_text = self.txt_search.Text.lower().strip() if self.txt_search.Text else ""
        filter_idx = self.cmb_filter.SelectedIndex if self.cmb_filter.SelectedIndex >= 0 else 0
        
        filtered = []
        for item in self.all_grid_items:
            # Search filter
            if search_text and search_text not in item.GridName.lower():
                continue
            # Type filter
            if filter_idx == 1 and not item.Is3D:  # 3D Only
                continue
            if filter_idx == 2 and item.Is3D:  # 2D Only
                continue
            filtered.append(item)
        
        # Update DataGrid
        self.dg_grids.ItemsSource = System.Array[System.Object](filtered)
        
        # Update summary cards
        total_3d = sum(1 for i in self.all_grid_items if i.Is3D)
        total_2d = sum(1 for i in self.all_grid_items if not i.Is3D)
        checked = sum(1 for i in self.all_grid_items if i.IsChecked)
        
        self.txt_total_grids.Text = str(len(self.all_grid_items))
        self.txt_is_3d.Text = str(total_3d)
        self.txt_is_2d.Text = str(total_2d)
        self.txt_selected_grids.Text = str(checked)
    
    # =====================================================
    # SELECTION HANDLERS
    # =====================================================
    def _on_select_all(self, sender, args):
        """Check all visible items"""
        for item in self.dg_grids.ItemsSource:
            item.IsChecked = True
        self._update_display()
    
    def _on_select_none(self, sender, args):
        """Uncheck all items"""
        for item in self.all_grid_items:
            item.IsChecked = False
        self._update_display()
    
    def _on_select_3d(self, sender, args):
        """Check only 3D grids"""
        for item in self.all_grid_items:
            item.IsChecked = item.Is3D
        self._update_display()
    
    def _on_select_2d(self, sender, args):
        """Check only 2D grids"""
        for item in self.all_grid_items:
            item.IsChecked = not item.Is3D
        self._update_display()
    
    # =====================================================
    # SWAP OPERATIONS
    # =====================================================
    def _get_checked_items(self):
        """Get all checked GridItems"""
        return [i for i in self.all_grid_items if i.IsChecked]
    
    def _on_swap_to_2d(self, sender, args):
        """Swap selected grids from 3D to 2D"""
        checked = self._get_checked_items()
        if not checked:
            WPFMessageBox.Show("No grids selected.\nPlease check the grids you want to swap.",
                              "Grid Swap", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Filter only 3D grids
        targets = [i for i in checked if i.Is3D]
        if not targets:
            WPFMessageBox.Show("No 3D grids found in selection.\nAll selected grids are already 2D.",
                              "Grid Swap", MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        view = self._get_selected_view()
        if view is None:
            return
        
        result = WPFMessageBox.Show(
            "Swap {} grid(s) from 3D to 2D in view '{}'?\n\nThis will change the extent type to view-specific (2D).".format(
                len(targets), view.Name),
            "Confirm Swap to 2D", MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        success = 0
        errors = []
        
        t = Transaction(doc, "DQT - Swap Grids to 2D")
        t.Start()
        try:
            for item in targets:
                try:
                    grid = item.Grid
                    # Set both ends to ViewSpecific (2D)
                    grid.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
                    grid.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
                    success += 1
                except Exception as e:
                    errors.append("{}: {}".format(item.GridName, str(e)))
            t.Commit()
        except Exception as e:
            t.RollBack()
            WPFMessageBox.Show("Transaction failed: {}".format(str(e)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        # Reload data
        self._load_grids()
        
        msg = "Successfully swapped {} grid(s) to 2D.".format(success)
        if errors:
            msg += "\n\nErrors ({}):\n{}".format(len(errors), "\n".join(errors[:10]))
        WPFMessageBox.Show(msg, "Grid Swap Complete", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_swap_to_3d(self, sender, args):
        """Swap selected grids from 2D to 3D"""
        checked = self._get_checked_items()
        if not checked:
            WPFMessageBox.Show("No grids selected.\nPlease check the grids you want to swap.",
                              "Grid Swap", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Filter only 2D grids
        targets = [i for i in checked if not i.Is3D]
        if not targets:
            WPFMessageBox.Show("No 2D grids found in selection.\nAll selected grids are already 3D.",
                              "Grid Swap", MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        view = self._get_selected_view()
        if view is None:
            return
        
        result = WPFMessageBox.Show(
            "Swap {} grid(s) from 2D to 3D in view '{}'?\n\nThis will change the extent type to model-wide (3D).\nNote: Grid positions will reset to model extents.".format(
                len(targets), view.Name),
            "Confirm Swap to 3D", MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        success = 0
        errors = []
        
        t = Transaction(doc, "DQT - Swap Grids to 3D")
        t.Start()
        try:
            for item in targets:
                try:
                    grid = item.Grid
                    # Set both ends to Model (3D)
                    grid.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.Model)
                    grid.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.Model)
                    success += 1
                except Exception as e:
                    errors.append("{}: {}".format(item.GridName, str(e)))
            t.Commit()
        except Exception as e:
            t.RollBack()
            WPFMessageBox.Show("Transaction failed: {}".format(str(e)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        # Reload data
        self._load_grids()
        
        msg = "Successfully swapped {} grid(s) to 3D.".format(success)
        if errors:
            msg += "\n\nErrors ({}):\n{}".format(len(errors), "\n".join(errors[:10]))
        WPFMessageBox.Show(msg, "Grid Swap Complete", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_toggle(self, sender, args):
        """Toggle selected grids: 3D becomes 2D, 2D becomes 3D"""
        checked = self._get_checked_items()
        if not checked:
            WPFMessageBox.Show("No grids selected.\nPlease check the grids you want to toggle.",
                              "Grid Swap", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        view = self._get_selected_view()
        if view is None:
            return
        
        count_3d_to_2d = sum(1 for i in checked if i.Is3D)
        count_2d_to_3d = sum(1 for i in checked if not i.Is3D)
        
        result = WPFMessageBox.Show(
            "Toggle {} grid(s) in view '{}'?\n\n  {} grid(s): 3D -> 2D\n  {} grid(s): 2D -> 3D".format(
                len(checked), view.Name, count_3d_to_2d, count_2d_to_3d),
            "Confirm Toggle", MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        success = 0
        errors = []
        
        t = Transaction(doc, "DQT - Toggle Grid Extents")
        t.Start()
        try:
            for item in checked:
                try:
                    grid = item.Grid
                    if item.Is3D:
                        # 3D -> 2D
                        new_type = DatumExtentType.ViewSpecific
                    else:
                        # 2D -> 3D
                        new_type = DatumExtentType.Model
                    
                    grid.SetDatumExtentType(DatumEnds.End0, view, new_type)
                    grid.SetDatumExtentType(DatumEnds.End1, view, new_type)
                    success += 1
                except Exception as e:
                    errors.append("{}: {}".format(item.GridName, str(e)))
            t.Commit()
        except Exception as e:
            t.RollBack()
            WPFMessageBox.Show("Transaction failed: {}".format(str(e)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        # Reload data
        self._load_grids()
        
        msg = "Successfully toggled {} grid(s).".format(success)
        if errors:
            msg += "\n\nErrors ({}):\n{}".format(len(errors), "\n".join(errors[:10]))
        WPFMessageBox.Show(msg, "Toggle Complete", MessageBoxButton.OK, MessageBoxImage.Information)
    
    # =====================================================
    # BUBBLE CONTROL OPERATIONS
    # =====================================================
    def _set_bubble_visibility(self, end, visible):
        """Set bubble visibility for checked grids at specified end.
        
        Args:
            end: DatumEnds.End0 (Start) or DatumEnds.End1 (End) or None (both)
            visible: True to show, False to hide
        """
        checked = self._get_checked_items()
        if not checked:
            WPFMessageBox.Show("No grids selected.\nPlease check the grids you want to modify.",
                              "Bubble Control", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        view = self._get_selected_view()
        if view is None:
            return
        
        # Build description
        if end == DatumEnds.End0:
            end_name = "Start"
        elif end == DatumEnds.End1:
            end_name = "End"
        else:
            end_name = "Start + End"
        
        action = "Show" if visible else "Hide"
        
        success = 0
        errors = []
        
        t = Transaction(doc, "DQT - {} {} Bubble".format(action, end_name))
        t.Start()
        try:
            for item in checked:
                try:
                    grid = item.Grid
                    ends_to_set = []
                    if end is None:
                        ends_to_set = [DatumEnds.End0, DatumEnds.End1]
                    else:
                        ends_to_set = [end]
                    
                    for e in ends_to_set:
                        if visible:
                            grid.ShowBubbleInView(e, view)
                        else:
                            grid.HideBubbleInView(e, view)
                    success += 1
                except Exception as ex:
                    errors.append("{}: {}".format(item.GridName, str(ex)))
            t.Commit()
        except Exception as ex:
            t.RollBack()
            WPFMessageBox.Show("Transaction failed: {}".format(str(ex)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        # Reload data
        self._load_grids()
        
        msg = "{} {} bubble for {} grid(s).".format(action, end_name, success)
        if errors:
            msg += "\n\nErrors ({}):\n{}".format(len(errors), "\n".join(errors[:10]))
        WPFMessageBox.Show(msg, "Bubble Control", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_bubble_start_on(self, sender, args):
        self._set_bubble_visibility(DatumEnds.End0, True)
    
    def _on_bubble_start_off(self, sender, args):
        self._set_bubble_visibility(DatumEnds.End0, False)
    
    def _on_bubble_end_on(self, sender, args):
        self._set_bubble_visibility(DatumEnds.End1, True)
    
    def _on_bubble_end_off(self, sender, args):
        self._set_bubble_visibility(DatumEnds.End1, False)
    
    def _on_bubble_all_on(self, sender, args):
        self._set_bubble_visibility(None, True)
    
    def _on_bubble_all_off(self, sender, args):
        self._set_bubble_visibility(None, False)
    
    # =====================================================
    # UI EVENT HANDLERS
    # =====================================================
    def _on_view_changed(self, sender, args):
        """Reload grids when view selection changes"""
        self._load_grids()
    
    def _on_filter_changed(self, sender, args):
        """Refilter display when filter changes"""
        self._update_display()
    
    def _on_search_changed(self, sender, args):
        """Refilter display when search text changes"""
        self._update_display()
    
    def _on_refresh(self, sender, args):
        """Reload all data"""
        self._load_views()
        self._load_grids()
    
    def _on_close(self, sender, args):
        """Close the window"""
        self.window.Close()
    
    def _on_row_double_click(self, sender, args):
        """Select grid in Revit on double-click"""
        if self.dg_grids.SelectedItem is not None:
            try:
                item = self.dg_grids.SelectedItem
                grid = item.Grid
                element_id = grid.Id
                uidoc.Selection.SetElementIds(System.Collections.Generic.List[ElementId]([element_id]))
                uidoc.ShowElements(element_id)
            except Exception as e:
                print("Error selecting grid: {}".format(str(e)))
    
    def show(self):
        """Show the window"""
        self.window.ShowDialog()


# =====================================================
# MAIN ENTRY POINT
# =====================================================
try:
    window = GridSwapWindow()
    window.show()
except Exception as e:
    from pyrevit import forms
    forms.alert("Error launching Grid Swap:\n{}".format(str(e)), title="Grid Swap Error")