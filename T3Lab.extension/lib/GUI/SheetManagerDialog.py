# -*- coding: utf-8 -*-
"""
Sheet Manager Dialog
Unified Sheet Manager including sheet lists and re-numbering inside a Lumina UI.
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
                             MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, WindowState)
from System.Windows.Controls import (RowDefinition, ColumnDefinition, Border,
                                      StackPanel, TextBlock, TextBox, Button,
                                      ComboBox, ComboBoxItem, DataGrid, Orientation,
                                      DataGridTextColumn, ScrollViewer)
from System.Windows.Media import SolidColorBrush
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

# CRITICAL: Import WPF Grid BEFORE Revit wildcard import
from System.Windows.Controls import Grid as WPFGrid

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet, Transaction

# Add Services/SheetManager directory to sys.path so we can import modules
GUI_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(GUI_DIR))
SERVICES_DIR = os.path.join(EXT_DIR, 'lib', 'Services', 'SheetManager')
if SERVICES_DIR not in sys.path:
    sys.path.append(SERVICES_DIR)

# Import Sheet Manager services
try:
    from sheet_core.revit_service import RevitService
    from sheet_core.data_models import ChangeTracker, SheetModel
    from excel_service import ExcelService
    from viewsheet_sets_service import ViewSheetSetsService
    from place_views_service import PlaceViewsService
    from custom_parameters_service import CustomParametersService
except Exception as e:
    # Print error in case imports fail
    print("Error importing services: {}".format(e))

doc = revit.doc
XAML_FILE = os.path.join(GUI_DIR, 'Tools', 'SheetManager.xaml')


# =====================================================
# RENUMBER WRAPPER MODEL
# =====================================================

class RenumberItem(INotifyPropertyChanged):
    """Wrapper class for Sheet Renumber preview grid"""
    def __init__(self, sheet_model):
        self._property_changed_handlers = []
        self.sheet_model = sheet_model
        self.orig_number = sheet_model.sheet_number
        self.name = sheet_model.sheet_name
        self._preview_number = sheet_model.sheet_number

    @property
    def preview_number(self):
        return self._preview_number
    @preview_number.setter
    def preview_number(self, value):
        if self._preview_number != value:
            self._preview_number = value
            self.OnPropertyChanged("preview_number")

    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        self._property_changed_handlers.remove(handler)

    def OnPropertyChanged(self, property_name):
        args = PropertyChangedEventArgs(property_name)
        for handler in self._property_changed_handlers:
            handler(self, args)


# =====================================================
# MAIN WINDOW CONTROLLER
# =====================================================

class SheetManagerWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self.doc = revit.doc
        
        # Initialize Core Services
        self.revit_service = RevitService(self.doc)
        self.change_tracker = ChangeTracker()
        
        try:
            self.excel_service = ExcelService()
        except:
            self.excel_service = None
            
        try:
            self.sheet_sets_service = ViewSheetSetsService(self.doc)
        except:
            self.sheet_sets_service = None
            
        try:
            self.place_views_service = PlaceViewsService(self.doc)
        except:
            self.place_views_service = None
            
        try:
            self.params_service = CustomParametersService(self.doc, __revit__.Application)
        except:
            self.params_service = None

        # Data collection
        self.all_sheets = []
        self.filtered_sheets = ObservableCollection[object]()
        self.renumber_items = ObservableCollection[object]()
        
        # Chrome controls
        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close.Click += self._close_chrome
        
        # Radio Tab navigation
        self.nav_sheets.Checked += self._on_tab_changed
        self.nav_renumber.Checked += self._on_tab_changed
        
        # Bind events: SHEETS Tab
        self.sheets_search_box.TextChanged += self._on_sheets_search_changed
        self.sheets_filter_combo.SelectionChanged += self._on_sheets_filter_changed
        self.sheets_select_all_btn.Click += self._on_sheets_select_all
        self.sheets_clear_btn.Click += self._on_sheets_clear_all
        
        self.sheets_sets_btn.Click += self._on_sheets_sets
        self.sheets_place_views_btn.Click += self._on_sheets_place_views
        self.sheets_custom_params_btn.Click += self._on_sheets_custom_params
        
        self.sheets_excel_btn.Click += self._on_sheets_excel
        self.sheets_refresh_btn.Click += self._on_sheets_refresh
        self.sheets_apply_btn.Click += self._on_sheets_apply
        self.sheets_close_btn.Click += self._on_close
        
        self.sheets_grid.SelectionChanged += self._on_sheets_selection_changed
        self.sheets_grid.CellEditEnding += self._on_sheets_cell_edit
        self.sheets_grid.ItemsSource = self.filtered_sheets
        
        # Bind events: RENUMBER Tab
        self.renum_refresh_btn.Click += self._on_renum_refresh
        self.renum_preview_btn.Click += self._on_renum_preview
        self.renum_run_btn.Click += self._on_renum_run
        self.renum_close_btn.Click += self._on_close
        self.renum_grid.ItemsSource = self.renumber_items

        # Load initial data
        self._load_sheets_data()
        self._apply_sheets_filters()
        self._load_renumber_preview_data()

    # ── Chrome Event Handlers ────────────────────────────────────
    def _minimize(self, sender, e):
        self.WindowState = WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def _close_chrome(self, sender, e):
        self.Close()
        
    def _on_close(self, sender, args):
        self.Close()

    def _on_tab_changed(self, sender, e):
        if not hasattr(self, 'tab_control'):
            return
        if self.nav_sheets.IsChecked:
            self.tab_control.SelectedIndex = 0
        elif self.nav_renumber.IsChecked:
            self.tab_control.SelectedIndex = 1
            # Refresh renumber tab data whenever switching to it
            self._load_renumber_preview_data()

    # ── SHEETS Tab Logics ─────────────────────────────────────────
    def _load_sheets_data(self):
        self.all_sheets = self.revit_service.get_all_sheets()
        self.change_tracker.clear_all()
        self._update_sheets_summary()

    def _apply_sheets_filters(self):
        self.filtered_sheets.Clear()
        
        search_text = self.sheets_search_box.Text.lower() if self.sheets_search_box.Text else ""
        filter_index = self.sheets_filter_combo.SelectedIndex
        
        for item in self.all_sheets:
            # Search
            if search_text:
                if search_text not in item.sheet_number.lower() and \
                   search_text not in item.sheet_name.lower():
                    continue
            # Placeholder Filter
            if filter_index == 1: # Placeholder Only
                # Check direct property on element
                if not item.element.IsPlaceholder:
                    continue
            elif filter_index == 2: # Non-Placeholder Only
                if item.element.IsPlaceholder:
                    continue
                    
            self.filtered_sheets.Add(item)
            
        self._update_sheets_summary()

    def _update_sheets_summary(self):
        self.sheets_total_text.Text = str(len(self.all_sheets))
        self.sheets_selected_text.Text = str(self.sheets_grid.SelectedItems.Count)
        
        categories = set()
        for s in self.all_sheets:
            # Categorization based designed_by or drawn_by if any
            if s.designed_by and s.designed_by != "-":
                categories.add(s.designed_by)
        self.sheets_categories_text.Text = str(len(categories)) if categories else "1"
        
        self.sheets_changes_text.Text = str(len(self.change_tracker.modified_items))

    def _on_sheets_search_changed(self, sender, args):
        self._apply_sheets_filters()

    def _on_sheets_filter_changed(self, sender, args):
        self._apply_sheets_filters()

    def _on_sheets_select_all(self, sender, args):
        self.sheets_grid.SelectAll()

    def _on_sheets_clear_all(self, sender, args):
        self.sheets_grid.UnselectAll()

    def _on_sheets_selection_changed(self, sender, args):
        self._update_sheets_summary()

    def _on_sheets_cell_edit(self, sender, args):
        from System.Windows.Controls import DataGridEditAction
        if args.EditAction == DataGridEditAction.Cancel:
            return
            
        try:
            item = args.Row.Item
            column = args.Column
            
            # Edit sheet number
            if column.Header == "Sheet Number":
                new_val = args.EditingElement.Text
                if item.sheet_number != new_val:
                    item.sheet_number = new_val
                    item.check_if_modified()
                    self.change_tracker.track_modification(item)
                    
            # Edit sheet name
            elif column.Header == "Sheet Name":
                new_val = args.EditingElement.Text
                if item.sheet_name != new_val:
                    item.sheet_name = new_val
                    item.check_if_modified()
                    self.change_tracker.track_modification(item)
                    
            # Edit designed_by
            elif column.Header == "Designed By":
                new_val = args.EditingElement.Text
                item.designed_by = new_val
                item.is_modified = True
                self.change_tracker.track_modification(item)
                
            # Edit checked_by
            elif column.Header == "Checked By":
                new_val = args.EditingElement.Text
                item.checked_by = new_val
                item.is_modified = True
                self.change_tracker.track_modification(item)
                
            # Edit approved_by
            elif column.Header == "Approved By":
                new_val = args.EditingElement.Text
                item.approved_by = new_val
                item.is_modified = True
                self.change_tracker.track_modification(item)
                
            # Edit drawn_by
            elif column.Header == "Drawn By":
                new_val = args.EditingElement.Text
                item.drawn_by = new_val
                item.is_modified = True
                self.change_tracker.track_modification(item)
                
            self._update_sheets_summary()
        except Exception as e:
            MessageBox.Show("Error editing sheet parameter: {}".format(str(e)), "Error")

    def _on_sheets_sets(self, sender, args):
        if self.sheet_sets_service:
            try:
                from viewsheet_sets_dialog import ViewSheetSetsDialog
                dialog = ViewSheetSetsDialog(self.doc, self.sheet_sets_service)
                dialog.ShowDialog()
                self._load_sheets_data()
                self._apply_sheets_filters()
            except Exception as e:
                MessageBox.Show("Error showing ViewSheet Sets dialog:\n{}".format(str(e)), "Error")

    def _on_sheets_place_views(self, sender, args):
        if self.place_views_service:
            try:
                from place_views_dialog import PlaceViewsDialog
                dialog = PlaceViewsDialog(self.doc, self.place_views_service)
                dialog.ShowDialog()
                self._load_sheets_data()
                self._apply_sheets_filters()
            except Exception as e:
                MessageBox.Show("Error showing Place Views dialog:\n{}".format(str(e)), "Error")

    def _on_sheets_custom_params(self, sender, args):
        if self.params_service:
            try:
                from custom_parameters_dialog import CustomParametersDialog
                dialog = CustomParametersDialog(self.doc, self.params_service)
                dialog.ShowDialog()
                self._load_sheets_data()
                self._apply_sheets_filters()
            except Exception as e:
                MessageBox.Show("Error showing Custom Parameters dialog:\n{}".format(str(e)), "Error")

    def _on_sheets_excel(self, sender, args):
        if not self.excel_service:
            MessageBox.Show("Excel Service is not initialized.", "Error")
            return
            
        from System.Windows.Forms import SaveFileDialog, OpenFileDialog, DialogResult
        result = MessageBox.Show(
            "Do you want to EXPORT the current list to Excel?\n(Select No if you want to IMPORT from Excel)",
            "Excel Import/Export", MessageBoxButton.YesNoCancel, MessageBoxImage.Question
        )
        
        if result == MessageBoxResult.Yes:
            # EXPORT
            sfd = SaveFileDialog()
            sfd.Filter = "Excel Files (*.xlsx)|*.xlsx"
            sfd.FileName = "T3Lab_SheetManager_Export.xlsx"
            if sfd.ShowDialog() == DialogResult.OK:
                try:
                    sheets_list = list(self.filtered_sheets)
                    # Convert model properties to dictionary format for excel_service
                    data_to_export = []
                    for s in sheets_list:
                        data_to_export.append({
                            "sheet_number": s.sheet_number,
                            "sheet_name": s.sheet_name,
                            "designed_by": s.designed_by,
                            "checked_by": s.checked_by,
                            "drawn_by": s.drawn_by,
                            "approved_by": s.approved_by,
                            "id": s.id.IntegerValue
                        })
                    self.excel_service.export_sheets(sfd.FileName, data_to_export)
                    MessageBox.Show("Successfully exported sheets to Excel.", "Export Successful")
                except Exception as ex:
                    MessageBox.Show("Error exporting Excel: {}".format(str(ex)), "Error")
                    
        elif result == MessageBoxResult.No:
            # IMPORT
            ofd = OpenFileDialog()
            ofd.Filter = "Excel Files (*.xlsx)|*.xlsx"
            if ofd.ShowDialog() == DialogResult.OK:
                try:
                    imported_data = self.excel_service.import_sheets(ofd.FileName)
                    if not imported_data:
                        MessageBox.Show("No valid sheet updates found in Excel file.", "Excel Import")
                        return
                        
                    t = Transaction(self.doc, "Excel Sync Sheet Parameters")
                    t.Start()
                    
                    sheets_dict = {s.id.IntegerValue: s for s in self.all_sheets}
                    success = 0
                    failed = 0
                    
                    for row in imported_data:
                        sh_id = row.get("id")
                        if sh_id in sheets_dict:
                            item = sheets_dict[sh_id]
                            try:
                                if "sheet_number" in row:
                                    item.sheet_number = str(row["sheet_number"])
                                if "sheet_name" in row:
                                    item.sheet_name = str(row["sheet_name"])
                                
                                # Assign standard parameters
                                for param_name in ["designed_by", "checked_by", "drawn_by", "approved_by"]:
                                    if param_name in row:
                                        setattr(item, param_name, str(row[param_name]))
                                        
                                # Lookup and update Revit parameters
                                element = item.element
                                for map_name, rev_name in [("designed_by", "Designed By"), 
                                                           ("checked_by", "Checked By"), 
                                                           ("drawn_by", "Drawn By"), 
                                                           ("approved_by", "Approved By")]:
                                    val = row.get(map_name)
                                    if val:
                                        param = element.LookupParameter(rev_name)
                                        if param and not param.IsReadOnly:
                                            param.Set(str(val))
                                            
                                # Apply sheet number & name update in transaction
                                self.revit_service.update_sheet(item)
                                item.commit_changes()
                                success += 1
                            except:
                                failed += 1
                                
                    t.Commit()
                    MessageBox.Show("Excel Sync Completed.\nUpdated: {}\nFailed: {}".format(success, failed), "Excel Import")
                    self._on_sheets_refresh(None, None)
                except Exception as ex:
                    MessageBox.Show("Error importing Excel: {}".format(str(ex)), "Error")

    def _on_sheets_refresh(self, sender, args):
        self._load_sheets_data()
        self._apply_sheets_filters()

    def _on_sheets_apply(self, sender, args):
        if not self.change_tracker.has_changes():
            MessageBox.Show("No pending changes to apply.", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
            return
            
        modified = len(self.change_tracker.modified_items)
        msg = "Apply changes?\n\nModified Sheets: {}".format(modified)
        
        result = MessageBox.Show(msg, "Confirm Changes", MessageBoxButton.YesNo, MessageBoxImage.Question)
        if result == MessageBoxResult.Yes:
            t = Transaction(self.doc, "Apply Sheet Manager Changes")
            t.Start()
            try:
                success = 0
                failed = 0
                for item in self.change_tracker.modified_items:
                    try:
                        # Update standard params in Revit
                        element = item.element
                        for val, p_name in [(item.designed_by, "Designed By"),
                                            (item.checked_by, "Checked By"),
                                            (item.drawn_by, "Drawn By"),
                                            (item.approved_by, "Approved By")]:
                            if val and val != "-":
                                p = element.LookupParameter(p_name)
                                if p and not p.IsReadOnly:
                                    p.Set(val)
                                    
                        # Update Number & Name
                        if self.revit_service.update_sheet(item):
                            item.commit_changes()
                            success += 1
                        else:
                            failed += 1
                    except:
                        failed += 1
                        
                t.Commit()
                msg = "Successfully updated: {}".format(success)
                if failed > 0:
                    msg += "\nFailed: {}".format(failed)
                MessageBox.Show(msg, "Apply Complete", MessageBoxButton.OK, MessageBoxImage.Information)
                self._load_sheets_data()
                self._apply_sheets_filters()
            except Exception as e:
                t.RollBack()
                MessageBox.Show("Error applying changes: {}".format(str(e)), "Error")

    # ── RENUMBER Tab Logics ───────────────────────────────────────
    def _load_renumber_preview_data(self):
        self.renumber_items.Clear()
        for s in self.filtered_sheets:
            self.renumber_items.Add(RenumberItem(s))

    def _on_renum_refresh(self, sender, args):
        self._load_sheets_data()
        self._apply_sheets_filters()
        self._load_renumber_preview_data()

    def _on_renum_preview(self, sender, args):
        selected_preview = [item for item in self.renum_grid.SelectedItems]
        if not selected_preview:
            MessageBox.Show("Please select sheets in the preview grid first.", "Info")
            return
            
        prefix = self.renum_prefix_box.Text or ""
        suffix = self.renum_suffix_box.Text or ""
        
        try:
            start_num = int(self.renum_start_box.Text)
            step_num = int(self.renum_step_box.Text)
        except ValueError:
            MessageBox.Show("Starting Number and Increment Step must be integers.", "Error")
            return

        for index, item in enumerate(selected_preview):
            new_num = "{}{}{}".format(prefix, start_num + index * step_num, suffix)
            item.preview_number = new_num

    def _on_renum_run(self, sender, args):
        selected_preview = [item for item in self.renum_grid.SelectedItems]
        if not selected_preview:
            MessageBox.Show("Please select sheets in the preview grid to renumber.", "Info")
            return
            
        # First preview them in case user hasn't previewed
        self._on_renum_preview(None, None)
        
        result = MessageBox.Show(
            "Renumber {} selected sheet(s)?".format(len(selected_preview)),
            "Confirm Renumber", MessageBoxButton.YesNo, MessageBoxImage.Question
        )
        
        if result == MessageBoxResult.Yes:
            t = Transaction(self.doc, "Batch Renumber Sheets")
            t.Start()
            try:
                success = 0
                failed = 0
                for item in selected_preview:
                    sheet = item.sheet_model.element
                    new_num = item.preview_number
                    try:
                        sheet.SheetNumber = new_num
                        item.sheet_model.sheet_number = new_num
                        item.sheet_model.commit_changes()
                        success += 1
                    except Exception as ex:
                        print("Renumber failed for sheet '{}': {}".format(item.name, ex))
                        failed += 1
                t.Commit()
                
                msg = "Renumbered successfully: {}".format(success)
                if failed > 0:
                    msg += "\nFailed: {}".format(failed)
                MessageBox.Show(msg, "Renumber Complete", MessageBoxButton.OK, MessageBoxImage.Information)
                
                self._on_renum_refresh(None, None)
            except Exception as e:
                t.RollBack()
                MessageBox.Show("Error during renumbering: {}".format(str(e)), "Error")


# =====================================================
# LAUNCHER FUNCTION
# =====================================================

def show_sheet_manager():
    """Launch the unified Sheet Manager Dialog"""
    try:
        window = SheetManagerWindow()
        window.ShowDialog()
    except Exception as e:
        print("\nFATAL ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        MessageBox.Show(
            "Error starting Sheet Manager:\n\n{}".format(str(e)),
            "Error", MessageBoxButton.OK, MessageBoxImage.Error
        )
