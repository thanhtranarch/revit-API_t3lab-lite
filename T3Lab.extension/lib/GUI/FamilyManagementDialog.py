# -*- coding: utf-8 -*-
"""
Family Management Dialog

WPF Window class handling batch renaming, prefixes/suffixes, case modification,
and worksets assignment for Revit Families.

Author: Tran Tien Thanh
"""

import os
import sys
import traceback
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import WindowState, MessageBox
from System.Windows.Controls import ComboBoxItem
from pyrevit import revit, forms, script

from Autodesk.Revit.DB import (
    FilteredElementCollector, FilteredWorksetCollector, WorksetKind,
    Family, FamilySymbol, ElementType, GroupType, AssemblyType, Group, AssemblyInstance,
    Transaction, BuiltInParameter, ElementId, Category
)

# DEFINE VARIABLES
# ==============================================================================
doc   = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

# HELPER SANITIZER
# ==============================================================================
def sanitize_name(name):
    """Remove Revit invalid characters from names to prevent API crashes."""
    invalid_chars = ['\\', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~']
    for char in invalid_chars:
        name = name.replace(char, '')
    return name.strip()

# DATA GRID VIEW-MODEL
# ==============================================================================
class FamilyRow(object):
    """View-model for a family or type row in the DataGrid."""
    def __init__(self, element, family_name, type_name, category_name, workset_id, is_loadable=True):
        self.IsSelected = False
        self.FamilyName = family_name
        self.TypeName = type_name
        self.CategoryName = category_name or "Unknown"
        self.WorksetId = workset_id
        
        # Reference to the actual Revit Element (Family, FamilySymbol, GroupType, AssemblyType)
        self.element = element
        self.is_loadable = is_loadable
        
        # Original state to detect edits and support Undo
        self.OriginalFamilyName = family_name
        self.OriginalTypeName = type_name
        self.OriginalWorksetId = workset_id

    @property
    def IsModified(self):
        return (self.FamilyName != self.OriginalFamilyName or 
                self.TypeName != self.OriginalTypeName or 
                self.WorksetId != self.OriginalWorksetId)

class WorksetItem(object):
    """Workset class wrapping for ComboBox binding."""
    def __init__(self, name, ws_id):
        self.Name = name
        self.Id = ws_id

# MAIN DIALOG WINDOW
# ==============================================================================
class FamilyManagementWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)
        self.DataContext = self
        self._all_rows = []
        self._visible_rows = []
        self._worksets = []
        
        self._load_worksets()
        self._refresh_data()
        self._update_counts()

    # --- CHROMATIC CHROMES ---
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

    # --- INITIALIZATIONS ---
    def _load_worksets(self):
        """Load project worksets for ComboBox assignment."""
        self._worksets = [WorksetItem("<No Workset / None>", -1)]
        if doc.IsWorkshared:
            f_collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
            for ws in f_collector.ToWorksets():
                self._worksets.append(WorksetItem(ws.Name, ws.Id.IntegerValue))
        self.Worksets = self._worksets # Bindable property

    def _refresh_data(self):
        """Gather and filter elements from Revit based on Scope and Category."""
        scope_all = self.rb_scope_all.IsChecked
        scope_view = self.rb_scope_view.IsChecked
        scope_selection = self.rb_scope_selection.IsChecked

        category_idx = self.cb_category.SelectedIndex
        # 0: Loadable Families, 1: System Families, 2: Model Groups, 3: Assemblies

        # Step 1: Filter collector by scope
        if scope_selection:
            selected_ids = uidoc.Selection.GetElementIds()
            if not selected_ids:
                self._all_rows = []
                self._visible_rows = []
                self.dg_families.ItemsSource = self._visible_rows
                self.status_text.Text = "No elements selected in Revit."
                return
            collector = FilteredElementCollector(doc, selected_ids)
        elif scope_view:
            collector = FilteredElementCollector(doc, doc.ActiveView.Id)
        else:
            collector = FilteredElementCollector(doc)

        rows = []
        
        # Loadable Families
        if category_idx == 0:
            if scope_selection or scope_view:
                # Find all FamilyInstances in the selection/view
                instances = collector.WhereElementIsNotElementType().ToElements()
                symbols_found = set()
                for inst in instances:
                    try:
                        symbol = inst.Symbol
                        if symbol and symbol.Id not in symbols_found:
                            symbols_found.add(symbol.Id)
                            family = symbol.Family
                            ws_id = inst.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger() if doc.IsWorkshared else -1
                            rows.append(FamilyRow(
                                element=symbol,
                                family_name=family.Name,
                                type_name=symbol.Name,
                                category_name=symbol.Category.Name if symbol.Category else "Generic Models",
                                workset_id=ws_id,
                                is_loadable=True
                            ))
                    except Exception:
                        pass
            else:
                # All loaded families in document
                families = FilteredElementCollector(doc).OfClass(Family).ToElements()
                for fam in families:
                    try:
                        if not fam.IsEditable:
                            continue # Skip non-editable system families listed here
                        for symbol_id in fam.GetFamilySymbolIds():
                            symbol = doc.GetElement(symbol_id)
                            if symbol:
                                ws_id = -1
                                # Try to get workset of first placed instance or the symbol itself
                                try:
                                    ws_id = symbol.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
                                except Exception:
                                    pass
                                rows.append(FamilyRow(
                                    element=symbol,
                                    family_name=fam.Name,
                                    type_name=symbol.Name,
                                    category_name=symbol.Category.Name if symbol.Category else "Generic Models",
                                    workset_id=ws_id,
                                    is_loadable=True
                                ))
                    except Exception:
                        pass

        # System Families
        elif category_idx == 1:
            # Collect element types that are subclasses of ElementType and represent system families
            types_collector = FilteredElementCollector(doc).OfClass(ElementType).ToElements()
            for t in types_collector:
                try:
                    # System families have a category but are not Loadable (i.e. not Family elements)
                    if hasattr(t, "FamilyName") and t.FamilyName and t.Category:
                        # Exclude symbols of loadable families by checking if family is editable (already handled in category 0)
                        # A quick check is to see if we can cast to FamilySymbol and look at its Family.IsEditable
                        is_loadable = False
                        if isinstance(t, FamilySymbol):
                            try:
                                if t.Family.IsEditable:
                                    is_loadable = True
                            except Exception:
                                pass
                        
                        if not is_loadable:
                            ws_id = -1
                            if doc.IsWorkshared:
                                try:
                                    ws_id = t.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
                                except Exception:
                                    pass
                            rows.append(FamilyRow(
                                element=t,
                                family_name=t.FamilyName,
                                type_name=t.Name,
                                category_name=t.Category.Name,
                                workset_id=ws_id,
                                is_loadable=False
                            ))
                except Exception:
                    pass

        # Model Groups
        elif category_idx == 2:
            group_types = FilteredElementCollector(doc).OfClass(GroupType).ToElements()
            for gt in group_types:
                try:
                    ws_id = -1
                    if doc.IsWorkshared:
                        try:
                            ws_id = gt.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
                        except Exception:
                            pass
                    rows.append(FamilyRow(
                        element=gt,
                        family_name="Model Group",
                        type_name=gt.Name,
                        category_name="Groups",
                        workset_id=ws_id,
                        is_loadable=False
                    ))
                except Exception:
                    pass

        # Assemblies
        elif category_idx == 3:
            assembly_types = FilteredElementCollector(doc).OfClass(AssemblyType).ToElements()
            for at in assembly_types:
                try:
                    ws_id = -1
                    if doc.IsWorkshared:
                        try:
                            ws_id = at.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsInteger()
                        except Exception:
                            pass
                    rows.append(FamilyRow(
                        element=at,
                        family_name="Assembly",
                        type_name=at.Name,
                        category_name="Assemblies",
                        workset_id=ws_id,
                        is_loadable=False
                    ))
                except Exception:
                    pass

        self._all_rows = rows
        self._apply_filter()
        self.status_text.Text = "Loaded {} rows.".format(len(self._all_rows))

    def _apply_filter(self):
        """Filter the loaded rows by search term."""
        search = (self.tb_search.Text or "").strip().lower()
        if not search:
            self._visible_rows = list(self._all_rows)
        else:
            self._visible_rows = []
            for r in self._all_rows:
                if (search in r.FamilyName.lower() or 
                    search in r.TypeName.lower() or 
                    search in r.CategoryName.lower()):
                    self._visible_rows.append(r)
        
        self.dg_families.ItemsSource = self._visible_rows
        self._update_counts()

    def _update_counts(self):
        """Update count labels in status bar."""
        total = len(self._visible_rows)
        selected = sum(1 for r in self._visible_rows if r.IsSelected)
        self.count_text.Text = "Total number of elements found {} | Selected {}".format(total, selected)

    # --- UI EVENTS & ACTIONS ---
    def scope_changed(self, sender, e):
        if hasattr(self, "rb_scope_all"):
            self._refresh_data()

    def category_changed(self, sender, e):
        if hasattr(self, "cb_category"):
            self._refresh_data()
            try:
                category_idx = self.cb_category.SelectedIndex
                if hasattr(self, "dg_families") and self.dg_families.Columns.Count > 1:
                    self.dg_families.Columns[1].IsReadOnly = (category_idx != 0)
            except Exception as ex:
                logger.debug("Error setting column IsReadOnly: {}".format(ex))

    def search_changed(self, sender, e):
        self._apply_filter()

    def refresh_click(self, sender, e):
        self._refresh_data()
        self.status_text.Text = "Data refreshed from Revit."

    def select_all_click(self, sender, e):
        for r in self._visible_rows:
            r.IsSelected = True
        self.dg_families.Items.Refresh()
        self._update_counts()

    def deselect_all_click(self, sender, e):
        for r in self._visible_rows:
            r.IsSelected = False
        self.dg_families.Items.Refresh()
        self._update_counts()

    # --- TEXT TRANSFORMATIONS ---
    def match_case_click(self, sender, e):
        pass

    def find_next_click(self, sender, e):
        """Highlights the next row matching the Find text."""
        find_val = (self.tb_find.Text or "").strip()
        if not find_val:
            return
        
        match_case = self.btn_match_case.IsChecked == True
        current_idx = self.dg_families.SelectedIndex
        
        # Look forward
        rows_len = len(self._visible_rows)
        for i in range(1, rows_len + 1):
            idx = (current_idx + i) % rows_len
            row = self._visible_rows[idx]
            
            f_match = find_val in row.FamilyName if match_case else find_val.lower() in row.FamilyName.lower()
            t_match = find_val in row.TypeName if match_case else find_val.lower() in row.TypeName.lower()
            
            if f_match or t_match:
                self.dg_families.SelectedIndex = idx
                self.dg_families.ScrollIntoView(row)
                break

    def find_all_click(self, sender, e):
        """Selects all rows matching the Find text."""
        find_val = (self.tb_find.Text or "").strip()
        if not find_val:
            return
        
        match_case = self.btn_match_case.IsChecked == True
        count = 0
        
        for row in self._visible_rows:
            f_match = find_val in row.FamilyName if match_case else find_val.lower() in row.FamilyName.lower()
            t_match = find_val in row.TypeName if match_case else find_val.lower() in row.TypeName.lower()
            
            if f_match or t_match:
                row.IsSelected = True
                count += 1
                
        self.dg_families.Items.Refresh()
        self._update_counts()
        self.status_text.Text = "Selected {} matching elements.".format(count)

    def replace_click(self, sender, e):
        """Replace Find with Replace in the currently selected row."""
        find_val = (self.tb_find.Text or "")
        replace_val = (self.tb_replace.Text or "")
        
        selected_row = self.dg_families.SelectedItem
        if not selected_row:
            forms.alert("Please select a row in the table first.")
            return
            
        match_case = self.btn_match_case.IsChecked == True
        
        # Family Name
        if match_case:
            selected_row.FamilyName = selected_row.FamilyName.replace(find_val, replace_val)
            selected_row.TypeName = selected_row.TypeName.replace(find_val, replace_val)
        else:
            # Case insensitive replace helper
            import re
            pattern = re.compile(re.escape(find_val), re.IGNORECASE)
            selected_row.FamilyName = pattern.sub(replace_val, selected_row.FamilyName)
            selected_row.TypeName = pattern.sub(replace_val, selected_row.TypeName)
            
        self.dg_families.Items.Refresh()

    def replace_all_click(self, sender, e):
        """Replace Find with Replace in all selected rows (or all visible rows if none selected)."""
        find_val = (self.tb_find.Text or "")
        replace_val = (self.tb_replace.Text or "")
        
        target_rows = [r for r in self._visible_rows if r.IsSelected]
        if not target_rows:
            target_rows = self._visible_rows # Fallback to all if none checked
            
        match_case = self.btn_match_case.IsChecked == True
        count = 0
        
        import re
        pattern = re.compile(re.escape(find_val), re.IGNORECASE)
        
        for r in target_rows:
            if match_case:
                r.FamilyName = r.FamilyName.replace(find_val, replace_val)
                r.TypeName = r.TypeName.replace(find_val, replace_val)
            else:
                r.FamilyName = pattern.sub(replace_val, r.FamilyName)
                r.TypeName = pattern.sub(replace_val, r.TypeName)
            count += 1
            
        self.dg_families.Items.Refresh()
        self.status_text.Text = "Staged Find & Replace for {} rows.".format(count)

    # --- PREFIX / SUFFIX ---
    def prefix_selected_click(self, sender, e):
        prefix = (self.tb_prefix.Text or "")
        for r in self._visible_rows:
            if r.IsSelected:
                r.FamilyName = prefix + r.FamilyName
                r.TypeName = prefix + r.TypeName
        self.dg_families.Items.Refresh()

    def prefix_all_click(self, sender, e):
        prefix = (self.tb_prefix.Text or "")
        for r in self._visible_rows:
            r.FamilyName = prefix + r.FamilyName
            r.TypeName = prefix + r.TypeName
        self.dg_families.Items.Refresh()

    def suffix_selected_click(self, sender, e):
        suffix = (self.tb_suffix.Text or "")
        for r in self._visible_rows:
            if r.IsSelected:
                r.FamilyName = r.FamilyName + suffix
                r.TypeName = r.TypeName + suffix
        self.dg_families.Items.Refresh()

    def suffix_all_click(self, sender, e):
        suffix = (self.tb_suffix.Text or "")
        for r in self._visible_rows:
            r.FamilyName = r.FamilyName + suffix
            r.TypeName = r.TypeName + suffix
        self.dg_families.Items.Refresh()

    # --- CASE CONVERSIONS ---
    def _apply_case_to_string(self, text, case_type):
        if case_type == "UPPER":
            return text.upper()
        elif case_type == "lower":
            return text.lower()
        elif case_type == "Title":
            # Simple Title Case
            return " ".join([w.capitalize() for w in text.split(" ")])
        elif case_type == "Sentence":
            # Sentence case
            if len(text) > 0:
                return text[0].upper() + text[1:].lower()
        return text

    def _apply_case_transformation(self, case_type):
        target_idx = self.cb_case_target.SelectedIndex
        # 0: Both Family & Type, 1: Family only, 2: Type only
        
        target_rows = [r for r in self._visible_rows if r.IsSelected]
        if not target_rows:
            target_rows = self._visible_rows # Fallback to all if none checked
            
        for r in target_rows:
            if target_idx == 0 or target_idx == 1:
                r.FamilyName = self._apply_case_to_string(r.FamilyName, case_type)
            if target_idx == 0 or target_idx == 2:
                r.TypeName = self._apply_case_to_string(r.TypeName, case_type)
                
        self.dg_families.Items.Refresh()
        self.status_text.Text = "Case changed to {}.".format(case_type)

    def case_upper_click(self, sender, e):
        self._apply_case_transformation("UPPER")

    def case_lower_click(self, sender, e):
        self._apply_case_transformation("lower")

    def case_title_click(self, sender, e):
        self._apply_case_transformation("Title")

    def case_sentence_click(self, sender, e):
        self._apply_case_transformation("Sentence")

    # --- OTHER ACTIONS ---
    def delete_row_click(self, sender, e):
        """Remove a family row from the staged list."""
        button = sender
        row = button.DataContext
        if row in self._all_rows:
            self._all_rows.remove(row)
        self._apply_filter()

    def clear_settings_click(self, sender, e):
        self.tb_find.Text = ""
        self.tb_replace.Text = ""
        self.tb_prefix.Text = ""
        self.tb_suffix.Text = ""
        self.status_text.Text = "Settings cleared."

    def undo_click(self, sender, e):
        """Reset DataGrid rows back to their original Revit values."""
        for r in self._all_rows:
            r.FamilyName = r.OriginalFamilyName
            r.TypeName = r.OriginalTypeName
            r.WorksetId = r.OriginalWorksetId
        self.dg_families.Items.Refresh()
        self.status_text.Text = "Undo completed (discarded staged changes)."

    def apply_click(self, sender, e):
        """Commit all renaming, case-changing, and workset changes back to Revit."""
        modified_rows = [r for r in self._all_rows if r.IsModified]
        if not modified_rows:
            forms.alert("No changes to apply.")
            return

        success_count = 0
        error_count = 0

        # Pre-build a dictionary mapping type ID to elements if any worksets are modified
        worksets_changed = any(doc.IsWorkshared and r.WorksetId != r.OriginalWorksetId for r in modified_rows)
        instances_by_type = {}
        if worksets_changed:
            try:
                all_instances = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
                for inst in all_instances:
                    try:
                        tid = inst.GetTypeId()
                        if tid and tid != ElementId.InvalidElementId:
                            tid_val = tid.IntegerValue
                            if tid_val not in instances_by_type:
                                instances_by_type[tid_val] = []
                            instances_by_type[tid_val].append(inst)
                    except Exception:
                        pass
            except Exception as ex:
                logger.debug("Error building instance type map: {}".format(ex))

        # Start Revit Transaction
        t = Transaction(doc, "T3Lab - Family Management Apply")
        t.Start()
        try:
            for r in modified_rows:
                try:
                    # 1. Rename Family (only for loadable families)
                    if r.FamilyName != r.OriginalFamilyName:
                        sanitized_fam = sanitize_name(r.FamilyName)
                        if r.is_loadable and hasattr(r.element, "Family"):
                            r.element.Family.Name = sanitized_fam
                        else:
                            # For system families, Model Groups, etc., Family Name is not editable
                            pass

                    # 2. Rename Type
                    if r.TypeName != r.OriginalTypeName:
                        sanitized_type = sanitize_name(r.TypeName)
                        r.element.Name = sanitized_type

                    # 3. Assign Workset (if workshared and workset changed)
                    if doc.IsWorkshared and r.WorksetId != r.OriginalWorksetId:
                        ws_val = r.WorksetId
                        # Set workset of the type
                        ws_param = r.element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                        if ws_param and not ws_param.IsReadOnly:
                            ws_param.Set(ws_val)
                            
                        # Set workset of all placed instances of this type
                        related_instances = instances_by_type.get(r.element.Id.IntegerValue, [])
                        for inst in related_instances:
                            try:
                                inst_ws_param = inst.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                                if inst_ws_param and not inst_ws_param.IsReadOnly:
                                    inst_ws_param.Set(ws_val)
                            except Exception:
                                pass

                    success_count += 1
                except Exception as ex:
                    error_count += 1
                    logger.debug("Error applying changes for row: {}\n{}".format(r.FamilyName, ex))

            t.Commit()
            forms.alert(
                "Successfully applied changes to {} elements.\nErrors encountered: {}".format(
                    success_count, error_count
                ),
                title="Apply Success"
            )
            # Reload to sync original state
            self._refresh_data()
        except Exception as ex:
            t.RollBack()
            forms.alert("Transaction failed: {}".format(ex), title="Transaction Error")

    def export_list_click(self, sender, e):
        """Export current DataGrid rows to a comma-separated text file."""
        dest_file = forms.save_file(filesfilter="Comma-Separated Values (*.csv)|*.csv", title="Export Family List")
        if not dest_file:
            return
            
        try:
            import csv
            with open(dest_file, "wb") as f:
                writer = csv.writer(f)
                writer.writerow(["Family Name", "Type Name", "Category", "Workset ID"])
                for r in self._visible_rows:
                    writer.writerow([
                        r.FamilyName.encode('utf-8'),
                        r.TypeName.encode('utf-8'),
                        r.CategoryName.encode('utf-8'),
                        r.WorksetId
                    ])
            self.status_text.Text = "Exported list to: {}".format(dest_file)
        except Exception as ex:
            forms.alert("Failed to export list:\n{}".format(ex))

    def tab_click(self, sender, e):
        """Switch view modes or trigger external panels (Export / Import / Worksets)."""
        tab_name = sender.Name
        if tab_name == "tab_edit":
            return

        # Trigger existing actions or dialogs
        gui_dir = os.path.dirname(os.path.abspath(__file__))
        lib_dir = os.path.dirname(gui_dir)
        ext_dir = os.path.dirname(lib_dir)

        if tab_name == "tab_export":
            # Launch Bulk Family Export (CAD to Family)
            try:
                cad_to_fam_dir = os.path.join(ext_dir, 'T3Lab.tab', 'Project.panel', 'Family Work.stack', 'CAD To Family.pushbutton')
                if os.path.exists(cad_to_fam_dir):
                    import imp
                    script_file = os.path.join(cad_to_fam_dir, 'script.py')
                    mod = imp.load_source('cad_to_family_module', script_file)
                    win = mod.BulkFamilyExportWindow()
                    self.Close()
                    win.ShowDialog()
            except Exception as ex:
                logger.debug("Failed to launch CAD To Family: {}".format(ex))
        elif tab_name == "tab_import":
            # Load families from local folder
            try:
                from GUI.FamilyLoaderDialog import show_family_loader
                self.Close()
                show_family_loader()
            except Exception as ex:
                logger.debug("Failed to launch Family Loader: {}".format(ex))
        elif tab_name == "tab_worksets":
            # Load worksets manager
            try:
                workset_dir = os.path.join(ext_dir, 'T3Lab.tab', 'Project.panel', 'Workset.stack', 'Workset.pushbutton')
                if os.path.exists(workset_dir):
                    import imp
                    script_file = os.path.join(workset_dir, 'script.py')
                    mod = imp.load_source('workset_module', script_file)
                    win = mod.WorksetManagerWindow()
                    self.Close()
                    win.ShowDialog()
            except Exception as ex:
                logger.debug("Failed to launch Workset Manager: {}".format(ex))

# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================
def show_family_management():
    # Load and show dialog
    xaml_path = os.path.join(os.path.dirname(__file__), "Tools", "FamilyManagement.xaml")
    dialog = FamilyManagementWindow(xaml_path)
    dialog.ShowDialog()
