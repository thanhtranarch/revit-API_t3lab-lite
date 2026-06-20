# -*- coding: utf-8 -*-
__title__ = "Split Columns"
__doc__ = "Split columns at selected levels with hosted elements report"

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import revit, DB, UI, forms
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Button, Label, CheckedListBox, FormBorderStyle, DockStyle, FormStartPosition, DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon, SaveFileDialog
from System.Drawing import Point, Size, Color, Font, FontStyle
import System
import os
from datetime import datetime

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# ============================================================
# Revit 2024-2027 compatibility helpers
# ============================================================
def _eid_int(eid):
    """Get integer value of an ElementId across Revit 2024-2027.
    .IntegerValue is deprecated in 2024+ and removed in 2026+; use .Value."""
    try:
        return eid.Value
    except:
        return eid.IntegerValue


class LevelSelectionForm(Form):
    def __init__(self, all_levels, column_info):
        self.all_levels = sorted(all_levels, key=lambda x: x.Elevation)
        self.column_info = column_info
        self.selected_levels = []
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self.Text = "Split Column - Select Levels"
        self.Width = 500
        self.Height = 600
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.FromArgb(254, 248, 231)
        
        header = Label()
        header.Text = "Select Levels to Split Column"
        header.Font = Font("Segoe UI", 10, FontStyle.Bold)
        header.BackColor = Color.FromArgb(240, 204, 136)
        header.ForeColor = Color.FromArgb(51, 51, 51)
        header.Height = 40
        header.Dock = DockStyle.Top
        header.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        header.Padding = System.Windows.Forms.Padding(10, 0, 0, 0)
        self.Controls.Add(header)
        
        info_label = Label()
        base_level = self.column_info['base_level']
        top_level = self.column_info['top_level']
        base_offset = self.column_info['base_offset']
        top_offset = self.column_info['top_offset']
        
        actual_base = base_level.Elevation + base_offset
        actual_top = top_level.Elevation + top_offset
        
        info_text = "Current Column:\n"
        info_text += "  Base: {} (Elev: {:.0f}mm, Offset: {:.0f}mm)\n".format(
            base_level.Name, base_level.Elevation * 304.8, base_offset * 304.8)
        info_text += "  Top: {} (Elev: {:.0f}mm, Offset: {:.0f}mm)\n".format(
            top_level.Name, top_level.Elevation * 304.8, top_offset * 304.8)
        info_text += "  Actual Range: {:.0f}mm to {:.0f}mm".format(
            min(actual_base, actual_top) * 304.8, max(actual_base, actual_top) * 304.8)
        
        info_label.Text = info_text
        info_label.Location = Point(10, 50)
        info_label.Size = Size(470, 80)
        info_label.Font = Font("Segoe UI", 9)
        self.Controls.Add(info_label)
        
        instruction = Label()
        instruction.Text = "Select levels where you want to split the column:"
        instruction.Location = Point(10, 140)
        instruction.Size = Size(470, 20)
        instruction.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(instruction)
        
        self.levels_list = CheckedListBox()
        self.levels_list.Location = Point(10, 165)
        self.levels_list.Size = Size(470, 350)
        self.levels_list.CheckOnClick = True
        self.levels_list.Font = Font("Segoe UI", 9)
        
        actual_min = min(actual_base, actual_top)
        actual_max = max(actual_base, actual_top)
        
        tolerance = 0.001
        
        for level in self.all_levels:
            level_elev = level.Elevation
            in_range = actual_min < level_elev < actual_max
            matches_start = (abs(level_elev - actual_base) < tolerance)
            matches_end = (abs(level_elev - actual_top) < tolerance)
            
            display_text = "{} (Elev: {:.0f}mm)".format(level.Name, level_elev * 304.8)
            
            if in_range:
                display_text += " [Within column range]"
            
            self.levels_list.Items.Add(display_text)
            
            if in_range and not matches_start and not matches_end:
                self.levels_list.SetItemChecked(self.levels_list.Items.Count - 1, True)
        
        self.Controls.Add(self.levels_list)
        
        ok_button = Button()
        ok_button.Text = "Split Column"
        ok_button.Location = Point(300, 525)
        ok_button.Size = Size(90, 30)
        ok_button.Font = Font("Segoe UI", 9)
        ok_button.Click += self.OnOK
        self.Controls.Add(ok_button)
        
        cancel_button = Button()
        cancel_button.Text = "Cancel"
        cancel_button.Location = Point(400, 525)
        cancel_button.Size = Size(70, 30)
        cancel_button.Font = Font("Segoe UI", 9)
        cancel_button.Click += self.OnCancel
        self.Controls.Add(cancel_button)
    
    def OnOK(self, sender, args):
        checked_indices = self.levels_list.CheckedIndices
        if checked_indices.Count == 0:
            forms.alert("Please select at least one level!", exitscript=False)
            return
        
        self.selected_levels = [self.all_levels[i] for i in checked_indices]
        self.DialogResult = DialogResult.OK
        self.Close()
    
    def OnCancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

def get_all_levels():
    levels = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
    return sorted(levels, key=lambda x: x.Elevation)

def get_column_info(column):
    base_level_id = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()
    top_level_id = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()
    
    base_level = doc.GetElement(base_level_id)
    top_level = doc.GetElement(top_level_id)
    
    base_offset_param = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
    top_offset_param = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
    
    base_offset = base_offset_param.AsDouble() if base_offset_param else 0.0
    top_offset = top_offset_param.AsDouble() if top_offset_param else 0.0
    
    return {
        'base_level': base_level,
        'top_level': top_level,
        'base_level_name': base_level.Name,
        'top_level_name': top_level.Name,
        'base_offset': base_offset,
        'top_offset': top_offset
    }

def get_hosted_elements_info(column):
    """Collect detailed information about hosted elements"""
    hosted_info = []
    
    try:
        dependent_ids = column.GetDependentElements(None)
        
        for dep_id in dependent_ids:
            dep_elem = doc.GetElement(dep_id)
            if dep_elem:
                category = dep_elem.Category
                if category:
                    cat_name = category.Name
                    if "Analytical" in cat_name:
                        continue
                    
                    elem_type = doc.GetElement(dep_elem.GetTypeId())
                    type_name = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if elem_type else "N/A"
                    
                    elem_name = "N/A"
                    name_param = dep_elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                    if name_param:
                        elem_name = name_param.AsString() if name_param.AsString() else "N/A"
                    
                    hosted_info.append({
                        'id': _eid_int(dep_elem.Id),
                        'category': cat_name,
                        'type': type_name,
                        'name': elem_name,
                        'element': dep_elem
                    })
        
        return hosted_info
    except Exception as e:
        print("Error getting hosted elements: {}".format(str(e)))
        return []

def get_save_file_path(original_column_id):
    """Show SaveFileDialog for user to choose save location"""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = "Column_Split_Report_{}_ID{}.csv".format(timestamp, original_column_id)
        
        save_dialog = SaveFileDialog()
        save_dialog.Title = "Save Hosted Elements Report"
        save_dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        save_dialog.FileName = default_filename
        
        try:
            user_profile = os.environ.get('USERPROFILE')
            if user_profile:
                desktop = os.path.join(user_profile, 'Desktop')
                if os.path.exists(desktop):
                    save_dialog.InitialDirectory = desktop
        except:
            pass
        
        if save_dialog.ShowDialog() == DialogResult.OK:
            return save_dialog.FileName
        else:
            return None
    except Exception as e:
        print("Error showing save dialog: {}".format(str(e)))
        return None

def export_hosted_elements_csv(original_column_id, hosted_elements, new_column_ids, output_path):
    """Export CSV report of hosted elements"""
    try:
        lines = []
        
        lines.append("="*80)
        lines.append("COLUMN SPLIT REPORT - HOSTED ELEMENTS")
        lines.append("Generated: {}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        lines.append("="*80)
        lines.append("")
        
        lines.append("ORIGINAL COLUMN")
        lines.append("-" * 80)
        lines.append("Column ID,{}".format(original_column_id))
        lines.append("Status,DELETED (split into {} segments)".format(len(new_column_ids)))
        lines.append("")
        
        lines.append("NEW COLUMN SEGMENTS")
        lines.append("-" * 80)
        lines.append("Segment,Column ID")
        for i, col_id in enumerate(new_column_ids, 1):
            lines.append("Segment {},{}".format(i, col_id))
        lines.append("")
        
        lines.append("HOSTED ELEMENTS REQUIRING MANUAL RECREATION")
        lines.append("-" * 80)
        lines.append("Element ID,Category,Type,Name,Action Required")
        
        for hosted in hosted_elements:
            line = "{},{},{},{},Manually recreate/rehost this element".format(
                hosted['id'],
                hosted['category'],
                hosted['type'],
                hosted['name']
            )
            lines.append(line)
        
        lines.append("")
        
        lines.append("SUMMARY")
        lines.append("-" * 80)
        lines.append("Total hosted elements,{}".format(len(hosted_elements)))
        lines.append("Action required,Manual recreation of all hosted elements")
        lines.append("")
        lines.append("="*80)
        lines.append("Copyright (c) {} by Dang Quoc Truong (DQT)".format(datetime.now().year))
        lines.append("="*80)
        
        with open(output_path, 'w') as f:
            for line in lines:
                f.write(line + '\n')
        
        return True, output_path
    except Exception as e:
        return False, str(e)

def split_column_at_levels(column, split_levels):
    info = get_column_info(column)
    
    location = column.Location
    if not isinstance(location, LocationPoint):
        return False, "Column does not have a point location", None, None
    
    point = location.Point
    column_symbol = doc.GetElement(column.GetTypeId())
    
    base_level = info['base_level']
    top_level = info['top_level']
    base_offset = info['base_offset']
    top_offset = info['top_offset']
    
    actual_base_elev = base_level.Elevation + base_offset
    actual_top_elev = top_level.Elevation + top_offset
    
    print("="*50)
    print("SPLITTING COLUMN")
    print("Original Column ID: {}".format(_eid_int(column.Id)))
    print("Original Base: {} (Elev: {:.0f}mm) + Offset: {:.0f}mm = {:.0f}mm".format(
        base_level.Name, base_level.Elevation * 304.8, base_offset * 304.8, actual_base_elev * 304.8))
    print("Original Top: {} (Elev: {:.0f}mm) + Offset: {:.0f}mm = {:.0f}mm".format(
        top_level.Name, top_level.Elevation * 304.8, top_offset * 304.8, actual_top_elev * 304.8))
    print("Split at levels: {}".format([lv.Name for lv in split_levels]))
    
    is_descending = actual_base_elev > actual_top_elev
    print("Column direction: {}".format("descending (top to bottom)" if is_descending else "ascending (bottom to top)"))
    
    tolerance = 0.001
    
    level_positions = []
    level_positions.append((base_level, 0, base_offset, actual_base_elev))
    
    for i, split_lv in enumerate(split_levels):
        split_elev = split_lv.Elevation
        matches_start = abs(split_elev - actual_base_elev) < tolerance
        matches_end = abs(split_elev - actual_top_elev) < tolerance
        
        if not matches_start and not matches_end:
            level_positions.append((split_lv, i + 1, 0.0, split_elev))
            print("Adding split level: {} at {:.0f}mm".format(split_lv.Name, split_elev * 304.8))
        else:
            print("Skipping split level {} (matches start/end)".format(split_lv.Name))
    
    level_positions.append((top_level, len(split_levels) + 1, top_offset, actual_top_elev))
    
    if is_descending:
        level_positions.sort(key=lambda x: x[3], reverse=True)
    else:
        level_positions.sort(key=lambda x: x[3])
    
    print("\nLevel positions after sorting by actual elevation:")
    for lv, pos, offset, actual_elev in level_positions:
        print("  {} (pos={}, offset={:.0f}mm, level_elev={:.0f}mm, actual_elev={:.0f}mm)".format(
            lv.Name, pos, offset * 304.8, lv.Elevation * 304.8, actual_elev * 304.8))
    print("="*50)
    
    columns_created = []
    new_column_ids = []
    
    for i in range(len(level_positions) - 1):
        current_base_level, current_base_pos, current_base_stored_offset, _ = level_positions[i]
        current_top_level, current_top_pos, current_top_stored_offset, _ = level_positions[i + 1]
        
        segment_base_offset = 0.0
        segment_top_offset = 0.0
        
        if current_base_pos == 0:
            segment_base_offset = base_offset
        
        if current_top_pos == len(split_levels) + 1:
            segment_top_offset = top_offset
        
        print("\nCreating column {}: {} -> {}".format(
            i+1, current_base_level.Name, current_top_level.Name))
        print("  Base position: {}, offset: {:.0f}mm".format(current_base_pos, segment_base_offset * 304.8))
        print("  Top position: {}, offset: {:.0f}mm".format(current_top_pos, segment_top_offset * 304.8))
        
        new_column = doc.Create.NewFamilyInstance(
            point, column_symbol, current_base_level, StructuralType.Column)
        
        new_column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).Set(current_top_level.Id)
        new_column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(segment_base_offset)
        new_column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).Set(segment_top_offset)
        
        new_column_ids.append(_eid_int(new_column.Id))
        
        col_desc = "Column {}: {}".format(i+1, current_base_level.Name)
        if segment_base_offset != 0:
            col_desc += " {:+.0f}mm".format(segment_base_offset * 304.8)
        col_desc += " -> {}".format(current_top_level.Name)
        if segment_top_offset != 0:
            col_desc += " {:+.0f}mm".format(segment_top_offset * 304.8)
        col_desc += " (ID: {})".format(_eid_int(new_column.Id))
        
        columns_created.append(col_desc)
    
    original_id = _eid_int(column.Id)
    doc.Delete(column.Id)
    print("\nOriginal column deleted (ID: {})".format(original_id))
    print("New column IDs: {}".format(new_column_ids))
    print("="*50)
    
    return True, columns_created, original_id, new_column_ids

def main():
    try:
        selection = uidoc.Selection
        selected_ids = selection.GetElementIds()
        
        if selected_ids.Count == 0:
            forms.alert("Please select at least one column!", exitscript=True)
            return
        
        columns = []
        for elem_id in selected_ids:
            elem = doc.GetElement(elem_id)
            if isinstance(elem, FamilyInstance) and elem.Category is not None and elem.Category.BuiltInCategory == BuiltInCategory.OST_StructuralColumns:
                columns.append(elem)
        
        if not columns:
            forms.alert("No structural columns selected!", exitscript=True)
            return
        
        if len(columns) > 1:
            forms.alert("Please select only ONE column at a time!", exitscript=True)
            return
        
    except Exception as ex:
        forms.alert("Error selecting columns: {}".format(str(ex)), exitscript=True)
        return
    
    column = columns[0]
    
    hosted_elements = get_hosted_elements_info(column)
    
    if hosted_elements:
        hosted_summary = "\n".join(["- {} [{}] (ID: {})".format(
            h['category'], h['type'], h['id']) for h in hosted_elements])
        
        info_msg = "INFO: This column has {} hosted element(s):\n\n{}\n\n".format(
            len(hosted_elements), hosted_summary)
        info_msg += "A CSV report will be exported with details for manual recreation.\n\n"
        info_msg += "Continue?"
        
        result = MessageBox.Show(
            info_msg,
            "Hosted Elements Detected",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Information
        )
        
        if result != DialogResult.Yes:
            return
    
    column_info = get_column_info(column)
    all_levels = get_all_levels()
    
    form = LevelSelectionForm(all_levels, column_info)
    result = form.ShowDialog()
    
    if result != DialogResult.OK:
        return
    
    selected_levels = form.selected_levels
    
    if not selected_levels:
        forms.alert("No levels selected!", exitscript=True)
        return
    
    t = Transaction(doc, "DQT - Split Column at Levels")
    t.Start()
    
    try:
        success, columns_desc, original_id, new_ids = split_column_at_levels(column, selected_levels)
        
        if success:
            t.Commit()
            
            report_path = None
            if hosted_elements:
                try:
                    report_path = get_save_file_path(original_id)
                    
                    if report_path:
                        export_success, export_result = export_hosted_elements_csv(
                            original_id, hosted_elements, new_ids, report_path)
                        
                        if not export_success:
                            print("ERROR: Failed to export report: {}".format(export_result))
                            report_path = None
                    else:
                        print("User cancelled save dialog")
                except Exception as e:
                    print("ERROR: Exception during report export: {}".format(str(e)))
                    import traceback
                    traceback.print_exc()
                    report_path = None
            
            message = "Column split successfully!\n\n"
            for col_info in columns_desc:
                message += col_info + "\n"
            
            if hosted_elements:
                message += "\n" + "="*50 + "\n"
                message += "HOSTED ELEMENTS REPORT\n"
                message += "="*50 + "\n"
                message += "{} hosted element(s) were detected.\n".format(len(hosted_elements))
                if report_path:
                    message += "CSV report saved to:\n{}".format(report_path)
                else:
                    message += "Report export was cancelled or failed."
            
            forms.alert(message)
        else:
            t.RollBack()
            forms.alert("Error: {}".format(columns_desc))
    
    except Exception as e:
        t.RollBack()
        forms.alert("Error: {}".format(str(e)), exitscript=True)
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()