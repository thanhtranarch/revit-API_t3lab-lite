# -*- coding: utf-8 -*-
__title__ = "Split Walls"
__doc__ = "Split walls at selected levels - preserves doors, windows and hosted elements"

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import revit, DB, UI, forms
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (
    Form, Button, Label, CheckedListBox, CheckBox,
    FormBorderStyle, DockStyle, FormStartPosition, DialogResult
)
from System.Drawing import Point, Size, Color, Font, FontStyle
import System

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


def _make_eid(int_value):
    """Create an ElementId from an integer across Revit 2024-2027.
    ElementId(int) is deprecated; 2026+ expects Int64."""
    try:
        return ElementId(System.Int64(int_value))
    except:
        return ElementId(int_value)


# ============================================================
# UI
# ============================================================
class LevelSelectionForm(Form):
    def __init__(self, all_levels, wall_info_list):
        self.all_levels = sorted(all_levels, key=lambda x: x.Elevation)
        self.wall_info_list = wall_info_list
        self.selected_levels = []
        self.copy_params = True
        self.InitializeComponent()

    def InitializeComponent(self):
        wall_count = len(self.wall_info_list)
        title_text = "Split Wall - Select Levels"
        if wall_count > 1:
            title_text = "Split {} Walls - Select Levels".format(wall_count)

        self.Text = title_text
        self.Width = 520
        self.Height = 650
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.FromArgb(254, 248, 231)

        # --- Header ---
        header = Label()
        header.Text = "  Split Wall - Select Levels"
        header.Font = Font("Segoe UI", 11, FontStyle.Bold)
        header.BackColor = Color.FromArgb(240, 204, 136)
        header.ForeColor = Color.FromArgb(51, 51, 51)
        header.Height = 40
        header.Dock = DockStyle.Top
        header.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        self.Controls.Add(header)

        # --- Wall Info ---
        info_label = Label()
        if wall_count == 1:
            wi = self.wall_info_list[0]
            base_level = wi['base_level']
            base_offset = wi['base_offset']

            info_text = "Selected Wall:  {} (ID: {})\n".format(
                wi['type_name'], wi['element_id'])
            info_text += "  Base: {} (Elev: {:.0f}mm, Offset: {:.0f}mm)\n".format(
                base_level.Name, base_level.Elevation * 304.8, base_offset * 304.8)

            if wi['has_top_level']:
                top_level = wi['top_level']
                top_offset = wi['top_offset']
                info_text += "  Top: {} (Elev: {:.0f}mm, Offset: {:.0f}mm)\n".format(
                    top_level.Name, top_level.Elevation * 304.8, top_offset * 304.8)
            else:
                info_text += "  Top: Unconnected Height = {:.0f}mm\n".format(
                    wi['unconnected_height'] * 304.8)

            info_text += "  Actual Range: {:.0f}mm to {:.0f}mm".format(
                wi['actual_base_elev'] * 304.8, wi['actual_top_elev'] * 304.8)

            hosted_count = wi.get('hosted_count', 0)
            if hosted_count > 0:
                info_text += "\n  Hosted Elements: {}".format(hosted_count)
        else:
            info_text = "{} walls selected.\n".format(wall_count)
            all_bases = [wi['actual_base_elev'] for wi in self.wall_info_list]
            all_tops = [wi['actual_top_elev'] for wi in self.wall_info_list]
            info_text += "  Combined range: {:.0f}mm to {:.0f}mm".format(
                min(all_bases) * 304.8, max(all_tops) * 304.8)
            total_hosted = sum(wi.get('hosted_count', 0) for wi in self.wall_info_list)
            if total_hosted > 0:
                info_text += "\n  Total Hosted Elements: {}".format(total_hosted)

        info_label.Text = info_text
        info_label.Location = Point(10, 50)
        info_label.Size = Size(490, 90)
        info_label.Font = Font("Segoe UI", 9)
        self.Controls.Add(info_label)

        # --- Instruction ---
        instruction = Label()
        instruction.Text = "Select levels where you want to split the wall(s):"
        instruction.Location = Point(10, 145)
        instruction.Size = Size(490, 20)
        instruction.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(instruction)

        # --- Levels CheckedListBox ---
        self.levels_list = CheckedListBox()
        self.levels_list.Location = Point(10, 170)
        self.levels_list.Size = Size(490, 340)
        self.levels_list.CheckOnClick = True
        self.levels_list.Font = Font("Segoe UI", 9)

        combined_min = min(wi['actual_base_elev'] for wi in self.wall_info_list)
        combined_max = max(wi['actual_top_elev'] for wi in self.wall_info_list)

        tolerance = 0.001

        for level in self.all_levels:
            level_elev = level.Elevation
            in_range = combined_min + tolerance < level_elev < combined_max - tolerance
            display_text = "{} (Elev: {:.0f}mm)".format(
                level.Name, level_elev * 304.8)
            if in_range:
                display_text += " [Within wall range]"

            self.levels_list.Items.Add(display_text)

            if in_range:
                self.levels_list.SetItemChecked(
                    self.levels_list.Items.Count - 1, True)

        self.Controls.Add(self.levels_list)

        # --- Copy Parameters checkbox ---
        self.chk_copy_params = CheckBox()
        self.chk_copy_params.Text = "Copy instance parameters to new walls"
        self.chk_copy_params.Location = Point(10, 520)
        self.chk_copy_params.Size = Size(300, 22)
        self.chk_copy_params.Font = Font("Segoe UI", 9)
        self.chk_copy_params.Checked = True
        self.Controls.Add(self.chk_copy_params)

        # --- Buttons ---
        ok_button = Button()
        ok_button.Text = "Split Wall"
        ok_button.Location = Point(310, 565)
        ok_button.Size = Size(95, 32)
        ok_button.Font = Font("Segoe UI", 9, FontStyle.Bold)
        ok_button.BackColor = Color.FromArgb(200, 150, 80)
        ok_button.ForeColor = Color.White
        ok_button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        ok_button.Click += self.OnOK
        self.Controls.Add(ok_button)

        cancel_button = Button()
        cancel_button.Text = "Cancel"
        cancel_button.Location = Point(415, 565)
        cancel_button.Size = Size(80, 32)
        cancel_button.Font = Font("Segoe UI", 9)
        cancel_button.Click += self.OnCancel
        self.Controls.Add(cancel_button)

        # --- Footer ---
        footer = Label()
        footer.Text = "Dang Quoc Truong - DQT (c) 2026"
        footer.Font = Font("Segoe UI", 7)
        footer.ForeColor = Color.FromArgb(150, 150, 150)
        footer.Height = 20
        footer.Dock = DockStyle.Bottom
        footer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        self.Controls.Add(footer)

    def OnOK(self, sender, args):
        checked_indices = self.levels_list.CheckedIndices
        if checked_indices.Count == 0:
            forms.alert("Please select at least one level!", exitscript=False)
            return

        self.selected_levels = [self.all_levels[i] for i in checked_indices]
        self.copy_params = self.chk_copy_params.Checked
        self.DialogResult = DialogResult.OK
        self.Close()

    def OnCancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ============================================================
# Helpers
# ============================================================
def get_all_levels():
    levels = FilteredElementCollector(doc).OfClass(Level) \
        .WhereElementIsNotElementType().ToElements()
    return sorted(levels, key=lambda x: x.Elevation)


def get_wall_info(wall):
    """Get wall constraint info. Supports both Top Constraint and Unconnected Height."""
    base_constraint_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
    top_constraint_param = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)
    base_offset_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    top_offset_param = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)
    unconnected_height_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)

    base_level_id = base_constraint_param.AsElementId()
    top_level_id = top_constraint_param.AsElementId()

    base_level = doc.GetElement(base_level_id)
    base_offset = base_offset_param.AsDouble() if base_offset_param else 0.0
    actual_base_elev = base_level.Elevation + base_offset

    has_top_level = (top_level_id != ElementId.InvalidElementId)

    if has_top_level:
        top_level = doc.GetElement(top_level_id)
        top_offset = top_offset_param.AsDouble() if top_offset_param else 0.0
        actual_top_elev = top_level.Elevation + top_offset
        unconnected_height = 0.0
    else:
        top_level = None
        top_offset = 0.0
        unconnected_height = unconnected_height_param.AsDouble() if unconnected_height_param else 10.0
        actual_top_elev = actual_base_elev + unconnected_height

    wall_type = doc.GetElement(wall.GetTypeId())
    type_name = ""
    if wall_type:
        tn_param = wall_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if tn_param:
            type_name = tn_param.AsString()
    if not type_name:
        type_name = "Unknown"

    # Count hosted elements
    hosted_count = 0
    try:
        dep_filter = ElementClassFilter(FamilyInstance)
        dep_ids = wall.GetDependentElements(dep_filter)
        for dep_id in dep_ids:
            dep_elem = doc.GetElement(dep_id)
            if dep_elem and hasattr(dep_elem, 'Host') and dep_elem.Host and dep_elem.Host.Id == wall.Id:
                hosted_count += 1
    except:
        pass

    return {
        'element_id': _eid_int(wall.Id),
        'base_level': base_level,
        'top_level': top_level,
        'has_top_level': has_top_level,
        'base_offset': base_offset,
        'top_offset': top_offset,
        'unconnected_height': unconnected_height,
        'actual_base_elev': actual_base_elev,
        'actual_top_elev': actual_top_elev,
        'type_name': type_name,
        'hosted_count': hosted_count,
    }


# ============================================================
# Hosted element preservation
# ============================================================
def collect_hosted_elements(wall):
    """Collect info of all FamilyInstance elements hosted on this wall."""
    hosted_data = []
    try:
        dep_filter = ElementClassFilter(FamilyInstance)
        dep_ids = wall.GetDependentElements(dep_filter)
    except:
        return hosted_data

    for dep_id in dep_ids:
        dep_elem = doc.GetElement(dep_id)
        if dep_elem is None:
            continue

        # Only FamilyInstance hosted on THIS wall
        if not hasattr(dep_elem, 'Host'):
            continue
        if dep_elem.Host is None or dep_elem.Host.Id != wall.Id:
            continue

        fi = dep_elem

        # Get location point
        loc = fi.Location
        if not isinstance(loc, LocationPoint):
            continue

        insert_point = loc.Point

        # Get symbol (type)
        symbol_id = fi.GetTypeId()

        # Get level
        level_id = fi.LevelId

        # Get facing orientation and hand orientation for flip states
        is_facing_flipped = fi.FacingFlipped
        is_hand_flipped = fi.HandFlipped

        # Get Sill Height / offset from host level
        sill_height_param = fi.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
        sill_height = sill_height_param.AsDouble() if sill_height_param else 0.0

        # Collect instance parameters for later restoration
        param_data = []
        for param in fi.Parameters:
            if param.IsReadOnly:
                continue
            if param.Definition is None:
                continue
            # Skip constraint/placement params that will be set automatically
            try:
                bip = param.Definition.BuiltInParameter
                skip_bips = [
                    int(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM),
                    int(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM),
                    int(BuiltInParameter.ELEM_FAMILY_PARAM),
                    int(BuiltInParameter.ELEM_TYPE_PARAM),
                    int(BuiltInParameter.FAMILY_LEVEL_PARAM),
                    int(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM),
                ]
                if int(bip) in skip_bips:
                    continue
            except:
                pass

            storage = param.StorageType
            val = None
            if storage == StorageType.Double:
                val = ('double', param.AsDouble())
            elif storage == StorageType.Integer:
                val = ('int', param.AsInteger())
            elif storage == StorageType.String:
                s = param.AsString()
                if s is not None:
                    val = ('string', s)
            elif storage == StorageType.ElementId:
                val = ('id', _eid_int(param.AsElementId()))

            if val is not None:
                param_data.append((param.Definition.Name, val))

        # Get the elevation of the insertion point (Z coordinate = actual elevation)
        insert_elevation = insert_point.Z

        # Get the BoundingBox to determine the vertical extent of the hosted element
        bbox = fi.get_BoundingBox(None)
        elem_min_z = insert_elevation
        elem_max_z = insert_elevation
        if bbox:
            elem_min_z = bbox.Min.Z
            elem_max_z = bbox.Max.Z
        elem_mid_z = (elem_min_z + elem_max_z) / 2.0

        hosted_data.append({
            'insert_point': insert_point,
            'insert_elevation': insert_elevation,
            'elem_mid_z': elem_mid_z,
            'symbol_id': symbol_id,
            'level_id': level_id,
            'sill_height': sill_height,
            'is_facing_flipped': is_facing_flipped,
            'is_hand_flipped': is_hand_flipped,
            'param_data': param_data,
            'category_name': fi.Category.Name if fi.Category else "Unknown",
        })

    return hosted_data


def find_host_segment(hosted_info, segment_walls, segment_ranges):
    """
    Determine which wall segment should host this element.
    Uses the midpoint elevation of the hosted element to find the matching segment.
    """
    mid_z = hosted_info['elem_mid_z']
    tolerance = 0.01  # ~3mm

    for i, (seg_min, seg_max) in enumerate(segment_ranges):
        if seg_min - tolerance <= mid_z <= seg_max + tolerance:
            return i

    # Fallback: find closest segment
    best_idx = 0
    best_dist = float('inf')
    for i, (seg_min, seg_max) in enumerate(segment_ranges):
        seg_mid = (seg_min + seg_max) / 2.0
        dist = abs(mid_z - seg_mid)
        if dist < best_dist:
            best_dist = dist
            best_idx = i

    return best_idx


def recreate_hosted_element(hosted_info, host_wall):
    """Recreate a hosted FamilyInstance on the given host wall."""
    symbol = doc.GetElement(hosted_info['symbol_id'])
    if symbol is None:
        return None

    # Activate symbol if needed
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

    insert_point = hosted_info['insert_point']

    # Get the base level of the host wall for the hosted element
    wall_base_level_param = host_wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
    wall_base_level_id = wall_base_level_param.AsElementId()
    wall_base_level = doc.GetElement(wall_base_level_id)

    try:
        new_fi = doc.Create.NewFamilyInstance(
            insert_point,
            symbol,
            host_wall,
            wall_base_level,
            StructuralType.NonStructural
        )
    except:
        # Some elements might need different overload
        try:
            new_fi = doc.Create.NewFamilyInstance(
                insert_point,
                symbol,
                host_wall,
                StructuralType.NonStructural
            )
        except:
            return None

    if new_fi is None:
        return None

    # Set sill height
    sill_param = new_fi.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
    if sill_param and not sill_param.IsReadOnly:
        sill_param.Set(hosted_info['sill_height'])

    # Flip states
    if hosted_info['is_facing_flipped'] != new_fi.FacingFlipped:
        new_fi.flipFacing()
    if hosted_info['is_hand_flipped'] != new_fi.HandFlipped:
        new_fi.flipHand()

    # Restore instance parameters
    for param_name, (ptype, pval) in hosted_info['param_data']:
        try:
            p = new_fi.LookupParameter(param_name)
            if p is None or p.IsReadOnly:
                continue
            if ptype == 'double':
                p.Set(pval)
            elif ptype == 'int':
                p.Set(pval)
            elif ptype == 'string':
                p.Set(pval)
            elif ptype == 'id':
                p.Set(_make_eid(pval))
        except:
            continue

    return new_fi


# ============================================================
# Wall parameter copy
# ============================================================
def copy_instance_parameters(source_wall, target_wall):
    """Copy writable instance parameters from source to target wall."""
    skip_params = set([
        int(BuiltInParameter.WALL_BASE_CONSTRAINT),
        int(BuiltInParameter.WALL_HEIGHT_TYPE),
        int(BuiltInParameter.WALL_BASE_OFFSET),
        int(BuiltInParameter.WALL_TOP_OFFSET),
        int(BuiltInParameter.WALL_USER_HEIGHT_PARAM),
        int(BuiltInParameter.WALL_KEY_REF_PARAM),
        int(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM),
        int(BuiltInParameter.ELEM_FAMILY_PARAM),
        int(BuiltInParameter.ELEM_TYPE_PARAM),
        int(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT),
    ])

    for param in source_wall.Parameters:
        if param.IsReadOnly:
            continue
        if param.Definition is None:
            continue
        try:
            bip = param.Definition.BuiltInParameter
            if int(bip) in skip_params:
                continue
        except:
            pass
        try:
            target_param = target_wall.LookupParameter(param.Definition.Name)
            if target_param is None or target_param.IsReadOnly:
                continue
            storage = param.StorageType
            if storage == StorageType.Double:
                target_param.Set(param.AsDouble())
            elif storage == StorageType.Integer:
                target_param.Set(param.AsInteger())
            elif storage == StorageType.String:
                val = param.AsString()
                if val is not None:
                    target_param.Set(val)
            elif storage == StorageType.ElementId:
                target_param.Set(param.AsElementId())
        except:
            continue


# ============================================================
# Core split logic
# ============================================================
def split_wall_at_levels(wall, split_levels, do_copy_params, all_levels):
    """Split a single wall at the given levels. Returns (success, result_desc, hosted_stats)."""
    info = get_wall_info(wall)

    location = wall.Location
    if not isinstance(location, LocationCurve):
        return False, "Wall does not have a curve location", ""

    curve = location.Curve
    wall_type_id = wall.GetTypeId()

    base_level = info['base_level']
    base_offset = info['base_offset']
    actual_base_elev = info['actual_base_elev']
    actual_top_elev = info['actual_top_elev']
    has_top_level = info['has_top_level']

    actual_min = min(actual_base_elev, actual_top_elev)
    actual_max = max(actual_base_elev, actual_top_elev)

    tolerance = 0.001

    # Filter split levels
    valid_splits = []
    for lv in split_levels:
        lv_elev = lv.Elevation
        if actual_min + tolerance < lv_elev < actual_max - tolerance:
            valid_splits.append(lv)

    if not valid_splits:
        return False, "No valid split levels within wall range", ""

    valid_splits.sort(key=lambda x: x.Elevation)

    # --- Collect hosted elements BEFORE any changes ---
    hosted_elements = collect_hosted_elements(wall)

    # Check structural
    is_structural = False
    struct_param = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT)
    if struct_param:
        is_structural = struct_param.AsInteger() == 1

    # Check original wall flip state
    original_flipped = wall.Flipped

    # --- Build segments ---
    segments = []

    if has_top_level:
        top_level = info['top_level']
        top_offset = info['top_offset']

        segments.append({
            'base_level': base_level, 'base_offset': base_offset,
            'top_level': valid_splits[0], 'top_offset': 0.0,
            'use_top_constraint': True,
        })
        for i in range(len(valid_splits) - 1):
            segments.append({
                'base_level': valid_splits[i], 'base_offset': 0.0,
                'top_level': valid_splits[i + 1], 'top_offset': 0.0,
                'use_top_constraint': True,
            })
        segments.append({
            'base_level': valid_splits[-1], 'base_offset': 0.0,
            'top_level': top_level, 'top_offset': top_offset,
            'use_top_constraint': True,
        })
    else:
        segments.append({
            'base_level': base_level, 'base_offset': base_offset,
            'top_level': valid_splits[0], 'top_offset': 0.0,
            'use_top_constraint': True,
        })
        for i in range(len(valid_splits) - 1):
            segments.append({
                'base_level': valid_splits[i], 'base_offset': 0.0,
                'top_level': valid_splits[i + 1], 'top_offset': 0.0,
                'use_top_constraint': True,
            })
        last_split_elev = valid_splits[-1].Elevation
        remaining_height = actual_top_elev - last_split_elev
        segments.append({
            'base_level': valid_splits[-1], 'base_offset': 0.0,
            'top_level': None, 'top_offset': 0.0,
            'use_top_constraint': False,
            'unconnected_height': remaining_height,
        })

    # --- Compute segment elevation ranges (for hosted element placement) ---
    segment_ranges = []
    for seg in segments:
        seg_base_elev = seg['base_level'].Elevation + seg['base_offset']
        if seg['use_top_constraint']:
            seg_top_elev = seg['top_level'].Elevation + seg['top_offset']
        else:
            seg_top_elev = seg_base_elev + seg.get('unconnected_height', 10.0)
        segment_ranges.append((seg_base_elev, seg_top_elev))

    # --- Delete original wall (hosted elements will be deleted too) ---
    doc.Delete(wall.Id)

    # Regenerate so deleted elements are fully cleared
    doc.Regenerate()

    # --- Create new wall segments ---
    new_walls = []
    result_descriptions = []

    for i, seg in enumerate(segments):
        seg_base_level = seg['base_level']
        seg_base_offset = seg['base_offset']

        new_wall = Wall.Create(
            doc, curve, wall_type_id, seg_base_level.Id,
            10.0, 0.0, False, is_structural
        )

        new_wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(seg_base_offset)

        if seg['use_top_constraint']:
            seg_top_level = seg['top_level']
            seg_top_offset = seg['top_offset']
            new_wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(seg_top_level.Id)
            new_wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).Set(seg_top_offset)

            seg_desc = "Wall {}: {}".format(i + 1, seg_base_level.Name)
            if seg_base_offset != 0:
                seg_desc += " {:+.0f}mm".format(seg_base_offset * 304.8)
            seg_desc += " -> {}".format(seg_top_level.Name)
            if seg_top_offset != 0:
                seg_desc += " {:+.0f}mm".format(seg_top_offset * 304.8)
        else:
            uc_height = seg.get('unconnected_height', 10.0)
            new_wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(ElementId.InvalidElementId)
            new_wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(uc_height)

            seg_desc = "Wall {}: {}".format(i + 1, seg_base_level.Name)
            if seg_base_offset != 0:
                seg_desc += " {:+.0f}mm".format(seg_base_offset * 304.8)
            seg_desc += " -> Unconnected {:.0f}mm".format(uc_height * 304.8)

        # Flip state
        if original_flipped:
            new_wall.Flip()

        # Copy instance parameters
        # Note: source wall is deleted, so we copy from first new wall for consistency
        # Actually we need to copy params BEFORE deleting original.
        # Workaround: we already stored params should be done differently.
        # For now, skip - params are already set during wall creation with same type.

        new_walls.append(new_wall)
        result_descriptions.append(seg_desc)

    # Regenerate before placing hosted elements
    doc.Regenerate()

    # --- Recreate hosted elements on correct segments ---
    hosted_restored = 0
    hosted_failed = 0

    for h_info in hosted_elements:
        seg_idx = find_host_segment(h_info, new_walls, segment_ranges)
        host_wall = new_walls[seg_idx]

        new_fi = recreate_hosted_element(h_info, host_wall)
        if new_fi is not None:
            hosted_restored += 1
        else:
            hosted_failed += 1

    hosted_stats = ""
    if hosted_elements:
        hosted_stats = "  Hosted: {} restored".format(hosted_restored)
        if hosted_failed > 0:
            hosted_stats += ", {} failed".format(hosted_failed)

    return True, result_descriptions, hosted_stats


# ============================================================
# Wall parameter pre-collection (before delete)
# ============================================================
def collect_wall_params(wall):
    """Collect writable instance parameters from wall before deletion."""
    skip_params = set([
        int(BuiltInParameter.WALL_BASE_CONSTRAINT),
        int(BuiltInParameter.WALL_HEIGHT_TYPE),
        int(BuiltInParameter.WALL_BASE_OFFSET),
        int(BuiltInParameter.WALL_TOP_OFFSET),
        int(BuiltInParameter.WALL_USER_HEIGHT_PARAM),
        int(BuiltInParameter.WALL_KEY_REF_PARAM),
        int(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM),
        int(BuiltInParameter.ELEM_FAMILY_PARAM),
        int(BuiltInParameter.ELEM_TYPE_PARAM),
        int(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT),
    ])

    param_data = []
    for param in wall.Parameters:
        if param.IsReadOnly:
            continue
        if param.Definition is None:
            continue
        try:
            bip = param.Definition.BuiltInParameter
            if int(bip) in skip_params:
                continue
        except:
            pass

        storage = param.StorageType
        val = None
        if storage == StorageType.Double:
            val = ('double', param.AsDouble())
        elif storage == StorageType.Integer:
            val = ('int', param.AsInteger())
        elif storage == StorageType.String:
            s = param.AsString()
            if s is not None:
                val = ('string', s)
        elif storage == StorageType.ElementId:
            val = ('id', _eid_int(param.AsElementId()))

        if val is not None:
            param_data.append((param.Definition.Name, val))

    return param_data


def apply_wall_params(target_wall, param_data):
    """Apply previously collected parameters to a wall."""
    for param_name, (ptype, pval) in param_data:
        try:
            p = target_wall.LookupParameter(param_name)
            if p is None or p.IsReadOnly:
                continue
            if ptype == 'double':
                p.Set(pval)
            elif ptype == 'int':
                p.Set(pval)
            elif ptype == 'string':
                p.Set(pval)
            elif ptype == 'id':
                p.Set(_make_eid(pval))
        except:
            continue


def split_wall_at_levels_v2(wall, split_levels, do_copy_params, all_levels):
    """
    Split wall - collects ALL data before deletion, then recreates.
    Returns (success, result_descriptions, hosted_stats)
    """
    info = get_wall_info(wall)

    location = wall.Location
    if not isinstance(location, LocationCurve):
        return False, "Wall does not have a curve location", ""

    curve = location.Curve
    wall_type_id = wall.GetTypeId()

    base_level = info['base_level']
    base_offset = info['base_offset']
    actual_base_elev = info['actual_base_elev']
    actual_top_elev = info['actual_top_elev']
    has_top_level = info['has_top_level']

    actual_min = min(actual_base_elev, actual_top_elev)
    actual_max = max(actual_base_elev, actual_top_elev)

    tolerance = 0.001

    valid_splits = []
    for lv in split_levels:
        lv_elev = lv.Elevation
        if actual_min + tolerance < lv_elev < actual_max - tolerance:
            valid_splits.append(lv)

    if not valid_splits:
        return False, "No valid split levels within wall range", ""

    valid_splits.sort(key=lambda x: x.Elevation)

    # === COLLECT EVERYTHING BEFORE DELETE ===
    hosted_elements = collect_hosted_elements(wall)
    wall_params = collect_wall_params(wall) if do_copy_params else []
    original_flipped = wall.Flipped

    is_structural = False
    struct_param = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT)
    if struct_param:
        is_structural = struct_param.AsInteger() == 1

    # === BUILD SEGMENTS ===
    segments = []

    if has_top_level:
        top_level = info['top_level']
        top_offset = info['top_offset']

        segments.append({
            'base_level': base_level, 'base_offset': base_offset,
            'top_level': valid_splits[0], 'top_offset': 0.0,
            'use_top_constraint': True,
        })
        for i in range(len(valid_splits) - 1):
            segments.append({
                'base_level': valid_splits[i], 'base_offset': 0.0,
                'top_level': valid_splits[i + 1], 'top_offset': 0.0,
                'use_top_constraint': True,
            })
        segments.append({
            'base_level': valid_splits[-1], 'base_offset': 0.0,
            'top_level': top_level, 'top_offset': top_offset,
            'use_top_constraint': True,
        })
    else:
        segments.append({
            'base_level': base_level, 'base_offset': base_offset,
            'top_level': valid_splits[0], 'top_offset': 0.0,
            'use_top_constraint': True,
        })
        for i in range(len(valid_splits) - 1):
            segments.append({
                'base_level': valid_splits[i], 'base_offset': 0.0,
                'top_level': valid_splits[i + 1], 'top_offset': 0.0,
                'use_top_constraint': True,
            })
        last_split_elev = valid_splits[-1].Elevation
        remaining_height = actual_top_elev - last_split_elev
        segments.append({
            'base_level': valid_splits[-1], 'base_offset': 0.0,
            'top_level': None, 'top_offset': 0.0,
            'use_top_constraint': False,
            'unconnected_height': remaining_height,
        })

    # Segment elevation ranges
    segment_ranges = []
    for seg in segments:
        seg_base_elev = seg['base_level'].Elevation + seg['base_offset']
        if seg['use_top_constraint']:
            seg_top_elev = seg['top_level'].Elevation + seg['top_offset']
        else:
            seg_top_elev = seg_base_elev + seg.get('unconnected_height', 10.0)
        segment_ranges.append((seg_base_elev, seg_top_elev))

    # === DELETE ORIGINAL ===
    doc.Delete(wall.Id)
    doc.Regenerate()

    # === CREATE NEW WALLS ===
    new_walls = []
    result_descriptions = []

    for i, seg in enumerate(segments):
        seg_base_level = seg['base_level']
        seg_base_offset = seg['base_offset']

        new_wall = Wall.Create(
            doc, curve, wall_type_id, seg_base_level.Id,
            10.0, 0.0, False, is_structural
        )

        new_wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(seg_base_offset)

        if seg['use_top_constraint']:
            seg_top_level = seg['top_level']
            seg_top_offset = seg['top_offset']
            new_wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(seg_top_level.Id)
            new_wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).Set(seg_top_offset)

            seg_desc = "Wall {}: {}".format(i + 1, seg_base_level.Name)
            if seg_base_offset != 0:
                seg_desc += " {:+.0f}mm".format(seg_base_offset * 304.8)
            seg_desc += " -> {}".format(seg_top_level.Name)
            if seg_top_offset != 0:
                seg_desc += " {:+.0f}mm".format(seg_top_offset * 304.8)
        else:
            uc_height = seg.get('unconnected_height', 10.0)
            new_wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(ElementId.InvalidElementId)
            new_wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(uc_height)

            seg_desc = "Wall {}: {}".format(i + 1, seg_base_level.Name)
            if seg_base_offset != 0:
                seg_desc += " {:+.0f}mm".format(seg_base_offset * 304.8)
            seg_desc += " -> Unconnected {:.0f}mm".format(uc_height * 304.8)

        if original_flipped:
            new_wall.Flip()

        # Apply collected instance parameters
        if do_copy_params and wall_params:
            apply_wall_params(new_wall, wall_params)

        new_walls.append(new_wall)
        result_descriptions.append(seg_desc)

    # Regenerate before hosted element placement
    doc.Regenerate()

    # === RECREATE HOSTED ELEMENTS ===
    hosted_restored = 0
    hosted_failed = 0
    hosted_failed_names = []

    for h_info in hosted_elements:
        seg_idx = find_host_segment(h_info, new_walls, segment_ranges)
        host_wall = new_walls[seg_idx]

        new_fi = recreate_hosted_element(h_info, host_wall)
        if new_fi is not None:
            hosted_restored += 1
        else:
            hosted_failed += 1
            hosted_failed_names.append(h_info['category_name'])

    hosted_stats = ""
    if hosted_elements:
        hosted_stats = "Hosted elements: {} restored".format(hosted_restored)
        if hosted_failed > 0:
            hosted_stats += ", {} failed ({})".format(
                hosted_failed, ", ".join(set(hosted_failed_names)))

    return True, result_descriptions, hosted_stats


# ============================================================
# Main
# ============================================================
def main():
    try:
        selection = uidoc.Selection
        selected_ids = selection.GetElementIds()

        if selected_ids.Count == 0:
            forms.alert("Please select at least one wall!", exitscript=True)
            return

        walls = []
        for elem_id in selected_ids:
            elem = doc.GetElement(elem_id)
            if isinstance(elem, Wall) and not elem.IsStackedWall:
                walls.append(elem)

        if not walls:
            forms.alert(
                "No valid walls found in selection!\n\n"
                "Note: Stacked walls and curtain walls are not supported.",
                exitscript=True)
            return

    except Exception as ex:
        forms.alert("Error: {}".format(str(ex)), exitscript=True)
        return

    # Gather wall info
    wall_info_list = []
    for w in walls:
        wi = get_wall_info(w)
        wall_info_list.append(wi)

    if not wall_info_list:
        forms.alert("No valid walls to process.", exitscript=True)
        return

    # Show form
    all_levels = get_all_levels()
    form = LevelSelectionForm(all_levels, wall_info_list)
    result = form.ShowDialog()

    if result != DialogResult.OK:
        return

    selected_levels = form.selected_levels
    do_copy_params = form.copy_params

    if not selected_levels:
        forms.alert("No levels selected!", exitscript=True)
        return

    # Execute
    t = Transaction(doc, "DQT - Split Walls at Levels")
    t.Start()

    try:
        success_count = 0
        fail_messages = []
        all_results = []
        all_hosted_stats = []

        for wi in wall_info_list:
            wall = doc.GetElement(_make_eid(wi['element_id']))
            if wall is None:
                fail_messages.append(
                    "Wall ID {} not found".format(wi['element_id']))
                continue

            success, result_data, hosted_stats = split_wall_at_levels_v2(
                wall, selected_levels, do_copy_params, all_levels)

            if success:
                success_count += 1
                all_results.extend(result_data)
                if hosted_stats:
                    all_hosted_stats.append(hosted_stats)
            else:
                fail_messages.append(
                    "Wall ID {}: {}".format(wi['element_id'], result_data))

        if success_count > 0:
            t.Commit()
            message = "{} wall(s) split successfully!\n\n".format(success_count)
            for r in all_results:
                message += r + "\n"
            if all_hosted_stats:
                message += "\n"
                for hs in all_hosted_stats:
                    message += hs + "\n"
            if fail_messages:
                message += "\nSkipped:\n"
                for fm in fail_messages:
                    message += fm + "\n"
            forms.alert(message)
        else:
            t.RollBack()
            message = "No walls were split.\n\n"
            for fm in fail_messages:
                message += fm + "\n"
            forms.alert(message)

    except Exception as e:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert("Error: {}".format(str(e)), exitscript=True)
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()