# -*- coding: utf-8 -*-
"""
Create Room Plan

Create Plan Views from Room List with WPF UI.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Create Plan Views"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
import os
import sys
import clr
import re

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import WindowState, Visibility, Thickness, CornerRadius, HorizontalAlignment, VerticalAlignment, FontWeights
from System.Windows.Controls import Border, Grid, TextBlock, Canvas
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Shapes import Line

from rpw import revit, DB
from Autodesk.Revit.DB import (
    Transaction,
    TransactionGroup,
    View,
    FilteredElementCollector,
    BuiltInCategory,
    ViewType,
    ViewFamilyType,
    ViewFamily,
    ViewPlan,
    ViewSheet,
    Viewport,
    ElevationMarker,
    SpatialElementBoundaryOptions,
    SpatialElementBoundaryLocation,
    ElementTransformUtils,
    XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms, script

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================
logger        = script.get_logger()
output        = script.get_output()
uidoc         = revit.uidoc
doc           = revit.doc
REVIT_VERSION = int(revit.doc.Application.VersionNumber)

SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR    = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
lib_dir    = os.path.join(EXT_DIR, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

XAML_FILE  = os.path.join(EXT_DIR, 'lib', 'GUI', 'Tools', 'CreateRoomPlan.xaml')


# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
# ==================================================

class RoomItem(object):
    """Represents a room item in the DataGrid."""
    def __init__(self, room_element, existing_plan_names=None):
        self.Element = room_element
        self.IsSelected = False
        p_num = room_element.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER)
        self.Number = p_num.AsString() if p_num else ""
        p_name = room_element.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
        self.Name = p_name.AsString() if p_name else ""

        type_param = room_element.LookupParameter("Room Type")
        self.RoomType = type_param.AsString() if type_param else ""

        try:
            level = doc.GetElement(room_element.LevelId)
            self.Level = level.Name if level else ""
        except Exception:
            self.Level = ""

        # Count floor plans matching this room number suffix (#Number)
        search_suffix = "(#{})".format(self.Number)
        if existing_plan_names:
            self.FloorPlanCount = sum(1 for name in existing_plan_names if search_suffix in name)
        else:
            self.FloorPlanCount = 0

        # Quantity to generate
        self.GenQty = 1


class CreateRoomPlanWindow(forms.WPFWindow):
    """WPF window for creating plan views from rooms."""

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._all_rooms = []
        self._first_sheet = None
        self._generated_sheets = []
        self._load_rooms()
        self._load_view_templates()
        self._load_plan_type_options()
        self._load_title_blocks()
        self._update_status()
        self._update_mockup()

    # ── Data loading ──────────────────────────────────
    def _load_rooms(self):
        """Collect all placed rooms from the document."""
        room_elements = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_Rooms) \
            .ToElements()

        # Collect names of all existing floor plans in the document
        view_elements = FilteredElementCollector(doc) \
            .OfClass(ViewPlan) \
            .WhereElementIsNotElementType() \
            .ToElements()
        
        existing_plan_names = [
            v.Name for v in view_elements
            if not v.IsTemplate and v.ViewType == ViewType.FloorPlan
        ]

        self._all_rooms = []
        for r in room_elements:
            # Skip unplaced rooms (area == 0 or no location)
            try:
                if r.Location is None:
                    continue
            except Exception:
                continue
            self._all_rooms.append(RoomItem(r, existing_plan_names))

        # Sort by number
        self._all_rooms.sort(key=lambda x: x.Number)
        self.room_datagrid.ItemsSource = self._all_rooms
        if self._all_rooms:
            self.room_datagrid.SelectedItem = self._all_rooms[0]

    def _load_view_templates(self):
        """Collect view templates for Floor Plan and Ceiling Plan."""
        view_elements = FilteredElementCollector(doc) \
            .OfClass(View) \
            .WhereElementIsNotElementType() \
            .ToElements()

        plan_templates = sorted([
            v.Name for v in view_elements
            if v.IsTemplate and v.ViewType == ViewType.FloorPlan
        ])
        rcp_templates = sorted([
            v.Name for v in view_elements
            if v.IsTemplate and v.ViewType == ViewType.CeilingPlan
        ])

        # Store template elements for later lookup
        self._plan_template_map = {
            v.Name: v.Id for v in view_elements
            if v.IsTemplate and v.ViewType == ViewType.FloorPlan
        }
        self._rcp_template_map = {
            v.Name: v.Id for v in view_elements
            if v.IsTemplate and v.ViewType == ViewType.CeilingPlan
        }

        # Populate combo boxes with "None" option
        self.cmb_plan_template.Items.Add("<None>")
        for name in plan_templates:
            self.cmb_plan_template.Items.Add(name)
        self.cmb_plan_template.SelectedIndex = 0

        self.cmb_rcp_template.Items.Add("<None>")
        for name in rcp_templates:
            self.cmb_rcp_template.Items.Add(name)
        self.cmb_rcp_template.SelectedIndex = 0

        # Elevation templates
        elev_templates = sorted([
            v.Name for v in view_elements
            if v.IsTemplate and v.ViewType == ViewType.Elevation
        ])
        self._elev_template_map = {
            v.Name: v.Id for v in view_elements
            if v.IsTemplate and v.ViewType == ViewType.Elevation
        }
        self.cmb_elev_template.Items.Add("<None>")
        for name in elev_templates:
            self.cmb_elev_template.Items.Add(name)
        self.cmb_elev_template.SelectedIndex = 0

    def _load_plan_type_options(self):
        """Load available view family types for Floor Plan / Ceiling Plan."""
        view_types = FilteredElementCollector(doc) \
            .OfClass(ViewFamilyType) \
            .WhereElementIsElementType() \
            .ToElements()

        self._floor_plan_type_id = None
        self._ceiling_plan_type_id = None
        self._elevation_type_id = None

        for vt in view_types:
            if vt.FamilyName == 'Floor Plan' and self._floor_plan_type_id is None:
                self._floor_plan_type_id = vt.Id
            elif vt.FamilyName == 'Ceiling Plan' and self._ceiling_plan_type_id is None:
                self._ceiling_plan_type_id = vt.Id
            elif vt.ViewFamily == ViewFamily.Elevation and self._elevation_type_id is None:
                self._elevation_type_id = vt.Id

    def _load_title_blocks(self):
        """Load all title block types into cmb_titleblock."""
        tb_types = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_TitleBlocks) \
            .WhereElementIsElementType() \
            .ToElements()

        self._titleblock_map = {}  # display_name -> ElementId
        for tb in tb_types:
            try:
                fam_name = tb.Family.Name
                type_name = tb.get_Parameter(
                    DB.BuiltInParameter.SYMBOL_NAME_PARAM
                )
                type_name = type_name.AsString() if type_name else ""
                display = "{}: {}".format(fam_name, type_name) if type_name else fam_name
            except Exception:
                display = str(tb.Id)
            self._titleblock_map[display] = tb.Id

        for name in sorted(self._titleblock_map.keys()):
            self.cmb_titleblock.Items.Add(name)
        if self._titleblock_map:
            self.cmb_titleblock.SelectedIndex = 0
        else:
            self.chk_layout_on_sheet.IsEnabled = False
            self.chk_layout_on_sheet.ToolTip = "No title blocks found in project"

    # ── Helpers ───────────────────────────────────────
    def _get_selected_rooms(self):
        """Return list of RoomItems that are checked."""
        return [r for r in self._all_rooms if r.IsSelected]

    def _update_status(self):
        """Update status bar text."""
        selected = len(self._get_selected_rooms())
        total = len(self._all_rooms)
        self.status_count.Text = "{} rooms".format(total)
        if selected > 0:
            self.status_text.Text = "{} room(s) selected".format(selected)
        else:
            self.status_text.Text = "Ready"

    def _get_offset(self):
        """Parse offset value from text box (in meters) and convert to feet."""
        try:
            val_meters = float(self.txt_offset.Text)
            return val_meters * 3.28084
        except (ValueError, TypeError):
            return 3.28084  # Default 1 meter in feet

    @staticmethod
    def _offset_bbox(bbox, offset=1):
        """Expand bounding box by offset in all directions."""
        new_bbox = DB.BoundingBoxXYZ()
        new_bbox.Min = DB.XYZ(bbox.Min.X - offset, bbox.Min.Y - offset, bbox.Min.Z - offset)
        new_bbox.Max = DB.XYZ(bbox.Max.X + offset, bbox.Max.Y + offset, bbox.Max.Z + offset)
        return new_bbox

    def _build_view_name(self, room_item, copy_index=0):
        """Build the plan view name from room info."""
        if "UNIT" in room_item.Name.upper() and room_item.RoomType:
            base = "ENLARGED PLAN - TYPE {} ({})".format(room_item.RoomType, room_item.Name)
        else:
            base = "ENLARGED PLAN - {} - (#{})".format(room_item.Name, room_item.Number)
        
        if copy_index > 0:
            return "{} - Copy {}".format(base, copy_index)
        return base

    def _find_plan_view_for_level(self, level_id):
        """Find an existing floor plan view for the given level."""
        views = FilteredElementCollector(doc) \
            .OfClass(ViewPlan) \
            .WhereElementIsNotElementType() \
            .ToElements()
        for v in views:
            if (not v.IsTemplate
                    and v.ViewType == ViewType.FloorPlan
                    and v.GenLevel is not None
                    and v.GenLevel.Id == level_id):
                return v
        return None

    def _get_boundary_wall_ids(self, room):
        """Return set of wall element ids forming the room boundary."""
        wall_ids = set()
        try:
            opt = SpatialElementBoundaryOptions()
            opt.SpatialElementBoundaryLocation = \
                SpatialElementBoundaryLocation.Finish
            segments_list = room.GetBoundarySegments(opt)
            if segments_list:
                for seg_loop in segments_list:
                    for seg in seg_loop:
                        elem = doc.GetElement(seg.ElementId)
                        if elem and isinstance(elem, DB.Wall):
                            wall_ids.add(seg.ElementId)
        except Exception:
            pass
        return wall_ids

    # ── Window chrome handlers ────────────────────────
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

    def nav_toggle_clicked(self, sender, e):
        """Handle sidebar toggle button clicks mutually exclusively, switch tabs, and update title subtitle."""
        try:
            # Set clicked button to checked, and all others to unchecked
            is_rooms = (sender.Name == "nav_rooms")
            self.nav_rooms.IsChecked = is_rooms
            self.nav_layout.IsChecked = not is_rooms

            # Switch TabControl index
            if is_rooms:
                self.main_tab_control.SelectedIndex = 0
                self.lbl_subtitle.Text = "Generate plan views from selected rooms"
            else:
                self.main_tab_control.SelectedIndex = 1
                self.lbl_subtitle.Text = "Configure sheet layout and preview viewport placement"
                self._update_mockup()
        except Exception as ex:
            logger.error("Navigation click error: {}".format(ex))

    # ── Layout / Preview handlers ─────────────────────
    def layout_toggle_changed(self, sender, e):
        """Enable/disable layout sub-options based on toggle."""
        enabled = bool(self.chk_layout_on_sheet.IsChecked)
        self.pnl_layout_options.IsEnabled = enabled
        self.pnl_layout_options.Opacity = 1.0 if enabled else 0.45

    def room_selection_changed(self, sender, e):
        """Handle selected room change to update the realtime preview."""
        self._update_status()
        self._update_mockup()

    def mockup_setting_changed(self, sender, e):
        """Handle view checkbox or ComboBox changes to redraw mockup."""
        self._update_mockup()

    def offset_changed(self, sender, e):
        """Handle offset textbox text changes to redraw mockup."""
        self._update_mockup()

    def _get_preview_room(self):
        """Get the room item to use for the layout preview."""
        try:
            row = self.room_datagrid.SelectedItem
            if row:
                return row
        except Exception:
            pass
        
        checked_rooms = self._get_selected_rooms()
        if checked_rooms:
            return checked_rooms[0]
            
        if self._all_rooms:
            return self._all_rooms[0]
            
        return None

    def _get_room_dimensions(self, room_item):
        """Get room bounding box width and depth in feet, with a fallback."""
        if room_item is None or room_item.Element is None:
            return 15.0, 12.0
        try:
            active_view = doc.ActiveView
            bbox = room_item.Element.get_BoundingBox(active_view)
            if bbox:
                w = bbox.Max.X - bbox.Min.X
                h = bbox.Max.Y - bbox.Min.Y
                if w > 0 and h > 0:
                    return w, h
        except Exception:
            pass
        return 15.0, 12.0

    def _get_template_scale(self, template_name):
        """Get the scale factor of a view template by name, defaulting to active view scale."""
        if not template_name or template_name == "<None>":
            try:
                return doc.ActiveView.Scale
            except Exception:
                return 50
        
        t_id = None
        if hasattr(self, '_plan_template_map') and template_name in self._plan_template_map:
            t_id = self._plan_template_map[template_name]
        elif hasattr(self, '_rcp_template_map') and template_name in self._rcp_template_map:
            t_id = self._rcp_template_map[template_name]
        elif hasattr(self, '_elev_template_map') and template_name in self._elev_template_map:
            t_id = self._elev_template_map[template_name]
            
        if t_id:
            try:
                template_view = doc.GetElement(t_id)
                if template_view:
                    return template_view.Scale
            except Exception:
                pass
        
        try:
            return doc.ActiveView.Scale
        except Exception:
            return 50

    def _get_preview_sheet_size(self):
        """Estimate sheet dimensions (width, height) in feet from the selected title block."""
        if not hasattr(self, 'cmb_titleblock') or self.cmb_titleblock.SelectedItem is None:
            return 2.759, 1.949
        tb_name = str(self.cmb_titleblock.SelectedItem).upper()
        if "A0" in tb_name:
            return 3.901, 2.759
        elif "A2" in tb_name:
            return 1.949, 1.378
        elif "A3" in tb_name:
            return 1.378, 0.974
        elif "A4" in tb_name:
            return 0.974, 0.689
        else: # Default A1: 841x594 mm = 2.759 x 1.949 ft
            return 2.759, 1.949

    def _draw_viewport(self, canvas, title, detail_num, w_px, h_px, x_px, y_px, bg_color, border_color):
        """Draw a viewport rectangle and Revit-style title mark on the WPF canvas."""
        # 1. Viewport Box
        box = Border()
        box.Width = w_px
        box.Height = h_px
        box.Background = SolidColorBrush(bg_color)
        box.BorderBrush = SolidColorBrush(border_color)
        box.BorderThickness = Thickness(1.5)
        box.CornerRadius = CornerRadius(4)
        Canvas.SetLeft(box, x_px)
        Canvas.SetTop(box, y_px)
        canvas.Children.Add(box)

        # 2. Revit Title Line
        line_y = y_px + h_px + 6
        line = Line()
        line.X1 = x_px
        line.Y1 = line_y
        line.X2 = x_px + max(w_px * 0.7, 40.0)
        line.Y2 = line_y
        line.Stroke = SolidColorBrush(Color.FromRgb(0x71, 0x71, 0x7A))
        line.StrokeThickness = 1
        canvas.Children.Add(line)

        # 3. Detail Circle
        circle = Border()
        circle.Width = 14
        circle.Height = 14
        circle.CornerRadius = CornerRadius(7)
        circle.BorderBrush = SolidColorBrush(border_color)
        circle.BorderThickness = Thickness(1)
        circle.Background = SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF))
        
        circle_text = TextBlock()
        circle_text.Text = str(detail_num)
        circle_text.FontSize = 8
        circle_text.FontWeight = FontWeights.Bold
        circle_text.Foreground = SolidColorBrush(border_color)
        circle_text.HorizontalAlignment = HorizontalAlignment.Center
        circle_text.VerticalAlignment = VerticalAlignment.Center
        circle.Child = circle_text
        
        Canvas.SetLeft(circle, x_px)
        Canvas.SetTop(circle, line_y + 3)
        canvas.Children.Add(circle)

        # 4. View Title Text
        title_text = TextBlock()
        title_text.Text = title
        title_text.FontSize = 8.5
        title_text.FontWeight = FontWeights.Bold
        title_text.Foreground = SolidColorBrush(Color.FromRgb(0x18, 0x18, 0x1B))
        Canvas.SetLeft(title_text, x_px + 18)
        Canvas.SetTop(title_text, line_y + 2)
        canvas.Children.Add(title_text)

    def _draw_viewport_to_canvas(self, canvas, title, detail_num, w_paper, h_paper, cx_sheet, cy_sheet, w_sheet, h_sheet, mockup_w, mockup_h, left_margin, top_margin, bg_color, border_color):
        """Map sheet coordinates to canvas coordinates and draw the viewport."""
        x_sheet_left = cx_sheet - w_paper / 2.0
        y_sheet_top = cy_sheet + h_paper / 2.0
        
        x_px = x_sheet_left * (float(mockup_w) / w_sheet) - left_margin
        y_px = float(mockup_h) - y_sheet_top * (float(mockup_h) / h_sheet) - top_margin
        
        w_px = w_paper * (float(mockup_w) / w_sheet)
        h_px = h_paper * (float(mockup_h) / h_sheet)
        
        self._draw_viewport(canvas, title, detail_num, w_px, h_px, x_px, y_px, bg_color, border_color)

    def _update_mockup(self):
        """Update the real-time layout mockup based on current settings and selected room."""
        if not hasattr(self, "combined_canvas") or self.combined_canvas is None:
            return
        
        # 1. Clear previous drawings
        self.combined_canvas.Children.Clear()
        self.plans_canvas.Children.Clear()
        self.elevations_canvas.Children.Clear()

        # 2. Get the room to preview
        room_item = self._get_preview_room()
        if not room_item:
            return

        w_room, h_room = self._get_room_dimensions(room_item)
        offset = self._get_offset()

        # 3. Check enabled view types and lookup templates
        do_floor = bool(self.chk_floor_plan.IsChecked)
        do_ceiling = bool(self.chk_ceiling_plan.IsChecked)
        do_elevations = bool(self.chk_elevations.IsChecked)

        plan_template = str(self.cmb_plan_template.SelectedItem) if self.cmb_plan_template.SelectedItem else "<None>"
        rcp_template = str(self.cmb_rcp_template.SelectedItem) if self.cmb_rcp_template.SelectedItem else "<None>"
        elev_template = str(self.cmb_elev_template.SelectedItem) if self.cmb_elev_template.SelectedItem else "<None>"

        plan_scale = self._get_template_scale(plan_template)
        rcp_scale = self._get_template_scale(rcp_template)
        elev_scale = self._get_template_scale(elev_template)

        # 4. Viewport sizes on paper (feet)
        w_floor_paper = (w_room + 2.0 * offset) / plan_scale
        h_floor_paper = (h_room + 2.0 * offset) / plan_scale

        w_rcp_paper = (w_room + 2.0 * offset) / rcp_scale
        h_rcp_paper = (h_room + 2.0 * offset) / rcp_scale

        h_elev_paper = 10.0 / elev_scale
        w_elev_s_n_paper = (w_room + 2.0 * offset) / elev_scale
        w_elev_w_e_paper = (h_room + 2.0 * offset) / elev_scale

        # 5. Sheet printable dimension (in feet)
        w_sheet, h_sheet = self._get_preview_sheet_size()

        # Usable boundaries inside sheet space
        margin = 0.066
        tb_strip = 0.164
        usable_h = h_sheet - margin - tb_strip
        baseline = tb_strip + margin / 2.0

        # Colors
        color_plan_bg = Color.FromRgb(0xEF, 0xF6, 0xFF)
        color_plan_border = Color.FromRgb(0x25, 0x63, 0xEB)

        color_rcp_bg = Color.FromRgb(0xF5, 0xF3, 0xFF)
        color_rcp_border = Color.FromRgb(0x7C, 0x3A, 0xED)

        color_elev_bg = Color.FromRgb(0xFF, 0xFB, 0xEB)
        color_elev_border = Color.FromRgb(0xD9, 0x77, 0x06)

        # 6. Draw Combined Canvas
        if bool(self.rdo_layout_combined.IsChecked):
            plan_zone_w = w_sheet * 0.45
            plan_views = []
            if do_floor:
                plan_views.append(("Floor Plan", w_floor_paper, h_floor_paper, color_plan_bg, color_plan_border))
            if do_ceiling:
                plan_views.append(("Ceiling Plan", w_rcp_paper, h_rcp_paper, color_rcp_bg, color_rcp_border))

            # Draw plans
            if len(plan_views) == 1:
                title, wp, hp, bg, border = plan_views[0]
                cx, cy = plan_zone_w / 2.0 + margin, baseline + usable_h / 2.0
                self._draw_viewport_to_canvas(self.combined_canvas, title, "1", wp, hp, cx, cy, w_sheet, h_sheet, 580, 380, 16, 16, bg, border)
            elif len(plan_views) >= 2:
                # Floor Plan
                title, wp, hp, bg, border = plan_views[0]
                cx, cy = plan_zone_w / 2.0 + margin, baseline + usable_h * 0.72
                self._draw_viewport_to_canvas(self.combined_canvas, title, "1", wp, hp, cx, cy, w_sheet, h_sheet, 580, 380, 16, 16, bg, border)
                # Ceiling Plan
                title, wp, hp, bg, border = plan_views[1]
                cx, cy = plan_zone_w / 2.0 + margin, baseline + usable_h * 0.28
                self._draw_viewport_to_canvas(self.combined_canvas, title, "2", wp, hp, cx, cy, w_sheet, h_sheet, 580, 380, 16, 16, bg, border)

            # Draw elevations on right half (2x2 grid)
            if do_elevations:
                right_x = w_sheet * 0.5
                elev_zone_w = w_sheet - right_x - margin
                elev_zone_h = usable_h
                positions = [
                    ("Elevation South", w_elev_s_n_paper, h_elev_paper, right_x + elev_zone_w * 0.25, baseline + elev_zone_h * 0.75, "3"),
                    ("Elevation West", w_elev_w_e_paper, h_elev_paper, right_x + elev_zone_w * 0.75, baseline + elev_zone_h * 0.75, "4"),
                    ("Elevation North", w_elev_s_n_paper, h_elev_paper, right_x + elev_zone_w * 0.25, baseline + elev_zone_h * 0.25, "5"),
                    ("Elevation East", w_elev_w_e_paper, h_elev_paper, right_x + elev_zone_w * 0.75, baseline + elev_zone_h * 0.25, "6"),
                ]
                for title, wp, hp, cx, cy, num in positions:
                    self._draw_viewport_to_canvas(self.combined_canvas, title, num, wp, hp, cx, cy, w_sheet, h_sheet, 580, 380, 16, 16, color_elev_bg, color_elev_border)

        # 7. Draw Separate Canvas (Plans and Elevations on separate sheets)
        else:
            # Sheet 1: Plans
            plan_views = []
            if do_floor:
                plan_views.append(("Floor Plan", w_floor_paper, h_floor_paper, color_plan_bg, color_plan_border))
            if do_ceiling:
                plan_views.append(("Ceiling Plan", w_rcp_paper, h_rcp_paper, color_rcp_bg, color_rcp_border))

            if len(plan_views) == 1:
                title, wp, hp, bg, border = plan_views[0]
                cx, cy = w_sheet / 2.0, baseline + usable_h / 2.0
                self._draw_viewport_to_canvas(self.plans_canvas, title, "1", wp, hp, cx, cy, w_sheet, h_sheet, 300, 200, 8, 8, bg, border)
            elif len(plan_views) >= 2:
                # Floor Plan
                title, wp, hp, bg, border = plan_views[0]
                cx, cy = w_sheet / 2.0, baseline + usable_h * 0.70
                self._draw_viewport_to_canvas(self.plans_canvas, title, "1", wp, hp, cx, cy, w_sheet, h_sheet, 300, 200, 8, 8, bg, border)
                # Ceiling Plan
                title, wp, hp, bg, border = plan_views[1]
                cx, cy = w_sheet / 2.0, baseline + usable_h * 0.28
                self._draw_viewport_to_canvas(self.plans_canvas, title, "2", wp, hp, cx, cy, w_sheet, h_sheet, 300, 200, 8, 8, bg, border)

            # Sheet 2: Elevations (2x2 grid)
            if do_elevations:
                positions = [
                    ("Elevation South", w_elev_s_n_paper, h_elev_paper, w_sheet * 0.25, baseline + usable_h * 0.75, "1"),
                    ("Elevation West", w_elev_w_e_paper, h_elev_paper, w_sheet * 0.75, baseline + usable_h * 0.75, "2"),
                    ("Elevation North", w_elev_s_n_paper, h_elev_paper, w_sheet * 0.25, baseline + usable_h * 0.25, "3"),
                    ("Elevation East", w_elev_w_e_paper, h_elev_paper, w_sheet * 0.75, baseline + usable_h * 0.25, "4"),
                ]
                for title, wp, hp, cx, cy, num in positions:
                    self._draw_viewport_to_canvas(self.elevations_canvas, title, num, wp, hp, cx, cy, w_sheet, h_sheet, 300, 200, 8, 8, color_elev_bg, color_elev_border)

    def open_sheet_clicked(self, sender, e):
        """Activate the selected generated sheet in Revit."""
        try:
            selected_idx = self.cmb_generated_sheets.SelectedIndex
            if selected_idx >= 0 and selected_idx < len(self._generated_sheets):
                selected_sheet = self._generated_sheets[selected_idx]
                uidoc.ActiveView = selected_sheet
            else:
                TaskDialog.Show("Create Room Plan", "Please select a valid generated sheet from the list.")
        except Exception as ex:
            logger.error("Open sheet error: {}".format(ex))

    # ── Toolbar handlers ──────────────────────────────
    def select_all_clicked(self, sender, e):
        for r in self._all_rooms:
            r.IsSelected = True
        self.room_datagrid.Items.Refresh()
        self._update_status()

    def select_none_clicked(self, sender, e):
        for r in self._all_rooms:
            r.IsSelected = False
        self.room_datagrid.Items.Refresh()
        self._update_status()

    def search_changed(self, sender, e):
        """Filter room list by search text."""
        query = self.txt_search.Text.strip().upper()
        if not query:
            self.room_datagrid.ItemsSource = self._all_rooms
        else:
            filtered = [
                r for r in self._all_rooms
                if query in r.Name.upper()
                or query in r.Number.upper()
                or query in (r.RoomType or "").upper()
                or query in (r.Level or "").upper()
            ]
            self.room_datagrid.ItemsSource = filtered
        self._update_status()

    # ── Main action ───────────────────────────────────
    def create_plans_clicked(self, sender, e):
        """Create plan views (and optionally lay them out on sheets) for selected rooms."""
        selected_rooms = self._get_selected_rooms()
        if not selected_rooms:
            TaskDialog.Show("Create Room Plan", "Please select at least one room.")
            return

        do_floor      = bool(self.chk_floor_plan.IsChecked)
        do_ceiling    = bool(self.chk_ceiling_plan.IsChecked)
        do_elevations = bool(self.chk_elevations.IsChecked)
        do_layout     = bool(self.chk_layout_on_sheet.IsChecked)

        if not do_floor and not do_ceiling and not do_elevations:
            TaskDialog.Show("Create Room Plan", "Please select at least one view type.")
            return

        # Template selections
        plan_template_name = self.cmb_plan_template.SelectedItem
        rcp_template_name  = self.cmb_rcp_template.SelectedItem
        elev_template_name = self.cmb_elev_template.SelectedItem
        plan_template_id = self._plan_template_map.get(plan_template_name) \
            if plan_template_name != "<None>" else None
        rcp_template_id  = self._rcp_template_map.get(rcp_template_name) \
            if rcp_template_name != "<None>" else None
        elev_template_id = self._elev_template_map.get(elev_template_name) \
            if elev_template_name != "<None>" else None

        offset         = self._get_offset()
        cropbox_visible = bool(self.chk_cropbox_visible.IsChecked)
        created_count  = 0
        error_count    = 0
        active_view    = doc.ActiveView

        # Collect per-room results for sheet layout
        room_results = []  # list of dicts

        for room_item in selected_rooms:
            room          = room_item.Element
            room_level_id = room.LevelId
            room_bbox     = room.get_BoundingBox(active_view)
            if room_bbox is None:
                error_count += 1
                continue

            new_bbox  = self._offset_bbox(room_bbox, offset)

            # Get quantity to generate for floor plans
            try:
                qty = int(room_item.GenQty)
                if qty < 0:
                    qty = 0
            except Exception:
                qty = 1

            result = {
                'room_item':    room_item,
                'floor_plan':   None,
                'ceiling_plan': None,
                'elevations':   [],
            }

            # ── Floor Plan ─────────────────────────────
            if do_floor and self._floor_plan_type_id and qty > 0:
                for idx in range(qty):
                    view_name = self._build_view_name(room_item, copy_index=idx)
                    try:
                        with Transaction(doc, "Create Floor Plan") as t:
                            t.Start()
                            vp = DB.ViewPlan.Create(doc, self._floor_plan_type_id, room_level_id)
                            vp.CropBoxActive  = True
                            vp.CropBoxVisible = cropbox_visible
                            vp.CropBox        = new_bbox
                            vp.Name           = view_name
                            t.Commit()
                        # Use first created floor plan as primary for sheet layout
                        if idx == 0:
                            result['floor_plan'] = vp
                        created_count += 1
                        if plan_template_id:
                            with Transaction(doc, "Assign Floor Plan Template") as t2:
                                t2.Start()
                                doc.GetElement(vp.Id).ViewTemplateId = plan_template_id
                                t2.Commit()
                    except Exception as ex:
                        error_count += 1
                        logger.error("Floor plan error for {}: {}".format(view_name, ex))

            # ── Ceiling Plan ────────────────────────────
            if do_ceiling and self._ceiling_plan_type_id:
                try:
                    with Transaction(doc, "Create Ceiling Plan") as t:
                        t.Start()
                        vp = DB.ViewPlan.Create(doc, self._ceiling_plan_type_id, room_level_id)
                        vp.CropBoxActive  = True
                        vp.CropBoxVisible = cropbox_visible
                        vp.CropBox        = new_bbox
                        vp.Name           = view_name
                        t.Commit()
                    result['ceiling_plan'] = vp
                    created_count += 1
                    if rcp_template_id:
                        with Transaction(doc, "Assign RCP Template") as t2:
                            t2.Start()
                            doc.GetElement(vp.Id).ViewTemplateId = rcp_template_id
                            t2.Commit()
                except Exception as ex:
                    error_count += 1
                    logger.error("Ceiling plan error for {}: {}".format(view_name, ex))

            # ── Interior Elevations ─────────────────────
            if do_elevations and self._elevation_type_id:
                try:
                    host_plan = self._find_plan_view_for_level(room_level_id)
                    if host_plan is None:
                        error_count += 1
                        logger.error("No floor plan found for level to host elevation marker")
                    else:
                        center = room.Location.Point
                        room_width = room_bbox.Max.X - room_bbox.Min.X
                        room_depth = room_bbox.Max.Y - room_bbox.Min.Y
                        max_dim    = max(room_width, room_depth) + offset * 2

                        with Transaction(doc, "Create Interior Elevations") as t:
                            t.Start()
                            scale  = host_plan.Scale
                            marker = ElevationMarker.CreateElevationMarker(
                                doc, self._elevation_type_id, center, scale
                            )
                            directions = ["South", "West", "North", "East"]
                            for idx in range(4):
                                try:
                                    ev = marker.CreateElevation(doc, host_plan.Id, idx)
                                    try:
                                        ev.Name = "INTERIOR ELEV - {} - {} ({})".format(
                                            room_item.Name, directions[idx], room_item.Number
                                        )
                                    except Exception:
                                        pass
                                    ev.CropBoxActive  = True
                                    ev.CropBoxVisible = cropbox_visible
                                    p_far = ev.get_Parameter(
                                        DB.BuiltInParameter.VIEWER_BOUND_FAR_CLIPPING)
                                    if p_far:
                                        p_far.Set(1)
                                    p_off = ev.get_Parameter(
                                        DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
                                    if p_off:
                                        p_off.Set(max_dim / 2 + offset)
                                    if elev_template_id:
                                        ev.ViewTemplateId = elev_template_id
                                    result['elevations'].append(ev)
                                    created_count += 1
                                except Exception as ex:
                                    error_count += 1
                                    logger.error("Elevation {} error: {}".format(
                                        directions[idx], ex))
                            t.Commit()
                except Exception as ex:
                    error_count += 1
                    logger.error("Elevation error for {}: {}".format(view_name, ex))

            room_results.append(result)

        # ── Sheet Layout ────────────────────────────────
        sheets_created = 0
        self._generated_sheets = []
        if do_layout and room_results:
            tb_name = self.cmb_titleblock.SelectedItem
            tb_id   = self._titleblock_map.get(tb_name) if tb_name else None
            if tb_id:
                combined = bool(self.rdo_layout_combined.IsChecked)
                for result in room_results:
                    try:
                        sheets = self._layout_views_on_sheets(
                            result, tb_id, combined
                        )
                        sheets_created += len(sheets)
                        for sh in sheets:
                            self._generated_sheets.append(sh)
                    except Exception as ex:
                        error_count += 1
                        logger.error("Sheet layout error for {}: {}".format(
                            result['room_item'].Name, ex))

                if self._generated_sheets:
                    # Clear and populate the combo box
                    self.cmb_generated_sheets.Items.Clear()
                    for sh in self._generated_sheets:
                        sheet_num = sh.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER)
                        num_str = sheet_num.AsString() if sheet_num else "???"
                        self.cmb_generated_sheets.Items.Add("{} - {}".format(num_str, sh.Name))
                    
                    # Select first sheet and enable controls
                    self.cmb_generated_sheets.SelectedIndex = 0
                    self.cmb_generated_sheets.IsEnabled = True
                    self.btn_open_sheet.IsEnabled = True
            else:
                logger.warning("No title block selected — skipping sheet layout.")

        # ── Result ─────────────────────────────────────
        msg = "{} view(s) created.".format(created_count)
        if sheets_created:
            msg += "\n{} sheet(s) created.".format(sheets_created)
        if error_count > 0:
            msg += "\n{} error(s) — see output for details.".format(error_count)

        self.status_text.Text = msg
        if not do_layout:
            TaskDialog.Show("Create Room Plan", msg)
            self.Close()
        else:
            # Keep dialog open so user can click Open Selected Sheet
            TaskDialog.Show("Create Room Plan", msg)

    # ── Sheet helpers ─────────────────────────────────
    def _get_sheet_size(self, sheet):
        """Return (width, height) of the title block in feet. Falls back to A1."""
        try:
            tb_elems = FilteredElementCollector(doc, sheet.Id) \
                .OfCategory(BuiltInCategory.OST_TitleBlocks) \
                .ToElements()
            if tb_elems:
                bb = tb_elems[0].get_BoundingBox(sheet)
                if bb:
                    w = bb.Max.X - bb.Min.X
                    h = bb.Max.Y - bb.Min.Y
                    if w > 0 and h > 0:
                        return (w, h)
        except Exception:
            pass
        # Default A1: 841x594 mm = 2.759 x 1.949 ft
        return (2.759, 1.949)

    def _build_sheet_number(self, room_item, suffix=""):
        """Build a sheet number like EPL-101-P or EPL-101-E."""
        base = "EPL-{}".format(room_item.Number)
        return "{}-{}".format(base, suffix) if suffix else base

    def _create_sheet(self, room_item, titleblock_id, name_suffix="", num_suffix=""):
        """Create a ViewSheet. Returns the sheet element."""
        sheet_name = self._build_view_name(room_item)
        if name_suffix:
            sheet_name = "{} — {}".format(sheet_name, name_suffix)
        sheet_num = self._build_sheet_number(room_item, num_suffix)

        with Transaction(doc, "Create Sheet") as t:
            t.Start()
            sheet = ViewSheet.Create(doc, titleblock_id)
            sheet.Name = sheet_name
            p_num = sheet.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER)
            if p_num and not p_num.IsReadOnly:
                try:
                    p_num.Set(sheet_num)
                except Exception:
                    pass
            t.Commit()
        return sheet

    def _place_viewport_centered(self, sheet, view_id, cx, cy):
        """
        Place a viewport at (cx, cy) in sheet coordinates.
        Returns the Viewport element, or None on failure.
        """
        try:
            with Transaction(doc, "Place Viewport") as t:
                t.Start()
                vp = Viewport.Create(doc, sheet.Id, view_id, XYZ(cx, cy, 0))
                t.Commit()
            return vp
        except Exception as ex:
            logger.error("Viewport place error: {}".format(ex))
            return None

    def _layout_combined(self, result, sheet):
        """
        Combined layout — plan left, 4 elevations right (2x2 grid).
        """
        w, h = self._get_sheet_size(sheet)
        margin    = 0.066   # ~20 mm in ft
        tb_strip  = 0.164   # ~50 mm title block strip at bottom
        usable_h  = h - margin - tb_strip
        baseline  = tb_strip + margin / 2

        plan_views = [v for v in [
            result.get('floor_plan'), result.get('ceiling_plan')
        ] if v is not None]

        elev_views = result.get('elevations', [])

        if plan_views:
            # Place plan(s) in left half
            plan_zone_w = w * 0.45
            if len(plan_views) == 1:
                self._place_viewport_centered(
                    sheet, plan_views[0].Id,
                    plan_zone_w / 2 + margin,
                    baseline + usable_h / 2
                )
            else:
                # Stack floor + ceiling vertically
                self._place_viewport_centered(
                    sheet, plan_views[0].Id,
                    plan_zone_w / 2 + margin,
                    baseline + usable_h * 0.72
                )
                self._place_viewport_centered(
                    sheet, plan_views[1].Id,
                    plan_zone_w / 2 + margin,
                    baseline + usable_h * 0.28
                )

        if elev_views:
            # 2x2 grid in right half
            right_x = w * 0.5
            elev_zone_w = w - right_x - margin
            elev_zone_h = usable_h
            positions = [
                (right_x + elev_zone_w * 0.25, baseline + elev_zone_h * 0.75),
                (right_x + elev_zone_w * 0.75, baseline + elev_zone_h * 0.75),
                (right_x + elev_zone_w * 0.25, baseline + elev_zone_h * 0.25),
                (right_x + elev_zone_w * 0.75, baseline + elev_zone_h * 0.25),
            ]
            for i, ev in enumerate(elev_views[:4]):
                cx, cy = positions[i]
                self._place_viewport_centered(sheet, ev.Id, cx, cy)

    def _layout_separate(self, result, plan_sheet, elev_sheet):
        """
        Separate layout — plans on plan_sheet, elevations on elev_sheet.
        """
        margin   = 0.066
        tb_strip = 0.164

        # ── Plans sheet ──────────────────────────────
        if plan_sheet:
            w, h = self._get_sheet_size(plan_sheet)
            usable_h = h - margin - tb_strip
            baseline = tb_strip + margin / 2
            plan_views = [v for v in [
                result.get('floor_plan'), result.get('ceiling_plan')
            ] if v is not None]
            if len(plan_views) == 1:
                self._place_viewport_centered(
                    plan_sheet, plan_views[0].Id,
                    w / 2, baseline + usable_h / 2
                )
            elif len(plan_views) >= 2:
                self._place_viewport_centered(
                    plan_sheet, plan_views[0].Id,
                    w / 2, baseline + usable_h * 0.70
                )
                self._place_viewport_centered(
                    plan_sheet, plan_views[1].Id,
                    w / 2, baseline + usable_h * 0.28
                )

        # ── Elevations sheet ──────────────────────────
        if elev_sheet:
            w, h = self._get_sheet_size(elev_sheet)
            usable_h = h - margin - tb_strip
            baseline = tb_strip + margin / 2
            elev_views = result.get('elevations', [])
            positions = [
                (w * 0.25, baseline + usable_h * 0.75),
                (w * 0.75, baseline + usable_h * 0.75),
                (w * 0.25, baseline + usable_h * 0.25),
                (w * 0.75, baseline + usable_h * 0.25),
            ]
            for i, ev in enumerate(elev_views[:4]):
                cx, cy = positions[i]
                self._place_viewport_centered(elev_sheet, ev.Id, cx, cy)

    def _layout_views_on_sheets(self, result, titleblock_id, combined):
        """
        Create sheet(s) and place viewports for a single room result.
        Returns list of created ViewSheet objects.
        """
        room_item  = result['room_item']
        has_plans  = result['floor_plan'] or result['ceiling_plan']
        has_elevs  = bool(result['elevations'])
        sheets     = []

        if combined:
            # One sheet for everything
            if has_plans or has_elevs:
                sheet = self._create_sheet(room_item, titleblock_id)
                self._layout_combined(result, sheet)
                sheets.append(sheet)
        else:
            # Plans on their own sheet
            if has_plans:
                p_sheet = self._create_sheet(
                    room_item, titleblock_id,
                    name_suffix="Plans", num_suffix="P"
                )
                sheets.append(p_sheet)
            else:
                p_sheet = None

            # Elevations on their own sheet
            if has_elevs:
                e_sheet = self._create_sheet(
                    room_item, titleblock_id,
                    name_suffix="Elevations", num_suffix="E"
                )
                sheets.append(e_sheet)
            else:
                e_sheet = None

            self._layout_separate(result, p_sheet, e_sheet)

        return sheets


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
# ==================================================
if __name__ == '__main__':
    try:
        window = CreateRoomPlanWindow()
        window.ShowDialog()
    except Exception as ex:
        logger.error("Create Room Plan error: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
