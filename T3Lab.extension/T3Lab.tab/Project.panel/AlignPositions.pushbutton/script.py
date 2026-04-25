# -*- coding: utf-8 -*-
"""
Align Positions
---------------
Snap element positions so their distance to a reference element
(Grid, Wall, or Column) becomes a clean multiple of 5 or 10 mm.

Workflow
  1. Pick a reference element (Grid / Wall / Column)
  2. Select surrounding elements
  3. Review the preview table
  4. Apply corrections

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""

__title__   = "Align\nPositions"
__author__  = "Tran Tien Thanh"
__version__ = "1.0.0"

# IMPORT LIBRARIES
# ==================================================
import os
import sys
import math
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('System.Data')

from System.Windows import WindowState
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind, Boolean
from System.Data import DataTable

from Autodesk.Revit.DB import (
    XYZ,
    Transaction,
    ElementTransformUtils,
    Grid,
    Wall,
    FamilyInstance,
    BuiltInCategory,
    LocationPoint,
    LocationCurve,
    FilteredElementCollector,
    Line,
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit, script

# Path setup
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# DEFINE VARIABLES
# ==================================================
doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()
output = script.get_output()
REVIT_VERSION = int(revit.doc.Application.VersionNumber)

# CLASS/FUNCTIONS
# ==================================================

# ════════════════════════════════════════════════════════════════
# CONSTANTS
# ════════════════════════════════════════════════════════════════
MM_PER_FOOT = 304.8
TOLERANCE_MM = 0.05        # ignore corrections smaller than this


def feet_to_mm(feet):
    return feet * MM_PER_FOOT


def mm_to_feet(mm):
    return mm / MM_PER_FOOT


def round_to_snap(value_mm, snap_mm):
    """Round *value_mm* to the nearest multiple of *snap_mm*."""
    return round(value_mm / snap_mm) * snap_mm


# ════════════════════════════════════════════════════════════════
# SELECTION FILTER
# ════════════════════════════════════════════════════════════════
class ReferenceFilter(ISelectionFilter):
    """Allow only Grid, Wall, or Column elements."""

    def AllowElement(self, element):
        if isinstance(element, Grid):
            return True
        if isinstance(element, Wall):
            return True
        if isinstance(element, FamilyInstance):
            cat = element.Category
            if cat is None:
                return False
            cat_id = cat.Id.IntegerValue
            if cat_id == int(BuiltInCategory.OST_Columns):
                return True
            if cat_id == int(BuiltInCategory.OST_StructuralColumns):
                return True
        return False

    def AllowReference(self, reference, position):
        return False


# ════════════════════════════════════════════════════════════════
# GEOMETRY HELPERS
# ════════════════════════════════════════════════════════════════
def get_element_location(element):
    """Return an XYZ location point for any element."""
    loc = element.Location
    if loc is not None:
        if isinstance(loc, LocationPoint):
            return loc.Point
        if isinstance(loc, LocationCurve):
            return loc.Curve.Evaluate(0.5, True)

    # Fallback: bounding-box centre
    bb = element.get_BoundingBox(None)
    if bb is not None:
        return XYZ(
            (bb.Min.X + bb.Max.X) / 2.0,
            (bb.Min.Y + bb.Max.Y) / 2.0,
            (bb.Min.Z + bb.Max.Z) / 2.0,
        )
    return None


def project_onto_curve(point, curve):
    """Project *point* onto *curve*; return the projected XYZ or None."""
    result = curve.Project(point)
    if result is not None:
        return result.XYZPoint
    return None


def get_element_orientation(element):
    """Returns the primary X direction vector for an element. Only XY plane."""
    vec = None
    if isinstance(element, Grid):
        vec = element.Curve.GetEndPoint(1) - element.Curve.GetEndPoint(0)
    elif isinstance(element, Wall):
        vec = element.Location.Curve.GetEndPoint(1) - element.Location.Curve.GetEndPoint(0)
    elif isinstance(element, FamilyInstance):
        vec = element.GetTransform().BasisX
    else:
        loc = element.Location
        if isinstance(loc, LocationCurve):
            vec = loc.Curve.GetEndPoint(1) - loc.Curve.GetEndPoint(0)
        elif isinstance(loc, LocationPoint):
            ang = loc.Rotation
            vec = XYZ(math.cos(ang), math.sin(ang), 0)

    if vec is not None:
        v_xy = XYZ(vec.X, vec.Y, 0)
        if v_xy.GetLength() > 1e-6:
            return v_xy.Normalize()
            
    return XYZ.BasisX




# ════════════════════════════════════════════════════════════════
# WPF WINDOW
# ════════════════════════════════════════════════════════════════
_GUI_DIR  = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))),
    'lib', 'GUI',
)
_XAML     = os.path.join(_GUI_DIR, 'Tools', 'AlignPositions.xaml')
_LOGO     = os.path.join(_GUI_DIR, 'T3Lab_logo.png')


class AlignPositionsWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)

        # State
        self._ref_element  = None
        self._ref_type     = None   # "Grid" | "Wall" | "Column"
        self._ref_data     = None   # tuple of geometry data
        self._selected_elements = []
        self._results      = []

        # DataTable for the DataGrid
        self._dt = DataTable()
        self._dt.Columns.Add("Apply",     Boolean)
        self._dt.Columns.Add("Category")
        self._dt.Columns.Add("Name")
        self._dt.Columns.Add("ElementId")
        self._dt.Columns.Add("StartPt")
        self._dt.Columns.Add("EndPt")
        self._dt.Columns.Add("LocDist")
        self._dt.Columns.Add("AngleStr")
        self._dt.Columns.Add("Correction")
        self.dg_elements.ItemsSource = self._dt.DefaultView

        # Initial button states
        self.btn_select.IsEnabled = False
        self.btn_apply.IsEnabled  = False

        self._load_logo()
        self._update_status("Ready - Pick a reference element to start")

    # ── logo & chrome ────────────────────────────────────────
    def _load_logo(self):
        try:
            if os.path.exists(_LOGO):
                bmp = BitmapImage()
                bmp.BeginInit()
                bmp.UriSource = Uri(_LOGO, UriKind.Absolute)
                bmp.EndInit()
                self.Icon = bmp
        except Exception as icon_ex:
            logger.warning("Could not set window icon: {}".format(icon_ex))

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

    # ── helpers ──────────────────────────────────────────────
    def _update_status(self, text):
        self.status_text.Text = text

    def _snap_feet(self):
        # Determine current snap logic
        is_metric = getattr(self, 'rb_metric', None) and self.rb_metric.IsChecked
        if is_metric:
            if self.rb_snap1.IsChecked: return 5.0 / MM_PER_FOOT
            if self.rb_snap2.IsChecked: return 10.0 / MM_PER_FOOT
            if getattr(self, 'rb_snap3', None) and self.rb_snap3.IsChecked: return 20.0 / MM_PER_FOOT
            return 5.0 / MM_PER_FOOT
        else:
            # Imperial => inches converted to feet
            if self.rb_snap1.IsChecked: return (1.0/8.0) / 12.0
            if self.rb_snap2.IsChecked: return (1.0/4.0) / 12.0
            if getattr(self, 'rb_snap3', None) and self.rb_snap3.IsChecked: return (1.0/2.0) / 12.0
            return (1.0/8.0) / 12.0

    # ── pick reference ───────────────────────────────────────
    def pick_reference_clicked(self, sender, e):
        self.Hide()
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                ReferenceFilter(),
                "Pick a reference element (Grid, Wall, or Column)",
            )
            self._set_reference(doc.GetElement(ref.ElementId))
        except OperationCanceledException:
            pass
        except Exception as ex:
            logger.error(str(ex))
        self.Show()

    def _set_reference(self, element):
        self._ref_element = element

        if isinstance(element, Grid):
            self._ref_type = "Grid"
            self.txt_ref_info.Text = "Grid: {}".format(element.Name)
            self.txt_ref_info.Foreground = \
                SolidColorBrush(Color.FromRgb(44, 62, 80))

        elif isinstance(element, Wall):
            self._ref_type = "Wall"
            w_mm = element.Width * MM_PER_FOOT
            self.txt_ref_info.Text = "Wall: Id {} (w={:.0f} mm)".format(
                element.Id.IntegerValue, w_mm)
            self.txt_ref_info.Foreground = \
                SolidColorBrush(Color.FromRgb(44, 62, 80))

        elif isinstance(element, FamilyInstance):
            self._ref_type = "Column"
            self.txt_ref_info.Text = "Column: {} (Id:{})".format(
                element.Symbol.Family.Name, element.Id.IntegerValue)
            self.txt_ref_info.Foreground = \
                SolidColorBrush(Color.FromRgb(44, 62, 80))

        # Setup local coordinate vectors
        self._ref_origin = get_element_location(element)
        self._ref_dir = get_element_orientation(element)

        # Reset downstream state
        self.btn_select.IsEnabled = True
        self._dt.Clear()
        self._results = []
        self._selected_elements = []
        self.btn_apply.IsEnabled = False
        self.txt_element_count.Text = "Scanning elements..."
        self._update_status("Reference set - Scanning for elements...")
        
        self._auto_find_elements()

    # ── select elements ──────────────────────────────────────
    def select_elements_clicked(self, sender, e):
        try:
            self._auto_find_elements()
        except Exception as ex:
            logger.error(str(ex))

    def _auto_find_elements(self):
        walls = list(FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(Wall).ToElements())
        grids = list(FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(Grid).ToElements())
        
        cols_arch = list(FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Columns).WhereElementIsNotElementType().ToElements())
        cols_str = list(FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsNotElementType().ToElements())

        elements = walls + grids + cols_arch + cols_str
        self._analyze(elements)

    # ── unit/snap value changed ──────────────────────────────
    def unit_changed(self, sender, e):
        if not hasattr(self, 'rb_metric') or not self.rb_metric:
            return
            
        from System.Windows import Visibility
        is_metric = self.rb_metric.IsChecked
        if is_metric:
            if hasattr(self, 'rb_snap1') and self.rb_snap1: self.rb_snap1.Content = " 5 mm"
            if hasattr(self, 'rb_snap2') and self.rb_snap2: self.rb_snap2.Content = " 10 mm"
            if hasattr(self, 'rb_snap3') and self.rb_snap3:
                self.rb_snap3.Visibility = Visibility.Collapsed
                if self.rb_snap3.IsChecked: self.rb_snap1.IsChecked = True
        else:
            if hasattr(self, 'rb_snap1') and self.rb_snap1: self.rb_snap1.Content = " 1/8 inch"
            if hasattr(self, 'rb_snap2') and self.rb_snap2: self.rb_snap2.Content = " 1/4 inch"
            if hasattr(self, 'rb_snap3') and self.rb_snap3:
                self.rb_snap3.Content = " 1/2 inch"
                self.rb_snap3.Visibility = Visibility.Visible

        if getattr(self, '_selected_elements', None) and getattr(self, '_ref_element', None):
            self._analyze(self._selected_elements)

    def snap_changed(self, sender, e):
        if self._selected_elements and self._ref_element:
            self._analyze(self._selected_elements)

    # ── analysis ─────────────────────────────────────────────
    def _analyze(self, elements):
        self._selected_elements = list(elements)
        self._dt.Clear()
        self._results = []
        snap_f = self._snap_feet()
        count = 0

        for elem in elements:
            if elem.Id == self._ref_element.Id:
                continue
            result = self._compute(elem, snap_f)
            if result is None:
                continue

            self._results.append(result)
            row = self._dt.NewRow()
            row["Apply"]      = True
            row["Category"]   = elem.Category.Name if elem.Category else "N/A"
            row["Name"]       = getattr(elem, 'Name', '') or "Id:{}".format(elem.Id.IntegerValue)
            row["ElementId"]  = str(elem.Id.IntegerValue)
            row["StartPt"]    = result.get("start_pt_str", "")
            row["EndPt"]      = result.get("end_pt_str", "")
            row["LocDist"]    = result["loc_str"]
            row["AngleStr"]   = result["angle_str"]
            row["Correction"] = result["corr_str"]
            self._dt.Rows.Add(row)
            count += 1

        self.btn_apply.IsEnabled = count > 0
        self.txt_element_count.Text = "{} element(s) need adjustment".format(count)
        self._update_status(
            "{} selected, {} need correction".format(len(elements), count))

    def _compute(self, elem, snap_feet):
        loc = get_element_location(elem)
        if loc is None or self._ref_origin is None:
            return None

        target_dir = get_element_orientation(elem)
        ref_x = self._ref_dir
        ref_y = XYZ(-ref_x.Y, ref_x.X, 0)

        is_curve_element = False
        c = None
        if isinstance(elem, Wall) and hasattr(elem.Location, "Curve"):
            c = elem.Location.Curve
            is_curve_element = c is not None and isinstance(c, Line)
        elif isinstance(elem, Grid):
            c = elem.Curve
            is_curve_element = c is not None and isinstance(c, Line)

        move_vector = None
        rot_correction = 0.0
        needs_move = False
        needs_rot = False
        new_curve = None
        
        diff_deg = 0.0
        start_pt_str = ""
        end_pt_str = ""

        # Format strings for UI using native settings
        is_metric = getattr(self, 'rb_metric', None) and self.rb_metric.IsChecked
        scale = MM_PER_FOOT if is_metric else 12.0
        fmt = "{:.2f}" if is_metric else "{:.3f}"
        fmt_corr = "{:+.2f}" if is_metric else "{:+.3f}"

        if is_curve_element:
            pt1 = c.GetEndPoint(0)
            pt2 = c.GetEndPoint(1)

            u1 = (pt1 - self._ref_origin).DotProduct(ref_x)
            v1 = (pt1 - self._ref_origin).DotProduct(ref_y)
            u2 = (pt2 - self._ref_origin).DotProduct(ref_x)
            v2 = (pt2 - self._ref_origin).DotProduct(ref_y)

            start_pt_str = (fmt + ", " + fmt).format(pt1.X * scale, pt1.Y * scale)
            end_pt_str = (fmt + ", " + fmt).format(pt2.X * scale, pt2.Y * scale)

            is_mostly_x = abs(u2 - u1) > abs(v2 - v1)

            if is_mostly_x:
                v_avg = (v1 + v2) / 2.0
                v_snap = round(v_avg / snap_feet) * snap_feet
                v1_snap, v2_snap = v_snap, v_snap
                u1_snap = round(u1 / snap_feet) * snap_feet
                u2_snap = round(u2 / snap_feet) * snap_feet
            else:
                u_avg = (u1 + u2) / 2.0
                u_snap = round(u_avg / snap_feet) * snap_feet
                u1_snap, u2_snap = u_snap, u_snap
                v1_snap = round(v1 / snap_feet) * snap_feet
                v2_snap = round(v2 / snap_feet) * snap_feet

            new_pt1 = self._ref_origin + u1_snap * ref_x + v1_snap * ref_y + XYZ(0,0,pt1.Z - self._ref_origin.Z)
            new_pt2 = self._ref_origin + u2_snap * ref_x + v2_snap * ref_y + XYZ(0,0,pt2.Z - self._ref_origin.Z)

            if new_pt1.DistanceTo(new_pt2) > 0.0026:
                if pt1.DistanceTo(new_pt1) > 0.001 or pt2.DistanceTo(new_pt2) > 0.001:
                    new_curve = Line.CreateBound(new_pt1, new_pt2)
                    needs_move = True
                    
                    # For grid fallback if needed
                    mid_pt = (pt1 + pt2) / 2.0
                    new_mid_pt = (new_pt1 + new_pt2) / 2.0
                    move_vector = new_mid_pt - mid_pt
                    
                    tar_theta = math.atan2(target_dir.Y, target_dir.X)
                    ref_theta = math.atan2(ref_x.Y, ref_x.X)
                    diff_rad = tar_theta - ref_theta
                    diff_deg = math.degrees(diff_rad) % 90
                    if diff_deg > 45: diff_deg = 90 - diff_deg

            dist_x_feet = (u1 + u2) / 2.0
            dist_y_feet = (v1 + v2) / 2.0
            cx_display = (u1_snap + u2_snap) / 2.0 - dist_x_feet
            cy_display = (v1_snap + v2_snap) / 2.0 - dist_y_feet

        else:
            # 1. Angle correction
            ref_theta = math.atan2(ref_x.Y, ref_x.X)
            tar_theta = math.atan2(target_dir.Y, target_dir.X)
            diff_rad = tar_theta - ref_theta
            
            snap_theta_rad = round(diff_rad / (math.pi/2)) * (math.pi/2)
            rot_correction = snap_theta_rad - diff_rad
            
            diff_deg = math.degrees(diff_rad) % 90
            if diff_deg > 45: diff_deg = 90 - diff_deg
            
            # 2. Local distance corrections
            vx = loc.X - self._ref_origin.X
            vy = loc.Y - self._ref_origin.Y
            v = XYZ(vx, vy, 0)
            
            dist_x_feet = v.DotProduct(ref_x)
            dist_y_feet = v.DotProduct(ref_y)
            
            snap_x_feet = round(dist_x_feet / snap_feet) * snap_feet
            snap_y_feet = round(dist_y_feet / snap_feet) * snap_feet
            
            corr_x_feet = snap_x_feet - dist_x_feet
            corr_y_feet = snap_y_feet - dist_y_feet
            
            move_vector = ref_x * corr_x_feet + ref_y * corr_y_feet
                
            if move_vector.GetLength() > 0.0026:
                needs_move = True
            
            if isinstance(elem, FamilyInstance):
                needs_rot = (abs(rot_correction) >= 0.001)

            cx_display = corr_x_feet
            cy_display = corr_y_feet

            start_pt_str = (fmt + ", " + fmt).format(loc.X * scale, loc.Y * scale)
            end_pt_str = "-"

        if not needs_move and not needs_rot:
            return None
            
        dx_display = dist_x_feet * scale
        dy_display = dist_y_feet * scale
        cxd = cx_display * scale
        cyd = cy_display * scale
        
        loc_str = (fmt + ", " + fmt).format(dx_display, dy_display)
        angle_str = "{:.1f}°".format(diff_deg)
        corr_str = (fmt_corr + ", " + fmt_corr).format(cxd, cyd) if needs_move else "0, 0"
        if needs_rot:
            corr_str += " | ⟲"
            
        return {
            "element": elem,
            "origin": loc,
            "loc_str": loc_str,
            "angle_str": angle_str,
            "corr_str": corr_str,
            "move_vector": move_vector if needs_move else None,
            "rot_correction": rot_correction if needs_rot else 0.0,
            "new_curve": new_curve,
            "start_pt_str": start_pt_str,
            "end_pt_str": end_pt_str
        }

    # ── apply ────────────────────────────────────────────────
    def apply_clicked(self, sender, e):
        if not self._results:
            return

        moved  = 0
        failed = 0
        
        self.Hide()

        t = Transaction(doc, "Align Positions")
        t.Start()
        try:
            for i, result in enumerate(self._results):
                # Respect the Apply checkbox
                if i < self._dt.Rows.Count:
                    if not self._dt.Rows[i]["Apply"]:
                        continue

                try:
                    elem_id = result["element"].Id
                    
                    if result.get("new_curve") is not None:
                        if isinstance(result["element"], Wall):
                            result["element"].Location.Curve = result["new_curve"]
                        elif isinstance(result["element"], Grid):
                            import Autodesk.Revit.DB as DB
                            try:
                                result["element"].SetCurveInView(DB.DatumExtentType.Model, doc.ActiveView, result["new_curve"])
                            except:
                                if result["move_vector"] is not None:
                                    ElementTransformUtils.MoveElement(doc, elem_id, result["move_vector"])
                    else:
                        if result["rot_correction"] != 0.0:
                            axis = Line.CreateBound(result["origin"], result["origin"] + XYZ.BasisZ)
                            ElementTransformUtils.RotateElement(doc, elem_id, axis, result["rot_correction"])
                        
                        if result["move_vector"] is not None:
                            ElementTransformUtils.MoveElement(
                                doc,
                                elem_id,
                                result["move_vector"],
                            )
                    moved += 1
                except Exception as ex:
                    logger.warning(
                        "Cannot move Id {}: {}".format(
                            result["element"].Id.IntegerValue, ex))
                    failed += 1

            t.Commit()
        except Exception as ex:
            t.RollBack()
            self._update_status("Error: {}".format(ex))
            self.Show()
            return

        msg = "Done! {} element(s) moved".format(moved)
        if failed:
            msg += ", {} failed".format(failed)
        self._update_status(msg)

        # Clear table
        self._dt.Clear()
        self._results = []
        self._selected_elements = []
        self.btn_apply.IsEnabled = False
        self.txt_element_count.Text = "{} element(s) moved".format(moved)
        
        self.Show()


# MAIN SCRIPT
# ==================================================
if __name__ == '__main__':
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)
    window = AlignPositionsWindow()
    window.ShowDialog()
