# -*- coding: utf-8 -*-
"""
Bulk Family Export

Export CAD blocks from imported DWG/DXF files as individual Revit families.
Select an imported CAD file, choose a family category, scan for blocks,
and export each block as a separate .rfa family file.

Author: T3Lab
"""

from __future__ import unicode_literals

__author__  = "T3Lab"
__title__   = "Bulk Family\nExport"
__version__ = "1.0.0"

# ─── Imports ──────────────────────────────────────────────────────────────────

import os
import re
import traceback

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import WindowState
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import (
    ImportInstance, FilteredElementCollector,
    Options, GeometryInstance,
    Line, Arc, XYZ, Plane,
    CurveArray, CurveArrArray,
    SketchPlane, SaveAsOptions,
    Transaction, ElementId,
)

# ─── Revit context ────────────────────────────────────────────────────────────

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
app   = __revit__.Application

logger = script.get_logger()

# ─── Category-to-template mapping ────────────────────────────────────────────

CATEGORY_TEMPLATES = [
    ("Generic Model",        ["Generic Model.rft", "Metric Generic Model.rft"]),
    ("Door",                 ["Door.rft", "Metric Door.rft"]),
    ("Window",               ["Window.rft", "Metric Window.rft"]),
    ("Furniture",            ["Furniture.rft", "Metric Furniture.rft"]),
    ("Plumbing Fixture",     ["Plumbing Fixture.rft", "Metric Plumbing Fixture.rft"]),
    ("Electrical Equipment", ["Electrical Equipment.rft"]),
    ("Mechanical Equipment", ["Mechanical Equipment.rft"]),
    ("Specialty Equipment",  ["Specialty Equipment.rft", "Metric Specialty Equipment.rft"]),
    ("Casework",             ["Casework.rft", "Metric Casework.rft"]),
    ("Columns",              ["Column.rft", "Metric Column.rft"]),
    ("Lighting Fixture",     ["Lighting Fixture.rft", "Metric Lighting Fixture.rft"]),
    ("Site",                 ["Site.rft", "Metric Site.rft"]),
    ("Entourage",            ["Entourage.rft", "Metric Entourage.rft"]),
]

DISCIPLINES = [
    "Architecture",
    "Structure",
    "Mechanical",
    "Electrical",
    "Plumbing",
    "Fire Protection",
    "General"
]


# ─── Data Model ───────────────────────────────────────────────────────────────

class BlockItem(object):
    """Represents a discovered CAD block for DataGrid binding."""

    def __init__(self, name, curve_count, instance_count, curves, layer_level=""):
        self.IsSelected    = True
        self.BlockName     = name
        self.CurveCount    = curve_count
        self.InstanceCount = instance_count
        self.LayerLevel    = layer_level
        self._curves       = curves          # internal use only


# ─── Main Window ──────────────────────────────────────────────────────────────

class BulkFamilyExportWindow(forms.WPFWindow):

    def __init__(self):
        ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(
            os.path.dirname(__file__))))
        xaml_path = os.path.join(ext_dir, 'lib', 'GUI', 'BulkFamilyExport.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        self._ext_dir       = ext_dir
        self._cad_instances = []
        self._block_items   = []

        self._load_logo()
        self._init_cad_files()
        self._init_disciplines()
        self._init_categories()
        self._update_status("Ready")

    # ── Logo ──────────────────────────────────────────────────────────────

    def _load_logo(self):
        """Load T3Lab logo into title bar and window icon."""
        try:
            logo_path = os.path.join(self._ext_dir, 'lib', 'GUI', 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                self.logo_image.Source = bitmap
                self.Icon = bitmap
        except Exception:
            pass

    # ── Initialisation ────────────────────────────────────────────────────

    def _init_cad_files(self):
        """Populate the CAD file ComboBox with all ImportInstances."""
        collector = FilteredElementCollector(doc).OfClass(ImportInstance)
        self.cad_file_combo.Items.Add("<All Imported CAD Files>")
        for inst in collector:
            name = self._get_cad_name(inst)
            self._cad_instances.append(inst)
            self.cad_file_combo.Items.Add(name)
        if self._cad_instances:
            self.cad_file_combo.SelectedIndex = 0

    def _get_cad_name(self, inst):
        """Return the Type Name for an ImportInstance."""
        try:
            type_id = inst.GetTypeId()
            if type_id and type_id != ElementId.InvalidElementId:
                elem_type = doc.GetElement(type_id)
                if elem_type and hasattr(elem_type, 'Name') and elem_type.Name:
                    return elem_type.Name
        except Exception:
            pass
        return inst.Name if hasattr(inst, 'Name') else "Unknown CAD Type"

    def _init_disciplines(self):
        """Populate the discipline ComboBox."""
        for name in DISCIPLINES:
            self.discipline_combo.Items.Add(name)
        self.discipline_combo.SelectedIndex = 6

    def _init_categories(self):
        """Populate the category ComboBox."""
        for name, _ in CATEGORY_TEMPLATES:
            self.category_combo.Items.Add(name)
        self.category_combo.SelectedIndex = 0

    def _update_status(self, text):
        try:
            self.status_text.Text = text
        except Exception:
            pass

    # ── Block Scanning ────────────────────────────────────────────────────

    def scan_blocks_clicked(self, sender, e):
        """Event handler: scan the selected CAD file for blocks."""
        if not self._cad_instances:
            forms.alert("No imported CAD files found in the document.")
            return

        idx = self.cad_file_combo.SelectedIndex - 1
        if idx < -1 or idx >= len(self._cad_instances):
            forms.alert("Please select a CAD file.")
            return

        self._update_status("Scanning blocks...")

        blocks = []
        try:
            if idx == -1:
                # All imported CAD files
                name_counts = {}
                for inst in self._cad_instances:
                    item = self._scan_entire_cad(inst)
                    if item:
                        base_name = item.BlockName
                        if base_name in name_counts:
                            name_counts[base_name] += 1
                            item.BlockName = "{}_{}".format(base_name, name_counts[base_name])
                        else:
                            name_counts[base_name] = 1
                        blocks.append(item)
            else:
                import_inst = self._cad_instances[idx]
                blocks = self._scan_blocks(import_inst)
                if not blocks:
                    # Fallback: treat the entire CAD as one block
                    item = self._scan_entire_cad(import_inst)
                    if item:
                        blocks.append(item)
        except Exception as ex:
            logger.error("Scan error:\n{}".format(traceback.format_exc()))
            forms.alert("Error scanning blocks:\n{}".format(str(ex)))
            self._update_status("Scan failed")
            return

        if not blocks:
            forms.alert(
                "No blocks or curves found in the selected CAD file(s).\n\n"
                "Make sure the CAD files contain geometry.")
            self._update_status("No geometry found")
            return

        self._block_items = blocks
        self.blocks_grid.ItemsSource = blocks
        self._update_status("Found {} unique item(s)".format(len(blocks)))
        self.block_count_text.Text = "{} items found".format(len(blocks))

    def _scan_entire_cad(self, import_inst):
        """Treat an ImportInstance as a single block and return a BlockItem."""
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True

        geom = import_inst.get_Geometry(opt)
        if not geom:
            return None

        min_len = getattr(app, 'ShortCurveTolerance', 0.00256)

        def is_curve(item):
            try:
                from Autodesk.Revit.DB import Curve as _Curve
                return (isinstance(item, _Curve)
                        and item.IsBound
                        and item.Length >= min_len)
            except Exception:
                return False

        def collect_curves(geo_elem):
            curves = []
            for item in geo_elem:
                if is_curve(item):
                    curves.append(item)
                elif isinstance(item, GeometryInstance):
                    nested = item.GetInstanceGeometry()
                    if nested:
                        curves.extend(collect_curves(nested))
            return curves

        curves = collect_curves(geom)
        if curves:
            name = self._get_cad_name(import_inst)
            
            level_name = ""
            try:
                level_param = import_inst.get_Parameter(DB.BuiltInParameter.IMPORT_BASE_LEVEL)
                if level_param and level_param.HasValue:
                    level_name = level_param.AsValueString()
                else:
                    level_id = import_inst.LevelId
                    if level_id and str(level_id) != "-1":
                        level_elem = doc.GetElement(level_id)
                        if level_elem:
                            level_name = level_elem.Name
            except Exception:
                pass
                
            return BlockItem(name, len(curves), 1, curves, layer_level=level_name)
        return None

    def _scan_blocks(self, import_inst):
        """Walk the geometry tree of an ImportInstance and return unique blocks
        as a list of BlockItem objects."""
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True

        geom = import_inst.get_Geometry(opt)
        if not geom:
            return []

        min_len = getattr(app, 'ShortCurveTolerance', 0.00256)
        found   = {}      # fingerprint -> {name, curves, count}
        counter = [0]     # mutable counter (IronPython 2 – no nonlocal)

        # ── helpers ──

        def is_curve(item):
            try:
                from Autodesk.Revit.DB import Curve as _Curve
                return (isinstance(item, _Curve)
                        and item.IsBound
                        and item.Length >= min_len)
            except Exception:
                return False

        def collect_curves(geo_elem):
            """Recursively collect all valid curves from a geometry element."""
            curves = []
            for item in geo_elem:
                if is_curve(item):
                    curves.append(item)
                elif isinstance(item, GeometryInstance):
                    nested = item.GetInstanceGeometry()
                    if nested:
                        curves.extend(collect_curves(nested))
            return curves

        def fingerprint(curves):
            """Rotation- and position-invariant fingerprint for a set of curves."""
            n     = len(curves)
            total = round(sum(c.Length for c in curves), 1)
            return (n, total)

        def style_name(geo_inst):
            """Try to retrieve the DWG layer name from the GraphicsStyle."""
            try:
                sid = geo_inst.GraphicsStyleId
                if sid and sid != ElementId.InvalidElementId:
                    style = doc.GetElement(sid)
                    if style and hasattr(style, 'Name') and style.Name:
                        return style.Name
            except Exception:
                pass
            return None

        def register(curves, geo_inst):
            """Register a block (or increment its instance count)."""
            fp = fingerprint(curves)
            if fp in found:
                found[fp]['count'] += 1
                return

            layer = style_name(geo_inst)
            counter[0] += 1
            
            block_name = ""
            try:
                if hasattr(geo_inst, 'Symbol') and geo_inst.Symbol:
                    block_name = geo_inst.Symbol.Name
            except Exception:
                pass

            if block_name:
                name = block_name
            elif layer:
                name = "{}_Block_{:03d}".format(layer, counter[0])
            else:
                name = "Block_{:03d}".format(counter[0])

            found[fp] = {'name': name, 'curves': curves, 'count': 1, 'layer': layer or ""}

        # ── walk geometry tree ──

        def walk(geo_elem, depth):
            for item in geo_elem:
                if not isinstance(item, GeometryInstance):
                    continue

                inst_geom = item.GetInstanceGeometry()
                if not inst_geom:
                    continue

                if depth == 0:
                    # Top-level container (DWG model space) – dive in
                    walk(inst_geom, depth + 1)
                else:
                    # Block reference at depth >= 1
                    curves = collect_curves(inst_geom)
                    if curves:
                        register(curves, item)

        walk(geom, 0)

        # Convert to sorted list of BlockItem
        items = []
        for data in sorted(found.values(), key=lambda d: d['name']):
            items.append(BlockItem(
                data['name'], len(data['curves']), data['count'], data['curves'], layer_level=data.get('layer', "")))
        return items

    # ── Folder picker ─────────────────────────────────────────────────────

    def browse_folder_clicked(self, sender, e):
        folder = forms.pick_folder()
        if folder:
            self.output_path.Text = folder

    # ── Select / Deselect ─────────────────────────────────────────────────

    def select_all_clicked(self, sender, e):
        for item in self._block_items:
            item.IsSelected = True
        self.blocks_grid.Items.Refresh()

    def deselect_all_clicked(self, sender, e):
        for item in self._block_items:
            item.IsSelected = False
        self.blocks_grid.Items.Refresh()

    # ── Export ────────────────────────────────────────────────────────────

    def export_clicked(self, sender, e):
        """Export each selected block as a .rfa family file."""
        output_folder = self.output_path.Text
        if not output_folder or not os.path.isdir(output_folder):
            forms.alert("Please select a valid output folder.")
            return

        selected = [b for b in self._block_items if b.IsSelected]
        if not selected:
            forms.alert("No blocks selected for export.")
            return

        cat_idx = self.category_combo.SelectedIndex
        if cat_idx < 0:
            forms.alert("Please select a family category.")
            return

        category_name = CATEGORY_TEMPLATES[cat_idx][0]
        
        disc_idx = self.discipline_combo.SelectedIndex
        if disc_idx < 0:
            forms.alert("Please select a discipline.")
            return
        discipline_name = DISCIPLINES[disc_idx]
        template_path = self._find_template(cat_idx)
        if not template_path:
            forms.alert(
                "Could not find family template for '{}'.\n\n"
                "Please ensure Revit family templates are installed.".format(
                    category_name))
            return

        self._update_status("Exporting {} block(s)...".format(len(selected)))
        success, failed = 0, 0

        for item in selected:
            try:
                self._update_status("Exporting: {}".format(item.BlockName))
                ok = self._export_block(item, template_path, output_folder, discipline_name, category_name)
                if ok:
                    success += 1
                else:
                    failed += 1
            except Exception:
                logger.error("Export '{}' failed:\n{}".format(
                    item.BlockName, traceback.format_exc()))
                failed += 1

        self._update_status(
            "Done: {} exported, {} failed".format(success, failed))
        forms.alert(
            "Export complete!\n\n"
            "Succeeded: {}\nFailed: {}\n\n"
            "Output folder:\n{}".format(success, failed, output_folder))

    # ── Template lookup ───────────────────────────────────────────────────

    def _find_template(self, cat_idx):
        """Locate the .rft family template for the chosen category."""
        _, template_names = CATEGORY_TEMPLATES[cat_idx]

        search_dirs = []

        # Revit-configured path (most reliable)
        try:
            tdir = app.FamilyTemplatePath
            if tdir and os.path.isdir(tdir):
                search_dirs.append(tdir)
        except Exception:
            pass

        # Standard fallback paths
        ver  = app.VersionNumber
        base = r"C:\ProgramData\Autodesk\RVT {}".format(ver)
        for sub in ("English", "", "English-Imperial", "English_I"):
            if sub:
                search_dirs.append(
                    os.path.join(base, "Family Templates", sub))
            else:
                search_dirs.append(os.path.join(base, "Family Templates"))

        for d in search_dirs:
            if not os.path.isdir(d):
                continue
            for tname in template_names:
                fp = os.path.join(d, tname)
                if os.path.isfile(fp):
                    return fp
        return None

    # ── Single-block export ───────────────────────────────────────────────

    def _export_block(self, block_item, template_path, output_folder, discipline_name, category_name):
        """Create a .rfa family from a block's curves and save it."""
        curves = block_item._curves
        if not curves:
            return False

        # Create a fresh family document
        fam_doc = app.NewFamilyDocument(template_path)

        # Compute bounding box and centre
        xs, ys = [], []
        for c in curves:
            try:
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                xs.extend([p0.X, p1.X])
                ys.extend([p0.Y, p1.Y])
            except Exception:
                pass

        if not xs:
            fam_doc.Close(False)
            return False

        is_door = "door" in category_name.lower()
        door_width = None
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
        
        if is_door:
            frame_xs = []
            frame_ys = []
            for curve in curves:
                if isinstance(curve, Arc):
                    try:
                        C = curve.Center
                        p0 = curve.GetEndPoint(0)
                        p1 = curve.GetEndPoint(1)
                        frame_xs.append(C.X)
                        frame_ys.append(C.Y)
                        if abs(p0.Y - C.Y) < abs(p1.Y - C.Y):
                            frame_xs.append(p0.X)
                            frame_ys.append(p0.Y)
                        else:
                            frame_xs.append(p1.X)
                            frame_ys.append(p1.Y)
                    except Exception:
                        pass
            if frame_xs:
                cx = (min(frame_xs) + max(frame_xs)) / 2.0
                cy = (min(frame_ys) + max(frame_ys)) / 2.0
                calc_w = max(frame_xs) - min(frame_xs)
                if calc_w > 0.01:
                    door_width = calc_w
            else:
                cx = (min_x + max_x) / 2.0
                cy = (min_y + max_y) / 2.0
                door_width = max_x - min_x
        else:
            cx = (min_x + max_x) / 2.0
            cy = (min_y + max_y) / 2.0
            
        half_w = max((max_x - min_x) / 2.0, 0.01)   # avoid zero
        half_h = max((max_y - min_y) / 2.0, 0.01)

        t = Transaction(fam_doc, 'Create Block Geometry')
        t.Start()
        try:
            # Find or create a Z-up sketch plane
            sketch_plane = None
            for sp in FilteredElementCollector(fam_doc).OfClass(SketchPlane):
                try:
                    if abs(sp.GetPlane().Normal.Z - 1.0) < 0.001:
                        sketch_plane = sp
                        break
                except Exception:
                    pass
            if not sketch_plane:
                sketch_plane = SketchPlane.Create(
                    fam_doc,
                    Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero))

            from Autodesk.Revit.DB import FamilyElementVisibility, FamilyElementVisibilityType, GraphicsStyleType, Transform

            THICKNESS = 0.1312
            HEIGHT = 7.2178
            extrusion_depth = HEIGHT if is_door else 1.0

            # Get SubCategories for Doors
            swing_gs = None
            frame_gs = None
            if is_door:
                try:
                    fam_cat = fam_doc.OwnerFamily.FamilyCategory
                    def get_or_create_subcat(name):
                        if fam_cat.SubCategories.Contains(name):
                            return fam_cat.SubCategories.get_Item(name)
                        return fam_doc.Settings.Categories.NewSubcategory(fam_cat, name)
                    
                    swing_subcat = get_or_create_subcat("Plan Swing")
                    frame_subcat = get_or_create_subcat("Frame/Mullion")
                    if swing_subcat:
                        swing_gs = swing_subcat.GetGraphicsStyle(GraphicsStyleType.Projection)
                    if frame_subcat:
                        frame_gs = frame_subcat.GetGraphicsStyle(GraphicsStyleType.Projection)
                except Exception:
                    pass

            # Setup Parameters via JSONtoFamily Binding Logic
            param_height = None
            param_material = None
            try:
                fam_mgr = fam_doc.FamilyManager
                for param in fam_mgr.Parameters:
                    pname = param.Definition.Name.lower()
                    if pname in ["height", "chiều cao"]:
                        fam_mgr.Set(param, extrusion_depth)
                        param_height = param
                    elif pname in ["width", "chiều rộng"]:
                        if door_width:
                            fam_mgr.Set(param, door_width)
                    elif pname in ["depth", "chiều sâu", "length", "chiều dài"]:
                        if not is_door and half_h * 2.0 > 0.01:
                            fam_mgr.Set(param, half_h * 2.0)
                    elif pname in ["material", "vật liệu"]:
                        param_material = param
            except Exception:
                pass

            # Generative logic
            if not is_door:
                # Bounding-rectangle extrusion for generic blocks
                c1 = XYZ(-half_w, -half_h, 0.0)
                c2 = XYZ( half_w, -half_h, 0.0)
                c3 = XYZ( half_w,  half_h, 0.0)
                c4 = XYZ(-half_w,  half_h, 0.0)

                rect = CurveArray()
                rect.Append(Line.CreateBound(c1, c2))
                rect.Append(Line.CreateBound(c2, c3))
                rect.Append(Line.CreateBound(c3, c4))
                rect.Append(Line.CreateBound(c4, c1))

                profile = CurveArrArray()
                profile.Append(rect)

                ext_box = fam_doc.FamilyCreate.NewExtrusion(
                    True, profile, sketch_plane, extrusion_depth)
                try:
                    if param_height:
                        end_p = ext_box.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                        if end_p:
                            fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(end_p, param_height)
                    if param_material:
                        mat_p = ext_box.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                        if mat_p:
                            fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(mat_p, param_material)
                except Exception:
                    pass

            top_sp = SketchPlane.Create(
                fam_doc,
                Plane.CreateByNormalAndOrigin(
                    XYZ.BasisZ, XYZ(0.0, 0.0, extrusion_depth) if not is_door else XYZ.Zero))

            for curve in curves:
                try:
                    if is_door:
                        translator = Transform.CreateTranslation(XYZ(-cx, -cy, 0.0))
                        new_c = curve.CreateTransformed(translator)
                        
                        if isinstance(curve, Line):
                            sym_line = fam_doc.FamilyCreate.NewSymbolicCurve(new_c, sketch_plane)
                            if frame_gs:
                                sym_line.Subcategory = frame_gs
                                
                        elif isinstance(curve, Arc):
                            sym_arc = fam_doc.FamilyCreate.NewSymbolicCurve(new_c, sketch_plane)
                            if swing_gs:
                                sym_arc.Subcategory = swing_gs
                                
                            # 3D Panel Extrusion: aim towards CLOSED position
                            ctr = curve.Center
                            nc = ctr + XYZ(-cx, -cy, 0.0)
                            
                            p0_orig = curve.GetEndPoint(0)
                            p1_orig = curve.GetEndPoint(1)
                            if abs(p0_orig.Y - ctr.Y) < abs(p1_orig.Y - ctr.Y):
                                p_closed_orig = p0_orig
                            else:
                                p_closed_orig = p1_orig
                                
                            np_closed = p_closed_orig + XYZ(-cx, -cy, 0.0)
                            
                            v_dir = (np_closed - nc).Normalize()
                            v_ortho = XYZ(-v_dir.Y, v_dir.X, 0.0)
                            half_t = THICKNESS / 2.0
                            
                            pt1 = nc + v_ortho * half_t
                            pt2 = nc - v_ortho * half_t
                            pt3 = pt2 + v_dir * curve.Radius
                            pt4 = pt1 + v_dir * curve.Radius
                            
                            p_rect = CurveArray()
                            p_rect.Append(Line.CreateBound(pt1, pt2))
                            p_rect.Append(Line.CreateBound(pt2, pt3))
                            p_rect.Append(Line.CreateBound(pt3, pt4))
                            p_rect.Append(Line.CreateBound(pt4, pt1))
                            
                            p_profile = CurveArrArray()
                            p_profile.Append(p_rect)
                            
                            panel_ext = fam_doc.FamilyCreate.NewExtrusion(
                                True, p_profile, sketch_plane, HEIGHT)
                                
                            try:
                                vis = FamilyElementVisibility(FamilyElementVisibilityType.Model)
                                vis.IsShownInTopBottom = False
                                panel_ext.SetVisibility(vis)
                                
                                if param_height:
                                    end_p = panel_ext.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                                    if end_p:
                                        fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(end_p, param_height)
                                if param_material:
                                    mat_p = panel_ext.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                                    if mat_p:
                                        fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(mat_p, param_material)
                            except Exception:
                                pass
                                
                    else:
                        translator = Transform.CreateTranslation(XYZ(-cx, -cy, extrusion_depth))
                        new_c = curve.CreateTransformed(translator)
                        fam_doc.FamilyCreate.NewModelCurve(new_c, top_sp)
                except Exception:
                    pass

            t.Commit()
        except Exception:
            try:
                t.RollBack()
            except Exception:
                pass
            fam_doc.Close(False)
            raise

        # Save .rfa to output folder
        safe_cad_name = block_item.BlockName.strip()
        if not safe_cad_name:
            safe_cad_name = "Family"
            
        base_name = "{} - {} - {}".format(discipline_name, category_name, safe_cad_name)
        base_name = re.sub(r'[\\/*?:"<>|]', "", base_name)
        
        save_path = os.path.join(output_folder, "{}.rfa".format(base_name))
        counter = 1
        while os.path.exists(save_path):
            save_path = os.path.join(output_folder, "{}_{}.rfa".format(base_name, counter))
            counter += 1

        try:
            opts = SaveAsOptions()
            opts.OverwriteExistingFile = True
            fam_doc.SaveAs(save_path, opts)
        finally:
            fam_doc.Close(False)

        logger.info("Exported: {}".format(save_path))
        return True

    # ── Window chrome handlers ────────────────────────────────────────────

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


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == '__main__':
    try:
        window = BulkFamilyExportWindow()
        window.ShowDialog()
    except Exception:
        logger.error(traceback.format_exc())
        forms.alert("Unexpected error. Check the pyRevit log for details.")
