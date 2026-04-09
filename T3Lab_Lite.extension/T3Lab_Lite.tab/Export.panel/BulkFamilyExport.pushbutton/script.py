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


# ─── Data Model ───────────────────────────────────────────────────────────────

class BlockItem(object):
    """Represents a discovered CAD block for DataGrid binding."""

    def __init__(self, name, curve_count, instance_count, curves):
        self.IsSelected    = True
        self.BlockName     = name
        self.CurveCount    = curve_count
        self.InstanceCount = instance_count
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
        for inst in collector:
            name = self._get_cad_name(inst)
            self._cad_instances.append(inst)
            self.cad_file_combo.Items.Add(name)
        if self._cad_instances:
            self.cad_file_combo.SelectedIndex = 0

    def _get_cad_name(self, inst):
        """Return the DWG/DXF source filename for an ImportInstance."""
        try:
            type_id = inst.GetTypeId()
            if type_id and type_id != ElementId.InvalidElementId:
                elem_type = doc.GetElement(type_id)
                if elem_type and hasattr(elem_type, 'Name') and elem_type.Name:
                    return elem_type.Name
        except Exception:
            pass
        return inst.Name if hasattr(inst, 'Name') else "Unknown CAD"

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

        idx = self.cad_file_combo.SelectedIndex
        if idx < 0 or idx >= len(self._cad_instances):
            forms.alert("Please select a CAD file.")
            return

        self._update_status("Scanning blocks...")
        import_inst = self._cad_instances[idx]

        try:
            blocks = self._scan_blocks(import_inst)
        except Exception as ex:
            logger.error("Scan error:\n{}".format(traceback.format_exc()))
            forms.alert("Error scanning blocks:\n{}".format(str(ex)))
            self._update_status("Scan failed")
            return

        if not blocks:
            forms.alert(
                "No blocks found in the selected CAD file.\n\n"
                "Make sure the CAD file contains block references.")
            self._update_status("No blocks found")
            return

        self._block_items = blocks
        self.blocks_grid.ItemsSource = blocks
        self._update_status("Found {} unique block(s)".format(len(blocks)))
        self.block_count_text.Text = "{} blocks found".format(len(blocks))

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
            if layer:
                name = "{}_Block_{:03d}".format(
                    re.sub(r'[^\w\-]', '_', layer), counter[0])
            else:
                name = "Block_{:03d}".format(counter[0])

            found[fp] = {'name': name, 'curves': curves, 'count': 1}

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
                data['name'], len(data['curves']), data['count'], data['curves']))
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
                ok = self._export_block(item, template_path, output_folder)
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

    def _export_block(self, block_item, template_path, output_folder):
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

        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
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

            # Bounding-rectangle extrusion (1 ft tall)
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

            extrusion_depth = 1.0
            fam_doc.FamilyCreate.NewExtrusion(
                True, profile, sketch_plane, extrusion_depth)

            # Model lines on the top face (Z = 1.0)
            top_sp = SketchPlane.Create(
                fam_doc,
                Plane.CreateByNormalAndOrigin(
                    XYZ.BasisZ, XYZ(0.0, 0.0, extrusion_depth)))

            for curve in curves:
                try:
                    p0 = curve.GetEndPoint(0)
                    p1 = curve.GetEndPoint(1)
                    np0 = XYZ(p0.X - cx, p0.Y - cy, extrusion_depth)
                    np1 = XYZ(p1.X - cx, p1.Y - cy, extrusion_depth)

                    if isinstance(curve, Line):
                        new_c = Line.CreateBound(np0, np1)
                    elif isinstance(curve, Arc):
                        ctr = curve.Center
                        nc  = XYZ(ctr.X - cx, ctr.Y - cy, extrusion_depth)
                        new_c = Arc.Create(
                            nc, curve.Radius,
                            curve.GetEndParameter(0),
                            curve.GetEndParameter(1),
                            XYZ.BasisX, XYZ.BasisY)
                    else:
                        new_c = Line.CreateBound(np0, np1)

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
        safe_name = re.sub(r'[^\w\-]', '',
                           block_item.BlockName.replace(' ', '_'))
        if not safe_name:
            safe_name = "Family"
        save_path = os.path.join(output_folder, "{}.rfa".format(safe_name))

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
