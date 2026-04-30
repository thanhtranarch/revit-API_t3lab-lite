# -*- coding: utf-8 -*-
"""
CAD to Beam
-----------
Create structural beams (Structural Framing) from CAD lines or layers.

Workflow:
  1. Select a CAD Link or Import.
  2. Select a Layer containing beam centerlines.
  3. Select a Beam Type (as Template) and Level.
  4. Specify Width/Height/Offset (Optional).
  5. Enable 'Read nearby Text' to auto-detect dimensions from Revit TextNotes.
  6. Click Generate to create beams.

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
--------------------------------------------------------
"""

__title__   = "CAD to\nBeam"
__author__  = "Tran Tien Thanh"
__version__ = "1.1.0"

import os
import sys
import clr
import re

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ImportInstance,
    BuiltInCategory,
    Transaction,
    Line,
    Structure,
    FamilySymbol,
    Level,
    Options,
    GeometryInstance,
    XYZ,
    TextNote
)
from pyrevit import forms, revit, script

# Path setup
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

_GUI_DIR = os.path.join(lib_dir, 'GUI')
_XAML = os.path.join(_GUI_DIR, 'Tools', 'CADtoBeam.xaml')

def parse_beam_dimensions(text):
    """Parse dimensions like 200x400 or B200x500 from string."""
    match = re.search(r'(\d+)\s*[xX*]\s*(\d+)', text)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None

def get_or_create_type(template_symbol, width, height):
    """Find or create a family symbol with specified width/height (mm)."""
    family = template_symbol.Family
    target_name = "{}x{}".format(width, height)
    
    # 1. Search existing
    for symbol_id in family.GetFamilySymbolIds():
        symbol = doc.GetElement(symbol_id)
        if revit.DB.Element.Name.__get__(symbol) == target_name:
            return symbol
        
        # Check parameters
        p_b = symbol.LookupParameter('b') or symbol.LookupParameter('Width') or symbol.LookupParameter('B')
        p_h = symbol.LookupParameter('h') or symbol.LookupParameter('Height') or symbol.LookupParameter('H')
        if p_b and p_h:
            if abs(p_b.AsDouble() * 304.8 - width) < 1.0 and abs(p_h.AsDouble() * 304.8 - height) < 1.0:
                return symbol
    
    # 2. Create New
    try:
        new_symbol = template_symbol.Duplicate(target_name)
        p_b = new_symbol.LookupParameter('b') or new_symbol.LookupParameter('Width') or new_symbol.LookupParameter('B')
        p_h = new_symbol.LookupParameter('h') or new_symbol.LookupParameter('Height') or new_symbol.LookupParameter('H')
        if p_b: p_b.Set(width / 304.8)
        if p_h: p_h.Set(height / 304.8)
        return new_symbol
    except Exception as ex:
        logger.debug("Could not create type {}: {}".format(target_name, ex))
        return template_symbol

# ════════════════════════════════════════════════════════════════
# CORE LOGIC
# ════════════════════════════════════════════════════════════════
class BeamGenerator:
    def __init__(self, doc):
        self.doc = doc

    def generate_beams(self, cad_instance, layer_name, base_beam_type, level, 
                       offset_mm=0.0, default_w=300, default_h=600, 
                       read_text=False, search_radius_mm=500):
        
        offset_ft = offset_mm / 304.8
        radius_ft = search_radius_mm / 304.8
        
        text_notes = []
        if read_text:
            text_notes = FilteredElementCollector(self.doc, self.doc.ActiveView.Id).OfClass(TextNote).ToElements()

        curves = self._extract_curves(cad_instance, layer_name)
        if not curves: return 0

        count = 0
        with Transaction(self.doc, "CAD to Beam Generated"):
            if not base_beam_type.IsActive: base_beam_type.Activate()
            
            for curve in curves:
                try:
                    target_w, target_h = default_w, default_h
                    if text_notes:
                        dims = self._find_nearby_dims(curve, text_notes, radius_ft)
                        if dims: target_w, target_h = dims
                    
                    current_type = base_beam_type
                    if target_w != default_w or target_h != default_h or read_text:
                        current_type = get_or_create_type(base_beam_type, target_w, target_h)
                    
                    if not current_type.IsActive: current_type.Activate()

                    beam = self.doc.Create.NewFamilyInstance(curve, current_type, level, Structure.StructuralType.Beam)
                    if abs(offset_ft) > 0.0001:
                        p = beam.get_Parameter(revit.DB.BuiltInParameter.STRUCTURAL_BEAM_Z_OFFSET_VALUE)
                        if p: p.Set(offset_ft)
                    count += 1
                except: pass
        return count

    def _extract_curves(self, cad_instance, layer_name):
        curves = []
        opt = Options()
        geom = cad_instance.get_Geometry(opt)
        for obj in geom:
            if isinstance(obj, GeometryInstance):
                sym_geom = obj.GetSymbolGeometry()
                inst_transform = obj.Transform
                for sym_obj in sym_geom:
                    g_style = self.doc.GetElement(sym_obj.GraphicsStyleId)
                    if g_style and g_style.GraphicsStyleCategory.Name == layer_name:
                        if isinstance(sym_obj, Line):
                            curves.append(sym_obj.CreateTransformed(inst_transform))
        return curves

    def _find_nearby_dims(self, curve, text_notes, radius_ft):
        mid_pt = curve.Evaluate(0.5, True)
        best_dist = radius_ft
        found_dims = None
        for tn in text_notes:
            dist = tn.Coord.DistanceTo(mid_pt)
            if dist < best_dist:
                dims = parse_beam_dimensions(tn.Text)
                if dims:
                    found_dims = dims
                    best_dist = dist
        return found_dims

# ════════════════════════════════════════════════════════════════
# UI / HEADLESS WRAPPER
# ════════════════════════════════════════════════════════════════
class CADtoBeamWindow(forms.WPFWindow):
    # ... (Keeping most of the UI logic, but calling BeamGenerator)
    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)
        self._populate_initial_data()
        self.generator = BeamGenerator(doc)

    def _populate_initial_data(self):
        cad_links = FilteredElementCollector(doc).OfClass(ImportInstance).ToElements()
        self.cad_map = { (c.get_Parameter(revit.DB.BuiltInParameter.IMPORT_SYMBOL_NAME).AsString() or str(c.Id)): c for c in cad_links }
        self.cb_cad_links.ItemsSource = sorted(self.cad_map.keys())
        beam_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(FamilySymbol).ToElements()
        self.beam_type_map = { "{}: {}".format(t.FamilyName, revit.DB.Element.Name.__get__(t)): t for t in beam_types }
        self.cb_beam_types.ItemsSource = sorted(self.beam_type_map.keys())
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        self.level_map = { l.Name: l for l in levels }
        self.cb_levels.ItemsSource = sorted(self.level_map.keys())

    def cad_link_changed(self, sender, e):
        name = self.cb_cad_links.SelectedItem
        if not name: return
        layers = set()
        geom = self.cad_map[name].get_Geometry(Options())
        for obj in geom:
            if isinstance(obj, GeometryInstance):
                for sym_obj in obj.GetSymbolGeometry():
                    style = doc.GetElement(sym_obj.GraphicsStyleId)
                    if style: layers.add(style.GraphicsStyleCategory.Name)
        self.cb_layers.ItemsSource = sorted(list(layers))

    def generate_clicked(self, sender, e):
        count = self.generator.generate_beams(
            self.cad_map[self.cb_cad_links.SelectedItem],
            self.cb_layers.SelectedItem,
            self.beam_type_map[self.cb_beam_types.SelectedItem],
            self.level_map[self.cb_levels.SelectedItem],
            float(self.txt_offset.Text or 0),
            float(self.txt_width.Text or 300),
            float(self.txt_height.Text or 600),
            self.chk_read_text.IsChecked,
            float(self.txt_radius.Text or 500)
        )
        forms.alert("Created {} beams.".format(count))
        self.Close()

def run_headless(args_json):
    import json
    try:
        data = json.loads(args_json)
        gen = BeamGenerator(doc)
        
        cad = doc.GetElement(revit.DB.ElementId(int(data["cad_id"])))
        beam_type = doc.GetElement(revit.DB.ElementId(int(data["type_id"])))
        level = doc.GetElement(revit.DB.ElementId(int(data["level_id"])))
        
        count = gen.generate_beams(
            cad, data["layer"], beam_type, level,
            data.get("offset", 0.0), data.get("width", 300), data.get("height", 600),
            data.get("read_text", False), data.get("radius", 500)
        )
        print(json.dumps({"status": "success", "count": count}))
    except Exception as ex:
        print(json.dumps({"status": "error", "message": str(ex)}))

if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_headless(sys.argv[1])
    else:
        window = CADtoBeamWindow()
        window.ShowDialog()
