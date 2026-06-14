# -*- coding: utf-8 -*-
"""
IFC-SG Auto Assign v3.0 - DQT
Assigns IFC Export parameters per Family Type based on IFC-SG standards.
Supports LTA Industry Mapping Excel with proper USERDEFINED vs Predefined handling.

Rules:
  Predefined type (no * in Excel col K):
    Export to IFC As = IfcEntity.PredefinedType (e.g. IfcDoor.DOOR)
    IfcObjectType = (empty)
  USERDEFINED (* prefix in Excel col K):
    Export to IFC As = IfcEntity (e.g. IfcDoor)
    IfcObjectType = name without * (e.g. BLASTDOOR)

Copyright (c) 2026 Dang Quoc Truong - DQT. All rights reserved.
"""

__title__ = "Auto Assign\nIFC Class"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = """IFC-SG Auto Assign v3.0 - Assign IFC classes per Family Type.
Supports LTA Industry Mapping Excel with USERDEFINED handling."""

import clr
clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

import System
from System.IO import MemoryStream
from System.Text import Encoding
from System.Windows.Markup import XamlReader
from System.Windows import Window, MessageBox as WPFMessageBox, WindowState
from System.Windows import MessageBoxButton, MessageBoxResult, MessageBoxImage
from System.Windows.Forms import OpenFileDialog, DialogResult as WFDialogResult
from System.Data import DataTable
from System.Windows.Controls import DataGridTextColumn, DataGridLength
from System.Windows.Data import Binding

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter, BuiltInCategory,
    Transaction, ElementId
)
from pyrevit import script, forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# Find T3Lab.extension parent folder dynamically
import os
import io

current_dir = os.path.dirname(__file__)
while current_dir and not current_dir.endswith('T3Lab.extension'):
    parent = os.path.dirname(current_dir)
    if parent == current_dir:
        break
    current_dir = parent

col_map_xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "AutoAssignColMap.xaml")
preview_xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "AutoAssignPreview.xaml")
main_xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "AutoAssignMain.xaml")

# =====================================================================
# Helpers
# =====================================================================

def _eid_int(eid):
    try: return eid.Value
    except: return eid.IntegerValue

def _load_xaml_file(file_path):
    with io.open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    return XamlReader.Load(MemoryStream(Encoding.UTF8.GetBytes(content)))

def _is_userdefined(raw_subtype):
    """Check if Excel subtype is USERDEFINED (has * prefix)."""
    return raw_subtype.startswith("*") if raw_subtype else False

def _clean_subtype(raw_subtype):
    """Remove * prefix from subtype."""
    return raw_subtype[1:] if raw_subtype and raw_subtype.startswith("*") else (raw_subtype or "")

def _build_export_values(entity, raw_subtype):
    """Build IFC export values based on USERDEFINED vs Predefined rules.
    Returns (export_as, object_type, display_class).
    
    Predefined: export_as=IfcDoor.DOOR, object_type="", display=IfcDoor.DOOR
    USERDEFINED: export_as=IfcDoor, object_type=BLASTDOOR, display=IfcDoor.BLASTDOOR
    """
    if not raw_subtype:
        return entity, "", entity
    
    clean = _clean_subtype(raw_subtype)
    display = "{}.{}".format(entity, clean)
    if _is_userdefined(raw_subtype):
        return entity, clean, display
    else:
        return "{}.{}".format(entity, clean), "", display


# =====================================================================
# Revit Category Mapping
# =====================================================================
REVIT_CAT_MAP = {
    "Areas": BuiltInCategory.OST_Areas,
    "Ceilings": BuiltInCategory.OST_Ceilings,
    "Columns": BuiltInCategory.OST_Columns,
    "Communication Devices": BuiltInCategory.OST_CommunicationDevices,
    "Curtain Panels": BuiltInCategory.OST_CurtainWallPanels,
    "Curtain Wall Panels": BuiltInCategory.OST_CurtainWallPanels,
    "Doors": BuiltInCategory.OST_Doors,
    "Duct Accessories": BuiltInCategory.OST_DuctAccessory,
    "Duct Fittings": BuiltInCategory.OST_DuctFitting,
    "Ducts": BuiltInCategory.OST_DuctCurves,
    "Electrical Equipment": BuiltInCategory.OST_ElectricalEquipment,
    "Electrical Fixtures": BuiltInCategory.OST_ElectricalFixtures,
    "Fire Alarm Devices": BuiltInCategory.OST_FireAlarmDevices,
    "Floors": BuiltInCategory.OST_Floors,
    "Furniture": BuiltInCategory.OST_Furniture,
    "Generic Models": BuiltInCategory.OST_GenericModel,
    "Lighting Devices": BuiltInCategory.OST_LightingDevices,
    "Lighting Fixtures": BuiltInCategory.OST_LightingFixtures,
    "Mechanical Equipment": BuiltInCategory.OST_MechanicalEquipment,
    "Parking": BuiltInCategory.OST_Parking,
    "Pipe Accessories": BuiltInCategory.OST_PipeAccessory,
    "Pipe Fittings": BuiltInCategory.OST_PipeFitting,
    "Pipes": BuiltInCategory.OST_PipeCurves,
    "Planting": BuiltInCategory.OST_Planting,
    "Plumbing Fixtures": BuiltInCategory.OST_PlumbingFixtures,
    "Railings": BuiltInCategory.OST_StairsRailing,
    "Ramps": BuiltInCategory.OST_Ramps,
    "Roofs": BuiltInCategory.OST_Roofs,
    "Rooms": BuiltInCategory.OST_Rooms,
    "Specialty Equipment": BuiltInCategory.OST_SpecialityEquipment,
    "Sprinklers": BuiltInCategory.OST_Sprinklers,
    "Stairs": BuiltInCategory.OST_Stairs,
    "Structural Columns": BuiltInCategory.OST_StructuralColumns,
    "Structural Foundations": BuiltInCategory.OST_StructuralFoundation,
    "Structural Framing": BuiltInCategory.OST_StructuralFraming,
    "Walls": BuiltInCategory.OST_Walls,
    "Windows": BuiltInCategory.OST_Windows,
}

# Built-in fallback: (entity, raw_subtype_with_star_if_userdefined)
FALLBACK_MAPPING = {
    BuiltInCategory.OST_Walls: ("IfcWall", ""),
    BuiltInCategory.OST_CurtainWallPanels: ("IfcCurtainWall", ""),
    BuiltInCategory.OST_Floors: ("IfcSlab", ""),
    BuiltInCategory.OST_Roofs: ("IfcRoof", ""),
    BuiltInCategory.OST_Ceilings: ("IfcCovering", "CEILING"),
    BuiltInCategory.OST_Doors: ("IfcDoor", ""),
    BuiltInCategory.OST_Windows: ("IfcWindow", ""),
    BuiltInCategory.OST_Columns: ("IfcColumn", ""),
    BuiltInCategory.OST_StructuralColumns: ("IfcColumn", ""),
    BuiltInCategory.OST_StructuralFraming: ("IfcBeam", ""),
    BuiltInCategory.OST_StructuralFoundation: ("IfcFooting", ""),
    BuiltInCategory.OST_Stairs: ("IfcStair", ""),
    BuiltInCategory.OST_Ramps: ("IfcRamp", ""),
    BuiltInCategory.OST_StairsRailing: ("IfcRailing", ""),
    BuiltInCategory.OST_Rooms: ("IfcSpace", ""),
    BuiltInCategory.OST_Areas: ("IfcSpace", ""),
    BuiltInCategory.OST_GenericModel: ("IfcBuildingElementProxy", ""),
    BuiltInCategory.OST_Furniture: ("IfcFurniture", ""),
    BuiltInCategory.OST_LightingFixtures: ("IfcLightFixture", ""),
    BuiltInCategory.OST_PlumbingFixtures: ("IfcSanitaryTerminal", ""),
    BuiltInCategory.OST_MechanicalEquipment: ("IfcPump", ""),
    BuiltInCategory.OST_SpecialityEquipment: ("IfcTransportElement", ""),
    BuiltInCategory.OST_PipeCurves: ("IfcPipeSegment", ""),
    BuiltInCategory.OST_PipeFitting: ("IfcPipeFitting", ""),
    BuiltInCategory.OST_PipeAccessory: ("IfcValve", ""),
    BuiltInCategory.OST_FlexPipeCurves: ("IfcPipeSegment", ""),
    BuiltInCategory.OST_DuctCurves: ("IfcDuctSegment", ""),
    BuiltInCategory.OST_DuctFitting: ("IfcDuctFitting", ""),
    BuiltInCategory.OST_DuctAccessory: ("IfcDamper", ""),
    BuiltInCategory.OST_Sprinklers: ("IfcSanitaryTerminal", ""),
    BuiltInCategory.OST_Planting: ("IfcGeographicElement", ""),
    BuiltInCategory.OST_Parking: ("IfcBuildingElementProxy", ""),
    BuiltInCategory.OST_FireAlarmDevices: ("IfcAlarm", ""),
    BuiltInCategory.OST_ElectricalEquipment: ("IfcUnitaryControlElement", ""),
    BuiltInCategory.OST_ElectricalFixtures: ("IfcSwitchingDevice", ""),
    BuiltInCategory.OST_LightingDevices: ("IfcSwitchingDevice", ""),
    BuiltInCategory.OST_CommunicationDevices: ("IfcOutlet", ""),
}


# =====================================================================
# IFC Standard Predefined Types (from IFC4 schema)
# =====================================================================
# For each IFC Entity, list of standard PredefinedType values
IFC_PREDEFINED_TYPES = {
    "IfcDoor": ["DOOR", "GATE", "TRAPDOOR"],
    "IfcWindow": ["WINDOW", "SKYLIGHT", "LIGHTDOME"],
    "IfcWall": ["MOVABLE", "PARAPET", "PARTITIONING", "PLUMBINGWALL",
                "SHEAR", "SOLIDWALL", "STANDARD", "POLYGONAL", "ELEMENTEDWALL",
                "RETAININGWALL"],
    "IfcSlab": ["FLOOR", "ROOF", "LANDING", "BASESLAB"],
    "IfcRoof": ["FLAT_ROOF", "SHED_ROOF", "GABLE_ROOF", "HIP_ROOF",
                "HIPPED_GABLE_ROOF", "GAMBREL_ROOF", "MANSARD_ROOF",
                "BARREL_ROOF", "RAINBOW_ROOF", "BUTTERFLY_ROOF",
                "PAVILION_ROOF", "DOME_ROOF", "FREEFORM"],
    "IfcStair": ["STRAIGHT_RUN_STAIR", "TWO_STRAIGHT_RUN_STAIR",
                 "QUARTER_WINDING_STAIR", "QUARTER_TURN_STAIR",
                 "HALF_WINDING_STAIR", "HALF_TURN_STAIR",
                 "TWO_QUARTER_WINDING_STAIR", "TWO_QUARTER_TURN_STAIR",
                 "THREE_QUARTER_WINDING_STAIR", "THREE_QUARTER_TURN_STAIR",
                 "SPIRAL_STAIR", "DOUBLE_RETURN_STAIR",
                 "CURVED_RUN_STAIR", "TWO_CURVED_RUN_STAIR"],
    "IfcRamp": ["STRAIGHT_RUN_RAMP", "TWO_STRAIGHT_RUN_RAMP",
                "QUARTER_TURN_RAMP", "TWO_QUARTER_TURN_RAMP",
                "HALF_TURN_RAMP", "SPIRAL_RAMP"],
    "IfcRailing": ["HANDRAIL", "GUARDRAIL", "BALUSTRADE"],
    "IfcColumn": ["COLUMN", "PILASTER"],
    "IfcBeam": ["BEAM", "JOIST", "HOLLOWCORE", "LINTEL", "SPANDREL", "T_BEAM"],
    "IfcCovering": ["CEILING", "FLOORING", "CLADDING", "ROOFING",
                    "MOLDING", "SKIRTINGBOARD", "INSULATION", "MEMBRANE",
                    "SLEEVING", "WRAPPING"],
    "IfcCurtainWall": ["CURTAINWALL"],
    "IfcFurniture": ["CHAIR", "TABLE", "DESK", "BED", "FILECABINET",
                     "SHELF", "SOFA"],
    "IfcSpace": ["SPACE", "PARKING", "GFA", "INTERNAL", "EXTERNAL"],
    "IfcFooting": ["CAISSON_FOUNDATION", "FOOTING_BEAM", "PAD_FOOTING",
                   "PILE_CAP", "STRIP_FOOTING"],
    "IfcPile": ["BORED", "DRIVEN", "JETGROUTING", "COHESION", "FRICTION",
                "SUPPORT"],
    "IfcSanitaryTerminal": ["BATH", "BIDET", "CISTERN", "SHOWER", "SINK",
                            "SANITARYFOUNTAIN", "TOILETPAN", "URINAL",
                            "WASHHANDBASIN", "WCSEAT"],
    "IfcLightFixture": ["POINTSOURCE", "DIRECTIONSOURCE", "SECURITYLIGHTING"],
    "IfcSwitchingDevice": ["CONTACTOR", "DIMMERSWITCH", "EMERGENCYSTOP",
                           "KEYPAD", "MOMENTARYSWITCH", "SELECTORSWITCH",
                           "STARTER", "SWITCHDISCONNECTOR", "TOGGLESWITCH"],
    "IfcOutlet": ["AUDIOVISUALOUTLET", "COMMUNICATIONSOUTLET", "POWEROUTLET",
                  "DATAOUTLET", "TELEPHONEOUTLET"],
    "IfcAlarm": ["BELL", "BREAKGLASSBUTTON", "LIGHT", "MANUALPULLBOX",
                 "SIREN", "WHISTLE"],
    "IfcPipeSegment": ["CULVERT", "FLEXIBLESEGMENT", "RIGIDSEGMENT", "GUTTER",
                       "SPOOL"],
    "IfcDuctSegment": ["RIGIDSEGMENT", "FLEXIBLESEGMENT"],
    "IfcPipeFitting": ["BEND", "CONNECTOR", "ENTRY", "EXIT", "JUNCTION",
                       "OBSTRUCTION", "TRANSITION"],
    "IfcDuctFitting": ["BEND", "CONNECTOR", "ENTRY", "EXIT", "JUNCTION",
                       "OBSTRUCTION", "TRANSITION"],
    "IfcValve": ["AIRRELEASE", "ANTIVACUUM", "CHANGEOVER", "CHECK",
                 "COMMISSIONING", "DIVERTING", "DRAWOFFCOCK", "DOUBLECHECK",
                 "DOUBLEREGULATING", "FAUCET", "FLUSHING", "GASCOCK",
                 "GASTAP", "ISOLATING", "MIXING", "PRESSUREREDUCING",
                 "PRESSURERELIEF", "REGULATING", "SAFETYCUTOFF", "STEAMTRAP",
                 "STOPCOCK"],
    "IfcDamper": ["BACKDRAFTDAMPER", "BALANCINGDAMPER", "BLASTDAMPER",
                  "CONTROLDAMPER", "FIREDAMPER", "FIRESMOKEDAMPER", "FUMEHOODEXHAUST",
                  "GRAVITYDAMPER", "GRAVITYRELIEFDAMPER", "RELIEFDAMPER",
                  "SMOKEDAMPER"],
    "IfcTransportElement": ["ELEVATOR", "ESCALATOR", "MOVINGWALKWAY",
                            "CRANEWAY", "LIFTINGGEAR"],
    "IfcGeographicElement": ["TERRAIN"],
    "IfcBuildingElementProxy": ["COMPLEX", "ELEMENT", "PARTIAL",
                                "PROVISIONFORVOID", "PROVISIONFORSPACE"],
    "IfcUnitaryControlElement": ["ALARMPANEL", "CONTROLPANEL", "GASDETECTIONPANEL",
                                  "INDICATORPANEL", "MIMICPANEL", "HUMIDISTAT",
                                  "THERMOSTAT", "WEATHERSTATION"],
    "IfcDistributionControlElement": [],
    "IfcDistributionElement": [],
}

# =====================================================================
# Keyword -> Predefined Type Suggestion
# =====================================================================
# Map common Family/Type name keywords to standard IFC predefined types
# Format: { entity: { keyword: predefined_type } }
KEYWORD_HINTS = {
    "IfcDoor": {
        "gate": "GATE", "trapdoor": "TRAPDOOR",
    },
    "IfcWindow": {
        "skylight": "SKYLIGHT", "lightdome": "LIGHTDOME",
    },
    "IfcWall": {
        "parapet": "PARAPET", "partition": "PARTITIONING",
        "movable": "MOVABLE", "shear": "SHEAR",
        "retaining": "RETAININGWALL", "rw-": "RETAININGWALL",
        "polygonal": "POLYGONAL", "plumbing": "PLUMBINGWALL",
    },
    "IfcSlab": {
        "floor": "FLOOR", "roof": "ROOF",
        "landing": "LANDING", "base": "BASESLAB",
    },
    "IfcRoof": {
        "flat": "FLAT_ROOF", "gable": "GABLE_ROOF",
        "hip": "HIP_ROOF", "shed": "SHED_ROOF",
        "barrel": "BARREL_ROOF", "dome": "DOME_ROOF",
        "mansard": "MANSARD_ROOF", "butterfly": "BUTTERFLY_ROOF",
    },
    "IfcStair": {
        "spiral": "SPIRAL_STAIR", "straight": "STRAIGHT_RUN_STAIR",
        "curved": "CURVED_RUN_STAIR", "quarter": "QUARTER_TURN_STAIR",
        "half-turn": "HALF_TURN_STAIR", "double-return": "DOUBLE_RETURN_STAIR",
    },
    "IfcRamp": {
        "spiral": "SPIRAL_RAMP", "straight": "STRAIGHT_RUN_RAMP",
        "quarter": "QUARTER_TURN_RAMP", "half-turn": "HALF_TURN_RAMP",
    },
    "IfcRailing": {
        "handrail": "HANDRAIL", "guardrail": "GUARDRAIL",
        "balustrade": "BALUSTRADE",
    },
    "IfcColumn": {
        "pilaster": "PILASTER",
    },
    "IfcBeam": {
        "joist": "JOIST", "lintel": "LINTEL",
        "hollowcore": "HOLLOWCORE", "spandrel": "SPANDREL",
        "t-beam": "T_BEAM", "tbeam": "T_BEAM",
    },
    "IfcCovering": {
        "ceiling": "CEILING", "flooring": "FLOORING",
        "cladding": "CLADDING", "roofing": "ROOFING",
        "skirting": "SKIRTINGBOARD", "molding": "MOLDING",
        "insulation": "INSULATION", "membrane": "MEMBRANE",
    },
    "IfcFurniture": {
        "chair": "CHAIR", "table": "TABLE", "desk": "DESK",
        "bed": "BED", "shelf": "SHELF", "shelving": "SHELF",
        "sofa": "SOFA", "couch": "SOFA", "filecabinet": "FILECABINET",
        "cabinet": "FILECABINET",
    },
    "IfcSanitaryTerminal": {
        "bath": "BATH", "bidet": "BIDET", "cistern": "CISTERN",
        "shower": "SHOWER", "sink": "SINK", "fountain": "SANITARYFOUNTAIN",
        "toilet": "TOILETPAN", "urinal": "URINAL",
        "basin": "WASHHANDBASIN", "washbasin": "WASHHANDBASIN",
        "wc": "WCSEAT",
    },
    "IfcSpace": {
        "parking": "PARKING", "internal": "INTERNAL",
        "external": "EXTERNAL", "gfa": "GFA",
    },
    "IfcFooting": {
        "pad": "PAD_FOOTING", "strip": "STRIP_FOOTING",
        "pile cap": "PILE_CAP", "pilecap": "PILE_CAP",
        "caisson": "CAISSON_FOUNDATION", "footing beam": "FOOTING_BEAM",
    },
    "IfcTransportElement": {
        "elevator": "ELEVATOR", "lift": "ELEVATOR",
        "escalator": "ESCALATOR", "walkway": "MOVINGWALKWAY",
        "moving walk": "MOVINGWALKWAY",
    },
    "IfcLightFixture": {
        "security": "SECURITYLIGHTING", "directional": "DIRECTIONSOURCE",
        "spotlight": "DIRECTIONSOURCE", "point": "POINTSOURCE",
    },
    "IfcSwitchingDevice": {
        "dimmer": "DIMMERSWITCH", "emergency": "EMERGENCYSTOP",
        "keypad": "KEYPAD", "toggle": "TOGGLESWITCH",
        "selector": "SELECTORSWITCH", "starter": "STARTER",
        "contactor": "CONTACTOR",
    },
    "IfcOutlet": {
        "data": "DATAOUTLET", "communication": "COMMUNICATIONSOUTLET",
        "telephone": "TELEPHONEOUTLET", "phone": "TELEPHONEOUTLET",
        "power": "POWEROUTLET", "audio": "AUDIOVISUALOUTLET",
        "av": "AUDIOVISUALOUTLET",
    },
    "IfcDamper": {
        "fire": "FIREDAMPER", "smoke": "SMOKEDAMPER",
        "fire smoke": "FIRESMOKEDAMPER", "control": "CONTROLDAMPER",
        "balancing": "BALANCINGDAMPER", "relief": "RELIEFDAMPER",
        "blast": "BLASTDAMPER", "backdraft": "BACKDRAFTDAMPER",
    },
    "IfcValve": {
        "check": "CHECK", "stop": "STOPCOCK", "gate": "STOPCOCK",
        "isolating": "ISOLATING", "mixing": "MIXING",
        "regulating": "REGULATING", "safety": "SAFETYCUTOFF",
        "relief": "PRESSURERELIEF", "faucet": "FAUCET", "tap": "FAUCET",
    },
    "IfcPipeFitting": {
        "bend": "BEND", "elbow": "BEND", "tee": "JUNCTION",
        "junction": "JUNCTION", "transition": "TRANSITION",
        "connector": "CONNECTOR", "reducer": "TRANSITION",
    },
    "IfcDuctFitting": {
        "bend": "BEND", "elbow": "BEND", "tee": "JUNCTION",
        "junction": "JUNCTION", "transition": "TRANSITION",
        "connector": "CONNECTOR", "reducer": "TRANSITION",
    },
}


# =====================================================================
# Excel Reader (zipfile + XML)
# =====================================================================
import zipfile
try:
    from xml.etree import ElementTree as ET
except:
    import xml.etree.ElementTree as ET

_NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def _col_to_idx(c):
    r = 0
    for ch in c.upper():
        r = r * 26 + (ord(ch) - 64)
    return r - 1

def _parse_ref(ref):
    c, r = "", ""
    for ch in ref:
        if ch.isalpha(): c += ch
        else: r += ch
    return _col_to_idx(c), int(r) - 1

def _read_xlsx(filepath):
    res = {"sheets": [], "sheet_data": {}}
    try: zf = zipfile.ZipFile(filepath, 'r')
    except: return None
    try:
        ss = []
        if "xl/sharedStrings.xml" in zf.namelist():
            for si in ET.fromstring(zf.read("xl/sharedStrings.xml")).findall("s:si", _NS):
                ss.append("".join(t.text or "" for t in si.iter(
                    "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")))
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        sinfo = [(s.get("name",""), s.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",""))
            for s in wb.findall(".//s:sheet", _NS)]
        rels = {r.get("Id",""): "xl/"+r.get("Target","")
            for r in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))}
        for sn, rid in sinfo:
            res["sheets"].append(sn)
            sp = rels.get(rid, "xl/worksheets/sheet{}.xml".format(len(res["sheets"])))
            if sp not in zf.namelist():
                res["sheet_data"][sn] = []; continue
            rd = {}; mc = mr = 0
            for row in ET.fromstring(zf.read(sp)).findall(".//s:sheetData/s:row", _NS):
                for cel in row.findall("s:c", _NS):
                    ref = cel.get("r","")
                    if not ref: continue
                    ci, ri = _parse_ref(ref)
                    if ci > mc: mc = ci
                    if ri > mr: mr = ri
                    ct = cel.get("t","")
                    v = cel.find("s:v", _NS)
                    val = ""
                    if v is not None and v.text:
                        if ct == "s":
                            try: val = ss[int(v.text)]
                            except: val = v.text
                        else: val = v.text
                    elif ct == "inlineStr":
                        isel = cel.find("s:is", _NS)
                        if isel is not None:
                            val = "".join(t.text or "" for t in isel.iter(
                                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
                    rd.setdefault(ri, {})[ci] = val
            res["sheet_data"][sn] = [[rd.get(ri,{}).get(ci,"")
                for ci in range(min(mc+1,30))] for ri in range(mr+1)]
    except: return None
    finally: zf.close()
    return res

def read_excel_headers(fp):
    x = _read_xlsx(fp)
    if not x: return None
    return {"sheets": x["sheets"], "headers": {
        s: [str(v).strip().replace("\n"," ") if v else "(empty)" for v in (x["sheet_data"].get(s,[[]])[0])]
        for s in x["sheets"]}}

def load_mapping_from_excel(fp, sheet, c_comp, c_ent, c_sub=0, c_rev=0):
    """Load mapping preserving raw_subtype (with * for USERDEFINED)."""
    x = _read_xlsx(fp)
    if not x: return None, []
    grid = x["sheet_data"].get(sheet, [])
    if len(grid) < 2: return {}, []
    raw_rows, cat_map, seen = [], {}, set()
    ic, ie, isu, ir = c_comp-1, c_ent-1, (c_sub-1 if c_sub>0 else -1), (c_rev-1 if c_rev>0 else -1)

    for r in range(1, len(grid)):
        row = grid[r]
        def cell(i): return row[i] if 0<=i<len(row) else ""
        comp = str(cell(ic)).strip()
        if not comp: continue
        ent = str(cell(ie)).strip()
        if not ent or ent.lower() in ("nan","n.a","n.a.",""): continue

        raw_sub = ""
        if isu >= 0:
            s = str(cell(isu)).strip()
            if s.lower() not in ("nan","n.a","n.a.",""):
                parts = [p.strip() for p in s.split(",") if p.strip()]
                if parts: raw_sub = parts[0]  # Keep * prefix!

        rev_cat = ""
        if ir >= 0:
            rev_cat = str(cell(ir)).strip()
            if rev_cat.lower() in ("nan","n.a","n.a.",""): rev_cat = ""

        export_as, obj_type, display = _build_export_values(ent, raw_sub)

        raw_rows.append({"component": comp, "entity": ent, "raw_subtype": raw_sub,
            "export_as": export_as, "object_type": obj_type, "display": display, "revit_cat": rev_cat})

        if rev_cat:
            matched = []
            for cn in REVIT_CAT_MAP:
                if cn.lower() == rev_cat.lower(): matched.append(cn); break
            if not matched:
                for cn in REVIT_CAT_MAP:
                    if cn.lower() in rev_cat.lower() or rev_cat.lower() in cn.lower():
                        matched.append(cn)
            for mc in matched:
                dk = (mc, comp, export_as)
                if dk in seen: continue
                seen.add(dk)
                cat_map.setdefault(mc, []).append({
                    "component": comp, "entity": ent, "raw_subtype": raw_sub,
                    "export_as": export_as, "object_type": obj_type, "display": display,
                })
    return cat_map, raw_rows


# =====================================================================
# IFC Parameter Get/Set
# =====================================================================

def _get_param(element, bip, *names):
    """Try BuiltInParameter first, then LookupParameter by names."""
    if bip:
        try:
            p = element.get_Parameter(bip)
            if p: return p
        except: pass
    for n in names:
        try:
            p = element.LookupParameter(n)
            if p: return p
        except: pass
    return None

def get_current_ifc(element):
    """Get current Export to IFC As value."""
    try:
        p = _get_param(element, BuiltInParameter.IFC_EXPORT_ELEMENT_AS,
                        "Export to IFC As", "IfcExportAs")
        if p and p.HasValue:
            v = p.AsString()
            if v and v.strip(): return v.strip()
    except: pass
    return ""

def get_current_objtype(element):
    """Get current IfcObjectType value."""
    try:
        p = _get_param(element, BuiltInParameter.IFC_EXPORT_ELEMENT_TYPE_AS,
                        "Type IfcObjectType[Type]", "IfcObjectType")
        if p and p.HasValue:
            v = p.AsString()
            if v and v.strip(): return v.strip()
    except: pass
    return ""

def set_ifc_values(element, export_as, object_type):
    """Set both Export to IFC As and IfcObjectType. Returns True on success."""
    ok = False
    try:
        p1 = _get_param(element, BuiltInParameter.IFC_EXPORT_ELEMENT_AS,
                         "Export to IFC As", "IfcExportAs")
        if p1 and not p1.IsReadOnly:
            p1.Set(export_as)
            ok = True

        # Set IfcObjectType (for USERDEFINED)
        p2 = _get_param(element, BuiltInParameter.IFC_EXPORT_ELEMENT_TYPE_AS,
                         "Type IfcObjectType[Type]", "IfcObjectType")
        if p2 and not p2.IsReadOnly:
            p2.Set(object_type or "")
    except: pass
    return ok

def collect_elements(categories):
    result = {}
    for bic in categories:
        try:
            elems = FilteredElementCollector(doc).OfCategory(bic) \
                        .WhereElementIsNotElementType().ToElements()
            if elems and len(elems) > 0:
                result[bic] = list(elems)
        except: continue
    return result


# =====================================================================
# Column Mapping Dialog
# =====================================================================
def show_column_mapping_dialog(excel_info):
    win = _load_xaml_file(col_map_xaml_path)
    result = {"ok": False}
    cs, cc, ce, csu, cr = (win.FindName(n) for n in
        ["cmbSheet","cmbComponent","cmbEntity","cmbSubtype","cmbRevit"])
    combos = [cc,ce,csu,cr]
    DET = {"component":["identified component","component"],
           "entity":["ifc4","ifc entity","ifc entities","entity"],
           "subtype":["sub type","subtype","predefined","ifc sub"],
           "revit":["revit representation","revit category","suggested revit"]}
    dc = {"component":cc,"entity":ce,"subtype":csu,"revit":cr}
    def pop(sn):
        hdrs = excel_info["headers"].get(sn,[])
        items = ["(not mapped)"] + ["{}: {}".format(i+1,h) for i,h in enumerate(hdrs)]
        for c in combos:
            c.Items.Clear()
            for it in items: c.Items.Add(it)
            c.SelectedIndex = 0
        for f,kws in DET.items():
            for idx,h in enumerate(hdrs):
                if any(k in h.lower().replace("\n"," ") for k in kws):
                    dc[f].SelectedIndex = idx+1; break
    for s in excel_info["sheets"]: cs.Items.Add(s)
    di = 0
    for i,s in enumerate(excel_info["sheets"]):
        if "mapping" in s.lower() or "pilot" in s.lower(): di=i; break
    cs.SelectedIndex = di; pop(excel_info["sheets"][di])
    def osc(s,a):
        if cs.SelectedIndex>=0: pop(excel_info["sheets"][cs.SelectedIndex])
    cs.SelectionChanged += osc
    def gc(c): return c.SelectedIndex if c.SelectedIndex>0 else 0
    def ok(s,a):
        if gc(cc)==0 or gc(ce)==0:
            WPFMessageBox.Show("Component + Entity required.","",MessageBoxButton.OK,MessageBoxImage.Warning); return
        result.update({"ok":True,"sheet":excel_info["sheets"][cs.SelectedIndex],
            "component":gc(cc),"entity":gc(ce),"subtype":gc(csu),"revit":gc(cr)})
        win.Close()
    win.FindName("btnOK").Click += ok
    win.FindName("btnCancel").Click += lambda s,a: win.Close()
    win.ShowDialog(); return result


# =====================================================================
# Family Type helpers
# =====================================================================
def _get_type_name(el):
    try:
        if hasattr(el,'Symbol') and el.Symbol is not None:
            sym = el.Symbol
            fn = ""
            try: fn = sym.FamilyName
            except:
                p = sym.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                if p and p.HasValue: fn = p.AsString()
            tn = ""
            try: tn = DB.Element.Name.__get__(sym)
            except:
                p = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if p and p.HasValue: tn = p.AsString()
            return "{} : {}".format(fn, tn) if fn else tn
        tp = el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if tp and tp.HasValue:
            te = doc.GetElement(tp.AsElementId())
            if te:
                try: return DB.Element.Name.__get__(te)
                except: pass
    except: pass
    return "(unknown)"

def _match_comp(type_name, comp_list):
    """Match Family Type name to Excel Component by keyword.
    Requires reasonably strong match to avoid false positives.
    Returns matched comp dict or None.
    """
    td = type_name.lower()
    best, bscore = None, 0
    for c in comp_list:
        cn = c["component"].lower().strip()
        if not cn:
            continue
        sc = 0
        # Strongest signal: full component name appears in family type name
        if cn in td:
            sc += 100
        else:
            # Match on individual keywords (>=3 chars to avoid noise)
            kws = [w for w in cn.split() if len(w) >= 3]
            for k in kws:
                if k in td:
                    sc += len(k) * 2
            # Boost if all keywords matched
            if kws and all(k in td for k in kws):
                sc += 30
        if sc > bscore:
            bscore, best = sc, c
    # Require a meaningful score; substring noise needs ≥6 to count
    return best if bscore >= 6 else None


def _suggest_predefined(type_name, entity):
    """Suggest a standard IFC predefined type from family type name keywords.
    Returns predefined_type string or "" if nothing matches.
    """
    if not entity or entity not in KEYWORD_HINTS:
        return ""
    td = type_name.lower()
    hints = KEYWORD_HINTS[entity]
    # Check longest keywords first (more specific)
    sorted_keys = sorted(hints.keys(), key=lambda k: -len(k))
    for kw in sorted_keys:
        if kw in td:
            return hints[kw]
    return ""


def _get_category_entity(bic, mapping, src):
    """Get the most common IFC Entity for a category.
    Used for Unmatched fallback - so we still know the entity.
    """
    if src == "Excel" and mapping:
        cat = ""
        try:
            c = DB.Category.GetCategory(doc, bic)
            if c: cat = c.Name
        except: pass
        for nk, dl in mapping.items():
            if nk.lower() == cat.lower() and dl:
                return dl[0]["entity"]
        for nk, dl in mapping.items():
            if (nk.lower() in cat.lower() or cat.lower() in nk.lower()) and dl:
                return dl[0]["entity"]
    fb = FALLBACK_MAPPING.get(bic)
    return fb[0] if fb else "IfcBuildingElementProxy"


# =====================================================================
# Build Mapping Rows (Family Type level) - 3 states
# =====================================================================
def _build_rows(elems_by_cat, mapping, src):
    """Build Family Type rows with 3-state classification:
    - Excel Match: family type matches Excel component name
    - Keyword Suggest: no Excel match, but standard IFC predefined type from name
    - Unmatched: cannot determine - assign only entity, no subtype (safe)
    """
    rows = []
    for bic, elems in elems_by_cat.items():
        cat = ""
        try:
            c = DB.Category.GetCategory(doc, bic)
            if c: cat = c.Name
        except: cat = str(bic)

        # Find Excel comp_list for this category
        cl = None
        if src == "Excel":
            for nk, dl in mapping.items():
                if nk.lower() == cat.lower(): cl = dl; break
            if cl is None:
                for nk, dl in mapping.items():
                    if nk.lower() in cat.lower() or cat.lower() in nk.lower():
                        cl = dl; break

        # Determine the default IFC entity for this category (used for fallback)
        default_entity = _get_category_entity(bic, mapping, src)

        # Group by family type
        tg = {}
        for el in elems:
            tn = _get_type_name(el)
            tg.setdefault(tn, []).append(el)

        for tn in sorted(tg):
            tel = tg[tn]
            # Current IFC values
            cv = {}
            for e in tel:
                cur = get_current_ifc(e)
                key = cur if cur else "(none)"
                cv[key] = cv.get(key, 0) + 1
            if len(cv) == 1:
                cur_d = list(cv.keys())[0]
            else:
                sv = sorted(cv.items(), key=lambda x: -x[1])
                cur_d = "{} (+{})".format(sv[0][0], len(cv) - 1)

            # State 1: Excel Match
            m = None
            if src == "Excel" and cl:
                m = _match_comp(tn, cl)

            if m:
                rows.append({
                    "category": cat, "family_type": tn, "count": len(tel),
                    "current_ifc": cur_d, "component": m["component"],
                    "entity": m["entity"], "raw_subtype": m["raw_subtype"],
                    "export_as": m["export_as"], "object_type": m["object_type"],
                    "display": m["display"], "source": "Excel Match",
                    "bic": bic, "elements": tel, "all_components": cl or [],
                    "default_entity": default_entity,
                })
                continue

            # State 2: Keyword Suggest (standard IFC predefined type)
            suggested_pt = _suggest_predefined(tn, default_entity)
            if suggested_pt:
                # Build Predefined assignment (no *)
                ea, ot, disp = _build_export_values(default_entity, suggested_pt)
                rows.append({
                    "category": cat, "family_type": tn, "count": len(tel),
                    "current_ifc": cur_d, "component": "(suggested)",
                    "entity": default_entity, "raw_subtype": suggested_pt,
                    "export_as": ea, "object_type": ot, "display": disp,
                    "source": "Keyword Suggest", "bic": bic, "elements": tel,
                    "all_components": cl or [],
                    "default_entity": default_entity,
                })
                continue

            # State 3: Unmatched - SAFE assignment (entity only, no subtype)
            ea_safe, ot_safe, disp_safe = _build_export_values(default_entity, "")
            rows.append({
                "category": cat, "family_type": tn, "count": len(tel),
                "current_ifc": cur_d, "component": "(none)",
                "entity": default_entity, "raw_subtype": "",
                "export_as": ea_safe, "object_type": ot_safe, "display": disp_safe,
                "source": "Unmatched", "bic": bic, "elements": tel,
                "all_components": cl or [],
                "default_entity": default_entity,
            })

    rows.sort(key=lambda x: (x["category"].lower(), x["family_type"].lower()))
    return rows


# =====================================================================
# Preview Dialog
# =====================================================================
def show_preview(elems_by_cat, mapping, src):
    win = _load_xaml(PREVIEW_XAML)
    result = {"ok":False,"mode":"assign_all","final_mapping":[]}
    dg = win.FindName("dgMapping")
    rbAll = win.FindName("rbAll")
    txtS = win.FindName("txtSearch")
    cmbC = win.FindName("cmbFilterCat")
    cmbSrc = win.FindName("cmbFilterSource")
    btnCh = win.FindName("btnChangeIFC")

    mrows = _build_rows(elems_by_cat, mapping, src)
    if not mrows: return result

    te = sum(r["count"] for r in mrows)
    tc = len(set(r["category"] for r in mrows))
    um = sum(1 for r in mrows if r["source"] == "Unmatched")
    win.FindName("txtSummary").Text = "Family Type mapping | Select rows + Change IFC for unmatched"
    win.FindName("txtCatCount").Text = str(tc)
    win.FindName("txtElemCount").Text = str(te)
    win.FindName("txtUnmatched").Text = str(um)
    win.FindName("txtSource").Text = src

    # DataTable with Current IFC column
    dt = DataTable()
    for col in ["category","family_type","count","current_ifc","component","entity",
                "subtype","object_type","new_ifc","source"]:
        dt.Columns.Add(col)

    for r in mrows:
        dr = dt.NewRow()
        dr["category"] = r["category"]
        dr["family_type"] = r["family_type"]
        dr["count"] = str(r["count"])
        dr["current_ifc"] = r["current_ifc"]
        dr["component"] = r["component"]
        dr["entity"] = r["entity"]
        dr["subtype"] = _clean_subtype(r["raw_subtype"]) if r["raw_subtype"] else ""
        dr["object_type"] = r["object_type"]
        dr["new_ifc"] = r["display"]
        dr["source"] = r["source"]
        dt.Rows.Add(dr)

    # Columns
    cols = [("Category","category",120),("Family / Type","family_type",200),
            ("Elem","count",45),("Current IFC","current_ifc",140),
            ("Component","component",120),("IFC Entity","entity",140),
            ("Subtype","subtype",100),("ObjectType","object_type",100),
            ("New IFC","new_ifc",180),("Source","source",85)]
    for h,bp,w in cols:
        c = DataGridTextColumn()
        c.Header = h; c.Binding = Binding(bp); c.Width = DataGridLength(w)
        dg.Columns.Add(c)
    dg.ItemsSource = dt.DefaultView

    # Filters
    cats = sorted(set(r["category"] for r in mrows))
    cmbC.Items.Add("(All)")
    for c in cats: cmbC.Items.Add(c)
    cmbC.SelectedIndex = 0
    srcs = sorted(set(r["source"] for r in mrows))
    cmbSrc.Items.Add("(All)")
    for s in srcs: cmbSrc.Items.Add(s)
    cmbSrc.SelectedIndex = 0

    def filt():
        se = txtS.Text.strip().lower() if txtS.Text else ""
        ca = str(cmbC.SelectedItem) if cmbC.SelectedIndex>0 else ""
        so = str(cmbSrc.SelectedItem) if cmbSrc.SelectedIndex>0 else ""
        parts = []
        if se:
            lp = ["{} LIKE '%{}%'".format(c, se.replace("'","''"))
                  for c in ["category","family_type","component","new_ifc"]]
            parts.append("({})".format(" OR ".join(lp)))
        if ca: parts.append("category='{}'".format(ca.replace("'","''")))
        if so: parts.append("source='{}'".format(so.replace("'","''")))
        try: dt.DefaultView.RowFilter = " AND ".join(parts) if parts else ""
        except: dt.DefaultView.RowFilter = ""

    txtS.TextChanged += lambda s,a: filt()
    cmbC.SelectionChanged += lambda s,a: filt()
    cmbSrc.SelectionChanged += lambda s,a: filt()

    # Change IFC (multi-select)
    def _find_idx(cat, ft):
        for i,r in enumerate(mrows):
            if r["category"]==cat and r["family_type"]==ft: return i
        return -1

    def _upd(i, sc):
        for k in ["component","entity","raw_subtype","export_as","object_type","display"]:
            mrows[i][k] = sc[k]
        mrows[i]["source"] = "Manual"
        dt.Rows[i]["component"] = sc["component"]
        dt.Rows[i]["entity"] = sc["entity"]
        dt.Rows[i]["subtype"] = _clean_subtype(sc["raw_subtype"])
        dt.Rows[i]["object_type"] = sc["object_type"]
        dt.Rows[i]["new_ifc"] = sc["display"]
        dt.Rows[i]["source"] = "Manual"

    def on_change(s, a):
        sel = []
        try:
            for it in dg.SelectedItems: sel.append(it)
        except: pass
        if not sel:
            WPFMessageBox.Show("Select rows first (Ctrl/Shift+Click).","",
                MessageBoxButton.OK, MessageBoxImage.Information); return
        idxs = [_find_idx(it["category"], it["family_type"]) for it in sel]
        idxs = [i for i in idxs if i>=0]
        if not idxs: return

        # Get IFC entity from first selected row
        fc = mrows[idxs[0]]["category"]
        entity = mrows[idxs[0]].get("default_entity", "") or mrows[idxs[0]].get("entity", "")

        # Build options from 3 sources:
        # 1) Standard IFC Predefined Types for this entity
        # 2) USERDEFINED options from Excel mapping
        # 3) Keep as base entity (no subtype)
        options_data = []  # list of (display_string, comp_dict)

        # Source 1: Standard IFC Predefined Types
        std_types = IFC_PREDEFINED_TYPES.get(entity, [])
        for pt in std_types:
            ea, ot, disp = _build_export_values(entity, pt)
            options_data.append((
                "[Predefined] {}.{}".format(entity, pt),
                {"component": "(predefined)", "entity": entity,
                 "raw_subtype": pt, "export_as": ea, "object_type": ot,
                 "display": disp}
            ))

        # Source 2: Excel USERDEFINED options for this category
        excel_cl = None
        for nk, dl in (mapping or {}).items():
            if nk.lower() == fc.lower() and dl:
                excel_cl = dl; break
        if not excel_cl:
            for nk, dl in (mapping or {}).items():
                if (nk.lower() in fc.lower() or fc.lower() in nk.lower()) and dl:
                    excel_cl = dl; break

        if excel_cl:
            for c in excel_cl:
                # Only USERDEFINED items add value here (Predefined already covered by source 1)
                if _is_userdefined(c.get("raw_subtype", "")):
                    options_data.append((
                        "[USERDEFINED] {} -> {}".format(c["component"], c["display"]),
                        c
                    ))

        # Source 3: Keep base entity
        ea_base, ot_base, disp_base = _build_export_values(entity, "")
        options_data.append((
            "[Base Entity] {} (no subtype)".format(entity),
            {"component": "(none)", "entity": entity,
             "raw_subtype": "", "export_as": ea_base,
             "object_type": ot_base, "display": disp_base}
        ))

        if not options_data:
            WPFMessageBox.Show("No IFC options available.","",
                MessageBoxButton.OK, MessageBoxImage.Information); return

        opts_str = [d[0] for d in options_data]
        pk = forms.CommandSwitchWindow.show(opts_str,
            message="Change {} rows ({}):".format(len(idxs), fc))
        if pk:
            sc = options_data[opts_str.index(pk)][1]
            for i in idxs: _upd(i, sc)
            win.FindName("txtUnmatched").Text = str(
                sum(1 for r in mrows if r["source"] == "Unmatched"))

    btnCh.Click += on_change

    def on_apply(s, a):
        # Warn if there are still Unmatched rows
        unm = sum(1 for r in mrows if r["source"] == "Unmatched")
        if unm > 0:
            res = WPFMessageBox.Show(
                "There are still {} Unmatched rows.\n\n".format(unm) +
                "Unmatched rows will be assigned only the IFC Entity (no subtype).\n"
                "Recommended: Filter Source = 'Unmatched' and use 'Change IFC' to review them.\n\n"
                "Do you want to continue anyway?",
                "Unmatched Rows", MessageBoxButton.YesNo, MessageBoxImage.Warning)
            if res != MessageBoxResult.Yes:
                return
        result["ok"] = True
        result["mode"] = "assign_all" if rbAll.IsChecked else "assign_empty"
        result["final_mapping"] = [{"export_as":r["export_as"],"object_type":r["object_type"],
            "category":r["category"],"family_type":r["family_type"],"elements":r["elements"]}
            for r in mrows]
        win.Close()
    win.FindName("btnApply").Click += on_apply
    win.FindName("btnCancel").Click += lambda s,a: win.Close()
    win.ShowDialog(); return result


# =====================================================================
# Apply Assignment
# =====================================================================
def apply_assignment(final_mapping, mode):
    stats = {"total":0,"success":0,"skipped":0,"failed":0,"by_category":{}}
    errors = []
    t = Transaction(doc, "DQT - Auto Assign IFC v3.0")
    t.Start()
    try:
        for fm in final_mapping:
            ea, ot = fm["export_as"], fm["object_type"]
            if not ea: continue
            lbl = "{} > {}".format(fm["category"], fm["family_type"])
            ok = sk = fl = 0
            for el in fm.get("elements",[]):
                stats["total"] += 1
                try:
                    if mode=="assign_empty" and get_current_ifc(el):
                        stats["skipped"]+=1; sk+=1; continue
                    if set_ifc_values(el, ea, ot):
                        stats["success"]+=1; ok+=1
                    else:
                        stats["failed"]+=1; fl+=1
                        errors.append("{} [{}] - read-only".format(lbl, _eid_int(el.Id)))
                except Exception as e:
                    stats["failed"]+=1; fl+=1
                    errors.append("{} [{}] - {}".format(lbl, _eid_int(el.Id), e))
            if ok+sk+fl>0:
                stats["by_category"][lbl] = {"success":ok,"skipped":sk,"failed":fl}
        t.Commit()
    except:
        if t.HasStarted() and not t.HasEnded(): t.RollBack()
        raise
    return stats, errors


# =====================================================================
# Main Dialog
# =====================================================================
class MainWindow(object):
    def __init__(self):
        self.win = _load_xaml_file(main_xaml_path)
        self.st = self.win.FindName("txtStatus")
        self.win.FindName("btnExcel").MouseLeftButtonUp += self.on_excel
        self.win.FindName("btnBuiltin").MouseLeftButtonUp += self.on_builtin

    def on_excel(self, s, a):
        dlg = OpenFileDialog()
        dlg.Title = "Select LTA Industry Mapping Excel"
        dlg.Filter = "Excel (*.xlsx)|*.xlsx|All|*.*"
        if dlg.ShowDialog() != WFDialogResult.OK: return
        self.st.Text = "Reading..."
        ei = read_excel_headers(dlg.FileName)
        if not ei:
            WPFMessageBox.Show("Cannot read Excel.","",MessageBoxButton.OK,MessageBoxImage.Error)
            self.st.Text = ""; return
        cr = show_column_mapping_dialog(ei)
        if not cr.get("ok"): self.st.Text = "Cancelled."; return
        self.st.Text = "Loading..."
        cm, rr = load_mapping_from_excel(dlg.FileName, cr["sheet"],
            cr["component"], cr["entity"], cr.get("subtype",0), cr.get("revit",0))
        if cm is None:
            WPFMessageBox.Show("Failed.","",MessageBoxButton.OK,MessageBoxImage.Error)
            self.st.Text = ""; return
        self.st.Text = "{} rows, {} cats.".format(len(rr), len(cm))
        self._run(cm, "Excel")

    def on_builtin(self, s, a):
        self._run(FALLBACK_MAPPING, "Built-in")

    def _run(self, mapping, src):
        cats = set()
        if src=="Excel":
            for cn in mapping:
                b = REVIT_CAT_MAP.get(cn)
                if b: cats.add(b)
        for b in FALLBACK_MAPPING: cats.add(b)
        self.st.Text = "Collecting..."
        ebc = collect_elements(list(cats))
        if not ebc:
            WPFMessageBox.Show("No elements.","",MessageBoxButton.OK,MessageBoxImage.Information)
            self.st.Text = ""; return
        self.win.Hide()
        pr = show_preview(ebc, mapping, src)
        if not pr.get("ok"):
            self.win.Show(); self.st.Text = "Cancelled."; return
        stats, errs = apply_assignment(pr["final_mapping"], pr["mode"])
        msg = "Done!\n\nTotal: {}\nSuccess: {}\nSkipped: {}\nFailed: {}\n".format(
            stats["total"], stats["success"], stats["skipped"], stats["failed"])
        if stats["by_category"]:
            msg += "\nBy Category:\n"
            for c, cs in sorted(stats["by_category"].items()):
                if cs["success"]>0:
                    msg += "  {} - {} OK".format(c, cs["success"])
                    if cs["skipped"]: msg += ", {} skip".format(cs["skipped"])
                    if cs["failed"]: msg += ", {} fail".format(cs["failed"])
                    msg += "\n"
        if errs:
            msg += "\nErrors ({}):\n".format(len(errs))
            for e in errs[:10]: msg += "  {}\n".format(e)
        ic = MessageBoxImage.Information if stats["failed"]==0 else MessageBoxImage.Warning
        WPFMessageBox.Show(msg, "IFC-SG Auto Assign", MessageBoxButton.OK, ic)
        self.win.Close()

    def show(self): self.win.ShowDialog()

if __name__ == "__main__":
    try: MainWindow().show()
    except:
        import traceback
        output.print_md("## Error\n```\n{}\n```".format(traceback.format_exc()))