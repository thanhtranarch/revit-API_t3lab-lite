# -*- coding: utf-8 -*-
"""
JSON to Family

Parametric Revit Family Generator (Metric Edition).
Parses a JSON schema to generate parametric Revit Families (Furniture, Casework, etc.).

SUPPORTED GEOMETRY TYPES:
  - Extrusion    : Profile extruded between two Z heights
  - Sweep        : 2D profile swept along a 3D path (CurveArray)
  - Revolve      : Profile rotated around an axis (in radians)
  - Blend        : Transition between two profiles at different heights
  - SweptBlend   : Blend profile swept along a single 3D path curve
  - ModelText    : Raised/embossed 3D text on a sketch plane
  - Void         : Any of the above with "is_solid": false, using "cuts" to cut parent solids

KNOWN LIMITATIONS:
  - Profile loops MUST be drawn continuously end-to-end (no gaps/self-intersections).
  - Blend: bottom and top profiles MUST have the same number of curve segments.
  - SweptBlend: path must be a single curve (use first segment if multiple provided).
  - Alignments may fail silently on complex curved geometry.
  - Sweeps: profile_2d drawn in 2D XY plane; API rotates it onto path automatically.
  - ModelText: requires an existing ModelTextType in the family template.

Credits: Based on JSONToFamily by Jonathan Bourne (manicooller/jonotools)

Author: T3Lab
"""

from __future__ import unicode_literals

__title__   = "JSON to Family"
__author__  = "T3Lab"
__version__ = "1.0.0"

# IMPORTS
import json
import clr
import os
import sys

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')
from System.Windows import WindowState, Clipboard
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir       = os.path.join(extension_dir, 'lib')

doc    = revit.doc
logger = script.get_logger()
output = script.get_output()

# =========================================================================
# METRIC CONVERSION
# =========================================================================
IS_METRIC = True
SCL = (1.0 / 304.8) if IS_METRIC else 1.0


# =========================================================================
# UI DIALOG
# =========================================================================

_XAML_PATH = os.path.join(lib_dir, 'GUI', 'Tools', 'JSONtoFamily.xaml')


class JsonInputDialog(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML_PATH)
        self.json_data = None
        logger.debug("JSONtoFamily dialog initialised from: {}".format(_XAML_PATH))

        # Load logo
        try:
            logo_path = os.path.join(lib_dir, 'GUI', 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                self.Icon = bitmap
                self.logo_image.Source = bitmap
        except Exception:
            pass

    def create_clicked(self, sender, args):
        self.json_data = self.json_tb.Text
        self.Close()

    def cancel_clicked(self, sender, args):
        self.Close()

    def copy_prompt_clicked(self, sender, e):
        try:
            script_dir = os.path.dirname(__file__)
            prompt_path = os.path.join(script_dir, "SYSTEM PROMPT.md")
            if not os.path.exists(prompt_path):
                forms.alert("SYSTEM PROMPT.md not found.", title="Copy Prompt")
                return
            with open(prompt_path, "r") as f:
                content = f.read()
            Clipboard.SetText(content)
            forms.alert("System Prompt copied to clipboard!", title="Copy Prompt")
        except Exception as ex:
            forms.alert("Failed to copy prompt:\n\n{}".format(ex), title="Copy Prompt Error")

    def help_button_clicked(self, sender, e):
        forms.alert(
            "JSON to Family - T3Lab\n\n"
            "Paste a JSON schema to generate a parametric Revit family.\n\n"
            "Supported geometry types:\n"
            "  Extrusion, Sweep, Revolve, Blend, SweptBlend, ModelText, Voids\n\n"
            "Run this tool inside an open Family Document (.rfa).",
            title="JSON to Family Help"
        )

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


# =========================================================================
# GEOMETRY HELPERS
# =========================================================================

def to_xyz(pt_list, scale=SCL):
    return XYZ(pt_list[0] * scale, pt_list[1] * scale, pt_list[2] * scale)

def to_vec(pt_list):
    return XYZ(pt_list[0], pt_list[1], pt_list[2])

def to_xyz_2d(pt_list, scale=SCL):
    return XYZ(pt_list[0] * scale, pt_list[1] * scale, 0.0)

def project_to_plane(pt, plane):
    vector = pt - plane.Origin
    distance = vector.DotProduct(plane.Normal)
    return pt - (plane.Normal * distance)

def create_curve_from_json(segment_data, sp_plane=None, is_2d_sweep=False, blend_offset=None):
    if is_2d_sweep:
        p1 = to_xyz_2d(segment_data["p1"])
        p2 = to_xyz_2d(segment_data["p2"])
    else:
        p1 = to_xyz(segment_data["p1"])
        p2 = to_xyz(segment_data["p2"])
        if sp_plane:
            p1 = project_to_plane(p1, sp_plane)
            p2 = project_to_plane(p2, sp_plane)
        if sp_plane and blend_offset is not None:
            offset_vec = sp_plane.Normal * blend_offset
            p1 += offset_vec
            p2 += offset_vec

    if segment_data.get("is_arc", False):
        if is_2d_sweep:
            p3 = to_xyz_2d(segment_data["p3"])
        else:
            p3 = to_xyz(segment_data["p3"])
            if sp_plane:
                p3 = project_to_plane(p3, sp_plane)
            if sp_plane and blend_offset is not None:
                p3 += sp_plane.Normal * blend_offset
        return Arc.Create(p1, p2, p3)
    else:
        return Line.CreateBound(p1, p2)

def build_sweep_profile(segments, is_2d=True, sp_plane=None):
    """Build a SweepProfile from a list of curve segments."""
    loop = CurveArray()
    for seg in segments:
        loop.Append(create_curve_from_json(seg, sp_plane=sp_plane, is_2d_sweep=is_2d))
    arr = CurveArrArray()
    arr.Append(loop)
    return doc.Application.Create.NewCurveLoopsProfile(arr)

def get_model_text_type(name=None):
    """Return a ModelTextType ElementId by name, or the first available."""
    collector = FilteredElementCollector(doc).OfClass(ModelTextType)
    for tt in collector:
        if name is None or tt.Name == name:
            return tt.Id
    return None


# =========================================================================
# MAIN GENERATION ENGINE
# =========================================================================

def generate_family_from_json(schema):
    with revit.Transaction("Generate Parametric Family from JSON"):

        # --- FIND REQUIRED VIEWS ---
        plan_view, elev_view = None, None
        for v in FilteredElementCollector(doc).OfClass(View):
            if v.IsTemplate:
                continue
            if v.ViewType == ViewType.FloorPlan and not plan_view:
                plan_view = v
            elif v.ViewType == ViewType.Elevation and not elev_view:
                if abs(v.ViewDirection.Y) > 0.99:
                    elev_view = v

        if not plan_view or not elev_view:
            forms.alert("Family template is missing a Plan or Front/Back Elevation view.", exitscript=True)

        if doc.FamilyManager.CurrentType is None:
            doc.FamilyManager.NewType("Standard")

        # --- STEP 1: PARAMETERS ---
        param_dict = {}
        for p in schema.get("parameters", []):
            p_type = p.get("type", "Length")
            try:
                if p_type == "Material":
                    fp = doc.FamilyManager.AddParameter(p["name"], GroupTypeId.Materials, SpecTypeId.Reference.Material, p["is_instance"])
                elif p_type == "YesNo":
                    fp = doc.FamilyManager.AddParameter(p["name"], GroupTypeId.Visibility, SpecTypeId.Boolean.YesNo, p["is_instance"])
                else:
                    fp = doc.FamilyManager.AddParameter(p["name"], GroupTypeId.Geometry, SpecTypeId.Length, p["is_instance"])
            except AttributeError:
                if p_type == "Material":
                    fp = doc.FamilyManager.AddParameter(p["name"], BuiltInParameterGroup.PG_MATERIALS, ParameterType.Material, p["is_instance"])
                elif p_type == "YesNo":
                    fp = doc.FamilyManager.AddParameter(p["name"], BuiltInParameterGroup.PG_VISIBILITY, ParameterType.YesNo, p["is_instance"])
                else:
                    fp = doc.FamilyManager.AddParameter(p["name"], BuiltInParameterGroup.PG_GEOMETRY, ParameterType.Length, p["is_instance"])

            if p.get("value") is not None:
                val = p["value"] * SCL if p_type == "Length" else p["value"]
                doc.FamilyManager.Set(fp, val)

            param_dict[p["name"]] = fp

        # --- STEP 2: REFERENCE PLANES ---
        rp_dict = {}
        for rp_data in schema.get("reference_planes", []):
            view = plan_view if rp_data["view"] == "Plan" else elev_view
            cut_vec = to_vec(rp_data["normal"]) if "normal" in rp_data else (XYZ.BasisZ if rp_data["view"] == "Plan" else XYZ.BasisY)
            rp = doc.FamilyCreate.NewReferencePlane(to_xyz(rp_data["p1"]), to_xyz(rp_data["p2"]), cut_vec, view)
            rp.Name = rp_data["name"]
            rp_dict[rp_data["name"]] = rp

        doc.Regenerate()

        # --- STEP 3: DIMENSIONS ---
        for dim_data in schema.get("dimensions", []):
            view = plan_view if dim_data["view"] == "Plan" else elev_view
            ref_array = ReferenceArray()
            for plane_name in dim_data["planes"]:
                ref_array.Append(rp_dict[plane_name].GetReference())
            dim_line = Line.CreateBound(to_xyz(dim_data["line_dir"][0]), to_xyz(dim_data["line_dir"][1]))
            try:
                new_dim = doc.FamilyCreate.NewDimension(view, dim_line, ref_array)
                new_dim.FamilyLabel = param_dict[dim_data["parameter"]]
            except Exception as e:
                logger.warning("Could not dimension '{}': {}".format(dim_data["parameter"], e))

        # --- STEP 4: GEOMETRY ---
        created_geometries = {}
        solid_forms, void_forms = [], []

        for geom_data in schema.get("geometry", []):
            created_geom = None
            is_solid = geom_data.get("is_solid", True)
            geom_type = geom_data["type"]

            # Build sketch plane (used by most types)
            if "sketch_plane_x" in geom_data:
                base_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, XYZ(geom_data["sketch_plane_x"] * SCL, 0, 0))
            elif "sketch_plane_y" in geom_data:
                base_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisY, XYZ(0, geom_data["sketch_plane_y"] * SCL, 0))
            else:
                z = geom_data.get("sketch_plane_z", 0.0) * SCL
                base_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, z))

            sketch_plane = SketchPlane.Create(doc, base_plane)
            sp_plane = sketch_plane.GetPlane()

            # ----------------------------------------------------------------
            # EXTRUSION
            # ----------------------------------------------------------------
            if geom_type == "Extrusion":
                profile = CurveArrArray()
                loop = CurveArray()
                for seg in geom_data.get("profile", []):
                    loop.Append(create_curve_from_json(seg, sp_plane=sp_plane))
                profile.Append(loop)
                # Support additional inner loops (holes)
                for inner in geom_data.get("inner_loops", []):
                    inner_loop = CurveArray()
                    for seg in inner:
                        inner_loop.Append(create_curve_from_json(seg, sp_plane=sp_plane))
                    profile.Append(inner_loop)

                created_geom = doc.FamilyCreate.NewExtrusion(
                    is_solid, profile, sketch_plane, geom_data["extrusion_end"] * SCL
                )
                created_geom.StartOffset = geom_data.get("extrusion_start", 0.0) * SCL

            # ----------------------------------------------------------------
            # SWEEP
            # ----------------------------------------------------------------
            elif geom_type == "Sweep":
                path = CurveArray()
                for seg in geom_data.get("path", []):
                    path.Append(create_curve_from_json(seg, sp_plane=sp_plane))

                sweep_profile = build_sweep_profile(geom_data.get("profile_2d", []), is_2d=True)
                created_geom = doc.FamilyCreate.NewSweep(
                    is_solid, path, sketch_plane, sweep_profile, 0, ProfilePlaneLocation.Start
                )

            # ----------------------------------------------------------------
            # REVOLVE
            # ----------------------------------------------------------------
            elif geom_type == "Revolve":
                profile = CurveArrArray()
                loop = CurveArray()
                for seg in geom_data.get("profile", []):
                    loop.Append(create_curve_from_json(seg, sp_plane=sp_plane))
                profile.Append(loop)

                axis_p1 = project_to_plane(to_xyz(geom_data["axis"]["p1"]), sp_plane)
                axis_p2 = project_to_plane(to_xyz(geom_data["axis"]["p2"]), sp_plane)
                axis_line = Line.CreateBound(axis_p1, axis_p2)
                created_geom = doc.FamilyCreate.NewRevolution(
                    is_solid, profile, sketch_plane, axis_line,
                    geom_data.get("start_angle", 0.0),
                    geom_data.get("end_angle", 6.283185307)
                )

            # ----------------------------------------------------------------
            # BLEND
            # ----------------------------------------------------------------
            elif geom_type == "Blend":
                bottom_profile = CurveArray()
                for seg in geom_data.get("bottom_profile", []):
                    bottom_profile.Append(create_curve_from_json(seg, sp_plane=sp_plane))

                second_end_offset = geom_data.get("second_end", 1.0) * SCL
                top_profile = CurveArray()
                for seg in geom_data.get("top_profile", []):
                    top_profile.Append(create_curve_from_json(seg, sp_plane=sp_plane, blend_offset=second_end_offset))

                created_geom = doc.FamilyCreate.NewBlend(is_solid, top_profile, bottom_profile, sketch_plane)
                if created_geom:
                    try:
                        created_geom.BottomOffset = geom_data.get("first_end", 0.0) * SCL
                        created_geom.TopOffset = second_end_offset
                    except Exception:
                        pass

            # ----------------------------------------------------------------
            # SWEPT BLEND
            # Profile transitions from profile_start -> profile_end along path.
            # Path must be a single curve; use first segment if multiple given.
            # JSON keys:
            #   "path"          : list of curve segments (first used as the path curve)
            #   "profile_start" : 2D profile curves at start of path
            #   "profile_end"   : 2D profile curves at end of path
            # ----------------------------------------------------------------
            elif geom_type == "SweptBlend":
                path_segs = geom_data.get("path", [])
                if not path_segs:
                    logger.warning("SweptBlend '{}': no path segments provided.".format(geom_data.get("id", "?")))
                else:
                    path_curve = create_curve_from_json(path_segs[0], sp_plane=sp_plane)

                    profile_start = build_sweep_profile(geom_data.get("profile_start", []), is_2d=True)
                    profile_end   = build_sweep_profile(geom_data.get("profile_end", []),   is_2d=True)

                    try:
                        created_geom = doc.FamilyCreate.NewSweptBlend(
                            is_solid, path_curve, sketch_plane, profile_start, profile_end
                        )
                    except Exception as e:
                        logger.error("SweptBlend creation failed: {}".format(e))

            # ----------------------------------------------------------------
            # MODEL TEXT  (chữ nổi 3D)
            # Creates actual 3D extruded text geometry in the family.
            # JSON keys:
            #   "text"       : string content (required)
            #   "position"   : [x, y, z] in mm — origin of the text baseline
            #   "depth"      : extrusion depth in mm (default 5)
            #   "h_align"    : "Left" | "Center" | "Right"  (default "Center")
            #   "text_type"  : name of an existing ModelTextType (optional)
            #   sketch_plane_z / _x / _y : defines the face the text sits on
            # ----------------------------------------------------------------
            elif geom_type == "ModelText":
                text_str   = geom_data.get("text", "Text")
                position   = to_xyz(geom_data.get("position", [0, 0, 0]))
                depth      = geom_data.get("depth", 5.0) * SCL

                h_align_map = {
                    "Left":   HorizontalAlign.Left,
                    "Center": HorizontalAlign.Center,
                    "Right":  HorizontalAlign.Right,
                }
                h_align = h_align_map.get(geom_data.get("h_align", "Center"), HorizontalAlign.Center)

                text_type_id = get_model_text_type(geom_data.get("text_type"))
                if text_type_id is None:
                    logger.warning("ModelText '{}': no ModelTextType found; skipping.".format(text_str))
                else:
                    try:
                        created_geom = doc.FamilyCreate.NewModelText(
                            text_str, text_type_id, sketch_plane, position, h_align, depth
                        )
                    except Exception as e:
                        logger.error("ModelText creation failed for '{}': {}".format(text_str, e))

            # ----------------------------------------------------------------
            # BIND MATERIAL & VISIBILITY PARAMETERS
            # ----------------------------------------------------------------
            if created_geom:
                if "id" in geom_data:
                    created_geometries[geom_data["id"]] = created_geom

                if is_solid:
                    solid_forms.append(created_geom)
                else:
                    void_forms.append(created_geom)

                mat_param_name = geom_data.get("material_param")
                if mat_param_name and mat_param_name in param_dict:
                    geom_mat_param = created_geom.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                    if geom_mat_param:
                        doc.FamilyManager.AssociateElementParameterToFamilyParameter(geom_mat_param, param_dict[mat_param_name])

                vis_param_name = geom_data.get("visible_param")
                if vis_param_name and vis_param_name in param_dict:
                    geom_vis_param = created_geom.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM)
                    if geom_vis_param:
                        doc.FamilyManager.AssociateElementParameterToFamilyParameter(geom_vis_param, param_dict[vis_param_name])

            doc.Regenerate()

            # ----------------------------------------------------------------
            # LOCK FACES TO REFERENCE PLANES
            # ----------------------------------------------------------------
            if "locks" in geom_data and created_geom:
                geom_opt = Options()
                geom_opt.ComputeReferences = True
                geometry_element = created_geom.get_Geometry(geom_opt)
                for geom_obj in geometry_element:
                    if isinstance(geom_obj, Solid):
                        for face in geom_obj.Faces:
                            if isinstance(face, PlanarFace):
                                normal = face.FaceNormal
                                for lock in geom_data["locks"]:
                                    req_norm = to_vec(lock["face_normal"])
                                    if normal.DotProduct(req_norm) > 0.99:
                                        target_rp = rp_dict[lock["plane"]]
                                        align_view = plan_view if abs(normal.Z) < 0.01 else elev_view
                                        try:
                                            alignment = doc.FamilyCreate.NewAlignment(align_view, target_rp.GetReference(), face.Reference)
                                            if alignment:
                                                alignment.IsLocked = True
                                        except Exception:
                                            pass

        # --- STEP 6: SEQUENTIAL VOID CUTTING ---
        doc.Regenerate()

        cuts_map = {}
        for geom_data in schema.get("geometry", []):
            if not geom_data.get("is_solid", True) and "cuts" in geom_data and "id" in geom_data:
                void_elem = created_geometries.get(geom_data["id"])
                if void_elem:
                    for solid_id in geom_data["cuts"]:
                        cuts_map.setdefault(solid_id, []).append(void_elem)

        for solid_id, void_list in cuts_map.items():
            current_target = created_geometries.get(solid_id)
            if not current_target:
                continue
            for void_elem in void_list:
                arr = CombinableElementArray()
                arr.Append(current_target)
                arr.Append(void_elem)
                try:
                    result = doc.CombineElements(arr)
                    if result:
                        current_target = result
                except Exception as e:
                    logger.error("Failed to cut '{}': {}".format(solid_id, e))


# =========================================================================
# EXECUTION
# =========================================================================
if __name__ == "__main__":
    logger.info("JSON to Family script started")

    if not doc.IsFamilyDocument:
        logger.warning("Script launched outside a Family Document – aborting")
        forms.alert(
            "This script must be run inside a Family Document (.rfa).\n\nOpen a Revit family file first.",
            title="Family Document Required",
            exitscript=True
        )

    dialog = JsonInputDialog()
    dialog.show_dialog()

    if dialog.json_data and dialog.json_data.strip() and dialog.json_data != "Paste your JSON schema here...":
        try:
            logger.info("Parsing JSON schema…")
            parsed_schema = json.loads(dialog.json_data)
            logger.info("Generating family from schema…")
            generate_family_from_json(parsed_schema)
            logger.info("Family generation completed successfully")
            forms.alert("Parametric family generated successfully!", title="Success")
        except ValueError as e:
            logger.error("Invalid JSON: {}".format(e))
            forms.alert("Invalid JSON:\n\n{}".format(e), title="JSON Error", exitscript=True)
        except Exception as e:
            logger.error("Error generating family: {}".format(e))
            forms.alert("Error generating family:\n\n{}".format(e), title="Generation Error", exitscript=True)
    else:
        logger.info("Dialog cancelled or no JSON provided")
        script.exit()
