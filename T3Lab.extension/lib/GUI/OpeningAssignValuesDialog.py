# -*- coding: utf-8 -*-
"""Opening Assign Values Dialog — assigns area values to filled region elements."""

import os
import math
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('RevitAPI')

from pyrevit import forms

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction, Options
)


class OpeningAssignValuesDialog(forms.WPFWindow):

    def __init__(self, revit_obj):
        self._app = revit_obj
        self._doc = revit_obj.ActiveUIDocument.Document
        self._uidoc = revit_obj.ActiveUIDocument

        xaml_path = os.path.join(os.path.dirname(__file__), 'Tools', 'OpeningAssignValues.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        self.btn_minimize.Click += self._on_minimize
        self.btn_maximize.Click += self._on_maximize
        self.btn_close_chrome.Click += self._on_close
        self.btn_execute.Click += self._on_execute

        self._load_counts()

    # ── Window chrome ──────────────────────────────────────────────────────────

    def _on_minimize(self, s, e):
        import System.Windows
        self.WindowState = System.Windows.WindowState.Minimized

    def _on_maximize(self, s, e):
        import System.Windows
        if self.WindowState == System.Windows.WindowState.Maximized:
            self.WindowState = System.Windows.WindowState.Normal
        else:
            self.WindowState = System.Windows.WindowState.Maximized

    def _on_close(self, s, e):
        self.Close()

    # ── Data loading ───────────────────────────────────────────────────────────

    def _collect_regions(self):
        active_view = self._doc.ActiveView
        if active_view is None:
            return [], []

        view_id = active_view.Id
        detail_components = FilteredElementCollector(self._doc, view_id)\
            .OfCategory(BuiltInCategory.OST_DetailComponents)\
            .ToElements()

        openings = []
        walls = []
        for dc in detail_components:
            if dc.Name != "Detail Filled Region":
                continue
            fam_type = dc.LookupParameter("Family and Type")
            if fam_type is None:
                continue
            val = fam_type.AsValueString()
            if val == "Filled region: _Area of Opening":
                openings.append(dc)
            elif val == "Filled region: _Area of Wall":
                walls.append(dc)

        return openings, walls

    def _load_counts(self):
        try:
            openings, walls = self._collect_regions()
            self.txt_opening_count.Text = str(len(openings))
            self.txt_wall_count.Text = str(len(walls))

            if len(openings) == 0 and len(walls) == 0:
                self.pnl_warning.Visibility = self._visible
                self.txt_warning.Text = "No filled regions found in the active view. Open a plan view that contains filled region elements."
                self.btn_execute.IsEnabled = False
            else:
                self.pnl_warning.Visibility = self._collapsed
                self.btn_execute.IsEnabled = True
                self.txt_status.Text = "Found {} opening region(s) and {} wall region(s) in active view.".format(
                    len(openings), len(walls))
        except Exception as ex:
            self.txt_status.Text = "Error loading: {}".format(str(ex))

    @property
    def _visible(self):
        from System.Windows import Visibility
        return Visibility.Visible

    @property
    def _collapsed(self):
        from System.Windows import Visibility
        return Visibility.Collapsed

    # ── Logic ─────────────────────────────────────────────────────────────────

    def _get_bbox(self, element):
        return element.get_BoundingBox(None)

    def _check_overlap(self, curve_wall, curve_opening):
        wallMinX = round(curve_wall.Min.X, 10)
        wallMinY = round(curve_wall.Min.Y, 10)
        wallMinZ = round(curve_wall.Min.Z, 10)
        wallMaxX = round(curve_wall.Max.X, 10)
        wallMaxY = round(curve_wall.Max.Y, 10)
        wallMaxZ = round(curve_wall.Max.Z, 10)

        openingMinX = round(curve_opening.Min.X, 10)
        openingMinY = round(curve_opening.Min.Y, 10)
        openingMinZ = round(curve_opening.Min.Z, 10)
        openingMaxX = round(curve_opening.Max.X, 10)
        openingMaxY = round(curve_opening.Max.Y, 10)
        openingMaxZ = round(curve_opening.Max.Z, 10)

        if openingMinX != openingMaxX and openingMinY != openingMaxY and openingMinZ != openingMaxZ:
            if not (openingMinZ >= wallMinZ and openingMaxZ <= wallMaxZ):
                return False
            minmax_wall_XY = (wallMaxX - wallMinX, wallMaxY - wallMinY)
            if minmax_wall_XY[0] == 0 or minmax_wall_XY[1] == 0:
                return False
            minmin_open_XY = (openingMinX - wallMinX, openingMinY - wallMinY)
            minmax_open_XY = (openingMaxX - wallMinX, openingMaxY - wallMinY)
            d1_1 = round(minmin_open_XY[0] / minmax_wall_XY[0], 10)
            d1_2 = round(minmin_open_XY[1] / minmax_wall_XY[1], 10)
            d2_1 = round(minmax_open_XY[0] / minmax_wall_XY[0], 10)
            d2_2 = round(minmax_open_XY[1] / minmax_wall_XY[1], 10)
            if d1_1 == d1_2 and d2_1 == d2_2 and 0 <= d1_1 <= 1 and 0 <= d2_1 <= 1:
                return True
            return False
        else:
            c1 = c2 = c3 = c4 = False
            if openingMinX == openingMaxX:
                c1 = openingMinY >= wallMinY
                c2 = openingMinZ >= wallMinZ
                c3 = openingMaxY <= wallMaxY
                c4 = openingMaxZ <= wallMaxZ
            elif openingMinY == openingMaxY:
                c1 = openingMinX >= wallMinX
                c2 = openingMinZ >= wallMinZ
                c3 = openingMaxX <= wallMaxX
                c4 = openingMaxZ <= wallMaxZ
            elif openingMinZ == openingMaxZ:
                c1 = openingMinX >= wallMinX
                c2 = openingMinY >= wallMinY
                c3 = openingMaxX <= wallMaxX
                c4 = openingMaxY <= wallMaxY
            return c1 and c2 and c3 and c4

    def _assign_level(self, fg_wall, level_dict):
        curve_wall = self._get_bbox(fg_wall)
        if curve_wall is None:
            return
        wallMinZ = curve_wall.Min.Z
        for level_name, level_elevation in level_dict.items():
            if level_elevation is not None and level_elevation < 0:
                p = fg_wall.LookupParameter("GMG_LEVEL OF ALLOWABLE OPENING")
                if p and not p.IsReadOnly:
                    p.Set("1ST FLOOR")
                break
            if level_elevation is not None and abs(wallMinZ - level_elevation) < 0.001:
                name_split = level_name.split(" ", 1)
                suffix = name_split[-1] if name_split else level_name
                p = fg_wall.LookupParameter("GMG_LEVEL OF ALLOWABLE OPENING")
                if p and not p.IsReadOnly:
                    p.Set(suffix)
                break

    # ── Execute handler ────────────────────────────────────────────────────────

    def _on_execute(self, s, e):
        self.btn_execute.IsEnabled = False
        self.txt_status.Text = "Running assignment..."
        self.pnl_result.Visibility = self._collapsed

        try:
            openings, walls = self._collect_regions()
            if not openings and not walls:
                self.txt_status.Text = "No elements found — nothing to assign."
                self.btn_execute.IsEnabled = True
                return

            active_view = self._doc.ActiveView
            view_id = active_view.Id

            # Collect levels
            levels = FilteredElementCollector(self._doc, view_id)\
                .OfCategory(BuiltInCategory.OST_Levels)\
                .ToElements()
            level_dict = {}
            for lv in levels:
                elev_p = lv.LookupParameter("Elevation")
                level_dict[lv.Name] = elev_p.AsDouble() if elev_p else None

            success_count = 0
            error_count = 0

            t = Transaction(self._doc, "Opening Assign Values")
            t.Start()
            try:
                for fg in openings:
                    try:
                        area_p = fg.LookupParameter("Area")
                        roundup_p = fg.LookupParameter("GMG_AREA ROUNDUP")
                        if area_p and roundup_p and not roundup_p.IsReadOnly:
                            roundup_p.Set(math.ceil(round(area_p.AsDouble(), 5)))
                            success_count += 1
                    except Exception:
                        error_count += 1

                for fg_wall in walls:
                    try:
                        area_p = fg_wall.LookupParameter("Area")
                        wall_area_p = fg_wall.LookupParameter("GMG_AREA OF WALL")
                        if area_p and wall_area_p and not wall_area_p.IsReadOnly:
                            wall_area_p.Set(math.floor(area_p.AsDouble()))

                        # Allowable opening percent from separation distance
                        separa = fg_wall.LookupParameter("GMG_SEPARATION DISTANCE")
                        allow_p = fg_wall.LookupParameter("GMG_ALLOWABLE OPENING")
                        if separa and allow_p and not allow_p.IsReadOnly:
                            sep_val = separa.AsString() or ""
                            if sep_val.endswith("10'"):
                                allow_p.Set("25%")
                            elif sep_val.endswith("15'"):
                                allow_p.Set("45%")
                            elif sep_val.endswith("20'"):
                                allow_p.Set("75%")
                            elif sep_val.endswith("25'") or sep_val.endswith("30'"):
                                allow_p.Set("NO LIMIT")

                        # Sum opening areas that overlap this wall
                        curve_wall = self._get_bbox(fg_wall)
                        area_opening = 0.0
                        for fg_op in openings:
                            curve_op = self._get_bbox(fg_op)
                            if curve_wall and curve_op and self._check_overlap(curve_wall, curve_op):
                                roundup_p = fg_op.LookupParameter("GMG_AREA ROUNDUP")
                                if roundup_p:
                                    area_opening += roundup_p.AsDouble()

                        opening_area_p = fg_wall.LookupParameter("GMG_AREA OF OPENING")
                        if opening_area_p and not opening_area_p.IsReadOnly:
                            opening_area_p.Set(area_opening)

                        # Percent
                        wall_area_val = wall_area_p.AsDouble() if wall_area_p else 0
                        percent_p = fg_wall.LookupParameter("GMG_PERCENT OF OPENING")
                        if percent_p and not percent_p.IsReadOnly and wall_area_val != 0:
                            percent_p.Set(area_opening / wall_area_val)

                        self._assign_level(fg_wall, level_dict)
                        success_count += 1
                    except Exception:
                        error_count += 1

                t.Commit()
            except Exception as tx_err:
                t.RollbackToSavepoint() if t.HasStarted() else None
                try:
                    t.RollBack()
                except Exception:
                    pass
                raise tx_err

            msg = "Assigned values to {} element(s).".format(success_count)
            if error_count > 0:
                msg += " {} error(s) skipped.".format(error_count)
            self.txt_result.Text = msg
            self.pnl_result.Visibility = self._visible
            self.txt_status.Text = "Done: " + msg

        except Exception as ex:
            self.txt_status.Text = "Error: {}".format(str(ex))
        finally:
            self.btn_execute.IsEnabled = True


def show_opening_assign_values(revit_obj):
    dlg = OpeningAssignValuesDialog(revit_obj)
    dlg.ShowDialog()
