# -*- coding: utf-8 -*-
"""Datum Manager — inline grid save/restore + alignment tools."""

import os
import pickle
import __builtin__
from tempfile import gettempdir
from collections import namedtuple

from pyrevit import revit, DB, UI, forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'DatumManager.xaml')
_GRID_PICKLE_PATH = os.path.join(gettempdir(), 'GridPlacement')

_Point = namedtuple('_Point', ['X', 'Y', 'Z'])

_SUPPORTED_VIEWS = [
    DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.Detail,
    DB.ViewType.AreaPlan, DB.ViewType.Section, DB.ViewType.Elevation,
]


def _save_grids(doc, uidoc):
    cView = doc.ActiveView
    if cView.ViewType not in _SUPPORTED_VIEWS:
        forms.alert('View type \'{}\' not supported.'.format(cView.ViewType))
        return

    selection = [doc.GetElement(eid) for eid in uidoc.Selection.GetElementIds()]
    GridLines = {}
    n = 0

    for el in selection:
        if not isinstance(el, DB.Grid):
            continue
        curves = el.GetCurvesInView(DB.DatumExtentType.ViewSpecific, cView)
        if len(curves) != 1:
            forms.alert('Grid \'{}\' has {} curves — skipped.'.format(el.Name, len(curves)))
            continue

        cCurve = curves[0]
        p0 = cCurve.GetEndPoint(0)
        p1 = cCurve.GetEndPoint(1)

        cgl = {
            'Name': el.Name,
            'Start': _Point(p0.X, p0.Y, p0.Z),
            'End':   _Point(p1.X, p1.Y, p1.Z),
            'StartBubble':        el.HasBubbleInView(DB.DatumEnds.End0, cView),
            'EndBubble':          el.HasBubbleInView(DB.DatumEnds.End1, cView),
            'StartBubbleVisible': el.IsBubbleVisibleInView(DB.DatumEnds.End0, cView),
            'EndBubbleVisible':   el.IsBubbleVisibleInView(DB.DatumEnds.End1, cView),
        }

        l0 = el.GetLeader(DB.DatumEnds.End0, cView)
        if l0:
            cgl['Leader0Elbow']  = _Point(l0.Elbow.X,  l0.Elbow.Y,  l0.Elbow.Z)
            cgl['Leader0End']    = _Point(l0.End.X,    l0.End.Y,    l0.End.Z)
            cgl['Leader0Anchor'] = _Point(l0.Anchor.X, l0.Anchor.Y, l0.Anchor.Z)

        l1 = el.GetLeader(DB.DatumEnds.End1, cView)
        if l1:
            cgl['Leader1Elbow']  = _Point(l1.Elbow.X,  l1.Elbow.Y,  l1.Elbow.Z)
            cgl['Leader1End']    = _Point(l1.End.X,    l1.End.Y,    l1.End.Z)
            cgl['Leader1Anchor'] = _Point(l1.Anchor.X, l1.Anchor.Y, l1.Anchor.Z)

        if isinstance(cCurve, DB.Arc):
            c = cCurve.Center
            cgl['Center'] = _Point(c.X, c.Y, c.Z)

        GridLines[el.Name] = cgl
        n += 1

    if n > 0:
        with open(_GRID_PICKLE_PATH, 'wb') as fp:
            pickle.dump(GridLines, fp)
        forms.alert('Saved {} grid position{}.'.format(n, 's' if n != 1 else ''))
    else:
        forms.alert('No grids selected. Select grids in the view first.')


def _restore_grids(doc, uidoc, all_in_view=False):
    cView = doc.ActiveView
    if cView.ViewType not in _SUPPORTED_VIEWS:
        forms.alert('View type \'{}\' not supported.'.format(cView.ViewType))
        return

    Axes = [doc.GetElement(eid) for eid in uidoc.Selection.GetElementIds()]
    if all_in_view or not Axes:
        Axes = list(DB.FilteredElementCollector(doc, cView.Id).OfClass(DB.Grid).ToElements())

    try:
        with open(_GRID_PICKLE_PATH, 'rb') as fp:
            GridLines = pickle.load(fp)
    except IOError:
        forms.alert('No saved grid positions found. Save positions first.')
        return

    n = 0
    for cAxis in Axes:
        if not isinstance(cAxis, DB.Grid):
            continue
        if cAxis.Name not in GridLines:
            continue

        curves = cAxis.GetCurvesInView(DB.DatumExtentType.ViewSpecific, cView)
        if len(curves) != 1:
            continue

        cCurve = curves[0]
        cData  = GridLines[cAxis.Name]

        tmp0 = cCurve.GetEndPoint(0)
        tmp1 = cCurve.GetEndPoint(1)

        if cView.ViewType in [DB.ViewType.Section, DB.ViewType.Elevation]:
            pt0 = DB.XYZ(tmp0.X, tmp0.Y, cData['Start'].Z)
            pt1 = DB.XYZ(tmp0.X, tmp0.Y, cData['End'].Z)
        else:
            pt0 = DB.XYZ(cData['Start'].X, cData['Start'].Y, tmp0.Z)
            pt1 = DB.XYZ(cData['End'].X,   cData['End'].Y,   tmp1.Z)

        if isinstance(cCurve, DB.Arc):
            ptRef = cCurve.Evaluate(0.5, True)
            gridline = DB.Arc.Create(pt0, pt1, ptRef)
        else:
            gridline = DB.Line.CreateBound(pt0, pt1)

        if cAxis.IsCurveValidInView(DB.DatumExtentType.ViewSpecific, cView, gridline):
            with revit.Transaction('Restore grid curve \'{}\''.format(cAxis.Name)):
                cAxis.SetCurveInView(DB.DatumExtentType.ViewSpecific, cView, gridline)

        with revit.Transaction('Restore grid placement \'{}\''.format(cAxis.Name)):
            if cData.get('StartBubble') and cData.get('StartBubbleVisible'):
                cAxis.ShowBubbleInView(DB.DatumEnds.End0, cView)
                if 'Leader0Anchor' in cData and not cAxis.GetLeader(DB.DatumEnds.End0, cView):
                    cAxis.AddLeader(DB.DatumEnds.End0, cView)
            else:
                cAxis.HideBubbleInView(DB.DatumEnds.End0, cView)

            if cData.get('EndBubble') and cData.get('EndBubbleVisible'):
                cAxis.ShowBubbleInView(DB.DatumEnds.End1, cView)
                if 'Leader1Anchor' in cData and not cAxis.GetLeader(DB.DatumEnds.End1, cView):
                    cAxis.AddLeader(DB.DatumEnds.End1, cView)
            else:
                cAxis.HideBubbleInView(DB.DatumEnds.End1, cView)

        n += 1

    forms.alert('Restored {} grid position{}.'.format(n, 's' if n != 1 else ''))


class DatumManagerWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit_app):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit_app

        self.btn_save_grids.Click        += self._on_save_grids
        self.btn_restore_grids.Click     += self._on_restore_grids
        self.btn_restore_all_grids.Click += self._on_restore_all_grids
        self.btn_align_gridlines.Click   += self._on_align_gridlines
        self.btn_convert_grid.Click      += self._on_convert_grid
        self.btn_align_levels.Click      += self._on_align_levels
        self.btn_convert_level.Click     += self._on_convert_level

        self.btn_minimize.Click     += self._minimize
        self.btn_maximize.Click     += self._maximize
        self.btn_close_chrome.Click += self._close_chrome
        self.PreviewKeyDown         += self._on_key_down

    def _launch(self, rel_path):
        script_path = os.path.normpath(os.path.join(self._script_dir, rel_path))
        self.Close()
        g = {'__name__': '__main__', '__file__': script_path,
             '__builtins__': __builtin__, '__revit__': self._revit}
        try:
            execfile(script_path, g)
        except Exception as ex:
            forms.alert("Error launching tool:\n{}".format(ex))

    def _on_save_grids(self, sender, e):
        try:
            _save_grids(revit.doc, revit.uidoc)
        except Exception as ex:
            forms.alert("Save Grids error:\n{}".format(ex))

    def _on_restore_grids(self, sender, e):
        try:
            _restore_grids(revit.doc, revit.uidoc, all_in_view=False)
        except Exception as ex:
            forms.alert("Restore Grids error:\n{}".format(ex))

    def _on_restore_all_grids(self, sender, e):
        try:
            _restore_grids(revit.doc, revit.uidoc, all_in_view=True)
        except Exception as ex:
            forms.alert("Restore All Grids error:\n{}".format(ex))

    def _on_align_gridlines(self, sender, e):
        self._launch("../Datum.pulldown/Gridline.pulldown/Align Gridline.pushbutton/script.py")

    def _on_convert_grid(self, sender, e):
        self._launch("../Datum.pulldown/Gridline.pulldown/ConvertGridline.pushbutton/script.py")

    def _on_align_levels(self, sender, e):
        self._launch("../Datum.pulldown/Level.pulldown/Align Level.pushbutton/script.py")

    def _on_convert_level(self, sender, e):
        self._launch("../Datum.pulldown/Level.pulldown/ConvertLevel.pushbutton/script.py")

    def _minimize(self, sender, e):
        import System.Windows
        self.WindowState = System.Windows.WindowState.Minimized

    def _maximize(self, sender, e):
        import System.Windows
        if self.WindowState == System.Windows.WindowState.Maximized:
            self.WindowState = System.Windows.WindowState.Normal
        else:
            self.WindowState = System.Windows.WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()

    def _on_key_down(self, sender, e):
        import System.Windows.Input as WI
        if e.Key == WI.Key.Escape:
            self.Close()


def show_datum_manager(script_dir, revit_app):
    DatumManagerWindow(script_dir, revit_app).ShowDialog()
