# -*- coding: utf-8 -*-
"""
Annotation Manager
------------------
Unified table showing all Dimensions and Text Notes.
  - Filter by category / kind / keyword
  - Double-click Name cell to rename inline
  - Jump to view, delete selected rows
  - Auto-rename all types based on their properties

Author: T3Lab (Tran Tien Thanh)
"""

__title__  = "Annotation\nManager"
__author__ = "T3Lab"

import os
import re
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Data')

from System.Windows import Visibility, WindowState
from System.Windows.Media.Imaging import BitmapImage
from System.Windows.Input import Key
from System import Uri, UriKind
from System.Data import DataTable
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

doc   = revit.doc
uidoc = revit.uidoc

# ============================================================
# SHARED COLOR TABLE
# ============================================================
_DIM_COLORS = {
    (255,128,128):"Light Coral",(255,255,128):"Light Yellow",(128,255,128):"Pale Green",
    (128,255,255):"Pale Cyan",(128,128,255):"Light Slate Blue",(255,128,255):"Orchid",
    (255,0,0):"Red",(255,255,0):"Yellow",(0,255,0):"Lime",(0,255,255):"Cyan",
    (0,0,255):"Blue",(255,0,255):"Magenta",(128,64,64):"Brown",(255,192,128):"Light Salmon",
    (128,255,192):"Aquamarine",(192,192,255):"Lavender",(192,128,255):"Medium Orchid",
    (128,0,0):"Maroon",(255,128,0):"Orange",(0,128,0):"Green",(0,128,128):"Teal",
    (0,0,128):"Navy",(128,0,128):"Purple",(128,64,0):"Saddle Brown",(192,128,64):"Peru",
    (0,128,64):"Dark Sea Green",(0,128,192):"Steel Blue",(64,128,255):"Dodger Blue",
    (128,0,192):"Dark Orchid",(0,0,0):"Black",(128,128,0):"Olive",(128,128,128):"Gray128",
    (0,192,192):"Medium Turquoise",(192,192,192):"Silver",(255,255,255):"White",
    (70,70,70):"Gray70",(128,0,64):"Dark Raspberry",(77,77,77):"Gray77",
}
_TXT_COLORS = {
    (255,0,0):"Red",(0,255,0):"Lime",(0,0,255):"Blue",(255,255,0):"Yellow",
    (0,255,255):"Cyan",(255,0,255):"Magenta",(0,0,0):"Black",(255,255,255):"White",
    (128,128,128):"Gray",(128,0,0):"Maroon",(0,128,0):"Green",(0,0,128):"Navy",
    (128,128,0):"Olive",(0,128,128):"Teal",(128,0,128):"Purple",(255,128,0):"Orange",
    (128,128,255):"LightBlue",(192,192,192):"Silver",
}

def _rgb(color_int):
    return (color_int & 255, (color_int >> 8) & 255, (color_int >> 16) & 255)

def _sanitize(v):
    if not v:
        return "N/A"
    return re.sub(r'[\\/:?"<>|=]', '', v).strip() or "N/A"

def _mm(param):
    return "{:.2f}mm".format(round(param.AsDouble() * 304.8, 2))


# ============================================================
# DIMENSION RENAME HELPERS
# ============================================================
def _dim_name(dt, origin):
    def gp(bip):
        try: return dt.get_Parameter(bip)
        except: return None

    discipline = "STR" if "STR" in origin.upper() else "ARC"
    p = gp(BuiltInParameter.TEXT_SIZE)
    size  = _mm(p) if p else "N/A"
    p = gp(BuiltInParameter.TEXT_FONT)
    font  = p.AsString() if p else "N/A"
    p = gp(BuiltInParameter.DIM_TEXT_BACKGROUND)
    bg    = p.AsValueString() if p else "N/A"
    p = gp(BuiltInParameter.LINE_COLOR)
    color = _DIM_COLORS.get(_rgb(p.AsInteger()), "RGB") if p else "N/A"
    p = gp(BuiltInParameter.DIM_PREFIX)
    pref  = _sanitize(p.AsString()) if p else "N/A"
    p = gp(BuiltInParameter.DIM_STYLE_CENTERLINE_SYMBOL)
    ctr   = "Center" if (p and p.AsElementId() != ElementId.InvalidElementId) else "N/A"
    p = gp(BuiltInParameter.SPOT_ELEV_IND_ELEVATION)
    elev  = _sanitize(p.AsString()) if p else "N/A"
    p = gp(BuiltInParameter.SPOT_ELEV_IND_TOP)
    top   = _sanitize(p.AsString()) if p else "N/A"
    p = gp(BuiltInParameter.SPOT_ELEV_IND_BOTTOM)
    bot   = _sanitize(p.AsString()) if p else "N/A"

    parts = ["LB", discipline, size, font, bg]
    if color != "Black": parts.append(color)
    if ctr  != "N/A":   parts.append(ctr)
    if pref != "N/A":   parts.append(pref)
    if elev != "N/A":
        parts.append(elev)
    else:
        if top != "N/A": parts.append(top)
        if bot != "N/A": parts.append(bot)
    return "_".join(parts)


# ============================================================
# TEXTNOTE RENAME HELPERS
# ============================================================
def _txt_name(tt, origin):
    def gp(bip):
        try: return tt.get_Parameter(bip)
        except: return None

    discipline = "STR" if "STR" in origin.upper() else "ARC"
    p = gp(BuiltInParameter.TEXT_SIZE)
    size   = _mm(p) if p else "N/A"
    p = gp(BuiltInParameter.TEXT_FONT)
    font   = p.AsString().replace(" ", "") if p else "N/A"
    p = gp(BuiltInParameter.TEXT_BACKGROUND)
    bg     = ("Opaque" if p.AsInteger() == 0 else "Transparent") if p else "N/A"
    p = gp(BuiltInParameter.TEXT_WIDTH_SCALE)
    factor = str(round(p.AsDouble(), 2)) if p else "N/A"
    p = gp(BuiltInParameter.LINE_COLOR)
    color  = _TXT_COLORS.get(_rgb(p.AsInteger()), "RGB") if p else "N/A"
    p = gp(BuiltInParameter.TEXT_BOX_VISIBILITY)
    border = p and p.AsInteger() == 1
    p = gp(BuiltInParameter.TEXT_STYLE_BOLD)
    bold   = p and p.AsInteger() == 1
    p = gp(BuiltInParameter.TEXT_STYLE_UNDERLINE)
    uline  = p and p.AsInteger() == 1
    p = gp(BuiltInParameter.TEXT_STYLE_ITALIC)
    italic = p and p.AsInteger() == 1

    parts = ["LB", discipline, size, font, bg, factor]
    if color != "Black": parts.append(color)
    if border:  parts.append("Border")
    if bold:    parts.append("B")
    if uline:   parts.append("U")
    if italic:  parts.append("I")
    return "_".join(parts)


# ============================================================
# XAML PATH
# ============================================================
_GUI_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))),
    'lib', 'GUI'
)
_XAML_PATH = os.path.join(_GUI_DIR, 'AnnotationManager.xaml')
_LOGO_PATH = os.path.join(_GUI_DIR, 'T3Lab_logo.png')


# ============================================================
# WINDOW CLASS
# ============================================================
class AnnotationManagerWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML_PATH)

        # ── DataTable: columns _id, _cat, Category, Kind, Name, Details ──
        self._dt = DataTable()
        for col in ["_id", "_cat", "Category", "Kind", "Name", "Details"]:
            self._dt.Columns.Add(col)
        self.dg_ann.ItemsSource = self._dt.DefaultView

        # Map: element-id string → Revit element
        self._elem_map = {}

        # Wire cell-edit event for inline rename
        self.dg_ann.CellEditEnding += self.on_cell_edit_ending

        # Load logo
        try:
            bitmap = BitmapImage()
            bitmap.BeginInit()
            bitmap.UriSource = Uri(_LOGO_PATH, UriKind.Absolute)
            bitmap.EndInit()
            self.Icon = bitmap
            self.logo_image.Source = bitmap
        except Exception:
            pass

        # Load all data immediately
        self.load_all()

    # ── helpers ─────────────────────────────────────────────────────────

    def _status(self, msg):
        self.status.Text = msg

    def _add_row(self, elem_id, cat_code, category, kind, name, details):
        row = self._dt.NewRow()
        row["_id"]      = elem_id
        row["_cat"]     = cat_code
        row["Category"] = category
        row["Kind"]     = kind
        row["Name"]     = name
        row["Details"]  = details
        self._dt.Rows.Add(row)

    # ── Window controls ──────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, args):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, args):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, args):
        self.Close()

    # ── Data loading ─────────────────────────────────────────────────────

    def load_all(self, sender=None, args=None):
        self._dt.Clear()
        self._elem_map = {}

        # ── Dimension types ──────────────────────────────────────────────
        dim_types = list(
            FilteredElementCollector(doc).OfClass(DimensionType)
            .WhereElementIsElementType().ToElements()
        )
        dim_insts = list(
            FilteredElementCollector(doc).OfClass(Dimension)
            .WhereElementIsNotElementType().ToElements()
        )

        # Count instances per dimension type
        dim_type_cnt = {}
        for d in dim_insts:
            tid = str(d.GetTypeId())
            dim_type_cnt[tid] = dim_type_cnt.get(tid, 0) + 1

        for dt in dim_types:
            name = dt.Name or "<unnamed>"
            cnt  = dim_type_cnt.get(str(dt.Id), 0)
            self._add_row(str(dt.Id), "DimType", "Dimension", "Type",
                          name, "{} instance(s)".format(cnt))
            self._elem_map[str(dt.Id)] = dt

        # ── Dimension instances ──────────────────────────────────────────
        for d in dim_insts:
            view      = doc.GetElement(d.OwnerViewId)
            view_name = view.Name if view else "?"
            self._add_row(str(d.Id), "DimInst", "Dimension", "Instance",
                          d.Name or "<no type>", view_name)
            self._elem_map[str(d.Id)] = d

        # ── Text Note types ──────────────────────────────────────────────
        txt_types = list(
            FilteredElementCollector(doc).OfClass(TextNoteType)
            .WhereElementIsElementType().ToElements()
        )
        txt_insts = list(
            FilteredElementCollector(doc).OfClass(TextNote)
            .WhereElementIsNotElementType().ToElements()
        )

        txt_type_cnt = {}
        for tn in txt_insts:
            tid = str(tn.GetTypeId())
            txt_type_cnt[tid] = txt_type_cnt.get(tid, 0) + 1

        for tt in txt_types:
            p    = tt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            name = p.AsString() if p else "<unnamed>"
            cnt  = txt_type_cnt.get(str(tt.Id), 0)
            self._add_row(str(tt.Id), "TxtType", "Text Note", "Type",
                          name, "{} instance(s)".format(cnt))
            self._elem_map[str(tt.Id)] = tt

        # ── Text Note instances ──────────────────────────────────────────
        for tn in txt_insts:
            view      = doc.GetElement(tn.ViewId)
            view_name = view.Name if view else "?"
            preview   = (tn.Text or "")[:60].replace("\n", " ").replace("\r", "")
            self._add_row(str(tn.Id), "TxtInst", "Text Note", "Instance",
                          preview, view_name)
            self._elem_map[str(tn.Id)] = tn

        # Reset filter and update count
        self.dg_ann.ItemsSource.RowFilter = ""
        total = len(self._elem_map)
        self.count_lbl.Text = "{} items".format(total)
        self._status("Loaded {} annotation items.".format(total))

    # ── Search / filter ──────────────────────────────────────────────────

    def do_search(self, sender, args):
        cat_idx  = self.cb_cat.SelectedIndex   # 0=All, 1=Dimension, 2=Text Note
        kind_idx = self.cb_kind.SelectedIndex  # 0=All, 1=Instance, 2=Type
        kw       = self.kw_search.Text.strip()

        filters = []
        if cat_idx == 1:
            filters.append("Category = 'Dimension'")
        elif cat_idx == 2:
            filters.append("Category = 'Text Note'")

        if kind_idx == 1:
            filters.append("Kind = 'Instance'")
        elif kind_idx == 2:
            filters.append("Kind = 'Type'")

        if kw:
            safe_kw = kw.replace("'", "''")
            filters.append("Name LIKE '%{}%'".format(safe_kw))

        row_filter = " AND ".join(filters) if filters else ""
        try:
            self.dg_ann.ItemsSource.RowFilter = row_filter
        except Exception as e:
            self._status("Filter error: {}".format(e))
            return

        visible = len(self.dg_ann.Items)
        self.count_lbl.Text = "{} items".format(visible)
        self._status("Showing {} of {} items.".format(visible, len(self._elem_map)))

    def search_on_enter(self, sender, args):
        if args.Key == Key.Enter:
            self.do_search(None, None)

    # ── Inline rename (CellEditEnding) ───────────────────────────────────

    def on_cell_edit_ending(self, sender, args):
        # Only handle the Name column on Commit
        if str(args.Column.Header) != "Name":
            return
        if str(args.EditAction) != "Commit":
            return

        tb       = args.EditingElement        # TextBox
        new_name = tb.Text.strip()
        if not new_name:
            args.Cancel = True
            return

        row      = args.Row.Item              # DataRowView
        elem_id  = str(row["_id"])
        cat_code = str(row["_cat"])
        old_name = str(row["Name"])

        if new_name == old_name:
            return

        elem = self._elem_map.get(elem_id)
        if not elem:
            return

        # Dimension instances don't have a user-editable name
        if cat_code == "DimInst":
            args.Cancel = True
            self._status("Dimension instances cannot be renamed directly — rename the Type instead.")
            return

        t = Transaction(doc, "Rename Annotation")
        t.Start()
        try:
            if cat_code in ("DimType", "TxtType"):
                elem.Name = new_name
            elif cat_code == "TxtInst":
                elem.Text = new_name
            t.Commit()
            self._status(u"Renamed: '{}' \u2192 '{}'.".format(old_name[:40], new_name[:40]))
        except Exception as e:
            t.RollBack()
            args.Cancel = True
            self._status("Rename failed: {}".format(e))

    # ── Jump to View ─────────────────────────────────────────────────────

    def do_jump(self, sender, args):
        selected = list(self.dg_ann.SelectedItems)
        if not selected:
            self._status("Select a row first.")
            return

        row      = selected[0]
        elem_id  = str(row["_id"])
        cat_code = str(row["_cat"])
        elem     = self._elem_map.get(elem_id)
        if not elem:
            return

        if cat_code == "DimInst":
            view = doc.GetElement(elem.OwnerViewId)
        elif cat_code == "TxtInst":
            view = doc.GetElement(elem.ViewId)
        else:
            self._status("Jump to View is only available for instances (not types).")
            return

        if view:
            uidoc.ActiveView = view
            uidoc.ShowElements(elem.Id)
            self._status("Jumped to view '{}' — '{}'.".format(
                view.Name, str(row["Name"])[:50]))

    # ── Delete selected ──────────────────────────────────────────────────

    def do_delete(self, sender, args):
        selected = list(self.dg_ann.SelectedItems)
        if not selected:
            self._status("Nothing selected.")
            return

        from pyrevit import forms as pf
        if not pf.alert(
            "Delete {} selected item(s)?\n\nDeleting a Type also removes all its instances.".format(len(selected)),
            title="Confirm Delete", yes=True, no=True
        ):
            return

        delete_ids = [str(r["_id"]) for r in selected]

        t = Transaction(doc, "Delete Selected Annotations")
        t.Start()
        deleted = 0
        errors  = 0
        ok_ids  = []
        for elem_id in delete_ids:
            elem = self._elem_map.get(elem_id)
            if elem:
                try:
                    doc.Delete(elem.Id)
                    ok_ids.append(elem_id)
                    deleted += 1
                except Exception:
                    errors += 1
        t.Commit()

        # Remove deleted rows from DataTable
        ok_set        = set(ok_ids)
        rows_to_del   = list(r for r in self._dt.Rows if str(r["_id"]) in ok_set)
        for r in rows_to_del:
            self._dt.Rows.Remove(r)
        for eid in ok_ids:
            self._elem_map.pop(eid, None)

        msg = "Deleted {} item(s).".format(deleted)
        if errors:
            msg += "  ({} could not be deleted — may be in use.)".format(errors)
        self.count_lbl.Text = "{} items".format(len(self.dg_ann.Items))
        self._status(msg)

    # ── Selection helpers ────────────────────────────────────────────────

    def do_select_all(self, sender, args):
        self.dg_ann.SelectAll()

    def do_clear(self, sender, args):
        self.dg_ann.UnselectAll()

    # ── Auto-rename all types ────────────────────────────────────────────

    def dim_rename_all(self, sender, args):
        from pyrevit import forms as pf
        if not pf.alert(
            "Auto-rename ALL DimensionTypes in this document?\nThis cannot be undone.",
            title="Confirm Rename", yes=True, no=True
        ):
            return
        t = Transaction(doc, "Rename Dimension Types")
        t.Start()
        count = 0
        try:
            for dt in FilteredElementCollector(doc).OfClass(DimensionType)\
                      .WhereElementIsElementType().ToElements():
                try:
                    origin = dt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    dt.Name = _dim_name(dt, origin)
                    count += 1
                except Exception:
                    pass
        finally:
            t.Commit()
        self._status("Renamed {} DimensionType(s).  Reloading table…".format(count))
        self.load_all()

    def txt_rename_all(self, sender, args):
        from pyrevit import forms as pf
        if not pf.alert(
            "Auto-rename ALL TextNoteTypes in this document?\nThis cannot be undone.",
            title="Confirm Rename", yes=True, no=True
        ):
            return
        t = Transaction(doc, "Rename TextNote Types")
        t.Start()
        count = 0
        try:
            for tt in FilteredElementCollector(doc).OfClass(TextNoteType)\
                      .WhereElementIsElementType().ToElements():
                try:
                    origin = tt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    tt.Name = _txt_name(tt, origin)
                    count += 1
                except Exception:
                    pass
        finally:
            t.Commit()
        self._status("Renamed {} TextNoteType(s).  Reloading table…".format(count))
        self.load_all()


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    win = AnnotationManagerWindow()
    win.show_dialog()
