# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')

from System.Windows import WindowState, Visibility
from System.Windows.Forms import OpenFileDialog, DialogResult
from System.Windows.Media import Brushes

from pyrevit import forms, DB, revit


RESOLUTION_MAP = {0: 150, 1: 300, 2: 600}

_VIEW_TYPES = {
    DB.ViewType.FloorPlan:    "Floor Plan",
    DB.ViewType.CeilingPlan:  "Ceiling Plan",
    DB.ViewType.DraftingView: "Drafting",
}

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'PDFImport.xaml')


class ViewItem(object):
    def __init__(self, name, type_label, view_id):
        self.Name        = name
        self.Type        = type_label
        self.view_id     = view_id
        self.IsSelected  = True
        self.PageDisplay = u"–"

    def set_page(self, n):
        self.PageDisplay = str(n) if n else u"–"


class PDFImportDialog(forms.WPFWindow):

    def __init__(self):
        self._pdf_path  = None
        self._all_items = []
        self._items     = []
        self._loading   = False  # blocks re-entrant checkbox events during DataGrid updates
        forms.WPFWindow.__init__(self, _XAML)
        self.Loaded += self._on_loaded

    def _on_loaded(self, sender, args):
        self._load_views()

    # ── Data ──────────────────────────────────────────────────────────────────

    def _load_views(self):
        items = []
        try:
            collector = (DB.FilteredElementCollector(revit.doc)
                           .OfCategory(DB.BuiltInCategory.OST_Views)
                           .WhereElementIsNotElementType())
            for v in collector:
                try:
                    if v.IsTemplate:
                        continue
                    label = _VIEW_TYPES.get(v.ViewType)
                    if label is None:
                        continue
                    items.append(ViewItem(v.Name or "Unnamed", label, v.Id))
                except Exception:
                    pass
        except Exception:
            pass

        items.sort(key=lambda x: x.Name)
        self._all_items = items
        self.pnl_loading.Visibility = Visibility.Collapsed
        self._refresh_list()

    def _refresh_list(self):
        """Rebuild filtered+sorted list and rebind ItemsSource. Call on load / filter / sort changes."""
        q = self.txt_search.Text.strip().lower()
        result = [i for i in self._all_items if not q or q in i.Name.lower()]

        if self.cmb_sort.SelectedIndex == 1:
            result.sort(key=lambda i: (i.Type, i.Name))
        else:
            result.sort(key=lambda i: i.Name)

        self._items = result
        self._assign_pages()

        # Guard: WPF fires Checked on each checkbox during ItemsSource bind because
        # IsSelected defaults to True; without this, view_selection_changed re-enters
        # _refresh_list infinitely.
        self._loading = True
        self.grid_views.ItemsSource = self._items
        self._loading = False

        self._update_status()

    def _refresh_pages(self):
        """Recompute page numbers and refresh display. Call on selection changes only."""
        self._assign_pages()
        self._loading = True
        self.grid_views.Items.Refresh()
        self._loading = False
        self._update_status()

    def _assign_pages(self):
        page = 1
        for item in self._items:
            if item.IsSelected:
                item.set_page(page)
                page += 1
            else:
                item.set_page(None)

    def _update_status(self):
        total    = len(self._items)
        selected = sum(1 for i in self._items if i.IsSelected)
        self.txt_view_count.Text     = "{} views".format(total)
        self.txt_selected_count.Text = "{} selected".format(selected)
        self.btn_import.IsEnabled    = bool(self._pdf_path and selected > 0)
        if not self._pdf_path:
            self.txt_status.Text = "Select a PDF file and choose target views"
        elif selected == 0:
            self.txt_status.Text = "Select at least one target view"
        else:
            self.txt_status.Text = u"Ready — {} page(s) will be imported".format(selected)

    # ── Events ────────────────────────────────────────────────────────────────

    def browse_pdf_clicked(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title       = "Select PDF File"
        dlg.Filter      = "PDF files (*.pdf)|*.pdf"
        dlg.Multiselect = False
        if dlg.ShowDialog() != DialogResult.OK:
            return
        self._pdf_path = dlg.FileName
        fname    = os.path.basename(self._pdf_path)
        size_kb  = os.path.getsize(self._pdf_path) / 1024.0
        size_str = ("{:.0f} KB".format(size_kb) if size_kb < 1024
                    else "{:.1f} MB".format(size_kb / 1024.0))
        self.txt_pdf_path.Text        = fname
        self.txt_pdf_path.Foreground  = Brushes.Black
        self.txt_filename.Text        = fname
        self.txt_file_size.Text       = "PDF · {}".format(size_str)
        self.pnl_file_info.Visibility = Visibility.Visible
        self._update_status()

    def sort_changed(self, sender, args):
        if self._all_items:
            self._refresh_list()

    def search_changed(self, sender, args):
        if self._all_items:
            self._refresh_list()

    def view_selection_changed(self, sender, args):
        if self._loading:
            return
        self._refresh_pages()

    def select_all_clicked(self, sender, args):
        for item in self._items:
            item.IsSelected = True
        self._refresh_pages()

    def select_none_clicked(self, sender, args):
        for item in self._items:
            item.IsSelected = False
        self._refresh_pages()

    def import_clicked(self, sender, args):
        selected = [i for i in self._items if i.IsSelected]
        if not selected:
            forms.alert("No views selected.", title="PDF Import")
            return

        resolution = RESOLUTION_MAP.get(self.cmb_resolution.SelectedIndex, 300)
        doc        = revit.doc
        imported   = 0
        errors     = []

        try:
            with revit.Transaction("PDF Import"):
                for idx, item in enumerate(selected):
                    try:
                        view     = doc.GetElement(item.view_id)
                        options  = DB.ImageTypeOptions(self._pdf_path, False, DB.ImageTypeSource.Import)
                        options.Resolution = resolution
                        options.PageNumber = idx + 1
                        img_type  = DB.ImageType.Create(doc, options)
                        placement = DB.ImagePlacementOptions()
                        placement.Location = DB.XYZ.Zero
                        DB.ImageInstance.Create(doc, view, img_type.Id, placement)
                        imported += 1
                    except Exception as ex:
                        errors.append(u"Page {} → '{}': {}".format(idx + 1, item.Name, str(ex)))
        except Exception as ex:
            forms.alert(u"Transaction failed:\n{}".format(str(ex)), title="PDF Import Error")
            return

        if errors:
            forms.alert(
                u"Imported {} of {} page(s).\n\nErrors:\n{}".format(
                    imported, len(selected), u"\n".join(errors)),
                title="PDF Import — Partial"
            )
        else:
            forms.alert(u"Done! {} page(s) imported.".format(imported), title="PDF Import")
        self.Close()

    # ── Chrome ────────────────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, args):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, args):
        self.WindowState = (WindowState.Normal if self.WindowState == WindowState.Maximized
                            else WindowState.Maximized)

    def close_button_clicked(self, sender, args):
        self.Close()


def show_pdf_import():
    PDFImportDialog().ShowDialog()
