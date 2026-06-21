# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')

from System import Action
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import WindowState, Visibility
from System.Windows.Controls import DataGridEditAction
from System.Windows.Forms import OpenFileDialog, DialogResult
from System.Windows.Threading import DispatcherPriority

from pyrevit import forms, DB, revit


RESOLUTION_MAP = {0: 150, 1: 300, 2: 600}

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'PDFImport.xaml')

_MODE_SEQUENTIAL = 0   # page N → view N
_MODE_ALL_IN_ONE = 1   # all pages → single view, placed in a horizontal strip

_GAP_FT = 0.15         # gap between images in all-in-one mode (feet, ~46 mm)
_FALLBACK_W_FT = 1.5   # fallback image width when Width property unavailable


class ViewItem(object):
    def __init__(self, name, type_label, view_id):
        self.Name        = name
        self.Type        = type_label
        self.view_id     = view_id
        self.IsSelected  = True
        self.PageNumber  = 0      # PDF page number assigned to this view (0 = none)
        self.PageDisplay = u"–"   # string shown in the PAGE column (user-editable)
        self.is_manual   = False  # True when user has hand-set the page number

    def set_page(self, n):
        self.PageNumber  = n if n else 0
        self.PageDisplay = str(n) if n else u"–"


class PDFImportDialog(forms.WPFWindow):

    def __init__(self):
        self._pdf_path  = None
        self._all_items = []
        self._items     = []
        self._loading   = False
        self._oc        = None   # ObservableCollection bound once, updated in-place
        self._mode      = _MODE_SEQUENTIAL
        forms.WPFWindow.__init__(self, _XAML)
        self.Loaded += self._on_loaded

    def _on_loaded(self, sender, args):
        # Bind the persistent OC once — never replace ItemsSource.
        # Replacing it forces WPF to destroy/recreate all row containers (the
        # root cause of the "empty grid" bug from the previous version).
        self._oc = ObservableCollection[object]()
        self.grid_views.ItemsSource = self._oc
        self.Dispatcher.BeginInvoke(DispatcherPriority.Background, Action(self._load_views))

    # ── Data ──────────────────────────────────────────────────────────────────

    def _load_views(self):
        items = []
        try:
            doc = revit.doc
            for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan):
                try:
                    if v.IsTemplate: continue
                    if v.ViewType == DB.ViewType.FloorPlan:
                        items.append(ViewItem(v.Name or "Unnamed", "Floor Plan", v.Id))
                    elif v.ViewType == DB.ViewType.CeilingPlan:
                        items.append(ViewItem(v.Name or "Unnamed", "Ceiling Plan", v.Id))
                except Exception:
                    pass
            for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewDrafting):
                try:
                    if v.IsTemplate: continue
                    items.append(ViewItem(v.Name or "Unnamed", "Drafting", v.Id))
                except Exception:
                    pass
        except Exception:
            pass

        items.sort(key=lambda x: x.Name)
        self._all_items = items
        self.pnl_loading.Visibility = Visibility.Collapsed
        self._refresh_list()

    def _refresh_list(self):
        q = self.txt_search.Text.strip().lower()
        result = [i for i in self._all_items if not q or q in i.Name.lower()]

        if self.cmb_sort.SelectedIndex == 1:
            result.sort(key=lambda i: (i.Type, i.Name))
        else:
            result.sort(key=lambda i: i.Name)

        self._items = result
        self._assign_pages()

        # Guard: WPF fires CheckBox.Checked during _oc.Add() because IsSelected
        # defaults to True; without this guard, view_selection_changed re-enters
        # _refresh_list for every added row.
        self._loading = True
        self._oc.Clear()
        for item in self._items:
            self._oc.Add(item)
        self._loading = False

        self._update_status()

    def _refresh_pages(self):
        self._assign_pages()
        # CommitEdit first: clicking a CheckBox in a non-ReadOnly DataGrid opens
        # an EditItem transaction; Items.Refresh() throws inside one.
        try:
            self.grid_views.CommitEdit()
            self.grid_views.Items.Refresh()
        except Exception:
            pass
        self._update_status()

    def _assign_pages(self):
        """Assign sequential page numbers to selected views.

        Views the user has hand-edited (is_manual=True) keep their value.
        Unselected views always show '–' regardless of manual state.
        """
        auto_page = 1
        for item in self._items:
            if item.IsSelected:
                if not item.is_manual:
                    item.set_page(auto_page)
                auto_page += 1
            else:
                # Deselected: always clear display; keep PageNumber for restore
                item.PageDisplay = u"–"

    def _update_status(self):
        total    = len(self._items)
        selected = [i for i in self._items if i.IsSelected]
        n_sel    = len(selected)

        self.txt_view_count.Text     = u"{} views".format(total)
        self.txt_selected_count.Text = u"{} selected".format(n_sel)
        self.btn_import.IsEnabled    = bool(self._pdf_path and n_sel > 0)

        if not self._pdf_path:
            self.txt_status.Text = u"Select a PDF file and choose target views"
        elif n_sel == 0:
            self.txt_status.Text = u"Select at least one target view"
        elif self._mode == _MODE_ALL_IN_ONE:
            self.txt_status.Text = u"All PDF pages → '{}'".format(selected[0].Name)
        else:
            self.txt_status.Text = u"Ready — {} page(s) will be imported".format(n_sel)

    def _set_ui_busy(self, busy):
        enabled = not busy
        self.rb_sequential.IsEnabled     = enabled
        self.rb_all_in_one.IsEnabled     = enabled
        self.btn_browse.IsEnabled        = enabled
        self.cmb_resolution.IsEnabled    = enabled
        self.cmb_sort.IsEnabled          = enabled
        self.txt_search.IsEnabled        = enabled
        self.btn_select_all.IsEnabled    = enabled and self._mode == _MODE_SEQUENTIAL
        self.btn_select_none.IsEnabled   = enabled
        self.btn_cancel.IsEnabled        = not busy
        if enabled:
            self.btn_import.IsEnabled = bool(
                self._pdf_path and any(i.IsSelected for i in self._items))
        else:
            self.btn_import.IsEnabled = False

    def _update_progress(self, done, total):
        if total <= 0:
            return
        try:
            track_w = self.pnl_progress_track.ActualWidth
            self.bar_progress.Width = (done / float(total)) * max(track_w, 0)
        except Exception:
            pass
        self.txt_progress_pct.Text = u"{}%".format(done * 100 // total)

    # ── Events ────────────────────────────────────────────────────────────────

    def mode_changed(self, sender, args):
        # Guard: XAML sets IsChecked="True" on rb_sequential before _on_loaded,
        # which fires this handler before _oc is initialised.
        if self._oc is None:
            return

        self._mode = (_MODE_SEQUENTIAL
                      if self.rb_sequential.IsChecked
                      else _MODE_ALL_IN_ONE)

        # Show PAGE column only in sequential mode
        page_col = self.grid_views.Columns[1]
        page_col.Visibility = (Visibility.Visible
                               if self._mode == _MODE_SEQUENTIAL
                               else Visibility.Collapsed)

        # Disable "All" button in all-in-one mode (single-select doesn't allow it)
        self.btn_select_all.IsEnabled = (self._mode == _MODE_SEQUENTIAL)

        if self._mode == _MODE_ALL_IN_ONE:
            # Enforce single-select: keep only the first currently-selected view
            first = next((i for i in self._items if i.IsSelected), None)
            self._loading = True
            for item in self._items:
                item.IsSelected = (item is first)
            self._loading = False
            try:
                self.grid_views.Items.Refresh()
            except Exception:
                pass

        self._update_status()

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
        size_str = (u"{:.0f} KB".format(size_kb) if size_kb < 1024
                    else u"{:.1f} MB".format(size_kb / 1024.0))
        self.txt_pdf_path.Text        = fname
        self.txt_filename.Text        = fname
        self.txt_file_size.Text       = u"PDF · {}".format(size_str)
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

        if self._mode == _MODE_ALL_IN_ONE:
            # Single-select: when a checkbox is checked, uncheck all others.
            # sender is the CheckBox — its DataContext is the ViewItem just changed.
            changed_item = getattr(sender, 'DataContext', None)
            if changed_item and getattr(changed_item, 'IsSelected', False):
                self._loading = True
                for item in self._items:
                    if item is not changed_item:
                        item.IsSelected = False
                self._loading = False
                try:
                    self.grid_views.Items.Refresh()
                except Exception:
                    pass
            self._update_status()
        else:
            self._refresh_pages()

    def select_all_clicked(self, sender, args):
        if self._mode == _MODE_ALL_IN_ONE:
            return
        self._loading = True
        for item in self._items:
            item.IsSelected = True
        self._loading = False
        self._refresh_pages()

    def select_none_clicked(self, sender, args):
        self._loading = True
        for item in self._items:
            item.IsSelected = False
        self._loading = False
        self._refresh_pages()

    def cell_edit_ending(self, sender, args):
        """Validate a manual page-number edit in the PAGE column (index 1)."""
        # Identify which column was edited
        try:
            col_idx = list(self.grid_views.Columns).index(args.Column)
        except ValueError:
            return
        if col_idx != 1:
            return
        if args.EditAction != DataGridEditAction.Commit:
            return

        item = args.Row.Item
        if not isinstance(item, ViewItem):
            return
        tb = args.EditingElement
        if tb is None:
            return

        raw = u""
        try:
            raw = str(tb.Text).strip()
            val = int(raw)
            if val > 0:
                item.PageDisplay = str(val)
                item.PageNumber  = val
                item.is_manual   = True
            else:
                args.Cancel = True  # reject zero or negative
        except Exception:
            args.Cancel = True  # reject non-integer input

        self._update_status()

    def import_clicked(self, sender, args):
        selected = [i for i in self._items if i.IsSelected]
        if not selected:
            forms.alert(u"No views selected.", title="PDF Import")
            return

        resolution = RESOLUTION_MAP.get(self.cmb_resolution.SelectedIndex, 300)
        doc        = revit.doc

        self._set_ui_busy(True)
        self.pnl_progress.Visibility = Visibility.Visible

        if self._mode == _MODE_ALL_IN_ONE:
            imported, errors = self._do_import_all_in_one(
                selected[0], resolution, doc)
            total = imported + len(errors)
        else:
            imported, errors, total = self._do_import_sequential(
                selected, resolution, doc)

        self.pnl_progress.Visibility = Visibility.Collapsed
        self._set_ui_busy(False)

        if errors:
            forms.alert(
                u"Imported {} of {} page(s).\n\nErrors:\n{}".format(
                    imported, total, u"\n".join(errors)),
                title=u"PDF Import — Partial")
        else:
            forms.alert(
                u"Done! {} page(s) imported successfully.".format(imported),
                title="PDF Import")
        self.Close()

    def _do_import_sequential(self, selected, resolution, doc):
        """Sequential mode: each selected view receives the PDF page in item.PageNumber."""
        imported = 0
        errors   = []
        total    = len(selected)
        self._update_progress(0, total)

        for item in selected:
            page_num = item.PageNumber if item.PageNumber > 0 else 1
            try:
                with revit.Transaction("PDF Import"):
                    view     = doc.GetElement(item.view_id)
                    options  = DB.ImageTypeOptions(
                        self._pdf_path, False, DB.ImageTypeSource.Import)
                    options.Resolution = resolution
                    options.PageNumber = page_num
                    img_type  = DB.ImageType.Create(doc, options)
                    placement = DB.ImagePlacementOptions()
                    placement.Location = DB.XYZ.Zero
                    DB.ImageInstance.Create(doc, view, img_type.Id, placement)
                    imported += 1
            except Exception as ex:
                errors.append(
                    u"Page {} → '{}': {}".format(page_num, item.Name, str(ex)))
            self._update_progress(imported, total)

        return imported, errors, total

    def _do_import_all_in_one(self, target_item, resolution, doc):
        """All-in-one mode: import every PDF page into one view, left-to-right strip.

        Uses ImageType.Width (internal feet) to calculate placement offsets so
        pages sit side-by-side without overlapping.
        """
        imported = 0
        errors   = []
        x_offset = 0.0
        page_num = 1

        view = doc.GetElement(target_item.view_id)

        while True:
            try:
                with revit.Transaction(
                        u"PDF Import — page {}".format(page_num)):
                    options = DB.ImageTypeOptions(
                        self._pdf_path, False, DB.ImageTypeSource.Import)
                    options.Resolution = resolution
                    options.PageNumber = page_num
                    img_type  = DB.ImageType.Create(doc, options)

                    # Width is in Revit internal units (feet).
                    try:
                        img_w = float(img_type.Width)
                        if img_w <= 0:
                            img_w = _FALLBACK_W_FT
                    except Exception:
                        img_w = _FALLBACK_W_FT

                    placement = DB.ImagePlacementOptions()
                    placement.Location = DB.XYZ(x_offset, 0, 0)
                    DB.ImageInstance.Create(doc, view, img_type.Id, placement)

                    x_offset += img_w + _GAP_FT
                    imported += 1
                    page_num += 1

            except Exception as ex:
                msg = str(ex)
                # Revit raises this specific message when PageNumber exceeds
                # the number of pages in the PDF — that's the normal end condition.
                if "PageNumber" in msg or "does not contain" in msg:
                    break
                errors.append(u"Page {}: {}".format(page_num, msg[:120]))
                break  # unexpected error — stop

        return imported, errors

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
