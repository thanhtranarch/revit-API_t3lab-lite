# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')

from System.Windows import WindowState, Visibility
from System.Windows.Forms import OpenFileDialog, DialogResult
from System.Windows.Threading import DispatcherPriority, DispatcherOperationCallback
from System.Collections.Generic import List as GList
from System import Type
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

from pyrevit import forms, DB, revit


RESOLUTION_MAP = {0: 150, 1: 300, 2: 600}

_TYPE_LABELS = {
    DB.ViewType.FloorPlan:    "Floor Plan",
    DB.ViewType.CeilingPlan:  "Ceiling Plan",
    DB.ViewType.Legend:       "Legend",
    DB.ViewType.DraftingView: "Drafting",
}

SUPPORTED_VIEW_TYPES = frozenset(_TYPE_LABELS.keys())

SORT_KEYS = {
    0: lambda v: v.Name,
    1: lambda v: (v.Type, v.Name),
}

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'PDFImport.xaml')


class ViewItem(object):
    """Lightweight data row — only display strings + ElementId, no Revit proxy."""

    def __init__(self, name, type_label, view_id):
        self.Name        = name
        self.Type        = type_label
        self.view_id     = view_id
        self.IsSelected  = True
        self.PageDisplay = u"–"
        self._page_no    = None

    def set_page(self, page):
        self._page_no = page
        self.PageDisplay = str(page) if page is not None else u"–"


class PDFImportDialog(forms.WPFWindow):
    """PDF Import tool window."""

    def __init__(self):
        # Initialise state BEFORE WPFWindow.__init__ so event guards work
        # during XAML load (ComboBox SelectionChanged fires then).
        self._pdf_path       = None
        self._all_items      = []
        self._filtered_items = []
        self._raw_views      = []
        self._ready          = False
        forms.WPFWindow.__init__(self, _XAML)
        # Defer heavy work until after the window has painted at least one frame.
        self.Loaded += self._on_loaded
        self._ready = True

    def _on_loaded(self, sender, args):
        """Window just rendered — queue the collector so the UI doesn't freeze (lag) during open."""
        if hasattr(self, 'txt_status'):
            self.txt_status.Text = "Loading views…"
        self._schedule(self._load_all_views)

    def _schedule(self, fn):
        """Queue the next step at Background priority so WPF repaints between collectors."""
        from System.Windows.Threading import DispatcherPriority
        from System.Windows.Threading import DispatcherOperationCallback
        self.Dispatcher.BeginInvoke(
            DispatcherPriority.Background,
            DispatcherOperationCallback(lambda _: fn() or None),
            None
        )

    def _load_all_views(self):
        """Single fast synchronous pass to collect and display views."""
        doc = revit.doc
        self._raw_views = []
        
        vt_floor    = DB.ViewType.FloorPlan
        vt_ceiling  = DB.ViewType.CeilingPlan
        vt_drafting = DB.ViewType.DraftingView
        vt_legend   = DB.ViewType.Legend
        
        try:
            collector = DB.FilteredElementCollector(doc)\
                          .OfCategory(DB.BuiltInCategory.OST_Views)\
                          .WhereElementIsNotElementType()
                          
            for v in collector:
                try:
                    # Match BatchOut's strict safety: Only allow specific safe subclasses.
                    # We drop support for generic DB.View (Legends) because corrupted
                    # internal views share this base class and crash Revit when queried.
                    if v.IsTemplate: continue
                    if not isinstance(v, (DB.ViewPlan, DB.ViewDrafting)): continue
                        
                    # Safely get name
                    v_name = v.Name
                    if not v_name:
                        v_name = "Unnamed View"
                        
                    if isinstance(v, DB.ViewPlan):
                        vt = getattr(v, 'ViewType', None)
                        if vt == vt_floor:
                            self._raw_views.append((v_name, _TYPE_LABELS[vt_floor], v.Id))
                        elif vt == vt_ceiling:
                            self._raw_views.append((v_name, _TYPE_LABELS[vt_ceiling], v.Id))
                    elif isinstance(v, DB.ViewDrafting):
                        self._raw_views.append((v_name, _TYPE_LABELS[vt_drafting], v.Id))
                except Exception:
                    pass
        except Exception as ex:
            import traceback
            from pyrevit import script
            logger = script.get_logger()
            logger.error("Error loading views: " + traceback.format_exc())
            
        self._finalize_load()

    def _finalize_load(self):
        """All collectors done — final sort, update ItemsSource."""
        self._raw_views.sort(key=lambda x: x[0])
        items = [ViewItem(n, t, vid) for n, t, vid in self._raw_views]

        self._all_items = items
        if hasattr(self, 'pnl_loading'):
            self.pnl_loading.Visibility = Visibility.Collapsed
        self._apply_filter(self.txt_search.Text)

    # ─── Filter / sort ───────────────────────────────────────────────────────

    def _apply_filter(self, query):
        """Filter + sort → single ItemsSource reassignment (list changed)."""
        q = (query or "").strip().lower()
        if q:
            self._filtered_items = [i for i in self._all_items if q in i.Name.lower()]
        else:
            self._filtered_items = list(self._all_items)
        idx = max(0, self.cmb_sort.SelectedIndex)
        self._filtered_items.sort(key=SORT_KEYS.get(idx, SORT_KEYS[0]))
        self._reassign_pages()
        self.grid_views.ItemsSource = self._filtered_items
        self._update_status_bar()

    def _apply_sort(self):
        """Sort in-place → lightweight Items.Refresh(), no rebind."""
        idx = max(0, self.cmb_sort.SelectedIndex)
        self._filtered_items.sort(key=SORT_KEYS.get(idx, SORT_KEYS[0]))
        self._reassign_pages()
        self.grid_views.Items.Refresh()

    def _reassign_pages(self):
        page = 1
        for item in self._filtered_items:
            if item.IsSelected:
                item.set_page(page)
                page += 1
            else:
                item.set_page(None)

    def _refresh_grid(self):
        """Recompute pages + light refresh, no rebind."""
        self._reassign_pages()
        self.grid_views.Items.Refresh()
        self._update_status_bar()

    def _update_status_bar(self):
        total    = len(self._filtered_items)
        selected = sum(1 for i in self._filtered_items if i.IsSelected)
        self.txt_view_count.Text     = "{} views".format(total)
        self.txt_selected_count.Text = "{} selected".format(selected)
        self.btn_import.IsEnabled    = bool(self._pdf_path and selected > 0)
        if not self._pdf_path:
            self.txt_status.Text = "Select a PDF file and choose target views"
        elif selected == 0:
            self.txt_status.Text = "Select at least one target view"
        else:
            self.txt_status.Text = "Ready — {} page(s) will be imported".format(selected)

    # ─── Events ──────────────────────────────────────────────────────────────

    def browse_pdf_clicked(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title = "Select PDF File"
        dlg.Filter = "PDF files (*.pdf)|*.pdf"
        dlg.Multiselect = False
        if dlg.ShowDialog() == DialogResult.OK:
            self._pdf_path = dlg.FileName
            fname = os.path.basename(self._pdf_path)
            fsize = os.path.getsize(self._pdf_path)
            size_kb = fsize / 1024.0
            size_str = (
                "{:.0f} KB".format(size_kb) if size_kb < 1024
                else "{:.1f} MB".format(size_kb / 1024.0)
            )
            self.txt_pdf_path.Text = fname
            from System.Windows.Media import Brushes
            self.txt_pdf_path.Foreground = Brushes.Black
            self.txt_filename.Text = fname
            self.txt_file_size.Text = "PDF · {}".format(size_str)
            self.pnl_file_info.Visibility = Visibility.Visible
            self._update_status_bar()

    def sort_changed(self, sender, args):
        if not self._ready:
            return
        self._apply_sort()

    def search_changed(self, sender, args):
        if not self._ready:
            return
        self._apply_filter(self.txt_search.Text)

    def view_selection_changed(self, sender, args):
        self._refresh_grid()

    def select_all_clicked(self, sender, args):
        for item in self._filtered_items:
            item.IsSelected = True
        self._refresh_grid()

    def select_none_clicked(self, sender, args):
        for item in self._filtered_items:
            item.IsSelected = False
        self._refresh_grid()

    def import_clicked(self, sender, args):
        selected = [i for i in self._filtered_items if i.IsSelected]
        if not selected:
            forms.alert("No views selected.", title="PDF Import")
            return

        resolution = RESOLUTION_MAP.get(self.cmb_resolution.SelectedIndex, 300)
        doc = revit.doc
        imported = 0
        errors = []
        total = len(selected)

        self.pnl_progress.Visibility = Visibility.Visible
        self.btn_import.IsEnabled = False
        self.btn_cancel.IsEnabled = False

        try:
            with revit.Transaction("PDF Import"):
                for idx, item in enumerate(selected):
                    page_number = idx + 1
                    try:
                        view = doc.GetElement(item.view_id)
                        options = DB.ImageTypeOptions(
                            self._pdf_path, False, DB.ImageTypeSource.Import
                        )
                        options.Resolution = resolution
                        options.PageNumber = page_number
                        image_type = DB.ImageType.Create(doc, options)
                        placement_options = DB.ImagePlacementOptions()
                        placement_options.Location = DB.XYZ.Zero
                        DB.ImageInstance.Create(
                            doc, view, image_type.Id, placement_options
                        )
                        imported += 1
                    except Exception as ex:
                        errors.append(
                            u"Page {} → '{}': {}".format(
                                page_number, item.Name, str(ex)
                            )
                        )

                    pct = int((idx + 1) * 100.0 / total)
                    self.txt_progress_pct.Text = "{}%".format(pct)
                    self.txt_status.Text = u"Importing… {}/{}".format(idx + 1, total)
                    track_w = self.pnl_progress_track.ActualWidth
                    if track_w > 0:
                        self.bar_progress.Width = track_w * (pct / 100.0)
                    self.UpdateLayout()

        except Exception as tx_err:
            forms.alert(
                u"Transaction failed:\n{}".format(str(tx_err)),
                title="PDF Import Error"
            )
        finally:
            self.pnl_progress.Visibility = Visibility.Collapsed
            self.btn_import.IsEnabled = True
            self.btn_cancel.IsEnabled = True

        if errors:
            forms.alert(
                u"Imported {} of {} page(s).\n\nErrors:\n{}".format(
                    imported, total, u"\n".join(errors)
                ),
                title="PDF Import — Partial"
            )
        else:
            forms.alert(
                u"Done! {} page(s) imported into {} view(s).".format(imported, total),
                title="PDF Import"
            )
        self.Close()

    # ─── Window chrome ───────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, args):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, args):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def close_button_clicked(self, sender, args):
        self.Close()


def show_pdf_import():
    PDFImportDialog().ShowDialog()
