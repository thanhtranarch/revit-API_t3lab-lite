# -*- coding: utf-8 -*-
"""
BatchLinkDialog.py
==================
WPF Dialog for working with Revit links (.rvt), in three tabs:

* **Link Models**   — browse a folder, filter backups, batch link models.
* **Link Worksets** — list the worksets inside each loaded link and open/close
                      them in batch; applying reloads the links.
* **View Display**  — set how each link draws in the active view:
                      By Host View / By Linked View / Custom, plus visibility
                      and halftone.

Revit API work lives in ``Snippets/_links.py``; this module is UI only.

Part of T3Lab Extension.
Author: T3Lab
"""

import os
import re
import datetime

import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import WindowState, Visibility
from System.Windows.Forms import FolderBrowserDialog, DialogResult, NativeWindow
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri, UriKind, IntPtr

from Autodesk.Revit.DB import (
    AttachmentType,
    ElementId,
    ImportPlacement,
    IFailuresPreprocessor,
    FailureProcessingResult,
    FailureSeverity,
    LinkLoadResultType,
    ModelPathUtils,
    RevitLinkInstance,
    RevitLinkOptions,
    RevitLinkType,
    Transaction,
    TransactionGroup,
)

from pyrevit import revit

from GUI.WPF_Base import T3WPFWindow, to_items_source
from GUI.T3Dialog import confirm as t3_confirm

from Snippets import _links

GUI_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(GUI_DIR, 'Tools', 'BatchLink.xaml')

TAB_LINK = 0
TAB_WORKSETS = 1
TAB_DISPLAY = 2

PRIMARY_LABELS = {
    TAB_LINK: "Link Selected Models",
    TAB_WORKSETS: "Apply Worksets",
    TAB_DISPLAY: "Apply Display Settings",
}


import uuid


class WarningSwallower(IFailuresPreprocessor):
    """Suppresses non-fatal warnings (duplicate types, shared coordinates) while linking."""
    __namespace__ = "T3Lab.BatchLink_" + uuid.uuid4().hex[:8]

    def PreprocessFailures(self, failuresAccessor):
        fail_list = failuresAccessor.GetFailureMessages()
        if fail_list.Count == 0:
            return FailureProcessingResult.Continue

        has_error = False
        for failure in fail_list:
            severity = failure.GetSeverity()
            if severity == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(failure)
            elif severity == FailureSeverity.Error:
                has_error = True
                failuresAccessor.ResolveFailure(failure)

        if has_error:
            return FailureProcessingResult.ProceedWithCommit
        return FailureProcessingResult.Continue


# ── ROW ITEMS ────────────────────────────────────────────────────────────────

class _Row(object):
    """Common check/status plumbing for every grid row in this window."""

    def __init__(self, status_text="Ready", severity="Success",
                 is_selected=False, is_enabled=True):
        self._status_text = status_text
        self._severity = severity
        self._is_selected = bool(is_selected)
        self._is_enabled = bool(is_enabled)

    @property
    def StatusText(self):
        return self._status_text

    @StatusText.setter
    def StatusText(self, value):
        self._status_text = value

    @property
    def Severity(self):
        return self._severity

    @Severity.setter
    def Severity(self, value):
        self._severity = value

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = bool(value)

    @property
    def IsEnabled(self):
        return self._is_enabled

    @IsEnabled.setter
    def IsEnabled(self, value):
        self._is_enabled = bool(value)


class RevitModelItem(_Row):
    """A discovered .rvt file on disk, waiting to be linked."""

    def __init__(self, full_path, file_name, file_size, modified_time,
                 status_text="Ready", severity="Success", is_selected=True, is_enabled=True):
        _Row.__init__(self, status_text, severity, is_selected, is_enabled)
        self.full_path = full_path
        self.file_name = file_name
        self.file_size = file_size
        self.modified_time = modified_time

    @property
    def FileName(self):
        return self.file_name

    @property
    def FileSize(self):
        if self.file_size < 1024 * 1024:
            return "{:.1f} KB".format(self.file_size / 1024.0)
        return "{:.1f} MB".format(self.file_size / (1024.0 * 1024.0))

    @property
    def ModifiedFormatted(self):
        try:
            dt = datetime.datetime.fromtimestamp(self.modified_time)
            return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            return ""


class LinkWorksetRow(_Row):
    """One Revit link on the Worksets tab."""

    def __init__(self, record, worksets):
        can_manage = bool(record.is_loaded and record.link_doc is not None and worksets)
        if not record.is_loaded:
            status, severity = record.status, "Warning"
        elif not worksets:
            status, severity = "Not workshared", "Warning"
        else:
            status, severity = "Loaded", "Success"
        _Row.__init__(self, status, severity, is_selected=False, is_enabled=can_manage)
        self.record = record
        self.worksets = worksets

    @property
    def LinkName(self):
        return self.record.name

    @property
    def WorksetSummary(self):
        if not self.worksets:
            return "—"
        opened = sum(1 for w in self.worksets if w.is_open)
        return "{} / {}".format(opened, len(self.worksets))


class WorksetRow(_Row):
    """One workset name across the links being edited."""

    def __init__(self, name, is_open, present_in, total_links):
        _Row.__init__(self,
                      "Open" if is_open else "Closed",
                      "Success" if is_open else "Warning",
                      is_selected=is_open, is_enabled=True)
        self.name = name
        self.present_in = present_in
        self.total_links = total_links

    @property
    def WorksetName(self):
        return self.name

    @property
    def Coverage(self):
        return "{} / {}".format(self.present_in, self.total_links)

    def refresh_status(self):
        self._status_text = "Open" if self._is_selected else "Closed"
        self._severity = "Success" if self._is_selected else "Warning"


class LinkDisplayRow(_Row):
    """One Revit link on the View Display tab."""

    def __init__(self, record, state, linked_view_name):
        enabled = record.instance_id != ElementId.InvalidElementId
        if not record.is_loaded:
            status, severity = record.status, "Warning"
        else:
            status, severity = "Loaded", "Success"
        _Row.__init__(self, status, severity, is_selected=False, is_enabled=enabled)
        self.record = record
        self.state = state
        self.linked_view_name = linked_view_name

    @property
    def LinkName(self):
        return self.record.name

    @property
    def DisplayMode(self):
        return _links.DISPLAY_MODES[self.state.mode_index]

    @property
    def LinkedViewName(self):
        return self.linked_view_name or "—"

    @property
    def GraphicsSummary(self):
        parts = ["Visible" if self.state.visible else "Hidden"]
        if self.state.halftone:
            parts.append("halftone")
        return ", ".join(parts)


def _show_dialog_owned(window, dlg):
    """Display a WinForms dialog properly parented to the WPF window."""
    try:
        hwnd = WindowInteropHelper(window).Handle
    except Exception:
        hwnd = IntPtr.Zero
    if hwnd == IntPtr.Zero:
        return dlg.ShowDialog()
    owner = NativeWindow()
    owner.AssignHandle(hwnd)
    try:
        return dlg.ShowDialog(owner)
    finally:
        try:
            owner.ReleaseHandle()
        except Exception:
            pass


class BatchLinkDialog(T3WPFWindow):
    """Main Window class for Batch Link Revit Models."""

    def __init__(self, doc=None):
        T3WPFWindow.__init__(self, XAML_FILE)
        self.doc = doc or getattr(revit, 'doc', None)
        self._all_files = []
        self._filtered_files = []
        self._is_busy = False
        self._active_tab = TAB_LINK

        # Worksets tab state
        self._ws_link_rows = []
        self._ws_rows = []
        self._ws_filtered = []
        self._ws_focus = None

        # Display tab state
        self._disp_rows = []
        self._disp_filtered = []
        self._disp_focus = None
        self._disp_views = []          # [(id_int, label)] for the focused link
        self._suspend_mode_event = True

        self._load_logo()
        self._init_options()
        self._init_display_options()
        self._init_initial_path()
        self._suspend_mode_event = False

    # ── SETUP ────────────────────────────────────────────────────────────────

    def _load_logo(self):
        """Load and bind the T3Lab logo to title bar and window icon."""
        try:
            logo_path = os.path.join(GUI_DIR, 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.CacheOption = BitmapCacheOption.OnLoad
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                bitmap.Freeze()
                if hasattr(self, 'logo_image') and self.logo_image:
                    self.logo_image.Source = bitmap
                self.Icon = bitmap
        except Exception:
            pass

    def _init_options(self):
        """Initialize positioning, attachment, and path type dropdowns."""
        self.cb_positioning.ItemsSource = to_items_source([
            "Auto - Internal Origin to Origin",
            "Auto - By Shared Coordinates",
            "Auto - Project Base Point to Project Base Point",
            "Auto - Center to Center"
        ])
        self.cb_positioning.SelectedIndex = 0

        self.cb_attachment.ItemsSource = to_items_source(["Overlay", "Attachment"])
        self.cb_attachment.SelectedIndex = 0

        self.cb_path_type.ItemsSource = to_items_source(["Absolute", "Relative"])
        self.cb_path_type.SelectedIndex = 0

    def _init_display_options(self):
        """Fill the display-settings dropdowns and warn on pre-2025 Revit."""
        self.cb_disp_mode.ItemsSource = to_items_source(list(_links.DISPLAY_MODES))
        self.cb_disp_mode.SelectedIndex = 0

        for key, _label, _kind, allows_custom in _links.CUSTOM_ASPECTS:
            combo = getattr(self, 'cb_asp_' + key, None)
            if combo is None:
                continue
            combo.ItemsSource = to_items_source(list(_links.aspect_modes(allows_custom)))
            combo.SelectedIndex = 1  # By Linked View is the useful default

        if not _links.display_api_available():
            self.banner_no_api.Visibility = Visibility.Visible
            self.cb_disp_mode.IsEnabled = False
            self.cb_disp_view.IsEnabled = False
            self.btn_preset_model.IsEnabled = False
            self.btn_preset_annotation.IsEnabled = False
            self.btn_preset_custom.IsEnabled = False

    def _init_initial_path(self):
        """Set default search path from active document location."""
        initial_dir = ""
        if self.doc:
            try:
                path = self.doc.PathName
                if path and os.path.exists(os.path.dirname(path)):
                    initial_dir = os.path.dirname(path)
            except Exception:
                pass
        if not initial_dir:
            initial_dir = os.path.expanduser("~\\Documents")

        self.tb_folder_path.Text = initial_dir
        self._scan_folder()

    # ── SHARED HELPERS ───────────────────────────────────────────────────────

    def _set_status(self, text):
        self.status_text.Text = text

    def _active_view(self):
        if not self.doc:
            return None
        try:
            return self.doc.ActiveView
        except Exception:
            return None

    def _begin_busy(self, phase):
        self._is_busy = True
        self.lbl_phase.Text = phase
        self.lbl_current_file.Text = ""
        self.progress_bar.Value = 0
        self.pnl_progress.Visibility = Visibility.Visible
        self.btn_primary.IsEnabled = False
        self.btn_cancel.IsEnabled = False
        self._do_events()

    def _step_busy(self, phase, current, index, total):
        self.lbl_phase.Text = "{} — {} / {}".format(phase, index, total)
        self.lbl_current_file.Text = current
        self.progress_bar.Value = int((float(index - 1) / max(1, total)) * 100)
        self._do_events()

    def _end_busy(self, summary):
        self._is_busy = False
        self.progress_bar.Value = 100
        self.btn_primary.IsEnabled = True
        self.btn_cancel.IsEnabled = True
        self._set_status(summary)
        self._do_events()

    # ── TAB 1: FOLDER SCAN ───────────────────────────────────────────────────

    def _scan_folder(self):
        """Scan the selected folder for .rvt files."""
        folder = self.tb_folder_path.Text.strip()
        self._all_files = []

        if not folder or not os.path.isdir(folder):
            self.lbl_file_count.Text = "0 models found"
            self._set_status("Select a valid folder")
            self._apply_filter()
            return

        include_subfolders = bool(self.chk_subfolders.IsChecked)
        ignore_backups = bool(self.chk_ignore_backups.IsChecked)
        backup_pattern = re.compile(r'\.\d{3,4}\.rvt$', re.IGNORECASE)

        existing_paths, existing_names = _links.existing_link_keys(self.doc)

        active_doc_path = ""
        if self.doc:
            try:
                if self.doc.PathName:
                    active_doc_path = os.path.normpath(self.doc.PathName).lower()
            except Exception:
                pass

        try:
            if include_subfolders:
                file_generator = (
                    os.path.join(root, f)
                    for root, _, files in os.walk(folder)
                    for f in files
                )
            else:
                file_generator = (
                    os.path.join(folder, f)
                    for f in os.listdir(folder)
                )

            for full_path in file_generator:
                if not full_path.lower().endswith('.rvt'):
                    continue
                file_name = os.path.basename(full_path)
                if file_name.startswith('~$'):
                    continue
                if ignore_backups and backup_pattern.search(file_name):
                    continue

                try:
                    stat = os.stat(full_path)
                    f_size = stat.st_size
                    m_time = stat.st_mtime
                except Exception:
                    f_size = 0
                    m_time = 0

                norm_path = os.path.normpath(full_path).lower()
                clean_name = file_name.lower()

                if active_doc_path and norm_path == active_doc_path:
                    status, severity, is_sel, is_enabled = "Host Model", "Warning", False, False
                elif norm_path in existing_paths or clean_name in existing_names:
                    status, severity, is_sel, is_enabled = "Already Linked", "Warning", False, True
                else:
                    status, severity, is_sel, is_enabled = "Ready", "Success", True, True

                self._all_files.append(RevitModelItem(
                    full_path=full_path,
                    file_name=file_name,
                    file_size=f_size,
                    modified_time=m_time,
                    status_text=status,
                    severity=severity,
                    is_selected=is_sel,
                    is_enabled=is_enabled
                ))

        except Exception as ex:
            self._set_status("Error reading folder: {}".format(str(ex)[:60]))

        self._all_files.sort(key=lambda x: x.FileName.lower())
        self._apply_filter()

    def _apply_filter(self):
        """Apply the search filter to the discovered models and refresh the grid."""
        query = self.tb_filter.Text.strip().lower()
        if not query:
            self._filtered_files = list(self._all_files)
        else:
            self._filtered_files = [f for f in self._all_files if query in f.FileName.lower()]

        self.grid_files.ItemsSource = to_items_source(self._filtered_files)
        has_items = len(self._filtered_files) > 0
        self.txt_empty.Visibility = Visibility.Collapsed if has_items else Visibility.Visible

        total_found = len(self._all_files)
        showing = len(self._filtered_files)
        if query:
            self.lbl_file_count.Text = "Showing {} / {} models".format(showing, total_found)
        else:
            self.lbl_file_count.Text = "{} models found".format(total_found)

        self._update_counts()

    def _update_counts(self):
        """Update the footer status and the header check state on tab 1."""
        selected_count = sum(1 for f in self._all_files if f.IsSelected and f.IsEnabled)

        if self._active_tab == TAB_LINK and not self._is_busy:
            self.btn_primary.Content = ("Link 1 Model" if selected_count == 1
                                       else "Link {} Models".format(selected_count))
            self._set_status("Ready — {} models selected".format(selected_count))

        self._sync_header_checkbox(self.chk_header, self._filtered_files)

    @staticmethod
    def _sync_header_checkbox(checkbox, rows):
        usable = [r for r in rows if r.IsEnabled]
        if not usable:
            checkbox.IsChecked = False
            return
        if all(r.IsSelected for r in usable):
            checkbox.IsChecked = True
        elif any(r.IsSelected for r in usable):
            checkbox.IsChecked = None
        else:
            checkbox.IsChecked = False

    def _refresh_grid(self):
        self.grid_files.ItemsSource = to_items_source(self._filtered_files)

    # ── TAB 2: LINK WORKSETS ─────────────────────────────────────────────────

    def _load_link_worksets(self):
        """Rebuild the link list on the Worksets tab."""
        self._ws_link_rows = []
        self._ws_focus = None
        records = _links.collect_links(self.doc)
        for rec in records:
            worksets = _links.get_link_worksets(rec.link_doc)
            self._ws_link_rows.append(LinkWorksetRow(rec, worksets))

        self.grid_ws_links.ItemsSource = to_items_source(self._ws_link_rows)
        has_items = len(self._ws_link_rows) > 0
        self.txt_ws_links_empty.Visibility = Visibility.Collapsed if has_items else Visibility.Visible

        manageable = sum(1 for r in self._ws_link_rows if r.IsEnabled)
        self.lbl_ws_link_count.Text = "{} links · {} with worksets".format(
            len(self._ws_link_rows), manageable)

        first = next((r for r in self._ws_link_rows if r.IsEnabled), None)
        if first is not None:
            first.IsSelected = True
            self._ws_focus = first
            self.grid_ws_links.SelectedItem = first
        else:
            self._ws_focus = None
        self._rebuild_workset_rows()

    def _ws_checked_rows(self):
        """Links the workset changes will be applied to."""
        if bool(self.chk_ws_apply_all_selected.IsChecked):
            checked = [r for r in self._ws_link_rows if r.IsSelected and r.IsEnabled]
            if checked:
                return checked
        if self._ws_focus is not None and self._ws_focus.IsEnabled:
            return [self._ws_focus]
        return []

    def _rebuild_workset_rows(self):
        """Rebuild the workset matrix from the focused link, scored over the checked links."""
        self._ws_rows = []
        focus = self._ws_focus

        if focus is None or not focus.worksets:
            self.lbl_ws_focus.Text = "Select a link on the left to list its worksets."
            self._apply_ws_filter()
            return

        targets = self._ws_checked_rows()
        if focus not in targets:
            targets = [focus] + targets
        total = len(targets)

        presence = {}
        for row in targets:
            for ws in row.worksets:
                presence[ws.name] = presence.get(ws.name, 0) + 1

        for ws in focus.worksets:
            self._ws_rows.append(WorksetRow(ws.name, ws.is_open,
                                            presence.get(ws.name, 1), total))

        self.lbl_ws_focus.Text = "Worksets of {} — changes apply to {} link{}".format(
            focus.LinkName, total, "" if total == 1 else "s")
        self._apply_ws_filter()

    def _apply_ws_filter(self):
        query = self.tb_ws_filter.Text.strip().lower()
        if not query:
            self._ws_filtered = list(self._ws_rows)
        else:
            self._ws_filtered = [w for w in self._ws_rows if query in w.WorksetName.lower()]

        self.grid_worksets.ItemsSource = to_items_source(self._ws_filtered)
        has_items = len(self._ws_filtered) > 0
        self.txt_ws_empty.Visibility = Visibility.Collapsed if has_items else Visibility.Visible

        opened = sum(1 for w in self._ws_rows if w.IsSelected)
        self.lbl_ws_count.Text = "{} / {} open".format(opened, len(self._ws_rows))
        self._sync_header_checkbox(self.chk_ws_header, self._ws_filtered)

        if self._active_tab == TAB_WORKSETS and not self._is_busy:
            links = len(self._ws_checked_rows())
            self._set_status("Ready — {} worksets open on {} link{}".format(
                opened, links, "" if links == 1 else "s"))

    def _refresh_ws_grid(self):
        for row in self._ws_rows:
            row.refresh_status()
        self.grid_worksets.ItemsSource = to_items_source(self._ws_filtered)

    def _apply_worksets(self):
        """Reload each targeted link with the workset list from the grid."""
        targets = self._ws_checked_rows()
        if not targets:
            self._set_status("Select at least one loaded, workshared link first.")
            return

        wanted_open = set(w.WorksetName for w in self._ws_rows if w.IsSelected)
        wanted_closed = set(w.WorksetName for w in self._ws_rows if not w.IsSelected)
        if not wanted_open:
            if not t3_confirm(
                    "No workset is checked, so every listed workset will be closed in "
                    "{} link{}.".format(len(targets), "" if len(targets) == 1 else "s"),
                    title="Close all worksets?", ok_text="Close them", owner=self):
                return

        self._begin_busy("Reloading links")
        total = len(targets)
        succeeded = 0
        failed = 0

        for index, row in enumerate(targets, 1):
            self._step_busy("Reloading links", row.LinkName, index, total)

            open_ids, close_ids = _links.split_workset_ids(
                row.worksets, wanted_open, wanted_closed)
            if not open_ids and not close_ids:
                row.StatusText = "No match"
                row.Severity = "Warning"
                continue

            ok, message = _links.apply_link_worksets(row.record.link_type, open_ids, close_ids)
            if ok:
                succeeded += 1
                row.StatusText = "Reloaded"
                row.Severity = "Success"
            else:
                failed += 1
                row.StatusText = message[:22] or "Failed"
                row.Severity = "Danger"
            self._do_events()

        self._refresh_ws_link_rows(targets)
        self._end_busy("Worksets applied: {} link{} reloaded, {} failed".format(
            succeeded, "" if succeeded == 1 else "s", failed))

    def _refresh_ws_link_rows(self, rows):
        """Re-read the worksets of the given links, keeping their result pills."""
        for row in rows:
            try:
                row.record.link_doc = row.record.instance.GetLinkDocument()
            except Exception:
                pass
            row.worksets = _links.get_link_worksets(row.record.link_doc)
        self.grid_ws_links.ItemsSource = to_items_source(self._ws_link_rows)
        if self._ws_focus is not None:
            for ws_row in self._ws_rows:
                match = next((w for w in self._ws_focus.worksets
                              if w.name == ws_row.WorksetName), None)
                if match is not None:
                    ws_row.IsSelected = match.is_open
            self._refresh_ws_grid()
            self._apply_ws_filter()

    # ── TAB 3: VIEW DISPLAY ──────────────────────────────────────────────────

    def _load_display_rows(self):
        """Rebuild the link list on the View Display tab for the active view."""
        self._disp_rows = []
        self._disp_focus = None
        view = self._active_view()

        if view is None:
            self.lbl_disp_scope.Text = "No active view"
            self.grid_disp_links.ItemsSource = to_items_source([])
            self.txt_disp_empty.Visibility = Visibility.Visible
            return

        self.lbl_disp_scope.Text = "Active view: {}".format(_links.element_name(view))

        for rec in _links.collect_links(self.doc):
            state = _links.get_link_display(view, rec)
            view_name = ""
            if state.mode_index != 0 and int(state.linked_view_id) > 0 and rec.link_doc is not None:
                try:
                    linked_view = rec.link_doc.GetElement(
                        _links.new_element_id(int(state.linked_view_id)))
                    view_name = _links.element_name(linked_view)
                except Exception:
                    view_name = ""
            self._disp_rows.append(LinkDisplayRow(rec, state, view_name))

        self._apply_disp_filter()

        # Focus a loaded link first: an unloaded one has no views to offer the
        # By Linked View preset, which makes the panel look broken on open.
        first = next((r for r in self._disp_rows
                      if r.IsEnabled and r.record.link_doc is not None), None)
        if first is None:
            first = next((r for r in self._disp_rows if r.IsEnabled), None)
        if first is not None:
            first.IsSelected = True
            self.grid_disp_links.SelectedItem = first
            if self._disp_focus is not first:
                self._focus_display_row(first)

    def _apply_disp_filter(self):
        query = self.tb_disp_filter.Text.strip().lower()
        if not query:
            self._disp_filtered = list(self._disp_rows)
        else:
            self._disp_filtered = [r for r in self._disp_rows if query in r.LinkName.lower()]

        self.grid_disp_links.ItemsSource = to_items_source(self._disp_filtered)
        has_items = len(self._disp_filtered) > 0
        self.txt_disp_empty.Visibility = Visibility.Collapsed if has_items else Visibility.Visible

        self.lbl_disp_count.Text = "{} links".format(len(self._disp_rows))
        self._sync_header_checkbox(self.chk_disp_header, self._disp_filtered)
        self._update_disp_status()

    def _update_disp_status(self):
        if self._active_tab != TAB_DISPLAY or self._is_busy:
            return
        count = len(self._disp_checked_rows())
        self._set_status("Ready — {} link{} selected in this view".format(
            count, "" if count == 1 else "s"))

    def _disp_checked_rows(self):
        checked = [r for r in self._disp_rows if r.IsSelected and r.IsEnabled]
        if checked:
            return checked
        if self._disp_focus is not None and self._disp_focus.IsEnabled:
            return [self._disp_focus]
        return []

    def _focus_display_row(self, row):
        """Load the focused link's views and mirror its current settings into the panel."""
        self._disp_focus = row
        self._disp_views = []
        self.cb_disp_view.ItemsSource = to_items_source([])

        if row is None:
            return

        view = self._active_view()
        if view is not None and row.record.link_doc is not None:
            try:
                self._disp_views = _links.get_link_views(row.record.link_doc, view.ViewType)
            except Exception:
                self._disp_views = []
            if not self._disp_views:
                self._disp_views = _links.get_link_views(row.record.link_doc)

        self.cb_disp_view.ItemsSource = to_items_source([label for _vid, label in self._disp_views])

        self._suspend_mode_event = True
        try:
            self.cb_disp_mode.SelectedIndex = row.state.mode_index
            current = int(row.state.linked_view_id)
            selected = -1
            for index, (vid, _label) in enumerate(self._disp_views):
                if int(vid) == current:
                    selected = index
                    break
            self.cb_disp_view.SelectedIndex = selected if selected >= 0 else (
                0 if self._disp_views else -1)

            for key, _label, _kind, allows_custom in _links.CUSTOM_ASPECTS:
                combo = getattr(self, 'cb_asp_' + key, None)
                if combo is None:
                    continue
                value = int(row.state.aspects.get(key, 0))
                limit = len(_links.aspect_modes(allows_custom)) - 1
                combo.SelectedIndex = max(0, min(limit, value))

            self.chk_disp_visible.IsChecked = bool(row.state.visible)
            self.chk_disp_halftone.IsChecked = bool(row.state.halftone)
        finally:
            self._suspend_mode_event = False

        self._sync_custom_panel()

    def _sync_custom_panel(self):
        is_custom = (self.cb_disp_mode.SelectedIndex == 2)
        self.pnl_custom.Visibility = Visibility.Visible if is_custom else Visibility.Collapsed
        needs_view = self.cb_disp_mode.SelectedIndex in (1, 2)
        self.cb_disp_view.IsEnabled = needs_view and bool(self._disp_views)

    def _selected_view_label(self):
        index = self.cb_disp_view.SelectedIndex
        if index < 0 or index >= len(self._disp_views):
            return None
        return self._disp_views[index][1]

    def _resolve_linked_view(self, row, label):
        """Find the view id inside `row`'s link whose label matches `label`."""
        if not label or row.record.link_doc is None:
            return -1
        view = self._active_view()
        candidates = []
        if view is not None:
            candidates = _links.get_link_views(row.record.link_doc, view.ViewType)
        if not candidates:
            candidates = _links.get_link_views(row.record.link_doc)
        for vid, vlabel in candidates:
            if vlabel == label:
                return vid
        return -1

    def _apply_display(self):
        """Write visibility, halftone and display settings to the active view."""
        view = self._active_view()
        if view is None:
            self._set_status("No active view — open a view before applying.")
            return

        targets = self._disp_checked_rows()
        if not targets:
            self._set_status("Select at least one link first.")
            return

        mode_index = max(0, self.cb_disp_mode.SelectedIndex)
        view_label = self._selected_view_label()
        api_ok = _links.display_api_available()

        if api_ok and mode_index == 1 and not view_label:
            self._set_status("Pick a view inside the link before using By Linked View.")
            return

        aspects = {}
        if mode_index == 2:
            for key, _label, _kind, _custom in _links.CUSTOM_ASPECTS:
                combo = getattr(self, 'cb_asp_' + key, None)
                if combo is not None:
                    aspects[key] = max(0, combo.SelectedIndex)

        want_visible = bool(self.chk_disp_visible.IsChecked)
        want_halftone = bool(self.chk_disp_halftone.IsChecked)

        self._begin_busy("Applying display settings")
        total = len(targets)
        succeeded = 0
        failed = 0

        transaction = Transaction(self.doc, "Batch Link — link display settings")
        try:
            transaction.Start()
            for index, row in enumerate(targets, 1):
                self._step_busy("Applying display settings", row.LinkName, index, total)

                messages = []
                ok = True

                vis_ok, vis_msg = _links.set_link_visibility(view, row.record, want_visible)
                if not vis_ok:
                    ok = False
                    messages.append(vis_msg)

                if want_visible:
                    ht_ok, ht_msg = _links.set_link_halftone(view, row.record, want_halftone)
                    if not ht_ok:
                        ok = False
                        messages.append(ht_msg)

                if api_ok:
                    linked_view_id = -1
                    if mode_index in (1, 2):
                        linked_view_id = self._resolve_linked_view(row, view_label)
                        if mode_index == 1 and linked_view_id <= 0:
                            ok = False
                            messages.append("No matching view in link")
                    if ok or mode_index == 0:
                        disp_ok, disp_msg = _links.set_link_display(
                            view, row.record, mode_index, linked_view_id, aspects)
                        if not disp_ok:
                            ok = False
                            messages.append(disp_msg)

                if ok:
                    succeeded += 1
                    row.StatusText = "Applied"
                    row.Severity = "Success"
                else:
                    failed += 1
                    text = (messages[0] if messages else "Failed")
                    row.StatusText = text[:22]
                    row.Severity = "Danger"
                self._do_events()

            transaction.Commit()
        except Exception as ex:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
            self._end_busy("Display settings failed: {}".format(str(ex)[:60]))
            return

        self._refresh_disp_rows(view, targets)
        self._end_busy("Display applied: {} link{} updated, {} failed".format(
            succeeded, "" if succeeded == 1 else "s", failed))

    def _refresh_disp_rows(self, view, rows):
        """Re-read display settings for the given links, keeping their result pills."""
        for row in rows:
            row.state = _links.get_link_display(view, row.record)
            row.linked_view_name = ""
            if int(row.state.linked_view_id) > 0 and row.record.link_doc is not None:
                try:
                    linked_view = row.record.link_doc.GetElement(
                        _links.new_element_id(int(row.state.linked_view_id)))
                    row.linked_view_name = _links.element_name(linked_view)
                except Exception:
                    row.linked_view_name = ""
        self.grid_disp_links.ItemsSource = to_items_source(self._disp_filtered)

    # ── TAB 1: LINKING ───────────────────────────────────────────────────────

    def _link_models(self):
        """Execute batch linking of the selected Revit models."""
        selected_items = [f for f in self._all_files if f.IsSelected and f.IsEnabled]
        if not selected_items:
            self._set_status("Please select at least one Revit model to link.")
            return

        pin_links = bool(self.chk_pin.IsChecked)
        skip_existing = bool(self.chk_skip_existing.IsChecked)

        placement_map = {
            0: ImportPlacement.Origin,
            1: ImportPlacement.Shared,
            2: ImportPlacement.Site,
            3: ImportPlacement.Centered,
        }
        placement = placement_map.get(self.cb_positioning.SelectedIndex, ImportPlacement.Origin)

        attachment_map = {0: AttachmentType.Overlay, 1: AttachmentType.Attachment}
        attachment_type = attachment_map.get(self.cb_attachment.SelectedIndex,
                                             AttachmentType.Overlay)
        is_relative = (self.cb_path_type.SelectedIndex == 1)

        self._begin_busy("Linking models")
        total = len(selected_items)
        success_count = 0
        skipped_count = 0
        failed_count = 0

        tg = TransactionGroup(self.doc, "Batch Link Revit Models")
        tg.Start()

        try:
            existing_paths, existing_names = _links.existing_link_keys(self.doc)

            for idx, item in enumerate(selected_items, 1):
                self._step_busy("Linking models", item.FileName, idx, total)

                norm_path = os.path.normpath(item.full_path).lower()
                clean_name = item.FileName.lower()

                if skip_existing and (norm_path in existing_paths or clean_name in existing_names):
                    item.StatusText = "Skipped"
                    item.Severity = "Warning"
                    skipped_count += 1
                    continue

                t = Transaction(self.doc, "Link: {}".format(item.FileName))
                try:
                    t.Start()
                    fail_opts = t.GetFailureHandlingOptions()
                    fail_opts.SetFailuresPreprocessor(WarningSwallower())
                    fail_opts.SetClearAfterRollback(True)
                    t.SetFailureHandlingOptions(fail_opts)

                    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(item.full_path)
                    link_opts = RevitLinkOptions(is_relative)
                    load_result = RevitLinkType.Create(self.doc, model_path, link_opts)

                    if load_result and (
                        load_result.LoadResult == LinkLoadResultType.LinkLoaded or
                        load_result.LoadResult == LinkLoadResultType.UsedExisting
                    ):
                        link_type_id = load_result.ElementId
                        if link_type_id != ElementId.InvalidElementId:
                            link_type = self.doc.GetElement(link_type_id)
                            if link_type and hasattr(link_type, 'AttachmentType'):
                                try:
                                    link_type.AttachmentType = attachment_type
                                except Exception:
                                    pass
                            instance = RevitLinkInstance.Create(self.doc, link_type_id, placement)
                            if instance and pin_links:
                                try:
                                    instance.Pinned = True
                                except Exception:
                                    pass
                            t.Commit()
                            item.StatusText = "Linked"
                            item.Severity = "Success"
                            success_count += 1
                            existing_paths.add(norm_path)
                            existing_names.add(clean_name)
                        else:
                            t.RollBack()
                            item.StatusText = "Invalid Type"
                            item.Severity = "Danger"
                            failed_count += 1
                    else:
                        t.RollBack()
                        res_name = str(load_result.LoadResult) if load_result else "Load Failed"
                        item.StatusText = res_name
                        item.Severity = "Danger"
                        failed_count += 1

                except Exception as ex:
                    if t.HasStarted() and not t.HasEnded():
                        t.RollBack()
                    err = str(ex).split("\n")[0]
                    if len(err) > 25:
                        err = err[:22] + "..."
                    item.StatusText = err or "Error"
                    item.Severity = "Danger"
                    failed_count += 1

            self.lbl_current_file.Text = "All operations finished"

            if success_count > 0:
                tg.Assimilate()
            else:
                tg.RollBack()

        except Exception as ex:
            if tg.HasStarted() and not tg.HasEnded():
                tg.RollBack()
            self._end_busy("Batch link failed: {}".format(str(ex)[:60]))
            return

        self._end_busy("Completed: {} linked, {} skipped, {} failed".format(
            success_count, skipped_count, failed_count))
        self._refresh_grid()

    # ── EVENT HANDLERS: WINDOW ───────────────────────────────────────────────

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

    def cancel_button_clicked(self, sender, e):
        self.Close()

    # ── EVENT HANDLERS: TABS ─────────────────────────────────────────────────

    def tab_chip_checked(self, sender, e):
        """Switch tabs from the pill bar, loading each tab's data on first use."""
        if self._is_busy:
            return
        try:
            index = int(sender.Tag)
        except Exception:
            index = TAB_LINK

        self._active_tab = index
        self.tab_control.SelectedIndex = index
        self.btn_primary.Content = PRIMARY_LABELS.get(index, "Apply")

        if index == TAB_LINK:
            self._update_counts()
        elif index == TAB_WORKSETS:
            self._load_link_worksets()
        elif index == TAB_DISPLAY:
            self._load_display_rows()

    def primary_button_clicked(self, sender, e):
        if self._is_busy:
            return
        if not self.doc:
            self._set_status("No active Revit document found.")
            return
        if self._active_tab == TAB_LINK:
            self._link_models()
        elif self._active_tab == TAB_WORKSETS:
            self._apply_worksets()
        elif self._active_tab == TAB_DISPLAY:
            self._apply_display()

    # ── EVENT HANDLERS: TAB 1 ────────────────────────────────────────────────

    def browse_folder_clicked(self, sender, e):
        dlg = FolderBrowserDialog()
        dlg.Description = "Select folder containing Revit models (.rvt) to link"
        if os.path.isdir(self.tb_folder_path.Text.strip()):
            dlg.SelectedPath = self.tb_folder_path.Text.strip()

        if _show_dialog_owned(self, dlg) == DialogResult.OK:
            self.tb_folder_path.Text = dlg.SelectedPath
            self._scan_folder()

    def folder_path_changed(self, sender, e):
        path = self.tb_folder_path.Text.strip()
        if os.path.isdir(path):
            self._scan_folder()

    def subfolders_toggled(self, sender, e):
        self._scan_folder()

    def ignore_backups_toggled(self, sender, e):
        self._scan_folder()

    def filter_text_changed(self, sender, e):
        self._apply_filter()

    def select_all_clicked(self, sender, e):
        for f in self._filtered_files:
            if f.IsEnabled:
                f.IsSelected = True
        self._refresh_grid()
        self._update_counts()

    def select_none_clicked(self, sender, e):
        for f in self._filtered_files:
            if f.IsEnabled:
                f.IsSelected = False
        self._refresh_grid()
        self._update_counts()

    def invert_selection_clicked(self, sender, e):
        for f in self._filtered_files:
            if f.IsEnabled:
                f.IsSelected = not f.IsSelected
        self._refresh_grid()
        self._update_counts()

    def header_checkbox_clicked(self, sender, e):
        check_val = bool(self.chk_header.IsChecked)
        for f in self._filtered_files:
            if f.IsEnabled:
                f.IsSelected = check_val
        self._refresh_grid()
        self._update_counts()

    def row_checkbox_clicked(self, sender, e):
        self._update_counts()

    # ── EVENT HANDLERS: TAB 2 ────────────────────────────────────────────────

    def reload_links_clicked(self, sender, e):
        self._load_link_worksets()

    def ws_link_selection_changed(self, sender, e):
        row = self.grid_ws_links.SelectedItem
        if row is None or row is self._ws_focus:
            return
        self._ws_focus = row
        self._rebuild_workset_rows()

    def ws_links_header_clicked(self, sender, e):
        check_val = bool(self.chk_ws_links_header.IsChecked)
        for row in self._ws_link_rows:
            if row.IsEnabled:
                row.IsSelected = check_val
        self.grid_ws_links.ItemsSource = to_items_source(self._ws_link_rows)
        self._rebuild_workset_rows()

    def ws_link_checkbox_clicked(self, sender, e):
        self._rebuild_workset_rows()

    def ws_filter_changed(self, sender, e):
        self._apply_ws_filter()

    def ws_open_all_clicked(self, sender, e):
        for row in self._ws_filtered:
            row.IsSelected = True
        self._refresh_ws_grid()
        self._apply_ws_filter()

    def ws_close_all_clicked(self, sender, e):
        for row in self._ws_filtered:
            row.IsSelected = False
        self._refresh_ws_grid()
        self._apply_ws_filter()

    def ws_invert_clicked(self, sender, e):
        for row in self._ws_filtered:
            row.IsSelected = not row.IsSelected
        self._refresh_ws_grid()
        self._apply_ws_filter()

    def ws_header_clicked(self, sender, e):
        check_val = bool(self.chk_ws_header.IsChecked)
        for row in self._ws_filtered:
            row.IsSelected = check_val
        self._refresh_ws_grid()
        self._apply_ws_filter()

    def ws_checkbox_clicked(self, sender, e):
        self._refresh_ws_grid()
        self._apply_ws_filter()

    # ── EVENT HANDLERS: TAB 3 ────────────────────────────────────────────────

    def disp_filter_changed(self, sender, e):
        self._apply_disp_filter()

    def disp_link_selection_changed(self, sender, e):
        row = self.grid_disp_links.SelectedItem
        if row is None or row is self._disp_focus:
            return
        self._focus_display_row(row)

    def disp_header_clicked(self, sender, e):
        check_val = bool(self.chk_disp_header.IsChecked)
        for row in self._disp_filtered:
            if row.IsEnabled:
                row.IsSelected = check_val
        self.grid_disp_links.ItemsSource = to_items_source(self._disp_filtered)
        self._update_disp_status()

    def disp_checkbox_clicked(self, sender, e):
        self._update_disp_status()

    def disp_mode_changed(self, sender, e):
        if self._suspend_mode_event:
            return
        self._sync_custom_panel()

    def preset_model_clicked(self, sender, e):
        """Model only — the link follows this view (no link annotations)."""
        self.cb_disp_mode.SelectedIndex = 0
        self.chk_disp_visible.IsChecked = True
        self._sync_custom_panel()
        self._set_status("Preset: By Host View — link draws with this view's model categories.")

    def preset_annotation_clicked(self, sender, e):
        """Model + Annotation — the link follows one of its own views."""
        self.cb_disp_mode.SelectedIndex = 1
        self.chk_disp_visible.IsChecked = True
        if self.cb_disp_view.SelectedIndex < 0 and self._disp_views:
            self.cb_disp_view.SelectedIndex = 0
        self._sync_custom_panel()
        if not self._disp_views:
            self._set_status("This link exposes no matching view — pick another link or view type.")
        else:
            self._set_status("Preset: By Linked View — link draws its own model and annotations.")

    def preset_custom_clicked(self, sender, e):
        self.cb_disp_mode.SelectedIndex = 2
        if self.cb_disp_view.SelectedIndex < 0 and self._disp_views:
            self.cb_disp_view.SelectedIndex = 0
        self._sync_custom_panel()
        self._set_status("Preset: Custom — set each aspect below, then apply.")

    def _pump_events(self):
        self._do_events()


def show_batch_link(doc=None):
    """Entry point to display BatchLinkDialog."""
    dlg = BatchLinkDialog(doc)
    dlg.ShowDialog()
