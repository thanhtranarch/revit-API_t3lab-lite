# -*- coding: utf-8 -*-
"""
ManaGroupDialog.py
==================
WPF Dialog for managing Revit groups — model groups, detail groups and attached
detail groups — in three tabs:

* **Rename**  — batch rename group types with find/replace, prefix, suffix,
                letter case and an illegal-character cleanup, with a live preview
                of every new name before anything is written.
* **Workset** — move the instances of the selected group types onto one workset,
                optionally taking the elements inside each group with them.
* **Cleanup** — audit every group type (unused, single instance, mixed worksets,
                name problems) then purge unused types or ungroup instances.

Revit API work lives in ``Snippets/_group_ops.py``; this module is UI only.

Part of T3Lab Extension.
Author: T3Lab
"""

import os

import clr
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import Visibility
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri, UriKind

from pyrevit import revit

from GUI.WPF_Base import T3WPFWindow, to_items_source
from GUI.T3Dialog import confirm as t3_confirm, show_warning as t3_warning

from Snippets import _group_ops

GUI_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(GUI_DIR, 'Tools', 'ManaGroup.xaml')

TAB_RENAME = 0
TAB_WORKSET = 1
TAB_CLEANUP = 2

PRIMARY_LABELS = {
    TAB_RENAME: "Apply Rename",
    TAB_WORKSET: "Apply Workset",
    TAB_CLEANUP: "Rescan Model",
}

KIND_ALL = "All groups"
KIND_FILTERS = (
    (KIND_ALL, None),
    ("Model groups", _group_ops.KIND_MODEL),
    ("Detail groups", _group_ops.KIND_DETAIL),
    ("Attached detail groups", _group_ops.KIND_ATTACHED),
)


# ── ROW ITEMS ────────────────────────────────────────────────────────────────

class GroupRow(object):
    """One group type as shown in a grid. Three rows share one record."""

    def __init__(self, record, status_text="Ready", severity="Success",
                 is_selected=False, is_enabled=True):
        self.record = record
        self._new_name = record.name
        self._manual = False
        self._status_text = status_text
        self._severity = severity
        self._is_selected = bool(is_selected)
        self._is_enabled = bool(is_enabled)

    # -- read-only columns ---------------------------------------------------

    @property
    def GroupName(self):
        return self.record.name

    @property
    def Kind(self):
        return self.record.kind

    @property
    def InstanceCount(self):
        return str(self.record.instance_count)

    @property
    def MemberCount(self):
        return str(self.record.member_count) if self.record.member_count else "—"

    @property
    def WorksetSummary(self):
        return self.record.workset_summary

    @property
    def Issues(self):
        found = self.record.audit_issues()
        return ", ".join(found) if found else "None"

    # -- editable / stateful -------------------------------------------------

    @property
    def NewName(self):
        return self._new_name

    @NewName.setter
    def NewName(self, value):
        self._new_name = value or ""

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

    def set_status(self, text, severity="Success"):
        self._status_text = text
        self._severity = severity


class ManaGroupDialog(T3WPFWindow):
    """Main window class for Group Manager."""

    def __init__(self, doc=None):
        T3WPFWindow.__init__(self, XAML_FILE)
        self.doc = doc or getattr(revit, 'doc', None)
        self._loading = True
        self._is_busy = False
        self._active_tab = TAB_RENAME

        self._records = []
        self._ren_rows = []
        self._ws_rows = []
        self._cln_rows = []
        self._worksets = []          # [(id_int, name)]

        self._load_logo()
        self._init_filters()
        self._reload_model()
        self._loading = False

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

    def _init_filters(self):
        """Fill the kind filter and the letter-case dropdown."""
        self.cb_kind.ItemsSource = to_items_source([label for label, _ in KIND_FILTERS])
        self.cb_kind.SelectedIndex = 0

        self.cb_case.ItemsSource = to_items_source(list(_group_ops.CASE_MODES))
        self.cb_case.SelectedIndex = 0

    # ── MODEL LOADING ────────────────────────────────────────────────────────

    def _reload_model(self):
        """Re-read every group type in the document and rebuild all three tabs."""
        self._records = _group_ops.collect_group_types(self.doc)

        # The Rename tab opens with everything checked so the preview column is
        # meaningful straight away; the other two tabs act on the model, so they
        # start empty and the user opts in.
        self._ren_rows = [GroupRow(r, is_selected=True) for r in self._records]
        self._ws_rows = [GroupRow(r, is_enabled=bool(r.instance_count))
                         for r in self._records]
        self._cln_rows = [GroupRow(r) for r in self._records]

        for row in self._ws_rows:
            if not row.record.instance_count:
                row.set_status("Not placed", "Warning")
        for row in self._cln_rows:
            found = row.record.audit_issues()
            if not found:
                row.set_status("Clean", "Success")
            elif "Unused" in found:
                row.set_status("Unused", "Warning")
            else:
                row.set_status("%d finding%s" % (len(found), "" if len(found) == 1 else "s"),
                               "Warning")

        self._load_worksets()
        self._recompute_names()
        self._apply_filter()
        for checkbox, rows in ((self.chk_ren_header, self._ren_rows),
                               (self.chk_ws_header, self._ws_rows),
                               (self.chk_cln_header, self._cln_rows)):
            self._sync_header_checkbox(checkbox, self._visible(rows))
        self._set_status("%d group type%s in this model — %d model, %d detail, %d attached."
                         % (len(self._records),
                            "" if len(self._records) == 1 else "s",
                            self._count_kind(_group_ops.KIND_MODEL),
                            self._count_kind(_group_ops.KIND_DETAIL),
                            self._count_kind(_group_ops.KIND_ATTACHED)))

    def _count_kind(self, kind):
        return sum(1 for r in self._records if r.kind == kind)

    def _load_worksets(self):
        """Fill the target workset dropdown, or explain why it is unavailable."""
        self._worksets = _group_ops.user_worksets(self.doc)
        usable = bool(self._worksets)

        self.cb_workset.ItemsSource = to_items_source([name for _, name in self._worksets])
        if usable:
            self.cb_workset.SelectedIndex = 0
        self.cb_workset.IsEnabled = usable
        self.chk_members.IsEnabled = usable
        self.banner_no_worksharing.Visibility = \
            Visibility.Collapsed if usable else Visibility.Visible

        if not usable:
            workshared = False
            try:
                workshared = bool(self.doc and self.doc.IsWorkshared)
            except Exception:
                pass
            self.lbl_no_worksharing.Text = (
                "No user workset was found in this model, so groups cannot be moved. "
                "Create a workset in Revit first — renaming and cleanup still work."
                if workshared else
                "This model is not workshared, so groups cannot be moved between worksets. "
                "Enable worksharing in Revit first — renaming and cleanup still work.")

    # ── FILTERING ────────────────────────────────────────────────────────────

    def _search_text(self):
        try:
            return (self.tb_search.Text or "").strip().lower()
        except Exception:
            return ""

    def _kind_filter(self):
        index = self.cb_kind.SelectedIndex
        if index is None or index < 0 or index >= len(KIND_FILTERS):
            return None
        return KIND_FILTERS[index][1]

    def _visible(self, rows):
        needle = self._search_text()
        kind = self._kind_filter()
        out = []
        for row in rows:
            if kind and row.record.kind != kind:
                continue
            if needle and needle not in row.record.name.lower():
                continue
            out.append(row)
        return out

    def _tabs(self):
        """(rows, grid, empty state, count label) for each of the three tabs."""
        return (
            (self._ren_rows, self.grid_rename, self.txt_ren_empty, self.lbl_ren_count),
            (self._ws_rows, self.grid_workset, self.txt_ws_empty, self.lbl_ws_count),
            (self._cln_rows, self.grid_cleanup, self.txt_cln_empty, self.lbl_cln_count))

    def _apply_filter(self):
        """Push the filtered rows into all three grids and update the counters.

        Reassigns ItemsSource, so it only runs on load or when the filter really
        changes — doing it on every checkbox click would reset the scroll
        position and drop the focus out of the NEW NAME box.
        """
        for rows, grid, empty, _label in self._tabs():
            shown = self._visible(rows)
            grid.ItemsSource = to_items_source(shown)
            empty.Visibility = Visibility.Collapsed if shown else Visibility.Visible
        self._update_counts()

    def _update_counts(self):
        """Refresh the "N group types · M checked" caption above each grid."""
        for rows, _grid, _empty, label in self._tabs():
            shown = self._visible(rows)
            checked = sum(1 for r in shown if r.IsSelected)
            label.Text = "%d group type%s · %d checked" % (
                len(shown), "" if len(shown) == 1 else "s", checked)

    def _refresh_grids(self):
        """Redraw every grid — the rows carry no INotifyPropertyChanged."""
        self.grid_rename.Items.Refresh()
        self.grid_workset.Items.Refresh()
        self.grid_cleanup.Items.Refresh()

    def _refresh_current_grid(self):
        """Redraw only the grid the user is looking at."""
        if self._active_tab == TAB_WORKSET:
            self.grid_workset.Items.Refresh()
        elif self._active_tab == TAB_CLEANUP:
            self.grid_cleanup.Items.Refresh()
        else:
            self.grid_rename.Items.Refresh()

    def _current_rows(self):
        """The row list belonging to the active tab."""
        if self._active_tab == TAB_WORKSET:
            return self._ws_rows
        if self._active_tab == TAB_CLEANUP:
            return self._cln_rows
        return self._ren_rows

    def _checked(self, rows):
        """Checked rows that are visible under the current filter and enabled."""
        return [r for r in self._visible(rows) if r.IsSelected and r.IsEnabled]

    # ── SHARED HELPERS ───────────────────────────────────────────────────────

    def _set_status(self, text):
        self.status_text.Text = text

    def _begin_busy(self, phase):
        self._is_busy = True
        self.lbl_phase.Text = phase
        self.lbl_current_item.Text = ""
        self.progress_bar.Value = 0
        self.pnl_progress.Visibility = Visibility.Visible
        self.btn_primary.IsEnabled = False
        self.btn_cancel.IsEnabled = False
        self._do_events()

    def _step_busy(self, phase, index, total, label):
        self.lbl_phase.Text = "%s — %d / %d" % (phase, index, total)
        self.lbl_current_item.Text = label or ""
        self.progress_bar.Value = int((float(index - 1) / max(1, total)) * 100)
        self._do_events()

    def _end_busy(self, summary):
        self._is_busy = False
        self.progress_bar.Value = 100
        self.btn_primary.IsEnabled = True
        self.btn_cancel.IsEnabled = True
        self._set_status(summary)
        self._do_events()

    @staticmethod
    def _sync_header_checkbox(checkbox, rows):
        """Tri-state the header checkbox from the rows below it."""
        if checkbox is None:
            return
        selectable = [r for r in rows if r.IsEnabled]
        checked = sum(1 for r in selectable if r.IsSelected)
        if not selectable or checked == 0:
            checkbox.IsChecked = False
        elif checked == len(selectable):
            checkbox.IsChecked = True
        else:
            checkbox.IsChecked = None

    def _header_for(self, tab=None):
        tab = self._active_tab if tab is None else tab
        if tab == TAB_WORKSET:
            return self.chk_ws_header
        if tab == TAB_CLEANUP:
            return self.chk_cln_header
        return self.chk_ren_header

    def _after_selection_change(self):
        self._sync_header_checkbox(self._header_for(), self._visible(self._current_rows()))
        if self._active_tab == TAB_RENAME:
            self._recompute_names()
        self._update_counts()
        self._refresh_current_grid()

    # ── TAB 1: RENAME ────────────────────────────────────────────────────────

    def _rule_values(self):
        return dict(
            find=(self.tb_find.Text or ""),
            replace=(self.tb_replace.Text or ""),
            match_case=bool(self.chk_match_case.IsChecked),
            prefix=(self.tb_prefix.Text or ""),
            suffix=(self.tb_suffix.Text or ""),
            case_mode=self._case_mode(),
            cleanup=bool(self.chk_cleanup.IsChecked))

    def _case_mode(self):
        index = self.cb_case.SelectedIndex
        if index is None or index < 0 or index >= len(_group_ops.CASE_MODES):
            return _group_ops.CASE_KEEP
        return _group_ops.CASE_MODES[index]

    def _recompute_names(self):
        """Re-apply the rename rules and re-validate every proposed name."""
        rules = self._rule_values()

        for row in self._ren_rows:
            if not row._manual:
                row.NewName = _group_ops.build_new_name(row.record.name, **rules)

        # Every name that will exist after the rename, to catch collisions.
        planned = {}
        for row in self._ren_rows:
            final = (row.NewName if row.IsSelected else row.record.name) or ""
            key = final.strip().lower()
            planned[key] = planned.get(key, 0) + 1

        for row in self._ren_rows:
            wanted = (row.NewName or "").strip()
            if not row.IsSelected:
                row.set_status("Not selected", "Warning")
                continue
            if not wanted:
                row.set_status("Empty name", "Danger")
                continue
            bad = _group_ops.illegal_chars_in(wanted)
            if bad:
                row.set_status("Illegal %s" % " ".join(bad), "Danger")
                continue
            if wanted == row.record.name:
                row.set_status("Unchanged", "Success")
                continue
            if planned.get(wanted.lower(), 0) > 1:
                row.set_status("Duplicate name", "Danger")
                continue
            row.set_status("Will rename", "Success")

    def _apply_rename(self):
        """Rename every checked group type whose proposed name is valid."""
        rows = self._checked(self._ren_rows)
        if not rows:
            self._set_status("Check at least one group type on the Rename tab first.")
            return

        blocked = [r for r in rows if r.Severity == "Danger"]
        pairs = [(r.record, (r.NewName or "").strip())
                 for r in rows
                 if r.Severity != "Danger" and (r.NewName or "").strip() != r.record.name]
        if not pairs:
            if blocked:
                self._set_status(
                    "Nothing to rename — %d checked name%s still has a problem to fix "
                    "(see the STATUS column)." % (
                        len(blocked), "" if len(blocked) == 1 else "s"))
            else:
                self._set_status(
                    "Nothing to rename — the %d checked group type%s already carry these names."
                    % (len(rows), "" if len(rows) == 1 else "s"))
            return

        if not t3_confirm(
                "Rename %d group type%s in this model?" % (
                    len(pairs), "" if len(pairs) == 1 else "s"),
                title="Rename groups",
                ok_text="Rename",
                details="Every placed instance keeps its geometry — only the type name changes.",
                owner=self):
            return

        by_record = {}
        for row in self._ren_rows:
            by_record[row.record.type_id] = row

        self._begin_busy("Renaming groups")
        try:
            results = _group_ops.rename_group_types(
                self.doc, pairs,
                progress=lambda i, t, label: self._step_busy("Renaming groups", i, t, label))
        except Exception as exc:
            self._end_busy("Rename failed — nothing was changed.")
            t3_warning("The rename was rolled back, so the model is unchanged.",
                       title="Rename failed", details=str(exc), owner=self)
            return

        renamed = failed = 0
        for record, ok, message in results:
            row = by_record.get(record.type_id)
            if ok:
                renamed += 1
                if row is not None:
                    row._manual = False
                    row.NewName = record.name
                    row.set_status("Renamed", "Success")
            else:
                failed += 1
                if row is not None:
                    row.set_status(message, "Danger")

        self._apply_filter()
        self._refresh_grids()
        self._end_busy("Renamed %d group type%s, %d failed." % (
            renamed, "" if renamed == 1 else "s", failed))

    # ── TAB 2: WORKSET ───────────────────────────────────────────────────────

    def _apply_workset(self):
        """Move the instances of every checked group type onto the target workset."""
        if not self._worksets:
            self._set_status("This model is not workshared — there is no workset to move to.")
            return

        rows = self._checked(self._ws_rows)
        if not rows:
            self._set_status("Check at least one placed group type on the Workset tab first.")
            return

        index = self.cb_workset.SelectedIndex
        if index is None or index < 0 or index >= len(self._worksets):
            self._set_status("Pick the workset the groups should move to.")
            return
        workset_value, workset_name = self._worksets[index]

        include_members = bool(self.chk_members.IsChecked)
        instances = sum(r.record.instance_count for r in rows)
        members = sum(r.record.instance_count * r.record.member_count
                      for r in rows) if include_members else 0

        detail = "%d group instance%s%s will move to \"%s\"." % (
            instances, "" if instances == 1 else "s",
            " and about %d member element%s" % (members, "" if members == 1 else "s")
            if include_members else "",
            workset_name)
        if not t3_confirm(detail,
                          title="Move groups to another workset",
                          ok_text="Move them",
                          owner=self):
            return

        by_record = {row.record.type_id: row for row in self._ws_rows}

        self._begin_busy("Setting worksets")
        try:
            results = _group_ops.apply_workset(
                self.doc, [r.record for r in rows], workset_value,
                include_members=include_members,
                progress=lambda i, t, label: self._step_busy("Setting worksets", i, t, label))
        except Exception as exc:
            self._end_busy("Workset change failed — nothing was changed.")
            t3_warning("The workset change was rolled back, so the model is unchanged.",
                       title="Workset change failed", details=str(exc), owner=self)
            return

        moved_total = failed_total = 0
        for record, moved, _skipped, failed, message in results:
            moved_total += moved
            failed_total += failed
            row = by_record.get(record.type_id)
            if row is not None:
                row.set_status(message, "Danger" if failed else "Success")

        self._refresh_worksets()
        self._end_busy("Moved %d element%s to \"%s\", %d failed." % (
            moved_total, "" if moved_total == 1 else "s", workset_name, failed_total))

    def _refresh_worksets(self):
        """Re-read the workset of every instance so the grid shows the new state."""
        for record in self._records:
            seen = []
            for instance in record.instances:
                name = _group_ops.instance_workset_name(self.doc, instance)
                if name and name not in seen:
                    seen.append(name)
            record.workset_names = seen
        self._apply_filter()
        self._refresh_grids()

    # ── TAB 3: CLEANUP ───────────────────────────────────────────────────────

    def _purge_selected(self):
        """Delete the checked group types that are no longer placed."""
        rows = self._checked(self._cln_rows)
        if not rows:
            self._set_status("Check at least one group type on the Cleanup tab first.")
            return

        unused = [r for r in rows if r.record.instance_count == 0]
        placed = len(rows) - len(unused)
        if not unused:
            self._set_status(
                "Nothing to purge — all %d checked group type%s are still placed in the model."
                % (len(rows), "" if len(rows) == 1 else "s"))
            return

        detail = "%d unused group type%s will be deleted from the project browser." % (
            len(unused), "" if len(unused) == 1 else "s")
        if placed:
            detail += " %d still-placed type%s stay untouched." % (
                placed, "" if placed == 1 else "s")
        if not t3_confirm(detail,
                          title="Purge unused group types",
                          ok_text="Purge",
                          danger=True,
                          owner=self):
            return

        self._begin_busy("Purging group types")
        try:
            results = _group_ops.purge_group_types(
                self.doc, [r.record for r in unused],
                progress=lambda i, t, label: self._step_busy("Purging group types", i, t, label))
        except Exception as exc:
            self._end_busy("Purge failed — nothing was deleted.")
            t3_warning("The purge was rolled back, so the model is unchanged.",
                       title="Purge failed", details=str(exc), owner=self)
            return

        purged = sum(1 for _r, ok, _m in results if ok)
        failed = len(results) - purged
        self._reload_model()
        self._end_busy("Purged %d group type%s, %d failed." % (
            purged, "" if purged == 1 else "s", failed))

    def _ungroup_selected(self):
        """Explode every instance of the checked group types."""
        rows = self._checked(self._cln_rows)
        if not rows:
            self._set_status("Check at least one group type on the Cleanup tab first.")
            return

        placed = [r for r in rows if r.record.instance_count]
        if not placed:
            self._set_status("Nothing to ungroup — none of the checked group types is placed.")
            return

        instances = sum(r.record.instance_count for r in placed)
        if not t3_confirm(
                "Ungroup %d instance%s of %d group type%s?" % (
                    instances, "" if instances == 1 else "s",
                    len(placed), "" if len(placed) == 1 else "s"),
                title="Ungroup instances",
                ok_text="Ungroup",
                danger=True,
                details="Members stay in the model as loose elements. "
                        "Ctrl+Z undoes the whole ungroup in one step.",
                owner=self):
            return

        self._begin_busy("Ungrouping")
        try:
            results = _group_ops.ungroup_instances(
                self.doc, [r.record for r in placed],
                progress=lambda i, t, label: self._step_busy("Ungrouping", i, t, label))
        except Exception as exc:
            self._end_busy("Ungroup failed — nothing was changed.")
            t3_warning("The ungroup was rolled back, so the model is unchanged.",
                       title="Ungroup failed", details=str(exc), owner=self)
            return

        ungrouped = sum(count for _r, count, _f, _m in results)
        failed = sum(f for _r, _c, f, _m in results)
        self._reload_model()
        self._end_busy("Ungrouped %d instance%s, %d failed." % (
            ungrouped, "" if ungrouped == 1 else "s", failed))

    # ── EVENT HANDLERS: WINDOW & TABS ────────────────────────────────────────

    def cancel_button_clicked(self, sender, e):
        if self._is_busy:
            return
        self.Close()

    def tab_chip_checked(self, sender, e):
        if getattr(self, '_loading', True):
            return
        try:
            self._active_tab = int(sender.Tag)
        except Exception:
            self._active_tab = TAB_RENAME
        self.tab_control.SelectedIndex = self._active_tab
        self.btn_primary.Content = PRIMARY_LABELS.get(self._active_tab, "Apply")
        self._sync_header_checkbox(self._header_for(), self._visible(self._current_rows()))

    def primary_button_clicked(self, sender, e):
        if self._is_busy:
            return
        if self._active_tab == TAB_RENAME:
            self._apply_rename()
        elif self._active_tab == TAB_WORKSET:
            self._apply_workset()
        else:
            self._reload_model()

    def refresh_clicked(self, sender, e):
        if self._is_busy:
            return
        self._reload_model()

    # ── EVENT HANDLERS: FILTERS ──────────────────────────────────────────────

    def search_text_changed(self, sender, e):
        if getattr(self, '_loading', True):
            return
        self._apply_filter()

    def kind_filter_changed(self, sender, e):
        if getattr(self, '_loading', True):
            return
        self._apply_filter()

    # ── EVENT HANDLERS: SELECTION ────────────────────────────────────────────

    def select_all_clicked(self, sender, e):
        for row in self._visible(self._current_rows()):
            if row.IsEnabled:
                row.IsSelected = True
        self._after_selection_change()

    def select_none_clicked(self, sender, e):
        for row in self._visible(self._current_rows()):
            row.IsSelected = False
        self._after_selection_change()

    def invert_selection_clicked(self, sender, e):
        for row in self._visible(self._current_rows()):
            if row.IsEnabled:
                row.IsSelected = not row.IsSelected
        self._after_selection_change()

    def select_unused_clicked(self, sender, e):
        for row in self._visible(self._cln_rows):
            row.IsSelected = row.record.instance_count == 0
        self._after_selection_change()

    def select_issues_clicked(self, sender, e):
        for row in self._visible(self._cln_rows):
            row.IsSelected = bool(row.record.audit_issues())
        self._after_selection_change()

    def rename_header_clicked(self, sender, e):
        self._header_clicked(sender, self._ren_rows)

    def workset_header_clicked(self, sender, e):
        self._header_clicked(sender, self._ws_rows)

    def cleanup_header_clicked(self, sender, e):
        self._header_clicked(sender, self._cln_rows)

    def _header_clicked(self, sender, rows):
        wanted = bool(sender.IsChecked)
        for row in self._visible(rows):
            if row.IsEnabled:
                row.IsSelected = wanted
        self._after_selection_change()

    def rename_checkbox_clicked(self, sender, e):
        self._after_selection_change()

    def workset_checkbox_clicked(self, sender, e):
        self._after_selection_change()

    def cleanup_checkbox_clicked(self, sender, e):
        self._after_selection_change()

    # ── EVENT HANDLERS: RENAME RULES ─────────────────────────────────────────

    def rule_changed(self, sender, e):
        if getattr(self, '_loading', True):
            return
        self._recompute_names()
        self.grid_rename.Items.Refresh()

    def rule_toggled(self, sender, e):
        self.rule_changed(sender, e)

    def new_name_changed(self, sender, e):
        """A name typed straight into the grid wins over the rules.

        TextChanged also fires while the grid realises its rows and the binding
        fills each box, so only a keyboard-focused box counts as a manual edit —
        otherwise scrolling the grid would pin every name against the rules.
        """
        if getattr(self, '_loading', True):
            return
        try:
            if not sender.IsKeyboardFocusWithin:
                return
        except Exception:
            return
        row = getattr(sender, 'DataContext', None)
        if isinstance(row, GroupRow):
            row._manual = True
            row.NewName = sender.Text
        self._recompute_names()

    def reset_names_clicked(self, sender, e):
        for row in self._ren_rows:
            row._manual = False
        self._recompute_names()
        self.grid_rename.Items.Refresh()
        self._set_status("New names reset to the current rename rules.")

    # ── EVENT HANDLERS: WORKSET & CLEANUP ────────────────────────────────────

    def members_toggled(self, sender, e):
        if getattr(self, '_loading', True):
            return
        if self.chk_members.IsChecked:
            self.lbl_ws_hint.Text = ("Group instances and every element inside them move to "
                                     "the chosen workset.")
        else:
            self.lbl_ws_hint.Text = ("Group instances move to the chosen workset. Members keep "
                                     "their own workset unless the option above is ticked.")

    def ungroup_clicked(self, sender, e):
        if self._is_busy:
            return
        self._ungroup_selected()

    def purge_clicked(self, sender, e):
        if self._is_busy:
            return
        self._purge_selected()


def show_group_manager(doc=None):
    """Entry point used by the pushbutton."""
    ManaGroupDialog(doc).ShowDialog()
