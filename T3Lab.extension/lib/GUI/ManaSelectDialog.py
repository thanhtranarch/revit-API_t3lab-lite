# -*- coding: utf-8 -*-
"""ManaSelect — unified controller for smart selection tools."""

import os
import sys
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import System
import System.Windows
from System.Windows import WindowState, Visibility
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSheet,
    ImportInstance,
    BuiltInCategory
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit, DB, script

# Add current folder to sys.path to find GUI dialog classes
sys.path.append(os.path.dirname(__file__))
# Add parent of Selection folder to find Selection
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

import QuickElementDialog

# Import dqt selection logic
from Selection.dqt_select import core as dqt_core
from Selection.dqt_select import compat as dqt_compat

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ManaSelect.xaml')


class ManaSelectWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)
        self.uidoc = revit.uidoc
        self.doc = revit.doc

        # Initialize and nest sub-panels
        self._init_sub_panels()

        # Connect navigation (left icon-rail sidebar toggles)
        self.nav_toggle_quick_select.Click += self._on_nav_toggle_clicked
        self.nav_toggle_select_similar.Click += self._on_nav_toggle_clicked
        self.nav_toggle_select_sheets.Click += self._on_nav_toggle_clicked

        # Connect run actions for tabs
        self.btn_pick_similar_seed.Click += self._on_pick_similar_seed
        self.btn_run_select_similar.Click += self._on_run_select_similar
        self.btn_run_select_sheets.Click += self._on_run_select_sheets

        # Chrome actions
        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close_chrome.Click += self._close_chrome

    def _init_sub_panels(self):
        """Loads Quick Select window grid and collapses its header to avoid duplicates."""
        try:
            self._quick_select_win = QuickElementDialog.QuickSelectWindow()
            quick_select_border = self._quick_select_win.Content
            # Content is Border, its Child is Grid. The Border is QuickElement's own
            # standalone window chrome (rounded corners + gray edge) - reparenting it
            # as-is would nest that chrome frame inside ManaSelect's own window chrome,
            # showing up as a gray gutter/border around the tab. Reparent the inner
            # Grid only and drop the outer chrome Border.
            quick_select_grid = quick_select_border.Child
            self._quick_select_win.Content = None
            quick_select_border.Child = None
            self.grid_quick_select.Children.Add(quick_select_grid)

            # Collapse sub-tool header and footer to unify status bar
            quick_select_grid.RowDefinitions[0].Height = System.Windows.GridLength(0)
            quick_select_grid.RowDefinitions[4].Height = System.Windows.GridLength(0)

            # Re-wire close
            self._quick_select_win.Close = self.Close
        except Exception as ex:
            print("Error loading Quick Select panel: {}".format(ex))

    def _on_nav_toggle_clicked(self, sender, e):
        """Switch active TabControl index, sync rail toggle state and update status bar text."""
        if sender == self.nav_toggle_quick_select:
            index = 0
            self.status_text.Text = "Quick Select — Query elements by categories, parameters and text filters"
        elif sender == self.nav_toggle_select_similar:
            index = 1
            self.status_text.Text = "Select Similar — Match Type, Family or Category of current selection"
        elif sender == self.nav_toggle_select_sheets:
            index = 2
            self.status_text.Text = "Select on Sheets — Find CAD imports or title blocks across drawings"
        else:
            return

        self.main_tab_control.SelectedIndex = index
        self.nav_toggle_quick_select.IsChecked = (index == 0)
        self.nav_toggle_select_similar.IsChecked = (index == 1)
        self.nav_toggle_select_sheets.IsChecked = (index == 2)

    # =========================================================================
    # TAB 2: SELECT SIMILAR
    # =========================================================================
    def _on_pick_similar_seed(self, sender, e):
        """Let the user pick a single seed element in the model, then apply it
        as the current Revit selection so it flows through to Apply Selection."""
        self.Hide()
        try:
            ref = self.uidoc.Selection.PickObject(ObjectType.Element, "Pick a seed element for Select Similar")
            elem = self.doc.GetElement(ref.ElementId)
            if elem is not None:
                self.uidoc.Selection.SetElementIds(dqt_compat.to_element_id_list([elem.Id]))
                cat_name = elem.Category.Name if elem.Category is not None else elem.__class__.__name__
                self.txt_similar_seed_status.Text = "Picked: {} (Id {})".format(
                    cat_name, dqt_compat.eid_int(elem.Id))
        except OperationCanceledException:
            pass
        finally:
            self.Show()

    def _on_run_select_similar(self, sender, e):
        """Execute select similar based on UI configurations."""
        if not self.uidoc.Selection.GetElementIds():
            forms.alert("Please select at least one seed element in the model first.", title="Select Similar")
            return

        scope = 'view' if self.rb_similar_scope_view.IsChecked else 'model'
        try:
            if self.rb_similar_mode_type.IsChecked:
                dqt_core.select_similar_type(mode=scope)
            elif self.rb_similar_mode_family.IsChecked:
                dqt_core.select_similar_family(mode=scope)
            else:
                dqt_core.select_similar_category(mode=scope)
        except Exception as ex:
            forms.alert("Error running Select Similar: {}".format(ex), title="Error")

    # =========================================================================
    # TAB 3: SELECT ON SHEETS
    # =========================================================================
    def _on_run_select_sheets(self, sender, e):
        """Execute selection on sheets based on target selection."""
        use_dwg = bool(self.rb_sheet_target_dwg.IsChecked)
        if use_dwg:
            sheets = self._get_target_sheets('Select DWGs', 'DQT - On Sheets: CAD Imports')
            if sheets:
                self._select_dwgs(sheets)
        else:
            sheets = self._get_target_sheets('Select Title Blocks', 'DQT - On Sheets: Title Blocks')
            if sheets:
                self._select_title_blocks(sheets)

    def _get_target_sheets(self, button_name, alert_title):
        sel_ids = self.uidoc.Selection.GetElementIds()
        sheets = [self.doc.GetElement(i) for i in sel_ids
                  if isinstance(self.doc.GetElement(i), ViewSheet)]
        if sheets:
            return sheets

        all_sheets = FilteredElementCollector(self.doc).OfClass(ViewSheet).ToElements()
        if not all_sheets:
            forms.alert('There are no sheets in this model.', title=alert_title)
            return None

        sheet_map = {'{} - {}'.format(s.SheetNumber, s.Name): s for s in all_sheets}
        chosen = forms.SelectFromList.show(
            sorted(sheet_map.keys()),
            title='DQT - Pick Sheets',
            button_name=button_name,
            multiselect=True,
        )
        if not chosen:
            return None
        return [sheet_map[c] for c in chosen]

    def _select_dwgs(self, sheets):
        sheet_ids = set(dqt_compat.eid_int(s.Id) for s in sheets)
        all_imports = (FilteredElementCollector(self.doc)
                       .OfClass(ImportInstance)
                       .WhereElementIsNotElementType()
                       .ToElements())

        dwg_ids = [imp.Id for imp in all_imports
                   if dqt_compat.eid_int(imp.OwnerViewId) in sheet_ids]

        if dwg_ids:
            self.uidoc.Selection.SetElementIds(dqt_compat.to_element_id_list(dwg_ids))
            dqt_compat.notify('Selected {} DWG(s) on {} sheet(s).'.format(
                       len(dwg_ids), len(sheets)),
                    title='DQT - On Sheets: CAD Imports')
        else:
            forms.alert('No DWGs found on the selected sheets.',
                        title='DQT - On Sheets: CAD Imports')

    def _select_title_blocks(self, sheets):
        sheet_ids = set(dqt_compat.eid_int(s.Id) for s in sheets)
        all_tb = (FilteredElementCollector(self.doc)
                  .OfCategory(BuiltInCategory.OST_TitleBlocks)
                  .WhereElementIsNotElementType()
                  .ToElements())

        tb_ids = [tb.Id for tb in all_tb
                  if dqt_compat.eid_int(tb.OwnerViewId) in sheet_ids]

        if tb_ids:
            self.uidoc.Selection.SetElementIds(dqt_compat.to_element_id_list(tb_ids))
            dqt_compat.notify('Selected {} title block(s) on {} sheet(s).'.format(
                       len(tb_ids), len(sheets)),
                    title='DQT - On Sheets: Title Blocks')
        else:
            forms.alert('No title blocks found on the selected sheets.',
                        title='DQT - On Sheets: Title Blocks')

    # =========================================================================
    # WINDOW CHROME
    # =========================================================================
    def _minimize(self, sender, e):
        self.WindowState = WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()


def show_dialog():
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)
    window = ManaSelectWindow()
    window.ShowDialog()
