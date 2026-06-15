# -*- coding: utf-8 -*-
__title__   = "On Sheets:\nSelect"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Select on Sheets (Unified)

Select elements on sheets: choose CAD Imports (DWGs) or Title Blocks.
___________________________________________________________
How-to:
- Select sheets in the Project Browser, OR
- Run with nothing selected and pick sheets from the list
- Choose a Target (CAD Imports / Title Blocks)
- Click "Select on Sheets" to apply the selection
___________________________________________________________
Dang Quoc Truong - DQT (c) 2026
"""

# ----------------------------------------------------------------- LIB PATH
import os as _os, sys as _sys
_here = _os.path.dirname(__file__)
for _up in range(6):
    _libp = _os.path.join(_here, 'lib')
    if _os.path.isdir(_os.path.join(_libp, 'dqt_select')) and _libp not in _sys.path:
        _sys.path.append(_libp)
        break
    _here = _os.path.dirname(_here)
# -----------------------------------------------------------------------------

# ----------------------------------------------------------------- IMPORTS
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (Window, WindowStartupLocation, Thickness,
                            FontWeights, ResizeMode)
from System.Windows.Controls import (StackPanel, RadioButton, Button,
                                     TextBlock, Orientation)
from System.Windows.Media import BrushConverter

from pyrevit import forms

from Autodesk.Revit.DB import (FilteredElementCollector, ViewSheet,
                               ImportInstance, BuiltInCategory)

from dqt_select.compat import to_element_id_list, eid_int, notify

# ----------------------------------------------------------------- VARIABLES
doc    = __revit__.ActiveUIDocument.Document   # noqa: F821
uidoc  = __revit__.ActiveUIDocument            # noqa: F821
# -----------------------------------------------------------------------------


def _brush(hex_color):
    """Convert a hex color string to a SolidColorBrush."""
    bc = BrushConverter()
    return bc.ConvertFromString(hex_color)


# ----------------------------------------------------------------- HELPERS
def get_target_sheets(button_name, alert_title):
    """Return sheets from current selection or ask the user to pick some."""
    sel_ids = uidoc.Selection.GetElementIds()
    sheets = [doc.GetElement(i) for i in sel_ids
              if isinstance(doc.GetElement(i), ViewSheet)]
    if sheets:
        return sheets

    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    if not all_sheets:
        forms.alert('There are no sheets in this model.',
                    title=alert_title, exitscript=True)

    sheet_map = {'{} - {}'.format(s.SheetNumber, s.Name): s
                 for s in all_sheets}
    chosen = forms.SelectFromList.show(
        sorted(sheet_map.keys()),
        title='DQT - Pick Sheets',
        button_name=button_name,
        multiselect=True,
    )
    if not chosen:
        forms.alert('No sheet selected. Cancelled.',
                    title=alert_title, exitscript=True)
    return [sheet_map[c] for c in chosen]


def _select_dwgs(sheets):
    """Select ImportInstance (DWG) elements on the given sheets."""
    sheet_ids = set(eid_int(s.Id) for s in sheets)

    all_imports = (FilteredElementCollector(doc)
                   .OfClass(ImportInstance)
                   .WhereElementIsNotElementType()
                   .ToElements())

    dwg_ids = [imp.Id for imp in all_imports
               if eid_int(imp.OwnerViewId) in sheet_ids]

    if dwg_ids:
        uidoc.Selection.SetElementIds(to_element_id_list(dwg_ids))
        notify('Selected {} DWG(s) on {} sheet(s).'.format(
                   len(dwg_ids), len(sheets)),
               title='DQT - On Sheets: CAD Imports')
    else:
        forms.alert('No DWGs found on the selected sheets.',
                    title='DQT - On Sheets: CAD Imports')


def _select_title_blocks(sheets):
    """Select Title Block elements on the given sheets."""
    sheet_ids = set(eid_int(s.Id) for s in sheets)

    all_tb = (FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_TitleBlocks)
              .WhereElementIsNotElementType()
              .ToElements())

    tb_ids = [tb.Id for tb in all_tb
              if eid_int(tb.OwnerViewId) in sheet_ids]

    if tb_ids:
        uidoc.Selection.SetElementIds(to_element_id_list(tb_ids))
        notify('Selected {} title block(s) on {} sheet(s).'.format(
                   len(tb_ids), len(sheets)),
               title='DQT - On Sheets: Title Blocks')
    else:
        forms.alert('No title blocks found on the selected sheets.',
                    title='DQT - On Sheets: Title Blocks')


# ----------------------------------------------------------------- DIALOG
class SelectOnSheetsDialog(Window):
    def __init__(self):
        self.Title = 'Select on Sheets'
        self.Width = 320
        self.Height = 190
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush('#F8FAFC')
        self.FontFamily = System.Windows.Media.FontFamily('Inter')

        root = StackPanel()
        root.Margin = Thickness(16, 14, 16, 14)

        # ---- Target section -------------------------------------------------
        target_label = TextBlock()
        target_label.Text = 'Target'
        target_label.FontSize = 12
        target_label.FontWeight = FontWeights.SemiBold
        target_label.Foreground = _brush('#0F172A')
        target_label.Margin = Thickness(0, 0, 0, 6)
        root.Children.Add(target_label)

        target_panel = StackPanel()
        target_panel.Orientation = Orientation.Horizontal
        target_panel.Margin = Thickness(0, 0, 0, 20)

        self._rb_dwg = RadioButton()
        self._rb_dwg.Content = 'CAD Imports'
        self._rb_dwg.GroupName = 'target_group'
        self._rb_dwg.IsChecked = True
        self._rb_dwg.FontSize = 12
        self._rb_dwg.Foreground = _brush('#0F172A')
        self._rb_dwg.Margin = Thickness(0, 0, 16, 0)
        target_panel.Children.Add(self._rb_dwg)

        self._rb_tb = RadioButton()
        self._rb_tb.Content = 'Title Blocks'
        self._rb_tb.GroupName = 'target_group'
        self._rb_tb.IsChecked = False
        self._rb_tb.FontSize = 12
        self._rb_tb.Foreground = _brush('#0F172A')
        target_panel.Children.Add(self._rb_tb)

        root.Children.Add(target_panel)

        # ---- Select button --------------------------------------------------
        btn_select = Button()
        btn_select.Content = 'Select on Sheets'
        btn_select.Height = 34
        btn_select.FontSize = 12
        btn_select.FontWeight = FontWeights.SemiBold
        btn_select.Background = _brush('#0F172A')
        btn_select.Foreground = _brush('#FFFFFF')
        btn_select.BorderThickness = Thickness(0)
        btn_select.Cursor = System.Windows.Input.Cursors.Hand
        btn_select.Click += self._on_select
        root.Children.Add(btn_select)

        self.Content = root

    def _on_select(self, sender, e):
        use_dwg = bool(self._rb_dwg.IsChecked)
        self.Close()

        if use_dwg:
            sheets = get_target_sheets(
                button_name='Select DWGs',
                alert_title='DQT - On Sheets: CAD Imports',
            )
            _select_dwgs(sheets)
        else:
            sheets = get_target_sheets(
                button_name='Select Title Blocks',
                alert_title='DQT - On Sheets: Title Blocks',
            )
            _select_title_blocks(sheets)


# ----------------------------------------------------------------- ENTRY
if __name__ == '__main__':
    dlg = SelectOnSheetsDialog()
    dlg.ShowDialog()
