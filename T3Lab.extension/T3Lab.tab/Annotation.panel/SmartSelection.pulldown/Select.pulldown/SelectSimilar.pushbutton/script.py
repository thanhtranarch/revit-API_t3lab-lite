# -*- coding: utf-8 -*-
__title__   = "Select\nSimilar"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Select Similar (Unified)

Select elements similar to the current selection.
Choose match mode (Category / Family / Type) and scope (Active View / Entire Model).
___________________________________________________________
How-to:
- Select one or more instances in the model or active view
- Run the tool
- Choose a Mode (Type / Family / Category) and Scope (Active View / Entire Model)
- Click "Select" to apply the selection
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
                                     TextBlock, Orientation, Grid,
                                     ColumnDefinition, RowDefinition,
                                     Border)
from System.Windows.Media import BrushConverter

from dqt_select.core import (select_similar_type, select_similar_family,
                              select_similar_category)
# -----------------------------------------------------------------------------


def _brush(hex_color):
    """Convert a hex color string to a SolidColorBrush."""
    bc = BrushConverter()
    return bc.ConvertFromString(hex_color)


# ----------------------------------------------------------------- DIALOG
class SelectSimilarDialog(Window):
    def __init__(self):
        self.Title = 'Select Similar'
        self.Width = 320
        self.Height = 230
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush('#F8FAFC')
        self.FontFamily = System.Windows.Media.FontFamily('Inter')

        root = StackPanel()
        root.Margin = Thickness(16, 14, 16, 14)

        # ---- Mode section ---------------------------------------------------
        mode_label = TextBlock()
        mode_label.Text = 'Mode'
        mode_label.FontSize = 12
        mode_label.FontWeight = FontWeights.SemiBold
        mode_label.Foreground = _brush('#0F172A')
        mode_label.Margin = Thickness(0, 0, 0, 6)
        root.Children.Add(mode_label)

        mode_panel = StackPanel()
        mode_panel.Orientation = Orientation.Horizontal
        mode_panel.Margin = Thickness(0, 0, 0, 14)

        self._rb_type = RadioButton()
        self._rb_type.Content = 'Type'
        self._rb_type.GroupName = 'mode_group'
        self._rb_type.IsChecked = True
        self._rb_type.FontSize = 12
        self._rb_type.Foreground = _brush('#0F172A')
        self._rb_type.Margin = Thickness(0, 0, 16, 0)
        mode_panel.Children.Add(self._rb_type)

        self._rb_family = RadioButton()
        self._rb_family.Content = 'Family'
        self._rb_family.GroupName = 'mode_group'
        self._rb_family.IsChecked = False
        self._rb_family.FontSize = 12
        self._rb_family.Foreground = _brush('#0F172A')
        self._rb_family.Margin = Thickness(0, 0, 16, 0)
        mode_panel.Children.Add(self._rb_family)

        self._rb_category = RadioButton()
        self._rb_category.Content = 'Category'
        self._rb_category.GroupName = 'mode_group'
        self._rb_category.IsChecked = False
        self._rb_category.FontSize = 12
        self._rb_category.Foreground = _brush('#0F172A')
        mode_panel.Children.Add(self._rb_category)

        root.Children.Add(mode_panel)

        # ---- Scope section --------------------------------------------------
        scope_label = TextBlock()
        scope_label.Text = 'Scope'
        scope_label.FontSize = 12
        scope_label.FontWeight = FontWeights.SemiBold
        scope_label.Foreground = _brush('#0F172A')
        scope_label.Margin = Thickness(0, 0, 0, 6)
        root.Children.Add(scope_label)

        scope_panel = StackPanel()
        scope_panel.Orientation = Orientation.Horizontal
        scope_panel.Margin = Thickness(0, 0, 0, 20)

        self._rb_view = RadioButton()
        self._rb_view.Content = 'Active View'
        self._rb_view.GroupName = 'scope_group'
        self._rb_view.IsChecked = True
        self._rb_view.FontSize = 12
        self._rb_view.Foreground = _brush('#0F172A')
        self._rb_view.Margin = Thickness(0, 0, 16, 0)
        scope_panel.Children.Add(self._rb_view)

        self._rb_model = RadioButton()
        self._rb_model.Content = 'Entire Model'
        self._rb_model.GroupName = 'scope_group'
        self._rb_model.IsChecked = False
        self._rb_model.FontSize = 12
        self._rb_model.Foreground = _brush('#0F172A')
        scope_panel.Children.Add(self._rb_model)

        root.Children.Add(scope_panel)

        # ---- Select button --------------------------------------------------
        btn_select = Button()
        btn_select.Content = 'Select'
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
        # Determine scope
        scope = 'view' if self._rb_view.IsChecked else 'model'

        # Determine mode and call the appropriate function
        if self._rb_type.IsChecked:
            self.Close()
            select_similar_type(mode=scope)
        elif self._rb_family.IsChecked:
            self.Close()
            select_similar_family(mode=scope)
        else:
            self.Close()
            select_similar_category(mode=scope)


# ----------------------------------------------------------------- ENTRY
if __name__ == '__main__':
    dlg = SelectSimilarDialog()
    dlg.ShowDialog()
