# -*- coding: utf-8 -*-
__title__ = "Datum\nManager"
__author__ = "Tran Tien Thanh & Dang Quoc Truong"
__doc__ = "Datum Manager — Manage Grids and Levels in one unified tool. Save/restore grid positions, align extents, and convert 2D/3D."

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import os
import System
import __builtin__

from System.Windows import (Window, WindowStartupLocation, Thickness,
                            HorizontalAlignment, VerticalAlignment,
                            FontWeights, TextWrapping, ResizeMode,
                            GridLength, GridUnitType)
from System.Windows.Controls import (Grid, RowDefinition, TabControl, TabItem,
                                     StackPanel, TextBlock, Button, ScrollViewer,
                                     Border)
from System.Windows.Media import BrushConverter, FontFamily

from pyrevit import forms


def _brush(h):
    return BrushConverter().ConvertFromString(h)


def _make_btn(label, desc, handler):
    sp = StackPanel()
    sp.Margin = Thickness(0, 0, 0, 12)
    b = Button()
    b.Content = label
    b.Height = 36
    b.HorizontalAlignment = HorizontalAlignment.Stretch
    b.Background = _brush("#0F172A")
    b.Foreground = _brush("#FFFFFF")
    b.FontWeight = FontWeights.SemiBold
    b.FontSize = 12
    b.BorderThickness = Thickness(0)
    b.Cursor = System.Windows.Input.Cursors.Hand
    b.Click += handler
    sp.Children.Add(b)
    if desc:
        d = TextBlock()
        d.Text = desc
        d.FontSize = 11
        d.Foreground = _brush("#64748B")
        d.TextWrapping = TextWrapping.Wrap
        d.Margin = Thickness(0, 3, 0, 0)
        sp.Children.Add(d)
    return sp


class DatumManagerWindow(Window):
    def __init__(self):
        self.Title = "Datum Manager"
        self.Width = 460
        self.Height = 480
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush("#F8FAFC")
        self.FontFamily = FontFamily("Hanken Grotesk")

        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(52)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))

        # Header
        hdr_sp = StackPanel()
        hdr_sp.Margin = Thickness(16, 10, 16, 8)
        t1 = TextBlock()
        t1.Text = "T3Lab"
        t1.FontSize = 9
        t1.FontWeight = FontWeights.Bold
        t1.Foreground = _brush("#64748B")
        hdr_sp.Children.Add(t1)
        t2 = TextBlock()
        t2.Text = "Datum Manager"
        t2.FontSize = 15
        t2.FontWeight = FontWeights.Bold
        t2.Foreground = _brush("#0F172A")
        hdr_sp.Children.Add(t2)
        Grid.SetRow(hdr_sp, 0)
        root.Children.Add(hdr_sp)

        tabs = TabControl()
        tabs.Margin = Thickness(0)
        Grid.SetRow(tabs, 1)

        # --- Tab: Grids ---
        grids_tab = TabItem()
        grids_tab.Header = "Grids"

        grids_inner = StackPanel()
        grids_inner.Margin = Thickness(16, 14, 16, 14)
        grids_inner.Children.Add(_make_btn(
            "Save Grids",
            "Snapshot current grid head/tail positions in this view",
            self._on_save_grids))
        grids_inner.Children.Add(_make_btn(
            "Restore Grids",
            "Restore saved positions to selected grids",
            self._on_restore_grids))
        grids_inner.Children.Add(_make_btn(
            "Restore All Grids",
            "Restore saved positions to all grids in the active view",
            self._on_restore_all_grids))
        grids_inner.Children.Add(_make_btn(
            "Align Gridlines",
            "Align 2D grid extents to a reference line",
            self._on_align_gridlines))
        grids_inner.Children.Add(_make_btn(
            "Convert Grid 2D/3D",
            "Toggle grids between 2D and 3D datum extent mode",
            self._on_convert_grid))

        grids_sv = ScrollViewer()
        grids_sv.Content = grids_inner
        grids_tab.Content = grids_sv

        # --- Tab: Levels ---
        levels_tab = TabItem()
        levels_tab.Header = "Levels"

        levels_inner = StackPanel()
        levels_inner.Margin = Thickness(16, 14, 16, 14)
        levels_inner.Children.Add(_make_btn(
            "Align Levels",
            "Align 2D level extents to a reference line",
            self._on_align_levels))
        levels_inner.Children.Add(_make_btn(
            "Convert Level 2D/3D",
            "Toggle levels between 2D and 3D datum extent mode",
            self._on_convert_level))

        levels_tab.Content = levels_inner

        tabs.Items.Add(grids_tab)
        tabs.Items.Add(levels_tab)
        root.Children.Add(tabs)

        # Status bar
        status_border = Border()
        status_border.Background = _brush("#F8FAFC")
        status_border.BorderBrush = _brush("#E2E8F0")
        status_border.BorderThickness = Thickness(0, 1, 0, 0)
        status_border.Padding = Thickness(14, 6)
        status_lbl = TextBlock()
        status_lbl.Text = "Click a button above to launch the tool"
        status_lbl.FontSize = 11
        status_lbl.FontStyle = System.Windows.FontStyles.Italic
        status_lbl.Foreground = _brush("#64748B")
        status_border.Child = status_lbl
        Grid.SetRow(status_border, 2)
        root.Children.Add(status_border)

        # Copyright overlay
        copyright = TextBlock()
        copyright.Text = "© Copyright by T3Lab"
        copyright.HorizontalAlignment = HorizontalAlignment.Right
        copyright.VerticalAlignment = VerticalAlignment.Bottom
        copyright.Margin = Thickness(0, 0, 14, 8)
        copyright.Foreground = _brush("#F59E0B")
        copyright.FontSize = 11
        copyright.IsHitTestVisible = False
        root.Children.Add(copyright)

        self.Content = root

    def _launch(self, rel_path):
        script_path = os.path.normpath(
            os.path.join(os.path.dirname(__file__), rel_path))
        self.Close()
        g = {'__name__': '__main__', '__file__': script_path,
             '__builtins__': __builtin__, '__revit__': __revit__}
        try:
            execfile(script_path, g)
        except Exception as ex:
            forms.alert("Error launching tool:\n{}".format(ex))

    # Grids handlers
    def _on_save_grids(self, sender, e):
        self._launch(
            "../../../Annotation.panel/Graphic 2.stack/Grids.pulldown/Save Grids.pushbutton/script.py")

    def _on_restore_grids(self, sender, e):
        self._launch(
            "../../../Annotation.panel/Graphic 2.stack/Grids.pulldown/Restore Grids.pushbutton/script.py")

    def _on_restore_all_grids(self, sender, e):
        self._launch(
            "../../../Annotation.panel/Graphic 2.stack/Grids.pulldown/Restore All Grids.pushbutton/script.py")

    def _on_align_gridlines(self, sender, e):
        self._launch(
            "../Datum.pulldown/Gridline.pulldown/Align Gridline.pushbutton/script.py")

    def _on_convert_grid(self, sender, e):
        self._launch(
            "../Datum.pulldown/Gridline.pulldown/ConvertGridline.pushbutton/script.py")

    # Levels handlers
    def _on_align_levels(self, sender, e):
        self._launch(
            "../Datum.pulldown/Level.pulldown/Align Level.pushbutton/script.py")

    def _on_convert_level(self, sender, e):
        self._launch(
            "../Datum.pulldown/Level.pulldown/ConvertLevel.pushbutton/script.py")


if __name__ == '__main__':
    win = DatumManagerWindow()
    win.ShowDialog()
