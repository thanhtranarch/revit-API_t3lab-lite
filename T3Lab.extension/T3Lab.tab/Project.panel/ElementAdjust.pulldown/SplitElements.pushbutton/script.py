# -*- coding: utf-8 -*-
__title__ = "Split\nElements"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Split Elements — Split Walls, Columns, or Floors at selected levels."

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


class SplitElementsWindow(Window):
    def __init__(self):
        self.Title = "Split Elements"
        self.Width = 440
        self.Height = 320
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
        t2.Text = "Split Elements"
        t2.FontSize = 15
        t2.FontWeight = FontWeights.Bold
        t2.Foreground = _brush("#0F172A")
        hdr_sp.Children.Add(t2)
        Grid.SetRow(hdr_sp, 0)
        root.Children.Add(hdr_sp)

        tabs = TabControl()
        tabs.Margin = Thickness(0)
        Grid.SetRow(tabs, 1)

        # --- Tab: Walls ---
        walls_tab = TabItem()
        walls_tab.Header = "Walls"

        walls_inner = StackPanel()
        walls_inner.Margin = Thickness(16, 14, 16, 14)
        walls_inner.Children.Add(_make_btn(
            "Split Walls",
            "Split walls at selected levels, preserving hosted elements",
            self._on_split_walls))

        walls_tab.Content = walls_inner

        # --- Tab: Columns ---
        columns_tab = TabItem()
        columns_tab.Header = "Columns"

        columns_inner = StackPanel()
        columns_inner.Margin = Thickness(16, 14, 16, 14)
        columns_inner.Children.Add(_make_btn(
            "Split Columns",
            "Split structural columns at selected levels",
            self._on_split_columns))

        columns_tab.Content = columns_inner

        # --- Tab: Floors ---
        floors_tab = TabItem()
        floors_tab.Header = "Floors"

        floors_inner = StackPanel()
        floors_inner.Margin = Thickness(16, 14, 16, 14)
        floors_inner.Children.Add(_make_btn(
            "Split Floors",
            "Split floors with multiple disconnected boundaries into separate floors",
            self._on_split_floors))

        floors_tab.Content = floors_inner

        tabs.Items.Add(walls_tab)
        tabs.Items.Add(columns_tab)
        tabs.Items.Add(floors_tab)
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

    # Split handlers
    def _on_split_walls(self, sender, e):
        self._launch("../Split.pulldown/Wall_Split.pushbutton/script.py")

    def _on_split_columns(self, sender, e):
        self._launch("../Split.pulldown/Column_Split.pushbutton/script.py")

    def _on_split_floors(self, sender, e):
        self._launch("../Split.pulldown/Floor_Split.pushbutton/script.py")


if __name__ == '__main__':
    win = SplitElementsWindow()
    win.ShowDialog()
