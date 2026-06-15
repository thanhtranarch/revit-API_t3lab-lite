# -*- coding: utf-8 -*-
__title__ = "Parameter\nManager"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Parameter Manager — Transfer, Text-to-Element, and Values-to-Region tools."

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import os
import System
import __builtin__

from System.Windows import (Window, WindowStartupLocation, Thickness,
                            HorizontalAlignment, FontWeights, FontStyles,
                            TextWrapping, ResizeMode)
from System.Windows.Controls import (Grid, RowDefinition, TabControl, TabItem,
                                     StackPanel, TextBlock, Button, ScrollViewer,
                                     Border)
from System.Windows.Media import BrushConverter, FontFamily
from System.Windows import GridLength, GridUnitType

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


class ToolWindow(Window):
    def __init__(self):
        self.Title = "Parameter Manager"
        self.Width = 440
        self.Height = 374
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush("#F8FAFC")
        self.FontFamily = FontFamily("Hanken Grotesk")

        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(52)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))

        # ── Header ──────────────────────────────────────────────────────────────
        hdr_sp = StackPanel()
        hdr_sp.Margin = Thickness(16, 10, 16, 8)

        t1 = TextBlock()
        t1.Text = "T3Lab"
        t1.FontSize = 9
        t1.FontWeight = FontWeights.Bold
        t1.Foreground = _brush("#64748B")
        hdr_sp.Children.Add(t1)

        t2 = TextBlock()
        t2.Text = "Parameter Manager"
        t2.FontSize = 15
        t2.FontWeight = FontWeights.Bold
        t2.Foreground = _brush("#0F172A")
        hdr_sp.Children.Add(t2)

        Grid.SetRow(hdr_sp, 0)
        root.Children.Add(hdr_sp)

        # ── Tabs ────────────────────────────────────────────────────────────────
        tabs = TabControl()
        tabs.Background = _brush("#F8FAFC")
        tabs.BorderThickness = Thickness(0)
        Grid.SetRow(tabs, 1)

        # Tab 1 — Transfer Params
        tab1 = TabItem()
        tab1.Header = "Transfer Params"
        sv1 = ScrollViewer()
        sv1.Padding = Thickness(16, 14, 16, 14)
        panel1 = StackPanel()
        panel1.Children.Add(_make_btn(
            "Transfer Parameter Values",
            "Transfer values from one parameter to another within the same elements",
            self._open_transfer_params
        ))
        sv1.Content = panel1
        tab1.Content = sv1
        tabs.Items.Add(tab1)

        # Tab 2 — Text to Element
        tab2 = TabItem()
        tab2.Header = "Text to Element"
        sv2 = ScrollViewer()
        sv2.Padding = Thickness(16, 14, 16, 14)
        panel2 = StackPanel()
        panel2.Children.Add(_make_btn(
            "Text to Element",
            "Write text annotation data to element parameters in bulk",
            self._open_text_to_element
        ))
        sv2.Content = panel2
        tab2.Content = sv2
        tabs.Items.Add(tab2)

        # Tab 3 — Values to Region
        tab3 = TabItem()
        tab3.Header = "Values to Region"
        sv3 = ScrollViewer()
        sv3.Padding = Thickness(16, 14, 16, 14)
        panel3 = StackPanel()
        panel3.Children.Add(_make_btn(
            "Values to Filled Region",
            "Map element parameter values to filled region hatch patterns",
            self._open_values_to_region
        ))
        sv3.Content = panel3
        tab3.Content = sv3
        tabs.Items.Add(tab3)

        root.Children.Add(tabs)

        # ── Status bar ──────────────────────────────────────────────────────────
        status_border = Border()
        status_border.Background = _brush("#F8FAFC")
        status_border.BorderBrush = _brush("#E2E8F0")
        status_border.BorderThickness = Thickness(0, 1, 0, 0)
        status_border.Padding = Thickness(14, 6)
        Grid.SetRow(status_border, 2)

        status_tb = TextBlock()
        status_tb.Text = "Click a button above to launch the tool"
        status_tb.FontSize = 11
        status_tb.Foreground = _brush("#64748B")
        status_tb.FontStyle = FontStyles.Italic
        status_border.Child = status_tb
        root.Children.Add(status_border)

        self.Content = root

        # ── Copyright overlay ────────────────────────────────────────────────────
        copyright_tb = TextBlock()
        copyright_tb.Text = "© Copyright by T3Lab"
        copyright_tb.HorizontalAlignment = HorizontalAlignment.Right
        copyright_tb.VerticalAlignment = System.Windows.VerticalAlignment.Bottom
        copyright_tb.Margin = Thickness(0, 0, 14, 8)
        copyright_tb.Foreground = _brush("#F59E0B")
        copyright_tb.FontSize = 11
        copyright_tb.IsHitTestVisible = False
        root.Children.Add(copyright_tb)

    # ── Launch helpers ───────────────────────────────────────────────────────

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

    def _open_transfer_params(self, sender, e):
        self._launch("../Transfer Para.pushbutton/script.py")

    def _open_text_to_element(self, sender, e):
        self._launch("../Text to element.pushbutton/script.py")

    def _open_values_to_region(self, sender, e):
        self._launch("../Values to Filled Region .pushbutton/script.py")


if __name__ == '__main__':
    win = ToolWindow()
    win.ShowDialog()
