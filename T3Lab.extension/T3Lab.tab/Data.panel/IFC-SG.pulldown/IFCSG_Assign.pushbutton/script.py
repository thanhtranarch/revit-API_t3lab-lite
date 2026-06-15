# -*- coding: utf-8 -*-
__title__ = "IFC-SG\nAssign"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "IFC-SG Assign — Auto or Manual assignment of IFC export classes."

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import os
import System
import __builtin__

from System.Windows import (Window, WindowStartupLocation, Thickness,
                            HorizontalAlignment, FontWeights, TextWrapping, ResizeMode)
from System.Windows.Controls import (Grid, RowDefinition, TabControl, TabItem,
                                     StackPanel, TextBlock, Button, ScrollViewer)
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
    b.HorizontalAlignment = HorizontalAlignment.Left
    b.MinWidth = 200
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


def _make_tab(header, content_panel):
    item = TabItem()
    item.Header = header
    sv = ScrollViewer()
    sv.Padding = Thickness(16, 14, 16, 14)
    sv.Content = content_panel
    item.Content = sv
    return item


class IFCSGAssignWindow(Window):
    def __init__(self):
        self.Title = "IFC-SG Assignment"
        self.Width = 440
        self.Height = 280
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush("#F8FAFC")
        self.FontFamily = FontFamily("Hanken Grotesk")

        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(52)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))

        # --- Header ---
        hdr_sp = StackPanel()
        hdr_sp.Margin = Thickness(16, 10, 16, 8)

        t1 = TextBlock()
        t1.Text = "T3Lab"
        t1.FontSize = 9
        t1.FontWeight = FontWeights.Bold
        t1.Foreground = _brush("#64748B")
        hdr_sp.Children.Add(t1)

        t2 = TextBlock()
        t2.Text = "IFC-SG Assignment"
        t2.FontSize = 15
        t2.FontWeight = FontWeights.Bold
        t2.Foreground = _brush("#0F172A")
        hdr_sp.Children.Add(t2)

        Grid.SetRow(hdr_sp, 0)
        root.Children.Add(hdr_sp)

        # --- Tabs ---
        tabs = TabControl()
        Grid.SetRow(tabs, 1)

        # Tab 1: Auto Assign
        p1 = StackPanel()
        p1.Children.Add(_make_btn(
            "Auto Assign IFC Classes",
            "Automatically assign IFC classes per Family Type using LTA mapping table",
            self._on_auto_assign
        ))
        tabs.Items.Add(_make_tab("Auto Assign", p1))

        # Tab 2: Manual Assign
        p2 = StackPanel()
        p2.Children.Add(_make_btn(
            "Manual Assign IFC Classes",
            "Manually review and assign IFC export classes with filtering and search",
            self._on_manual_assign
        ))
        tabs.Items.Add(_make_tab("Manual Assign", p2))

        root.Children.Add(tabs)
        self.Content = root

    def _launch(self, rel_path):
        script_path = os.path.normpath(
            os.path.join(os.path.dirname(__file__), rel_path))
        self.Close()
        g = {
            '__name__': '__main__',
            '__file__': script_path,
            '__builtins__': __builtin__,
            '__revit__': __revit__,
        }
        try:
            execfile(script_path, g)
        except Exception as ex:
            forms.alert("Error launching tool:\n{}".format(ex))

    def _on_auto_assign(self, sender, e):
        self._launch("../Auto Assign.pushbutton/script.py")

    def _on_manual_assign(self, sender, e):
        self._launch("../Manual Assign.pushbutton/script.py")


if __name__ == '__main__':
    win = IFCSGAssignWindow()
    win.ShowDialog()
