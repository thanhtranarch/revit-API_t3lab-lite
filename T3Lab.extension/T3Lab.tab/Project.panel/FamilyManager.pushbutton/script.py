# -*- coding: utf-8 -*-
__title__ = "Family\nManager"
__author__ = "Tran Tien Thanh & Dang Quoc Truong"
__doc__ = "Family Manager — Manage families and load new families in one dialog."

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
        self.Title = "Family Manager"
        self.Width = 440
        self.Height = 314
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
        t2.Text = "Family Manager"
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

        # Tab 1 — Family Management
        tab1 = TabItem()
        tab1.Header = "Family Management"
        sv1 = ScrollViewer()
        sv1.Padding = Thickness(16, 14, 16, 14)
        panel1 = StackPanel()
        panel1.Children.Add(_make_btn(
            "Open Family Management",
            "Batch rename families/types, modify case, customize prefix/suffix and assign worksets",
            self._open_family_management
        ))
        sv1.Content = panel1
        tab1.Content = sv1
        tabs.Items.Add(tab1)

        # Tab 2 — Load Family
        tab2 = TabItem()
        tab2.Header = "Load Family"
        sv2 = ScrollViewer()
        sv2.Padding = Thickness(16, 14, 16, 14)
        panel2 = StackPanel()
        panel2.Children.Add(_make_btn(
            "Load Family",
            "Load Revit families from local disk or cloud library",
            self._open_load_family
        ))
        sv2.Content = panel2
        tab2.Content = sv2
        tabs.Items.Add(tab2)

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

    def _open_family_management(self, sender, e):
        self._launch("../Family Work 2.stack/Family Management.pushbutton/script.py")

    def _open_load_family(self, sender, e):
        self._launch("../Family Work 2.stack/Load Family.pushbutton/script.py")


if __name__ == '__main__':
    win = ToolWindow()
    win.ShowDialog()
