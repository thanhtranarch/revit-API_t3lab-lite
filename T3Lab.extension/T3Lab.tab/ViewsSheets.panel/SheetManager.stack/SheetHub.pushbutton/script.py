# -*- coding: utf-8 -*-
__title__ = "Sheet\nHub"
__author__ = "Dang Quoc Truong & Tran Tien Thanh"
__doc__ = "Sheet Hub — Manage Sheets and Re-number Sheets in one dialog."

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import os
import System
import __builtin__

from System.Windows import (Window, WindowStartupLocation, Thickness,
                            HorizontalAlignment, VerticalAlignment,
                            FontWeights, FontStyles, TextWrapping, ResizeMode,
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


def _make_tab(header, content_panel):
    item = TabItem()
    item.Header = header
    sv = ScrollViewer()
    sv.Padding = Thickness(16, 14, 16, 14)
    sv.Content = content_panel
    item.Content = sv
    return item


class SheetHubWindow(Window):
    def __init__(self):
        self.Title = "Sheet Hub"
        self.Width = 440
        self.Height = 320
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush("#F8FAFC")
        self.FontFamily = FontFamily("Hanken Grotesk")

        root = Grid()
        # 3 rows: header [52], content [*], status bar [34]
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(52)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))

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
        t2.Text = "Sheet Hub"
        t2.FontSize = 15
        t2.FontWeight = FontWeights.Bold
        t2.Foreground = _brush("#0F172A")
        hdr_sp.Children.Add(t2)

        Grid.SetRow(hdr_sp, 0)
        root.Children.Add(hdr_sp)

        # --- Tabs ---
        tabs = TabControl()
        Grid.SetRow(tabs, 1)

        # Tab 1: Manage Sheets
        p1 = StackPanel()
        p1.Children.Add(_make_btn(
            "Open Sheet Manager",
            "Advanced sheet management — create, rename, and organize sheets",
            self._on_sheet_manager
        ))
        tabs.Items.Add(_make_tab("Manage Sheets", p1))

        # Tab 2: Re-number Sheets
        p2 = StackPanel()
        p2.Children.Add(_make_btn(
            "Re-number Sheets",
            "Bulk re-number sheets with prefix, suffix, and sequence options",
            self._on_sheet_renumber
        ))
        tabs.Items.Add(_make_tab("Re-number Sheets", p2))

        root.Children.Add(tabs)

        # --- Status bar ---
        status_border = Border()
        status_border.Background = _brush("#F8FAFC")
        status_border.BorderBrush = _brush("#E2E8F0")
        status_border.BorderThickness = Thickness(0, 1, 0, 0)
        status_border.Padding = Thickness(14, 6, 14, 6)

        status_tb = TextBlock()
        status_tb.Text = "Click a button above to launch the tool"
        status_tb.FontSize = 11
        status_tb.Foreground = _brush("#64748B")
        status_tb.FontStyle = FontStyles.Italic
        status_border.Child = status_tb

        Grid.SetRow(status_border, 2)
        root.Children.Add(status_border)

        # --- Copyright overlay ---
        copyright_tb = TextBlock()
        copyright_tb.Text = "© Copyright by T3Lab"
        copyright_tb.HorizontalAlignment = HorizontalAlignment.Right
        copyright_tb.VerticalAlignment = VerticalAlignment.Bottom
        copyright_tb.Margin = Thickness(0, 0, 14, 8)
        copyright_tb.Foreground = _brush("#F59E0B")
        copyright_tb.FontSize = 11
        copyright_tb.IsHitTestVisible = False
        root.Children.Add(copyright_tb)

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

    def _on_sheet_manager(self, sender, e):
        self._launch("../SheetManager.pushbutton/script.py")

    def _on_sheet_renumber(self, sender, e):
        self._launch("../Sheet re-number.pushbutton/script.py")


if __name__ == '__main__':
    win = SheetHubWindow()
    win.ShowDialog()
