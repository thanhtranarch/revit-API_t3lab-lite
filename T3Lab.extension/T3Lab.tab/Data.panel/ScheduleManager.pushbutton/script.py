# -*- coding: utf-8 -*-
__title__ = "Schedule\nManager"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Schedule Manager — Export/Import Excel and Duplicate Schedules."

# ─── Imports ──────────────────────────────────────────────────────────────────
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import os
import __builtin__
import System

from System.Windows import Window, WindowStartupLocation, WindowStyle, ResizeMode, Thickness
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, TabControl, TabItem,
    StackPanel, TextBlock, Button, Border
)
from System.Windows.Media import BrushConverter, FontFamily
from System.Windows import GridLength, GridUnitType
from System.Windows import FontWeights, HorizontalAlignment, VerticalAlignment
from System.Windows import TextWrapping, CornerRadius
from System.Windows.Controls import ScrollBarVisibility

from pyrevit import forms

# ─── Helper ───────────────────────────────────────────────────────────────────
def _brush(hex_color):
    return BrushConverter().ConvertFromString(hex_color)


# ─── Window ───────────────────────────────────────────────────────────────────
class ScheduleManagerWindow(Window):

    def __init__(self):
        self.Title = "Schedule Manager"
        self.Width = 420
        self.Height = 280
        self.Background = _brush("#F8FAFC")
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.FontFamily = FontFamily("Hanken Grotesk, Inter")

        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))

        # ── Header ──
        header = TextBlock()
        header.Text = "Schedule Manager"
        header.FontSize = 16
        header.FontWeight = FontWeights.Bold
        header.Foreground = _brush("#0F172A")
        header.Margin = Thickness(16, 12, 16, 8)
        Grid.SetRow(header, 0)
        root.Children.Add(header)

        # ── TabControl ──
        tabs = TabControl()
        tabs.Margin = Thickness(12, 0, 12, 12)
        tabs.Background = _brush("#F8FAFC")
        Grid.SetRow(tabs, 1)
        root.Children.Add(tabs)

        # Tab 1 — Export / Import Excel
        tab1 = TabItem()
        tab1.Header = "Export / Import Excel"

        panel1 = StackPanel()
        panel1.Margin = Thickness(12, 12, 12, 12)

        desc1 = TextBlock()
        desc1.Text = "Export schedules to Excel, edit, and import back into Revit."
        desc1.Foreground = _brush("#64748B")
        desc1.FontSize = 12
        desc1.TextWrapping = TextWrapping.Wrap
        desc1.Margin = Thickness(0, 0, 0, 16)
        panel1.Children.Add(desc1)

        btn1 = Button()
        btn1.Content = "Open Export / Import Tool"
        btn1.Background = _brush("#0F172A")
        btn1.Foreground = _brush("#FFFFFF")
        btn1.Height = 38
        btn1.Width = 220
        btn1.HorizontalAlignment = HorizontalAlignment.Left
        btn1.FontSize = 12
        btn1.FontWeight = FontWeights.SemiBold
        btn1.Cursor = System.Windows.Input.Cursors.Hand
        btn1.Click += self._on_export_import
        panel1.Children.Add(btn1)

        tab1.Content = panel1
        tabs.Items.Add(tab1)

        # Tab 2 — Duplicate Schedules
        tab2 = TabItem()
        tab2.Header = "Duplicate Schedules"

        panel2 = StackPanel()
        panel2.Margin = Thickness(12, 12, 12, 12)

        desc2 = TextBlock()
        desc2.Text = "Duplicate existing schedules with template and naming options."
        desc2.Foreground = _brush("#64748B")
        desc2.FontSize = 12
        desc2.TextWrapping = TextWrapping.Wrap
        desc2.Margin = Thickness(0, 0, 0, 16)
        panel2.Children.Add(desc2)

        btn2 = Button()
        btn2.Content = "Run Schedule Duplicator"
        btn2.Background = _brush("#0F172A")
        btn2.Foreground = _brush("#FFFFFF")
        btn2.Height = 38
        btn2.Width = 220
        btn2.HorizontalAlignment = HorizontalAlignment.Left
        btn2.FontSize = 12
        btn2.FontWeight = FontWeights.SemiBold
        btn2.Cursor = System.Windows.Input.Cursors.Hand
        btn2.Click += self._on_duplicate
        panel2.Children.Add(btn2)

        tab2.Content = panel2
        tabs.Items.Add(tab2)

        self.Content = root

    # ── Handlers ──────────────────────────────────────────────────────────────

    def _on_export_import(self, sender, e):
        self._launch('../ScheduleExportImportPro/script.py')

    def _on_duplicate(self, sender, e):
        self._launch('../Schedule_Copy/script.py')

    def _launch(self, rel_path):
        script_path = os.path.normpath(
            os.path.join(os.path.dirname(__file__), rel_path)
        )
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
            forms.alert("Error launching tool: {}".format(ex))


# ─── Entry Point ──────────────────────────────────────────────────────────────
if __name__ == '__main__':
    win = ScheduleManagerWindow()
    win.ShowDialog()
