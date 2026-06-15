# -*- coding: utf-8 -*-
__title__ = "Workset\nManager"
__author__ = "Tran Tien Thanh"
__doc__ = "Workset Manager — Create Central File and Manage Worksets."

# ─── Imports ──────────────────────────────────────────────────────────────────
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import os
import __builtin__

from System.Windows import Window, WindowStartupLocation, WindowStyle, ResizeMode, Thickness
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, TabControl, TabItem,
    StackPanel, TextBlock, Button, Border
)
from System.Windows.Media import BrushConverter, FontFamily
from System.Windows import GridLength, GridUnitType
from System.Windows import FontWeights, HorizontalAlignment, VerticalAlignment
from System.Windows import TextWrapping

from Autodesk.Revit.DB import SaveAsOptions, WorksharingSaveAsOptions
from Autodesk.Revit.UI import TaskDialog

from pyrevit import forms

# ─── Revit context ────────────────────────────────────────────────────────────
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# ─── Helper ───────────────────────────────────────────────────────────────────
def _brush(hex_color):
    return BrushConverter().ConvertFromString(hex_color)


def _create_central():
    file_name = forms.save_file()
    if file_name:
        base = file_name.rsplit(".", 1)[0]
        path = base + ".rvt"
        try:
            opts = WorksharingSaveAsOptions()
            opts.SaveAsCentral = True
            sao = SaveAsOptions()
            sao.MaximumBackups = 10
            sao.OverwriteExistingFile = True
            sao.SetWorksharingOptions(opts)
            doc.SaveAs(path, sao)
            TaskDialog.Show("Success", "Central File created at:\n" + path)
        except Exception as e:
            forms.alert("Failed: {}".format(e))
    else:
        forms.alert("No file path selected.")


# ─── Window ───────────────────────────────────────────────────────────────────
class WorksetManagerWindow(Window):

    def __init__(self):
        self.Title = "Workset Manager"
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
        header.Text = "Workset Manager"
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

        # Tab 1 — Create Central File
        tab1 = TabItem()
        tab1.Header = "Create Central File"

        panel1 = StackPanel()
        panel1.Margin = Thickness(12, 12, 12, 12)

        desc1 = TextBlock()
        desc1.Text = "Save the current project as a Central File for worksharing."
        desc1.Foreground = _brush("#64748B")
        desc1.FontSize = 12
        desc1.TextWrapping = TextWrapping.Wrap
        desc1.Margin = Thickness(0, 0, 0, 16)
        panel1.Children.Add(desc1)

        btn1 = Button()
        btn1.Content = "Create Central File"
        btn1.Background = _brush("#0F172A")
        btn1.Foreground = _brush("#FFFFFF")
        btn1.Height = 38
        btn1.Width = 220
        btn1.HorizontalAlignment = HorizontalAlignment.Left
        btn1.FontSize = 12
        btn1.FontWeight = FontWeights.SemiBold
        btn1.Cursor = System.Windows.Input.Cursors.Hand
        btn1.Click += self._on_create_central
        panel1.Children.Add(btn1)

        tab1.Content = panel1
        tabs.Items.Add(tab1)

        # Tab 2 — Manage Worksets
        tab2 = TabItem()
        tab2.Header = "Manage Worksets"

        panel2 = StackPanel()
        panel2.Margin = Thickness(12, 12, 12, 12)

        desc2 = TextBlock()
        desc2.Text = "View and manage user worksets in the active project."
        desc2.Foreground = _brush("#64748B")
        desc2.FontSize = 12
        desc2.TextWrapping = TextWrapping.Wrap
        desc2.Margin = Thickness(0, 0, 0, 16)
        panel2.Children.Add(desc2)

        btn2 = Button()
        btn2.Content = "Open Workset Manager"
        btn2.Background = _brush("#0F172A")
        btn2.Foreground = _brush("#FFFFFF")
        btn2.Height = 38
        btn2.Width = 220
        btn2.HorizontalAlignment = HorizontalAlignment.Left
        btn2.FontSize = 12
        btn2.FontWeight = FontWeights.SemiBold
        btn2.Cursor = System.Windows.Input.Cursors.Hand
        btn2.Click += self._on_manage_worksets
        panel2.Children.Add(btn2)

        tab2.Content = panel2
        tabs.Items.Add(tab2)

        self.Content = root

    # ── Handlers ──────────────────────────────────────────────────────────────

    def _on_create_central(self, sender, e):
        _create_central()

    def _on_manage_worksets(self, sender, e):
        self._launch('../Workset.stack/Workset.pushbutton/script.py')

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
    win = WorksetManagerWindow()
    win.ShowDialog()
