# -*- coding: utf-8 -*-
__title__ = "CAD to\nElements"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "CAD to Elements — Convert CAD linework into Walls, Floors, or Beams."

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


def _make_btn(label, desc, handler, bg="#0F172A"):
    sp = StackPanel()
    sp.Margin = Thickness(0, 0, 0, 12)
    b = Button()
    b.Content = label
    b.Height = 36
    b.HorizontalAlignment = HorizontalAlignment.Stretch
    b.Background = _brush(bg)
    b.Foreground = _brush("#FFFFFF")
    b.FontFamily = FontFamily("Hanken Grotesk, Inter")
    b.FontWeight = FontWeights.SemiBold
    b.FontSize = 12
    b.BorderThickness = Thickness(0)
    b.Cursor = System.Windows.Input.Cursors.Hand
    b.Click += handler
    sp.Children.Add(b)
    if desc:
        d = TextBlock()
        d.Text = desc
        d.FontFamily = FontFamily("Inter")
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


class CADToElementsWindow(Window):
    def __init__(self):
        self.Title = "CAD to Elements"
        self.Width = 440
        self.Height = 340
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.FontFamily = FontFamily("Hanken Grotesk, Inter")
        self.Background = _brush("#F8FAFC")

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
        t2.Text = "CAD to Elements"
        t2.FontSize = 15
        t2.FontWeight = FontWeights.Bold
        t2.Foreground = _brush("#0F172A")
        hdr_sp.Children.Add(t2)

        Grid.SetRow(hdr_sp, 0)
        root.Children.Add(hdr_sp)

        # --- Tabs ---
        tabs = TabControl()
        tabs.Background = _brush("#FFFFFF")
        tabs.BorderBrush = _brush("#E2E8F0")
        tabs.BorderThickness = Thickness(1)
        tabs.FontFamily = FontFamily("Hanken Grotesk, Inter")
        tabs.FontSize = 12

        # Tab: Walls
        walls_panel = StackPanel()
        walls_panel.Children.Add(
            _make_btn(
                "Convert to Walls",
                "Convert CAD linework into Revit wall elements",
                self._on_cad_to_wall,
            )
        )
        tabs.Items.Add(_make_tab("Walls", walls_panel))

        # Tab: Floors
        floors_panel = StackPanel()
        floors_panel.Children.Add(
            _make_btn(
                "Convert to Floors",
                "Convert CAD linework into Revit floor elements",
                self._on_cad_to_floor,
            )
        )
        tabs.Items.Add(_make_tab("Floors", floors_panel))

        # Tab: Beams
        beams_panel = StackPanel()
        beams_panel.Children.Add(
            _make_btn(
                "Convert to Beams",
                "Convert CAD linework into Revit beam structural elements",
                self._on_cad_to_beam,
            )
        )
        tabs.Items.Add(_make_tab("Beams", beams_panel))

        Grid.SetRow(tabs, 1)
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

    # ── Launch helpers ───────────────────────────────────────────────────────

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
            forms.alert("Error launching tool:\n{}".format(ex))

    def _on_cad_to_wall(self, sender, e):
        self._launch("../CadtoWall.pushbutton/script.py")

    def _on_cad_to_floor(self, sender, e):
        self._launch("../CadtoFloor.pushbutton/script.py")

    def _on_cad_to_beam(self, sender, e):
        self._launch("../Beam.pushbutton/script.py")


if __name__ == '__main__':
    win = CADToElementsWindow()
    win.ShowDialog()
