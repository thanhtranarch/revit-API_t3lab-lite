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
                            FontWeights, TextWrapping, ResizeMode)
from System.Windows.Controls import (Grid, RowDefinition, TabControl, TabItem,
                                     StackPanel, TextBlock, Button, ScrollViewer)
from System.Windows.Media import BrushConverter, FontFamily
from System.Windows import GridLength, GridUnitType

from pyrevit import forms


def _brush(h):
    return BrushConverter().ConvertFromString(h)


def _make_btn(label, desc, handler, bg="#0F172A"):
    sp = StackPanel()
    sp.Margin = Thickness(0, 0, 0, 12)
    b = Button()
    b.Content = label
    b.Height = 36
    b.HorizontalAlignment = HorizontalAlignment.Left
    b.MinWidth = 200
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


class CADToElementsWindow(Window):
    def __init__(self):
        self.Title = "CAD to Elements"
        self.Width = 440
        self.Height = 280
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.FontFamily = FontFamily("Hanken Grotesk, Inter")
        self.Background = _brush("#FFFFFF")

        root = Grid()
        root.Margin = Thickness(16, 16, 16, 16)

        tabs = TabControl()
        tabs.Background = _brush("#FFFFFF")
        tabs.BorderBrush = _brush("#E2E8F0")
        tabs.BorderThickness = Thickness(1)
        tabs.FontFamily = FontFamily("Hanken Grotesk, Inter")
        tabs.FontSize = 12

        # ── Tab: Walls ──────────────────────────────────────────────────────
        walls_panel = StackPanel()
        walls_panel.Margin = Thickness(12, 12, 12, 12)
        walls_panel.Children.Add(
            _make_btn(
                "Convert to Walls",
                "Convert CAD linework into Revit wall elements",
                self._on_cad_to_wall,
            )
        )

        walls_scroll = ScrollViewer()
        walls_scroll.Content = walls_panel
        walls_scroll.VerticalScrollBarVisibility = (
            System.Windows.Controls.ScrollBarVisibility.Auto
        )

        tab_walls = TabItem()
        tab_walls.Header = "Walls"
        tab_walls.Content = walls_scroll
        tabs.Items.Add(tab_walls)

        # ── Tab: Floors ─────────────────────────────────────────────────────
        floors_panel = StackPanel()
        floors_panel.Margin = Thickness(12, 12, 12, 12)
        floors_panel.Children.Add(
            _make_btn(
                "Convert to Floors",
                "Convert CAD linework into Revit floor elements",
                self._on_cad_to_floor,
            )
        )

        floors_scroll = ScrollViewer()
        floors_scroll.Content = floors_panel
        floors_scroll.VerticalScrollBarVisibility = (
            System.Windows.Controls.ScrollBarVisibility.Auto
        )

        tab_floors = TabItem()
        tab_floors.Header = "Floors"
        tab_floors.Content = floors_scroll
        tabs.Items.Add(tab_floors)

        # ── Tab: Beams ──────────────────────────────────────────────────────
        beams_panel = StackPanel()
        beams_panel.Margin = Thickness(12, 12, 12, 12)
        beams_panel.Children.Add(
            _make_btn(
                "Convert to Beams",
                "Convert CAD linework into Revit beam structural elements",
                self._on_cad_to_beam,
            )
        )

        beams_scroll = ScrollViewer()
        beams_scroll.Content = beams_panel
        beams_scroll.VerticalScrollBarVisibility = (
            System.Windows.Controls.ScrollBarVisibility.Auto
        )

        tab_beams = TabItem()
        tab_beams.Header = "Beams"
        tab_beams.Content = beams_scroll
        tabs.Items.Add(tab_beams)

        root.Children.Add(tabs)
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
