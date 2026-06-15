# -*- coding: utf-8 -*-
__title__ = "View\nHub"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "View Hub — Manage Views, View Templates, and Room Plans in one place."

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


class ViewHubWindow(Window):
    def __init__(self):
        self.Title = "View Hub"
        self.Width = 440
        self.Height = 340
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
        t2.Text = "View Hub"
        t2.FontSize = 15
        t2.FontWeight = FontWeights.Bold
        t2.Foreground = _brush("#0F172A")
        hdr_sp.Children.Add(t2)

        Grid.SetRow(hdr_sp, 0)
        root.Children.Add(hdr_sp)

        # --- Tabs ---
        tabs = TabControl()
        Grid.SetRow(tabs, 1)

        # Tab 1: Manage Views
        p1 = StackPanel()
        p1.Children.Add(_make_btn(
            "Open View Manager",
            "Advanced view management — filter, rename, and organize views",
            self._on_view_manager
        ))
        tabs.Items.Add(_make_tab("Manage Views", p1))

        # Tab 2: View Templates
        p2 = StackPanel()
        p2.Children.Add(_make_btn(
            "Open View Templates",
            "Manage and apply view templates across views",
            self._on_view_templates
        ))
        tabs.Items.Add(_make_tab("View Templates", p2))

        # Tab 3: Room Plans
        p3 = StackPanel()
        p3.Children.Add(_make_btn(
            "Create Room Plans",
            "Generate plan views for each room in the model",
            self._on_room_plans
        ))
        tabs.Items.Add(_make_tab("Room Plans", p3))

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

    def _on_view_manager(self, sender, e):
        self._launch("../ViewManager.pushbutton/script.py")

    def _on_view_templates(self, sender, e):
        self._launch("../ViewTemplate.pushbutton/script.py")

    def _on_room_plans(self, sender, e):
        self._launch("../Create Room Plan.pushbutton/script.py")


if __name__ == '__main__':
    win = ViewHubWindow()
    win.ShowDialog()
