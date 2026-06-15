# -*- coding: utf-8 -*-
__title__ = "Cleanup\nManager"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Cleanup Manager — Smart Purge, Advanced Purge, and Smart Delete in one dialog."

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


class CleanupManagerWindow(Window):
    def __init__(self):
        self.Title = "Cleanup Manager"
        self.Width = 440
        self.Height = 340
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

        # ── Tab: Purge ──────────────────────────────────────────────────────
        purge_panel = StackPanel()
        purge_panel.Margin = Thickness(12, 12, 12, 12)
        purge_panel.Children.Add(
            _make_btn(
                "Smart Purge",
                "Standard purge with safety checks — removes unused elements",
                self._on_smart_purge,
            )
        )
        purge_panel.Children.Add(
            _make_btn(
                "Advanced Purge",
                "Deep purge including dependent elements (use with caution)",
                self._on_advanced_purge,
                bg="#EF4444",
            )
        )

        purge_scroll = ScrollViewer()
        purge_scroll.Content = purge_panel
        purge_scroll.VerticalScrollBarVisibility = (
            System.Windows.Controls.ScrollBarVisibility.Auto
        )

        tab_purge = TabItem()
        tab_purge.Header = "Purge"
        tab_purge.Content = purge_scroll
        tabs.Items.Add(tab_purge)

        # ── Tab: Delete ─────────────────────────────────────────────────────
        delete_panel = StackPanel()
        delete_panel.Margin = Thickness(12, 12, 12, 12)
        delete_panel.Children.Add(
            _make_btn(
                "Smart Delete",
                "Delete elements with dependency checking and reporting",
                self._on_smart_delete,
            )
        )

        delete_scroll = ScrollViewer()
        delete_scroll.Content = delete_panel
        delete_scroll.VerticalScrollBarVisibility = (
            System.Windows.Controls.ScrollBarVisibility.Auto
        )

        tab_delete = TabItem()
        tab_delete.Header = "Delete"
        tab_delete.Content = delete_scroll
        tabs.Items.Add(tab_delete)

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

    def _on_smart_purge(self, sender, e):
        self._launch("../SmartPurge.pushbutton/script.py")

    def _on_advanced_purge(self, sender, e):
        self._launch("../AdvancedPurge.pushbutton/script.py")

    def _on_smart_delete(self, sender, e):
        self._launch("../SmartDelete.pushbutton/script.py")


if __name__ == '__main__':
    win = CleanupManagerWindow()
    win.ShowDialog()
