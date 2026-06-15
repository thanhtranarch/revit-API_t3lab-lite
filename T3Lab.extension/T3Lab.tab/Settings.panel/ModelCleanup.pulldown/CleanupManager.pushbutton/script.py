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
                            FontWeights, TextWrapping, ResizeMode,
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


class CleanupManagerWindow(Window):
    def __init__(self):
        self.Title = "Cleanup Manager"
        self.Width = 440
        self.Height = 380
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.FontFamily = FontFamily("Hanken Grotesk, Inter")
        self.Background = _brush("#FFFFFF")

        root = Grid()
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(52)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(34)))

        # Header
        hdr_sp = StackPanel()
        hdr_sp.Margin = Thickness(16, 10, 16, 8)
        t1 = TextBlock()
        t1.Text = "T3Lab"
        t1.FontSize = 9
        t1.FontWeight = FontWeights.Bold
        t1.Foreground = _brush("#64748B")
        hdr_sp.Children.Add(t1)
        t2 = TextBlock()
        t2.Text = "Cleanup Manager"
        t2.FontSize = 15
        t2.FontWeight = FontWeights.Bold
        t2.Foreground = _brush("#0F172A")
        hdr_sp.Children.Add(t2)
        Grid.SetRow(hdr_sp, 0)
        root.Children.Add(hdr_sp)

        tabs = TabControl()
        tabs.Background = _brush("#FFFFFF")
        tabs.BorderBrush = _brush("#E2E8F0")
        tabs.BorderThickness = Thickness(1)
        tabs.FontFamily = FontFamily("Hanken Grotesk, Inter")
        tabs.FontSize = 12
        tabs.Margin = Thickness(0)
        Grid.SetRow(tabs, 1)

        # ── Tab: Purge ──────────────────────────────────────────────────────
        purge_panel = StackPanel()
        purge_panel.Margin = Thickness(16, 14, 16, 14)
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
        delete_panel.Margin = Thickness(16, 14, 16, 14)
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

        # Status bar
        status_border = Border()
        status_border.Background = _brush("#F8FAFC")
        status_border.BorderBrush = _brush("#E2E8F0")
        status_border.BorderThickness = Thickness(0, 1, 0, 0)
        status_border.Padding = Thickness(14, 6)
        status_lbl = TextBlock()
        status_lbl.Text = "Click a button above to launch the tool"
        status_lbl.FontSize = 11
        status_lbl.FontStyle = System.Windows.FontStyles.Italic
        status_lbl.Foreground = _brush("#64748B")
        status_border.Child = status_lbl
        Grid.SetRow(status_border, 2)
        root.Children.Add(status_border)

        # Copyright overlay
        copyright = TextBlock()
        copyright.Text = "© Copyright by T3Lab"
        copyright.HorizontalAlignment = HorizontalAlignment.Right
        copyright.VerticalAlignment = VerticalAlignment.Bottom
        copyright.Margin = Thickness(0, 0, 14, 8)
        copyright.Foreground = _brush("#F59E0B")
        copyright.FontSize = 11
        copyright.IsHitTestVisible = False
        root.Children.Add(copyright)

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
