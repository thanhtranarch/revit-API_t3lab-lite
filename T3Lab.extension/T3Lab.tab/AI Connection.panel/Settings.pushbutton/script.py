# -*- coding: utf-8 -*-
"""Settings Dialog

Configure T3Lab API keys, LLM settings, and copy local MCP configurations.
"""
__title__ = "Settings"
__author__ = "T3Lab & Dang Quoc Truong"

import os
import sys
import clr
import json

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import (
    Window, WindowStartupLocation, Thickness, CornerRadius,
    HorizontalAlignment, VerticalAlignment, Visibility,
    GridLength, GridUnitType, FontWeights, TextWrapping,
    MessageBox, Clipboard
)
from System.Windows.Controls import *
from System.Windows.Media import BrushConverter, FontFamily

# Setup paths
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import script, forms
try:
    from core.server import get_t3labai_server
    HAS_SERVER = True
except:
    HAS_SERVER = False

class SettingsWindow(Window):
    def __init__(self):
        self.Title = "T3Lab AI & MCP Settings"
        self.Width = 500
        self.Height = 420
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = BrushConverter().ConvertFromString("#F8FAFC")
        
        # Main layout
        grid = Grid()
        grid.RowDefinitions.Add(RowDefinition())
        grid.RowDefinitions.Add(RowDefinition())
        
        # Row 1: Content
        stack = StackPanel()
        stack.Padding = Thickness(20)
        
        title = TextBlock()
        title.Text = "T3Lab AI Settings"
        title.FontSize = 18
        title.FontWeight = FontWeights.Bold
        title.Foreground = BrushConverter().ConvertFromString("#0F172A")
        title.Margin = Thickness(0, 0, 0, 15)
        stack.Children.Add(title)

        # Port Setting
        port_label = TextBlock()
        port_label.Text = "MCP Port Server:"
        port_label.FontWeight = FontWeights.SemiBold
        port_label.Foreground = BrushConverter().ConvertFromString("#0F172A")
        port_label.Margin = Thickness(0, 0, 0, 5)
        stack.Children.Add(port_label)
        
        self.port_tb = TextBox()
        server = get_t3labai_server() if HAS_SERVER else None
        self.port_tb.Text = str(server.port) if server else "48884"
        self.port_tb.Padding = Thickness(8, 6)
        self.port_tb.Margin = Thickness(0, 0, 0, 15)
        self.port_tb.TextChanged += self.port_changed
        stack.Children.Add(self.port_tb)

        # Config Snippet Label
        config_label = TextBlock()
        config_label.Text = "Claude Desktop / Cline Configuration Snippet:"
        config_label.FontWeight = FontWeights.SemiBold
        config_label.Foreground = BrushConverter().ConvertFromString("#0F172A")
        config_label.Margin = Thickness(0, 0, 0, 5)
        stack.Children.Add(config_label)

        # JSON Box
        bridge_path = os.path.join(LIB_DIR, "core", "bridge.py").replace("\\", "/")
        port_str = self.port_tb.Text
        self.snippet_text = (
            '{\n'
            '  "mcpServers": {\n'
            '    "t3lab-revit": {\n'
            '      "command": "python",\n'
            '      "args": [\n'
            '        "' + bridge_path + '",\n'
            '        "' + port_str + '"\n'
            '      ]\n'
            '    }\n'
            '  }\n'
            '}'
        )

        self.config_box = TextBox()
        self.config_box.Text = self.snippet_text
        self.config_box.IsReadOnly = True
        self.config_box.Height = 150
        self.config_box.FontFamily = FontFamily("Consolas")
        self.config_box.Padding = Thickness(10)
        self.config_box.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.config_box.Margin = Thickness(0, 0, 0, 15)
        stack.Children.Add(self.config_box)

        # Copy Button
        copy_btn = Button()
        copy_btn.Content = "Copy Configuration to Clipboard"
        copy_btn.Padding = Thickness(12, 8)
        copy_btn.Background = BrushConverter().ConvertFromString("#0F172A")
        copy_btn.Foreground = BrushConverter().ConvertFromString("#FFFFFF")
        copy_btn.BorderThickness = Thickness(0)
        copy_btn.Cursor = System.Windows.Input.Cursors.Hand
        copy_btn.Click += self.copy_clicked
        stack.Children.Add(copy_btn)

        grid.Children.Add(stack)
        self.Content = grid

    def copy_clicked(self, sender, e):
        try:
            Clipboard.SetText(self.snippet_text)
            MessageBox.Show("Configuration copied successfully!", "Success")
        except Exception as ex:
            MessageBox.Show("Failed to copy configuration: {}".format(ex), "Error")

    def port_changed(self, sender, e):
        bridge_path = os.path.join(LIB_DIR, "core", "bridge.py").replace("\\", "/")
        port_str = self.port_tb.Text
        self.snippet_text = (
            '{\n'
            '  "mcpServers": {\n'
            '    "t3lab-revit": {\n'
            '      "command": "python",\n'
            '      "args": [\n'
            '        "' + bridge_path + '",\n'
            '        "' + port_str + '"\n'
            '      ]\n'
            '    }\n'
            '  }\n'
            '}'
        )
        self.config_box.Text = self.snippet_text

def main():
    window = SettingsWindow()
    window.ShowDialog()

if __name__ == '__main__':
    main()
