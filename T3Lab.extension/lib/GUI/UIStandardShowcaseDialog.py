# -*- coding: utf-8 -*-
"""
T3Lab UI Standard Showcase Dialog
GUI classes for reviewing standard UI components and styling.
"""

import os
import sys
import clr
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows import (Thickness, GridLength, GridUnitType,
                            HorizontalAlignment, VerticalAlignment, FontWeights,
                            MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, WindowState)
from System.Windows.Controls import (RowDefinition, ColumnDefinition, Border,
                                      StackPanel, TextBlock, TextBox, Button,
                                      ComboBox, ComboBoxItem, DataGrid, Orientation,
                                      DataGridTextColumn, DataGridCheckBoxColumn,
                                      ScrollViewer, TabControl, TabItem, CheckBox)
from System.Windows.Media import SolidColorBrush, BrushConverter
from System.Windows.Data import Binding
from System.Collections.ObjectModel import ObservableCollection

from pyrevit import revit, DB, forms

# XAML Path
GUI_DIR = os.path.dirname(__file__)  # [repo]/T3Lab.extension/lib/GUI
REPO_DIR = os.path.dirname(os.path.dirname(os.path.dirname(GUI_DIR)))  # [repo]
XAML_FILE = os.path.join(REPO_DIR, '.claude', 'standard', 'UIStandardShowcase.xaml')


class ShowcaseItem(object):
    """Simple model for UI Standard Showcase Grid binding"""
    def __init__(self, item_id, name, category, status):
        self._id = item_id
        self._name = name
        self._category = category
        self._status = status

    @property
    def id(self):
        return self._id

    @property
    def name(self):
        return self._name

    @property
    def category(self):
        return self._category

    @property
    def status(self):
        return self._status


class UIStandardShowcaseWindow(forms.WPFWindow):
    """Window class for T3Lab UI Standard Showcase"""
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._load_sample_data()

    def setup_icon(self):
        """Override pyrevit's setup_icon to load custom T3Lab icon and handle exceptions gracefully."""
        try:
            # Resolve the pushbutton's icon.png path relative to this file
            current_dir = os.path.dirname(__file__)  # lib/GUI
            extension_dir = os.path.dirname(os.path.dirname(current_dir))  # T3Lab.extension
            icon_path = os.path.join(extension_dir, "T3Lab.tab", "Standard.panel", "UIStandard.pushbutton", "icon.png")
            if os.path.exists(icon_path):
                self.set_icon(icon_path)
            else:
                # Fallback to default pyRevit icon
                super(UIStandardShowcaseWindow, self).setup_icon()
        except Exception as e:
            # Silence icon loading errors to prevent window initialization crash
            print("Warning: Failed to load window icon: {}".format(e))

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self.Close()

    def _load_sample_data(self):
        # Create list of sample compliant and non-compliant items
        items = [
            ShowcaseItem("104231", "01_Plan_Mặt bằng định vị cột", "Floor Plan", "Compliant"),
            ShowcaseItem("104232", "02_Plan_Mặt bằng kích thước cột", "Floor Plan", "Compliant"),
            ShowcaseItem("104235", "Detail_Chi tiết nối thép cột", "Detail View", "Compliant"),
            ShowcaseItem("104240", "3D_Phối cảnh tổng thể", "3D View", "Non-Compliant (Template missing)"),
            ShowcaseItem("104245", "Section_Mặt cắt đứng dọc nhà", "Section", "Compliant"),
            ShowcaseItem("104250", "Elevation_Mặt đứng trục A-D", "Elevation", "Compliant"),
            ShowcaseItem("104255", "Schedule_Thống kê cốt thép dầm", "Schedule", "Compliant"),
            ShowcaseItem("104260", "Legend_Ký hiệu ghi chú chung", "Legend", "Compliant"),
            ShowcaseItem("104265", "Drafting_Chi tiết cấu tạo sê nô", "Drafting View", "Non-Compliant (Naming standard)"),
            ShowcaseItem("104272", "Elevation_Mặt đứng trục E-H", "Elevation", "Needs Review")
        ]
        self.sample_grid.ItemsSource = ObservableCollection[object](items)


def show_ui_standard_showcase():
    """Launch the UI Standard Showcase Dialog"""
    try:
        window = UIStandardShowcaseWindow()
        window.ShowDialog()
    except Exception as e:
        print("\nFATAL ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        
        MessageBox.Show(
            "Error starting UI Standard Showcase:\n\n{}".format(str(e)),
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        )
