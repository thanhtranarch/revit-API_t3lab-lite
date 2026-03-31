# -*- coding: utf-8 -*-
"""
Load Family

Load Family from Dropbox

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ ="Tran Tien Thanh"
__title__ = "Load Family"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from rpw import *
from pyrevit import *
from pyrevit import forms, script
import math
from System.Windows import Window
from System.Windows.Controls import TextBlock, Button, StackPanel, Border, Image
from System.Windows.Media import Brushes
from System.Windows.Media.Imaging import BitmapImage
from System import Uri

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# CLASS/FUNCTIONS
# ==================================================
class LoadFamilyWindow(forms.WPFWindow):
    def __init__(self):
        try:
            xaml_file = script.get_bundle_file('ui.xaml')
            super(LoadFamilyWindow, self).__init__(xaml_file)
        except Exception as e:
    forms.alert(f"Lỗi khi load XAML: {e}")

        
        # Load dữ liệu ban đầu
        self.populate_families()

    def populate_families(self):
        """Tạo danh sách các Family dưới dạng Button có hình ảnh xem trước."""
        families = [
            {"name": "Base Cabinet - 2 Bin", "image": "icon.png"},
            {"name": "Base Cabinet - 3 Drawers", "image": "icon.png"},
            {"name": "Base Cabinet - 4 Drawers", "image": "icon.png"}
        ]
        
        for family in families:
            item = Border()
            item.BorderBrush = Brushes.Gray
            item.BorderThickness = DB.Thickness(1)
            item.Margin = DB.Thickness(5)
            item.Width = 120
            item.Height = 150
            
            stack = StackPanel()
            
            img = Image()
            img.Source = BitmapImage(Uri(family["image"], UriKind.Relative))
            img.Height = 80
            
            text = TextBlock()
            text.Text = family["name"]
            text.TextAlignment = DB.TextAlignment.Center
            
            stack.Children.Add(img)
            stack.Children.Add(text)
            
            item.Child = stack
            self.familyList.Children.Add(item)

    def OnLoadClick(self, sender, args):
        """Hàm xử lý khi bấm Load."""
        forms.alert("Family Loaded!")

    def OnCancelClick(self, sender, args):
        """Đóng cửa sổ khi bấm Cancel."""
        self.Close()
    def on_search_changed(self, sender, args):
        search_text = self.searchBox.Text
        forms.alert(f"Searching: {search_text}")
        
# MAIN SCRIPT
# ==================================================
#get elements

#run
if __name__ == "__main__":
    win = LoadFamilyWindow()
    win.show()