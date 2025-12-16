# -*- coding: utf-8 -*-
"""
Family Loader Dialog
Load Revit families from folders with category organization
"""
__title__ = "Family Loader"
__author__ = "T3Lab"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
import os
import sys
import clr

# .NET Imports
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Uri
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows import Window, Visibility
from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapImage
from System.Windows.Controls import TreeViewItem
from System.Windows.Forms import FolderBrowserDialog, DialogResult

# pyRevit Imports
from pyrevit import revit, DB, forms, script

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#====================================================================================================
logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
#====================================================================================================

class FamilyItem(INotifyPropertyChanged):
    """Represents a family file with its properties"""

    def __init__(self, name, full_path, category, thumbnail_path=None, parent_window=None):
        self._is_checked = False
        self.Name = name
        self.FullPath = full_path
        self.Category = category
        self.Thumbnail = self._load_thumbnail(thumbnail_path)
        self.parent_window = parent_window

    def _load_thumbnail(self, thumbnail_path):
        """Load thumbnail image or return default"""
        try:
            if thumbnail_path and os.path.exists(thumbnail_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(thumbnail_path)
                bitmap.DecodePixelWidth = 90
                bitmap.EndInit()
                return bitmap
        except:
            pass
        return None

    @property
    def IsChecked(self):
        return self._is_checked

    @IsChecked.setter
    def IsChecked(self, value):
        if self._is_checked != value:
            self._is_checked = value
            self.OnPropertyChanged("IsChecked")
            # Notify parent window to update count
            if self.parent_window:
                self.parent_window.update_result_count()

    def OnPropertyChanged(self, propertyName):
        if self.PropertyChanged is not None:
            self.PropertyChanged(self, PropertyChangedEventArgs(propertyName))

    # INotifyPropertyChanged implementation
    PropertyChanged = None


class FamilyLoaderWindow(Window):
    """Main window for Family Loader"""

    def __init__(self):
        # Initialize the base Window class first
        Window.__init__(self)

        # Load XAML
        xaml_path = os.path.join(os.path.dirname(__file__), 'FamilyLoader.xaml')
        try:
            with open(xaml_path, 'r') as f:
                xaml_content = f.read()
            self.ui = XamlReader.Parse(xaml_content)
            # Set the parsed content
            self.Content = self.ui
        except Exception as e:
            forms.alert("Error loading XAML: {}".format(str(e)))
            return

        # Get named controls
        self.txt_current_folder = self.ui.FindName('txt_current_folder')
        self.txt_search = self.ui.FindName('txt_search')
        self.tree_categories = self.ui.FindName('tree_categories')
        self.items_families = self.ui.FindName('items_families')
        self.txt_result_count = self.ui.FindName('txt_result_count')
        self.txt_selected_count = self.ui.FindName('txt_selected_count')
        self.btn_load = self.ui.FindName('btn_load')

        # Initialize variables
        self.current_folder = None
        self.all_families = []
        self.filtered_families = ObservableCollection[object]()
        self.category_structure = {}

        # Bind to ItemsControl
        self.items_families.ItemsSource = self.filtered_families

        # Result
        self.loaded_families = []

    # ╔═╗╦  ╦╔═╗╔╗╔╔╦╗  ╦ ╦╔═╗╔╗╔╔╦╗╦  ╔═╗╦═╗╔═╗
    # ║╣ ╚╗╔╝║╣ ║║║ ║   ╠═╣╠═╣║║║ ║║║  ║╣ ╠╦╝╚═╗
    # ╚═╝ ╚╝ ╚═╝╝╚╝ ╩   ╩ ╩╩ ╩╝╚╝═╩╝╩═╝╚═╝╩╚═╚═╝
    #====================================================================================================

    def select_folder_clicked(self, sender, e):
        """Handle folder selection"""
        dialog = FolderBrowserDialog()
        dialog.Description = "Select folder containing Revit families"

        if dialog.ShowDialog() == DialogResult.OK:
            self.current_folder = dialog.SelectedPath
            self.txt_current_folder.Text = self.current_folder
            self.scan_families()

    def scan_families(self):
        """Scan selected folder for .rfa files"""
        if not self.current_folder:
            return

        self.all_families = []
        self.category_structure = {}

        try:
            # Walk through directory
            for root, dirs, files in os.walk(self.current_folder):
                for file in files:
                    if file.lower().endswith('.rfa'):
                        full_path = os.path.join(root, file)
                        relative_path = os.path.relpath(root, self.current_folder)

                        # Use folder name as category
                        category = relative_path if relative_path != '.' else 'Root'

                        # Create family item
                        family_name = os.path.splitext(file)[0]
                        family_item = FamilyItem(family_name, full_path, category, parent_window=self)
                        self.all_families.append(family_item)

                        # Add to category structure
                        if category not in self.category_structure:
                            self.category_structure[category] = []
                        self.category_structure[category].append(family_item)

            # Update UI
            self.update_category_tree()
            self.update_family_display()

            logger.info("Found {} families in {} categories".format(
                len(self.all_families),
                len(self.category_structure)
            ))

        except Exception as ex:
            logger.error("Error scanning families: {}".format(ex))
            forms.alert("Error scanning folder: {}".format(ex), exitscript=False)

    def update_category_tree(self):
        """Update the category tree view with hierarchical structure"""
        self.tree_categories.Items.Clear()

        # Add "All" item
        all_item = TreeViewItem()
        all_item.Header = "All ({})".format(len(self.all_families))
        all_item.Tag = "ALL"
        all_item.IsExpanded = True
        self.tree_categories.Items.Add(all_item)

        # Build hierarchical tree structure
        tree_dict = {}

        for category, families in self.category_structure.items():
            # Split category path
            if category == 'Root':
                parts = ['Root']
            else:
                parts = category.split(os.sep)

            # Build nested structure
            current_dict = tree_dict
            for i, part in enumerate(parts):
                if part not in current_dict:
                    current_dict[part] = {'_families': [], '_children': {}}
                current_dict = current_dict[part]['_children']

            # Add families to the leaf
            path_key = os.sep.join(parts) if parts != ['Root'] else 'Root'
            if path_key in self.category_structure:
                tree_dict_leaf = tree_dict
                for part in parts:
                    tree_dict_leaf = tree_dict_leaf[part]
                tree_dict_leaf['_families'] = self.category_structure[path_key]

        # Recursively add tree items
        def add_tree_items(parent_item, tree_data, path_prefix=""):
            for folder_name, data in sorted(tree_data.items()):
                folder_path = os.path.join(path_prefix, folder_name) if path_prefix else folder_name

                # Count all families in this folder and subfolders
                total_families = self._count_families_in_tree(data)

                # Create tree item
                item = TreeViewItem()
                item.Header = "{} ({})".format(folder_name, total_families)
                item.Tag = folder_path if folder_path != 'Root' else 'Root'
                item.IsExpanded = True
                parent_item.Items.Add(item)

                # Add children recursively
                if data['_children']:
                    add_tree_items(item, data['_children'], folder_path)

        add_tree_items(self.tree_categories, tree_dict)

    def _count_families_in_tree(self, tree_node):
        """Count all families in a tree node and its children"""
        count = len(tree_node.get('_families', []))
        for child in tree_node.get('_children', {}).values():
            count += self._count_families_in_tree(child)
        return count

    def update_family_display(self, families=None):
        """Update the family display grid"""
        if families is None:
            families = self.all_families

        self.filtered_families.Clear()
        for family in families:
            self.filtered_families.Add(family)

        self.update_result_count()

    def update_result_count(self):
        """Update the result count text"""
        count = len(self.filtered_families)
        self.txt_result_count.Text = "{} families found".format(count)

        # Update selected count
        selected = sum(1 for f in self.filtered_families if f.IsChecked)
        self.txt_selected_count.Text = "{} families selected".format(selected)

        # Enable/disable load button
        self.btn_load.IsEnabled = selected > 0

    def category_selected(self, sender, e):
        """Handle category selection"""
        selected_item = self.tree_categories.SelectedItem
        if not selected_item:
            return

        tag = selected_item.Tag

        if tag == "ALL":
            self.update_family_display(self.all_families)
        else:
            # Show families in selected folder and all subfolders
            filtered = [f for f in self.all_families
                       if f.Category == tag or f.Category.startswith(tag + os.sep)]
            self.update_family_display(filtered)

    def search_text_changed(self, sender, e):
        """Handle search text changes"""
        search_text = self.txt_search.Text.lower()

        if not search_text:
            # Get current category selection
            selected_item = self.tree_categories.SelectedItem
            if selected_item and selected_item.Tag != "ALL":
                filtered = [f for f in self.all_families if f.Category == selected_item.Tag]
                self.update_family_display(filtered)
            else:
                self.update_family_display(self.all_families)
        else:
            # Filter by search text
            filtered = [f for f in self.all_families
                       if search_text in f.Name.lower() or
                          search_text in f.Category.lower()]
            self.update_family_display(filtered)

    def select_all_clicked(self, sender, e):
        """Select all families"""
        for family in self.filtered_families:
            family.IsChecked = True
        self.update_result_count()

    def select_none_clicked(self, sender, e):
        """Deselect all families"""
        for family in self.filtered_families:
            family.IsChecked = False
        self.update_result_count()

    def load_clicked(self, sender, e):
        """Load selected families into Revit"""
        selected_families = [f for f in self.all_families if f.IsChecked]

        if not selected_families:
            forms.alert("Please select at least one family to load.", exitscript=False)
            return

        # Load families
        success_count = 0
        fail_count = 0

        with revit.Transaction("Load Families"):
            for family in selected_families:
                try:
                    # Load family
                    loaded = doc.LoadFamily(family.FullPath)
                    if loaded:
                        success_count += 1
                        self.loaded_families.append(family.FullPath)
                    else:
                        fail_count += 1
                except Exception as ex:
                    logger.error("Failed to load {}: {}".format(family.Name, ex))
                    fail_count += 1

        # Show result
        message = "Loaded {} families successfully.".format(success_count)
        if fail_count > 0:
            message += "\n{} families failed to load.".format(fail_count)

        forms.alert(message, exitscript=False)

        # Close dialog
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, e):
        """Cancel and close dialog"""
        self.Close()


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
#====================================================================================================

def show_family_loader():
    """Show the family loader dialog"""
    try:
        window = FamilyLoaderWindow()
        window.ShowDialog()
        return window.loaded_families
    except Exception as ex:
        logger.error("Error showing family loader: {}".format(ex))
        forms.alert("Error: {}".format(ex))
        return []


if __name__ == '__main__':
    show_family_loader()
