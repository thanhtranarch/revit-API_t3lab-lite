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
import json
import traceback

# .NET Imports
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System
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

# Config file path
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".t3lab")
CONFIG_FILE = os.path.join(CONFIG_DIR, "family_loader_config.json")

# ╦ ╦╔═╗╦  ╔═╗╔═╗╦═╗  ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠═╣║╣ ║  ╠═╝║╣ ╠╦╝  ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╩ ╩╚═╝╩═╝╩  ╚═╝╩╚═  ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝
#====================================================================================================

def load_config():
    """Load configuration from JSON file"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                logger.info("Loaded config from: {}".format(CONFIG_FILE))
                return config
        else:
            logger.info("No config file found at: {}".format(CONFIG_FILE))
            return {}
    except Exception as ex:
        logger.error("Failed to load config: {}".format(ex))
        logger.error(traceback.format_exc())
        return {}

def save_config(config):
    """Save configuration to JSON file"""
    try:
        # Create config directory if it doesn't exist
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)
            logger.info("Created config directory: {}".format(CONFIG_DIR))

        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=4)
        logger.info("Saved config to: {}".format(CONFIG_FILE))
        return True
    except Exception as ex:
        logger.error("Failed to save config: {}".format(ex))
        logger.error(traceback.format_exc())
        return False

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
        except Exception as ex:
            # Silently ignore thumbnail loading errors
            logger.debug("Failed to load thumbnail {}: {}".format(thumbnail_path, ex))
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
                try:
                    self.parent_window.update_result_count()
                except Exception as ex:
                    # Silently ignore update errors to prevent cascade failures
                    logger.debug("Error updating count from IsChecked: {}".format(ex))

    def OnPropertyChanged(self, propertyName):
        if self.PropertyChanged is not None:
            self.PropertyChanged(self, PropertyChangedEventArgs(propertyName))

    # INotifyPropertyChanged implementation
    PropertyChanged = None


class FamilyLoaderWindow(Window):
    """Main window for Family Loader"""

    def __init__(self):
        try:
            logger.info("Initializing FamilyLoaderWindow...")

            # Initialize the base Window class first
            Window.__init__(self)

            # Set Window properties
            self.Title = "Load Autodesk Family"
            self.Height = 700
            self.Width = 1000
            self.MinHeight = 500
            self.MinWidth = 800
            self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

            # Load XAML
            xaml_path = os.path.join(os.path.dirname(__file__), 'FamilyLoader.xaml')
            logger.info("Loading XAML from: {}".format(xaml_path))

            try:
                with open(xaml_path, 'r') as f:
                    xaml_content = f.read()
                self.ui = XamlReader.Parse(xaml_content)
                # Set the parsed content
                self.Content = self.ui
                logger.info("XAML loaded successfully")
            except Exception as e:
                logger.error("Error loading XAML: {}".format(str(e)))
                logger.error(traceback.format_exc())
                forms.alert("Error loading XAML: {}".format(str(e)))
                raise

            # Get named controls
            logger.info("Getting named controls from XAML...")
            self.btn_select_folder = self.ui.FindName('btn_select_folder')
            self.txt_current_folder = self.ui.FindName('txt_current_folder')
            self.txt_search = self.ui.FindName('txt_search')
            self.tree_categories = self.ui.FindName('tree_categories')
            self.items_families = self.ui.FindName('items_families')
            self.txt_result_count = self.ui.FindName('txt_result_count')
            self.txt_selected_count = self.ui.FindName('txt_selected_count')
            self.btn_select_all = self.ui.FindName('btn_select_all')
            self.btn_select_none = self.ui.FindName('btn_select_none')
            self.btn_load = self.ui.FindName('btn_load')
            self.btn_cancel = self.ui.FindName('btn_cancel')

            # Wire up event handlers
            logger.info("Wiring up event handlers...")
            self.btn_select_folder.Click += self.select_folder_clicked
            self.txt_search.TextChanged += self.search_text_changed
            self.tree_categories.SelectedItemChanged += self.category_selected
            self.btn_select_all.Click += self.select_all_clicked
            self.btn_select_none.Click += self.select_none_clicked
            self.btn_load.Click += self.load_clicked
            self.btn_cancel.Click += self.cancel_clicked
            self.Loaded += self.window_loaded

            # Initialize variables
            logger.info("Initializing variables...")
            self.config = load_config()
            self.current_folder = self.config.get('last_folder', None)
            self.all_families = []
            self.filtered_families = ObservableCollection[object]()
            self.category_structure = {}
            self._is_updating = False  # Flag to prevent UI updates during batch operations

            # Bind to ItemsControl
            self.items_families.ItemsSource = self.filtered_families

            # Result
            self.loaded_families = []

            logger.info("Family Loader window initialized successfully")
            if self.current_folder:
                logger.info("Last used folder: {}".format(self.current_folder))

        except Exception as ex:
            logger.error("Critical error in FamilyLoaderWindow.__init__: {}".format(ex))
            logger.error(traceback.format_exc())
            raise

    # ╔═╗╦  ╦╔═╗╔╗╔╔╦╗  ╦ ╦╔═╗╔╗╔╔╦╗╦  ╔═╗╦═╗╔═╗
    # ║╣ ╚╗╔╝║╣ ║║║ ║   ╠═╣╠═╣║║║ ║║║  ║╣ ╠╦╝╚═╗
    # ╚═╝ ╚╝ ╚═╝╝╚╝ ╩   ╩ ╩╩ ╩╝╚╝═╩╝╩═╝╚═╝╩╚═╚═╝
    #====================================================================================================

    def window_loaded(self, sender, e):
        """Handle window loaded event - auto-show folder dialog if no folder is set"""
        try:
            logger.info("Window loaded event triggered")

            # Check if all UI elements are ready
            if not hasattr(self, 'txt_current_folder') or not self.txt_current_folder:
                logger.error("UI elements not ready in window_loaded")
                return

            if self.current_folder:
                # Check if saved folder still exists
                if os.path.exists(self.current_folder):
                    logger.info("Loading families from saved folder: {}".format(self.current_folder))
                    self.txt_current_folder.Text = self.current_folder
                    self.scan_families()
                else:
                    logger.warning("Saved folder no longer exists: {}".format(self.current_folder))
                    self.current_folder = None
                    self.select_folder_clicked(None, None)
            else:
                # No saved folder - show dialog
                logger.info("No saved folder found, showing folder selection dialog")
                self.select_folder_clicked(None, None)
        except Exception as ex:
            logger.error("Error in window_loaded: {}".format(ex))
            logger.error(traceback.format_exc())

    def select_folder_clicked(self, sender, e):
        """Handle folder selection"""
        try:
            dialog = FolderBrowserDialog()
            dialog.Description = "Select folder containing Revit families"

            if dialog.ShowDialog() == DialogResult.OK:
                self.current_folder = dialog.SelectedPath
                self.txt_current_folder.Text = self.current_folder
                logger.info("User selected folder: {}".format(self.current_folder))

                # Save folder to config
                self.config['last_folder'] = self.current_folder
                if save_config(self.config):
                    logger.info("Folder path saved to config")

                self.scan_families()
        except Exception as ex:
            logger.error("Error in select_folder_clicked: {}".format(ex))
            logger.error(traceback.format_exc())
            forms.alert("Error selecting folder: {}".format(ex), exitscript=False)

    def scan_families(self):
        """Scan selected folder for .rfa files"""
        if not self.current_folder:
            logger.warning("No current folder set for scanning")
            return

        logger.info("Starting to scan folder: {}".format(self.current_folder))

        # Set updating flag to prevent UI updates during scan
        self._is_updating = True

        self.all_families = []
        self.category_structure = {}
        scan_errors = 0

        try:
            # Walk through directory
            for root, dirs, files in os.walk(self.current_folder):
                for file in files:
                    if file.lower().endswith('.rfa'):
                        try:
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
                        except Exception as item_ex:
                            scan_errors += 1
                            logger.warning("Failed to process family {}: {}".format(file, item_ex))
                            # Continue scanning other families

            # Re-enable UI updates
            self._is_updating = False

            # Update UI
            self.update_category_tree()
            self.update_family_display()

            logger.info("Scan completed: {} families found in {} categories".format(
                len(self.all_families),
                len(self.category_structure)
            ))

            if scan_errors > 0:
                logger.warning("Encountered {} errors while scanning families".format(scan_errors))

        except Exception as ex:
            # Re-enable UI updates even on error
            self._is_updating = False
            logger.error("Critical error scanning families: {}".format(ex))
            logger.error(traceback.format_exc())
            forms.alert("Error scanning folder: {}".format(ex), exitscript=False)

    def update_category_tree(self):
        """Update the category tree view with hierarchical structure"""
        try:
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
            logger.debug("Category tree updated with {} categories".format(len(self.category_structure)))
        except Exception as ex:
            logger.error("Error updating category tree: {}".format(ex))
            logger.error(traceback.format_exc())

    def _count_families_in_tree(self, tree_node):
        """Count all families in a tree node and its children"""
        count = len(tree_node.get('_families', []))
        for child in tree_node.get('_children', {}).values():
            count += self._count_families_in_tree(child)
        return count

    def update_family_display(self, families=None):
        """Update the family display grid"""
        try:
            if families is None:
                families = self.all_families

            self.filtered_families.Clear()
            for family in families:
                self.filtered_families.Add(family)

            self.update_result_count()
            logger.debug("Family display updated with {} families".format(len(families)))
        except Exception as ex:
            logger.error("Error updating family display: {}".format(ex))
            logger.error(traceback.format_exc())

    def update_result_count(self):
        """Update the result count text"""
        # Skip updates during batch operations
        if self._is_updating:
            return

        try:
            count = len(self.filtered_families)
            self.txt_result_count.Text = "{} families found".format(count)

            # Update selected count
            selected = sum(1 for f in self.filtered_families if f.IsChecked)
            self.txt_selected_count.Text = "{} families selected".format(selected)

            # Enable/disable load button
            self.btn_load.IsEnabled = selected > 0
        except Exception as ex:
            logger.error("Error updating result count: {}".format(ex))
            logger.error(traceback.format_exc())

    def category_selected(self, sender, e):
        """Handle category selection"""
        try:
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
                logger.debug("Category selected: {} ({} families)".format(tag, len(filtered)))
        except Exception as ex:
            logger.error("Error in category_selected: {}".format(ex))
            logger.error(traceback.format_exc())

    def search_text_changed(self, sender, e):
        """Handle search text changes"""
        try:
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
                logger.debug("Search: '{}' found {} families".format(search_text, len(filtered)))
        except Exception as ex:
            logger.error("Error in search_text_changed: {}".format(ex))
            logger.error(traceback.format_exc())

    def select_all_clicked(self, sender, e):
        """Select all families"""
        try:
            for family in self.filtered_families:
                family.IsChecked = True
            self.update_result_count()
            logger.debug("Selected all {} families".format(len(self.filtered_families)))
        except Exception as ex:
            logger.error("Error in select_all_clicked: {}".format(ex))
            logger.error(traceback.format_exc())

    def select_none_clicked(self, sender, e):
        """Deselect all families"""
        try:
            for family in self.filtered_families:
                family.IsChecked = False
            self.update_result_count()
            logger.debug("Deselected all families")
        except Exception as ex:
            logger.error("Error in select_none_clicked: {}".format(ex))
            logger.error(traceback.format_exc())

    def load_clicked(self, sender, e):
        """Load selected families into Revit"""
        try:
            selected_families = [f for f in self.all_families if f.IsChecked]

            if not selected_families:
                forms.alert("Please select at least one family to load.", exitscript=False)
                return

            logger.info("Starting to load {} selected families".format(len(selected_families)))

            # Load families
            success_count = 0
            fail_count = 0
            failed_families = []

            with revit.Transaction("Load Families"):
                for family in selected_families:
                    try:
                        logger.debug("Attempting to load: {}".format(family.FullPath))

                        # Check if file exists before attempting to load
                        if not os.path.exists(family.FullPath):
                            logger.error("Family file not found: {}".format(family.FullPath))
                            fail_count += 1
                            failed_families.append(family.Name)
                            continue

                        # Load family
                        loaded = doc.LoadFamily(family.FullPath)
                        if loaded:
                            success_count += 1
                            self.loaded_families.append(family.FullPath)
                            logger.info("Successfully loaded: {}".format(family.Name))
                        else:
                            fail_count += 1
                            failed_families.append(family.Name)
                            logger.warning("LoadFamily returned False for: {}".format(family.Name))
                    except Exception as ex:
                        logger.error("Failed to load {}: {}".format(family.Name, ex))
                        logger.error(traceback.format_exc())
                        fail_count += 1
                        failed_families.append(family.Name)

            # Show result
            message = "Loaded {} families successfully.".format(success_count)
            if fail_count > 0:
                message += "\n{} families failed to load.".format(fail_count)
                if len(failed_families) <= 5:
                    message += "\nFailed families: {}".format(", ".join(failed_families))

            logger.info("Load completed: {} success, {} failed".format(success_count, fail_count))
            forms.alert(message, exitscript=False)

            # Close dialog
            self.DialogResult = True
            self.Close()

        except Exception as ex:
            logger.error("Critical error in load_clicked: {}".format(ex))
            logger.error(traceback.format_exc())
            forms.alert("Error loading families: {}".format(ex), exitscript=False)

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
        logger.info("=" * 80)
        logger.info("Family Loader Dialog Starting")
        logger.info("=" * 80)

        window = FamilyLoaderWindow()
        window.ShowDialog()

        logger.info("Family Loader Dialog Closed")
        logger.info("Loaded {} families".format(len(window.loaded_families)))
        logger.info("=" * 80)

        return window.loaded_families
    except Exception as ex:
        logger.error("Critical error showing family loader: {}".format(ex))
        logger.error(traceback.format_exc())
        forms.alert("Error: {}".format(ex))
        return []


if __name__ == '__main__':
    show_family_loader()
