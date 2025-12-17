# -*- coding: utf-8 -*-
"""
Family Loader Dialog
Load Revit families from folders with category organization
COMPREHENSIVE FIX: Memory leaks, threading, error handling, and stability improvements
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
import time
import datetime
import threading

# .NET Imports
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System
from System import Uri, Action
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows import Window, Visibility
from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapImage
from System.Windows.Controls import TreeViewItem
from System.Windows.Forms import FolderBrowserDialog, DialogResult
from System.Windows.Threading import Dispatcher

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
    """Load configuration from JSON file with validation and corruption recovery"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)

                # Validate config structure
                if not isinstance(config, dict):
                    raise ValueError("Invalid config format: expected dict, got {}".format(type(config)))

                logger.info("Loaded config from: {}".format(CONFIG_FILE))
                return config
        else:
            logger.info("No config file found at: {}".format(CONFIG_FILE))
            return {}
    except (ValueError, json.JSONDecodeError) as ex:
        logger.error("Config file corrupted, recreating: {}".format(ex))
        # Backup corrupted file
        try:
            backup_path = CONFIG_FILE + ".corrupted.{}".format(int(time.time()))
            if os.path.exists(CONFIG_FILE):
                os.rename(CONFIG_FILE, backup_path)
                logger.info("Backed up corrupted config to: {}".format(backup_path))
        except Exception as backup_ex:
            logger.error("Failed to backup corrupted config: {}".format(backup_ex))
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

def is_valid_rfa_file(file_path):
    """Validate if file is a valid Revit family file"""
    try:
        # Check file exists
        if not os.path.exists(file_path):
            return False

        # Check file size (min 1KB, max 500MB)
        size = os.path.getsize(file_path)
        if size < 1024 or size > 500 * 1024 * 1024:
            logger.debug("File size out of range: {} bytes for {}".format(size, file_path))
            return False

        # Try to open file to check if accessible
        with open(file_path, 'rb') as f:
            # Read header to validate (Revit files are OLE/Structured storage)
            header = f.read(8)
            if not header.startswith(b'\xD0\xCF\x11\xE0'):
                logger.debug("Invalid file header for: {}".format(file_path))
                return False

        return True
    except Exception as ex:
        logger.debug("File validation failed for {}: {}".format(file_path, ex))
        return False

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
#====================================================================================================

class FamilyLoadOptions(DB.IFamilyLoadOptions):
    """Custom IFamilyLoadOptions to handle family conflicts automatically"""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        """Handle when family already exists in project"""
        # Always overwrite existing families
        overwriteParameterValues = True
        logger.debug("Family found in project, overwriting: {}".format(
            familyInUse.Name if hasattr(familyInUse, 'Name') else 'Unknown'
        ))
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        """Handle when shared family is found"""
        # Always use the source (file) version
        overwriteParameterValues = True
        source = DB.FamilySource.Family
        logger.debug("Shared family found, using source version")
        return True

class FamilyItem(INotifyPropertyChanged):
    """Represents a family file with its properties"""

    def __init__(self, name, full_path, category, thumbnail_path=None):
        self._is_checked = False
        self._is_disposed = False
        self.Name = name
        self.FullPath = full_path
        self.Category = category
        self.Thumbnail = self._load_thumbnail(thumbnail_path)

    def _load_thumbnail(self, thumbnail_path):
        """Load thumbnail image or return default"""
        try:
            if thumbnail_path and os.path.exists(thumbnail_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(thumbnail_path)
                bitmap.DecodePixelWidth = 90
                bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
                bitmap.EndInit()
                bitmap.Freeze()  # Make bitmap immutable for thread safety and memory optimization
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

    def OnPropertyChanged(self, propertyName):
        try:
            if not self._is_disposed and self.PropertyChanged is not None:
                self.PropertyChanged(self, PropertyChangedEventArgs(propertyName))
        except Exception as ex:
            # Silently ignore PropertyChanged errors
            logger.debug("Error in OnPropertyChanged: {}".format(ex))

    def Dispose(self):
        """Clean up resources to prevent memory leaks"""
        try:
            if not self._is_disposed:
                # Clear thumbnail reference
                self.Thumbnail = None
                # Clear event handlers
                self.PropertyChanged = None
                self._is_disposed = True
        except Exception as ex:
            logger.debug("Error disposing FamilyItem: {}".format(ex))

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
            self._cancel_requested = False  # Flag for cancellation
            self._scan_thread = None  # Background scan thread
            self._seen_family_names = {}  # Track duplicate family names

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
        """Scan selected folder for .rfa files with background threading"""
        if not self.current_folder:
            logger.warning("No current folder set for scanning")
            return

        # Reset cancellation flag
        self._cancel_requested = False

        # Disable UI controls during scan
        self.btn_select_folder.IsEnabled = False
        self.btn_load.IsEnabled = False
        self.txt_current_folder.Text = "{} (Scanning...)".format(self.current_folder)

        logger.info("=" * 80)
        logger.info("FAMILY SCAN STARTED: {}".format(datetime.datetime.now()))
        logger.info("Folder: {}".format(self.current_folder))
        logger.info("=" * 80)

        # Start background scan thread
        self._scan_thread = threading.Thread(target=self._scan_families_worker)
        self._scan_thread.daemon = True
        self._scan_thread.start()

    def _scan_families_worker(self):
        """Background worker for scanning families"""
        start_time = time.time()
        scan_errors = 0
        permission_errors = 0
        validation_errors = 0
        timeout_seconds = 300  # 5 minutes timeout

        temp_families = []
        temp_category_structure = {}
        temp_seen_names = {}

        try:
            logger.info("Walking through directory structure...")

            # Walk through directory with error handling
            for root, dirs, files in os.walk(self.current_folder, followlinks=False):
                # Check for cancellation
                if self._cancel_requested:
                    logger.info("Scan cancelled by user")
                    self._scan_complete([], {}, cancelled=True)
                    return

                # Check for timeout
                if time.time() - start_time > timeout_seconds:
                    logger.error("Scan timeout after {} seconds".format(timeout_seconds))
                    self._scan_complete([], {}, timeout=True)
                    return

                # Test directory accessibility
                try:
                    _ = os.listdir(root)
                except (PermissionError, OSError) as access_ex:
                    logger.warning("Skipping inaccessible folder {}: {}".format(root, access_ex))
                    permission_errors += 1
                    dirs[:] = []  # Don't descend into this directory
                    continue

                # Process files
                for file in files:
                    # Check cancellation
                    if self._cancel_requested:
                        logger.info("Scan cancelled by user")
                        self._scan_complete([], {}, cancelled=True)
                        return

                    if file.lower().endswith('.rfa'):
                        try:
                            full_path = os.path.join(root, file)
                            relative_path = os.path.relpath(root, self.current_folder)

                            # Validate file
                            if not is_valid_rfa_file(full_path):
                                logger.debug("Skipping invalid .rfa file: {}".format(full_path))
                                validation_errors += 1
                                continue

                            # Use folder name as category
                            category = relative_path if relative_path != '.' else 'Root'

                            # Create family name with duplicate detection
                            family_name = os.path.splitext(file)[0]
                            if family_name in temp_seen_names:
                                # Append folder name to disambiguate
                                logger.warning("Duplicate family name: {} in {} and {}".format(
                                    family_name,
                                    temp_seen_names[family_name],
                                    full_path
                                ))
                                folder_name = os.path.basename(root)
                                family_name = "{} ({})".format(family_name, folder_name)
                            else:
                                temp_seen_names[family_name] = full_path

                            # Create family item
                            family_item = FamilyItem(family_name, full_path, category)
                            temp_families.append(family_item)

                            # Add to category structure
                            if category not in temp_category_structure:
                                temp_category_structure[category] = []
                            temp_category_structure[category].append(family_item)

                            # Update progress every 50 families
                            if len(temp_families) % 50 == 0:
                                logger.info("Scanned {} families so far...".format(len(temp_families)))
                                # Update UI progress on main thread
                                self._update_scan_progress(len(temp_families))

                        except Exception as item_ex:
                            scan_errors += 1
                            logger.warning("Failed to process family {}: {}".format(file, item_ex))
                            logger.debug(traceback.format_exc())
                            # Continue scanning other families

            # Calculate duration
            duration = time.time() - start_time

            logger.info("Directory walk completed in {:.2f} seconds".format(duration))
            logger.info("Found {} families in {} categories".format(
                len(temp_families),
                len(temp_category_structure)
            ))

            if scan_errors > 0:
                logger.warning("Encountered {} file processing errors".format(scan_errors))
            if permission_errors > 0:
                logger.warning("Encountered {} permission errors (folders skipped)".format(permission_errors))
            if validation_errors > 0:
                logger.warning("Skipped {} invalid .rfa files".format(validation_errors))

            # Complete scan on UI thread
            self._scan_complete(temp_families, temp_category_structure)

        except Exception as ex:
            logger.error("Critical error in scan worker: {}".format(ex))
            logger.error(traceback.format_exc())
            self._scan_complete([], {}, error=str(ex))

    def _update_scan_progress(self, count):
        """Update scan progress on UI thread"""
        try:
            if self.Dispatcher:
                self.Dispatcher.Invoke(
                    Action(lambda: self._update_scan_progress_ui(count))
                )
        except Exception as ex:
            logger.debug("Error updating progress: {}".format(ex))

    def _update_scan_progress_ui(self, count):
        """Update progress UI (called on UI thread)"""
        try:
            self.txt_current_folder.Text = "{} (Scanning... {} families found)".format(
                self.current_folder, count
            )
        except Exception as ex:
            logger.debug("Error updating progress UI: {}".format(ex))

    def _scan_complete(self, families, category_structure, error=None, cancelled=False, timeout=False):
        """Handle scan completion on UI thread"""
        try:
            if self.Dispatcher:
                self.Dispatcher.Invoke(
                    Action(lambda: self._scan_complete_ui(families, category_structure, error, cancelled, timeout))
                )
        except Exception as ex:
            logger.error("Error invoking scan complete: {}".format(ex))

    def _scan_complete_ui(self, families, category_structure, error=None, cancelled=False, timeout=False):
        """Complete scan and update UI (called on UI thread)"""
        try:
            # Clean up old families to prevent memory leaks
            for old_family in self.all_families:
                if hasattr(old_family, 'Dispose'):
                    old_family.Dispose()

            # Update data
            self.all_families = families
            self.category_structure = category_structure

            # Re-enable UI
            self.btn_select_folder.IsEnabled = True
            self.txt_current_folder.Text = self.current_folder

            # Handle different completion states
            if error:
                logger.error("Scan failed with error: {}".format(error))
                forms.alert("Error scanning folder: {}".format(error), exitscript=False)
            elif cancelled:
                logger.info("Scan cancelled by user")
                forms.alert("Scan cancelled", exitscript=False)
            elif timeout:
                logger.error("Scan timeout")
                forms.alert("Scan timeout: Operation took too long (>5 minutes)", exitscript=False)
            else:
                # Update UI with results
                logger.info("Updating category tree...")
                self.update_category_tree()

                logger.info("Updating family display...")
                self.update_family_display()

                logger.info("=" * 80)
                logger.info("FAMILY SCAN COMPLETED: {}".format(datetime.datetime.now()))
                logger.info("Total families: {}".format(len(self.all_families)))
                logger.info("Total categories: {}".format(len(self.category_structure)))
                logger.info("=" * 80)

        except Exception as ex:
            logger.error("Critical error in scan complete UI: {}".format(ex))
            logger.error(traceback.format_exc())
            self.btn_select_folder.IsEnabled = True
            self.txt_current_folder.Text = self.current_folder
            forms.alert("Error completing scan: {}".format(ex), exitscript=False)

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
        """Update the family display grid with proper event cleanup"""
        try:
            if families is None:
                families = self.all_families

            # Unsubscribe old events to prevent memory leaks
            for old_family in self.filtered_families:
                try:
                    old_family.PropertyChanged -= self.on_family_property_changed
                except:
                    pass  # Ignore if not subscribed

            # Clear collection
            self.filtered_families.Clear()

            # Add new families and subscribe events
            for family in families:
                # Subscribe to PropertyChanged event to update count when checkbox changes
                try:
                    family.PropertyChanged += self.on_family_property_changed
                except:
                    pass  # Ignore if already subscribed
                self.filtered_families.Add(family)

            self.update_result_count()
            logger.debug("Family display updated with {} families".format(len(families)))
        except Exception as ex:
            logger.error("Error updating family display: {}".format(ex))
            logger.error(traceback.format_exc())

    def on_family_property_changed(self, sender, e):
        """Handle property changed event from family items"""
        try:
            if e.PropertyName == "IsChecked" and not self._is_updating:
                self.update_result_count()
        except Exception as ex:
            logger.debug("Error in on_family_property_changed: {}".format(ex))

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
        """Load selected families into Revit with comprehensive error handling"""
        try:
            selected_families = [f for f in self.all_families if f.IsChecked]

            if not selected_families:
                forms.alert("Please select at least one family to load.", exitscript=False)
                return

            # Validate document state
            if doc.IsReadOnly:
                forms.alert("Cannot load families: Document is read-only.\nPlease open a modifiable document.", exitscript=False)
                return

            if doc.IsModifiable:
                forms.alert("Cannot load families: Document is currently being modified.\nPlease finish current operation first.", exitscript=False)
                return

            # Warn if document is workshared
            if doc.IsWorkshared and not doc.IsDetached:
                result = forms.alert(
                    "Document is workshared. Families will be loaded to central model.\n\nDo you want to continue?",
                    yes=True, no=True, exitscript=False
                )
                if not result:
                    logger.info("User cancelled loading due to workshared warning")
                    return

            logger.info("=" * 80)
            logger.info("FAMILY LOADING STARTED: {}".format(datetime.datetime.now()))
            logger.info("Selected families: {}".format(len(selected_families)))
            logger.info("=" * 80)

            start_time = time.time()

            # Disable UI during load
            self.btn_load.IsEnabled = False
            self.btn_cancel.IsEnabled = False

            # Load families with individual transactions
            success_count = 0
            fail_count = 0
            failed_families = []
            load_options = FamilyLoadOptions()

            for i, family in enumerate(selected_families):
                try:
                    logger.debug("[{}/{}] Attempting to load: {}".format(
                        i + 1, len(selected_families), family.FullPath
                    ))

                    # Check if file exists and is valid
                    if not os.path.exists(family.FullPath):
                        logger.error("Family file not found: {}".format(family.FullPath))
                        fail_count += 1
                        failed_families.append((family.Name, "File not found"))
                        continue

                    if not is_valid_rfa_file(family.FullPath):
                        logger.error("Invalid .rfa file: {}".format(family.FullPath))
                        fail_count += 1
                        failed_families.append((family.Name, "Invalid file format"))
                        continue

                    # Use individual transaction for each family
                    # This prevents one failure from rolling back all others
                    with revit.Transaction("Load Family: {}".format(family.Name)):
                        try:
                            # Load family with options to handle conflicts
                            loaded = doc.LoadFamily(family.FullPath, load_options)

                            if loaded:
                                success_count += 1
                                self.loaded_families.append(family.FullPath)
                                logger.info("[{}/{}] Successfully loaded: {}".format(
                                    i + 1, len(selected_families), family.Name
                                ))
                            else:
                                fail_count += 1
                                failed_families.append((family.Name, "LoadFamily returned False"))
                                logger.warning("[{}/{}] LoadFamily returned False for: {}".format(
                                    i + 1, len(selected_families), family.Name
                                ))

                        except DB.InvalidOperationException as inv_ex:
                            fail_count += 1
                            error_msg = "Invalid operation: {}".format(str(inv_ex))
                            failed_families.append((family.Name, error_msg))
                            logger.error("InvalidOperationException loading {}: {}".format(family.Name, inv_ex))

                        except DB.Exceptions.CorruptModelException as corrupt_ex:
                            fail_count += 1
                            error_msg = "Corrupt file"
                            failed_families.append((family.Name, error_msg))
                            logger.error("Corrupt family file {}: {}".format(family.Name, corrupt_ex))

                        except Exception as load_ex:
                            fail_count += 1
                            error_msg = str(load_ex)[:50]  # Truncate long errors
                            failed_families.append((family.Name, error_msg))
                            logger.error("Failed to load {}: {}".format(family.Name, load_ex))
                            logger.debug(traceback.format_exc())

                except Exception as outer_ex:
                    fail_count += 1
                    failed_families.append((family.Name, "Transaction error"))
                    logger.error("Transaction error for {}: {}".format(family.Name, outer_ex))
                    logger.error(traceback.format_exc())

            # Calculate duration
            duration = time.time() - start_time

            logger.info("=" * 80)
            logger.info("FAMILY LOADING COMPLETED: {}".format(datetime.datetime.now()))
            logger.info("Duration: {:.2f} seconds".format(duration))
            logger.info("Success: {}, Failed: {}".format(success_count, fail_count))
            logger.info("=" * 80)

            # Re-enable UI
            self.btn_load.IsEnabled = True
            self.btn_cancel.IsEnabled = True

            # Show result
            message = "Successfully loaded {} families in {:.1f} seconds.".format(success_count, duration)
            if fail_count > 0:
                message += "\n\n{} families failed to load.".format(fail_count)
                if len(failed_families) <= 10:
                    message += "\n\nFailed families:"
                    for fam_name, error in failed_families:
                        message += "\n- {}: {}".format(fam_name, error)
                else:
                    message += "\n\nShowing first 10 failures:"
                    for fam_name, error in failed_families[:10]:
                        message += "\n- {}: {}".format(fam_name, error)
                    message += "\n... and {} more (check log for details)".format(len(failed_families) - 10)

            forms.alert(message, exitscript=False)

            # Close dialog if any families were loaded successfully
            if success_count > 0:
                self.DialogResult = True
                self.Close()

        except Exception as ex:
            # Re-enable UI on error
            try:
                self.btn_load.IsEnabled = True
                self.btn_cancel.IsEnabled = True
            except:
                pass

            logger.error("Critical error in load_clicked: {}".format(ex))
            logger.error(traceback.format_exc())
            forms.alert("Critical error loading families:\n{}".format(str(ex)[:200]), exitscript=False)

    def cancel_clicked(self, sender, e):
        """Cancel and close dialog (or cancel scan if in progress)"""
        try:
            # If scan is in progress, cancel it
            if self._scan_thread and self._scan_thread.is_alive():
                logger.info("User requested scan cancellation")
                self._cancel_requested = True
                # Don't close dialog, let scan complete
                forms.alert("Cancelling scan...", exitscript=False)
                return

            # Clean up resources before closing
            self._cleanup()

            # Close dialog
            self.Close()
        except Exception as ex:
            logger.error("Error in cancel_clicked: {}".format(ex))
            self.Close()

    def _cleanup(self):
        """Clean up resources to prevent memory leaks"""
        try:
            logger.info("Cleaning up Family Loader resources...")

            # Unsubscribe all PropertyChanged events
            for family in self.filtered_families:
                try:
                    family.PropertyChanged -= self.on_family_property_changed
                except:
                    pass

            # Dispose all family items
            for family in self.all_families:
                if hasattr(family, 'Dispose'):
                    family.Dispose()

            # Clear collections
            self.filtered_families.Clear()
            self.all_families = []
            self.category_structure = {}

            logger.info("Cleanup completed successfully")
        except Exception as ex:
            logger.error("Error during cleanup: {}".format(ex))


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
