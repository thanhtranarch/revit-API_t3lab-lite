# -*- coding: utf-8 -*-
"""
Fill Pattern (Hatching) Manager
Manage and rename Fill Patterns in Revit using DQT shared library

Copyright (c) 2024 Dang Quoc Truong (DQT)
All rights reserved.
"""
__title__ = "Fill Pattern\nManager"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Manage and rename Fill Patterns (Hatching) - By DQT"

import clr
clr.AddReference('System')

import sys
import os

# Add System.ComponentModel for INotifyPropertyChanged
import System
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

# CRITICAL: Add lib path - MUST be before imports
script_dir = os.path.dirname(__file__)
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(script_dir)))
lib_path = os.path.join(extension_dir, 'lib')

print("="*60)
print("DEBUG: Path Information")
print("="*60)
print("script_dir: {}".format(script_dir))
print("extension_dir: {}".format(extension_dir))
print("lib_path: {}".format(lib_path))
print("lib exists: {}".format(os.path.exists(lib_path)))

if os.path.exists(lib_path):
    print("\nFiles in lib:")
    for f in os.listdir(lib_path):
        print("  - {}".format(f))
print("="*60 + "\n")

# Add to path if exists
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Try importing
try:
    from pyrevit import revit, forms
    from Autodesk.Revit.DB import (FilteredElementCollector, FillPatternElement, 
                                    FillPatternTarget, BuiltInParameter, ElementId)
    
    from base_manager import BaseManagerWindow, BaseItem
    from ui_components import show_warning, show_info, show_error, ask_yes_no
    
    print("SUCCESS: All imports successful!\n")
    
except ImportError as e:
    print("\nIMPORT ERROR: {}".format(str(e)))
    print("\nCurrent sys.path:")
    for p in sys.path:
        print("  - {}".format(p))
    raise

doc = revit.doc


# ============================================================================
# HELPER FUNCTION FOR REVIT 2024/2025+ COMPATIBILITY
# ============================================================================

def get_element_id_value(element_id):
    """Get integer value from ElementId - compatible with Revit 2024 and 2025+
    
    Args:
        element_id: Revit ElementId object
        
    Returns:
        int/long: Element ID as integer
    """
    try:
        # Revit 2025+ uses .Value (returns long)
        return element_id.Value
    except AttributeError:
        # Revit 2024 and earlier use .IntegerValue
        return element_id.IntegerValue


# ============================================================================
# FILL PATTERN ITEM CLASS
# ============================================================================

class FillPatternItem(INotifyPropertyChanged):
    """Wrapper class for FillPatternElement
    
    Implements INotifyPropertyChanged for WPF data binding
    """
    
    def __init__(self, element):
        # Initialize event handler list for INotifyPropertyChanged
        self._property_changed_handlers = []
        
        # Store element reference
        self._element = element
        self._is_selected = False
        
        # Get name directly from element (FillPattern specific)
        try:
            self._name = element.Name if element.Name else "Unnamed"
        except:
            self._name = "Unnamed"
        
        # Get ID using compatibility helper
        self._id = get_element_id_value(element.Id)
        
        # Get fill pattern specific properties
        self._pattern_type = self.get_pattern_type()
        self._target = self.get_target_type()
        
        # Get pattern settings
        self._pattern_settings = self.get_pattern_settings()
        self._grid_count = self.get_grid_count()
    
    def get_pattern_type(self):
        """Get pattern type (Drafting or Model)"""
        try:
            fill_pattern = self._element.GetFillPattern()
            if fill_pattern:
                target = fill_pattern.Target
                if target == FillPatternTarget.Drafting:
                    return "Drafting"
                elif target == FillPatternTarget.Model:
                    return "Model"
            return "Unknown"
        except:
            return "Unknown"
    
    def get_target_type(self):
        """Get target type for backward compatibility"""
        return self._pattern_type
    
    def get_grid_count(self):
        """Get number of grids in pattern"""
        try:
            fill_pattern = self._element.GetFillPattern()
            if fill_pattern:
                grids = fill_pattern.GetFillGrids()
                return len(grids) if grids else 0
            return 0
        except:
            return 0
    
    def get_pattern_settings(self):
        """Get pattern settings description (Parallel lines / Crosshatch)"""
        try:
            fill_pattern = self._element.GetFillPattern()
            if not fill_pattern:
                return "N/A"
            
            grids = fill_pattern.GetFillGrids()
            if not grids or len(grids) == 0:
                return "Solid fill"
            
            grid_count = len(grids)
            
            # Determine pattern type based on grid count and angles
            if grid_count == 1:
                settings = "Parallel lines"
            elif grid_count == 2:
                # Check if it's crosshatch (perpendicular lines)
                grid1 = grids[0]
                grid2 = grids[1]
                
                angle1 = grid1.Angle
                angle2 = grid2.Angle
                
                # Convert to degrees
                import math
                angle1_deg = math.degrees(angle1)
                angle2_deg = math.degrees(angle2)
                
                # Check if angles are approximately perpendicular (90 degrees apart)
                angle_diff = abs(angle1_deg - angle2_deg)
                if abs(angle_diff - 90) < 5 or abs(angle_diff - 270) < 5:
                    settings = "Crosshatch"
                else:
                    settings = "Custom ({} grids)".format(grid_count)
            else:
                settings = "Custom ({} grids)".format(grid_count)
            
            # Add spacing info from first grid
            try:
                first_grid = grids[0]
                offset = first_grid.Offset
                spacing = first_grid.Shift
                
                # Convert from feet to mm
                offset_mm = offset * 304.8
                
                if offset_mm > 0:
                    settings += " - {:.1f}mm".format(offset_mm)
            except:
                pass
            
            return settings
            
        except Exception as ex:
            print("Error getting pattern settings for {}: {}".format(self._name, str(ex)))
            return "N/A"
    
    # Properties from BaseItem that need to be exposed
    @property
    def Element(self):
        return self._element
    
    @property
    def Id(self):
        return self._id
    
    @property
    def Name(self):
        return self._name
    
    @property
    def IsSelected(self):
        return self._is_selected
    
    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self.OnPropertyChanged("IsSelected")
    
    @property
    def PatternType(self):
        return self._pattern_type
    
    @property
    def Target(self):
        return self._target
    
    @property
    def FillPattern(self):
        """Get underlying FillPatternElement for easy access"""
        return self._element
    
    @property
    def PatternSettings(self):
        """Get pattern settings description"""
        return self._pattern_settings
    
    @property
    def GridCount(self):
        """Get number of grids"""
        return self._grid_count
    
    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    
    def OnPropertyChanged(self, prop_name):
        """Raise PropertyChanged event for data binding"""
        args = PropertyChangedEventArgs(prop_name)
        for handler in self._property_changed_handlers:
            handler(self, args)


# ============================================================================
# FILL PATTERN MANAGER WINDOW
# ============================================================================

class FillPatternManager(BaseManagerWindow):
    """Fill Pattern Manager using shared library"""
    
    def __init__(self):
        # Configuration for Fill Pattern Manager
        config = {
            'title': 'FILL PATTERN MANAGER',
            'subtitle': 'Manage and rename Fill Patterns (Hatching)',
            'element_type': FillPatternElement,
            'instance_type': None,  # No usage calculation
            'item_class': FillPatternItem,
            
            # Feature flags
            'has_batch_rename': True,
            'has_edit_properties': False,
            'has_duplicate': True,
            'has_delete': True,
            
            # Extra columns specific to Fill Patterns
            'extra_columns': [
                {
                    'name': 'Type',
                    'binding': 'PatternType',
                    'width': 80,
                    'sortable': True
                },
                {
                    'name': 'Pattern Settings',
                    'binding': 'PatternSettings',
                    'width': 250,
                    'sortable': True
                },
                {
                    'name': 'Grids',
                    'binding': 'GridCount',
                    'width': 60,
                    'sortable': True
                }
            ],
            
            # Edit properties config (for future use)
            'edit_config': {
                'properties': []
            }
        }
        
        # Call parent constructor with doc and config
        super(FillPatternManager, self).__init__(doc, config)
    
    def calculate_usage(self):
        """Override to disable usage calculation completely"""
        # No usage calculation needed
        pass
    
    def on_rename_click(self, sender, args):
        """Override rename to handle fill pattern renaming properly"""
        from pyrevit import forms
        from Autodesk.Revit.DB import Transaction, FilteredElementCollector
        
        selected = self.get_selected_items()
        if not selected:
            show_warning("Please select one fill pattern to rename!")
            return
        
        if len(selected) > 1:
            show_warning("Please select only one fill pattern to rename!")
            return
        
        item = selected[0]
        
        # Ask for new name
        new_name = forms.ask_for_string(
            prompt="Enter new name for fill pattern:",
            default=item.Name,
            title="Rename Fill Pattern"
        )
        
        if not new_name or new_name.strip() == "":
            return
        
        if new_name == item.Name:
            return
        
        # Sanitize name
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            new_name = new_name.replace(char, '')
        new_name = new_name.strip()
        
        if not new_name:
            show_error("Name cannot be empty!")
            return
        
        # Check conflicts
        all_patterns = FilteredElementCollector(doc).OfClass(FillPatternElement)
        for pattern in all_patterns:
            if pattern.Id != item.Element.Id:
                try:
                    if pattern.Name == new_name:
                        show_error("A fill pattern with name '{}' already exists!".format(new_name))
                        return
                except:
                    continue
        
        # Check if system pattern
        try:
            test_name = item.Element.Name
            if test_name and ("Solid fill" in test_name or test_name.startswith("<")):
                show_error("Cannot rename system fill patterns!")
                return
        except:
            pass
        
        # Rename
        t = Transaction(doc, "Rename Fill Pattern")
        t.Start()
        
        try:
            item.Element.Name = new_name
            t.Commit()
            
            show_info("Fill pattern renamed successfully!")
            
            self.load_items()
            self.update_stats()
            
        except Exception as ex:
            t.RollBack()
            error_msg = str(ex)
            
            if "read-only" in error_msg.lower():
                show_error("Cannot rename system fill patterns!")
            elif "duplicate" in error_msg.lower():
                show_error("A fill pattern with this name already exists!")
            else:
                show_error("Failed to rename:\n\n{}".format(error_msg))
    
    def on_duplicate_click(self, sender, args):
        """Override duplicate"""
        from Autodesk.Revit.DB import Transaction
        
        selected = self.get_selected_items()
        if not selected:
            show_warning("Please select at least one fill pattern to duplicate!")
            return
        
        t = Transaction(doc, "Duplicate Fill Patterns")
        t.Start()
        
        try:
            success_count = 0
            error_count = 0
            
            for item in selected:
                try:
                    original_name = item.Name
                    new_name = "Copy of " + original_name
                    
                    new_id = item.FillPattern.Duplicate(new_name)
                    
                    if new_id != ElementId.InvalidElementId:
                        success_count += 1
                    else:
                        error_count += 1
                except Exception as ex:
                    print("Error duplicating {}: {}".format(item.Name, str(ex)))
                    error_count += 1
            
            t.Commit()
            
            msg = "Successfully duplicated {} fill pattern(s)!".format(success_count)
            if error_count > 0:
                msg += "\nFailed: {}".format(error_count)
            
            show_info(msg)
            
            self.load_items()
            self.update_stats()
            
        except Exception as ex:
            t.RollBack()
            show_error("Error duplicating: {}".format(str(ex)))
    
    def on_delete_click(self, sender, args):
        """Override delete"""
        from Autodesk.Revit.DB import Transaction
        
        selected = self.get_selected_items()
        if not selected:
            show_warning("Please select at least one fill pattern to delete!")
            return
        
        if not ask_yes_no("Delete {} fill pattern(s)?".format(len(selected))):
            return
        
        t = Transaction(doc, "Delete Fill Patterns")
        t.Start()
        
        try:
            success_count = 0
            error_count = 0
            
            for item in selected:
                try:
                    doc.Delete(item.Element.Id)
                    success_count += 1
                except Exception as ex:
                    print("Error deleting {}: {}".format(item.Name, str(ex)))
                    error_count += 1
            
            t.Commit()
            
            msg = "Deleted: {}".format(success_count)
            if error_count > 0:
                msg += "\nFailed: {} (may be system patterns or in use)".format(error_count)
            
            show_info(msg)
            
            self.load_items()
            self.update_stats()
            
        except Exception as ex:
            t.RollBack()
            show_error("Error deleting: {}".format(str(ex)))


# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == '__main__':
    try:
        window = FillPatternManager()
        window.ShowDialog()
    except Exception as e:
        print("\nFATAL ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        MessageBox.Show(
            "Error starting Fill Pattern Manager:\n\n{}".format(str(e)),
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        )