# -*- coding: utf-8 -*-
"""
Line Style Manager v2.0
Manage and rename Line Styles (Categories) in Revit using DQT shared library

Compatible with Revit 2024, 2025, 2026, 2027

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""
__title__ = "Line Style\nManager"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Manage and rename Line Styles - By DQT"

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

# Add to path if exists
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Try importing
try:
    from pyrevit import revit, forms
    from Autodesk.Revit.DB import (Category, GraphicsStyleType, Transaction,
                                    Color, ElementId)
    
    from base_manager import BaseManagerWindow, BaseItem
    from ui_components import show_warning, show_info, show_error, ask_yes_no
    from batch_rename_dialog import BatchRenameDialog
    
except ImportError as e:
    print("\nIMPORT ERROR: {}".format(str(e)))
    raise

doc = revit.doc


# ============================================================================
# REVIT API VERSION COMPATIBILITY
# ============================================================================

def _eid_int(element_id):
    """Get integer value from ElementId - compatible with Revit 2024-2027
    
    Revit 2024+: ElementId.Value (long)
    Revit 2023-: ElementId.IntegerValue (int)
    
    Args:
        element_id: Autodesk.Revit.DB.ElementId
        
    Returns:
        int: Element ID as integer
    """
    if element_id is None:
        return -1
    
    try:
        # Try Revit 2024+ API first
        return int(element_id.Value)
    except AttributeError:
        # Fallback to Revit 2023- API
        try:
            return element_id.IntegerValue
        except:
            return -1


def _get_invalid_element_id():
    """Get InvalidElementId - compatible with Revit 2024-2027
    
    Revit 2024+: ElementId.InvalidElementId is deprecated, use constructor
    Revit 2023-: ElementId.InvalidElementId works fine
    
    Returns:
        ElementId: Invalid element ID
    """
    try:
        # Revit 2024+ - use constructor with -1
        return ElementId(-1)
    except:
        # Revit 2023- - use static property
        return ElementId.InvalidElementId


# ============================================================================
# LINE STYLE ITEM CLASS
# ============================================================================

class LineStyleItem(INotifyPropertyChanged):
    """Wrapper class for Line Style (Category)
    
    Implements INotifyPropertyChanged for WPF data binding
    """
    
    def __init__(self, category):
        # Initialize event handler list for INotifyPropertyChanged
        self._property_changed_handlers = []
        
        # Store category reference (Line Styles are Categories!)
        self._category = category
        self._is_selected = False
        
        # Get name
        try:
            self._name = category.Name if category.Name else "Unnamed"
        except:
            self._name = "Unnamed"
        
        # Use compatibility function for ElementId
        self._id = _eid_int(category.Id)
        self._usage_count = 0
        self._usage_percentage = 0.0
        
        # Get line style specific properties
        self._color = self.get_color_display()
        self._weight = self.get_weight_display()
        self._pattern = self.get_pattern_display()
    
    def get_color_display(self):
        """Get line color as RGB string"""
        try:
            color = self._category.LineColor
            return "RGB({},{},{})".format(color.Red, color.Green, color.Blue)
        except:
            return "N/A"
    
    def get_color_brush(self):
        """Get WPF brush for color preview"""
        try:
            from System.Windows.Media import SolidColorBrush, Color as WPFColor
            
            color = self._category.LineColor
            wpf_color = WPFColor.FromRgb(color.Red, color.Green, color.Blue)
            return SolidColorBrush(wpf_color)
        except:
            from System.Windows.Media import Brushes
            return Brushes.Gray
    
    def get_weight_display(self):
        """Get line weight"""
        try:
            weight = self._category.GetLineWeight(GraphicsStyleType.Projection)
            return str(weight)
        except:
            return "N/A"
    
    def get_pattern_display(self):
        """Get line pattern name"""
        try:
            pattern_id = self._category.GetLinePatternId(GraphicsStyleType.Projection)
            invalid_id = _get_invalid_element_id()
            
            # Use compatibility function for comparison
            if pattern_id and _eid_int(pattern_id) != _eid_int(invalid_id):
                pattern = doc.GetElement(pattern_id)
                if pattern:
                    return pattern.Name
            return "Solid"
        except:
            return "Solid"
    
    # Additional properties for batch rename Type/Size/Segments
    @property
    def TypeInfo(self):
        """Type info for batch rename - returns Color"""
        return self.get_color_display()
    
    @property
    def SizeInfo(self):
        """Size info for batch rename - returns Weight"""
        return self.get_weight_display()
    
    @property
    def SegmentInfo(self):
        """Segment info for batch rename - returns Pattern name"""
        return self.get_pattern_display()
    
    # Properties that BaseManager expects
    @property
    def Element(self):
        """Return category as 'element' for compatibility with BatchRenameDialog"""
        return self._category
    
    @property
    def Category(self):
        """Direct access to category"""
        return self._category
    
    @property
    def Id(self):
        return self._id
    
    @property
    def Name(self):
        return self._name
    
    @Name.setter
    def Name(self, value):
        if self._name != value:
            self._name = value
            self.OnPropertyChanged("Name")
    
    @property
    def Color(self):
        return self._color
    
    @property
    def ColorBrush(self):
        """WPF brush for color preview"""
        return self.get_color_brush()
    
    @property
    def Weight(self):
        return self._weight
    
    @property
    def Pattern(self):
        return self._pattern
    
    @property
    def UsageCount(self):
        return self._usage_count
    
    @UsageCount.setter
    def UsageCount(self, value):
        if self._usage_count != value:
            self._usage_count = value
            self.OnPropertyChanged("UsageCount")
            self.OnPropertyChanged("UsagePercentageFormatted")
            self.OnPropertyChanged("Usage")
    
    @property
    def UsagePercentage(self):
        return self._usage_percentage
    
    @UsagePercentage.setter
    def UsagePercentage(self, value):
        if self._usage_percentage != value:
            self._usage_percentage = value
            self.OnPropertyChanged("UsagePercentage")
            self.OnPropertyChanged("UsagePercentageFormatted")
            self.OnPropertyChanged("Usage")
    
    @property
    def UsagePercentageFormatted(self):
        """Formatted percentage for display"""
        return "{:.1f}".format(self._usage_percentage)
    
    @property
    def Usage(self):
        """Usage display string"""
        if self._usage_count == 0:
            return "0"
        return "{} ({:.1f}%)".format(self._usage_count, self._usage_percentage)
    
    @property
    def IsSelected(self):
        return self._is_selected
    
    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self.OnPropertyChanged("IsSelected")
    
    # INotifyPropertyChanged implementation
    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    
    def OnPropertyChanged(self, prop_name):
        args = PropertyChangedEventArgs(prop_name)
        for handler in self._property_changed_handlers:
            handler(self, args)


# ============================================================================
# USAGE CALCULATION
# ============================================================================

def calculate_linestyle_usage(doc, line_style_items):
    """
    Calculate usage for line styles by reading GraphicsStyle parameter from CurveElements.
    Line styles are stored in BUILDING_CURVE_GSTYLE parameter, not in Category!
    """
    
    print("\n" + "="*60)
    print("CALCULATING LINE STYLE USAGE")
    print("="*60)
    
    # Reset all counts
    for item in line_style_items:
        item.UsageCount = 0
        item.UsagePercentage = 0.0
    
    # Build lookup by category ID
    category_lookup = {}
    for item in line_style_items:
        if item.Category:
            cat_id = _eid_int(item.Category.Id)
            category_lookup[cat_id] = item
            print("Tracking: ID {} = '{}'".format(cat_id, item.Name))
    
    # Usage dictionary
    usage_dict = {}
    
    # Count ModelLine and DetailLine elements
    from Autodesk.Revit.DB import FilteredElementCollector, CurveElement, BuiltInParameter
    
    try:
        # Get all CurveElements (ModelLine, DetailLine, etc.)
        collector = FilteredElementCollector(doc).OfClass(CurveElement)
        
        total_curves = 0
        matched_curves = 0
        
        for curve in collector:
            total_curves += 1
            
            try:
                # Get GraphicsStyle from parameter (NOT from Category!)
                style_param = curve.get_Parameter(BuiltInParameter.BUILDING_CURVE_GSTYLE)
                
                if not style_param:
                    continue
                
                style_id = style_param.AsElementId()
                
                if not style_id or _eid_int(style_id) <= 0:
                    continue
                
                # Get GraphicsStyle element
                style_elem = doc.GetElement(style_id)
                
                if not style_elem:
                    continue
                
                # Get the category of GraphicsStyle
                if hasattr(style_elem, 'GraphicsStyleCategory'):
                    style_cat = style_elem.GraphicsStyleCategory
                    
                    if style_cat:
                        cat_id = _eid_int(style_cat.Id)
                        
                        if cat_id in category_lookup:
                            if cat_id not in usage_dict:
                                usage_dict[cat_id] = 0
                            usage_dict[cat_id] += 1
                            matched_curves += 1
                    
            except Exception as ex:
                pass
        
        print("\nTotal curves: {}".format(total_curves))
        print("Matched curves: {}".format(matched_curves))
        print("\nUsage by category:")
        
        for cat_id, count in sorted(usage_dict.items(), key=lambda x: x[1], reverse=True)[:10]:
            item = category_lookup.get(cat_id)
            name = item.Name if item else "Unknown"
            print("  {} = {} uses".format(name, count))
                
    except Exception as ex:
        print("Error calculating line style usage: {}".format(str(ex)))
    
    # Calculate totals
    total_refs = sum(usage_dict.values())
    
    # Update items
    for item in line_style_items:
        if item.Category:
            cat_id = _eid_int(item.Category.Id)
            
            if cat_id in usage_dict:
                item.UsageCount = usage_dict[cat_id]
                
                if total_refs > 0:
                    item.UsagePercentage = (float(usage_dict[cat_id]) / total_refs) * 100
                else:
                    item.UsagePercentage = 0.0
    
    print("\nTotal usage: {}".format(total_refs))
    print("="*60 + "\n")


# ============================================================================
# LINE STYLE MANAGER
# ============================================================================

class LineStyleManager(BaseManagerWindow):
    """Manager window for Line Styles"""
    
    def __init__(self):
        # Configuration
        config = {
            'title': 'Line Style Manager',
            'subtitle': 'Manage Line Styles (Categories) - Copyright by Dang Quoc Truong - DQT © 2025',
            'element_type': None,  # Line styles are Categories, not Elements
            'instance_type': None,
            'item_class': LineStyleItem,
            'has_batch_rename': True,
            'has_edit_properties': False,
            'has_duplicate': False,
            'has_delete': True,  # Enable delete button
            'extra_columns': []  # All columns handled in add_standard_columns override
        }
        
        # Store doc reference
        self.doc = doc
        self.config = config
        
        # Call parent constructor manually since we need custom load
        BaseManagerWindow.__init__(self, doc, config)
    
    def load_items(self):
        """Override to load Line Styles (subcategories of Lines category)"""
        self.all_items.Clear()
        
        try:
            # Get the "Lines" category
            from Autodesk.Revit.DB import BuiltInCategory
            
            lines_category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
            
            if not lines_category:
                show_error("Could not find Lines category!")
                return
            
            # Get all subcategories under Lines - these are the Line Styles
            if lines_category.SubCategories:
                for subcat in lines_category.SubCategories:
                    try:
                        # OPTION: Skip system line styles (those with < >)
                        # Uncomment the next 2 lines to hide system styles:
                        # if subcat.Name.startswith('<') and subcat.Name.endswith('>'):
                        #     continue
                        
                        item = self.config['item_class'](subcat)
                        self.all_items.Add(item)
                    except Exception as ex:
                        print("Error loading line style '{}': {}".format(
                            subcat.Name if hasattr(subcat, 'Name') else 'Unknown', str(ex)))
            
            # DO NOT add the main Lines category - only subcategories
            # (Lines category represents default lines without specific style)
                
        except Exception as ex:
            print("Error loading line styles: {}".format(str(ex)))
            show_error("Failed to load line styles:\n\n{}".format(str(ex)))
        
        # Apply filter
        self.apply_filter()
    
    def add_standard_columns(self, grid):
        """Override to match old Line Style Manager column order"""
        from System.Windows.Data import BindingMode
        from System.Windows.Controls import DataGridCheckBoxColumn, DataGridTextColumn, DataGridLength
        from System.Windows.Controls import DataGridTemplateColumn, Orientation
        from System.Windows.Data import Binding
        from System.Windows import VerticalAlignment
        import System.Windows
        
        # Select column
        col_select = DataGridCheckBoxColumn()
        col_select.Header = "Select"
        col_select.Width = DataGridLength(60)
        binding = Binding("IsSelected")
        binding.Mode = BindingMode.TwoWay
        binding.UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
        col_select.Binding = binding
        col_select.CanUserSort = False
        col_select.IsReadOnly = False
        grid.Columns.Add(col_select)
        
        # LINE STYLE NAME column (wide)
        col_name = DataGridTextColumn()
        col_name.Header = "LINE STYLE NAME"
        col_name.Width = DataGridLength(300)
        col_name.Binding = Binding("Name")
        col_name.IsReadOnly = True
        col_name.CanUserSort = True
        col_name.SortMemberPath = "Name"
        grid.Columns.Add(col_name)
        
        # WEIGHT column (before Color)
        col_weight = DataGridTextColumn()
        col_weight.Header = "WEIGHT"
        col_weight.Width = DataGridLength(80)
        col_weight.Binding = Binding("Weight")
        col_weight.IsReadOnly = True
        col_weight.CanUserSort = True
        col_weight.SortMemberPath = "Weight"
        grid.Columns.Add(col_weight)
        
        # COLOR column - with color preview
        from System.Windows.Controls import DataGridTemplateColumn
        from System.Windows import DataTemplate, FrameworkElementFactory
        from System.Windows.Controls import StackPanel as WPFStackPanel, TextBlock as WPFTextBlock, Border as WPFBorder
        from System.Windows.Media import SolidColorBrush as WPFBrush
        
        col_color = DataGridTemplateColumn()
        col_color.Header = "COLOR"
        col_color.Width = DataGridLength(180)
        col_color.CanUserSort = True
        col_color.SortMemberPath = "Color"
        
        # Create data template programmatically
        # Template structure: StackPanel > Border (colored) + TextBlock (RGB text)
        factory = FrameworkElementFactory(WPFStackPanel)
        factory.SetValue(WPFStackPanel.OrientationProperty, Orientation.Horizontal)
        
        # Color preview border
        border_factory = FrameworkElementFactory(WPFBorder)
        border_factory.SetValue(WPFBorder.WidthProperty, 16.0)
        border_factory.SetValue(WPFBorder.HeightProperty, 16.0)
        border_factory.SetValue(WPFBorder.MarginProperty, System.Windows.Thickness(2))
        border_factory.SetValue(WPFBorder.BorderThicknessProperty, System.Windows.Thickness(1))
        border_factory.SetValue(WPFBorder.BorderBrushProperty, WPFBrush(System.Windows.Media.Colors.Gray))
        
        # Bind background to ColorBrush property
        from System.Windows.Data import Binding as WPFBinding
        color_binding = WPFBinding("ColorBrush")
        border_factory.SetBinding(WPFBorder.BackgroundProperty, color_binding)
        
        factory.AppendChild(border_factory)
        
        # Text block for RGB string
        text_factory = FrameworkElementFactory(WPFTextBlock)
        text_factory.SetValue(WPFTextBlock.MarginProperty, System.Windows.Thickness(5, 0, 0, 0))
        text_factory.SetValue(WPFTextBlock.VerticalAlignmentProperty, VerticalAlignment.Center)
        
        text_binding = WPFBinding("Color")
        text_factory.SetBinding(WPFTextBlock.TextProperty, text_binding)
        
        factory.AppendChild(text_factory)
        
        template = DataTemplate()
        template.VisualTree = factory
        
        col_color.CellTemplate = template
        grid.Columns.Add(col_color)
        
        # PATTERN column
        col_pattern = DataGridTextColumn()
        col_pattern.Header = "PATTERN"
        col_pattern.Width = DataGridLength(150)
        col_pattern.Binding = Binding("Pattern")
        col_pattern.IsReadOnly = True
        col_pattern.CanUserSort = True
        col_pattern.SortMemberPath = "Pattern"
        grid.Columns.Add(col_pattern)
        
        # LINE COUNT column (separate from usage %)
        col_count = DataGridTextColumn()
        col_count.Header = "LINE COUNT"
        col_count.Width = DataGridLength(100)
        col_count.Binding = Binding("UsageCount")
        col_count.IsReadOnly = True
        col_count.CanUserSort = True
        col_count.SortMemberPath = "UsageCount"
        grid.Columns.Add(col_count)
        
        # USAGE % column - FORMATTED
        col_percent = DataGridTextColumn()
        col_percent.Header = "USAGE %"
        col_percent.Width = DataGridLength(80)
        col_percent.Binding = Binding("UsagePercentageFormatted")
        col_percent.IsReadOnly = True
        col_percent.CanUserSort = True
        col_percent.SortMemberPath = "UsagePercentage"
        grid.Columns.Add(col_percent)
        
        # ID column
        col_id = DataGridTextColumn()
        col_id.Header = "ID"
        col_id.Width = DataGridLength(100)
        col_id.Binding = Binding("Id")
        col_id.IsReadOnly = True
        col_id.CanUserSort = True
        col_id.SortMemberPath = "Id"
        grid.Columns.Add(col_id)
    
    def calculate_usage(self):
        """Override to use custom calculation for line styles"""
        calculate_linestyle_usage(doc, self.all_items)
    
    def on_rename_click(self, sender, args):
        """Override rename using create-copy-transfer-delete method (like old code)"""
        selected = self.get_selected_items()
        if not selected:
            show_warning("Please select one line style to rename!")
            return
        
        if len(selected) > 1:
            show_warning("Please select only one line style to rename!")
            return
        
        item = selected[0]
        
        # Call the actual rename method
        success = self.rename_line_style_item(item)
        
        if success:
            # Refresh
            self.load_items()
            self.calculate_usage()
            self.update_stats()
    
    def rename_line_style_item(self, item, new_name=None, skip_prompts=False):
        """
        Rename a single line style item using create-copy-transfer-delete method.
        
        Args:
            item: LineStyleItem to rename
            new_name: New name (if None, will prompt user)
            skip_prompts: If True, skip user prompts and warnings
            
        Returns:
            True if successful, False otherwise
        """
        old_style = item.Category
        
        # Check if system style
        if item.Name.startswith('<') and item.Name.endswith('>'):
            if not skip_prompts:
                show_error("Cannot rename system line style '{}'!".format(item.Name))
            return False
        
        # Ask for new name if not provided
        if not new_name:
            new_name = forms.ask_for_string(
                prompt="Enter new name for line style:",
                default=item.Name,
                title="Rename Line Style"
            )
        
        if not new_name or new_name.strip() == "":
            return False
        
        if new_name == item.Name:
            return False
        
        # Sanitize name
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            new_name = new_name.replace(char, '')
        new_name = new_name.strip()
        
        if not new_name:
            if not skip_prompts:
                show_error("Name cannot be empty after removing invalid characters!")
            return False
        
        # Check for conflicts
        from Autodesk.Revit.DB import BuiltInCategory, GraphicsStyleType
        
        lines_category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
        
        if lines_category.SubCategories:
            for subcat in lines_category.SubCategories:
                if _eid_int(subcat.Id) != _eid_int(old_style.Id):
                    try:
                        if subcat.Name == new_name:
                            if not skip_prompts:
                                show_error("A line style with name '{}' already exists!".format(new_name))
                            return False
                    except:
                        continue
        
        # Perform rename using create-copy-transfer-delete method
        try:
            from Autodesk.Revit.DB import TransactionGroup
            
            # Wrap everything in TransactionGroup for better undo
            tg = TransactionGroup(doc, "Rename Line Style '{}'".format(item.Name))
            tg.Start()
            
            # Step 1: Create new subcategory
            new_subcategory = None
            t1 = Transaction(doc, "Create New Line Style")
            t1.Start()
            
            try:
                new_subcategory = doc.Settings.Categories.NewSubcategory(lines_category, new_name)
                
                # Copy properties
                new_subcategory.LineColor = old_style.LineColor
                
                # Copy projection line weight
                new_subcategory.SetLineWeight(
                    old_style.GetLineWeight(GraphicsStyleType.Projection), 
                    GraphicsStyleType.Projection
                )
                
                # Copy cut line weight (if different)
                try:
                    cut_weight = old_style.GetLineWeight(GraphicsStyleType.Cut)
                    new_subcategory.SetLineWeight(cut_weight, GraphicsStyleType.Cut)
                except:
                    pass
                
                # Copy projection line pattern
                pattern_id = old_style.GetLinePatternId(GraphicsStyleType.Projection)
                if pattern_id and _eid_int(pattern_id) > 0:
                    new_subcategory.SetLinePatternId(pattern_id, GraphicsStyleType.Projection)
                
                # Copy cut line pattern
                try:
                    cut_pattern_id = old_style.GetLinePatternId(GraphicsStyleType.Cut)
                    if cut_pattern_id and _eid_int(cut_pattern_id) > 0:
                        new_subcategory.SetLinePatternId(cut_pattern_id, GraphicsStyleType.Cut)
                except:
                    pass
                
                t1.Commit()
                
            except Exception as ex:
                t1.RollBack()
                tg.RollBack()
                if not skip_prompts:
                    show_error("Failed to create new line style:\n\n{}".format(str(ex)))
                return False
            
            if not new_subcategory:
                tg.RollBack()
                if not skip_prompts:
                    show_error("Failed to create new line style!")
                return False
            
            # Step 2: Transfer all lines to new style
            from Autodesk.Revit.DB import FilteredElementCollector, CurveElement, BuiltInParameter
            
            lines_changed = 0
            lines_to_change = []
            
            # Find lines using old style
            collector = FilteredElementCollector(doc).OfClass(CurveElement)
            
            for curve in collector:
                try:
                    style_param = curve.get_Parameter(BuiltInParameter.BUILDING_CURVE_GSTYLE)
                    
                    if style_param:
                        style_id = style_param.AsElementId()
                        
                        if style_id:
                            style_elem = doc.GetElement(style_id)
                            
                            if style_elem and hasattr(style_elem, 'GraphicsStyleCategory'):
                                style_cat = style_elem.GraphicsStyleCategory
                                
                                if style_cat and _eid_int(style_cat.Id) == _eid_int(old_style.Id):
                                    lines_to_change.append(curve)
                except:
                    pass
            
            # Change lines to new style
            if lines_to_change:
                t2 = Transaction(doc, "Transfer Lines to New Style")
                t2.Start()
                
                try:
                    new_graphics_style = new_subcategory.GetGraphicsStyle(GraphicsStyleType.Projection)
                    
                    for line in lines_to_change:
                        try:
                            line.LineStyle = new_graphics_style
                            lines_changed += 1
                        except:
                            continue
                    
                    t2.Commit()
                    
                except Exception as ex:
                    t2.RollBack()
                    tg.RollBack()
                    if not skip_prompts:
                        show_error("Failed to transfer lines:\n\n{}".format(str(ex)))
                    return False
            
            # Step 3: Delete old style
            t3 = Transaction(doc, "Delete Old Line Style")
            t3.Start()
            
            try:
                doc.Delete(old_style.Id)
                t3.Commit()
                
                # Assimilate transaction group - makes it one undo operation
                tg.Assimilate()
                
                if not skip_prompts:
                    show_info(
                        "Line style renamed successfully!\n\n"
                        "Old name: '{}'\n"
                        "New name: '{}'\n"
                        "Lines transferred: {}"
                        .format(item.Name, new_name, lines_changed)
                    )
                
                return True
                
            except Exception as ex:
                t3.RollBack()
                tg.RollBack()
                if not skip_prompts:
                    show_error(
                        "Failed to delete old line style:\n\n{}\n\n"
                        "New style created but old style remains.\n"
                        "Operation rolled back."
                        .format(str(ex))
                    )
                return False
                
        except Exception as ex:
            if not skip_prompts:
                show_error("Failed to rename line style:\n\n{}".format(str(ex)))
            return False
    
    def on_batch_rename_click(self, sender, args):
        """Override batch rename to use shared BatchRenameDialog with custom logic"""
        selected = self.get_selected_items()
        
        if not selected:
            show_warning("Please select line styles to batch rename!")
            return
        
        # Filter out system styles
        system_styles = []
        renameable_styles = []
        
        for item in selected:
            if item.Name.startswith('<') and item.Name.endswith('>'):
                system_styles.append(item.Name)
            else:
                renameable_styles.append(item)
        
        if system_styles:
            warning_msg = "The following system line styles cannot be renamed and will be skipped:\n\n"
            warning_msg += "\n".join("  - " + name for name in system_styles[:10])
            if len(system_styles) > 10:
                warning_msg += "\n  ... and {} more".format(len(system_styles) - 10)
            
            if not renameable_styles:
                warning_msg += "\n\nNo renameable line styles selected!"
                show_warning(warning_msg)
                return
            else:
                warning_msg += "\n\n{} line styles will be renamed.".format(len(renameable_styles))
                show_info(warning_msg)
        
        if not renameable_styles:
            show_warning("No renameable line styles selected!")
            return
        
        # Create custom dialog class that overrides apply_batch_rename
        dialog = LineStyleBatchRenameDialog(doc, renameable_styles, self)
        
        # Show dialog
        dialog.ShowDialog()


class LineStyleBatchRenameDialog(BatchRenameDialog):
    """Custom batch rename dialog for Line Styles that uses create-copy-transfer-delete method"""
    
    def __init__(self, doc, selected_items, parent):
        """Initialize with parent manager reference"""
        self.line_style_manager = parent
        BatchRenameDialog.__init__(self, doc, selected_items, parent)
    
    def get_type_info(self, item):
        """Override to return Color information for line styles"""
        try:
            if hasattr(item, 'TypeInfo'):
                color_info = item.TypeInfo  # RGB(r,g,b)
                print("    DEBUG: Got Color from LineStyleItem: {}".format(color_info))
                return color_info
            
            # Fallback to getting color directly from category
            if hasattr(item, 'Category'):
                try:
                    color = item.Category.LineColor
                    color_str = "RGB({},{},{})".format(color.Red, color.Green, color.Blue)
                    print("    DEBUG: Got Color from Category: {}".format(color_str))
                    return color_str
                except:
                    pass
            
            return "Line"
        except Exception as ex:
            print("    DEBUG: Error getting line style type info: {}".format(str(ex)))
            return "Line"
    
    def get_size_info(self, item):
        """Override to return Weight information for line styles"""
        try:
            if hasattr(item, 'SizeInfo'):
                weight_info = item.SizeInfo  # Weight number
                print("    DEBUG: Got Weight from LineStyleItem: {}".format(weight_info))
                return weight_info
            
            # Fallback to getting weight directly from category
            if hasattr(item, 'Category'):
                try:
                    weight = item.Category.GetLineWeight(GraphicsStyleType.Projection)
                    weight_str = str(weight)
                    print("    DEBUG: Got Weight from Category: {}".format(weight_str))
                    return weight_str
                except:
                    pass
            
            return None
        except Exception as ex:
            print("    DEBUG: Error getting line style size info: {}".format(str(ex)))
            return None
    
    def get_segment_info(self, item):
        """Override to return Pattern name for line styles"""
        try:
            if hasattr(item, 'SegmentInfo'):
                pattern_info = item.SegmentInfo  # Pattern name
                print("    DEBUG: Got Pattern from LineStyleItem: {}".format(pattern_info))
                # Don't return "Solid" as segment info - it's the default
                if pattern_info and pattern_info != "Solid":
                    return pattern_info
                return None
            
            # Fallback to getting pattern directly from category
            if hasattr(item, 'Category'):
                try:
                    pattern_id = item.Category.GetLinePatternId(GraphicsStyleType.Projection)
                    invalid_id = _get_invalid_element_id()
                    
                    if pattern_id and _eid_int(pattern_id) != _eid_int(invalid_id):
                        pattern = doc.GetElement(pattern_id)
                        if pattern and pattern.Name != "Solid":
                            pattern_name = pattern.Name
                            print("    DEBUG: Got Pattern from Category: {}".format(pattern_name))
                            return pattern_name
                except:
                    pass
            
            return None
        except Exception as ex:
            print("    DEBUG: Error getting line style segment info: {}".format(str(ex)))
            return None
    
    def apply_batch_rename(self):
        """Override to use custom line style rename logic"""
        print("\n" + "="*60)
        print("BATCH RENAME LINE STYLES")
        print("Total items: {}".format(len(self.selected_items)))
        print("-"*60)
        
        success_count = 0
        skip_count = 0
        error_count = 0
        
        from batch_rename_dialog import sanitize_revit_name
        
        for item in self.selected_items:
            old_name = item.Name
            new_name = self.apply_rename_rules(item, old_name)
            
            print("\nProcessing: '{}' -> '{}'".format(old_name, new_name))
            
            # Sanitize name
            new_name = sanitize_revit_name(new_name)
            
            # Skip if no change
            if new_name == old_name:
                print("  Skipped: No change in name")
                skip_count += 1
                continue
            
            # Use LineStyleManager's custom rename method
            if self.line_style_manager.rename_line_style_item(item, new_name, skip_prompts=True):
                success_count += 1
                print("  Success!")
            else:
                error_count += 1
                print("  Error!")
        
        print("\n" + "="*60)
        print("Completed: {} success, {} skipped, {} errors".format(
            success_count, skip_count, error_count))
        print("="*60 + "\n")
        
        # Show results
        self.show_results(success_count, skip_count, error_count)
        
        # Auto-refresh parent
        if self.parent:
            try:
                print("\nAuto-refreshing parent window...")
                self.parent.load_items()
                self.parent.calculate_usage()
                self.parent.update_stats()
                print("Parent window refreshed successfully!")
            except Exception as ex:
                print("Warning: Could not refresh parent window: {}".format(str(ex)))
        
        # Close dialog
        self.result = True
        self.Close()
    
    def on_delete_click(self, sender, args):
        """Override delete to handle line style (category) deletion with usage warnings"""
        selected = self.get_selected_items()
        
        if not selected:
            show_warning("Please select line styles to delete!")
            return
        
        # Check for system/read-only line styles
        system_styles = []
        deletable_styles = []
        styles_in_use = []
        
        for item in selected:
            # Check if system style (typically those with < >)
            is_system = item.Name.startswith('<') and item.Name.endswith('>')
            
            if is_system:
                system_styles.append(item.Name)
            else:
                deletable_styles.append(item)
                
                # Check if in use
                if item.UsageCount > 0:
                    styles_in_use.append((item.Name, item.UsageCount))
        
        # Show warnings for system styles
        if system_styles:
            show_error(
                "Cannot delete system line styles:\n\n{}\n\n"
                "These are built-in Revit line styles that cannot be removed."
                .format('\n'.join('  - ' + name for name in system_styles[:10]))
            )
            
            # If all selected are system styles, stop here
            if not deletable_styles:
                return
        
        # Show warning for styles in use
        if styles_in_use:
            # Build warning message
            usage_list = []
            for name, count in styles_in_use[:10]:  # Show max 10
                usage_list.append("  - '{}': {} uses".format(name, count))
            
            if len(styles_in_use) > 10:
                usage_list.append("  ... and {} more".format(len(styles_in_use) - 10))
            
            warning_msg = (
                "WARNING: The following line styles are currently in use:\n\n{}\n\n"
                "Deleting them will affect existing lines in the project!\n\n"
                "Lines using these styles may revert to default style.\n\n"
                "Are you sure you want to continue?"
                .format('\n'.join(usage_list))
            )
            
            if not ask_yes_no(warning_msg, title="Line Styles In Use - Confirm Delete"):
                return
        else:
            # Confirm deletion for unused styles
            if len(deletable_styles) == 1:
                msg = "Are you sure you want to delete line style '{}'?".format(deletable_styles[0].Name)
            else:
                msg = "Are you sure you want to delete {} line styles?".format(len(deletable_styles))
            
            if not ask_yes_no(msg, title="Confirm Delete"):
                return
        
        # Perform deletion
        t = Transaction(doc, "Delete Line Styles")
        t.Start()
        
        deleted_count = 0
        failed_deletes = []
        
        try:
            for item in deletable_styles:
                try:
                    # Delete the subcategory
                    doc.Delete(item.Category.Id)
                    deleted_count += 1
                    
                except Exception as ex:
                    error_msg = str(ex)
                    
                    # Collect failed deletes with reason
                    if "cannot be deleted" in error_msg.lower():
                        failed_deletes.append("'{}' - Cannot be deleted (system style)".format(item.Name))
                    elif "in use" in error_msg.lower():
                        failed_deletes.append("'{}' - Still in use".format(item.Name))
                    else:
                        failed_deletes.append("'{}' - {}".format(item.Name, error_msg[:50]))
            
            t.Commit()
            
            # Show results
            if deleted_count > 0 and not failed_deletes:
                show_info("Successfully deleted {} line style(s)!".format(deleted_count))
            elif deleted_count > 0 and failed_deletes:
                show_warning(
                    "Deleted {} line style(s), but {} failed:\n\n{}"
                    .format(deleted_count, len(failed_deletes), 
                           '\n'.join(failed_deletes[:5]))
                )
            elif failed_deletes:
                show_error(
                    "Failed to delete line styles:\n\n{}"
                    .format('\n'.join(failed_deletes[:10]))
                )
            
            # Refresh
            self.load_items()
            self.calculate_usage()
            self.update_stats()
            
        except Exception as ex:
            t.RollBack()
            show_error("Error during deletion:\n\n{}".format(str(ex)))


# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':
    try:
        manager = LineStyleManager()
        manager.ShowDialog()
    except Exception as ex:
        import traceback
        error_msg = "Error: {}\n\n{}".format(str(ex), traceback.format_exc())
        print(error_msg)
        from Autodesk.Revit.UI import TaskDialog
        TaskDialog.Show("Error", error_msg)