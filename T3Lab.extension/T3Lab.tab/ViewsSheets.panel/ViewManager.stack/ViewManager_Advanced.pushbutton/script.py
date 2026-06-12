# -*- coding: utf-8 -*-
"""Advanced View Manager with Sheet Manager Style UI
Enhanced view management with summary cards and modern layout
Copyright: Dang Quoc Truong (DQT)
"""

__title__ = "Advanced\nView Manager"
__author__ = "Dang Quoc Truong (DQT)"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, GridLength, GridUnitType, Thickness
from System.Windows.Controls import *
from System.Windows.Media import SolidColorBrush, Color, Brushes
from System.Windows.Forms import SaveFileDialog, OpenFileDialog, DialogResult
from System.Collections.ObjectModel import ObservableCollection
import System

# =====================================================
# REVIT VERSION COMPATIBILITY (2024-2027)
# =====================================================

def _eid_int(eid):
    """Get integer value from ElementId - compatible with Revit 2024-2027.
    Revit 2024-2025: ElementId.IntegerValue (int)
    Revit 2026+: ElementId.Value (long, IntegerValue removed)
    """
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue

def _make_eid(int_val):
    """Create ElementId from integer - compatible with Revit 2024-2027.
    Revit 2024-2025: ElementId(int)
    Revit 2026+: ElementId(long)
    """
    try:
        return ElementId(int(int_val))
    except:
        try:
            return ElementId(System.Int64(int(int_val)))
        except:
            return ElementId(int_val)

# =====================================================
# CONFIG - DQT COLORS
# =====================================================

class Config:
    PRIMARY_COLOR = "#0F172A"      # DQT Gold
    BACKGROUND_COLOR = "#F8FAFC"   # Light cream
    ACCENT_COLOR = "#E8A317"       # Dark gold
    CARD_BG = "#FFFFFF"            # White cards
    BORDER_COLOR = "#DDDDDD"       # Light gray
    TEXT_PRIMARY = "#0F172A"       # Dark gray
    TEXT_SECONDARY = "#475569"     # Medium gray

# =====================================================
# ENHANCED VIEW ITEM
# =====================================================

class EnhancedViewItem:
    """Enhanced view item with all properties"""
    
    def __init__(self, view, doc):
        self.element = view
        self.doc = doc
        self.id = view.Id
        self.name = view.Name
        self.view_type = self._get_view_type_name(view)
        self.view_template = self._get_view_template(view)
        self.scale = self._get_scale(view)
        self.detail_level = self._get_detail_level(view)
        self.on_sheets = self._get_sheet_count(view)
        self.title_on_sheet = self._get_title_on_sheet(view)
        self.referencing_sheet = self._get_referencing_sheet(view)
        self.sheet_number = self._get_sheet_number(view)
        self.sheet_name = self._get_sheet_name(view)
        self.level_name = self._get_level_name(view)
        
        # Crop Box data
        crop_data = self._get_crop_data(view)
        self.crop_active = crop_data[0]
        self.crop_visible = crop_data[1]
        self.crop_min = crop_data[2]   # "x,y,z" string
        self.crop_max = crop_data[3]   # "x,y,z" string
    
    def _get_crop_data(self, view):
        """Get crop box data: (active, visible, min_str, max_str)"""
        try:
            crop_active = "Yes" if view.CropBoxActive else "No"
        except:
            crop_active = "No"
        
        try:
            crop_visible = "Yes" if view.CropBoxVisible else "No"
        except:
            crop_visible = "No"
        
        crop_min = ""
        crop_max = ""
        try:
            if view.CropBoxActive or True:  # Always export crop box if available
                bb = view.CropBox
                if bb is not None:
                    mn = bb.Min
                    mx = bb.Max
                    # Round to 6 decimal places (Revit internal units = feet)
                    crop_min = "{},{},{}".format(
                        round(mn.X, 6), round(mn.Y, 6), round(mn.Z, 6))
                    crop_max = "{},{},{}".format(
                        round(mx.X, 6), round(mx.Y, 6), round(mx.Z, 6))
        except:
            pass
        
        return (crop_active, crop_visible, crop_min, crop_max)
    
    def _get_level_name(self, view):
        """Get the associated level name for plan views"""
        try:
            if hasattr(view, 'GenLevel') and view.GenLevel is not None:
                return view.GenLevel.Name
        except:
            pass
        try:
            level_param = view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL)
            if level_param and level_param.AsString():
                return level_param.AsString()
        except:
            pass
        return ""
        
    def _get_view_type_name(self, view):
        view_type_dict = {
            ViewType.FloorPlan: "Floor Plan",
            ViewType.CeilingPlan: "Ceiling Plan",
            ViewType.Elevation: "Elevation",
            ViewType.Section: "Section",
            ViewType.ThreeD: "3D View",
            ViewType.DraftingView: "Drafting View",
            ViewType.EngineeringPlan: "Structural Plan",
            ViewType.AreaPlan: "Area Plan",
            ViewType.Detail: "Detail View",
            ViewType.Legend: "Legend",
            ViewType.Schedule: "Schedule",
            ViewType.DrawingSheet: "Sheet"
        }
        return view_type_dict.get(view.ViewType, str(view.ViewType))
    
    def _get_view_template(self, view):
        try:
            template_id = view.ViewTemplateId
            if template_id and template_id != ElementId.InvalidElementId:
                template = self.doc.GetElement(template_id)
                return template.Name if template else "None"
            return "None"
        except:
            return "None"
    
    def _get_scale(self, view):
        try:
            return view.Scale if hasattr(view, 'Scale') else 0
        except:
            return 0
    
    def _get_detail_level(self, view):
        try:
            detail_dict = {
                ViewDetailLevel.Coarse: "Coarse",
                ViewDetailLevel.Medium: "Medium",
                ViewDetailLevel.Fine: "Fine"
            }
            return detail_dict.get(view.DetailLevel, "N/A")
        except:
            return "N/A"
    
    def _get_sheet_count(self, view):
        try:
            count = 0
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            for vp in collector:
                if vp.ViewId == view.Id:
                    count += 1
            return count
        except:
            return 0
    
    def _get_title_on_sheet(self, view):
        try:
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            for vp in collector:
                if vp.ViewId == view.Id:
                    title_param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                    if title_param:
                        return title_param.AsString() or "N/A"
                    return "N/A"
            return "N/A"
        except:
            return "N/A"
    
    def _get_referencing_sheet(self, view):
        try:
            param = view.get_Parameter(BuiltInParameter.VIEW_REFERENCING_SHEET)
            if param:
                return param.AsString() or "N/A"
            return "N/A"
        except:
            return "N/A"
    
    def _get_sheet_number(self, view):
        try:
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            for vp in collector:
                if vp.ViewId == view.Id:
                    sheet = self.doc.GetElement(vp.SheetId)
                    if sheet:
                        return sheet.SheetNumber or "N/A"
            return "N/A"
        except:
            return "N/A"
    
    def _get_sheet_name(self, view):
        try:
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            for vp in collector:
                if vp.ViewId == view.Id:
                    sheet = self.doc.GetElement(vp.SheetId)
                    if sheet:
                        return sheet.Name or "N/A"
            return "N/A"
        except:
            return "N/A"

# =====================================================
# PREVIEW ITEM FOR BATCH RENAME
# =====================================================

class PreviewItem:
    def __init__(self, old_name, new_name):
        self.old_name = old_name
        self.new_name = new_name

# =====================================================
# BATCH RENAME DIALOG
# =====================================================

class BatchRenameDialog(Window):
    """Dialog for batch renaming views"""
    
    def __init__(self, views, doc):
        self.views = views
        self.doc = doc
        self.preview_items = ObservableCollection[object]()
        
        self.Title = "Batch Rename Views - Dang Quoc Truong (DQT)"
        self.Width = 800
        self.Height = 600
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner
        
        # Colors
        bg_color = Color.FromArgb(255, 254, 248, 231)
        self.Background = SolidColorBrush(bg_color)
        
        self._build_ui()
        self._update_preview()
    
    def _build_ui(self):
        main_grid = Grid()
        main_grid.Margin = Thickness(15)
        
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        
        # Title
        title_text = TextBlock()
        title_text.Text = "Batch Rename {0} View(s)".format(len(self.views))
        title_text.FontSize = 18
        title_text.FontWeight = System.Windows.FontWeights.Bold
        title_text.Margin = Thickness(0, 0, 0, 15)
        Grid.SetRow(title_text, 0)
        main_grid.Children.Add(title_text)
        
        # Options
        options = self._create_options()
        Grid.SetRow(options, 1)
        main_grid.Children.Add(options)
        
        # Preview
        preview = self._create_preview()
        Grid.SetRow(preview, 2)
        main_grid.Children.Add(preview)
        
        # Buttons
        buttons = self._create_buttons()
        Grid.SetRow(buttons, 3)
        main_grid.Children.Add(buttons)
        
        # Copyright
        copyright = TextBlock()
        copyright.Text = "Copyright (c) Dang Quoc Truong (DQT)"
        copyright.FontSize = 10
        copyright.Foreground = Brushes.Gray
        copyright.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        copyright.Margin = Thickness(0, 10, 0, 0)
        Grid.SetRow(copyright, 4)
        main_grid.Children.Add(copyright)
        
        self.Content = main_grid
    
    def _create_options(self):
        border = Border()
        border.BorderBrush = Brushes.Gray
        border.BorderThickness = Thickness(1)
        border.Padding = Thickness(10)
        border.CornerRadius = System.Windows.CornerRadius(5)
        
        stack = StackPanel()
        
        # Find/Replace
        find_grid = Grid()
        find_grid.Margin = Thickness(0, 0, 0, 10)
        find_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(80)))
        find_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(200)))
        find_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(80)))
        find_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(200)))
        
        find_lbl = TextBlock()
        find_lbl.Text = "Find:"
        find_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(find_lbl, 0)
        
        self.find_box = TextBox()
        self.find_box.Margin = Thickness(5, 0, 15, 0)
        self.find_box.TextChanged += self._on_option_changed
        Grid.SetColumn(self.find_box, 1)
        
        replace_lbl = TextBlock()
        replace_lbl.Text = "Replace:"
        replace_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(replace_lbl, 2)
        
        self.replace_box = TextBox()
        self.replace_box.Margin = Thickness(5, 0, 0, 0)
        self.replace_box.TextChanged += self._on_option_changed
        Grid.SetColumn(self.replace_box, 3)
        
        find_grid.Children.Add(find_lbl)
        find_grid.Children.Add(self.find_box)
        find_grid.Children.Add(replace_lbl)
        find_grid.Children.Add(self.replace_box)
        
        # Prefix
        prefix_grid = Grid()
        prefix_grid.Margin = Thickness(0, 0, 0, 10)
        prefix_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(80)))
        prefix_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(200)))
        
        prefix_lbl = TextBlock()
        prefix_lbl.Text = "Add Prefix:"
        prefix_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(prefix_lbl, 0)
        
        self.prefix_box = TextBox()
        self.prefix_box.Margin = Thickness(5, 0, 0, 0)
        self.prefix_box.TextChanged += self._on_option_changed
        Grid.SetColumn(self.prefix_box, 1)
        
        prefix_grid.Children.Add(prefix_lbl)
        prefix_grid.Children.Add(self.prefix_box)
        
        # Suffix
        suffix_grid = Grid()
        suffix_grid.Margin = Thickness(0, 0, 0, 10)
        suffix_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(80)))
        suffix_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(200)))
        
        suffix_lbl = TextBlock()
        suffix_lbl.Text = "Add Suffix:"
        suffix_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(suffix_lbl, 0)
        
        self.suffix_box = TextBox()
        self.suffix_box.Margin = Thickness(5, 0, 0, 0)
        self.suffix_box.TextChanged += self._on_option_changed
        Grid.SetColumn(self.suffix_box, 1)
        
        suffix_grid.Children.Add(suffix_lbl)
        suffix_grid.Children.Add(self.suffix_box)
        
        # Case
        case_grid = Grid()
        case_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(80)))
        case_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(150)))
        
        case_lbl = TextBlock()
        case_lbl.Text = "Change Case:"
        case_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
        Grid.SetColumn(case_lbl, 0)
        
        self.case_combo = ComboBox()
        self.case_combo.Margin = Thickness(5, 0, 0, 0)
        self.case_combo.Items.Add("No Change")
        self.case_combo.Items.Add("UPPERCASE")
        self.case_combo.Items.Add("lowercase")
        self.case_combo.Items.Add("Title Case")
        self.case_combo.SelectedIndex = 0
        self.case_combo.SelectionChanged += self._on_option_changed
        Grid.SetColumn(self.case_combo, 1)
        
        case_grid.Children.Add(case_lbl)
        case_grid.Children.Add(self.case_combo)
        
        stack.Children.Add(find_grid)
        stack.Children.Add(prefix_grid)
        stack.Children.Add(suffix_grid)
        stack.Children.Add(case_grid)
        
        border.Child = stack
        return border
    
    def _create_preview(self):
        border = Border()
        border.BorderBrush = Brushes.Gray
        border.BorderThickness = Thickness(1)
        border.Padding = Thickness(10)
        border.Margin = Thickness(0, 10, 0, 10)
        
        stack = StackPanel()
        
        preview_lbl = TextBlock()
        preview_lbl.Text = "Preview (showing first 20)"
        preview_lbl.FontWeight = System.Windows.FontWeights.Bold
        preview_lbl.Margin = Thickness(0, 0, 0, 10)
        
        self.preview_grid = DataGrid()
        self.preview_grid.IsReadOnly = True
        self.preview_grid.AutoGenerateColumns = False
        self.preview_grid.ItemsSource = self.preview_items
        self.preview_grid.MaxHeight = 300
        
        old_col = DataGridTextColumn()
        old_col.Header = "Current Name"
        old_col.Binding = System.Windows.Data.Binding("old_name")
        old_col.Width = DataGridLength(1, DataGridLengthUnitType.Star)
        
        new_col = DataGridTextColumn()
        new_col.Header = "New Name"
        new_col.Binding = System.Windows.Data.Binding("new_name")
        new_col.Width = DataGridLength(1, DataGridLengthUnitType.Star)
        
        self.preview_grid.Columns.Add(old_col)
        self.preview_grid.Columns.Add(new_col)
        
        stack.Children.Add(preview_lbl)
        stack.Children.Add(self.preview_grid)
        
        border.Child = stack
        return border
    
    def _create_buttons(self):
        stack = StackPanel()
        stack.Orientation = Orientation.Horizontal
        stack.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        
        apply_btn = Button()
        apply_btn.Content = "Apply Rename"
        apply_btn.Width = 120
        apply_btn.Height = 35
        apply_btn.Margin = Thickness(0, 0, 10, 0)
        apply_btn.Click += self._on_apply
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 100
        cancel_btn.Height = 35
        cancel_btn.Click += self._on_cancel
        
        stack.Children.Add(apply_btn)
        stack.Children.Add(cancel_btn)
        
        return stack
    
    def _apply_rename_rules(self, name):
        new_name = name
        
        if self.find_box.Text:
            new_name = new_name.replace(self.find_box.Text, self.replace_box.Text)
        
        if self.prefix_box.Text:
            new_name = self.prefix_box.Text + new_name
        
        if self.suffix_box.Text:
            new_name = new_name + self.suffix_box.Text
        
        case_option = str(self.case_combo.SelectedItem) if self.case_combo.SelectedItem else "No Change"
        if case_option == "UPPERCASE":
            new_name = new_name.upper()
        elif case_option == "lowercase":
            new_name = new_name.lower()
        elif case_option == "Title Case":
            new_name = new_name.title()
        
        return new_name
    
    def _update_preview(self):
        self.preview_items.Clear()
        
        preview_count = min(20, len(self.views))
        
        for i in range(preview_count):
            view = self.views[i]
            old_name = view.name
            new_name = self._apply_rename_rules(old_name)
            
            item = PreviewItem(old_name, new_name)
            self.preview_items.Add(item)
    
    def _on_option_changed(self, sender, args):
        self._update_preview()
    
    def _on_apply(self, sender, args):
        t = Transaction(self.doc, "Batch Rename Views")
        t.Start()
        
        try:
            renamed = 0
            failed = 0
            
            for view_item in self.views:
                old_name = view_item.name
                new_name = self._apply_rename_rules(old_name)
                
                if new_name == old_name:
                    continue
                
                try:
                    view_item.element.Name = new_name
                    renamed += 1
                except:
                    failed += 1
            
            t.Commit()
            
            msg = "Renamed {0} view(s)".format(renamed)
            if failed > 0:
                msg += "\nFailed: {0} view(s)".format(failed)
            
            MessageBox.Show(msg, "Complete")
            self.DialogResult = True
            self.Close()
            
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {0}".format(str(e)), "Error")
    
    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

# =====================================================
# MAIN WINDOW - SHEET MANAGER STYLE
# =====================================================

class AdvancedViewManagerWindow(Window):
    """Advanced view manager with Sheet Manager style UI"""
    
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.all_views = []
        self.filtered_views = ObservableCollection[object]()
        
        # Initialize text references
        self.total_value_text = None
        self.selected_value_text = None
        self.types_value_text = None
        self.filters_value_text = None
        
        self.Title = "Advanced View Manager - Dang Quoc Truong (DQT)"
        self.Width = 1400
        self.Height = 850
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        # Background
        bg_color = Color.FromArgb(255, 254, 248, 231)
        self.Background = SolidColorBrush(bg_color)
        
        # Build UI first (creates summary card text elements)
        self._build_ui()
        
        # Then load data (will update summary cards)
        self._load_all_views()
        self._apply_filters()
    
    def _build_ui(self):
        main_grid = Grid()
        
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(80)))   # Title header
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(100)))  # Summary cards
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # Content
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))   # Actions
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(35)))   # Footer
        
        # Title header
        title_header = self._create_title_header()
        Grid.SetRow(title_header, 0)
        main_grid.Children.Add(title_header)
        
        # Summary cards
        summary = self._create_summary_cards()
        Grid.SetRow(summary, 1)
        main_grid.Children.Add(summary)
        
        # Content area (filters + grid)
        content = self._create_content_area()
        Grid.SetRow(content, 2)
        main_grid.Children.Add(content)
        
        # Actions
        actions = self._create_actions()
        Grid.SetRow(actions, 3)
        main_grid.Children.Add(actions)
        
        # Footer
        footer = self._create_footer()
        Grid.SetRow(footer, 4)
        main_grid.Children.Add(footer)
        
        self.Content = main_grid
    
    def _create_title_header(self):
        """Create title header like Sheet Manager"""
        border = Border()
        # Background màu cream nhạt hơn
        bg_color = Color.FromArgb(255, 254, 248, 231)
        border.Background = SolidColorBrush(bg_color)
        border.Padding = Thickness(20, 12, 20, 8)
        
        stack = StackPanel()
        stack.Orientation = Orientation.Horizontal
        stack.VerticalAlignment = System.Windows.VerticalAlignment.Bottom
        
        # Main title - màu gold nhạt như Sheet Manager
        title = TextBlock()
        title.Text = "View Manager"
        title.FontSize = 32
        title.FontWeight = System.Windows.FontWeights.Normal
        # Gold nhạt giống Sheet Manager
        gold_light = Color.FromArgb(255, 240, 204, 136)
        title.Foreground = SolidColorBrush(gold_light)
        
        # Spacer
        spacer = TextBlock()
        spacer.Text = "  "
        
        # Version - nhỏ hơn, cùng baseline
        version = TextBlock()
        version.Text = "v1.0.0"
        version.FontSize = 16
        version.FontWeight = System.Windows.FontWeights.Normal
        gold_light2 = Color.FromArgb(255, 240, 204, 136)
        version.Foreground = SolidColorBrush(gold_light2)
        version.VerticalAlignment = System.Windows.VerticalAlignment.Bottom
        version.Margin = Thickness(0, 0, 0, 4)
        
        stack.Children.Add(title)
        stack.Children.Add(spacer)
        stack.Children.Add(version)
        
        border.Child = stack
        return border
    
    def _create_summary_cards(self):
        """Create summary cards panel like Sheet Manager"""
        main_stack = StackPanel()
        
        # Subtitle bar - màu gold đậm
        subtitle_border = Border()
        # Gold đậm giống Sheet Manager
        gold_dark = Color.FromArgb(255, 218, 165, 32)  # Goldenrod
        subtitle_border.Background = SolidColorBrush(gold_dark)
        subtitle_border.Padding = Thickness(20, 10, 20, 10)
        
        subtitle_stack = StackPanel()
        subtitle_stack.Orientation = Orientation.Horizontal
        
        subtitle_text = TextBlock()
        subtitle_text.Text = "View Manager v1.0"
        subtitle_text.FontSize = 14
        subtitle_text.FontWeight = System.Windows.FontWeights.Bold
        subtitle_text.Foreground = Brushes.Black
        
        subtitle_author = TextBlock()
        subtitle_author.Text = "by Dang Quoc Truong (DQT)"
        subtitle_author.FontSize = 11
        subtitle_author.Foreground = Brushes.Black
        subtitle_author.Margin = Thickness(10, 2, 0, 0)
        
        subtitle_stack.Children.Add(subtitle_text)
        subtitle_stack.Children.Add(subtitle_author)
        subtitle_border.Child = subtitle_stack
        
        # Cards grid with proper background
        cards_border = Border()
        # Background cream nhạt
        cream = Color.FromArgb(255, 254, 248, 231)
        cards_border.Background = SolidColorBrush(cream)
        cards_border.Padding = Thickness(10, 10, 10, 10)
        
        cards_grid = Grid()
        cards_grid.Margin = Thickness(0, 0, 0, 10)
        
        # 4 EQUAL width columns using Star like Sheet Manager
        cards_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        cards_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        cards_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        cards_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        # Card 1: TOTAL
        card1 = self._create_info_card("TOTAL", "", 0)
        Grid.SetColumn(card1, 0)
        
        # Card 2: SELECTED
        card2 = self._create_info_card("SELECTED", "0", 1)
        Grid.SetColumn(card2, 1)
        
        # Card 3: CATEGORIES
        card3 = self._create_info_card("CATEGORIES", "", 2)
        Grid.SetColumn(card3, 2)
        
        # Card 4: FILTERS
        card4 = self._create_info_card("FILTERS", "", 3)
        Grid.SetColumn(card4, 3)
        
        cards_grid.Children.Add(card1)
        cards_grid.Children.Add(card2)
        cards_grid.Children.Add(card3)
        cards_grid.Children.Add(card4)
        
        cards_border.Child = cards_grid
        
        main_stack.Children.Add(subtitle_border)
        main_stack.Children.Add(cards_border)
        
        return main_stack
    
    def _create_info_card(self, title, value, index):
        """Create individual info card - Sheet Manager EXACT specs"""
        card_border = Border()
        card_border.Background = Brushes.White
        # Border color
        border_color = Color.FromArgb(255, 212, 184, 122)  # #CBD5E1 from Sheet Manager
        card_border.BorderBrush = SolidColorBrush(border_color)
        card_border.BorderThickness = Thickness(1)
        card_border.CornerRadius = System.Windows.CornerRadius(4)
        card_border.Padding = Thickness(10, 6, 10, 6)  # EXACT from Sheet Manager
        
        # Margin based on position like Sheet Manager
        if index == 0:
            card_border.Margin = Thickness(0, 0, 4, 0)
        elif index == 3:
            card_border.Margin = Thickness(4, 0, 0, 0)
        else:
            card_border.Margin = Thickness(4, 0, 4, 0)
        
        stack = StackPanel()
        
        # Title label
        title_text = TextBlock()
        title_text.Text = title
        title_text.FontSize = 9  # EXACT from Sheet Manager
        title_text.FontWeight = System.Windows.FontWeights.Bold
        # Gray color #666
        label_gray = Color.FromArgb(255, 102, 102, 102)
        title_text.Foreground = SolidColorBrush(label_gray)
        
        # Value text
        value_text = TextBlock()
        value_text.Text = value if value else "..."
        value_text.FontSize = 22  # EXACT from Sheet Manager
        value_text.FontWeight = System.Windows.FontWeights.Bold
        
        # Colors for different cards - like Sheet Manager
        if index == 3:  # FILTERS card - #666
            value_color = Color.FromArgb(255, 102, 102, 102)
            value_text.Foreground = SolidColorBrush(value_color)
        elif index == 2:  # CATEGORIES - #10B981 green
            value_color = Color.FromArgb(255, 76, 175, 80)
            value_text.Foreground = SolidColorBrush(value_color)
        elif index == 1:  # SELECTED - #E5B85C orange
            value_color = Color.FromArgb(255, 229, 184, 92)
            value_text.Foreground = SolidColorBrush(value_color)
        else:  # TOTAL - black
            value_text.Foreground = Brushes.Black
        
        # Store value text for updates
        if index == 0:
            self.total_value_text = value_text
        elif index == 1:
            self.selected_value_text = value_text
        elif index == 2:
            self.types_value_text = value_text
        elif index == 3:
            self.filters_value_text = value_text
        
        stack.Children.Add(title_text)
        stack.Children.Add(value_text)
        
        card_border.Child = stack
        return card_border
    
    def _create_content_area(self):
        """Create content area with left filters and right grid"""
        grid = Grid()
        grid.Margin = Thickness(10)
        # Background cream
        cream = Color.FromArgb(255, 254, 248, 231)
        grid.Background = SolidColorBrush(cream)
        
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(200)))  # Left filters
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))  # Main grid
        
        # Left panel
        left_panel = self._create_left_panel()
        Grid.SetColumn(left_panel, 0)
        grid.Children.Add(left_panel)
        
        # Main grid
        main_panel = self._create_main_grid()
        Grid.SetColumn(main_panel, 1)
        grid.Children.Add(main_panel)
        
        return grid
    
    def _create_left_panel(self):
        """Create left filter panel - Sheet Manager exact style"""
        border = Border()
        border.Background = Brushes.White
        border_gray = Color.FromArgb(255, 230, 230, 230)
        border.BorderBrush = SolidColorBrush(border_gray)
        border.BorderThickness = Thickness(1)
        border.Padding = Thickness(12)
        border.Margin = Thickness(0, 0, 8, 0)
        
        stack = StackPanel()
        
        # SEARCH section
        search_title = TextBlock()
        search_title.Text = "SEARCH"
        search_title.FontSize = 10
        search_title.FontWeight = System.Windows.FontWeights.Bold
        search_title.Foreground = Brushes.Black
        search_title.Margin = Thickness(0, 0, 0, 5)
        
        self.search_box = TextBox()
        self.search_box.Height = 24
        self.search_box.Margin = Thickness(0, 0, 0, 15)
        self.search_box.TextChanged += self._on_search_changed
        
        # FILTER section
        filter_title = TextBlock()
        filter_title.Text = "FILTER"
        filter_title.FontSize = 10
        filter_title.FontWeight = System.Windows.FontWeights.Bold
        filter_title.Foreground = Brushes.Black
        filter_title.Margin = Thickness(0, 0, 0, 5)
        
        self.type_combo = ComboBox()
        self.type_combo.Items.Add("All Sheets")
        self.type_combo.Items.Add("Floor Plan")
        self.type_combo.Items.Add("Ceiling Plan")
        self.type_combo.Items.Add("Section")
        self.type_combo.Items.Add("Elevation")
        self.type_combo.Items.Add("3D View")
        self.type_combo.SelectedIndex = 0
        self.type_combo.Height = 24
        self.type_combo.Margin = Thickness(0, 0, 0, 15)
        self.type_combo.SelectionChanged += self._on_filter_changed
        
        # QUICK SELECT section
        select_title = TextBlock()
        select_title.Text = "QUICK SELECT"
        select_title.FontSize = 10
        select_title.FontWeight = System.Windows.FontWeights.Bold
        select_title.Foreground = Brushes.Black
        select_title.Margin = Thickness(0, 0, 0, 5)
        
        select_all_btn = Button()
        select_all_btn.Content = "Select All"
        select_all_btn.Height = 26
        select_all_btn.Margin = Thickness(0, 0, 0, 4)
        select_all_btn.Click += self._on_select_all
        
        clear_btn = Button()
        clear_btn.Content = "Clear All"
        clear_btn.Height = 26
        clear_btn.Margin = Thickness(0, 0, 0, 15)
        clear_btn.Click += self._on_clear_all
        
        # MORE FILTERS
        more_title = TextBlock()
        more_title.Text = "MORE FILTERS"
        more_title.FontSize = 10
        more_title.FontWeight = System.Windows.FontWeights.Bold
        more_title.Foreground = Brushes.Black
        more_title.Margin = Thickness(0, 0, 0, 5)
        
        # Template
        template_lbl = TextBlock()
        template_lbl.Text = "Has Template:"
        template_lbl.FontSize = 9
        template_lbl.Margin = Thickness(0, 0, 0, 3)
        
        self.template_combo = ComboBox()
        self.template_combo.Items.Add("All Views")
        self.template_combo.Items.Add("With Template")
        self.template_combo.Items.Add("Without Template")
        self.template_combo.SelectedIndex = 0
        self.template_combo.Height = 24
        self.template_combo.Margin = Thickness(0, 0, 0, 8)
        self.template_combo.SelectionChanged += self._on_filter_changed
        
        # Sheets
        sheets_lbl = TextBlock()
        sheets_lbl.Text = "On Sheets:"
        sheets_lbl.FontSize = 9
        sheets_lbl.Margin = Thickness(0, 0, 0, 3)
        
        self.sheets_combo = ComboBox()
        self.sheets_combo.Items.Add("All Views")
        self.sheets_combo.Items.Add("On Sheets")
        self.sheets_combo.Items.Add("Not On Sheets")
        self.sheets_combo.SelectedIndex = 0
        self.sheets_combo.Height = 24
        self.sheets_combo.SelectionChanged += self._on_filter_changed
        
        stack.Children.Add(search_title)
        stack.Children.Add(self.search_box)
        stack.Children.Add(filter_title)
        stack.Children.Add(self.type_combo)
        stack.Children.Add(select_title)
        stack.Children.Add(select_all_btn)
        stack.Children.Add(clear_btn)
        stack.Children.Add(more_title)
        stack.Children.Add(template_lbl)
        stack.Children.Add(self.template_combo)
        stack.Children.Add(sheets_lbl)
        stack.Children.Add(self.sheets_combo)
        
        border.Child = stack
        return border
    
    def _create_main_grid(self):
        """Create main data grid"""
        border = Border()
        border.Background = Brushes.White
        border.BorderBrush = Brushes.LightGray
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(5)
        border.Padding = Thickness(10)
        
        self.data_grid = DataGrid()
        self.data_grid.IsReadOnly = False
        self.data_grid.AutoGenerateColumns = False
        self.data_grid.SelectionMode = DataGridSelectionMode.Extended
        self.data_grid.CanUserSortColumns = True
        self.data_grid.AlternatingRowBackground = System.Windows.Media.Brushes.WhiteSmoke
        self.data_grid.ItemsSource = self.filtered_views
        self.data_grid.SelectionChanged += self._on_selection_changed
        self.data_grid.CellEditEnding += self._on_cell_edit
        self.data_grid.PreviewMouseRightButtonDown += self._on_header_right_click  # RIGHT-CLICK
        
        # Track custom parameter columns
        self.custom_columns = {}  # {col_name: param_name}
        
        # Editable columns
        columns = [
            ("View Name", "name", 200, False),
            ("Type", "view_type", 120, True),
            ("Level", "level_name", 100, True),
            ("Scale", "scale", 60, False),
            ("Detail Level", "detail_level", 100, False),
            ("Title on Sheet", "title_on_sheet", 120, False),
            ("Sheet Number", "sheet_number", 100, True),
            ("Sheet Name", "sheet_name", 150, True),
            ("On Sheets", "on_sheets", 80, True)
        ]
        
        for header, binding, width, readonly in columns:
            col = DataGridTextColumn()
            col.Header = header
            col.Binding = System.Windows.Data.Binding(binding)
            col.Width = DataGridLength(width)
            col.IsReadOnly = readonly
            self.data_grid.Columns.Add(col)
        
        # Template combo column - insert after Level (index 3)
        template_col = DataGridComboBoxColumn()
        template_col.Header = "View Template"
        template_col.Width = DataGridLength(150)
        template_col.SelectedItemBinding = System.Windows.Data.Binding("view_template")
        self.template_items = self._get_all_templates()
        template_col.ItemsSource = self.template_items
        self.data_grid.Columns.Insert(3, template_col)
        
        # Detail combo column - insert after Scale (index 5)
        detail_col = DataGridComboBoxColumn()
        detail_col.Header = "Detail Level"
        detail_col.Width = DataGridLength(100)
        detail_col.SelectedItemBinding = System.Windows.Data.Binding("detail_level")
        detail_col.ItemsSource = ["Coarse", "Medium", "Fine"]
        self.data_grid.Columns.Insert(5, detail_col)
        
        border.Child = self.data_grid
        return border
    
    def _create_actions(self):
        """Create action buttons"""
        border = Border()
        border.BorderBrush = Brushes.LightGray
        border.BorderThickness = Thickness(0, 1, 0, 0)
        border.Padding = Thickness(20, 10, 20, 10)
        
        stack = StackPanel()
        stack.Orientation = Orientation.Horizontal
        stack.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        
        # Excel button
        excel_btn = Button()
        excel_btn.Content = "Excel"
        excel_btn.Width = 100
        excel_btn.Height = 35
        excel_btn.Margin = Thickness(0, 0, 10, 0)
        green_color = Color.FromArgb(255, 76, 175, 80)
        excel_btn.Background = SolidColorBrush(green_color)
        excel_btn.Foreground = Brushes.White
        excel_btn.Click += self._on_excel
        
        # Refresh button - NEW!
        refresh_btn = Button()
        refresh_btn.Content = "Refresh"
        refresh_btn.Width = 100
        refresh_btn.Height = 35
        refresh_btn.Margin = Thickness(0, 0, 10, 0)
        blue_color = Color.FromArgb(255, 33, 150, 243)  # Material Blue
        refresh_btn.Background = SolidColorBrush(blue_color)
        refresh_btn.Foreground = Brushes.White
        refresh_btn.Click += self._on_refresh
        
        rename_btn = Button()
        rename_btn.Content = "Batch Rename"
        rename_btn.Width = 120
        rename_btn.Height = 35
        rename_btn.Margin = Thickness(0, 0, 10, 0)
        rename_btn.Click += self._on_batch_rename
        
        dup_btn = Button()
        dup_btn.Content = "Duplicate"
        dup_btn.Width = 100
        dup_btn.Height = 35
        dup_btn.Margin = Thickness(0, 0, 10, 0)
        dup_btn.Click += self._on_duplicate
        
        del_btn = Button()
        del_btn.Content = "Delete"
        del_btn.Width = 100
        del_btn.Height = 35
        del_btn.Margin = Thickness(0, 0, 10, 0)
        red_color = Color.FromArgb(255, 244, 67, 54)
        del_btn.Background = SolidColorBrush(red_color)
        del_btn.Foreground = Brushes.White
        del_btn.Click += self._on_delete
        
        close_btn = Button()
        close_btn.Content = "Close"
        close_btn.Width = 100
        close_btn.Height = 35
        close_btn.Click += self._on_close
        
        stack.Children.Add(excel_btn)
        stack.Children.Add(refresh_btn)  # NEW!
        stack.Children.Add(rename_btn)
        stack.Children.Add(dup_btn)
        stack.Children.Add(del_btn)
        stack.Children.Add(close_btn)
        
        border.Child = stack
        return border
    
    def _create_footer(self):
        """Create copyright footer - Sheet Manager exact style"""
        border = Border()
        # Gold đậm như Sheet Manager
        gold_dark = Color.FromArgb(255, 218, 165, 32)
        border.Background = SolidColorBrush(gold_dark)
        border.Padding = Thickness(20, 10, 20, 10)
        
        text = TextBlock()
        text.Text = "(c) 2026 Dang Quoc Truong (DQT) - All Rights Reserved"
        text.FontSize = 10
        text.Foreground = Brushes.Black
        text.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        text.VerticalAlignment = System.Windows.VerticalAlignment.Center
        
        border.Child = text
        return border
    
    def _load_all_views(self):
        """Load views"""
        self.all_views = []
        
        collector = FilteredElementCollector(self.doc)\
            .OfClass(View)\
            .WhereElementIsNotElementType()
        
        for view in collector:
            if view.ViewType in [ViewType.ProjectBrowser, ViewType.SystemBrowser,
                                ViewType.Undefined, ViewType.Internal]:
                continue
            
            if view.IsTemplate:
                continue
            
            try:
                item = EnhancedViewItem(view, self.doc)
                self.all_views.append(item)
            except:
                pass
        
        self._update_summary_cards()
    
    def _get_all_templates(self):
        """Get all view templates"""
        templates = ["None"]
        
        collector = FilteredElementCollector(self.doc)\
            .OfClass(View)\
            .WhereElementIsNotElementType()
        
        for view in collector:
            if view.IsTemplate:
                templates.append(view.Name)
        
        return templates
    
    def _apply_filters(self):
        """Apply filters"""
        self.filtered_views.Clear()
        
        type_filter = str(self.type_combo.SelectedItem) if self.type_combo.SelectedItem else "All Sheets"
        template_filter = str(self.template_combo.SelectedItem) if self.template_combo.SelectedItem else "All Views"
        sheets_filter = str(self.sheets_combo.SelectedItem) if self.sheets_combo.SelectedItem else "All Views"
        search_text = self.search_box.Text.lower() if hasattr(self, 'search_box') and self.search_box.Text else ""
        
        for view in self.all_views:
            # Search filter
            if search_text and search_text not in view.name.lower():
                continue
            
            # Type filter - fix: "All Sheets" should show all
            if type_filter != "All Sheets" and view.view_type != type_filter:
                continue
            
            # Template filter
            if template_filter == "With Template" and view.view_template == "None":
                continue
            elif template_filter == "Without Template" and view.view_template != "None":
                continue
            
            # Sheets filter
            if sheets_filter == "On Sheets" and view.on_sheets == 0:
                continue
            elif sheets_filter == "Not On Sheets" and view.on_sheets > 0:
                continue
            
            self.filtered_views.Add(view)
        
        self._update_summary_cards()
    
    def _update_summary_cards(self):
        """Update summary card values"""
        # TOTAL
        if hasattr(self, 'total_value_text') and self.total_value_text is not None:
            total = len(self.all_views)
            self.total_value_text.Text = str(total)
            self.total_value_text.InvalidateVisual()
            self.total_value_text.UpdateLayout()
        
        # CATEGORIES
        if hasattr(self, 'types_value_text') and self.types_value_text is not None:
            types = set(v.view_type for v in self.all_views)
            type_count = len(types)
            self.types_value_text.Text = str(type_count)
            self.types_value_text.InvalidateVisual()
            self.types_value_text.UpdateLayout()
        
        # FILTERS
        if hasattr(self, 'filters_value_text') and self.filters_value_text is not None:
            self.filters_value_text.Text = "Active"
            self.filters_value_text.InvalidateVisual()
            self.filters_value_text.UpdateLayout()
        
        # Force update
        self.InvalidateVisual()
        self.UpdateLayout()
    
    def _on_select_all(self, sender, args):
        """Select all views in grid"""
        self.data_grid.SelectAll()
    
    def _on_clear_all(self, sender, args):
        """Clear all selections"""
        self.data_grid.UnselectAll()
    
    def _on_selection_changed(self, sender, args):
        """Update selected count"""
        if hasattr(self, 'selected_value_text'):
            selected = self.data_grid.SelectedItems
            self.selected_value_text.Text = str(len(selected))
    
    def _on_search_changed(self, sender, args):
        """Search changed"""
        self._apply_filters()
    
    def _on_filter_changed(self, sender, args):
        """Filter changed"""
        self._apply_filters()
    
    def _on_cell_edit(self, sender, args):
        """Handle cell edit"""
        if args.EditAction == DataGridEditAction.Cancel:
            return
        
        try:
            item = args.Row.Item
            column = args.Column
            
            if column.Header == "View Name":
                new_name = args.EditingElement.Text
                self._update_view_name(item, new_name)
            
            elif column.Header == "View Template":
                new_template = args.EditingElement.SelectedItem
                self._update_view_template(item, new_template)
            
            elif column.Header == "Scale":
                new_scale = args.EditingElement.Text
                self._update_scale(item, new_scale)
            
            elif column.Header == "Detail Level":
                new_detail = args.EditingElement.SelectedItem
                self._update_detail_level(item, new_detail)
            
            elif column.Header == "Title on Sheet":
                new_title = args.EditingElement.Text
                self._update_title_on_sheet(item, new_title)
        
        except Exception as e:
            MessageBox.Show("Error: {0}".format(str(e)), "Error")
    
    def _update_view_name(self, item, new_name):
        """Update view name"""
        if not new_name or new_name.strip() == "":
            MessageBox.Show("View name cannot be empty", "Invalid Name")
            return
        
        t = Transaction(self.doc, "Rename View")
        t.Start()
        
        try:
            view = item.element
            view.Name = new_name
            item.name = new_name
            t.Commit()
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Failed: {0}".format(str(e)), "Error")
    
    def _update_view_template(self, item, template_name):
        """Update template"""
        t = Transaction(self.doc, "Update Template")
        t.Start()
        
        try:
            view = item.element
            
            if template_name == "None":
                view.ViewTemplateId = ElementId.InvalidElementId
            else:
                collector = FilteredElementCollector(self.doc)\
                    .OfClass(View)\
                    .WhereElementIsNotElementType()
                
                for template in collector:
                    if template.IsTemplate and template.Name == template_name:
                        view.ViewTemplateId = template.Id
                        break
            
            item.view_template = template_name
            t.Commit()
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Failed: {0}".format(str(e)), "Error")
    
    def _update_scale(self, item, scale_str):
        """Update scale"""
        t = Transaction(self.doc, "Update Scale")
        t.Start()
        
        try:
            view = item.element
            
            try:
                scale_value = int(scale_str)
                if scale_value > 0:
                    view.Scale = scale_value
                    item.scale = scale_value
            except:
                MessageBox.Show("Scale must be positive integer", "Invalid")
                t.RollBack()
                return
            
            t.Commit()
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Failed: {0}".format(str(e)), "Error")
    
    def _update_detail_level(self, item, detail_str):
        """Update detail level"""
        t = Transaction(self.doc, "Update Detail")
        t.Start()
        
        try:
            view = item.element
            
            detail_map = {
                "Coarse": ViewDetailLevel.Coarse,
                "Medium": ViewDetailLevel.Medium,
                "Fine": ViewDetailLevel.Fine
            }
            
            if detail_str in detail_map:
                view.DetailLevel = detail_map[detail_str]
                item.detail_level = detail_str
            
            t.Commit()
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Failed: {0}".format(str(e)), "Error")
    
    def _update_title_on_sheet(self, item, title_str):
        """Update title on sheet"""
        t = Transaction(self.doc, "Update Title")
        t.Start()
        
        try:
            view = item.element
            
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            updated = False
            for vp in collector:
                if vp.ViewId == view.Id:
                    title_param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                    if title_param:
                        title_param.Set(title_str)
                        item.title_on_sheet = title_str
                        updated = True
                        break
            
            if not updated:
                MessageBox.Show("View not on sheet", "Cannot Update")
                t.RollBack()
                return
            
            t.Commit()
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Failed: {0}".format(str(e)), "Error")
    
    def _on_batch_rename(self, sender, args):
        """Batch rename"""
        selected = list(self.data_grid.SelectedItems)
        
        if not selected:
            MessageBox.Show("Select views", "No Selection")
            return
        
        dialog = BatchRenameDialog(selected, self.doc)
        dialog.Owner = self
        result = dialog.ShowDialog()
        
        if result:
            self._load_all_views()
            self._apply_filters()
    
    def _on_duplicate(self, sender, args):
        """Duplicate"""
        selected = list(self.data_grid.SelectedItems)
        if not selected:
            MessageBox.Show("Select views", "No Selection")
            return
        
        t = Transaction(self.doc, "Duplicate")
        t.Start()
        
        try:
            count = 0
            for item in selected:
                try:
                    view = item.element
                    new_id = view.Duplicate(ViewDuplicateOption.Duplicate)
                    new_view = self.doc.GetElement(new_id)
                    new_view.Name = view.Name + " - Copy"
                    count += 1
                except:
                    pass
            
            t.Commit()
            MessageBox.Show("Duplicated {0} view(s)".format(count), "Success")
            
            self._load_all_views()
            self._apply_filters()
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {0}".format(str(e)), "Error")
    
    def _on_delete(self, sender, args):
        """Delete"""
        selected = list(self.data_grid.SelectedItems)
        if not selected:
            MessageBox.Show("Select views", "No Selection")
            return
        
        result = MessageBox.Show("Delete {0} view(s)?".format(len(selected)), 
                                "Confirm", 
                                MessageBoxButton.YesNo)
        
        if result != System.Windows.MessageBoxResult.Yes:
            return
        
        t = Transaction(self.doc, "Delete")
        t.Start()
        
        try:
            count = 0
            for item in selected:
                try:
                    self.doc.Delete(item.id)
                    count += 1
                except:
                    pass
            
            t.Commit()
            MessageBox.Show("Deleted {0} view(s)".format(count), "Success")
            
            self._load_all_views()
            self._apply_filters()
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {0}".format(str(e)), "Error")
    
    
    def _on_excel(self, sender, args):
        """Excel Export/Import menu"""
        menu = ContextMenu()
        
        export_item = MenuItem()
        export_item.Header = "Export to Excel..."
        export_item.Click += self._on_export_excel
        menu.Items.Add(export_item)
        
        import_item = MenuItem()
        import_item.Header = "Import from Excel (Update Existing)..."
        import_item.Click += self._on_import_excel
        menu.Items.Add(import_item)
        
        sep = System.Windows.Controls.Separator()
        menu.Items.Add(sep)
        
        create_item = MenuItem()
        create_item.Header = "Import from Excel (Create New Views)..."
        create_item.Click += self._on_create_views_from_excel
        menu.Items.Add(create_item)
        
        menu.PlacementTarget = sender
        menu.IsOpen = True
    
    # =====================================================
    # XLSX WRITER - Pure Python (No COM / No Interop)
    # =====================================================
    
    def _write_xlsx(self, filepath, headers, rows, hidden_cols=None, header_colors=None):
        """Write data to .xlsx file using pure Python XML + zipfile.
        No COM, no Interop, no Excel installation required.
        
        Args:
            filepath: output .xlsx path
            headers: list of header strings
            rows: list of lists (each inner list = one row of values)
            hidden_cols: list of 0-based column indices to hide
            header_colors: dict {col_index_0based: hex_color} e.g. {0: 'D4E6A5'}
        """
        import zipfile
        import os
        
        if hidden_cols is None:
            hidden_cols = []
        if header_colors is None:
            header_colors = {}
        
        def _col_letter(idx):
            """Convert 0-based column index to Excel column letter (A, B, ..., Z, AA, AB...)"""
            result = ""
            idx += 1
            while idx > 0:
                idx -= 1
                result = chr(65 + idx % 26) + result
                idx //= 26
            return result
        
        def _escape_xml(val):
            if val is None:
                return ""
            s = str(val)
            # Strip invalid XML control characters
            cleaned = []
            for ch in s:
                code = ord(ch)
                if code == 0x9 or code == 0xA or code == 0xD:
                    cleaned.append(ch)
                elif code >= 0x20:
                    cleaned.append(ch)
            s = "".join(cleaned)
            s = s.replace("&", "&amp;")
            s = s.replace("<", "&lt;")
            s = s.replace(">", "&gt;")
            s = s.replace('"', "&quot;")
            s = s.replace("'", "&apos;")
            return s
        
        # Build unique fill colors
        default_header_color = "0F172A"  # DQT Gold
        fill_colors = [default_header_color]  # index 0
        for ci, color in sorted(header_colors.items()):
            if color not in fill_colors:
                fill_colors.append(color)
        
        # Build fills XML
        fills_xml = '<fills count="{}">\n'.format(len(fill_colors) + 2)
        fills_xml += '<fill><patternFill patternType="none"/></fill>\n'
        fills_xml += '<fill><patternFill patternType="gray125"/></fill>\n'
        for color in fill_colors:
            fills_xml += '<fill><patternFill patternType="solid"><fgColor rgb="FF{}"/><bgColor indexed="64"/></patternFill></fill>\n'.format(color)
        fills_xml += '</fills>\n'
        
        # Build cellXfs (styles) XML
        # Style 0: default
        # Style 1: bold header with default gold
        # Style 2+: bold header with custom colors
        num_styles = 2 + len(fill_colors)
        xfs_xml = '<cellXfs count="{}">\n'.format(num_styles)
        xfs_xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" />\n'  # style 0: default
        xfs_xml += '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"/>\n'  # style 1: bold + default fill
        for fi in range(len(fill_colors)):
            xfs_xml += '<xf numFmtId="0" fontId="1" fillId="{}" borderId="0" applyFont="1" applyFill="1"/>\n'.format(fi + 2)
        xfs_xml += '</cellXfs>\n'
        
        styles_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        styles_xml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        styles_xml += '<fonts count="2">\n'
        styles_xml += '<font><sz val="11"/><name val="Calibri"/></font>\n'
        styles_xml += '<font><b/><sz val="11"/><name val="Calibri"/></font>\n'
        styles_xml += '</fonts>\n'
        styles_xml += fills_xml
        styles_xml += '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n'
        styles_xml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'
        styles_xml += xfs_xml
        styles_xml += '</styleSheet>'
        
        # Build shared strings
        all_strings = []
        string_map = {}
        
        def _get_string_idx(val):
            s = str(val) if val is not None else ""
            if s not in string_map:
                string_map[s] = len(all_strings)
                all_strings.append(s)
            return string_map[s]
        
        # Build sheet data
        sheet_rows = []
        
        # Header row
        header_cells = []
        for ci, h in enumerate(headers):
            col = _col_letter(ci)
            si = _get_string_idx(h)
            # Determine style for this header
            color = header_colors.get(ci, default_header_color)
            if color in fill_colors:
                style_id = fill_colors.index(color) + 2  # +2 because fillId offset
            else:
                style_id = 1
            header_cells.append('<c r="{}1" t="s" s="{}"><v>{}</v></c>'.format(col, style_id, si))
        sheet_rows.append('<row r="1">{}</row>'.format("".join(header_cells)))
        
        # Data rows
        for ri, row_data in enumerate(rows):
            row_num = ri + 2
            cells = []
            for ci, val in enumerate(row_data):
                col = _col_letter(ci)
                ref = "{}{}".format(col, row_num)
                
                if val is None or val == "":
                    cells.append('<c r="{}"><v></v></c>'.format(ref))
                elif isinstance(val, (int, float)):
                    cells.append('<c r="{}"><v>{}</v></c>'.format(ref, val))
                else:
                    si = _get_string_idx(val)
                    cells.append('<c r="{}" t="s"><v>{}</v></c>'.format(ref, si))
            
            sheet_rows.append('<row r="{}">{}</row>'.format(row_num, "".join(cells)))
        
        # Column definitions (for hidden columns and widths)
        cols_xml = '<cols>\n'
        for ci in range(len(headers)):
            width = 15
            hidden = ' hidden="1"' if ci in hidden_cols else ''
            cols_xml += '<col min="{}" max="{}" width="{}" bestFit="1" customWidth="1"{}/>'.format(ci+1, ci+1, width, hidden)
        cols_xml += '\n</cols>\n'
        
        # Sheet XML
        sheet_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        sheet_xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        sheet_xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        sheet_xml += cols_xml
        sheet_xml += '<sheetData>\n'
        sheet_xml += '\n'.join(sheet_rows)
        sheet_xml += '\n</sheetData>\n'
        sheet_xml += '</worksheet>'
        
        # Shared strings XML
        sst_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        sst_xml += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">\n'.format(
            len(all_strings), len(all_strings))
        for s in all_strings:
            sst_xml += '<si><t>{}</t></si>\n'.format(_escape_xml(s))
        sst_xml += '</sst>'
        
        # Workbook XML
        workbook_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        workbook_xml += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        workbook_xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        workbook_xml += '<sheets><sheet name="Views" sheetId="1" r:id="rId1"/></sheets>\n'
        workbook_xml += '</workbook>'
        
        # Relationships
        workbook_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        workbook_rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        workbook_rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\n'
        workbook_rels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
        workbook_rels += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n'
        workbook_rels += '</Relationships>'
        
        rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        rels_xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        rels_xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
        rels_xml += '</Relationships>'
        
        content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        content_types += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        content_types += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        content_types += '<Default Extension="xml" ContentType="application/xml"/>\n'
        content_types += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
        content_types += '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
        content_types += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n'
        content_types += '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n'
        content_types += '</Types>'
        
        # Delete existing file
        if os.path.exists(filepath):
            os.remove(filepath)
        
        # Write ZIP (xlsx is a ZIP archive)
        with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('[Content_Types].xml', content_types)
            zf.writestr('_rels/.rels', rels_xml)
            zf.writestr('xl/workbook.xml', workbook_xml)
            zf.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
            zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
            zf.writestr('xl/styles.xml', styles_xml)
            zf.writestr('xl/sharedStrings.xml', sst_xml)
    
    # =====================================================
    # XLSX READER - Pure Python (No COM / No Interop)
    # =====================================================
    
    def _read_xlsx(self, filepath):
        """Read .xlsx file using pure Python zipfile + XML parsing.
        Handles encoding issues in IronPython by sanitizing XML before parsing.
        Returns: (headers_list, rows_as_lists)
        """
        import zipfile
        import re
        
        try:
            from xml.etree import ElementTree as ET
        except:
            import xml.etree.ElementTree as ET
        
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        
        def _sanitize_xml(raw_bytes):
            """Remove invalid XML characters and fix encoding for IronPython.
            Handles BOM, control chars, and encoding declaration mismatches.
            """
            # Decode bytes to string
            text = None
            for enc in ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']:
                try:
                    text = raw_bytes.decode(enc)
                    break
                except:
                    continue
            
            if text is None:
                text = raw_bytes.decode('utf-8', errors='replace')
            
            # Remove BOM if present
            if text.startswith('\xef\xbb\xbf'):
                text = text[3:]
            if text.startswith('\ufeff'):
                text = text[1:]
            
            # Remove invalid XML 1.0 control characters (except tab, newline, carriage return)
            # Valid: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
            cleaned = []
            for ch in text:
                code = ord(ch)
                if code == 0x9 or code == 0xA or code == 0xD:
                    cleaned.append(ch)
                elif code >= 0x20 and code <= 0xD7FF:
                    cleaned.append(ch)
                elif code >= 0xE000 and code <= 0xFFFD:
                    cleaned.append(ch)
                # Skip all other control chars
            text = "".join(cleaned)
            
            # Fix encoding declaration to match actual content (UTF-8 string)
            # Replace any encoding="..." in XML declaration with encoding="UTF-8"
            text = re.sub(
                r'<\?xml[^?]*\?>',
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                text,
                count=1
            )
            
            return text
        
        def _parse_xml(raw_bytes):
            """Parse XML bytes with sanitization fallback."""
            # Try direct parse first (fastest)
            try:
                return ET.fromstring(raw_bytes)
            except:
                pass
            
            # Sanitize and retry
            text = _sanitize_xml(raw_bytes)
            try:
                return ET.fromstring(text.encode('utf-8'))
            except:
                pass
            
            # Last resort: strip XML declaration entirely
            text = re.sub(r'<\?xml[^?]*\?>', '', text, count=1).strip()
            return ET.fromstring(text.encode('utf-8'))
        
        def _find_all(root, tag):
            """Find all elements by tag name, namespace-agnostic."""
            # Try with namespace
            results = root.findall('.//{%s}%s' % (ns, tag))
            if results:
                return results
            # Try without namespace
            results = root.findall('.//' + tag)
            return results
        
        def _find(elem, tag):
            """Find first child element by tag name, namespace-agnostic."""
            result = elem.find('{%s}%s' % (ns, tag))
            if result is not None:
                return result
            return elem.find(tag)
        
        with zipfile.ZipFile(filepath, 'r') as zf:
            # Read shared strings
            shared_strings = []
            if 'xl/sharedStrings.xml' in zf.namelist():
                sst_data = zf.read('xl/sharedStrings.xml')
                sst_root = _parse_xml(sst_data)
                for si in _find_all(sst_root, 'si'):
                    texts = []
                    t_elem = _find(si, 't')
                    if t_elem is not None and t_elem.text:
                        texts.append(t_elem.text)
                    else:
                        for t in _find_all(si, 't'):
                            if t.text:
                                texts.append(t.text)
                    shared_strings.append("".join(texts))
            
            # Read sheet1
            sheet_data = zf.read('xl/worksheets/sheet1.xml')
            sheet_root = _parse_xml(sheet_data)
        
        # Parse cell references
        def _col_from_ref(ref):
            col = ""
            for ch in ref:
                if ch.isalpha():
                    col += ch
                else:
                    break
            return col
        
        def _col_to_index(col_str):
            idx = 0
            for ch in col_str.upper():
                idx = idx * 26 + (ord(ch) - ord('A') + 1)
            return idx - 1
        
        def _row_from_ref(ref):
            num = ""
            for ch in ref:
                if ch.isdigit():
                    num += ch
            return int(num) if num else 0
        
        # Read all cells into a dict: (row, col) -> value
        cells = {}
        max_row = 0
        max_col = 0
        
        # Find all rows in sheetData
        all_rows = _find_all(sheet_root, 'row')
        
        # Filter to only rows that are direct children of sheetData
        sheet_data_elem = _find(sheet_root, 'sheetData')
        if sheet_data_elem is None:
            # Try deeper search
            for sd in _find_all(sheet_root, 'sheetData'):
                sheet_data_elem = sd
                break
        
        if sheet_data_elem is not None:
            # Get row elements from sheetData
            row_tag_ns = '{%s}row' % ns
            row_elems = list(sheet_data_elem)
            if not row_elems:
                row_elems = sheet_data_elem.findall(row_tag_ns)
            if not row_elems:
                row_elems = sheet_data_elem.findall('row')
            
            for row_elem in row_elems:
                # Get cell elements
                cell_tag_ns = '{%s}c' % ns
                cell_elems = list(row_elem)
                if not cell_elems:
                    cell_elems = row_elem.findall(cell_tag_ns)
                if not cell_elems:
                    cell_elems = row_elem.findall('c')
                
                for cell in cell_elems:
                    ref = cell.get('r', '')
                    if not ref:
                        continue
                    
                    col_idx = _col_to_index(_col_from_ref(ref))
                    row_idx = _row_from_ref(ref) - 1  # 0-based
                    
                    cell_type = cell.get('t', '')
                    v_elem = _find(cell, 'v')
                    
                    value = None
                    if v_elem is not None and v_elem.text is not None:
                        if cell_type == 's':
                            try:
                                si = int(v_elem.text)
                                value = shared_strings[si] if si < len(shared_strings) else ""
                            except:
                                value = v_elem.text
                        elif cell_type == 'b':
                            value = v_elem.text == '1'
                        else:
                            try:
                                fval = float(v_elem.text)
                                if fval == int(fval):
                                    value = int(fval)
                                else:
                                    value = fval
                            except:
                                value = v_elem.text
                    else:
                        # Check for inline string
                        is_elem = None
                        is_parent = _find(cell, 'is')
                        if is_parent is not None:
                            is_elem = _find(is_parent, 't')
                        if is_elem is not None and is_elem.text:
                            value = is_elem.text
                    
                    if value is not None:
                        cells[(row_idx, col_idx)] = value
                        if row_idx > max_row:
                            max_row = row_idx
                        if col_idx > max_col:
                            max_col = col_idx
        
        if not cells:
            return [], []
        
        # Extract headers (row 0)
        headers = []
        for ci in range(max_col + 1):
            val = cells.get((0, ci), "")
            headers.append(str(val) if val else "")
        
        # Extract data rows
        rows = []
        for ri in range(1, max_row + 1):
            row = []
            has_data = False
            for ci in range(max_col + 1):
                val = cells.get((ri, ci))
                row.append(val)
                if val is not None and val != "":
                    has_data = True
            if has_data:
                rows.append(row)
        
        return headers, rows
    
    # =====================================================
    # EXPORT
    # =====================================================
    
    def _on_export_excel(self, sender, args):
        """Export views to Excel - pure Python, no COM needed"""
        try:
            dialog = SaveFileDialog()
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            dialog.Title = "Export Views to Excel"
            dialog.FileName = "Views_Export.xlsx"
            
            if dialog.ShowDialog() != DialogResult.OK:
                return
            
            filepath = dialog.FileName
            
            # Headers
            base_headers = ["Element ID", "View Name", "Type", "Level", "View Template", "Scale", "Detail Level", 
                          "Title on Sheet", "Sheet Number", "Sheet Name", "On Sheets",
                          "Crop Active", "Crop Visible", "Crop Min", "Crop Max"]
            
            all_headers = base_headers[:]
            if self.custom_columns:
                for col_name in self.custom_columns.keys():
                    all_headers.append(col_name)
            
            # Build rows
            rows = []
            for view_item in self.filtered_views:
                try:
                    row = [
                        _eid_int(view_item.id),
                        view_item.name or "",
                        view_item.view_type or "",
                        view_item.level_name or "",
                        view_item.view_template or "",
                        view_item.scale or "",
                        view_item.detail_level or "",
                        view_item.title_on_sheet or "",
                        view_item.sheet_number or "",
                        view_item.sheet_name or "",
                        view_item.on_sheets or "",
                        view_item.crop_active or "",
                        view_item.crop_visible or "",
                        view_item.crop_min or "",
                        view_item.crop_max or "",
                    ]
                    
                    # Custom parameter columns
                    for i in range(len(self.custom_columns)):
                        binding_name = "param_{}".format(i)
                        value = getattr(view_item, binding_name, "") if hasattr(view_item, binding_name) else ""
                        row.append(value or "")
                    
                    rows.append(row)
                except:
                    continue
            
            # Header colors: col 0 = green (Element ID), others = gold, custom = cream
            header_colors = {0: 'D4E6A5'}
            for ci in range(1, len(base_headers)):
                header_colors[ci] = '0F172A'
            for ci in range(len(base_headers), len(all_headers)):
                header_colors[ci] = 'CCEBFF'
            
            # Write
            self._write_xlsx(filepath, all_headers, rows, hidden_cols=[0], header_colors=header_colors)
            
            custom_msg = ""
            if self.custom_columns:
                custom_msg = "\n\nIncluding {} custom parameter column(s): {}".format(
                    len(self.custom_columns), 
                    ", ".join(self.custom_columns.keys())
                )
            
            MessageBox.Show(
                "Exported {0} views to:\n{1}\n\nTip: Don't delete column A (Element ID) - it's needed for import!{2}".format(
                    len(rows), filepath, custom_msg),
                "Export Successful",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            )
        
        except Exception as e:
            MessageBox.Show("Export error: {0}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    # =====================================================
    # IMPORT (Update Existing)
    # =====================================================
    
    def _on_import_excel(self, sender, args):
        """Import views from Excel - update existing views"""
        try:
            dialog = OpenFileDialog()
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            dialog.Title = "Import Views from Excel (Update Existing)"
            
            if dialog.ShowDialog() != DialogResult.OK:
                return
            
            filepath = dialog.FileName
            
            headers, rows = self._read_xlsx(filepath)
            
            if not headers or not rows:
                MessageBox.Show("No data found in Excel file", "No Data",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            # Map header names to column indices
            header_map = {}
            for i, h in enumerate(headers):
                header_map[h.strip()] = i
            
            # Detect custom parameter columns (not in base set)
            base_names = {"Element ID", "View Name", "Type", "Level", "View Template", "Scale", 
                         "Detail Level", "Title on Sheet", "Sheet Number", "Sheet Name", "On Sheets",
                         "Crop Active", "Crop Visible", "Crop Min", "Crop Max"}
            custom_param_cols = {}
            for h, ci in header_map.items():
                if h and h not in base_names:
                    custom_param_cols[ci] = h
            
            # Build updates list
            updates = []
            for row in rows:
                def _get(col_name):
                    idx = header_map.get(col_name)
                    if idx is not None and idx < len(row):
                        return row[idx]
                    return None
                
                element_id = _get("Element ID")
                view_name = _get("View Name")
                
                if not element_id and not view_name:
                    continue
                
                update = {
                    'element_id': int(float(element_id)) if element_id else None,
                    'view_name': str(view_name) if view_name else None,
                    'template': str(_get("View Template") or ""),
                    'scale': _get("Scale"),
                    'detail_level': str(_get("Detail Level") or ""),
                    'title': None,
                    'custom_params': {},
                    'crop_active': str(_get("Crop Active") or "").strip(),
                    'crop_visible': str(_get("Crop Visible") or "").strip(),
                    'crop_min': str(_get("Crop Min") or "").strip(),
                    'crop_max': str(_get("Crop Max") or "").strip(),
                }
                
                for ci, param_name in custom_param_cols.items():
                    if ci < len(row) and row[ci]:
                        update['custom_params'][param_name] = str(row[ci])
                
                updates.append(update)
            
            if not updates:
                MessageBox.Show("No valid data found in Excel file", "No Data",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            custom_msg = ""
            if custom_param_cols:
                custom_msg = "\n\nIncluding {} custom parameter column(s)".format(len(custom_param_cols))
            
            result = MessageBox.Show(
                "Update {0} existing views from Excel?{1}".format(len(updates), custom_msg),
                "Confirm Import",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question
            )
            
            if result != System.Windows.MessageBoxResult.Yes:
                return
            
            self._apply_excel_updates(updates)
        
        except Exception as e:
            MessageBox.Show("Import error: {0}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    # =====================================================
    # IMPORT (Create New Views)
    # =====================================================
    
    def _on_create_views_from_excel(self, sender, args):
        """Create new views from Excel file"""
        try:
            dialog = OpenFileDialog()
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            dialog.Title = "Import Excel - Create New Views"
            
            if dialog.ShowDialog() != DialogResult.OK:
                return
            
            filepath = dialog.FileName
            
            headers, rows = self._read_xlsx(filepath)
            
            if not headers or not rows:
                MessageBox.Show("No data found in Excel file", "No Data",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            header_map = {}
            for i, h in enumerate(headers):
                header_map[h.strip()] = i
            
            view_defs = []
            for row in rows:
                def _get(col_name):
                    idx = header_map.get(col_name)
                    if idx is not None and idx < len(row):
                        return row[idx]
                    return None
                
                view_name = _get("View Name")
                view_type = _get("Type")
                
                if not view_name or not view_type:
                    continue
                
                view_def = {
                    'name': str(view_name).strip(),
                    'type': str(view_type).strip(),
                    'level': str(_get("Level") or "").strip(),
                    'template': str(_get("View Template") or ""),
                    'scale': _get("Scale"),
                    'detail_level': str(_get("Detail Level") or ""),
                    'crop_active': str(_get("Crop Active") or "").strip(),
                    'crop_visible': str(_get("Crop Visible") or "").strip(),
                    'crop_min': str(_get("Crop Min") or "").strip(),
                    'crop_max': str(_get("Crop Max") or "").strip(),
                }
                view_defs.append(view_def)
            
            if not view_defs:
                MessageBox.Show("No valid view definitions found.\n\n"
                              "Required columns:\n"
                              "  B: View Name\n"
                              "  C: Type (Floor Plan, Ceiling Plan, Drafting View, etc.)",
                              "No Data", MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            supported_types = ["Floor Plan", "Ceiling Plan", "Structural Plan",
                              "Drafting View", "Area Plan", "3D View", "Section", "Legend"]
            
            creatable = []
            skipped = []
            
            for vd in view_defs:
                if vd['type'] in supported_types:
                    creatable.append(vd)
                else:
                    skipped.append(vd)
            
            if not creatable:
                skip_msg = "\n".join(["  - {} ({})".format(s['name'], s['type']) for s in skipped[:10]])
                MessageBox.Show(
                    "No views can be created.\n\n"
                    "Supported types: {}\n\n"
                    "Skipped:\n{}".format(", ".join(supported_types), skip_msg),
                    "Cannot Create", MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            type_counts = {}
            for vd in creatable:
                t = vd['type']
                type_counts[t] = type_counts.get(t, 0) + 1
            
            summary_lines = ["  {} x {}".format(cnt, tp) for tp, cnt in type_counts.items()]
            skip_msg = ""
            if skipped:
                skip_msg = "\n\nSkipped ({} unsupported):\n".format(len(skipped))
                skip_msg += "\n".join(["  - {} ({})".format(s['name'], s['type']) for s in skipped[:5]])
                if len(skipped) > 5:
                    skip_msg += "\n  ... and {} more".format(len(skipped) - 5)
            
            result = MessageBox.Show(
                "Create {} new views?\n\n{}{}\n\n"
                "Level matching: Uses 'Level' column if available, "
                "otherwise tries to match level name from view name.\n"
                "Views with names already existing will be skipped.".format(
                    len(creatable), "\n".join(summary_lines), skip_msg),
                "Confirm Create Views",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question
            )
            
            if result != System.Windows.MessageBoxResult.Yes:
                return
            
            self._create_views_from_defs(creatable)
        
        except Exception as e:
            MessageBox.Show("Create views error: {0}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    # =====================================================
    # CREATE VIEWS ENGINE
    # =====================================================
    
    def _create_views_from_defs(self, view_defs):
        """Create views from definition list"""
        existing_names = set()
        collector = FilteredElementCollector(self.doc)\
            .OfClass(View)\
            .WhereElementIsNotElementType()
        for v in collector:
            if not v.IsTemplate:
                existing_names.add(v.Name)
        
        levels = {}
        for lvl in FilteredElementCollector(self.doc).OfClass(Level):
            levels[lvl.Name] = lvl
        
        vf_types = {}
        for vft in FilteredElementCollector(self.doc).OfClass(ViewFamilyType):
            fam = vft.ViewFamily
            if fam not in vf_types:
                vf_types[fam] = vft
        
        templates = {}
        for v in FilteredElementCollector(self.doc).OfClass(View).WhereElementIsNotElementType():
            if v.IsTemplate:
                templates[v.Name] = v.Id
        
        area_schemes = {}
        try:
            for scheme in FilteredElementCollector(self.doc).OfClass(AreaScheme):
                area_schemes[scheme.Name] = scheme.Id
        except:
            pass
        
        t = Transaction(self.doc, "DQT - Create Views from Excel")
        t.Start()
        
        try:
            created = 0
            failed = []
            dup_skipped = 0
            
            detail_map = {
                'Coarse': ViewDetailLevel.Coarse,
                'Medium': ViewDetailLevel.Medium,
                'Fine': ViewDetailLevel.Fine
            }
            
            type_to_family = {
                "Floor Plan": ViewFamily.FloorPlan,
                "Ceiling Plan": ViewFamily.CeilingPlan,
                "Structural Plan": ViewFamily.StructuralPlan,
                "Drafting View": ViewFamily.Drafting,
                "3D View": ViewFamily.ThreeDimensional,
                "Section": ViewFamily.Section,
                "Area Plan": ViewFamily.AreaPlan,
                "Legend": ViewFamily.Legend,
            }
            
            for vd in view_defs:
                view_name = vd['name']
                view_type = vd['type']
                
                if view_name in existing_names:
                    dup_skipped += 1
                    continue
                
                try:
                    new_view = None
                    vf = type_to_family.get(view_type)
                    
                    if vf is None:
                        failed.append((view_name, "Unsupported type: {}".format(view_type)))
                        continue
                    
                    vft = vf_types.get(vf)
                    if vft is None:
                        failed.append((view_name, "No ViewFamilyType for {}".format(view_type)))
                        continue
                    
                    if view_type in ["Floor Plan", "Ceiling Plan", "Structural Plan"]:
                        # Priority 1: Use explicit "Level" column from Excel
                        # Priority 2: Fall back to matching level from view name
                        matched_level = None
                        
                        level_hint = vd.get('level', '').strip()
                        if level_hint:
                            # Exact match on Level column
                            matched_level = levels.get(level_hint)
                            # Case-insensitive fallback
                            if matched_level is None:
                                for lvl_name, lvl in levels.items():
                                    if lvl_name.lower() == level_hint.lower():
                                        matched_level = lvl
                                        break
                        
                        # Fallback: match from view name
                        if matched_level is None:
                            matched_level = self._match_level(view_name, levels)
                        
                        if matched_level is None:
                            available_levels = ", ".join(sorted(levels.keys()))
                            failed.append((view_name, 
                                "No matching Level found.\n"
                                "    Level hint: '{}'\n"
                                "    Available levels: {}".format(
                                    level_hint or "(empty)", available_levels)))
                            continue
                        
                        # Try ViewPlan.Create first
                        try:
                            new_view = ViewPlan.Create(self.doc, vft.Id, matched_level.Id)
                        except:
                            # ViewPlan.Create fails if a plan of same type+level already exists
                            # Fallback: find existing view of same type+level and Duplicate it
                            existing_plan = None
                            target_vt = {
                                "Floor Plan": ViewType.FloorPlan,
                                "Ceiling Plan": ViewType.CeilingPlan,
                                "Structural Plan": ViewType.EngineeringPlan,
                            }.get(view_type)
                            
                            for v in FilteredElementCollector(self.doc).OfClass(ViewPlan):
                                try:
                                    if (v.ViewType == target_vt 
                                        and not v.IsTemplate 
                                        and hasattr(v, 'GenLevel') 
                                        and v.GenLevel is not None
                                        and v.GenLevel.Id == matched_level.Id):
                                        existing_plan = v
                                        break
                                except:
                                    continue
                            
                            if existing_plan:
                                try:
                                    new_id = existing_plan.Duplicate(ViewDuplicateOption.Duplicate)
                                    new_view = self.doc.GetElement(new_id)
                                except Exception as dup_ex:
                                    failed.append((view_name, 
                                        "Cannot create (plan already exists for Level '{}') "
                                        "and Duplicate also failed: {}".format(
                                            matched_level.Name, str(dup_ex))))
                                    continue
                            else:
                                failed.append((view_name, 
                                    "Cannot create {} for Level '{}'. "
                                    "A plan may already exist for this level+type combination.".format(
                                        view_type, matched_level.Name)))
                                continue
                    
                    elif view_type == "Drafting View":
                        new_view = ViewDrafting.Create(self.doc, vft.Id)
                    
                    elif view_type == "3D View":
                        new_view = View3D.CreateIsometric(self.doc, vft.Id)
                    
                    elif view_type == "Section":
                        bb = BoundingBoxXYZ()
                        bb.Min = XYZ(-10, -10, -10)
                        bb.Max = XYZ(10, 10, 10)
                        transform = Transform.Identity
                        transform.Origin = XYZ(0, 0, 0)
                        transform.BasisX = XYZ(1, 0, 0)
                        transform.BasisY = XYZ(0, 0, 1)
                        transform.BasisZ = XYZ(0, -1, 0)
                        bb.Transform = transform
                        new_view = ViewSection.CreateSection(self.doc, vft.Id, bb)
                    
                    elif view_type == "Area Plan":
                        matched_level = None
                        level_hint = vd.get('level', '').strip()
                        if level_hint:
                            matched_level = levels.get(level_hint)
                            if matched_level is None:
                                for lvl_name, lvl in levels.items():
                                    if lvl_name.lower() == level_hint.lower():
                                        matched_level = lvl
                                        break
                        if matched_level is None:
                            matched_level = self._match_level(view_name, levels)
                        if matched_level is None:
                            failed.append((view_name, "No matching Level."))
                            continue
                        scheme_id = list(area_schemes.values())[0] if area_schemes else None
                        if scheme_id:
                            new_view = ViewPlan.CreateAreaPlan(self.doc, scheme_id, matched_level.Id)
                        else:
                            failed.append((view_name, "No Area Scheme in project."))
                            continue
                    
                    elif view_type == "Legend":
                        existing_legend = None
                        for v in FilteredElementCollector(self.doc).OfClass(View):
                            if v.ViewType == ViewType.Legend and not v.IsTemplate:
                                existing_legend = v
                                break
                        if existing_legend:
                            new_id = existing_legend.Duplicate(ViewDuplicateOption.Duplicate)
                            new_view = self.doc.GetElement(new_id)
                        else:
                            failed.append((view_name, "No Legend to duplicate."))
                            continue
                    
                    if new_view:
                        try:
                            new_view.Name = view_name
                        except:
                            pass
                        
                        if vd.get('scale'):
                            try:
                                new_view.Scale = int(vd['scale'])
                            except:
                                pass
                        
                        if vd.get('detail_level') and vd['detail_level'] in detail_map:
                            try:
                                new_view.DetailLevel = detail_map[vd['detail_level']]
                            except:
                                pass
                        
                        if vd.get('template') and vd['template'] != "None" and vd['template'] in templates:
                            try:
                                new_view.ViewTemplateId = templates[vd['template']]
                            except:
                                pass
                        
                        # Apply Crop Box
                        self._apply_crop_box(new_view, vd)
                        
                        created += 1
                        existing_names.add(view_name)
                
                except Exception as ex:
                    failed.append((view_name, str(ex)))
            
            t.Commit()
            
            msg = "Created {} new view(s)!".format(created)
            if dup_skipped > 0:
                msg += "\n{} skipped (name already exists).".format(dup_skipped)
            if failed:
                msg += "\n\nFailed ({}):\n".format(len(failed))
                for name, reason in failed[:15]:
                    # Truncate long reasons
                    short_reason = reason.split('\n')[0] if '\n' in reason else reason
                    if len(short_reason) > 80:
                        short_reason = short_reason[:77] + "..."
                    msg += "  - {}: {}\n".format(name, short_reason)
                if len(failed) > 15:
                    msg += "  ... and {} more\n".format(len(failed) - 15)
            
            MessageBox.Show(msg, "Create Views Complete",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            self._refresh_all_data()
        
        except Exception as e:
            if t.HasStarted():
                t.RollBack()
            MessageBox.Show("Error creating views: {0}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _match_level(self, view_name, levels):
        """Match a Level from view name. Exact match first, then longest substring."""
        view_name_lower = view_name.lower()
        
        for lvl_name, lvl in levels.items():
            if lvl_name.lower() == view_name_lower:
                return lvl
        
        best_match = None
        best_len = 0
        for lvl_name, lvl in levels.items():
            if lvl_name.lower() in view_name_lower:
                if len(lvl_name) > best_len:
                    best_match = lvl
                    best_len = len(lvl_name)
        return best_match
    
    def _apply_crop_box(self, view, vd):
        """Apply crop box settings from view definition dict.
        vd keys: crop_active, crop_visible, crop_min, crop_max
        crop_min/crop_max format: "x,y,z" (float, Revit internal units = feet)
        """
        try:
            crop_min_str = vd.get('crop_min', '').strip()
            crop_max_str = vd.get('crop_max', '').strip()
            
            if crop_min_str and crop_max_str:
                # Parse coordinates
                min_parts = [float(v.strip()) for v in crop_min_str.split(',')]
                max_parts = [float(v.strip()) for v in crop_max_str.split(',')]
                
                if len(min_parts) == 3 and len(max_parts) == 3:
                    new_bb = view.CropBox
                    if new_bb is None:
                        new_bb = BoundingBoxXYZ()
                    
                    new_bb.Min = XYZ(min_parts[0], min_parts[1], min_parts[2])
                    new_bb.Max = XYZ(max_parts[0], max_parts[1], max_parts[2])
                    
                    view.CropBox = new_bb
            
            # Set crop active
            crop_active = vd.get('crop_active', '').strip().lower()
            if crop_active == 'yes':
                view.CropBoxActive = True
            elif crop_active == 'no':
                view.CropBoxActive = False
            
            # Set crop visible
            crop_visible = vd.get('crop_visible', '').strip().lower()
            if crop_visible == 'yes':
                view.CropBoxVisible = True
            elif crop_visible == 'no':
                view.CropBoxVisible = False
        except:
            pass  # Crop box is optional - don't fail the view creation
    
    def _apply_excel_updates(self, updates):
        """Apply updates from Excel import"""
        t = Transaction(self.doc, "Import from Excel")
        t.Start()
        
        try:
            count = 0
            skipped = 0
            custom_param_updates = 0
            custom_param_errors = []
            
            for update in updates:
                view = None
                
                # Try to find view by Element ID first (most reliable)
                if update.get('element_id'):
                    try:
                        view_elem = self.doc.GetElement(_make_eid(update['element_id']))
                        if view_elem and isinstance(view_elem, View):
                            view = view_elem
                    except:
                        pass
                
                # Fallback: find by view name if Element ID failed
                if not view and update.get('view_name'):
                    for v in self.all_views:
                        if v.name == update['view_name']:
                            view = v.element
                            break
                
                if not view:
                    skipped += 1
                    continue
                
                # Update name
                if update.get('view_name') and update['view_name'] != view.Name:
                    try:
                        view.Name = update['view_name']
                    except:
                        pass
                
                # Update template
                if update.get('template') and update['template'] != "None":
                    templates = FilteredElementCollector(self.doc)\
                        .OfClass(View)\
                        .WhereElementIsElementType()
                    for tmpl in templates:
                        if tmpl.Name == update['template']:
                            try:
                                view.ViewTemplateId = tmpl.Id
                            except:
                                pass
                            break
                
                # Update scale
                if update.get('scale'):
                    try:
                        view.Scale = int(update['scale'])
                    except:
                        pass
                
                # Update detail level
                if update.get('detail_level'):
                    detail_map = {
                        'Coarse': ViewDetailLevel.Coarse,
                        'Medium': ViewDetailLevel.Medium,
                        'Fine': ViewDetailLevel.Fine
                    }
                    if update['detail_level'] in detail_map:
                        try:
                            view.DetailLevel = detail_map[update['detail_level']]
                        except:
                            pass
                
                # Update custom parameters (if any)
                if update.get('custom_params'):
                    for param_name, param_value in update['custom_params'].items():
                        try:
                            param = view.LookupParameter(param_name)
                            if param:
                                if param.IsReadOnly:
                                    if param_name not in [e[0] for e in custom_param_errors]:
                                        custom_param_errors.append((param_name, "Read-only parameter"))
                                    continue
                                
                                # Try to set value based on storage type
                                success = False
                                if param.StorageType == StorageType.String:
                                    param.Set(str(param_value))
                                    success = True
                                elif param.StorageType == StorageType.Integer:
                                    try:
                                        param.Set(int(float(param_value)))
                                        success = True
                                    except:
                                        pass
                                elif param.StorageType == StorageType.Double:
                                    try:
                                        param.Set(float(param_value))
                                        success = True
                                    except:
                                        pass
                                
                                if success:
                                    custom_param_updates += 1
                            else:
                                if param_name not in [e[0] for e in custom_param_errors]:
                                    custom_param_errors.append((param_name, "Parameter not found"))
                        except Exception as e:
                            if param_name not in [e[0] for e in custom_param_errors]:
                                custom_param_errors.append((param_name, str(e)))
                
                # Update crop box
                self._apply_crop_box(view, update)
                
                count += 1
            
            t.Commit()
            
            msg = "Updated {0} views from Excel!".format(count)
            if skipped > 0:
                msg += "\n{0} views skipped (not found).".format(skipped)
            
            if custom_param_updates > 0:
                msg += "\n\nCustom parameters: {0} updates applied successfully!".format(custom_param_updates)
            
            if custom_param_errors:
                msg += "\n\nWarnings:"
                for param_name, error in custom_param_errors[:5]:  # Show first 5 errors
                    msg += "\n- {}: {}".format(param_name, error)
                if len(custom_param_errors) > 5:
                    msg += "\n... and {} more errors".format(len(custom_param_errors) - 5)
            
            MessageBox.Show(msg, "Import Successful" if count > 0 else "Import Complete",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            
            # Refresh ALL data including custom parameters
            self._refresh_all_data()
        
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error applying updates: {0}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _refresh_all_data(self):
        """Refresh views and custom parameter values"""
        # Reload all views from Revit
        self._load_all_views()
        
        # Re-populate custom parameter columns if any exist
        if self.custom_columns:
            for i, (col_name, param_name) in enumerate(self.custom_columns.items()):
                binding_name = "param_{}".format(i)
                
                # Update all items with fresh parameter values
                for item in self.all_views:
                    try:
                        param = item.element.LookupParameter(param_name)
                        if param and param.HasValue:
                            if param.StorageType == StorageType.String:
                                value = param.AsString() or ""
                            elif param.StorageType == StorageType.Integer:
                                value = str(param.AsInteger())
                            elif param.StorageType == StorageType.Double:
                                value = str(param.AsDouble())
                            elif param.StorageType == StorageType.ElementId:
                                elem_id = param.AsElementId()
                                if elem_id and _eid_int(elem_id) > 0:
                                    elem = self.doc.GetElement(elem_id)
                                    value = elem.Name if elem else str(_eid_int(elem_id))
                                else:
                                    value = ""
                            else:
                                value = param.AsValueString() or ""
                        else:
                            value = ""
                        
                        setattr(item, binding_name, value)
                    except:
                        setattr(item, binding_name, "")
        
        # Reapply filters
        self._apply_filters()
        
        # Refresh grid
        self.data_grid.Items.Refresh()
    
    def _on_refresh(self, sender, args):
        """Refresh button handler"""
        try:
            self._refresh_all_data()
            
            MessageBox.Show("Views refreshed successfully!", "Refresh Complete",
                          MessageBoxButton.OK, MessageBoxImage.Information)
        
        except Exception as e:
            MessageBox.Show("Error refreshing views: {0}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_close(self, sender, args):
        """Close"""
        self.Close()
    
    def _on_header_right_click(self, sender, args):
        """Show context menu on header right-click"""
        try:
            # Check if click is on header
            hit_test = System.Windows.Media.VisualTreeHelper.HitTest(self.data_grid, args.GetPosition(self.data_grid))
            
            if hit_test and hit_test.VisualHit:
                # Walk up tree to find if we clicked header
                element = hit_test.VisualHit
                while element:
                    if isinstance(element, System.Windows.Controls.Primitives.DataGridColumnHeader):
                        # Show context menu
                        menu = ContextMenu()
                        
                        # Add parameter column
                        add_item = MenuItem()
                        add_item.Header = "Add Parameter Column..."
                        add_item.Click += self._on_add_parameter_column
                        menu.Items.Add(add_item)
                        
                        # Remove custom columns (if any exist)
                        if self.custom_columns:
                            separator = System.Windows.Controls.Separator()
                            menu.Items.Add(separator)
                            
                            for col_name in self.custom_columns.keys():
                                remove_item = MenuItem()
                                remove_item.Header = "Remove '{}'".format(col_name)
                                remove_item.Tag = col_name
                                remove_item.Click += self._on_remove_parameter_column
                                menu.Items.Add(remove_item)
                        
                        menu.PlacementTarget = element
                        menu.IsOpen = True
                        args.Handled = True
                        return
                    
                    element = System.Windows.Media.VisualTreeHelper.GetParent(element)
        
        except:
            pass
    
    def _on_add_parameter_column(self, sender, args):
        """Add a custom parameter column"""
        try:
            if not self.all_views:
                MessageBox.Show("No views found in project", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
                return
            
            # Get UNION of all parameters from all views (not just common)
            # This way we show ALL parameters that exist in ANY view
            all_params = set()
            
            # Sample multiple views to collect all possible parameters
            sample_size = min(100, len(self.all_views))  # Check up to 100 views
            
            for view_item in self.all_views[:sample_size]:
                if not view_item.element:
                    continue
                
                view = view_item.element
                for param in view.Parameters:
                    if param.Definition and param.Definition.Name:
                        param_name = param.Definition.Name
                        all_params.add(param_name)
            
            if not all_params:
                MessageBox.Show("No parameters found", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
                return
            
            # Sort parameters
            params = sorted(list(all_params))
            
            # Create selection dialog
            dialog = Window()
            dialog.Title = "Add Parameter Column"
            dialog.Width = 500
            dialog.Height = 400
            dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
            
            main_grid = Grid()
            main_grid.Margin = Thickness(20)
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
            
            # Title
            title_panel = StackPanel()
            Grid.SetRow(title_panel, 0)
            
            title = TextBlock()
            title.Text = "Select Parameter to Add as Column"
            title.FontSize = 14
            title.FontWeight = System.Windows.FontWeights.Bold
            title.Margin = Thickness(0, 0, 0, 5)
            title_panel.Children.Add(title)
            
            instruction = TextBlock()
            instruction.Text = "Choose a view parameter from the list below:"
            instruction.FontSize = 11
            gray_color = Color.FromArgb(255, 100, 100, 100)
            instruction.Foreground = SolidColorBrush(gray_color)
            title_panel.Children.Add(instruction)
            
            main_grid.Children.Add(title_panel)
            
            # ListBox with search
            list_container = Grid()
            list_container.Margin = Thickness(0, 10, 0, 10)
            Grid.SetRow(list_container, 1)
            
            list_container.RowDefinitions.Add(RowDefinition(Height=GridLength(35)))
            list_container.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
            
            # Search box
            search_label = TextBlock()
            search_label.Text = "Search:"
            search_label.Margin = Thickness(0, 0, 0, 5)
            search_label.FontSize = 10
            Grid.SetRow(search_label, 0)
            list_container.Children.Add(search_label)
            
            search_box = TextBox()
            search_box.Margin = Thickness(50, 0, 0, 5)
            search_box.Padding = Thickness(5)
            Grid.SetRow(search_box, 0)
            
            # ListBox
            from System.Windows.Controls import ScrollViewer, ListBox, ListBoxItem
            scroll = ScrollViewer()
            scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
            gray_border = Color.FromArgb(255, 200, 200, 200)
            scroll.BorderBrush = SolidColorBrush(gray_border)
            scroll.BorderThickness = Thickness(1)
            Grid.SetRow(scroll, 1)
            
            param_listbox = ListBox()
            param_listbox.Padding = Thickness(5)
            
            # Add parameters to ListBox
            for param_name in params:
                item = ListBoxItem()
                item.Content = param_name
                item.Padding = Thickness(8, 6, 8, 6)
                item.FontSize = 12
                param_listbox.Items.Add(item)
            
            if param_listbox.Items.Count > 0:
                param_listbox.SelectedIndex = 0
            
            scroll.Content = param_listbox
            list_container.Children.Add(scroll)
            
            # Search functionality
            def on_search_changed(s, e):
                search_text = search_box.Text.lower()
                param_listbox.Items.Clear()
                
                for param_name in params:
                    if not search_text or search_text in param_name.lower():
                        item = ListBoxItem()
                        item.Content = param_name
                        item.Padding = Thickness(8, 6, 8, 6)
                        item.FontSize = 12
                        param_listbox.Items.Add(item)
                
                if param_listbox.Items.Count > 0:
                    param_listbox.SelectedIndex = 0
            
            search_box.TextChanged += on_search_changed
            list_container.Children.Add(search_box)
            
            main_grid.Children.Add(list_container)
            
            # Info text
            info_text = TextBlock()
            info_text.Text = "{} parameters available".format(len(params))
            info_text.FontSize = 10
            gray_info = Color.FromArgb(255, 120, 120, 120)
            info_text.Foreground = SolidColorBrush(gray_info)
            info_text.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            info_text.Margin = Thickness(0, 0, 0, 10)
            Grid.SetRow(info_text, 2)
            main_grid.Children.Add(info_text)
            
            # Buttons
            btn_panel = StackPanel()
            btn_panel.Orientation = Orientation.Horizontal
            btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            btn_panel.VerticalAlignment = System.Windows.VerticalAlignment.Bottom
            Grid.SetRow(btn_panel, 2)
            
            result_holder = [False]
            
            def on_ok(s, e):
                if param_listbox.SelectedIndex < 0:
                    MessageBox.Show("Please select a parameter", "Info",
                                  MessageBoxButton.OK, MessageBoxImage.Information)
                    return
                result_holder[0] = True
                dialog.Close()
            
            def on_cancel(s, e):
                result_holder[0] = False
                dialog.Close()
            
            ok_btn = Button()
            ok_btn.Content = "Add Column"
            ok_btn.Width = 100
            ok_btn.Height = 32
            ok_btn.Margin = Thickness(5, 0, 5, 0)
            green_color = Color.FromArgb(255, 76, 175, 80)
            ok_btn.Background = SolidColorBrush(green_color)
            ok_btn.Foreground = Brushes.White
            ok_btn.FontWeight = System.Windows.FontWeights.SemiBold
            ok_btn.Click += on_ok
            btn_panel.Children.Add(ok_btn)
            
            cancel_btn = Button()
            cancel_btn.Content = "Cancel"
            cancel_btn.Width = 100
            cancel_btn.Height = 32
            cancel_btn.Click += on_cancel
            btn_panel.Children.Add(cancel_btn)
            
            main_grid.Children.Add(btn_panel)
            dialog.Content = main_grid
            
            dialog.ShowDialog()
            
            if not result_holder[0] or param_listbox.SelectedIndex < 0:
                return
            
            param_name = param_listbox.SelectedItem.Content
            
            # Check if already exists
            if param_name in self.custom_columns.values():
                MessageBox.Show("This parameter is already displayed", "Info",
                              MessageBoxButton.OK, MessageBoxImage.Information)
                return
            
            # Add column
            col_index = len(self.custom_columns)
            col_name = param_name
            binding_name = "param_{}".format(col_index)
            
            new_col = DataGridTextColumn()
            new_col.Header = col_name
            new_col.Binding = System.Windows.Data.Binding(binding_name)
            new_col.Width = DataGridLength(150)
            new_col.IsReadOnly = True
            self.data_grid.Columns.Add(new_col)
            
            # Track it
            self.custom_columns[col_name] = param_name
            
            # Update all items with parameter value
            populated_count = 0
            for item in self.all_views:
                try:
                    param = item.element.LookupParameter(param_name)
                    if param and param.HasValue:
                        if param.StorageType == StorageType.String:
                            value = param.AsString() or ""
                        elif param.StorageType == StorageType.Integer:
                            value = str(param.AsInteger())
                        elif param.StorageType == StorageType.Double:
                            value = str(param.AsDouble())
                        elif param.StorageType == StorageType.ElementId:
                            elem_id = param.AsElementId()
                            if elem_id and _eid_int(elem_id) > 0:
                                elem = self.doc.GetElement(elem_id)
                                value = elem.Name if elem else str(_eid_int(elem_id))
                            else:
                                value = ""
                        else:
                            value = param.AsValueString() or ""
                        
                        if value:
                            populated_count += 1
                    else:
                        value = ""
                    
                    setattr(item, binding_name, value)
                except:
                    setattr(item, binding_name, "")
            
            # Refresh grid
            self.data_grid.Items.Refresh()
            
            msg = "Parameter column '{}' added successfully!".format(param_name)
            empty_count = len(self.all_views) - populated_count
            if empty_count > 0:
                msg += "\n\n{} of {} views have this parameter with values.".format(
                    populated_count, len(self.all_views))
                msg += "\n{} views don't have this parameter or have empty values.".format(empty_count)
            
            MessageBox.Show(msg, "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        
        except Exception as e:
            MessageBox.Show("Error adding parameter: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_remove_parameter_column(self, sender, args):
        """Remove a custom parameter column"""
        try:
            col_name = sender.Tag
            
            if col_name not in self.custom_columns:
                return
            
            # Find and remove column
            col_to_remove = None
            for col in self.data_grid.Columns:
                if col.Header == col_name:
                    col_to_remove = col
                    break
            
            if col_to_remove:
                self.data_grid.Columns.Remove(col_to_remove)
                del self.custom_columns[col_name]
                
                MessageBox.Show("Parameter column '{}' removed".format(col_name),
                              "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        
        except Exception as e:
            MessageBox.Show("Error removing column: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)

# =====================================================
# MAIN
# =====================================================

if __name__ == "__main__":
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    
    try:
        window = AdvancedViewManagerWindow(doc, uidoc)
        window.ShowDialog()
    except Exception as e:
        import traceback
        TaskDialog.Show("Error", str(e) + "\n\n" + traceback.format_exc())