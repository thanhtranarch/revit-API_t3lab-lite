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
clr.AddReference('Microsoft.Office.Interop.Excel')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, GridLength, GridUnitType, Thickness
from System.Windows.Controls import *
from System.Windows.Media import SolidColorBrush, Color, Brushes
from System.Windows.Forms import SaveFileDialog, OpenFileDialog, DialogResult
from System.Collections.ObjectModel import ObservableCollection
import Microsoft.Office.Interop.Excel as Excel
import System

# =====================================================
# CONFIG - DQT COLORS
# =====================================================

class Config:
    PRIMARY_COLOR = "#F0CC88"      # DQT Gold
    BACKGROUND_COLOR = "#FEF8E7"   # Light cream
    ACCENT_COLOR = "#E8A317"       # Dark gold
    CARD_BG = "#FFFFFF"            # White cards
    BORDER_COLOR = "#DDDDDD"       # Light gray
    TEXT_PRIMARY = "#333333"       # Dark gray
    TEXT_SECONDARY = "#666666"     # Medium gray

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
                except Exception as e:
                    print("Failed: {0}".format(str(e)))
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
        
        print("DEBUG: Window initialized")
        print("DEBUG: total_value_text exists: {0}".format(self.total_value_text is not None))
        print("DEBUG: selected_value_text exists: {0}".format(self.selected_value_text is not None))
        print("DEBUG: types_value_text exists: {0}".format(self.types_value_text is not None))
        print("DEBUG: filters_value_text exists: {0}".format(self.filters_value_text is not None))
    
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
        border_color = Color.FromArgb(255, 212, 184, 122)  # #D4B87A from Sheet Manager
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
        elif index == 2:  # CATEGORIES - #4CAF50 green
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
        
        # Template combo column
        template_col = DataGridComboBoxColumn()
        template_col.Header = "View Template"
        template_col.Width = DataGridLength(150)
        template_col.SelectedItemBinding = System.Windows.Data.Binding("view_template")
        self.template_items = self._get_all_templates()
        template_col.ItemsSource = self.template_items
        self.data_grid.Columns.Insert(2, template_col)
        
        # Detail combo column
        detail_col = DataGridComboBoxColumn()
        detail_col.Header = "Detail Level"
        detail_col.Width = DataGridLength(100)
        detail_col.SelectedItemBinding = System.Windows.Data.Binding("detail_level")
        detail_col.ItemsSource = ["Coarse", "Medium", "Fine"]
        self.data_grid.Columns.Insert(4, detail_col)
        
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
        text.Text = "(c) 2024 Dang Quoc Truong (DQT) - All Rights Reserved"
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
        
        print("DEBUG: Starting to load views...")
        
        for view in collector:
            if view.ViewType in [ViewType.ProjectBrowser, ViewType.SystemBrowser,
                                ViewType.Undefined, ViewType.Internal]:
                continue
            
            if view.IsTemplate:
                continue
            
            try:
                item = EnhancedViewItem(view, self.doc)
                self.all_views.append(item)
            except Exception as e:
                print("ERROR loading view {0}: {1}".format(view.Name, str(e)))
        
        print("DEBUG: Loaded {0} views".format(len(self.all_views)))
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
        print("DEBUG: _update_summary_cards called")
        
        # TOTAL - tổng views trong project
        if hasattr(self, 'total_value_text') and self.total_value_text is not None:
            total = len(self.all_views)
            self.total_value_text.Text = str(total)
            self.total_value_text.InvalidateVisual()  # Force refresh
            self.total_value_text.UpdateLayout()
            print("DEBUG: Updated TOTAL to {0}".format(total))
        else:
            print("DEBUG: total_value_text not available")
        
        # CATEGORIES - số loại view khác nhau
        if hasattr(self, 'types_value_text') and self.types_value_text is not None:
            types = set(v.view_type for v in self.all_views)
            type_count = len(types)
            self.types_value_text.Text = str(type_count)
            self.types_value_text.InvalidateVisual()
            self.types_value_text.UpdateLayout()
            print("DEBUG: Updated CATEGORIES to {0}".format(type_count))
        else:
            print("DEBUG: types_value_text not available")
        
        # FILTERS - always "Active"
        if hasattr(self, 'filters_value_text') and self.filters_value_text is not None:
            self.filters_value_text.Text = "Active"
            self.filters_value_text.InvalidateVisual()
            self.filters_value_text.UpdateLayout()
            print("DEBUG: Updated FILTERS to Active")
        else:
            print("DEBUG: filters_value_text not available")
        
        # Force update entire window
        self.InvalidateVisual()
        self.UpdateLayout()
        
        print("DEBUG: Filtered views count = {0}".format(len(self.filtered_views)))
    
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
        import_item.Header = "Import from Excel..."
        import_item.Click += self._on_import_excel
        menu.Items.Add(import_item)
        
        menu.PlacementTarget = sender
        menu.IsOpen = True
    
    def _on_export_excel(self, sender, args):
        """Export views to Excel using COM Interop"""
        excel_app = None
        wb = None
        
        try:
            dialog = SaveFileDialog()
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            dialog.Title = "Export Views to Excel"
            dialog.FileName = "Views_Export.xlsx"
            
            if dialog.ShowDialog() != DialogResult.OK:
                return
            
            filepath = dialog.FileName
            
            # Use COM Interop Excel
            excel_app = Excel.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            wb = excel_app.Workbooks.Add()
            ws = wb.Sheets[1]
            ws.Name = "Views"
            
            # Headers - BASE COLUMNS + CUSTOM PARAMETER COLUMNS
            base_headers = ["Element ID", "View Name", "Type", "View Template", "Scale", "Detail Level", 
                          "Title on Sheet", "Sheet Number", "Sheet Name", "On Sheets"]
            
            # Add custom parameter columns
            all_headers = base_headers[:]
            custom_col_bindings = []
            
            if self.custom_columns:
                for col_name in self.custom_columns.keys():
                    all_headers.append(col_name)
                    # Find binding name for this column
                    for i in range(len(self.custom_columns)):
                        binding_name = "param_{}".format(i)
                        if hasattr(self.filtered_views[0] if self.filtered_views else None, binding_name):
                            custom_col_bindings.append(binding_name)
                            break
            
            # Write headers
            for col, header in enumerate(all_headers, 1):
                cell = ws.Cells[1, col]
                cell.Value2 = header
                cell.Font.Bold = True
                if col == 1:
                    # Element ID column - special color
                    cell.Interior.Color = 0xD4E6A5  # Light green in BGR
                elif col > len(base_headers):
                    # Custom parameter columns - light yellow/cream
                    cell.Interior.Color = 0xCCF0FF  # Light cream/yellow in BGR
                else:
                    cell.Interior.Color = 0x88CCF0  # DQT Gold
            
            # Data
            row = 2
            for view_item in self.filtered_views:
                try:
                    # Helper function to safely convert values for Excel
                    def safe_value(val):
                        if val is None:
                            return ""
                        if isinstance(val, (int, float)):
                            return val
                        try:
                            return str(val) if val else ""
                        except:
                            return ""
                    
                    # Base columns with safe conversion
                    ws.Cells[row, 1].Value2 = safe_value(view_item.id.IntegerValue)
                    ws.Cells[row, 2].Value2 = safe_value(view_item.name)
                    ws.Cells[row, 3].Value2 = safe_value(view_item.view_type)
                    ws.Cells[row, 4].Value2 = safe_value(view_item.view_template)
                    ws.Cells[row, 5].Value2 = safe_value(view_item.scale)
                    ws.Cells[row, 6].Value2 = safe_value(view_item.detail_level)
                    ws.Cells[row, 7].Value2 = safe_value(view_item.title_on_sheet)
                    ws.Cells[row, 8].Value2 = safe_value(view_item.sheet_number)
                    ws.Cells[row, 9].Value2 = safe_value(view_item.sheet_name)
                    ws.Cells[row, 10].Value2 = safe_value(view_item.on_sheets)
                    
                    # Custom parameter columns
                    col_idx = 11
                    for i in range(len(self.custom_columns)):
                        binding_name = "param_{}".format(i)
                        if hasattr(view_item, binding_name):
                            value = getattr(view_item, binding_name, "")
                            ws.Cells[row, col_idx].Value2 = safe_value(value)
                            col_idx += 1
                    
                    row += 1
                except Exception as e:
                    # Skip problematic rows but continue
                    print("Error exporting row {}: {}".format(row, str(e)))
                    row += 1
                    continue
            
            # Hide Element ID column (but keep data)
            ws.Columns[1].Hidden = True
            
            ws.Columns.AutoFit()
            
            wb.SaveAs(filepath)
            wb.Close()
            excel_app.Quit()
            
            custom_msg = ""
            if self.custom_columns:
                custom_msg = "\n\nIncluding {} custom parameter column(s): {}".format(
                    len(self.custom_columns), 
                    ", ".join(self.custom_columns.keys())
                )
            
            MessageBox.Show(
                "Exported {0} views to:\n{1}\n\nTip: Don't delete column A (Element ID) - it's needed for import!{2}".format(
                    row-2, filepath, custom_msg),
                "Export Successful",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            )
        
        except Exception as e:
            if wb:
                try:
                    wb.Close(False)
                except:
                    pass
            if excel_app:
                try:
                    excel_app.Quit()
                except:
                    pass
            MessageBox.Show("Export error: {0}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_import_excel(self, sender, args):
        """Import views from Excel using COM Interop"""
        excel_app = None
        wb = None
        
        try:
            dialog = OpenFileDialog()
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            dialog.Title = "Import Views from Excel"
            
            if dialog.ShowDialog() != DialogResult.OK:
                return
            
            filepath = dialog.FileName
            
            excel_app = Excel.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            wb = excel_app.Workbooks.Open(filepath)
            ws = wb.Sheets[1]
            
            # Read headers to detect custom parameter columns
            header_row = 1
            col_idx = 1
            headers = []
            custom_param_cols = {}  # {col_index: param_name}
            
            while True:
                header = self._get_cell_value(ws.Cells[header_row, col_idx])
                if not header:
                    break
                headers.append(str(header))
                
                # Check if this is a custom parameter (after base columns)
                if col_idx > 10:  # After base 10 columns
                    custom_param_cols[col_idx] = str(header)
                
                col_idx += 1
                if col_idx > 100:  # Safety limit
                    break
            
            updates = []
            row = 2
            empty_rows = 0
            
            while empty_rows < 5:
                # Read Element ID from column A
                element_id = self._get_cell_value(ws.Cells[row, 1])
                view_name = self._get_cell_value(ws.Cells[row, 2])
                
                if not element_id and not view_name:
                    empty_rows += 1
                    row += 1
                    continue
                
                empty_rows = 0
                
                # Read base columns
                new_template = self._get_cell_value(ws.Cells[row, 4])
                new_scale = self._get_cell_value(ws.Cells[row, 5])
                new_detail = self._get_cell_value(ws.Cells[row, 6])
                
                update = {
                    'element_id': int(element_id) if element_id else None,
                    'view_name': str(view_name) if view_name else None,
                    'template': str(new_template) if new_template else None,
                    'scale': new_scale,
                    'detail_level': str(new_detail) if new_detail else None,
                    'title': None,
                    'custom_params': {}
                }
                
                # Read custom parameter values
                for col_idx, param_name in custom_param_cols.items():
                    param_value = self._get_cell_value(ws.Cells[row, col_idx])
                    if param_value:
                        update['custom_params'][param_name] = str(param_value)
                
                updates.append(update)
                
                row += 1
                if row > 10000:
                    break
            
            wb.Close()
            excel_app.Quit()
            
            if not updates:
                MessageBox.Show("No data found in Excel file", "No Data",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            custom_msg = ""
            if custom_param_cols:
                custom_msg = "\n\nIncluding {} custom parameter column(s)".format(len(custom_param_cols))
            
            result = MessageBox.Show(
                "Import {0} view updates from Excel?{1}".format(len(updates), custom_msg),
                "Confirm Import",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question
            )
            
            if result != System.Windows.MessageBoxResult.Yes:
                return
            
            self._apply_excel_updates(updates)
        
        except Exception as e:
            if wb:
                try:
                    wb.Close(False)
                except:
                    pass
            if excel_app:
                try:
                    excel_app.Quit()
                except:
                    pass
            MessageBox.Show("Import error: {0}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _get_cell_value(self, cell):
        """Safely get cell value from Excel"""
        try:
            val = cell.Value2
            return val if val is not None else None
        except:
            try:
                return cell.Value
            except:
                return None
    
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
                        view_elem = self.doc.GetElement(ElementId(update['element_id']))
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
                                if elem_id and elem_id.IntegerValue > 0:
                                    elem = self.doc.GetElement(elem_id)
                                    value = elem.Name if elem else str(elem_id.IntegerValue)
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
        
        except Exception as e:
            print("Right-click error: {}".format(str(e)))
    
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
                            if elem_id and elem_id.IntegerValue > 0:
                                elem = self.doc.GetElement(elem_id)
                                value = elem.Name if elem else str(elem_id.IntegerValue)
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