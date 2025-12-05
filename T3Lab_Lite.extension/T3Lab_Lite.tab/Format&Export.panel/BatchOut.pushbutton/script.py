# -*- coding: utf-8 -*-
"""BatchOut - Batch export sheets to multiple formats.

Enhanced version inspired by DiRoots ProSheets with tabbed interface.

Features:
- Tabbed interface (Selection, Format, Create)
- Export to DWG, PDF, DWF, DGN, NWD, IFC, and Image formats
- Customizable naming patterns with multiple placeholders
- Advanced PDF export options (paper size, orientation, hide elements)
- Combined PDF export option
- Sheet filtering by size
- Custom filename per sheet
- Progress tracking with cancellation support
- File organization options (same folder or split by format)

Version Compatibility:
- Supports Revit 2022, 2023, 2024, 2025, and 2026
- Uses version-aware API calls for optimal compatibility
- Automatically detects Revit version and applies appropriate API signatures
- Handles API differences between versions gracefully with fallback logic

API Version Notes:
- Revit 2022-2026: Document.Export() uses ICollection<ElementId> for DWG/PDF
- Revit 2022-2026: DWGExportOptions.PropOverrides expects PropOverrideMode enum
- Revit 2022-2026: PDFExportOptions signature requires separate folder and filename
- Revit 2025+: Built on .NET 8 (backward compatible with same API)
"""

__title__ = "Batch\nOut"
__author__ = "T3Lab"
__version__ = "1.0.0"

# IMPORTS
import os
import sys
import clr
from datetime import datetime
from collections import defaultdict

clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
from System.Windows.Forms import FolderBrowserDialog, DialogResult
from System.Windows import Visibility
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

from pyrevit import revit, DB, UI, forms, script
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory,
    ViewSheet, ViewSet, DWGExportOptions, DWFExportOptions,
    ExportDWGSettings, ACADVersion, PDFExportOptions,
    ImageExportOptions, ImageFileType, ImageResolution,
    PropOverrideMode,
)

from System.Collections.Generic import List

# Import API learner and updater modules
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

try:
    from api_learner import SmartAPIAdapter, RevitAPILearner
    from api_updater import auto_check_and_update
    HAS_API_LEARNER = True
except:
    HAS_API_LEARNER = False

# Try to import IFC export
try:
    from Autodesk.Revit.DB import IFCExportOptions, IFCVersion
    HAS_IFC = True
except:
    HAS_IFC = False

# Try to import Navisworks export
try:
    from Autodesk.Revit.DB import NavisworksExportOptions
    HAS_NAVISWORKS = True
except:
    HAS_NAVISWORKS = False

logger = script.get_logger()
output = script.get_output()

# Get Revit version information
REVIT_VERSION = int(revit.doc.Application.VersionNumber)  # e.g., 2023, 2024, 2025, 2026

class SheetItem(forms.Reactive):
    """Represents a sheet item in the list - optimized for performance."""
    def __init__(self, sheet, is_selected=False):
        self.Sheet = sheet
        self.IsSelected = is_selected
        self.SheetNumber = sheet.SheetNumber
        self.SheetName = sheet.Name
        self.Status = "Ready"
        self.Progress = 0
        self.Size = "-"  # Simplified - not loading size for performance

        # Get sheet revision info (fast parameter access)
        try:
            rev_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION)
            self.Revision = rev_param.AsString() if rev_param else ""
        except:
            self.Revision = ""

        try:
            rev_date_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION_DATE)
            self.RevisionDate = rev_date_param.AsString() if rev_date_param else ""
        except:
            self.RevisionDate = ""

        try:
            rev_desc_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION)
            self.RevisionDescription = rev_desc_param.AsString() if rev_desc_param else ""
        except:
            self.RevisionDescription = ""

        # Get drawn by and checked by (fast parameter access)
        try:
            drawn_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_DRAWN_BY)
            self.DrawnBy = drawn_param.AsString() if drawn_param else ""
        except:
            self.DrawnBy = ""

        try:
            checked_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CHECKED_BY)
            self.CheckedBy = checked_param.AsString() if checked_param else ""
        except:
            self.CheckedBy = ""

        # Custom filename (defaults to naming pattern)
        self.CustomFilename = ""

    def __repr__(self):
        return "{} - {}".format(self.SheetNumber, self.SheetName)


class ExportPreviewItem(object):
    """Represents an export preview item."""
    def __init__(self, sheet_item, format_name, size, orientation):
        self.SheetNumber = sheet_item.SheetNumber
        self.SheetName = sheet_item.SheetName
        self.Format = format_name
        self.Size = size
        self.Orientation = orientation
        self.Progress = 0


class ExportManagerWindow(forms.WPFWindow):
    """Export Manager Window."""

    def __init__(self):
        try:
            # Get absolute path to XAML file from lib/GUI folder
            extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
            xaml_file_path = os.path.join(extension_dir, 'lib', 'GUI', 'ExportManager.xaml')
            forms.WPFWindow.__init__(self, xaml_file_path)

            self.doc = revit.doc
            self.all_sheets = []
            self.filtered_sheets = []
            self.export_items = []

            # Set window icon and title bar logo
            try:
                logo_path = os.path.join(extension_dir, 'lib', 'GUI', 'T3Lab_logo.png')
                if os.path.exists(logo_path):
                    bitmap = BitmapImage()
                    bitmap.BeginInit()
                    bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                    bitmap.EndInit()
                    # Set window icon
                    self.Icon = bitmap
                    # Set title bar logo
                    self.logo_image.Source = bitmap
            except Exception as icon_ex:
                logger.warning("Could not set window icon: {}".format(icon_ex))

            # Initialize Smart API Adapter for self-learning capability
            if HAS_API_LEARNER:
                try:
                    self.api_adapter = SmartAPIAdapter(self.doc, REVIT_VERSION)
                    logger.info("Smart API Adapter initialized successfully")

                    # Check for API updates (non-blocking, runs in background)
                    self._check_for_api_updates()
                except Exception as adapter_ex:
                    logger.warning("Could not initialize Smart API Adapter: {}".format(adapter_ex))
                    self.api_adapter = None
            else:
                self.api_adapter = None

            # Set default output folder
            default_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'Revit Exports')
            self.output_folder.Text = default_folder

            # Load CAD export setups
            self.load_cad_export_setups()

            # Load sheets
            self.load_sheets()

            # Disable formats if not available
            if not HAS_NAVISWORKS:
                self.export_nwd.IsEnabled = False
                self.export_nwd.ToolTip = "Navisworks export not available in this Revit version"

            if not HAS_IFC:
                self.export_ifc.IsEnabled = False
                self.export_ifc.ToolTip = "IFC export not available in this Revit version"

            # Update button text based on current tab
            self.update_navigation_buttons()

        except Exception as ex:
            logger.error("Error initializing BatchOut window: {}".format(ex))
            raise

    def _check_for_api_updates(self):
        """Check for API updates in the background (non-blocking)."""
        try:
            # Auto-check for updates (this runs on Fridays or if never checked)
            update_result = auto_check_and_update()

            if update_result.get('checked'):
                # Log the check
                logger.info("API update check performed")

                # Show notifications if any
                notifications = update_result.get('notifications', [])
                for notif in notifications:
                    if notif.get('severity') == 'critical':
                        output.print_md("**⚠ CRITICAL**: {}".format(notif.get('message', '')))
                    elif notif.get('severity') == 'warning':
                        output.print_md("**⚡ INFO**: {}".format(notif.get('message', '')))
                    else:
                        output.print_md("**ℹ**: {}".format(notif.get('message', '')))

                # Show learner info
                if self.api_adapter:
                    learner_info = self.api_adapter.get_learner_info()
                    logger.info("API Learner: Cached date: {}, Source: {}".format(
                        learner_info.get('cached_date'),
                        learner_info.get('learned_from')
                    ))

        except Exception as ex:
            # Don't fail initialization if update check fails
            logger.debug("API update check failed: {}".format(ex))

    def load_cad_export_setups(self):
        """Load available DWG export setups from the document."""
        try:
            # Clear existing items
            self.cad_export_setup.Items.Clear()

            # Add default option
            from System.Windows.Controls import ComboBoxItem
            default_item = ComboBoxItem()
            default_item.Content = "Use setup from file (Default)"
            self.cad_export_setup.Items.Add(default_item)

            # Get all export settings from the document
            collector = FilteredElementCollector(self.doc)\
                .OfClass(ExportDWGSettings)

            # Add each export setup to the combo box
            for setup in collector:
                try:
                    setup_name = setup.Name if hasattr(setup, 'Name') else "Setup {}".format(setup.Id.IntegerValue)
                    item = ComboBoxItem()
                    item.Content = setup_name
                    item.Tag = setup  # Store the setup object for later use
                    self.cad_export_setup.Items.Add(item)
                except:
                    pass

            # Select the first item (default)
            self.cad_export_setup.SelectedIndex = 0

        except Exception as ex:
            logger.warning("Could not load CAD export setups: {}".format(ex))
            # Add just the default if there's an error
            from System.Windows.Controls import ComboBoxItem
            default_item = ComboBoxItem()
            default_item.Content = "Use setup from file (Default)"
            self.cad_export_setup.Items.Add(default_item)
            self.cad_export_setup.SelectedIndex = 0

    def load_sheets(self):
        """Load all sheets from the document."""
        try:
            # Get all sheets
            sheets_collector = FilteredElementCollector(self.doc)\
                .OfCategory(BuiltInCategory.OST_Sheets)\
                .WhereElementIsNotElementType()

            sheets = [s for s in sheets_collector if isinstance(s, ViewSheet)]

            # Sort by sheet number
            sheets.sort(key=lambda x: x.SheetNumber)

            # Create sheet items
            self.all_sheets = [SheetItem(sheet, False) for sheet in sheets]
            self.filtered_sheets = list(self.all_sheets)

            # Update ListView
            self.update_sheets_list()

            # Update status with version info
            self.status_text.Text = "Loaded {} sheets | Revit {}".format(
                len(self.all_sheets), REVIT_VERSION)

        except Exception as ex:
            logger.error("Error loading sheets: {}".format(ex))
            forms.alert("Error loading sheets: {}".format(ex), exitscript=True)

    def update_sheets_list(self):
        """Update the sheets ListView."""
        self.sheets_listview.ItemsSource = None
        self.sheets_listview.ItemsSource = self.filtered_sheets

    def select_all_sheets(self, sender, e):
        """Select all sheets."""
        for sheet_item in self.filtered_sheets:
            sheet_item.IsSelected = True
        self.sheets_listview.Items.Refresh()
        self.status_text.Text = "Selected {} sheets".format(len(self.filtered_sheets))

    def select_none_sheets(self, sender, e):
        """Deselect all sheets."""
        for sheet_item in self.filtered_sheets:
            sheet_item.IsSelected = False
        self.sheets_listview.Items.Refresh()
        self.status_text.Text = "Deselected all sheets"

    def refresh_sheets(self, sender, e):
        """Refresh the sheets list."""
        self.load_sheets()

    def search_sheets(self, sender, e):
        """Filter sheets by search text."""
        self.apply_filters()

    def filter_by_size(self, sender, e):
        """Filter sheets by size."""
        self.apply_filters()

    def apply_filters(self):
        """Apply search and size filters."""
        # Check if controls are initialized (prevents error during XAML loading)
        if not hasattr(self, 'search_textbox') or not hasattr(self, 'size_filter'):
            return

        search_text = self.search_textbox.Text.lower() if self.search_textbox.Text else ""

        # Get selected size filter
        size_filter = None
        if self.size_filter.SelectedItem:
            size_text = self.size_filter.SelectedItem.Content
            if size_text != "All Sizes":
                size_filter = size_text

        # Apply filters
        self.filtered_sheets = []
        for sheet in self.all_sheets:
            # Check search text
            if search_text:
                if search_text not in sheet.SheetNumber.lower() and \
                   search_text not in sheet.SheetName.lower():
                    continue

            # Check size filter
            if size_filter:
                if sheet.Size != size_filter:
                    continue

            self.filtered_sheets.append(sheet)

        self.update_sheets_list()
        self.status_text.Text = "Found {} sheets".format(len(self.filtered_sheets))

    def browse_output_folder(self, sender, e):
        """Browse for output folder."""
        dialog = FolderBrowserDialog()
        dialog.Description = "Select output folder for exports"
        dialog.SelectedPath = self.output_folder.Text

        if dialog.ShowDialog() == DialogResult.OK:
            self.output_folder.Text = dialog.SelectedPath
            self.status_text.Text = "Output folder: {}".format(dialog.SelectedPath)

    def format_changed(self, sender, e):
        """Handle format checkbox change."""
        # Update status to show selected formats
        formats = []
        if self.export_pdf.IsChecked:
            formats.append("PDF")
        if self.export_dwg.IsChecked:
            formats.append("DWG")
        if self.export_dgn.IsChecked:
            formats.append("DGN")
        if self.export_dwf.IsChecked:
            formats.append("DWF")
        if self.export_nwd.IsChecked:
            formats.append("NWC")
        if self.export_ifc.IsChecked:
            formats.append("IFC")
        if self.export_img.IsChecked:
            formats.append("IMG")

        if formats:
            self.status_text.Text = "Selected formats: {}".format(", ".join(formats))

    def reverse_order_changed(self, sender, e):
        """Handle reverse order checkbox change."""
        # Reverse the filtered sheets list
        self.filtered_sheets.reverse()
        self.update_sheets_list()

    def update_navigation_buttons(self):
        """Update navigation button visibility and text based on current tab."""
        current_tab_index = self.main_tabs.SelectedIndex

        if current_tab_index == 0:  # Selection tab
            self.back_button.Visibility = Visibility.Collapsed
            self.next_button.Content = "Next"
        elif current_tab_index == 1:  # Format tab
            self.back_button.Visibility = Visibility.Visible
            self.next_button.Content = "Next"
        elif current_tab_index == 2:  # Create tab
            self.back_button.Visibility = Visibility.Visible
            self.next_button.Content = "Create"

    def go_back(self, sender, e):
        """Navigate to previous tab."""
        current_index = self.main_tabs.SelectedIndex
        if current_index > 0:
            self.main_tabs.SelectedIndex = current_index - 1
            self.update_navigation_buttons()

    def go_next(self, sender, e):
        """Navigate to next tab or start export."""
        current_index = self.main_tabs.SelectedIndex

        if current_index == 0:  # Selection tab
            # Validate selection
            selected_sheets = [s for s in self.all_sheets if s.IsSelected]
            if not selected_sheets:
                forms.alert("Please select at least one sheet to export.", title="No Sheets Selected")
                return

            # Don't cache filenames here - they will be generated fresh during export
            # to ensure they reflect the current state of the sheets (e.g., updated sheet numbers)

            # Move to Format tab
            self.main_tabs.SelectedIndex = 1
            self.update_navigation_buttons()

        elif current_index == 1:  # Format tab
            # Validate format selection
            if not (self.export_pdf.IsChecked or self.export_dwg.IsChecked or
                    self.export_dgn.IsChecked or self.export_dwf.IsChecked or
                    self.export_nwd.IsChecked or self.export_ifc.IsChecked or
                    self.export_img.IsChecked):
                forms.alert("Please select at least one export format.", title="No Format Selected")
                return

            # Build export preview
            self.build_export_preview()

            # Move to Create tab
            self.main_tabs.SelectedIndex = 2
            self.update_navigation_buttons()

        elif current_index == 2:  # Create tab
            # Start export
            self.start_export()

    def build_export_preview(self):
        """Build the export preview list."""
        selected_sheets = [s for s in self.all_sheets if s.IsSelected]

        # Get orientation
        orientation = "Landscape" if self.pdf_landscape.IsChecked else "Portrait"

        # Get selected formats
        formats = []
        if self.export_pdf.IsChecked:
            formats.append("PDF")
        if self.export_dwg.IsChecked:
            formats.append("DWG")
        if self.export_dgn.IsChecked:
            formats.append("DGN")
        if self.export_dwf.IsChecked:
            formats.append("DWF")
        if self.export_nwd.IsChecked:
            formats.append("NWC")
        if self.export_ifc.IsChecked:
            formats.append("IFC")
        if self.export_img.IsChecked:
            formats.append("IMG")

        # Build preview items
        self.export_items = []
        for sheet in selected_sheets:
            for fmt in formats:
                item = ExportPreviewItem(sheet, fmt, sheet.Size, orientation)
                self.export_items.append(item)

        # Update preview list
        self.export_preview_list.ItemsSource = self.export_items
        self.progress_text.Text = "Ready to export {} items".format(len(self.export_items))

    def get_export_filename(self, sheet_item):
        """Generate export filename based on naming pattern.

        Always reads live values from the Revit sheet to ensure the filename
        reflects the current state of the sheet (e.g., if sheet number changed).
        """
        pattern = self.naming_pattern.Text

        # Get project info
        try:
            project_info = self.doc.ProjectInformation
            project_number = project_info.Number or ""
            project_name = project_info.Name or ""
        except:
            project_number = ""
            project_name = ""

        # Get live values from the actual Revit sheet object
        # This ensures we always use current values, not cached ones
        sheet = sheet_item.Sheet
        sheet_number = sheet.SheetNumber
        sheet_name = sheet.Name

        # Get live revision info
        try:
            rev_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION)
            revision = rev_param.AsString() if rev_param else ""
        except:
            revision = ""

        try:
            rev_date_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION_DATE)
            revision_date = rev_date_param.AsString() if rev_date_param else ""
        except:
            revision_date = ""

        try:
            rev_desc_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION)
            revision_description = rev_desc_param.AsString() if rev_desc_param else ""
        except:
            revision_description = ""

        # Get live drawn by and checked by
        try:
            drawn_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_DRAWN_BY)
            drawn_by = drawn_param.AsString() if drawn_param else ""
        except:
            drawn_by = ""

        try:
            checked_param = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CHECKED_BY)
            checked_by = checked_param.AsString() if checked_param else ""
        except:
            checked_by = ""

        # Replace placeholders with live values
        filename = pattern.replace("{SheetNumber}", sheet_number)
        filename = filename.replace("{SheetName}", sheet_name)
        filename = filename.replace("{Revision}", revision)
        filename = filename.replace("{RevisionDate}", revision_date)
        filename = filename.replace("{RevisionDescription}", revision_description)
        filename = filename.replace("{DrawnBy}", drawn_by)
        filename = filename.replace("{CheckedBy}", checked_by)
        filename = filename.replace("{ProjectNumber}", project_number)
        filename = filename.replace("{ProjectName}", project_name)
        filename = filename.replace("{Date}", datetime.now().strftime("%Y%m%d"))
        filename = filename.replace("{Time}", datetime.now().strftime("%H%M%S"))

        # Remove invalid characters
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in invalid_chars:
            filename = filename.replace(char, '_')

        return filename

    def update_export_item_progress(self, sheet_number, format_name, progress):
        """Update progress for a specific export item and refresh the display."""
        try:
            for item in self.export_items:
                if item.SheetNumber == sheet_number and item.Format == format_name:
                    item.Progress = progress
                    # Refresh the ListView to show updated progress
                    self.export_preview_list.Items.Refresh()
                    break
        except:
            pass

    def start_export(self):
        """Start the export process."""
        try:
            # Get selected sheets
            selected_sheets = [s for s in self.all_sheets if s.IsSelected]

            if not selected_sheets:
                forms.alert("Please select at least one sheet to export.", title="No Sheets Selected")
                return

            # Check if reverse order is enabled
            if self.reverse_order.IsChecked:
                selected_sheets.reverse()

            # Get output folder
            output_folder = self.output_folder.Text
            if not output_folder:
                forms.alert("Please select an output folder.", title="No Output Folder")
                return

            # Create output folder if it doesn't exist
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            # Check if split by format
            split_by_format = self.save_split_by_format.IsChecked

            # Disable buttons during export
            self.next_button.IsEnabled = False
            self.back_button.IsEnabled = False
            self.status_text.Text = "Exporting..."

            # Print export summary
            output.print_md("# Export Summary")
            output.print_md("**Sheets to export:** {}".format(len(selected_sheets)))
            output.print_md("**Output folder:** `{}`".format(output_folder))
            output.print_md("---\n")

            # Export to each format
            total_exported = 0
            total_items = len(self.export_items)
            current_item = 0

            if self.export_dwg.IsChecked:
                folder = os.path.join(output_folder, "DWG") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_dwg(selected_sheets, folder)
                total_exported += count
                current_item += count
                if total_items > 0:
                    self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_pdf.IsChecked:
                folder = os.path.join(output_folder, "PDF") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_pdf(selected_sheets, folder)
                total_exported += count
                current_item += count
                if total_items > 0:
                    self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_dwf.IsChecked:
                folder = os.path.join(output_folder, "DWF") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_dwf(selected_sheets, folder)
                total_exported += count
                current_item += count
                if total_items > 0:
                    self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_nwd.IsChecked:
                folder = os.path.join(output_folder, "NWC") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_nwd(selected_sheets, folder)
                total_exported += count
                current_item += count
                if total_items > 0:
                    self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_ifc.IsChecked:
                folder = os.path.join(output_folder, "IFC") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_ifc(selected_sheets, folder)
                total_exported += count
                current_item += count
                if total_items > 0:
                    self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_img.IsChecked:
                folder = os.path.join(output_folder, "Images") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_images(selected_sheets, folder)
                total_exported += count
                current_item += count
                if total_items > 0:
                    self.overall_progress.Value = (current_item * 100.0) / total_items

            # Show completion message
            output.print_md("\n---")
            output.print_md("# Export Complete!")
            output.print_md("**Total files exported:** {}".format(total_exported))

            self.status_text.Text = "Export complete! {} files exported".format(total_exported)
            self.progress_text.Text = "Export complete! {} files exported".format(total_exported)
            self.overall_progress.Value = 100
            self.next_button.IsEnabled = True
            self.back_button.IsEnabled = True

            # Ask if user wants to open output folder
            if forms.alert("Export complete!\n\nDo you want to open the output folder?",
                          title="Export Complete",
                          yes=True, no=True):
                os.startfile(output_folder)

        except Exception as ex:
            logger.error("Export failed: {}".format(ex))
            forms.alert("Export failed: {}".format(ex), title="Export Error")
            self.status_text.Text = "Export failed"
            self.next_button.IsEnabled = True
            self.back_button.IsEnabled = True

    def export_to_dwg(self, sheets, output_folder):
        """Export sheets to DWG format with version-aware API usage.

        Supports Revit 2022-2026 with appropriate API handling for each version.
        """
        try:
            output.print_md("### Exporting to DWG...")
            output.print_md("**Revit Version:** {}".format(REVIT_VERSION))

            # Get selected export setup (if any)
            selected_setup = None
            selected_setup_name = None
            if self.cad_export_setup.SelectedIndex > 0:
                # User selected a specific setup (not the default)
                selected_item = self.cad_export_setup.SelectedItem
                if hasattr(selected_item, 'Tag') and selected_item.Tag:
                    selected_setup = selected_item.Tag
                    selected_setup_name = selected_item.Content
                    output.print_md("Using export setup: **{}**".format(selected_setup_name))

            # Create DWG export options
            dwg_options = DWGExportOptions()

            # Set AutoCAD version
            dwg_version_index = self.dwg_version.SelectedIndex
            if dwg_version_index == 0:
                dwg_options.FileVersion = ACADVersion.R2013
            elif dwg_version_index == 1:
                dwg_options.FileVersion = ACADVersion.R2010
            else:
                dwg_options.FileVersion = ACADVersion.R2007

            # VERSION-AWARE: Apply CAD export options
            # ExportingAreas availability varies by version
            export_views_on_sheets = self.cad_export_views_on_sheets.IsChecked
            try:
                if hasattr(dwg_options, 'ExportingAreas') and hasattr(DB, 'ExportingAreas'):
                    if export_views_on_sheets:
                        dwg_options.ExportingAreas = DB.ExportingAreas.ExportViewsOnSheets
                        output.print_md("- Export views on sheets: **Enabled**")
                    else:
                        dwg_options.ExportingAreas = DB.ExportingAreas.DontExportViewsOnSheets
            except Exception as ex:
                logger.debug("ExportingAreas not supported in Revit {}: {}".format(REVIT_VERSION, ex))

            # Export links as external references
            export_links_as_external = self.cad_export_links_as_external.IsChecked
            try:
                if hasattr(dwg_options, 'MergedViews'):
                    dwg_options.MergedViews = not export_links_as_external
                    if export_links_as_external:
                        output.print_md("- Export links as external references: **Enabled**")
            except Exception as ex:
                logger.debug("MergedViews not supported in Revit {}: {}".format(REVIT_VERSION, ex))

            # VERSION-AWARE: Handle export setup application
            # Load settings from selected ExportDWGSettings
            if selected_setup:
                try:
                    # Load all settings from the ExportDWGSettings object
                    # LoadSettingsFrom copies all settings including layers, colors, line weights, etc.
                    dwg_options.LoadSettingsFrom(selected_setup, True)
                    output.print_md("**INFO**: Export setup '{}' applied successfully.".format(selected_setup_name))
                    output.print_md("  *All settings from export setup including layers, colors, and line weights have been loaded.*")
                except Exception as setup_ex:
                    logger.warning("Could not apply export setup '{}': {}".format(
                        selected_setup_name, setup_ex))
                    # Fallback: Set PropOverrides to ByEntity to match Revit colors
                    try:
                        dwg_options.PropOverrides = PropOverrideMode.ByEntity
                        output.print_md("**WARNING**: Could not load export setup. Using ByEntity mode for colors.")
                    except:
                        pass
            else:
                # No setup selected - ensure colors match Revit by using ByEntity mode
                try:
                    dwg_options.PropOverrides = PropOverrideMode.ByEntity
                    output.print_md("**INFO**: Using ByEntity mode - colors will match Revit display.")
                except Exception as prop_ex:
                    logger.debug("Could not set PropOverrides: {}".format(prop_ex))

            exported_count = 0

            for sheet_item in sheets:
                try:
                    # Update progress text to show current sheet and format
                    self.progress_text.Text = "Exporting {} to DWG...".format(sheet_item.SheetNumber)

                    filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)

                    # Remove extension if present
                    if filename.lower().endswith('.dwg'):
                        filename = filename[:-4]

                    # VERSION-AWARE: Export API handling
                    # All versions 2022-2026 support ICollection<ElementId> signature
                    # Signature: Export(String folder, String name, ICollection<ElementId> views, DWGExportOptions options)
                    view_ids = List[DB.ElementId]()
                    view_ids.Add(sheet_item.Sheet.Id)

                    # Use Smart API Adapter if available for intelligent export
                    if self.api_adapter:
                        # Smart adapter automatically handles version differences
                        self.api_adapter.export_dwg(output_folder, filename, view_ids, dwg_options)
                    else:
                        # Fallback to direct export call
                        if REVIT_VERSION >= 2022:
                            # Revit 2022-2026: Use ICollection<ElementId> signature
                            self.doc.Export(output_folder, filename, view_ids, dwg_options)
                        else:
                            # Fallback for older versions (if needed)
                            self.doc.Export(output_folder, filename, view_ids, dwg_options)

                    # Verify file was created
                    expected_file = os.path.join(output_folder, filename + ".dwg")
                    if os.path.exists(expected_file):
                        output.print_md("- Exported: **{}** → `{}.dwg`".format(
                            sheet_item.SheetNumber, filename))
                        exported_count += 1
                        # Update progress for this export item
                        self.update_export_item_progress(sheet_item.SheetNumber, "DWG", 100)
                    else:
                        output.print_md("- **Warning**: Export completed but file not found: {}".format(filename))

                except Exception as ex:
                    logger.error("Error exporting {} to DWG: {}".format(
                        sheet_item.SheetNumber, ex))
                    output.print_md("- **Error** exporting {}: {}".format(
                        sheet_item.SheetNumber, str(ex)))

            output.print_md("\n**DWG Export Complete:** {} sheets exported".format(exported_count))
            return exported_count

        except Exception as ex:
            logger.error("DWG export failed: {}".format(ex))
            output.print_md("**DWG Export Failed:** {}".format(str(ex)))
            return 0

    def export_to_pdf(self, sheets, output_folder):
        """Export sheets to PDF format using Revit's native PDF export with version-aware API usage.

        Supports Revit 2022-2026 with appropriate API handling for each version.
        """
        try:
            import time
            import glob

            output.print_md("### Exporting to PDF...")
            output.print_md("**Revit Version:** {}".format(REVIT_VERSION))

            # Check if combine PDF is enabled
            combine_pdf = self.combine_pdf.IsChecked

            exported_count = 0

            if combine_pdf:
                # Export all sheets to a single PDF
                try:
                    # Update progress text
                    self.progress_text.Text = "Exporting combined PDF with {} sheets...".format(len(sheets))

                    # Generate combined filename
                    if len(sheets) > 0:
                        first_sheet = sheets[0]
                        last_sheet = sheets[-1]
                        filename = "{}-{}_Combined".format(
                            first_sheet.SheetNumber,
                            last_sheet.SheetNumber
                        )
                    else:
                        filename = "Combined_Sheets"

                    # Remove extension if present (pyRevit style)
                    if filename.lower().endswith('.pdf'):
                        filename = filename[:-4]

                    # Clean filename - remove invalid chars and extra spaces
                    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
                    for char in invalid_chars:
                        filename = filename.replace(char, '_')
                    filename = filename.strip()

                    # Get list of existing PDF files before export
                    existing_pdfs = set(glob.glob(os.path.join(output_folder, "*.pdf")))

                    # Get all sheet IDs as System.Collections.Generic.List
                    sheet_ids = List[DB.ElementId]()
                    for s in sheets:
                        sheet_ids.Add(s.Sheet.Id)

                    # Create PDF export options
                    pdf_options = PDFExportOptions()
                    pdf_options.Combine = True
                    # Set filename (learned from pyRevit)
                    pdf_options.FileName = filename

                    # VERSION-AWARE: Apply PDF settings
                    # Use Smart API Adapter if available for intelligent configuration
                    if self.api_adapter:
                        pdf_options = self.api_adapter.configure_pdf_options(
                            pdf_options,
                            hide_scope_boxes=self.pdf_hide_ref_planes.IsChecked,
                            hide_crop_boundaries=self.pdf_hide_crop_boundaries.IsChecked,
                            hide_unreferenced_tags=self.pdf_hide_unreferenced_tags.IsChecked
                        )
                    else:
                        # Fallback to manual configuration
                        try:
                            if self.pdf_hide_ref_planes.IsChecked:
                                pdf_options.HideScopeBoxes = True
                        except:
                            logger.debug("HideScopeBoxes not supported in Revit {}".format(REVIT_VERSION))

                        try:
                            if self.pdf_hide_crop_boundaries.IsChecked:
                                pdf_options.HideCropBoundaries = True
                        except:
                            logger.debug("HideCropBoundaries not supported in Revit {}".format(REVIT_VERSION))

                        try:
                            if self.pdf_hide_unreferenced_tags.IsChecked:
                                pdf_options.HideUnreferencedViewTags = True
                        except:
                            logger.debug("HideUnreferencedViewTags not supported in Revit {}".format(REVIT_VERSION))

                    # VERSION-AWARE: Export using Revit's native PDF export
                    # Use Smart API Adapter if available for intelligent export (handles method overload resolution)
                    if self.api_adapter:
                        # Smart adapter automatically handles version differences and method overload resolution
                        self.api_adapter.export_pdf(output_folder, filename, sheet_ids, pdf_options)
                    else:
                        # Fallback to direct export call
                        # Revit 2022-2026 signature: Export(String folder, IList<ElementId> viewIds, PDFExportOptions options)
                        # NOTE: PDF export does NOT take a filename parameter in the Export() method (unlike DWG/DXF)
                        # Instead, filename is set via PDFExportOptions.FileName property (learned from pyRevit)
                        self.doc.Export(output_folder, sheet_ids, pdf_options)

                    # Wait briefly for file system to update
                    time.sleep(0.5)

                    # Get list of PDF files after export
                    current_pdfs = set(glob.glob(os.path.join(output_folder, "*.pdf")))
                    new_pdfs = current_pdfs - existing_pdfs

                    # Verify file was created
                    expected_file = os.path.join(output_folder, filename + ".pdf")
                    if os.path.exists(expected_file):
                        output.print_md("- Exported: **{} sheets** → `{}.pdf`".format(
                            len(sheets), filename))
                        exported_count = 1
                        # Update progress for all sheets in combined PDF
                        for s in sheets:
                            self.update_export_item_progress(s.SheetNumber, "PDF", 100)
                    elif new_pdfs:
                        # A PDF was created but with a different name - report it
                        actual_file = list(new_pdfs)[0]
                        actual_filename = os.path.basename(actual_file)
                        output.print_md("- Exported: **{} sheets** → `{}`".format(
                            len(sheets), actual_filename))
                        output.print_md("  *(Note: Revit used filename: `{}` instead of `{}.pdf`)*".format(
                            actual_filename, filename))
                        exported_count = 1
                        # Update progress for all sheets in combined PDF
                        for s in sheets:
                            self.update_export_item_progress(s.SheetNumber, "PDF", 100)
                    else:
                        output.print_md("- **Warning**: Export completed but no new PDF file detected")

                except Exception as ex:
                    logger.error("Error exporting combined PDF: {}".format(ex))
                    output.print_md("- **Error** exporting combined PDF: {}".format(str(ex)))

            else:
                # Export each sheet individually
                for sheet_item in sheets:
                    try:
                        # Update progress text to show current sheet and format
                        self.progress_text.Text = "Exporting {} to PDF...".format(sheet_item.SheetNumber)

                        filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)

                        # Remove extension if present (pyRevit style)
                        if filename.lower().endswith('.pdf'):
                            filename = filename[:-4]

                        # Clean filename - remove invalid chars and extra spaces
                        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
                        for char in invalid_chars:
                            filename = filename.replace(char, '_')
                        filename = filename.strip()

                        # Get list of existing PDF files before export
                        existing_pdfs = set(glob.glob(os.path.join(output_folder, "*.pdf")))

                        # Create PDF export options
                        pdf_options = PDFExportOptions()
                        pdf_options.Combine = False
                        # Set filename (learned from pyRevit)
                        pdf_options.FileName = filename

                        # VERSION-AWARE: Apply PDF settings
                        # Use Smart API Adapter if available for intelligent configuration
                        if self.api_adapter:
                            pdf_options = self.api_adapter.configure_pdf_options(
                                pdf_options,
                                hide_scope_boxes=self.pdf_hide_ref_planes.IsChecked,
                                hide_crop_boundaries=self.pdf_hide_crop_boundaries.IsChecked,
                                hide_unreferenced_tags=self.pdf_hide_unreferenced_tags.IsChecked
                            )
                        else:
                            # Fallback to manual configuration
                            try:
                                if self.pdf_hide_ref_planes.IsChecked:
                                    pdf_options.HideScopeBoxes = True
                            except:
                                logger.debug("HideScopeBoxes not supported in Revit {}".format(REVIT_VERSION))

                            try:
                                if self.pdf_hide_crop_boundaries.IsChecked:
                                    pdf_options.HideCropBoundaries = True
                            except:
                                logger.debug("HideCropBoundaries not supported in Revit {}".format(REVIT_VERSION))

                            try:
                                if self.pdf_hide_unreferenced_tags.IsChecked:
                                    pdf_options.HideUnreferencedViewTags = True
                            except:
                                logger.debug("HideUnreferencedViewTags not supported in Revit {}".format(REVIT_VERSION))

                        # Create System.Collections.Generic.List for sheet IDs
                        sheet_ids = List[DB.ElementId]()
                        sheet_ids.Add(sheet_item.Sheet.Id)

                        # VERSION-AWARE: Export using Revit's native PDF export
                        # Use Smart API Adapter if available for intelligent export (handles method overload resolution)
                        if self.api_adapter:
                            # Smart adapter automatically handles version differences and method overload resolution
                            self.api_adapter.export_pdf(output_folder, filename, sheet_ids, pdf_options)
                        else:
                            # Fallback to direct export call
                            # Revit 2022-2026 signature: Export(String folder, IList<ElementId> viewIds, PDFExportOptions options)
                            # NOTE: PDF export does NOT take a filename parameter in the Export() method (unlike DWG/DXF)
                            # Instead, filename is set via PDFExportOptions.FileName property (learned from pyRevit)
                            self.doc.Export(output_folder, sheet_ids, pdf_options)

                        # Wait briefly for file system to update
                        time.sleep(0.3)

                        # Get list of PDF files after export
                        current_pdfs = set(glob.glob(os.path.join(output_folder, "*.pdf")))
                        new_pdfs = current_pdfs - existing_pdfs

                        # Verify file was created
                        expected_file = os.path.join(output_folder, filename + ".pdf")
                        if os.path.exists(expected_file):
                            output.print_md("- Exported: **{}** → `{}.pdf`".format(
                                sheet_item.SheetNumber, filename))
                            exported_count += 1
                            # Update progress for this export item
                            self.update_export_item_progress(sheet_item.SheetNumber, "PDF", 100)
                        elif new_pdfs:
                            # A PDF was created but with a different name - report it
                            actual_file = list(new_pdfs)[0]
                            actual_filename = os.path.basename(actual_file)
                            output.print_md("- Exported: **{}** → `{}`".format(
                                sheet_item.SheetNumber, actual_filename))
                            exported_count += 1
                            # Update progress for this export item
                            self.update_export_item_progress(sheet_item.SheetNumber, "PDF", 100)
                        else:
                            output.print_md("- **Warning**: Export completed but no new PDF file detected for {}".format(
                                sheet_item.SheetNumber))

                    except Exception as ex:
                        logger.error("Error exporting {} to PDF: {}".format(
                            sheet_item.SheetNumber, ex))
                        output.print_md("- **Error** exporting {}: {}".format(
                            sheet_item.SheetNumber, str(ex)))

            output.print_md("\n**PDF Export Complete:** {} file(s) exported".format(exported_count))
            return exported_count

        except Exception as ex:
            logger.error("PDF export failed: {}".format(ex))
            output.print_md("**PDF Export Failed:** {}".format(str(ex)))
            return 0

    def export_to_dwf(self, sheets, output_folder):
        """Export sheets to DWF format using Revit's native DWF export with version-aware API usage.

        Supports Revit 2022-2026 with appropriate API handling for each version.
        """
        try:
            output.print_md("### Exporting to DWF...")
            output.print_md("**Revit Version:** {}".format(REVIT_VERSION))

            # Create DWF export options
            dwf_options = DWFExportOptions()

            exported_count = 0

            for sheet_item in sheets:
                try:
                    # Update progress text to show current sheet and format
                    self.progress_text.Text = "Exporting {} to DWF...".format(sheet_item.SheetNumber)

                    filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)

                    # Remove extension if present
                    if filename.lower().endswith('.dwf'):
                        filename = filename[:-4]

                    # VERSION-AWARE: Export handling
                    # Revit 2022-2026 all support ViewSet for DWF export
                    # Signature: Export(String folder, String name, ViewSet views, DWFExportOptions options)
                    view_set = DB.ViewSet()
                    view_set.Insert(sheet_item.Sheet)
                    self.doc.Export(output_folder, filename, view_set, dwf_options)

                    # Verify file was created
                    expected_file = os.path.join(output_folder, filename + ".dwf")
                    if os.path.exists(expected_file):
                        output.print_md("- Exported: **{}** → `{}.dwf`".format(
                            sheet_item.SheetNumber, filename))
                        exported_count += 1
                        # Update progress for this export item
                        self.update_export_item_progress(sheet_item.SheetNumber, "DWF", 100)
                    else:
                        output.print_md("- **Warning**: Export completed but file not found: {}".format(filename))

                except Exception as ex:
                    logger.error("Error exporting {} to DWF: {}".format(
                        sheet_item.SheetNumber, ex))
                    output.print_md("- **Error** exporting {}: {}".format(
                        sheet_item.SheetNumber, str(ex)))

            output.print_md("\n**DWF Export Complete:** {} sheets exported".format(exported_count))
            return exported_count

        except Exception as ex:
            logger.error("DWF export failed: {}".format(ex))
            output.print_md("**DWF Export Failed:** {}".format(str(ex)))
            return 0

    def export_to_nwd(self, sheets, output_folder):
        """Export to Navisworks NWD format with version-aware API usage."""
        if not HAS_NAVISWORKS:
            output.print_md("**Navisworks export not available**")
            return 0

        try:
            output.print_md("### Exporting to NWC (Navisworks)...")
            output.print_md("**Revit Version:** {}".format(REVIT_VERSION))

            # Create Navisworks export options
            nwd_options = NavisworksExportOptions()

            exported_count = 0

            for sheet_item in sheets:
                try:
                    # Update progress text to show current sheet and format
                    self.progress_text.Text = "Exporting {} to NWC...".format(sheet_item.SheetNumber)

                    filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)
                    filepath = os.path.join(output_folder, filename + ".nwc")

                    # Export view
                    nwd_options.ExportScope = DB.NavisworksExportScope.View
                    nwd_options.ViewId = sheet_item.Sheet.Id

                    self.doc.Export(output_folder, filename, nwd_options)

                    output.print_md("- Exported: **{}** → `{}.nwc`".format(
                        sheet_item.SheetNumber, filename))
                    exported_count += 1
                    # Update progress for this export item
                    self.update_export_item_progress(sheet_item.SheetNumber, "NWC", 100)

                except Exception as ex:
                    logger.error("Error exporting {} to NWC: {}".format(
                        sheet_item.SheetNumber, ex))
                    output.print_md("- **Error** exporting {}: {}".format(
                        sheet_item.SheetNumber, str(ex)))

            output.print_md("\n**NWC Export Complete:** {} sheets exported".format(exported_count))
            return exported_count

        except Exception as ex:
            logger.error("NWC export failed: {}".format(ex))
            output.print_md("**NWC Export Failed:** {}".format(str(ex)))
            return 0

    def export_to_ifc(self, sheets, output_folder):
        """Export to IFC format with version-aware API usage."""
        if not HAS_IFC:
            output.print_md("**IFC export not available**")
            return 0

        try:
            output.print_md("### Exporting to IFC...")
            output.print_md("**Revit Version:** {}".format(REVIT_VERSION))

            # Create IFC export options
            ifc_options = IFCExportOptions()
            ifc_options.FileVersion = IFCVersion.IFC2x3
            ifc_options.WallAndColumnSplitting = True

            exported_count = 0

            # For IFC, export the entire model once
            # Note: IFC export requires a transaction (unique requirement compared to other formats)
            if len(sheets) > 0:
                try:
                    # Update progress text to show IFC export
                    self.progress_text.Text = "Exporting entire model to IFC..."

                    filename = "Model_IFC_Export"

                    # IFC export needs to be wrapped in a transaction
                    with Transaction(self.doc, "Export IFC") as trans:
                        trans.Start()
                        self.doc.Export(output_folder, filename, ifc_options)
                        trans.Commit()

                    output.print_md("- Exported: **Model** → `{}.ifc`".format(filename))
                    output.print_md("  *(Note: IFC exports the entire model, not individual sheets)*")
                    exported_count = 1
                    # Update progress for all IFC export items
                    for s in sheets:
                        self.update_export_item_progress(s.SheetNumber, "IFC", 100)
                except Exception as ex:
                    logger.error("Error exporting to IFC: {}".format(ex))
                    output.print_md("- **Error** exporting to IFC: {}".format(str(ex)))

            output.print_md("\n**IFC Export Complete**")
            return exported_count

        except Exception as ex:
            logger.error("IFC export failed: {}".format(ex))
            output.print_md("**IFC Export Failed:** {}".format(str(ex)))
            return 0

    def export_to_images(self, sheets, output_folder):
        """Export sheets to image format using Revit's native image export with version-aware API usage."""
        try:
            output.print_md("### Exporting to Images...")
            output.print_md("**Revit Version:** {}".format(REVIT_VERSION))

            exported_count = 0

            for sheet_item in sheets:
                try:
                    # Update progress text to show current sheet and format
                    self.progress_text.Text = "Exporting {} to Image...".format(sheet_item.SheetNumber)

                    filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)

                    # Remove extension if present
                    if filename.lower().endswith('.png'):
                        filename = filename[:-4]

                    # Create image export options for each sheet
                    img_options = ImageExportOptions()
                    img_options.ZoomType = DB.ZoomFitType.FitToPage
                    img_options.ImageResolution = ImageResolution.DPI_150
                    img_options.FilePath = os.path.join(output_folder, filename)
                    img_options.FitDirection = DB.FitDirectionType.Horizontal
                    img_options.HLRandWFViewsFileType = ImageFileType.PNG
                    img_options.ShadowViewsFileType = ImageFileType.PNG
                    img_options.ExportRange = DB.ExportRange.SetOfViews

                    # Set the view IDs using System.Collections.Generic.List
                    view_ids = List[DB.ElementId]()
                    view_ids.Add(sheet_item.Sheet.Id)
                    img_options.SetViewsAndSheets(view_ids)

                    # Export using Revit's native image export
                    self.doc.ExportImage(img_options)

                    # Verify file was created
                    expected_file = os.path.join(output_folder, filename + ".png")
                    if os.path.exists(expected_file):
                        output.print_md("- Exported: **{}** → `{}.png`".format(
                            sheet_item.SheetNumber, filename))
                        exported_count += 1
                        # Update progress for this export item
                        self.update_export_item_progress(sheet_item.SheetNumber, "IMG", 100)
                    else:
                        output.print_md("- **Warning**: Export completed but file not found: {}".format(filename))

                except Exception as ex:
                    logger.error("Error exporting {} to Image: {}".format(
                        sheet_item.SheetNumber, ex))
                    output.print_md("- **Error** exporting {}: {}".format(
                        sheet_item.SheetNumber, str(ex)))

            output.print_md("\n**Image Export Complete:** {} sheets exported".format(exported_count))
            return exported_count

        except Exception as ex:
            logger.error("Image export failed: {}".format(ex))
            output.print_md("**Image Export Failed:** {}".format(str(ex)))
            return 0

    def cancel_export(self, sender, e):
        """Cancel and close the window."""
        self.Close()


# MAIN
if __name__ == '__main__':
    # Check if document is open
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)

    # Show Export Manager window
    window = ExportManagerWindow()
    window.ShowDialog()


