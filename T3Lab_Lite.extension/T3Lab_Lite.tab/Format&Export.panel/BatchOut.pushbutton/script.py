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
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

from pyrevit import revit, DB, UI, forms, script
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory,
    ViewSheet, ViewSet, DWGExportOptions, DWFExportOptions,
    ExportDWGSettings, ACADVersion, PDFExportOptions,
    ImageExportOptions, ImageFileType, ImageResolution,
)
from System.Collections.Generic import List

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

            # Set default output folder
            default_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'Revit Exports')
            self.output_folder.Text = default_folder

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

            # Update status
            self.status_text.Text = "Loaded {} sheets".format(len(self.all_sheets))

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

            # Generate custom filenames if not set
            for sheet in selected_sheets:
                if not sheet.CustomFilename:
                    sheet.CustomFilename = self.get_export_filename(sheet)

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
        """Generate export filename based on naming pattern."""
        pattern = self.naming_pattern.Text

        # Get project info
        try:
            project_info = self.doc.ProjectInformation
            project_number = project_info.Number or ""
            project_name = project_info.Name or ""
        except:
            project_number = ""
            project_name = ""

        # Replace placeholders
        filename = pattern.replace("{SheetNumber}", sheet_item.SheetNumber)
        filename = filename.replace("{SheetName}", sheet_item.SheetName)
        filename = filename.replace("{Revision}", sheet_item.Revision)
        filename = filename.replace("{RevisionDate}", sheet_item.RevisionDate)
        filename = filename.replace("{RevisionDescription}", sheet_item.RevisionDescription)
        filename = filename.replace("{DrawnBy}", sheet_item.DrawnBy)
        filename = filename.replace("{CheckedBy}", sheet_item.CheckedBy)
        filename = filename.replace("{ProjectNumber}", project_number)
        filename = filename.replace("{ProjectName}", project_name)
        filename = filename.replace("{Date}", datetime.now().strftime("%Y%m%d"))
        filename = filename.replace("{Time}", datetime.now().strftime("%H%M%S"))

        # Remove invalid characters
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in invalid_chars:
            filename = filename.replace(char, '_')

        return filename

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
                self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_pdf.IsChecked:
                folder = os.path.join(output_folder, "PDF") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_pdf(selected_sheets, folder)
                total_exported += count
                current_item += count
                self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_dwf.IsChecked:
                folder = os.path.join(output_folder, "DWF") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_dwf(selected_sheets, folder)
                total_exported += count
                current_item += count
                self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_nwd.IsChecked:
                folder = os.path.join(output_folder, "NWC") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_nwd(selected_sheets, folder)
                total_exported += count
                current_item += count
                self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_ifc.IsChecked:
                folder = os.path.join(output_folder, "IFC") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_ifc(selected_sheets, folder)
                total_exported += count
                current_item += count
                self.overall_progress.Value = (current_item * 100.0) / total_items

            if self.export_img.IsChecked:
                folder = os.path.join(output_folder, "Images") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_images(selected_sheets, folder)
                total_exported += count
                current_item += count
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
        """Export sheets to DWG format."""
        try:
            output.print_md("### Exporting to DWG...")

            # Create DWG export options
            dwg_options = DWGExportOptions()
            dwg_version_index = self.dwg_version.SelectedIndex
            if dwg_version_index == 0:
                dwg_options.FileVersion = ACADVersion.R2013
            elif dwg_version_index == 1:
                dwg_options.FileVersion = ACADVersion.R2010
            else:
                dwg_options.FileVersion = ACADVersion.R2007

            exported_count = 0

            for sheet_item in sheets:
                try:
                    filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)

                    # Remove extension if present
                    if filename.lower().endswith('.dwg'):
                        filename = filename[:-4]

                    # Create ViewSet for export (using Revit API ViewSet class)
                    view_set = DB.ViewSet()
                    view_set.Insert(sheet_item.Sheet)

                    # Export using Revit's native DWG export
                    self.doc.Export(output_folder, filename, dwg_options, view_set)

                    # Verify file was created
                    expected_file = os.path.join(output_folder, filename + ".dwg")
                    if os.path.exists(expected_file):
                        output.print_md("- Exported: **{}** → `{}.dwg`".format(
                            sheet_item.SheetNumber, filename))
                        exported_count += 1
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
        """Export sheets to PDF format using Revit's native PDF export."""
        try:
            output.print_md("### Exporting to PDF...")

            # Check if combine PDF is enabled
            combine_pdf = self.combine_pdf.IsChecked

            exported_count = 0

            if combine_pdf:
                # Export all sheets to a single PDF
                try:
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

                    # Remove extension if present
                    if filename.lower().endswith('.pdf'):
                        filename = filename[:-4]

                    # Get all sheet IDs as System.Collections.Generic.List
                    sheet_ids = List[DB.ElementId]()
                    for s in sheets:
                        sheet_ids.Add(s.Sheet.Id)

                    # Create PDF export options
                    pdf_options = PDFExportOptions()
                    pdf_options.Combine = True

                    # Apply PDF settings
                    if self.pdf_hide_ref_planes.IsChecked:
                        pdf_options.HideScopeBoxes = True
                    if self.pdf_hide_crop_boundaries.IsChecked:
                        pdf_options.HideCropBoundaries = True
                    if self.pdf_hide_unreferenced_tags.IsChecked:
                        pdf_options.HideUnreferencedViewTags = True

                    # Export using Revit's native PDF export
                    # Pass filename directly to Export method, not via FileName property
                    self.doc.Export(output_folder, filename, sheet_ids, pdf_options)

                    # Verify file was created
                    expected_file = os.path.join(output_folder, filename + ".pdf")
                    if os.path.exists(expected_file):
                        output.print_md("- Exported: **{} sheets** → `{}.pdf`".format(
                            len(sheets), filename))
                        exported_count = 1
                    else:
                        output.print_md("- **Warning**: Export completed but file not found: {}".format(filename))

                except Exception as ex:
                    logger.error("Error exporting combined PDF: {}".format(ex))
                    output.print_md("- **Error** exporting combined PDF: {}".format(str(ex)))

            else:
                # Export each sheet individually
                for sheet_item in sheets:
                    try:
                        filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)

                        # Remove extension if present
                        if filename.lower().endswith('.pdf'):
                            filename = filename[:-4]

                        # Create PDF export options
                        pdf_options = PDFExportOptions()
                        pdf_options.Combine = False

                        # Apply PDF settings
                        if self.pdf_hide_ref_planes.IsChecked:
                            pdf_options.HideScopeBoxes = True
                        if self.pdf_hide_crop_boundaries.IsChecked:
                            pdf_options.HideCropBoundaries = True
                        if self.pdf_hide_unreferenced_tags.IsChecked:
                            pdf_options.HideUnreferencedViewTags = True

                        # Create System.Collections.Generic.List for sheet IDs
                        sheet_ids = List[DB.ElementId]()
                        sheet_ids.Add(sheet_item.Sheet.Id)

                        # Export using Revit's native PDF export
                        # Pass filename directly to Export method, not via FileName property
                        self.doc.Export(output_folder, filename, sheet_ids, pdf_options)

                        # Verify file was created
                        expected_file = os.path.join(output_folder, filename + ".pdf")
                        if os.path.exists(expected_file):
                            output.print_md("- Exported: **{}** → `{}.pdf`".format(
                                sheet_item.SheetNumber, filename))
                            exported_count += 1
                        else:
                            output.print_md("- **Warning**: Export completed but file not found: {}".format(filename))

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
        """Export sheets to DWF format using Revit's native DWF export."""
        try:
            output.print_md("### Exporting to DWF...")

            # Create DWF export options
            dwf_options = DWFExportOptions()

            exported_count = 0

            for sheet_item in sheets:
                try:
                    filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)

                    # Remove extension if present
                    if filename.lower().endswith('.dwf'):
                        filename = filename[:-4]

                    # Create ViewSet for export
                    view_set = DB.ViewSet()
                    view_set.Insert(sheet_item.Sheet)

                    # Export using Revit's native DWF export
                    self.doc.Export(output_folder, filename, dwf_options, view_set)

                    # Verify file was created
                    expected_file = os.path.join(output_folder, filename + ".dwf")
                    if os.path.exists(expected_file):
                        output.print_md("- Exported: **{}** → `{}.dwf`".format(
                            sheet_item.SheetNumber, filename))
                        exported_count += 1
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
        """Export to Navisworks NWD format."""
        if not HAS_NAVISWORKS:
            output.print_md("**Navisworks export not available**")
            return 0

        try:
            output.print_md("### Exporting to NWC (Navisworks)...")

            # Create Navisworks export options
            nwd_options = NavisworksExportOptions()

            exported_count = 0

            for sheet_item in sheets:
                try:
                    filename = sheet_item.CustomFilename or self.get_export_filename(sheet_item)
                    filepath = os.path.join(output_folder, filename + ".nwc")

                    # Export view
                    nwd_options.ExportScope = DB.NavisworksExportScope.View
                    nwd_options.ViewId = sheet_item.Sheet.Id

                    self.doc.Export(output_folder, filename, nwd_options)

                    output.print_md("- Exported: **{}** → `{}.nwc`".format(
                        sheet_item.SheetNumber, filename))
                    exported_count += 1

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
        """Export to IFC format."""
        if not HAS_IFC:
            output.print_md("**IFC export not available**")
            return 0

        try:
            output.print_md("### Exporting to IFC...")

            # Create IFC export options
            ifc_options = IFCExportOptions()
            ifc_options.FileVersion = IFCVersion.IFC2x3
            ifc_options.WallAndColumnSplitting = True

            exported_count = 0

            # For IFC, export the entire model once
            if len(sheets) > 0:
                try:
                    filename = "Model_IFC_Export"
                    self.doc.Export(output_folder, filename, ifc_options)
                    output.print_md("- Exported: **Model** → `{}.ifc`".format(filename))
                    output.print_md("  *(Note: IFC exports the entire model, not individual sheets)*")
                    exported_count = 1
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
        """Export sheets to image format using Revit's native image export."""
        try:
            output.print_md("### Exporting to Images...")

            exported_count = 0

            for sheet_item in sheets:
                try:
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
