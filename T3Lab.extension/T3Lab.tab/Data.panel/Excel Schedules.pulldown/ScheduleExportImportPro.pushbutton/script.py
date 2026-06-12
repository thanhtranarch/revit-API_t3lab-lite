# -*- coding: utf-8 -*-
"""
Schedule Link Pro - Export/Import Schedule to Excel
Copyright (c) 2025 by Dang Quoc Truong (DQT)

DiRoots-style: Read data directly from elements, not from schedule cells
"""

__title__ = "Schedule\nLink Pro"
__doc__ = "Export Schedule to Excel, edit and import back"
__author__ = "Dang Quoc Truong (DQT)"

import os
import traceback

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

# Try to load Excel Interop - if fails, use zipfile+XML approach
Excel = None
USE_ZIPFILE = False
USE_COM = False
Excel = None

try:
    # Method 1: Direct reference (works if Office PIA installed)
    clr.AddReference('Microsoft.Office.Interop.Excel')
    import Microsoft.Office.Interop.Excel as Excel
except:
    try:
        # Method 2: Load from GAC with full name (Office 2013/2016)
        clr.AddReference('Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
        import Microsoft.Office.Interop.Excel as Excel
    except:
        try:
            # Method 3: Try Office 2019/365 version
            clr.AddReference('Microsoft.Office.Interop.Excel, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
            import Microsoft.Office.Interop.Excel as Excel
        except:
            # Method 4: Fallback to zipfile (most reliable for Revit 2026)
            USE_ZIPFILE = True
            import zipfile
            try:
                from xml.etree import ElementTree as ET
            except:
                import xml.etree.ElementTree as ET

from System.Windows.Forms import (
    OpenFileDialog, SaveFileDialog, DialogResult, MessageBox,
    MessageBoxButtons, MessageBoxIcon
)
from System.Windows import (
    Window, Thickness, TextWrapping, HorizontalAlignment, 
    VerticalAlignment, GridLength, Visibility, FontWeights
)
from System.Windows.Controls import (
    StackPanel, DockPanel, Grid, RowDefinition, ColumnDefinition,
    Button, TextBlock, ComboBox, ComboBoxItem, Border, ScrollViewer,
    ProgressBar, CheckBox, ListBox, ListBoxItem, SelectionMode
)
from System.Windows.Media import SolidColorBrush
from System.Windows.Media import Color as WpfColor

import System
from System import Activator, Type

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ScheduleFieldType,
    Transaction, ElementId, BuiltInParameter, StorageType, Element
)

from pyrevit import revit

doc = revit.doc


def get_element_id_value(eid):
    """Get integer value from ElementId - compatible with Revit 2024, 2025, and 2026"""
    if eid is None:
        return -1
    try:
        # Revit 2026+ uses .Value
        return eid.Value
    except AttributeError:
        try:
            # Revit 2024/2025 uses .IntegerValue
            return eid.IntegerValue
        except:
            return -1

# =============================================================================
# CONSTANTS
# =============================================================================
PRIMARY = "#FFE6A800"
BACKGROUND = "#FFF8FAFC"  
WHITE = "#FFFFFFFF"
GRAY = "#FF808080"
RED = "#FFFF0000"
GREEN = "#FF008000"
COPYRIGHT = "Copyright (c) 2025 by Dang Quoc Truong (DQT)"

DEBUG_LOG = []
UPDATE_LOG = []

def log_debug(msg):
    global DEBUG_LOG
    DEBUG_LOG.append(str(msg))

def log_update(msg):
    global UPDATE_LOG
    UPDATE_LOG.append(str(msg))

def brush(hex_color):
    h = hex_color.lstrip('#')
    if len(h) == 8:
        a, r, g, b = [int(h[i:i+2], 16) for i in range(0, 8, 2)]
    else:
        a = 255
        r, g, b = [int(h[i:i+2], 16) for i in range(0, 6, 2)]
    return SolidColorBrush(WpfColor.FromArgb(a, r, g, b))

def get_cell_value(cell):
    try:
        val = cell.Value2
        return val if val is not None else None
    except:
        try:
            return cell.Value
        except:
            return None

def safe_str(val):
    if val is None:
        return ""
    try:
        # Handle float that is actually an integer (1.0 -> "1")
        if isinstance(val, float):
            if val == int(val):
                return str(int(val))
            else:
                return str(val)
        return str(val).strip()
    except:
        return ""

def safe_int(val):
    if val is None:
        return None
    try:
        if isinstance(val, (int, float)):
            return int(val)
        s = str(val).strip()
        return int(float(s)) if s else None
    except:
        return None


# =============================================================================
# SCHEDULE EXTRACTOR - DiRoots Style
# =============================================================================
class ScheduleExtractor:
    
    @staticmethod
    def get_all_schedules():
        result = {}
        for sch in FilteredElementCollector(doc).OfClass(ViewSchedule):
            if not sch.IsTitleblockRevisionSchedule and not sch.IsInternalKeynoteSchedule:
                result[sch.Name] = sch
        return result
    
    @staticmethod
    def extract_from_cells(schedule):
        """Extract data directly from schedule cells with formatting - keeps formatting exactly like Revit displays"""
        global DEBUG_LOG
        DEBUG_LOG = []
        
        if not schedule:
            return None
        
        try:
            from Autodesk.Revit.DB import SectionType, Color
            
            table = schedule.GetTableData()
            body = table.GetSectionData(SectionType.Body)
            
            num_rows = body.NumberOfRows
            num_cols = body.NumberOfColumns
            
            log_debug("Schedule cells: {} rows x {} cols".format(num_rows, num_cols))
            
            if num_rows == 0 or num_cols == 0:
                return None
            
            # Get column widths
            column_widths = []
            for c in range(num_cols):
                try:
                    # GetColumnWidth returns width in feet, convert to approximate Excel units
                    width_feet = body.GetColumnWidth(c)
                    # 1 foot ≈ 72 points, Excel column width is ~7 pixels per unit
                    width_excel = max(8, int(width_feet * 72 / 7 * 12))  # Scale factor
                    column_widths.append(width_excel)
                except:
                    column_widths.append(15)  # Default width
            
            log_debug("Column widths: {}".format(column_widths))
            
            # Read header row (row 0) with formatting
            headers = []
            header_formats = []
            for c in range(num_cols):
                try:
                    header_text = schedule.GetCellText(SectionType.Body, 0, c) or ""
                    headers.append(header_text.strip())
                    
                    # Try to get cell style/color
                    try:
                        style = body.GetTableCellStyle(0, c)
                        bg_color = style.BackgroundColor
                        if bg_color and bg_color.IsValid:
                            header_formats.append({
                                'bg_color': (bg_color.Red, bg_color.Green, bg_color.Blue),
                                'bold': True
                            })
                        else:
                            header_formats.append({'bg_color': (200, 200, 200), 'bold': True})
                    except:
                        header_formats.append({'bg_color': (200, 200, 200), 'bold': True})
                except:
                    headers.append("")
                    header_formats.append({'bg_color': (200, 200, 200), 'bold': True})
            
            log_debug("Headers from cells: {}".format(headers))
            
            # Create fields info (read-only for cell export)
            fields = []
            for h in headers:
                fields.append({
                    'name': h,
                    'param_id': None,
                    'can_edit': False,
                    'field_type': None,
                    'field_type_str': 'Cell'
                })
            
            # Read data rows (starting from row 1) with formatting
            rows = []
            row_formats = []
            for r in range(1, num_rows):
                row_data = []
                row_fmt = []
                for c in range(num_cols):
                    try:
                        cell_text = schedule.GetCellText(SectionType.Body, r, c) or ""
                        row_data.append(cell_text)
                        
                        # Try to get cell style/color
                        try:
                            style = body.GetTableCellStyle(r, c)
                            bg_color = style.BackgroundColor
                            if bg_color and bg_color.IsValid:
                                row_fmt.append({
                                    'bg_color': (bg_color.Red, bg_color.Green, bg_color.Blue)
                                })
                            else:
                                row_fmt.append({'bg_color': None})
                        except:
                            row_fmt.append({'bg_color': None})
                    except:
                        row_data.append("")
                        row_fmt.append({'bg_color': None})
                
                # Include row if it has any data
                if any(v.strip() for v in row_data):
                    rows.append(row_data)
                    row_formats.append(row_fmt)
            
            log_debug("Data rows: {}".format(len(rows)))
            
            # Debug first few rows
            for i in range(min(3, len(rows))):
                log_debug("Row {}: {}".format(i, rows[i][:5] if len(rows[i]) > 5 else rows[i]))
            
            return {
                'headers': headers,
                'header_formats': header_formats,
                'fields': fields,
                'rows': rows,
                'row_formats': row_formats,
                'column_widths': column_widths,
                'element_ids': [],  # No element IDs in cell mode
                'schedule_name': schedule.Name,
                'num_cols': len(headers),
                'valid_element_count': len(rows),
                'total_rows': len(rows),
                'from_cells': True
            }
            
        except Exception as e:
            log_debug("Extract from cells error: " + str(e))
            log_debug(traceback.format_exc())
            return None
    
    @staticmethod
    def extract(schedule):
        global DEBUG_LOG
        DEBUG_LOG = []
        
        if not schedule:
            return None
        
        try:
            definition = schedule.Definition
            
            # Get fields info from schedule definition
            fields = []
            for i in range(definition.GetFieldCount()):
                f = definition.GetField(i)
                if not f.IsHidden:
                    param_id = None
                    try:
                        pid = get_element_id_value(f.ParameterId)
                        if f.ParameterId and pid != -1:
                            param_id = pid
                    except:
                        pass
                    
                    field_type = f.FieldType
                    field_name = f.GetName()
                    
                    is_editable = field_type not in [
                        ScheduleFieldType.Formula, 
                        ScheduleFieldType.Count,
                        ScheduleFieldType.ElementType
                    ]
                    
                    fields.append({
                        'name': field_name,
                        'param_id': param_id,
                        'can_edit': is_editable,
                        'field_type': field_type,
                        'field_type_str': str(field_type)
                    })
            
            headers = [f['name'] for f in fields]
            log_debug("Fields: {}".format(headers))
            log_debug("Field count: {}".format(len(fields)))
            
            for f in fields:
                log_debug("  {} -> param_id={}, field_type={}".format(
                    f['name'], f['param_id'], f['field_type_str']))
            
            # Get elements from schedule view
            collector = FilteredElementCollector(doc, schedule.Id)
            elements = list(collector.ToElements())
            log_debug("Elements in schedule: {}".format(len(elements)))
            
            # Build data rows directly from elements
            rows = []
            element_ids = []
            
            for elem in elements:
                eid = get_element_id_value(elem.Id)
                element_ids.append(eid)
                
                row_data = []
                for field in fields:
                    value = ScheduleExtractor._get_field_value(elem, field)
                    row_data.append(value)
                
                rows.append(row_data)
            
            log_debug("Rows built: {}".format(len(rows)))
            
            for i in range(min(3, len(rows))):
                eid = element_ids[i]
                log_debug("Row {}: ID={}, Data={}".format(i, eid, rows[i]))
            
            return {
                'headers': headers,
                'fields': fields,
                'rows': rows,
                'element_ids': element_ids,
                'schedule_name': schedule.Name,
                'num_cols': len(fields),
                'valid_element_count': len(elements),
                'total_rows': len(rows)
            }
            
        except Exception as e:
            log_debug("Extract error: " + str(e))
            log_debug(traceback.format_exc())
            return None
    
    @staticmethod
    def _get_element_name(elem):
        """Get element name - works in IronPython"""
        try:
            # Method 1: Direct property access
            if hasattr(elem, 'Name'):
                name = elem.Name
                if name:
                    return name
        except:
            pass
        
        try:
            # Method 2: Use Element.Name.GetValue() for IronPython
            name = Element.Name.GetValue(elem)
            if name:
                return name
        except:
            pass
        
        return ""
    
    @staticmethod
    def _get_field_value(elem, field):
        """Get field value from element"""
        try:
            field_name = field['name']
            param_id = field['param_id']
            field_type = field['field_type']
            field_name_lower = field_name.lower().strip()
            
            # Handle Count field type
            if field_type == ScheduleFieldType.Count:
                return "1"
            
            # Get element type
            type_id = elem.GetTypeId()
            elem_type = doc.GetElement(type_id) if type_id and get_element_id_value(type_id) != -1 else None
            
            # Handle Family field (param_id -1002051)
            if param_id == -1002051 or field_name_lower == 'family':
                if elem_type:
                    # Try FamilyName property
                    if hasattr(elem_type, 'FamilyName'):
                        fname = elem_type.FamilyName
                        if fname:
                            return fname
                    # Try Family.Name
                    if hasattr(elem_type, 'Family') and elem_type.Family:
                        fname = ScheduleExtractor._get_element_name(elem_type.Family)
                        if fname:
                            return fname
                return ""
            
            # Handle Type field (param_id -1002050)
            if param_id == -1002050 or field_name_lower == 'type':
                if elem_type:
                    # Get Type name using _get_element_name helper
                    type_name = ScheduleExtractor._get_element_name(elem_type)
                    if type_name:
                        return type_name
                return ""
            
            # Handle ElementType field type (parameters from Type)
            if field_type == ScheduleFieldType.ElementType:
                if elem_type:
                    param = None
                    if param_id and param_id < 0:
                        try:
                            bip = BuiltInParameter(param_id)
                            param = elem_type.get_Parameter(bip)
                        except:
                            pass
                    if not param:
                        param = elem_type.LookupParameter(field_name)
                    
                    if param:
                        return ScheduleExtractor._get_param_value(param)
                return ""
            
            # Regular parameter - try instance first
            param = None
            
            if param_id and param_id < 0:
                try:
                    bip = BuiltInParameter(param_id)
                    param = elem.get_Parameter(bip)
                except:
                    pass
            
            if not param:
                param = elem.LookupParameter(field_name)
            
            # If not found on instance, try type
            if not param and elem_type:
                if param_id and param_id < 0:
                    try:
                        bip = BuiltInParameter(param_id)
                        param = elem_type.get_Parameter(bip)
                    except:
                        pass
                if not param:
                    param = elem_type.LookupParameter(field_name)
            
            if not param:
                return ""
            
            return ScheduleExtractor._get_param_value(param)
            
        except Exception as e:
            return ""
    
    @staticmethod
    def _get_param_value(param):
        """Extract value from parameter - returns display string with units"""
        try:
            # First try AsValueString - returns formatted value with units like Revit displays
            value_string = param.AsValueString()
            if value_string:
                return value_string
            
            # Fallback to raw value conversion
            storage = param.StorageType
            
            if storage == StorageType.String:
                return param.AsString() or ""
            elif storage == StorageType.Integer:
                v = param.AsInteger()
                return str(v) if v is not None else ""
            elif storage == StorageType.Double:
                v = param.AsDouble()
                if v is not None:
                    return str(round(v, 6))
                return ""
            elif storage == StorageType.ElementId:
                eid = param.AsElementId()
                if eid and get_element_id_value(eid) != -1:
                    ref_elem = doc.GetElement(eid)
                    if ref_elem:
                        return ScheduleExtractor._get_element_name(ref_elem)
                return ""
            return ""
        except:
            return ""


# =============================================================================
# EXCEL MANAGER - Uses zipfile+XML when Interop not available
# =============================================================================

def _col_letter(col_num):
    """Convert column number (1-based) to Excel letter (A, B, ..., Z, AA, AB, ...)"""
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(col_num % 26 + ord('A')) + result
        col_num //= 26
    return result


def _create_xlsx_zipfile(filepath, sheets_data):
    """Create XLSX file using zipfile + XML with formatting support"""
    import zipfile
    try:
        from xml.etree import ElementTree as ET
    except:
        import xml.etree.ElementTree as ET
    
    def to_xml_string(element):
        """Convert element to XML string - IronPython compatible"""
        try:
            return ET.tostring(element, encoding='unicode')
        except (LookupError, TypeError):
            try:
                xml_bytes = ET.tostring(element, encoding='utf-8')
                result = xml_bytes.decode('utf-8') if isinstance(xml_bytes, bytes) else xml_bytes
                if result.startswith('<?xml'):
                    idx = result.find('?>')
                    if idx > 0:
                        result = result[idx+2:].lstrip()
                return result
            except:
                xml_bytes = ET.tostring(element)
                if isinstance(xml_bytes, bytes):
                    return xml_bytes.decode('utf-8')
                return str(xml_bytes)
    
    def rgb_to_hex(r, g, b):
        """Convert RGB to hex string for Excel"""
        return "FF{:02X}{:02X}{:02X}".format(int(r), int(g), int(b))
    
    # Namespace definitions
    NS = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
        'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
        'ws': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
        'ss': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
        'st': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
    }
    
    # Collect unique colors for styles
    colors_set = set()
    colors_set.add((200, 200, 200))  # Default header gray
    
    for sheet_data in sheets_data:
        for fmt in sheet_data.get('header_formats', []):
            if fmt.get('bg_color'):
                colors_set.add(tuple(fmt['bg_color']))
        for row_fmt in sheet_data.get('row_formats', []):
            for fmt in row_fmt:
                if fmt.get('bg_color'):
                    colors_set.add(tuple(fmt['bg_color']))
    
    colors_list = list(colors_set)
    color_to_fill_id = {}
    for idx, color in enumerate(colors_list):
        color_to_fill_id[color] = idx + 2  # +2 because 0=none, 1=gray125
    
    # Collect all unique strings
    shared_strings = []
    string_index = {}
    
    def get_string_index(s):
        s = str(s) if s is not None else ""
        if s not in string_index:
            string_index[s] = len(shared_strings)
            shared_strings.append(s)
        return string_index[s]
    
    # Pre-process all sheets to collect strings
    for sheet_data in sheets_data:
        is_from_cells = sheet_data.get('from_cells', False)
        for header in sheet_data.get('headers', []):
            get_string_index(str(header) if header else "")
        for row in sheet_data.get('rows', []):
            for cell in row:
                val_str = str(cell) if cell is not None else ""
                if is_from_cells:
                    get_string_index(val_str)
                else:
                    try:
                        if val_str.strip():
                            float(val_str)
                    except:
                        get_string_index(val_str)
        get_string_index("")
    
    # Create styles XML with fills
    def create_styles_xml():
        styleSheet = ET.Element('styleSheet', {'xmlns': NS['main']})
        
        # Fonts
        fonts = ET.SubElement(styleSheet, 'fonts', {'count': '2'})
        # Font 0 - normal
        font0 = ET.SubElement(fonts, 'font')
        ET.SubElement(font0, 'sz', {'val': '11'})
        ET.SubElement(font0, 'name', {'val': 'Calibri'})
        # Font 1 - bold
        font1 = ET.SubElement(fonts, 'font')
        ET.SubElement(font1, 'b')
        ET.SubElement(font1, 'sz', {'val': '11'})
        ET.SubElement(font1, 'name', {'val': 'Calibri'})
        
        # Fills
        fills = ET.SubElement(styleSheet, 'fills', {'count': str(len(colors_list) + 2)})
        # Fill 0 - none
        fill0 = ET.SubElement(fills, 'fill')
        ET.SubElement(fill0, 'patternFill', {'patternType': 'none'})
        # Fill 1 - gray125
        fill1 = ET.SubElement(fills, 'fill')
        ET.SubElement(fill1, 'patternFill', {'patternType': 'gray125'})
        # Custom fills
        for color in colors_list:
            fill = ET.SubElement(fills, 'fill')
            pf = ET.SubElement(fill, 'patternFill', {'patternType': 'solid'})
            ET.SubElement(pf, 'fgColor', {'rgb': rgb_to_hex(color[0], color[1], color[2])})
            ET.SubElement(pf, 'bgColor', {'indexed': '64'})
        
        # Borders
        borders = ET.SubElement(styleSheet, 'borders', {'count': '2'})
        # Border 0 - none
        border0 = ET.SubElement(borders, 'border')
        ET.SubElement(border0, 'left')
        ET.SubElement(border0, 'right')
        ET.SubElement(border0, 'top')
        ET.SubElement(border0, 'bottom')
        # Border 1 - thin
        border1 = ET.SubElement(borders, 'border')
        left = ET.SubElement(border1, 'left', {'style': 'thin'})
        ET.SubElement(left, 'color', {'indexed': '64'})
        right = ET.SubElement(border1, 'right', {'style': 'thin'})
        ET.SubElement(right, 'color', {'indexed': '64'})
        top = ET.SubElement(border1, 'top', {'style': 'thin'})
        ET.SubElement(top, 'color', {'indexed': '64'})
        bottom = ET.SubElement(border1, 'bottom', {'style': 'thin'})
        ET.SubElement(bottom, 'color', {'indexed': '64'})
        
        # Cell style xfs
        cellStyleXfs = ET.SubElement(styleSheet, 'cellStyleXfs', {'count': '1'})
        ET.SubElement(cellStyleXfs, 'xf', {'numFmtId': '0', 'fontId': '0', 'fillId': '0', 'borderId': '0'})
        
        # Cell xfs - actual cell formats
        # Style 0: normal, Style 1: bold header with default gray
        # Style 2+: bold with custom colors
        num_styles = 2 + len(colors_list)
        cellXfs = ET.SubElement(styleSheet, 'cellXfs', {'count': str(num_styles)})
        # Style 0 - normal with border
        ET.SubElement(cellXfs, 'xf', {
            'numFmtId': '0', 'fontId': '0', 'fillId': '0', 'borderId': '1',
            'xfId': '0', 'applyBorder': '1'
        })
        # Style 1 - bold header with default gray
        ET.SubElement(cellXfs, 'xf', {
            'numFmtId': '0', 'fontId': '1', 'fillId': '2', 'borderId': '1',
            'xfId': '0', 'applyFont': '1', 'applyFill': '1', 'applyBorder': '1'
        })
        # Styles for each color
        for idx, color in enumerate(colors_list):
            fill_id = idx + 2
            ET.SubElement(cellXfs, 'xf', {
                'numFmtId': '0', 'fontId': '0', 'fillId': str(fill_id), 'borderId': '1',
                'xfId': '0', 'applyFill': '1', 'applyBorder': '1'
            })
        
        # Cell styles
        cellStyles = ET.SubElement(styleSheet, 'cellStyles', {'count': '1'})
        ET.SubElement(cellStyles, 'cellStyle', {'name': 'Normal', 'xfId': '0', 'builtinId': '0'})
        
        return to_xml_string(styleSheet)
    
    def get_style_id(color_tuple, is_header=False):
        """Get style ID for a color"""
        if is_header and not color_tuple:
            return 1  # Default bold header style
        if color_tuple and color_tuple in color_to_fill_id:
            return 1 + color_to_fill_id[color_tuple]  # +1 for the header style offset
        return 0  # Normal style
    
    # Create workbook XML
    def create_workbook_xml():
        wb = ET.Element('workbook', {'xmlns': NS['main'], 'xmlns:r': NS['r']})
        sheets_el = ET.SubElement(wb, 'sheets')
        for idx, sheet_data in enumerate(sheets_data):
            name = sheet_data.get('sheet_name', 'Sheet{}'.format(idx + 1))[:31]
            name = name.replace('/', '_').replace('\\', '_').replace('*', '_')
            name = name.replace('?', '_').replace('[', '_').replace(']', '_')
            ET.SubElement(sheets_el, 'sheet', {
                'name': name, 'sheetId': str(idx + 1), 'r:id': 'rId{}'.format(idx + 1)
            })
        return to_xml_string(wb)
    
    # Create worksheet XML with formatting
    def create_worksheet_xml(sheet_data):
        ws = ET.Element('worksheet', {'xmlns': NS['main']})
        
        # Column widths
        column_widths = sheet_data.get('column_widths', [])
        if column_widths:
            cols = ET.SubElement(ws, 'cols')
            for idx, width in enumerate(column_widths):
                ET.SubElement(cols, 'col', {
                    'min': str(idx + 1), 'max': str(idx + 1),
                    'width': str(width), 'customWidth': '1'
                })
        
        sheet_data_el = ET.SubElement(ws, 'sheetData')
        
        headers = sheet_data.get('headers', [])
        rows_data = sheet_data.get('rows', [])
        element_ids = sheet_data.get('element_ids', [])
        is_from_cells = sheet_data.get('from_cells', False)
        header_formats = sheet_data.get('header_formats', [])
        row_formats = sheet_data.get('row_formats', [])
        
        # Header row
        row_el = ET.SubElement(sheet_data_el, 'row', {'r': '1'})
        
        if is_from_cells:
            for col_idx, header in enumerate(headers):
                cell_ref = '{}{}'.format(_col_letter(col_idx + 1), 1)
                # Get header color
                hdr_color = None
                if col_idx < len(header_formats) and header_formats[col_idx].get('bg_color'):
                    hdr_color = tuple(header_formats[col_idx]['bg_color'])
                style_id = get_style_id(hdr_color, is_header=True)
                
                cell_el = ET.SubElement(row_el, 'c', {'r': cell_ref, 't': 's', 's': str(style_id)})
                ET.SubElement(cell_el, 'v').text = str(get_string_index(header))
        else:
            # Element ID header
            cell_el = ET.SubElement(row_el, 'c', {'r': 'A1', 't': 's', 's': '1'})
            ET.SubElement(cell_el, 'v').text = str(get_string_index('Element ID'))
            
            for col_idx, header in enumerate(headers):
                cell_ref = '{}{}'.format(_col_letter(col_idx + 2), 1)
                cell_el = ET.SubElement(row_el, 'c', {'r': cell_ref, 't': 's', 's': '1'})
                ET.SubElement(cell_el, 'v').text = str(get_string_index(header))
        
        # Data rows
        for row_idx, row_data in enumerate(rows_data):
            excel_row = row_idx + 2
            row_el = ET.SubElement(sheet_data_el, 'row', {'r': str(excel_row)})
            row_fmt = row_formats[row_idx] if row_idx < len(row_formats) else []
            
            if is_from_cells:
                for col_idx, value in enumerate(row_data):
                    cell_ref = '{}{}'.format(_col_letter(col_idx + 1), excel_row)
                    val_str = str(value) if value is not None else ""
                    
                    # Get cell color
                    cell_color = None
                    if col_idx < len(row_fmt) and row_fmt[col_idx].get('bg_color'):
                        cell_color = tuple(row_fmt[col_idx]['bg_color'])
                    style_id = get_style_id(cell_color)
                    
                    cell_el = ET.SubElement(row_el, 'c', {'r': cell_ref, 't': 's', 's': str(style_id)})
                    ET.SubElement(cell_el, 'v').text = str(get_string_index(val_str))
            else:
                # Element ID column
                eid = element_ids[row_idx] if row_idx < len(element_ids) else None
                cell_ref = 'A{}'.format(excel_row)
                if eid:
                    cell_el = ET.SubElement(row_el, 'c', {'r': cell_ref, 's': '0'})
                    ET.SubElement(cell_el, 'v').text = str(eid)
                else:
                    cell_el = ET.SubElement(row_el, 'c', {'r': cell_ref, 't': 's', 's': '0'})
                    ET.SubElement(cell_el, 'v').text = str(get_string_index(""))
                
                for col_idx, value in enumerate(row_data):
                    cell_ref = '{}{}'.format(_col_letter(col_idx + 2), excel_row)
                    val_str = str(value) if value is not None else ""
                    
                    is_numeric = False
                    try:
                        if val_str and val_str.strip():
                            float(val_str)
                            is_numeric = True
                    except:
                        pass
                    
                    if is_numeric:
                        cell_el = ET.SubElement(row_el, 'c', {'r': cell_ref, 's': '0'})
                        ET.SubElement(cell_el, 'v').text = val_str
                    else:
                        cell_el = ET.SubElement(row_el, 'c', {'r': cell_ref, 't': 's', 's': '0'})
                        ET.SubElement(cell_el, 'v').text = str(get_string_index(val_str))
        
        return to_xml_string(ws)
    
    # Create shared strings XML
    def create_shared_strings_xml():
        sst = ET.Element('sst', {
            'xmlns': NS['main'],
            'count': str(len(shared_strings)),
            'uniqueCount': str(len(shared_strings))
        })
        for s in shared_strings:
            si = ET.SubElement(sst, 'si')
            t = ET.SubElement(si, 't')
            t.text = s
        return to_xml_string(sst)
    
    # Create content types XML
    def create_content_types_xml():
        types = ET.Element('Types', {'xmlns': NS['ct']})
        ET.SubElement(types, 'Default', {'Extension': 'rels', 'ContentType': 'application/vnd.openxmlformats-package.relationships+xml'})
        ET.SubElement(types, 'Default', {'Extension': 'xml', 'ContentType': 'application/xml'})
        ET.SubElement(types, 'Override', {'PartName': '/xl/workbook.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
        ET.SubElement(types, 'Override', {'PartName': '/xl/sharedStrings.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})
        ET.SubElement(types, 'Override', {'PartName': '/xl/styles.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'})
        for idx in range(len(sheets_data)):
            ET.SubElement(types, 'Override', {
                'PartName': '/xl/worksheets/sheet{}.xml'.format(idx + 1),
                'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
            })
        return to_xml_string(types)
    
    # Create relationships XML
    def create_rels_xml():
        rels = ET.Element('Relationships', {'xmlns': NS['rel']})
        ET.SubElement(rels, 'Relationship', {
            'Id': 'rId1',
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
            'Target': 'xl/workbook.xml'
        })
        return to_xml_string(rels)
    
    # Create workbook relationships XML
    def create_workbook_rels_xml():
        rels = ET.Element('Relationships', {'xmlns': NS['rel']})
        for idx in range(len(sheets_data)):
            ET.SubElement(rels, 'Relationship', {
                'Id': 'rId{}'.format(idx + 1),
                'Type': NS['ws'],
                'Target': 'worksheets/sheet{}.xml'.format(idx + 1)
            })
        ET.SubElement(rels, 'Relationship', {
            'Id': 'rId{}'.format(len(sheets_data) + 1),
            'Type': NS['ss'],
            'Target': 'sharedStrings.xml'
        })
        ET.SubElement(rels, 'Relationship', {
            'Id': 'rId{}'.format(len(sheets_data) + 2),
            'Type': NS['st'],
            'Target': 'styles.xml'
        })
        return to_xml_string(rels)
    
    # Write the XLSX file
    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?>' + create_content_types_xml())
        zf.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?>' + create_rels_xml())
        zf.writestr('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?>' + create_workbook_xml())
        zf.writestr('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?>' + create_workbook_rels_xml())
        zf.writestr('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8"?>' + create_shared_strings_xml())
        zf.writestr('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8"?>' + create_styles_xml())
        
        for idx, sheet_data in enumerate(sheets_data):
            zf.writestr('xl/worksheets/sheet{}.xml'.format(idx + 1), 
                       '<?xml version="1.0" encoding="UTF-8"?>' + create_worksheet_xml(sheet_data))
    
    return True


def _read_xlsx_zipfile(filepath):
    """Read XLSX file using zipfile + XML (no Office dependency)"""
    import zipfile
    try:
        from xml.etree import ElementTree as ET
    except:
        import xml.etree.ElementTree as ET
    
    NS = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    
    with zipfile.ZipFile(filepath, 'r') as zf:
        # Read shared strings
        shared_strings = []
        try:
            ss_xml = zf.read('xl/sharedStrings.xml')
            ss_root = ET.fromstring(ss_xml)
            for si in ss_root.findall('.//{%s}si' % NS['main']):
                t = si.find('.//{%s}t' % NS['main'])
                shared_strings.append(t.text if t is not None and t.text else "")
        except:
            pass
        
        # Read workbook to get sheet names
        wb_xml = zf.read('xl/workbook.xml')
        wb_root = ET.fromstring(wb_xml)
        sheets_info = []
        for sheet in wb_root.findall('.//{%s}sheet' % NS['main']):
            sheets_info.append(sheet.get('name'))
        
        # Read first data sheet
        data_sheet_idx = 0
        meta_sheet_idx = None
        
        for idx, name in enumerate(sheets_info):
            if name == "_Meta":
                meta_sheet_idx = idx
            elif not name.startswith("_") and data_sheet_idx == 0:
                data_sheet_idx = idx
        
        def read_sheet(sheet_idx):
            sheet_xml = zf.read('xl/worksheets/sheet{}.xml'.format(sheet_idx + 1))
            sheet_root = ET.fromstring(sheet_xml)
            
            rows_data = {}
            for row_el in sheet_root.findall('.//{%s}row' % NS['main']):
                row_num = int(row_el.get('r'))
                cells = {}
                for cell_el in row_el.findall('.//{%s}c' % NS['main']):
                    cell_ref = cell_el.get('r')
                    cell_type = cell_el.get('t', '')
                    v_el = cell_el.find('{%s}v' % NS['main'])
                    
                    if v_el is not None and v_el.text:
                        if cell_type == 's':
                            idx = int(v_el.text)
                            value = shared_strings[idx] if idx < len(shared_strings) else ""
                        else:
                            value = v_el.text
                    else:
                        value = ""
                    
                    # Extract column from cell reference
                    col_str = ''.join(c for c in cell_ref if c.isalpha())
                    col_num = 0
                    for c in col_str:
                        col_num = col_num * 26 + (ord(c.upper()) - ord('A') + 1)
                    cells[col_num] = value
                
                rows_data[row_num] = cells
            
            return rows_data
        
        result = {
            'sheets_info': sheets_info,
            'shared_strings': shared_strings
        }
        
        if meta_sheet_idx is not None:
            result['meta'] = read_sheet(meta_sheet_idx)
        
        result['data'] = read_sheet(data_sheet_idx)
        
        return result


class ExcelManager:
    
    @staticmethod
    def _create_excel_app():
        """Create Excel application - tries Interop first, then COM"""
        if USE_COM:
            excel_type = Type.GetTypeFromProgID("Excel.Application")
            return Activator.CreateInstance(excel_type)
        elif Excel:
            return Excel.ApplicationClass()
        else:
            raise Exception("Excel not available")
    
    @staticmethod
    def _rgb_to_ole(r, g, b):
        """Convert RGB to OLE color (BGR format)"""
        return b * 65536 + g * 256 + r
    
    @staticmethod
    def export_multiple_to_excel(filepath, schedules_data, from_cells=False):
        """Export multiple schedules to single Excel file"""
        if USE_ZIPFILE:
            return ExcelManager._export_multiple_zipfile(filepath, schedules_data, from_cells)
        else:
            return ExcelManager._export_with_formatting(filepath, schedules_data, from_cells)
    
    @staticmethod
    def _export_multiple_zipfile(filepath, schedules_data, from_cells=False):
        """Export using zipfile (no Office needed)"""
        try:
            sheets = []
            for data in schedules_data:
                sheet_data = {
                    'sheet_name': data['schedule_name'],
                    'headers': data['headers'],
                    'rows': data['rows'],
                    'element_ids': data.get('element_ids', []),
                    'from_cells': data.get('from_cells', False) or from_cells
                }
                sheets.append(sheet_data)
            
            _create_xlsx_zipfile(filepath, sheets)
            return True, None
        except Exception as e:
            return False, str(e)
    
    @staticmethod
    def _export_with_formatting(filepath, schedules_data, from_cells=False):
        """Export using Excel Interop/COM with full formatting"""
        excel_app = None
        wb = None
        
        try:
            excel_app = ExcelManager._create_excel_app()
            
            # Set properties - handle both Interop and COM
            try:
                excel_app.Visible = False
            except:
                pass
            try:
                excel_app.DisplayAlerts = False
            except:
                pass
            
            wb = excel_app.Workbooks.Add()
            
            # Remove extra sheets
            try:
                while wb.Sheets.Count > 1:
                    wb.Sheets[wb.Sheets.Count].Delete()
            except:
                pass
            
            first_sheet = True
            
            for data in schedules_data:
                schedule_name = data['schedule_name']
                headers = data['headers']
                fields = data['fields']
                rows = data['rows']
                element_ids = data.get('element_ids', [])
                is_from_cells = data.get('from_cells', False) or from_cells
                
                # Get formatting info if available
                header_formats = data.get('header_formats', [])
                row_formats = data.get('row_formats', [])
                column_widths = data.get('column_widths', [])
                
                if first_sheet:
                    ws = wb.Sheets[1]
                    first_sheet = False
                else:
                    try:
                        ws = wb.Sheets.Add(After=wb.Sheets[wb.Sheets.Count])
                    except:
                        ws = wb.Sheets.Add()
                
                safe_name = schedule_name[:31].replace("/", "_").replace("\\", "_").replace("*", "_")
                safe_name = safe_name.replace("?", "_").replace("[", "_").replace("]", "_")
                try:
                    ws.Name = safe_name
                except:
                    pass
                
                if is_from_cells:
                    # Apply column widths
                    for col_idx, width in enumerate(column_widths):
                        try:
                            ws.Columns[col_idx + 1].ColumnWidth = width
                        except:
                            pass
                    
                    # Header row with formatting
                    for col, header in enumerate(headers):
                        try:
                            cell = ws.Cells[1, col + 1]
                            cell.Value2 = header
                            try:
                                cell.Font.Bold = True
                            except:
                                pass
                            
                            # Apply header background color
                            try:
                                if col < len(header_formats) and header_formats[col].get('bg_color'):
                                    rgb = header_formats[col]['bg_color']
                                    cell.Interior.Color = ExcelManager._rgb_to_ole(rgb[0], rgb[1], rgb[2])
                                else:
                                    cell.Interior.Color = 0xD0D0D0  # Default gray
                            except:
                                pass
                            
                            # Add border
                            try:
                                cell.Borders.LineStyle = 1  # xlContinuous
                            except:
                                pass
                        except:
                            pass
                    
                    # Data rows with formatting
                    for row_idx, row_data in enumerate(rows):
                        excel_row = row_idx + 2
                        row_fmt = row_formats[row_idx] if row_idx < len(row_formats) else []
                        
                        for col_idx, value in enumerate(row_data):
                            try:
                                cell = ws.Cells[excel_row, col_idx + 1]
                                cell.Value2 = value
                                
                                # Apply cell background color if available
                                try:
                                    if col_idx < len(row_fmt) and row_fmt[col_idx].get('bg_color'):
                                        rgb = row_fmt[col_idx]['bg_color']
                                        cell.Interior.Color = ExcelManager._rgb_to_ole(rgb[0], rgb[1], rgb[2])
                                except:
                                    pass
                                
                                # Add border
                                try:
                                    cell.Borders.LineStyle = 1
                                except:
                                    pass
                            except:
                                pass
                    
                    # Auto-fit if no column widths provided
                    if not column_widths:
                        try:
                            ws.Columns.AutoFit()
                        except:
                            pass
                else:
                    # Normal export with Element ID
                    try:
                        ws.Cells[1, 1].Value2 = "Element ID"
                        ws.Cells[1, 1].Font.Bold = True
                        ws.Cells[1, 1].Interior.Color = 0xA5E6D4
                    except:
                        pass
                    
                    for col, header in enumerate(headers):
                        try:
                            cell = ws.Cells[1, col + 2]
                            cell.Value2 = header
                            cell.Font.Bold = True
                            is_editable = col < len(fields) and fields[col].get('can_edit', True)
                            cell.Interior.Color = 0xCDFF97 if is_editable else 0xE0E0E0
                        except:
                            pass
                    
                    for row_idx, row_data in enumerate(rows):
                        excel_row = row_idx + 2
                        eid = element_ids[row_idx] if row_idx < len(element_ids) else None
                        try:
                            ws.Cells[excel_row, 1].Value2 = eid if eid else ""
                        except:
                            pass
                        
                        for col_idx, value in enumerate(row_data):
                            try:
                                ws.Cells[excel_row, col_idx + 2].Value2 = value
                            except:
                                pass
                    
                    try:
                        ws.Columns.AutoFit()
                    except:
                        pass
            
            # Add meta sheet if not from_cells mode (skip for COM as it may fail)
            if not from_cells and not USE_COM:
                try:
                    meta = wb.Sheets.Add(After=wb.Sheets[wb.Sheets.Count])
                    meta.Name = "_Meta"
                    
                    meta.Cells[1, 1].Value2 = "ScheduleCount"
                    meta.Cells[1, 2].Value2 = len(schedules_data)
                    
                    row = 3
                    for idx, data in enumerate(schedules_data):
                        meta.Cells[row, 1].Value2 = "Schedule_{}".format(idx)
                        meta.Cells[row, 2].Value2 = data['schedule_name']
                        meta.Cells[row, 3].Value2 = len(data['fields'])
                        row += 1
                        
                        for f in data['fields']:
                            meta.Cells[row, 1].Value2 = f['name']
                            meta.Cells[row, 2].Value2 = f['param_id'] if f['param_id'] else ""
                            meta.Cells[row, 3].Value2 = 1 if f['can_edit'] else 0
                            row += 1
                        
                        row += 1
                    
                    try:
                        meta.Visible = False
                    except:
                        pass
                except:
                    pass
            
            try:
                wb.Sheets[1].Activate()
            except:
                pass
            
            wb.SaveAs(filepath)
            wb.Close()
            excel_app.Quit()
            
            return True, None
            
        except Exception as e:
            if wb:
                try: wb.Close(False)
                except: pass
            if excel_app:
                try: excel_app.Quit()
                except: pass
            return False, str(e)
    
    @staticmethod
    def _export_multiple_interop(filepath, schedules_data, from_cells=False):
        """Export using Excel Interop - legacy method"""
        return ExcelManager._export_with_formatting(filepath, schedules_data, from_cells)
    
    @staticmethod
    def export_to_excel(filepath, data):
        """Export single schedule to Excel"""
        return ExcelManager.export_multiple_to_excel(filepath, [data], data.get('from_cells', False))
    
    @staticmethod
    def import_from_excel(filepath):
        """Import Excel file"""
        if USE_ZIPFILE:
            return ExcelManager._import_zipfile(filepath)
        else:
            return ExcelManager._import_interop(filepath)
    
    @staticmethod
    def _import_zipfile(filepath):
        """Import using zipfile"""
        try:
            result = _read_xlsx_zipfile(filepath)
            
            meta = result.get('meta', {})
            data_rows = result.get('data', {})
            
            if not meta:
                return None, "Invalid file format - no _Meta sheet"
            
            # Parse meta sheet
            schedule_name = meta.get(1, {}).get(2, "")
            field_count = int(meta.get(2, {}).get(2, 0) or 0)
            
            fields = []
            for i in range(field_count):
                row_num = i + 4
                name = meta.get(row_num, {}).get(1, "")
                param_id_str = meta.get(row_num, {}).get(2, "")
                can_edit_str = meta.get(row_num, {}).get(3, "0")
                
                param_id = int(param_id_str) if param_id_str and str(param_id_str).lstrip('-').isdigit() else None
                can_edit = str(can_edit_str) == "1"
                
                if name:
                    fields.append({'name': name, 'param_id': param_id, 'can_edit': can_edit})
            
            # Parse data sheet
            if not data_rows:
                return None, "No data in file"
            
            # Check header row
            header_row = data_rows.get(1, {})
            h1 = header_row.get(1, "")
            
            if h1.lower() in ['element id', 'elementid', 'id']:
                eid_col = 1
                data_start = 2
            else:
                eid_col = None
                data_start = 1
            
            # Get headers
            headers = []
            col = data_start
            while col in header_row or col <= data_start + 20:
                h = header_row.get(col, "")
                if h:
                    headers.append(h)
                elif col > data_start + 5 and not any(header_row.get(c) for c in range(col, col + 5)):
                    break
                col += 1
            
            # Get data rows
            rows = []
            element_ids = []
            
            row_nums = sorted([r for r in data_rows.keys() if r > 1])
            
            for row_num in row_nums:
                row_data = data_rows[row_num]
                
                if eid_col:
                    eid_val = row_data.get(eid_col, "")
                    try:
                        element_ids.append(int(float(eid_val)) if eid_val else None)
                    except:
                        element_ids.append(None)
                
                row_values = []
                for col_idx in range(len(headers)):
                    col_num = data_start + col_idx
                    row_values.append(row_data.get(col_num, ""))
                
                if any(v for v in row_values):
                    rows.append(row_values)
                elif eid_col and element_ids and element_ids[-1]:
                    rows.append(row_values)
            
            return {
                'schedule_name': schedule_name,
                'headers': headers,
                'fields': fields,
                'rows': rows,
                'element_ids': element_ids
            }, None
            
        except Exception as e:
            return None, str(e)
    
    @staticmethod
    def _import_interop(filepath):
        """Import using Excel Interop"""
        excel_app = None
        wb = None
        
        try:
            excel_app = Excel.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            wb = excel_app.Workbooks.Open(filepath)
            
            meta_sheet = None
            data_sheet = None
            
            for i in range(1, wb.Sheets.Count + 1):
                sheet = wb.Sheets[i]
                if sheet.Name == "_Meta":
                    meta_sheet = sheet
                elif not sheet.Name.startswith("_"):
                    data_sheet = sheet
            
            if not meta_sheet or not data_sheet:
                wb.Close()
                excel_app.Quit()
                return None, "Invalid file format"
            
            schedule_name = safe_str(get_cell_value(meta_sheet.Cells[1, 2]))
            field_count = safe_int(get_cell_value(meta_sheet.Cells[2, 2])) or 0
            
            fields = []
            for i in range(field_count):
                name = safe_str(get_cell_value(meta_sheet.Cells[i + 4, 1]))
                param_id = safe_int(get_cell_value(meta_sheet.Cells[i + 4, 2]))
                can_edit = safe_int(get_cell_value(meta_sheet.Cells[i + 4, 3])) == 1
                if name:
                    fields.append({'name': name, 'param_id': param_id, 'can_edit': can_edit})
            
            h1 = safe_str(get_cell_value(data_sheet.Cells[1, 1]))
            
            if h1.lower() in ['element id', 'elementid', 'id']:
                eid_col = 1
                data_start = 2
            else:
                eid_col = None
                data_start = 1
            
            headers = []
            col = data_start
            while True:
                h = safe_str(get_cell_value(data_sheet.Cells[1, col]))
                if not h:
                    break
                headers.append(h)
                col += 1
                if col > 100:
                    break
            
            rows, element_ids = [], []
            row, empty = 2, 0
            
            while empty < 5:
                has_data = False
                
                if eid_col:
                    eid_val = get_cell_value(data_sheet.Cells[row, eid_col])
                    try:
                        element_ids.append(int(float(eid_val)) if eid_val else None)
                    except:
                        element_ids.append(None)
                
                row_data = []
                for col_idx in range(len(headers)):
                    val = safe_str(get_cell_value(data_sheet.Cells[row, data_start + col_idx]))
                    row_data.append(val)
                    if val:
                        has_data = True
                
                if has_data:
                    rows.append(row_data)
                    empty = 0
                else:
                    empty += 1
                
                row += 1
                if row > 10000:
                    break
            
            wb.Close()
            excel_app.Quit()
            
            return {
                'schedule_name': schedule_name,
                'headers': headers,
                'fields': fields,
                'rows': rows,
                'element_ids': element_ids
            }, None
            
        except Exception as e:
            if wb:
                try: wb.Close(False)
                except: pass
            if excel_app:
                try: excel_app.Quit()
                except: pass
            return None, str(e)


# =============================================================================
# CHANGE DETECTOR
# =============================================================================
class ChangeDetector:
    @staticmethod
    def find_changes(current_data, imported_data):
        changes = []
        if not current_data or not imported_data:
            return changes
        
        cur_headers = current_data['headers']
        imp_headers = imported_data.get('headers', [])
        imp_fields = imported_data.get('fields', [])
        cur_col_map = {h: idx for idx, h in enumerate(cur_headers)}
        
        cur_by_eid = {}
        for row_idx, eid in enumerate(current_data['element_ids']):
            if eid:
                cur_by_eid[eid] = {
                    'row_idx': row_idx,
                    'row_data': current_data['rows'][row_idx]
                }
        
        for imp_idx, imp_row in enumerate(imported_data['rows']):
            imp_eid = imported_data['element_ids'][imp_idx] if imp_idx < len(imported_data['element_ids']) else None
            
            if not imp_eid:
                continue
            
            cur_info = cur_by_eid.get(imp_eid)
            if not cur_info:
                continue
            
            cur_row = cur_info['row_data']
            
            for col_idx, header in enumerate(imp_headers):
                is_editable = col_idx < len(imp_fields) and imp_fields[col_idx].get('can_edit', True)
                param_id = imp_fields[col_idx].get('param_id') if col_idx < len(imp_fields) else None
                
                if not is_editable:
                    continue
                
                imp_val = safe_str(imp_row[col_idx] if col_idx < len(imp_row) else "")
                cur_col = cur_col_map.get(header)
                if cur_col is None:
                    continue
                cur_val = safe_str(cur_row[cur_col] if cur_col < len(cur_row) else "")
                
                if imp_val != cur_val:
                    changes.append({
                        'imp_row_index': imp_idx,
                        'field_name': header,
                        'param_id': param_id,
                        'element_id': imp_eid,
                        'old_value': cur_val,
                        'new_value': imp_val,
                        'excel_row': imp_idx + 2
                    })
        
        return changes


# =============================================================================
# MODEL UPDATER
# =============================================================================
class ModelUpdater:
    
    @staticmethod
    def apply_changes(schedule, changes, progress_cb=None):
        global UPDATE_LOG
        UPDATE_LOG = []
        
        if not changes:
            return 0, [], 0
        
        success, errors, skipped = 0, [], 0
        t = Transaction(doc, "Schedule Link Pro - Update")
        t.Start()
        
        try:
            for idx, ch in enumerate(changes):
                if progress_cb:
                    progress_cb(idx + 1, len(changes))
                
                elem = None
                elem_id = ch.get('element_id')
                if elem_id:
                    try:
                        elem = doc.GetElement(ElementId(elem_id))
                    except:
                        pass
                
                if not elem:
                    skipped += 1
                    continue
                
                ok, err = ModelUpdater._set_param(elem, ch)
                if ok:
                    success += 1
                    log_update("OK: ID {} - {} = '{}'".format(
                        elem_id, ch['field_name'], 
                        ch['new_value'][:30] if ch['new_value'] else ''))
                else:
                    errors.append("ID {}: {}".format(elem_id, err))
            
            t.Commit()
            
        except Exception as e:
            t.RollBack()
            return 0, [str(e)], 0
        
        return success, errors, skipped
    
    @staticmethod
    def _set_param(elem, change):
        try:
            param = None
            field_name = change['field_name']
            param_id = change['param_id']
            value = change['new_value']
            
            if param_id and param_id < 0:
                try:
                    bip = BuiltInParameter(param_id)
                    param = elem.get_Parameter(bip)
                except:
                    pass
            
            if not param:
                param = elem.LookupParameter(field_name)
            
            if not param:
                for p in elem.Parameters:
                    try:
                        if p.Definition and p.Definition.Name == field_name:
                            param = p
                            break
                    except:
                        pass
            
            if not param:
                elem_type = doc.GetElement(elem.GetTypeId())
                if elem_type:
                    if param_id and param_id < 0:
                        try:
                            bip = BuiltInParameter(param_id)
                            param = elem_type.get_Parameter(bip)
                        except:
                            pass
                    if not param:
                        param = elem_type.LookupParameter(field_name)
            
            if not param:
                return False, "'{}' not found".format(field_name)
            
            if param.IsReadOnly:
                return False, "'{}' read-only".format(field_name)
            
            storage = param.StorageType
            
            if storage == StorageType.String:
                param.Set(str(value) if value else "")
            elif storage == StorageType.Integer:
                if value:
                    param.Set(int(float(value)))
                else:
                    param.Set(0)
            elif storage == StorageType.Double:
                if value:
                    param.Set(float(value))
                else:
                    param.Set(0.0)
            else:
                return False, "Unsupported type"
            
            return True, ""
            
        except Exception as e:
            return False, str(e)


# =============================================================================
# MAIN WINDOW
# =============================================================================
class MainWindow(Window):
    
    def __init__(self):
        self.Title = "Schedule Link Pro - DQT"
        self.Width = 1100
        self.Height = 750
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = brush(BACKGROUND)
        
        self.schedules = ScheduleExtractor.get_all_schedules()
        self.current_schedule = None
        self.current_data = None
        self.imported_data = None
        self.changes = []
        
        self._build_ui()
        self._load_schedules()
    
    def _build_ui(self):
        main = Grid()
        main.Margin = Thickness(10)
        
        for _ in range(4):
            main.RowDefinitions.Add(RowDefinition())
        main.RowDefinitions[0].Height = GridLength.Auto
        main.RowDefinitions[1].Height = GridLength.Auto
        main.RowDefinitions[3].Height = GridLength.Auto
        
        copyright_top = TextBlock()
        copyright_top.Text = COPYRIGHT
        copyright_top.FontSize = 10
        copyright_top.Foreground = brush(GRAY)
        copyright_top.HorizontalAlignment = HorizontalAlignment.Center
        copyright_top.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(copyright_top, 0)
        main.Children.Add(copyright_top)
        
        top = self._build_top_bar()
        Grid.SetRow(top, 1)
        main.Children.Add(top)
        
        scroll = ScrollViewer()
        scroll.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.Margin = Thickness(0, 10, 0, 10)
        
        self.grid_container = StackPanel()
        scroll.Content = self.grid_container
        
        Grid.SetRow(scroll, 2)
        main.Children.Add(scroll)
        
        bottom = self._build_bottom_bar()
        Grid.SetRow(bottom, 3)
        main.Children.Add(bottom)
        
        self.Content = main
    
    def _build_top_bar(self):
        panel = StackPanel()
        
        title = TextBlock()
        title.Text = "Schedule Link Pro"
        title.FontSize = 20
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(title)
        
        # Mode selection
        mode_row = DockPanel()
        mode_row.Margin = Thickness(0, 0, 0, 5)
        
        self.single_mode_cb = CheckBox()
        self.single_mode_cb.Content = "Single Schedule"
        self.single_mode_cb.IsChecked = True
        self.single_mode_cb.Margin = Thickness(0, 0, 20, 0)
        self.single_mode_cb.Checked += self._on_mode_change
        mode_row.Children.Add(self.single_mode_cb)
        
        self.multi_mode_cb = CheckBox()
        self.multi_mode_cb.Content = "Multi-Schedule Export"
        self.multi_mode_cb.IsChecked = False
        self.multi_mode_cb.Checked += self._on_mode_change
        mode_row.Children.Add(self.multi_mode_cb)
        
        panel.Children.Add(mode_row)
        
        # Single schedule row
        self.single_row = DockPanel()
        self.single_row.Margin = Thickness(0, 0, 0, 10)
        
        lbl = TextBlock()
        lbl.Text = "Schedule: "
        lbl.Width = 70
        lbl.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(lbl, System.Windows.Controls.Dock.Left)
        self.single_row.Children.Add(lbl)
        
        self.preview_btn = self._btn("Preview", self._on_preview)
        self.preview_btn.Width = 100
        DockPanel.SetDock(self.preview_btn, System.Windows.Controls.Dock.Right)
        self.single_row.Children.Add(self.preview_btn)
        
        self.sch_combo = ComboBox()
        self.sch_combo.Height = 28
        self.sch_combo.Margin = Thickness(0, 0, 10, 0)
        self.single_row.Children.Add(self.sch_combo)
        
        panel.Children.Add(self.single_row)
        
        # Multi schedule panel (hidden by default)
        self.multi_panel = StackPanel()
        self.multi_panel.Visibility = Visibility.Collapsed
        
        multi_lbl = TextBlock()
        multi_lbl.Text = "Select schedules to export:"
        multi_lbl.Margin = Thickness(0, 0, 0, 5)
        self.multi_panel.Children.Add(multi_lbl)
        
        # Select All / None buttons
        select_row = StackPanel()
        select_row.Orientation = System.Windows.Controls.Orientation.Horizontal
        select_row.Margin = Thickness(0, 0, 0, 5)
        
        self.select_all_btn = Button()
        self.select_all_btn.Content = "Select All"
        self.select_all_btn.Width = 80
        self.select_all_btn.Height = 24
        self.select_all_btn.Margin = Thickness(0, 0, 10, 0)
        self.select_all_btn.Click += self._on_select_all
        select_row.Children.Add(self.select_all_btn)
        
        self.select_none_btn = Button()
        self.select_none_btn.Content = "Select None"
        self.select_none_btn.Width = 80
        self.select_none_btn.Height = 24
        self.select_none_btn.Click += self._on_select_none
        select_row.Children.Add(self.select_none_btn)
        
        self.multi_panel.Children.Add(select_row)
        
        # Schedule checkboxes in scrollable list
        self.schedule_list_scroll = ScrollViewer()
        self.schedule_list_scroll.MaxHeight = 150
        self.schedule_list_scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        
        self.schedule_list = StackPanel()
        self.schedule_list_scroll.Content = self.schedule_list
        self.multi_panel.Children.Add(self.schedule_list_scroll)
        
        self.selected_count_text = TextBlock()
        self.selected_count_text.Text = "0 schedules selected"
        self.selected_count_text.Margin = Thickness(0, 5, 0, 0)
        self.selected_count_text.Foreground = brush(GRAY)
        self.multi_panel.Children.Add(self.selected_count_text)
        
        panel.Children.Add(self.multi_panel)
        
        self.info_text = TextBlock()
        self.info_text.Text = "Select a schedule and click Preview"
        self.info_text.Foreground = brush(GRAY)
        self.info_text.TextWrapping = TextWrapping.Wrap
        panel.Children.Add(self.info_text)
        
        return panel
    
    def _on_mode_change(self, sender, e):
        if sender == self.single_mode_cb and self.single_mode_cb.IsChecked:
            self.multi_mode_cb.IsChecked = False
            self.single_row.Visibility = Visibility.Visible
            self.multi_panel.Visibility = Visibility.Collapsed
            self.export_btn.Content = "Export to Excel"
        elif sender == self.multi_mode_cb and self.multi_mode_cb.IsChecked:
            self.single_mode_cb.IsChecked = False
            self.single_row.Visibility = Visibility.Collapsed
            self.multi_panel.Visibility = Visibility.Visible
            self.export_btn.Content = "Export All Selected"
            self.export_btn.IsEnabled = True
            self._update_selected_count()
    
    def _on_select_all(self, s, e):
        for child in self.schedule_list.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = True
        self._update_selected_count()
    
    def _on_select_none(self, s, e):
        for child in self.schedule_list.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = False
        self._update_selected_count()
    
    def _update_selected_count(self):
        count = sum(1 for child in self.schedule_list.Children 
                   if isinstance(child, CheckBox) and child.IsChecked)
        self.selected_count_text.Text = "{} schedules selected".format(count)
        if count > 0:
            self.selected_count_text.Foreground = brush(GREEN)
        else:
            self.selected_count_text.Foreground = brush(GRAY)
    
    def _on_schedule_check_changed(self, s, e):
        self._update_selected_count()
    
    def _build_bottom_bar(self):
        panel = StackPanel()
        
        self.progress = ProgressBar()
        self.progress.Height = 20
        self.progress.Margin = Thickness(0, 0, 0, 10)
        self.progress.Visibility = Visibility.Collapsed
        panel.Children.Add(self.progress)
        
        # Export options row
        options_row = StackPanel()
        options_row.Orientation = System.Windows.Controls.Orientation.Horizontal
        options_row.Margin = Thickness(0, 0, 0, 10)
        
        self.keep_formatting_cb = CheckBox()
        self.keep_formatting_cb.Content = "Keep Formatting of Schedule (read-only export, includes calculated fields)"
        self.keep_formatting_cb.IsChecked = False
        self.keep_formatting_cb.Margin = Thickness(0, 0, 20, 0)
        options_row.Children.Add(self.keep_formatting_cb)
        
        panel.Children.Add(options_row)
        
        btn_row = DockPanel()
        
        right_panel = StackPanel()
        right_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        DockPanel.SetDock(right_panel, System.Windows.Controls.Dock.Right)
        
        self.import_btn = self._btn("Import Excel", self._on_import)
        self.import_btn.Width = 120
        self.import_btn.Margin = Thickness(5, 0, 0, 0)
        right_panel.Children.Add(self.import_btn)
        
        self.update_btn = self._btn("Update Model", self._on_update)
        self.update_btn.Width = 120
        self.update_btn.Margin = Thickness(5, 0, 0, 0)
        self.update_btn.IsEnabled = False
        self.update_btn.Background = brush(GREEN)
        right_panel.Children.Add(self.update_btn)
        
        btn_row.Children.Add(right_panel)
        
        left_panel = StackPanel()
        left_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        
        self.export_btn = self._btn("Export to Excel", self._on_export)
        self.export_btn.Width = 130
        self.export_btn.IsEnabled = False
        left_panel.Children.Add(self.export_btn)
        
        self.debug_btn = self._btn("Debug Info", self._on_debug)
        self.debug_btn.Width = 100
        self.debug_btn.Margin = Thickness(10, 0, 0, 0)
        left_panel.Children.Add(self.debug_btn)
        
        btn_row.Children.Add(left_panel)
        panel.Children.Add(btn_row)
        
        self.status_text = TextBlock()
        self.status_text.Text = ""
        self.status_text.Margin = Thickness(0, 10, 0, 5)
        self.status_text.TextWrapping = TextWrapping.Wrap
        panel.Children.Add(self.status_text)
        
        copyright_bottom = TextBlock()
        copyright_bottom.Text = COPYRIGHT
        copyright_bottom.FontSize = 10
        copyright_bottom.Foreground = brush(GRAY)
        copyright_bottom.HorizontalAlignment = HorizontalAlignment.Center
        copyright_bottom.Margin = Thickness(0, 10, 0, 0)
        panel.Children.Add(copyright_bottom)
        
        return panel
    
    def _btn(self, text, handler):
        b = Button()
        b.Content = text
        b.Height = 32
        b.MinWidth = 80
        b.Background = brush(PRIMARY)
        b.Padding = Thickness(10, 5, 10, 5)
        b.Click += handler
        return b
    
    def _load_schedules(self):
        self.sch_combo.Items.Clear()
        self.schedule_list.Children.Clear()
        self.schedule_checkboxes = {}
        
        for name in sorted(self.schedules.keys()):
            # For combo box
            item = ComboBoxItem()
            item.Content = name
            self.sch_combo.Items.Add(item)
            
            # For multi-select list
            cb = CheckBox()
            cb.Content = name
            cb.Margin = Thickness(0, 2, 0, 2)
            cb.Checked += self._on_schedule_check_changed
            cb.Unchecked += self._on_schedule_check_changed
            self.schedule_list.Children.Add(cb)
            self.schedule_checkboxes[name] = cb
        
        if self.sch_combo.Items.Count > 0:
            self.sch_combo.SelectedIndex = 0
    
    def _build_grid(self, data, change_lookup=None):
        self.grid_container.Children.Clear()
        
        headers = data.get('headers', [])
        rows = data.get('rows', [])
        element_ids = data.get('element_ids', [])
        fields = data.get('fields', [])
        
        grid = Grid()
        grid.Background = brush(WHITE)
        
        num_cols = len(headers) + 1
        for i in range(num_cols):
            col_def = ColumnDefinition()
            col_def.Width = GridLength.Auto
            grid.ColumnDefinitions.Add(col_def)
        
        num_rows = len(rows) + 1
        for i in range(num_rows):
            row_def = RowDefinition()
            row_def.Height = GridLength.Auto
            grid.RowDefinitions.Add(row_def)
        
        self._add_cell(grid, 0, 0, "Element ID", True, False, False)
        for col_idx, header in enumerate(headers):
            is_editable = col_idx < len(fields) and fields[col_idx].get('can_edit', True)
            self._add_cell(grid, 0, col_idx + 1, header, True, is_editable, False)
        
        for row_idx, row_data in enumerate(rows):
            eid = element_ids[row_idx] if row_idx < len(element_ids) else None
            self._add_cell(grid, row_idx + 1, 0, str(eid) if eid else "", False, False, False)
            
            for col_idx in range(len(headers)):
                val = row_data[col_idx] if col_idx < len(row_data) else ""
                is_changed = False
                old_val = None
                if change_lookup:
                    key = (row_idx, headers[col_idx])
                    if key in change_lookup:
                        is_changed = True
                        old_val = change_lookup[key]['old_value']
                
                is_editable = col_idx < len(fields) and fields[col_idx].get('can_edit', True)
                self._add_cell(grid, row_idx + 1, col_idx + 1, val, False, is_editable, is_changed, old_val)
        
        self.grid_container.Children.Add(grid)
    
    def _add_cell(self, grid, row, col, text, is_header, is_editable, is_changed, old_value=None):
        border = Border()
        border.BorderBrush = brush("#FFD0D0D0")
        border.BorderThickness = Thickness(1)
        border.Padding = Thickness(6, 3, 6, 3)
        
        if is_header:
            if col == 0:
                border.Background = brush("#FFA5E6D4")
            elif is_editable:
                border.Background = brush("#FFCDFF97")
            else:
                border.Background = brush("#FFE0E0E0")
        elif is_changed:
            border.Background = brush("#FFFFF3CD")
        elif row % 2 == 0:
            border.Background = brush("#FFF8F8F8")
        else:
            border.Background = brush(WHITE)
        
        tb = TextBlock()
        tb.Text = text
        tb.TextWrapping = TextWrapping.NoWrap
        
        if is_header:
            tb.FontWeight = FontWeights.Bold
        elif is_changed:
            tb.Foreground = brush(RED)
            tb.FontWeight = FontWeights.Bold
            if old_value:
                tb.ToolTip = "Was: " + str(old_value)
        
        border.Child = tb
        Grid.SetRow(border, row)
        Grid.SetColumn(border, col)
        grid.Children.Add(border)
    
    def _on_debug(self, s, e):
        global DEBUG_LOG, UPDATE_LOG
        
        msg = "=== DEBUG LOG ===\n"
        msg += "\n".join(DEBUG_LOG[-50:]) if DEBUG_LOG else "(empty)"
        msg += "\n\n=== UPDATE LOG ===\n"
        msg += "\n".join(UPDATE_LOG[-25:]) if UPDATE_LOG else "(none)"
        
        MessageBox.Show(msg, "Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
    
    def _on_preview(self, s, e):
        global DEBUG_LOG
        DEBUG_LOG = []
        
        if not self.sch_combo.SelectedItem:
            return
        
        name = self.sch_combo.SelectedItem.Content
        schedule = self.schedules.get(name)
        
        if not schedule:
            return
        
        self.current_schedule = schedule
        self.current_data = ScheduleExtractor.extract(schedule)
        
        if not self.current_data:
            self.info_text.Text = "Error extracting schedule"
            self.info_text.Foreground = brush(RED)
            return
        
        self._build_grid(self.current_data)
        
        editable = sum(1 for f in self.current_data.get('fields', []) if f.get('can_edit', True))
        total = self.current_data.get('valid_element_count', 0)
        
        self.info_text.Text = "{} cols ({} editable) | {} elements".format(
            len(self.current_data['headers']), editable, total)
        self.info_text.Foreground = brush("#FF000000")
        self.status_text.Text = ""
        
        self.export_btn.IsEnabled = True
        self.changes = []
        self.update_btn.IsEnabled = False
    
    def _on_export(self, s, e):
        # Check if multi-mode
        if self.multi_mode_cb.IsChecked:
            self._on_export_multi()
            return
        
        # Single schedule export
        if not self.sch_combo.SelectedItem:
            return
        
        name = self.sch_combo.SelectedItem.Content
        schedule = self.schedules.get(name)
        
        if not schedule:
            return
        
        # Check if Keep Formatting is enabled
        keep_formatting = self.keep_formatting_cb.IsChecked
        
        if keep_formatting:
            # Extract from cells for full data including calculated fields
            export_data = ScheduleExtractor.extract_from_cells(schedule)
        else:
            # Use current_data if available, otherwise extract
            if self.current_data and self.current_data.get('schedule_name') == name:
                export_data = self.current_data
            else:
                export_data = ScheduleExtractor.extract(schedule)
        
        if not export_data:
            self.status_text.Text = "Error extracting schedule"
            self.status_text.Foreground = brush(RED)
            return
        
        dlg = SaveFileDialog()
        dlg.Filter = "Excel (*.xlsx)|*.xlsx"
        safe_name = name.replace(" ", "_").replace("/", "_")
        dlg.FileName = "{}.xlsx".format(safe_name)
        
        if dlg.ShowDialog() != DialogResult.OK:
            return
        
        ok, err = ExcelManager.export_to_excel(dlg.FileName, export_data)
        
        if ok:
            count = export_data.get('valid_element_count', 0)
            mode_text = "(formatted)" if keep_formatting else ""
            self.status_text.Text = "Exported {} rows {}!".format(count, mode_text)
            self.status_text.Foreground = brush(GREEN)
            MessageBox.Show("Exported {} rows {}!".format(count, mode_text), "Success", 
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
            os.startfile(dlg.FileName)
        else:
            self.status_text.Text = "Export failed: " + str(err)
            self.status_text.Foreground = brush(RED)
    
    def _on_export_multi(self):
        """Export multiple selected schedules to single Excel file"""
        # Get selected schedules
        selected_names = []
        for name, cb in self.schedule_checkboxes.items():
            if cb.IsChecked:
                selected_names.append(name)
        
        if not selected_names:
            MessageBox.Show("Please select at least one schedule!", "No Selection", 
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        dlg = SaveFileDialog()
        dlg.Filter = "Excel (*.xlsx)|*.xlsx"
        dlg.FileName = "Combined_Schedules.xlsx"
        
        if dlg.ShowDialog() != DialogResult.OK:
            return
        
        # Check if Keep Formatting is enabled
        keep_formatting = self.keep_formatting_cb.IsChecked
        
        # Show progress
        self.progress.Visibility = Visibility.Visible
        self.progress.Maximum = len(selected_names)
        self.progress.Value = 0
        
        # Extract data from all selected schedules
        schedules_data = []
        total_rows = 0
        
        for idx, name in enumerate(selected_names):
            schedule = self.schedules.get(name)
            if schedule:
                if keep_formatting:
                    data = ScheduleExtractor.extract_from_cells(schedule)
                else:
                    data = ScheduleExtractor.extract(schedule)
                    
                if data:
                    schedules_data.append(data)
                    total_rows += data.get('valid_element_count', 0)
            self.progress.Value = idx + 1
        
        self.progress.Visibility = Visibility.Collapsed
        
        if not schedules_data:
            MessageBox.Show("No data extracted from selected schedules!", "Error", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        
        # Export to Excel
        ok, err = ExcelManager.export_multiple_to_excel(dlg.FileName, schedules_data, keep_formatting)
        
        if ok:
            mode_text = "(formatted)" if keep_formatting else ""
            self.status_text.Text = "Exported {} schedules ({} rows) {}!".format(
                len(schedules_data), total_rows, mode_text)
            self.status_text.Foreground = brush(GREEN)
            MessageBox.Show(
                "Exported {} schedules with {} total rows {}!\n\nEach schedule is in a separate sheet.".format(
                    len(schedules_data), total_rows, mode_text), 
                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            os.startfile(dlg.FileName)
        else:
            self.status_text.Text = "Export failed: " + str(err)
            self.status_text.Foreground = brush(RED)
    
    def _on_import(self, s, e):
        dlg = OpenFileDialog()
        dlg.Filter = "Excel (*.xlsx)|*.xlsx"
        
        if dlg.ShowDialog() != DialogResult.OK:
            return
        
        data, err = ExcelManager.import_from_excel(dlg.FileName)
        
        if err:
            self.status_text.Text = "Import failed: " + str(err)
            self.status_text.Foreground = brush(RED)
            return
        
        self.imported_data = data
        schedule = self.schedules.get(data['schedule_name'])
        if not schedule:
            self.status_text.Text = "Schedule '{}' not found!".format(data['schedule_name'])
            self.status_text.Foreground = brush(RED)
            return
        
        self.current_schedule = schedule
        self.current_data = ScheduleExtractor.extract(schedule)
        
        self.changes = ChangeDetector.find_changes(self.current_data, data)
        
        change_lookup = {(ch['imp_row_index'], ch['field_name']): ch for ch in self.changes}
        self._build_grid(data, change_lookup)
        
        if self.changes:
            self.status_text.Text = "{} changes detected".format(len(self.changes))
            self.status_text.Foreground = brush("#FF0000FF")
            self.update_btn.IsEnabled = True
            self.info_text.Text = "IMPORTED: {} | {} rows".format(
                data['schedule_name'], len(data['rows']))
            self.info_text.Foreground = brush(RED)
        else:
            self.status_text.Text = "No changes detected"
            self.status_text.Foreground = brush(GRAY)
            self.update_btn.IsEnabled = False
    
    def _on_update(self, s, e):
        global UPDATE_LOG
        UPDATE_LOG = []
        
        if not self.changes or not self.current_schedule:
            return
        
        r = MessageBox.Show(
            "Apply {} changes to model?".format(len(self.changes)), 
            "Confirm Update", 
            MessageBoxButtons.YesNo, 
            MessageBoxIcon.Question)
        
        if r != DialogResult.Yes:
            return
        
        self.progress.Visibility = Visibility.Visible
        self.progress.Maximum = len(self.changes)
        
        success, errors, skipped = ModelUpdater.apply_changes(
            self.current_schedule, 
            self.changes,
            lambda cur, total: setattr(self.progress, 'Value', cur))
        
        self.progress.Visibility = Visibility.Collapsed
        
        msg = "Updated: {}\nSkipped: {}\nErrors: {}".format(success, skipped, len(errors))
        
        if success > 0:
            self.status_text.Text = "Updated {} elements!".format(success)
            self.status_text.Foreground = brush(GREEN)
            MessageBox.Show(msg, "Update Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
        else:
            self.status_text.Text = "No elements updated"
            self.status_text.Foreground = brush(RED)
            MessageBox.Show(msg, "Update Result", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        
        self.changes = []
        self.update_btn.IsEnabled = False
        self._on_preview(None, None)


if __name__ == "__main__":
    try:
        MainWindow().ShowDialog()
    except Exception as e:
        MessageBox.Show(str(e) + "\n\n" + traceback.format_exc(), "Error", 
                       MessageBoxButtons.OK, MessageBoxIcon.Error)