# -*- coding: utf-8 -*-
"""
Schedule Manager Dialog

WPF event handler for ScheduleManager.xaml.
Provides two tools:
  - Excel Link: export/import Revit schedule to/from .xlsx
  - Duplicator: batch-duplicate schedules with optional view template

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__ = "Tran Tien Thanh"
__title__  = "Schedule Manager Dialog"

# ============================================================
# IMPORTS
# ============================================================
import os
import zipfile

try:
    from xml.etree import ElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET

import clr
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows import Visibility, MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Controls import (
    CheckBox as WPFCheckBox,
    ComboBoxItem,
    ListBoxItem,
    TextBlock,
)

from pyrevit import revit, DB, forms

# ============================================================
# XAML PATH
# ============================================================
_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ScheduleManager.xaml')


# ============================================================
# REVIT VERSION HELPER
# ============================================================
def _eid_int(element_id):
    """ElementId integer value - compatible across Revit 2024/2025/2026."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


# ============================================================
# SCHEDULE COLLECTOR
# ============================================================
def _get_all_schedules(doc):
    """Return sorted list of (name, ViewSchedule) for all non-internal schedules."""
    result = []
    collector = DB.FilteredElementCollector(doc)\
                  .OfClass(DB.ViewSchedule)\
                  .WhereElementIsNotElementType()\
                  .ToElements()
    for sch in collector:
        try:
            if sch.IsTitleblockRevisionSchedule:
                continue
            if sch.IsInternalKeynoteSchedule:
                continue
            result.append((sch.Name, sch))
        except Exception:
            pass
    result.sort(key=lambda x: x[0])
    return result


def _get_view_templates(doc):
    """Return sorted list of (name, View|None) for schedule-compatible view templates."""
    result = [("(None)", None)]
    collector = DB.FilteredElementCollector(doc)\
                  .OfClass(DB.View)\
                  .WhereElementIsNotElementType()\
                  .ToElements()
    for vt in collector:
        try:
            if not vt.IsTemplate:
                continue
            if vt.ViewType in (DB.ViewType.Schedule, DB.ViewType.Undefined):
                result.append((vt.Name, vt))
        except Exception:
            pass
    result.sort(key=lambda x: x[0])
    return result


# ============================================================
# SCHEDULE DATA EXTRACTION
# ============================================================

def _extract_schedule_data(schedule, doc, keep_formatting=False):
    """
    Extract header + rows from a ViewSchedule.

    keep_formatting=True  -> read cell text via GetCellText (preserves calc
                             fields and formatting, no element ID writeback).
    keep_formatting=False -> read from element parameters (supports writeback).

    Returns a dict with keys:
      schedule_name, headers, rows, element_ids, fields, from_cells
    """
    if keep_formatting:
        return _extract_from_cells(schedule)
    return _extract_from_elements(schedule, doc)


def _extract_from_cells(schedule):
    """Read cell text directly - preserves formatting, no element ID column."""
    try:
        from Autodesk.Revit.DB import SectionType
        table = schedule.GetTableData()
        body  = table.GetSectionData(SectionType.Body)
        n_rows = body.NumberOfRows
        n_cols = body.NumberOfColumns
        if n_rows == 0 or n_cols == 0:
            return None

        headers = []
        for c in range(n_cols):
            try:
                headers.append(schedule.GetCellText(SectionType.Body, 0, c) or "")
            except Exception:
                headers.append("")

        rows = []
        for r in range(1, n_rows):
            row_data = []
            for c in range(n_cols):
                try:
                    row_data.append(schedule.GetCellText(SectionType.Body, r, c) or "")
                except Exception:
                    row_data.append("")
            if any(v.strip() for v in row_data):
                rows.append(row_data)

        return {
            'schedule_name': schedule.Name,
            'headers': headers,
            'fields': [{'name': h, 'param_id': None, 'can_edit': False} for h in headers],
            'rows': rows,
            'element_ids': [],
            'from_cells': True,
        }
    except Exception:
        return None


def _extract_from_elements(schedule, doc):
    """Read from element parameters - supports writeback."""
    try:
        definition = schedule.Definition
        fields = []
        for i in range(definition.GetFieldCount()):
            f = definition.GetField(i)
            if f.IsHidden:
                continue
            param_id = None
            try:
                pid = _eid_int(f.ParameterId)
                if pid != -1:
                    param_id = pid
            except Exception:
                pass
            field_type = f.FieldType
            is_editable = field_type not in (
                DB.ScheduleFieldType.Formula,
                DB.ScheduleFieldType.Count,
                DB.ScheduleFieldType.ElementType,
            )
            fields.append({
                'name': f.GetName(),
                'param_id': param_id,
                'can_edit': is_editable,
                'field_type': field_type,
            })

        headers = [fld['name'] for fld in fields]
        collector = DB.FilteredElementCollector(doc, schedule.Id).ToElements()
        rows, element_ids = [], []

        for elem in collector:
            eid = _eid_int(elem.Id)
            element_ids.append(eid)
            row_data = [_get_param_value(elem, fld, doc) for fld in fields]
            rows.append(row_data)

        return {
            'schedule_name': schedule.Name,
            'headers': headers,
            'fields': fields,
            'rows': rows,
            'element_ids': element_ids,
            'from_cells': False,
        }
    except Exception:
        return None


def _get_param_value(elem, field, doc):
    """Read one field value from an element."""
    try:
        field_name = field['name']
        param_id   = field['param_id']
        field_type = field.get('field_type')

        elem_type = None
        try:
            type_id = elem.GetTypeId()
            if type_id and _eid_int(type_id) != -1:
                elem_type = doc.GetElement(type_id)
        except Exception:
            pass

        # ElementType fields come from the type
        if field_type == DB.ScheduleFieldType.ElementType and elem_type:
            param = None
            if param_id and param_id < 0:
                try:
                    param = elem_type.get_Parameter(DB.BuiltInParameter(param_id))
                except Exception:
                    pass
            if not param:
                param = elem_type.LookupParameter(field_name)
            return _read_param(param) if param else ""

        # Instance parameter
        param = None
        if param_id and param_id < 0:
            try:
                param = elem.get_Parameter(DB.BuiltInParameter(param_id))
            except Exception:
                pass
        if not param:
            param = elem.LookupParameter(field_name)

        # Fall back to type
        if not param and elem_type:
            if param_id and param_id < 0:
                try:
                    param = elem_type.get_Parameter(DB.BuiltInParameter(param_id))
                except Exception:
                    pass
            if not param:
                param = elem_type.LookupParameter(field_name)

        return _read_param(param) if param else ""
    except Exception:
        return ""


def _read_param(param):
    """Return display string for a Parameter."""
    try:
        val = param.AsValueString()
        if val:
            return val
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString() or ""
        if st == DB.StorageType.Integer:
            return str(param.AsInteger())
        if st == DB.StorageType.Double:
            return str(round(param.AsDouble(), 6))
        if st == DB.StorageType.ElementId:
            eid = param.AsElementId()
            if eid and _eid_int(eid) != -1:
                ref = param.Element.Document.GetElement(eid)
                if ref:
                    try:
                        return ref.Name
                    except Exception:
                        return str(_eid_int(eid))
        return ""
    except Exception:
        return ""


# ============================================================
# EXCEL EXPORT  (zipfile - no Office interop required)
# ============================================================

def _col_letter(col_num):
    """1-based column index -> Excel column letter (A, B, ..., Z, AA, ...)."""
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(col_num % 26 + ord('A')) + result
        col_num //= 26
    return result


def _export_to_xlsx(filepath, schedules_data):
    """Write one or more schedule datasets to an .xlsx file using zipfile."""
    ns_main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns_r    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    ns_ct   = 'http://schemas.openxmlformats.org/package/2006/content-types'
    ns_rel  = 'http://schemas.openxmlformats.org/package/2006/relationships'

    # Collect all shared strings
    shared  = []
    str_idx = {}

    def _si(s):
        s = str(s) if s is not None else ""
        if s not in str_idx:
            str_idx[s] = len(shared)
            shared.append(s)
        return str_idx[s]

    for sd in schedules_data:
        for h in sd.get('headers', []):
            _si(h)
        if not sd.get('from_cells', False):
            _si('Element ID')
        for row in sd.get('rows', []):
            for cell in row:
                _si(cell)

    def _to_str(el):
        try:
            return ET.tostring(el, encoding='unicode')
        except Exception:
            raw = ET.tostring(el)
            return raw.decode('utf-8') if isinstance(raw, bytes) else str(raw)

    def _workbook():
        wb = ET.Element('workbook', {'xmlns': ns_main, 'xmlns:r': ns_r})
        sh = ET.SubElement(wb, 'sheets')
        for i, sd in enumerate(schedules_data):
            name = sd['schedule_name'][:31]\
                     .replace('/', '_').replace('\\', '_').replace('*', '_')\
                     .replace('?', '_').replace('[', '_').replace(']', '_')
            ET.SubElement(sh, 'sheet', {
                'name': name, 'sheetId': str(i+1), 'r:id': 'rId{}'.format(i+1)
            })
        return _to_str(wb)

    def _styles():
        ss = ET.Element('styleSheet', {'xmlns': ns_main})
        fonts = ET.SubElement(ss, 'fonts', {'count': '2'})
        f0 = ET.SubElement(fonts, 'font')
        ET.SubElement(f0, 'sz', {'val': '11'}); ET.SubElement(f0, 'name', {'val': 'Calibri'})
        f1 = ET.SubElement(fonts, 'font')
        ET.SubElement(f1, 'b'); ET.SubElement(f1, 'sz', {'val': '11'})
        ET.SubElement(f1, 'name', {'val': 'Calibri'})
        fills = ET.SubElement(ss, 'fills', {'count': '3'})
        ET.SubElement(ET.SubElement(fills, 'fill'), 'patternFill', {'patternType': 'none'})
        ET.SubElement(ET.SubElement(fills, 'fill'), 'patternFill', {'patternType': 'gray125'})
        hf = ET.SubElement(fills, 'fill')
        pf = ET.SubElement(hf, 'patternFill', {'patternType': 'solid'})
        ET.SubElement(pf, 'fgColor', {'rgb': 'FFCDCDCD'})
        borders = ET.SubElement(ss, 'borders', {'count': '1'})
        b0 = ET.SubElement(borders, 'border')
        ET.SubElement(b0, 'left'); ET.SubElement(b0, 'right')
        ET.SubElement(b0, 'top'); ET.SubElement(b0, 'bottom')
        csx = ET.SubElement(ss, 'cellStyleXfs', {'count': '1'})
        ET.SubElement(csx, 'xf', {'numFmtId': '0', 'fontId': '0', 'fillId': '0', 'borderId': '0'})
        cx = ET.SubElement(ss, 'cellXfs', {'count': '2'})
        ET.SubElement(cx, 'xf', {
            'numFmtId': '0', 'fontId': '0', 'fillId': '0', 'borderId': '0', 'xfId': '0'
        })
        ET.SubElement(cx, 'xf', {
            'numFmtId': '0', 'fontId': '1', 'fillId': '2', 'borderId': '0',
            'xfId': '0', 'applyFont': '1', 'applyFill': '1'
        })
        cs = ET.SubElement(ss, 'cellStyles', {'count': '1'})
        ET.SubElement(cs, 'cellStyle', {'name': 'Normal', 'xfId': '0', 'builtinId': '0'})
        return _to_str(ss)

    def _sst():
        root = ET.Element('sst', {
            'xmlns': ns_main,
            'count': str(len(shared)),
            'uniqueCount': str(len(shared)),
        })
        for s in shared:
            si = ET.SubElement(root, 'si')
            t  = ET.SubElement(si, 't')
            t.text = s
        return _to_str(root)

    def _worksheet(sd):
        ws  = ET.Element('worksheet', {'xmlns': ns_main})
        shd = ET.SubElement(ws, 'sheetData')
        headers    = sd.get('headers', [])
        rows       = sd.get('rows', [])
        eids       = sd.get('element_ids', [])
        from_cells = sd.get('from_cells', False)

        # Header row
        hdr_el = ET.SubElement(shd, 'row', {'r': '1'})
        if from_cells:
            for c, h in enumerate(headers):
                cel = ET.SubElement(hdr_el, 'c', {
                    'r': '{}{}'.format(_col_letter(c+1), 1), 't': 's', 's': '1'
                })
                ET.SubElement(cel, 'v').text = str(_si(h))
        else:
            cel = ET.SubElement(hdr_el, 'c', {'r': 'A1', 't': 's', 's': '1'})
            ET.SubElement(cel, 'v').text = str(_si('Element ID'))
            for c, h in enumerate(headers):
                cel = ET.SubElement(hdr_el, 'c', {
                    'r': '{}{}'.format(_col_letter(c+2), 1), 't': 's', 's': '1'
                })
                ET.SubElement(cel, 'v').text = str(_si(h))

        # Data rows
        for r_idx, row in enumerate(rows):
            excel_row = r_idx + 2
            row_el = ET.SubElement(shd, 'row', {'r': str(excel_row)})
            if from_cells:
                for c, val in enumerate(row):
                    cel = ET.SubElement(row_el, 'c', {
                        'r': '{}{}'.format(_col_letter(c+1), excel_row), 't': 's'
                    })
                    ET.SubElement(cel, 'v').text = str(_si(val))
            else:
                eid = eids[r_idx] if r_idx < len(eids) else None
                cel = ET.SubElement(row_el, 'c', {'r': 'A{}'.format(excel_row)})
                ET.SubElement(cel, 'v').text = str(eid) if eid else '0'
                for c, val in enumerate(row):
                    cell_ref = '{}{}'.format(_col_letter(c+2), excel_row)
                    val_str = str(val) if val is not None else ""
                    is_num = False
                    try:
                        if val_str.strip():
                            float(val_str)
                            is_num = True
                    except Exception:
                        pass
                    if is_num:
                        cel = ET.SubElement(row_el, 'c', {'r': cell_ref})
                        ET.SubElement(cel, 'v').text = val_str
                    else:
                        cel = ET.SubElement(row_el, 'c', {'r': cell_ref, 't': 's'})
                        ET.SubElement(cel, 'v').text = str(_si(val_str))
        return _to_str(ws)

    def _rels():
        r = ET.Element('Relationships', {'xmlns': ns_rel})
        ET.SubElement(r, 'Relationship', {
            'Id': 'rId1',
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
            'Target': 'xl/workbook.xml',
        })
        return _to_str(r)

    def _wb_rels():
        r = ET.Element('Relationships', {'xmlns': ns_rel})
        for i in range(len(schedules_data)):
            ET.SubElement(r, 'Relationship', {
                'Id': 'rId{}'.format(i+1),
                'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                'Target': 'worksheets/sheet{}.xml'.format(i+1),
            })
        ET.SubElement(r, 'Relationship', {
            'Id': 'rId{}'.format(len(schedules_data)+1),
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
            'Target': 'sharedStrings.xml',
        })
        ET.SubElement(r, 'Relationship', {
            'Id': 'rId{}'.format(len(schedules_data)+2),
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
            'Target': 'styles.xml',
        })
        return _to_str(r)

    def _ct():
        t = ET.Element('Types', {'xmlns': ns_ct})
        ET.SubElement(t, 'Default', {
            'Extension': 'rels',
            'ContentType': 'application/vnd.openxmlformats-package.relationships+xml',
        })
        ET.SubElement(t, 'Default', {
            'Extension': 'xml', 'ContentType': 'application/xml',
        })
        ET.SubElement(t, 'Override', {
            'PartName': '/xl/workbook.xml',
            'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        })
        ET.SubElement(t, 'Override', {
            'PartName': '/xl/sharedStrings.xml',
            'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
        })
        ET.SubElement(t, 'Override', {
            'PartName': '/xl/styles.xml',
            'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
        })
        for i in range(len(schedules_data)):
            ET.SubElement(t, 'Override', {
                'PartName': '/xl/worksheets/sheet{}.xml'.format(i+1),
                'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
            })
        return _to_str(t)

    decl = '<?xml version="1.0" encoding="UTF-8"?>'
    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml',           decl + _ct())
        zf.writestr('_rels/.rels',                   decl + _rels())
        zf.writestr('xl/workbook.xml',               decl + _workbook())
        zf.writestr('xl/_rels/workbook.xml.rels',    decl + _wb_rels())
        zf.writestr('xl/sharedStrings.xml',          decl + _sst())
        zf.writestr('xl/styles.xml',                 decl + _styles())
        for i, sd in enumerate(schedules_data):
            zf.writestr(
                'xl/worksheets/sheet{}.xml'.format(i+1),
                decl + _worksheet(sd)
            )


# ============================================================
# EXCEL IMPORT  (zipfile)
# ============================================================

def _import_from_xlsx(filepath):
    """
    Read the first non-hidden sheet from an xlsx file.

    Returns (data_dict, error_str).
    data_dict keys: schedule_name, headers, rows, element_ids
    """
    try:
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        with zipfile.ZipFile(filepath, 'r') as zf:
            # Shared strings
            shared = []
            try:
                ss_root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
                for si in ss_root.findall('.//{%s}si' % ns):
                    t = si.find('.//{%s}t' % ns)
                    shared.append(t.text if t is not None and t.text else "")
            except Exception:
                pass

            # Sheet names
            wb_root     = ET.fromstring(zf.read('xl/workbook.xml'))
            sheets_info = [
                sh.get('name', '')
                for sh in wb_root.findall('.//{%s}sheet' % ns)
            ]

            schedule_name = sheets_info[0] if sheets_info else "Schedule"

            def _read_sheet(idx):
                sheet_xml  = zf.read('xl/worksheets/sheet{}.xml'.format(idx+1))
                sheet_root = ET.fromstring(sheet_xml)
                rows_data  = {}
                for row_el in sheet_root.findall('.//{%s}row' % ns):
                    rn    = int(row_el.get('r', 0))
                    cells = {}
                    for cell_el in row_el.findall('.//{%s}c' % ns):
                        cell_ref  = cell_el.get('r', '')
                        cell_type = cell_el.get('t', '')
                        v_el      = cell_el.find('{%s}v' % ns)
                        if v_el is not None and v_el.text:
                            if cell_type == 's':
                                idx2 = int(v_el.text)
                                value = shared[idx2] if idx2 < len(shared) else ""
                            else:
                                value = v_el.text
                        else:
                            value = ""
                        col_str = ''.join(c for c in cell_ref if c.isalpha())
                        col_num = 0
                        for ch in col_str:
                            col_num = col_num * 26 + (ord(ch.upper()) - ord('A') + 1)
                        cells[col_num] = value
                    rows_data[rn] = cells
                return rows_data

            data_rows = _read_sheet(0)

        if not data_rows:
            return None, "No data found in file."

        header_row = data_rows.get(1, {})
        h1 = header_row.get(1, "").strip().lower()
        if h1 in ('element id', 'elementid', 'id', 'element_id'):
            eid_col    = 1
            data_start = 2
        else:
            eid_col    = None
            data_start = 1

        # Collect headers
        headers = []
        col = data_start
        while True:
            h = header_row.get(col, "").strip()
            if not h:
                if col > data_start + 5:
                    break
            else:
                headers.append(h)
            col += 1
            if col > data_start + 200:
                break

        rows, element_ids = [], []
        row_nums = sorted(r for r in data_rows if r > 1)
        for rn in row_nums:
            row_data = data_rows[rn]
            if eid_col:
                eid_val = row_data.get(eid_col, "")
                try:
                    element_ids.append(int(float(eid_val)) if eid_val else None)
                except Exception:
                    element_ids.append(None)
            row_values = [row_data.get(data_start + ci, "") for ci in range(len(headers))]
            if any(v for v in row_values):
                rows.append(row_values)
            elif eid_col and element_ids and element_ids[-1]:
                rows.append(row_values)

        return {
            'schedule_name': schedule_name,
            'headers':       headers,
            'rows':          rows,
            'element_ids':   element_ids,
        }, None

    except Exception as ex:
        return None, "Import error: {}".format(ex)


# ============================================================
# CHANGE DETECTION + MODEL UPDATER
# ============================================================

def _find_changes(current_data, imported_data):
    """
    Compare imported xlsx data against live Revit data.
    Returns list of change dicts (element_id, field_name, param_id,
                                  old_value, new_value).
    """
    changes = []
    if not current_data or not imported_data:
        return changes

    cur_headers = current_data['headers']
    cur_col_map = {h: i for i, h in enumerate(cur_headers)}
    cur_fields  = current_data.get('fields', [])

    cur_by_eid = {}
    for ri, eid in enumerate(current_data['element_ids']):
        if eid:
            cur_by_eid[eid] = current_data['rows'][ri]

    imp_headers = imported_data.get('headers', [])

    for imp_idx, imp_row in enumerate(imported_data['rows']):
        imp_eid = (imported_data['element_ids'][imp_idx]
                   if imp_idx < len(imported_data['element_ids']) else None)
        if not imp_eid:
            continue
        cur_row = cur_by_eid.get(imp_eid)
        if cur_row is None:
            continue

        for ci, header in enumerate(imp_headers):
            cur_ci = cur_col_map.get(header)
            if cur_ci is None:
                continue
            # Skip non-editable fields
            if cur_ci < len(cur_fields) and not cur_fields[cur_ci].get('can_edit', True):
                continue
            param_id = cur_fields[cur_ci].get('param_id') if cur_ci < len(cur_fields) else None

            imp_val = str(imp_row[ci]).strip() if ci < len(imp_row) else ""
            cur_val = str(cur_row[cur_ci]).strip() if cur_ci < len(cur_row) else ""

            if imp_val != cur_val:
                changes.append({
                    'element_id': imp_eid,
                    'field_name': header,
                    'param_id':   param_id,
                    'old_value':  cur_val,
                    'new_value':  imp_val,
                })

    return changes


def _apply_changes(changes, doc):
    """
    Write changes back to Revit elements inside a single transaction.
    Returns (success_count, error_list, skipped_count).
    """
    if not changes:
        return 0, [], 0

    success, errors, skipped = 0, [], 0

    t = DB.Transaction(doc, "T3Lab: Schedule Manager - Apply Changes")
    t.Start()
    try:
        for ch in changes:
            eid = ch['element_id']
            if not eid:
                skipped += 1
                continue
            try:
                elem = doc.GetElement(DB.ElementId(eid))
            except Exception:
                elem = None
            if not elem:
                skipped += 1
                continue

            ok, err_msg = _set_param_value(elem, ch, doc)
            if ok:
                success += 1
            else:
                errors.append("ID {}: {}".format(eid, err_msg))

        t.Commit()
    except Exception as ex:
        t.RollBack()
        return 0, [str(ex)], 0

    return success, errors, skipped


def _set_param_value(elem, change, doc):
    """Set a single parameter value on an element. Returns (bool, error_str)."""
    try:
        field_name = change['field_name']
        param_id   = change['param_id']
        value      = change['new_value']

        param = None
        if param_id and param_id < 0:
            try:
                param = elem.get_Parameter(DB.BuiltInParameter(param_id))
            except Exception:
                pass
        if not param:
            param = elem.LookupParameter(field_name)

        # Try element type as fallback
        if not param:
            elem_type = None
            try:
                type_id = elem.GetTypeId()
                if type_id and _eid_int(type_id) != -1:
                    elem_type = doc.GetElement(type_id)
            except Exception:
                pass
            if elem_type:
                if param_id and param_id < 0:
                    try:
                        param = elem_type.get_Parameter(DB.BuiltInParameter(param_id))
                    except Exception:
                        pass
                if not param:
                    param = elem_type.LookupParameter(field_name)

        if not param:
            return False, "'{}' not found".format(field_name)
        if param.IsReadOnly:
            return False, "'{}' is read-only".format(field_name)

        st = param.StorageType
        if st == DB.StorageType.String:
            param.Set(str(value) if value else "")
        elif st == DB.StorageType.Integer:
            param.Set(int(float(value)) if value else 0)
        elif st == DB.StorageType.Double:
            param.Set(float(value) if value else 0.0)
        else:
            return False, "Unsupported StorageType"

        return True, ""
    except Exception as ex:
        return False, str(ex)


# ============================================================
# PREVIEW RENDERER
# ============================================================

def _render_preview(container, all_data):
    """
    Render lightweight preview tables into a WPF StackPanel (ctr_grid).
    Shows up to 50 rows per schedule.
    """
    from System.Windows.Controls import (
        Grid as WPFGrid, ColumnDefinition, RowDefinition,
        Border as WPFBorder,
    )
    from System.Windows.Media import BrushConverter

    MAX_ROWS = 50
    bc = BrushConverter()

    def _b(hex_str):
        try:
            return bc.ConvertFromString(hex_str)
        except Exception:
            return None

    container.Children.Clear()

    for data in all_data:
        headers = data.get('headers', [])
        rows    = data.get('rows', [])
        name    = data.get('schedule_name', '')

        # Schedule name label
        lbl = TextBlock()
        lbl.Text       = name
        lbl.FontSize   = 13
        lbl.FontWeight = System.Windows.FontWeights.SemiBold
        lbl.Foreground = _b('#18181B')
        lbl.Margin     = System.Windows.Thickness(0, 8, 0, 4)
        container.Children.Add(lbl)

        n_cols = len(headers)
        if n_cols == 0:
            continue

        grid = WPFGrid()
        grid.Margin = System.Windows.Thickness(0, 0, 0, 12)

        for _ in range(n_cols):
            cd = ColumnDefinition()
            cd.MinWidth = 80
            grid.ColumnDefinitions.Add(cd)

        def _add_cell(g, text, row, col, is_header):
            brd = WPFBorder()
            brd.BorderBrush     = _b('#E4E4E7')
            brd.BorderThickness = System.Windows.Thickness(0, 0, 1, 1)
            brd.Background      = _b('#F4F4F6') if is_header else _b('#FFFFFF')
            brd.Padding         = System.Windows.Thickness(6, 3, 6, 3)
            tb = TextBlock()
            tb.Text       = str(text) if text is not None else ""
            tb.FontSize   = 11
            tb.FontWeight = (System.Windows.FontWeights.SemiBold
                             if is_header else System.Windows.FontWeights.Normal)
            tb.Foreground = _b('#18181B')
            brd.Child = tb
            WPFGrid.SetRow(brd, row)
            WPFGrid.SetColumn(brd, col)
            g.Children.Add(brd)

        # Header row
        grid.RowDefinitions.Add(RowDefinition())
        for ci, h in enumerate(headers):
            _add_cell(grid, h, 0, ci, is_header=True)

        # Data rows
        display_rows = rows[:MAX_ROWS]
        for ri, row in enumerate(display_rows):
            grid.RowDefinitions.Add(RowDefinition())
            for ci in range(n_cols):
                val = row[ci] if ci < len(row) else ""
                _add_cell(grid, val, ri+1, ci, is_header=False)

        if len(rows) > MAX_ROWS:
            grid.RowDefinitions.Add(RowDefinition())
            more = TextBlock()
            more.Text      = "... {} more row(s) not shown".format(len(rows) - MAX_ROWS)
            more.FontSize  = 11
            more.Foreground = _b('#71717A')
            more.Margin    = System.Windows.Thickness(4, 2, 0, 2)
            WPFGrid.SetRow(more, len(display_rows)+1)
            WPFGrid.SetColumnSpan(more, n_cols)
            grid.Children.Add(more)

        container.Children.Add(grid)


# ============================================================
# MAIN WINDOW CLASS
# ============================================================

class ScheduleManagerWindow(forms.WPFWindow):
    """
    WPF host for ScheduleManager.xaml.

    Excel Link tab  - export to / import from xlsx, apply changes back to Revit.
    Duplicator tab  - batch-duplicate schedules with optional view template.
    """

    def __init__(self, script_dir, revit_obj):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit      = revit_obj
        self._doc        = revit_obj.ActiveUIDocument.Document

        # ---- Excel Link state ----
        self._all_schedules   = []   # [(name, ViewSchedule), ...]
        self._current_data    = None # dict from _extract_schedule_data
        self._imported_data   = None # dict from _import_from_xlsx
        self._pending_changes = []

        # ---- Duplicator state ----
        self._dup_schedules = []     # [(name, ViewSchedule), ...]
        self._dup_templates = []     # [(name, View|None), ...]

        # Wire window chrome
        self.btn_minimize.Click     += self._minimize
        self.btn_maximize.Click     += self._maximize
        self.btn_close_chrome.Click += self._close

        # Wire sidebar navigation
        self.nav_excel_link.Click += self._nav_excel_link_clicked
        self.nav_duplicator.Click += self._nav_duplicator_clicked

        # Wire Excel Link controls
        self.cmb_schedule.SelectionChanged += self._on_schedule_changed
        self.chk_multi_mode.Checked        += self._on_multi_mode_toggled
        self.chk_multi_mode.Unchecked      += self._on_multi_mode_toggled
        self.btn_preview.Click             += self._on_preview
        self.btn_export.Click              += self._on_export
        self.btn_import.Click              += self._on_import
        self.btn_update.Click              += self._on_update

        # Wire Duplicator controls
        self.btn_dup_run.Click += self._on_dup_run

        # Populate data
        self._load_all_schedules()
        self._populate_dup_controls()
        self._set_status("Ready")

    # ------------------------------------------------------------------
    # WINDOW CHROME
    # ------------------------------------------------------------------

    def _minimize(self, sender, e):
        self.WindowState = System.Windows.WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == System.Windows.WindowState.Maximized:
            self.WindowState = System.Windows.WindowState.Normal
        else:
            self.WindowState = System.Windows.WindowState.Maximized

    def _close(self, sender, e):
        self.Close()

    # Aliases so XAML Click="..." attributes also work when pyRevit
    # routes them to method names matching the XAML attribute value.
    def minimize_button_clicked(self, sender, e):
        self._minimize(sender, e)

    def maximize_button_clicked(self, sender, e):
        self._maximize(sender, e)

    def close_button_clicked(self, sender, e):
        self._close(sender, e)

    # ------------------------------------------------------------------
    # SIDEBAR NAVIGATION
    # ------------------------------------------------------------------

    def _nav_excel_link_clicked(self, sender, e):
        self.tab_main.SelectedItem = self.tab_excel_link

    def _nav_duplicator_clicked(self, sender, e):
        self.tab_main.SelectedItem = self.tab_duplicator

    # XAML Click= aliases
    def nav_excel_link_clicked(self, sender, e):
        self._nav_excel_link_clicked(sender, e)

    def nav_duplicator_clicked(self, sender, e):
        self._nav_duplicator_clicked(sender, e)

    # ------------------------------------------------------------------
    # STATUS / INFO HELPERS
    # ------------------------------------------------------------------

    def _set_status(self, text):
        try:
            self.txt_statusbar.Text = text
        except Exception:
            pass

    def _set_excel_status(self, text):
        try:
            self.txt_excel_status.Text = text
        except Exception:
            pass

    def _set_info(self, text):
        try:
            self.txt_info.Text = text
        except Exception:
            pass

    def _set_progress(self, value, visible=True):
        try:
            self.prg_excel.Value      = value
            self.prg_excel.Visibility = (Visibility.Visible if visible
                                         else Visibility.Collapsed)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # LOAD SCHEDULES (Excel Link tab)
    # ------------------------------------------------------------------

    def _load_all_schedules(self):
        """Populate cmb_schedule and pnl_schedule_list."""
        self._all_schedules = _get_all_schedules(self._doc)
        self.cmb_schedule.Items.Clear()
        self.pnl_schedule_list.Children.Clear()

        for name, _ in self._all_schedules:
            # Single-select ComboBox
            ci = ComboBoxItem()
            ci.Content = name
            self.cmb_schedule.Items.Add(ci)

            # Multi-select checklist
            chk = WPFCheckBox()
            chk.Content = name
            try:
                chk.Style = self.FindResource('T3CheckBox')
            except Exception:
                pass
            chk.Margin = System.Windows.Thickness(0, 2, 0, 2)
            self.pnl_schedule_list.Children.Add(chk)

        if self._all_schedules:
            self.cmb_schedule.SelectedIndex = 0

        n = len(self._all_schedules)
        self._set_info("{} schedule(s) found in the project.".format(n))
        self._set_status("Loaded {} schedule(s).".format(n))

    # ------------------------------------------------------------------
    # EXCEL LINK - SELECTION HELPER
    # ------------------------------------------------------------------

    def _get_selected_schedules(self):
        """
        Return list of (name, ViewSchedule) for the active selection.
        Single mode: ComboBox selection.
        Multi mode: all checked items in pnl_schedule_list.
        """
        if self.chk_multi_mode.IsChecked:
            selected = []
            for chk in self.pnl_schedule_list.Children:
                if isinstance(chk, WPFCheckBox) and chk.IsChecked:
                    name  = chk.Content
                    match = [sch for n, sch in self._all_schedules if n == name]
                    if match:
                        selected.append((name, match[0]))
            return selected
        else:
            idx = self.cmb_schedule.SelectedIndex
            if 0 <= idx < len(self._all_schedules):
                return [self._all_schedules[idx]]
            return []

    # ------------------------------------------------------------------
    # EXCEL LINK - EVENT HANDLERS
    # ------------------------------------------------------------------

    def _on_schedule_changed(self, sender, e):
        """Clear stale preview / import state when user picks a new schedule."""
        self._current_data    = None
        self._imported_data   = None
        self._pending_changes = []
        self.btn_update.IsEnabled = False
        try:
            self.ctr_grid.Children.Clear()
        except Exception:
            pass
        self._set_excel_status("")

    def _on_multi_mode_toggled(self, sender, e):
        """Show or hide the multi-schedule checklist (pnl_multi_list)."""
        if self.chk_multi_mode.IsChecked:
            self.pnl_multi_list.Visibility = Visibility.Visible
        else:
            self.pnl_multi_list.Visibility = Visibility.Collapsed

    def _on_preview(self, sender, e):
        selected = self._get_selected_schedules()
        if not selected:
            MessageBox.Show("Please select at least one schedule.", "No Selection",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        keep_fmt = bool(self.chk_keep_formatting.IsChecked)
        self._set_progress(0, True)
        self._set_excel_status("Loading preview...")

        try:
            all_data = []
            for i, (name, sch) in enumerate(selected):
                data = _extract_schedule_data(sch, self._doc, keep_formatting=keep_fmt)
                if data:
                    all_data.append(data)
                self._set_progress(int((i+1) * 80.0 / len(selected)))

            if not all_data:
                self._set_excel_status("No data could be extracted from the selected schedule(s).")
                self._set_progress(0, False)
                return

            # Cache for single-schedule change detection
            if len(all_data) == 1 and not self.chk_multi_mode.IsChecked:
                self._current_data = all_data[0]

            _render_preview(self.ctr_grid, all_data)
            self._set_progress(100)
            total_rows = sum(len(d['rows']) for d in all_data)
            self._set_excel_status(
                "Preview: {} schedule(s), {} row(s).".format(len(all_data), total_rows)
            )
            self._set_status("Preview loaded.")
        except Exception as ex:
            self._set_excel_status("Preview error: {}".format(ex))
        finally:
            self._set_progress(0, False)

    def _on_export(self, sender, e):
        selected = self._get_selected_schedules()
        if not selected:
            MessageBox.Show("Please select at least one schedule.", "No Selection",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        default_name = (selected[0][0] if len(selected) == 1 else "Schedules_Export") + '.xlsx'
        save_path = forms.save_file(file_ext='xlsx', default_name=default_name)
        if not save_path:
            return

        keep_fmt = bool(self.chk_keep_formatting.IsChecked)
        self._set_progress(0, True)
        self._set_excel_status("Exporting...")

        try:
            all_data = []
            for i, (name, sch) in enumerate(selected):
                data = _extract_schedule_data(sch, self._doc, keep_formatting=keep_fmt)
                if data:
                    all_data.append(data)
                self._set_progress(int((i+1) * 70.0 / len(selected)))

            if not all_data:
                self._set_excel_status("No data to export.")
                self._set_progress(0, False)
                return

            _export_to_xlsx(save_path, all_data)

            if len(all_data) == 1 and not self.chk_multi_mode.IsChecked:
                self._current_data = all_data[0]

            self._set_progress(100)
            self._set_excel_status("Exported {} schedule(s) to: {}".format(
                len(all_data), os.path.basename(save_path)))
            self._set_status("Export complete.")
            MessageBox.Show(
                "Exported {} schedule(s) successfully.".format(len(all_data)),
                "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information
            )
        except Exception as ex:
            self._set_excel_status("Export error: {}".format(ex))
            MessageBox.Show("Export failed:\n\n{}".format(ex), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)
        finally:
            self._set_progress(0, False)

    def _on_import(self, sender, e):
        if self.chk_multi_mode.IsChecked:
            MessageBox.Show(
                "Import is only supported in single-schedule mode.\n"
                "Please uncheck 'Multi-Schedule' and select one schedule.",
                "Multi-Mode Not Supported",
                MessageBoxButton.OK, MessageBoxImage.Warning
            )
            return

        open_path = forms.pick_file(file_ext='xlsx')
        if not open_path:
            return

        self._set_excel_status("Importing: {}".format(os.path.basename(open_path)))
        self._set_progress(20, True)

        try:
            data, err = _import_from_xlsx(open_path)
            if err:
                self._set_excel_status("Import error: {}".format(err))
                MessageBox.Show("Import failed:\n\n{}".format(err), "Error",
                                MessageBoxButton.OK, MessageBoxImage.Error)
                return

            self._imported_data = data

            # Ensure we have live current data to diff against
            if not self._current_data:
                idx = self.cmb_schedule.SelectedIndex
                if 0 <= idx < len(self._all_schedules):
                    _, sch = self._all_schedules[idx]
                    self._current_data = _extract_schedule_data(
                        sch, self._doc, keep_formatting=False)

            if not self._current_data:
                self._set_excel_status(
                    "Import loaded but no live schedule data available for comparison. "
                    "Click Preview first, then Import.")
                self._set_progress(0, False)
                return

            self._set_progress(60)
            self._pending_changes = _find_changes(self._current_data, self._imported_data)
            n = len(self._pending_changes)
            self.btn_update.IsEnabled = n > 0
            self._set_progress(100)
            self._set_excel_status(
                "Imported: {} row(s), {} change(s) detected.".format(
                    len(data['rows']), n)
            )
            self._set_status("Import complete.")
        except Exception as ex:
            self._set_excel_status("Import error: {}".format(ex))
            MessageBox.Show("Import error:\n\n{}".format(ex), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)
        finally:
            self._set_progress(0, False)

    def _on_update(self, sender, e):
        if not self._pending_changes:
            MessageBox.Show("No pending changes to apply.", "Info",
                            MessageBoxButton.OK, MessageBoxImage.Information)
            return

        n = len(self._pending_changes)
        result = MessageBox.Show(
            "Apply {} change(s) to the Revit model?".format(n),
            "Confirm Apply", MessageBoxButton.YesNo, MessageBoxImage.Warning
        )
        if result != System.Windows.MessageBoxResult.Yes:
            return

        self._set_progress(0, True)
        self._set_excel_status("Applying {} change(s)...".format(n))

        try:
            success, errors, skipped = _apply_changes(self._pending_changes, self._doc)
            self._pending_changes = []
            self.btn_update.IsEnabled = False
            self._set_progress(100)

            msg = "Applied {} / {} change(s).".format(success, n)
            if skipped:
                msg += " {} skipped (element not found).".format(skipped)
            self._set_excel_status(msg)
            self._set_status("Update complete.")

            if errors:
                detail = "\n".join(errors[:20])
                if len(errors) > 20:
                    detail += "\n... and {} more".format(len(errors) - 20)
                MessageBox.Show("Some changes could not be applied:\n\n{}".format(detail),
                                "Partial Success", MessageBoxButton.OK, MessageBoxImage.Warning)
            else:
                MessageBox.Show(msg, "Update Complete",
                                MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            self._set_excel_status("Update error: {}".format(ex))
            MessageBox.Show("Update failed:\n\n{}".format(ex), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)
        finally:
            self._set_progress(0, False)

    # XAML Click= aliases for Excel Link tab buttons
    def btn_preview_clicked(self, sender, e):
        self._on_preview(sender, e)

    def btn_export_clicked(self, sender, e):
        self._on_export(sender, e)

    def btn_import_clicked(self, sender, e):
        self._on_import(sender, e)

    def btn_update_clicked(self, sender, e):
        self._on_update(sender, e)

    # ------------------------------------------------------------------
    # DUPLICATOR - POPULATE CONTROLS
    # ------------------------------------------------------------------

    def _populate_dup_controls(self):
        """Fill lst_dup_schedules, cmb_dup_method, cmb_dup_template."""
        self._dup_schedules = _get_all_schedules(self._doc)
        self.lst_dup_schedules.Items.Clear()
        for name, _ in self._dup_schedules:
            item = ListBoxItem()
            item.Content = name
            self.lst_dup_schedules.Items.Add(item)

        self.cmb_dup_method.Items.Clear()
        for method in ("Duplicate", "As Dependent", "As Independent"):
            ci = ComboBoxItem()
            ci.Content = method
            self.cmb_dup_method.Items.Add(ci)
        self.cmb_dup_method.SelectedIndex = 0

        self.txt_dup_basename.Text = ""

        self._dup_templates = _get_view_templates(self._doc)
        self.cmb_dup_template.Items.Clear()
        for name, _ in self._dup_templates:
            ci = ComboBoxItem()
            ci.Content = name
            self.cmb_dup_template.Items.Add(ci)
        if self._dup_templates:
            self.cmb_dup_template.SelectedIndex = 0

    # ------------------------------------------------------------------
    # DUPLICATOR - RUN
    # ------------------------------------------------------------------

    def _dup_log(self, text):
        try:
            cur = self.txt_dup_log.Text or ""
            self.txt_dup_log.Text = cur + text + "\n"
        except Exception:
            pass

    def _set_dup_status(self, text):
        try:
            self.txt_dup_status.Text = text
        except Exception:
            pass

    def _on_dup_run(self, sender, e):
        # Collect selected schedules from list box
        selected_items = [
            self._dup_schedules[i]
            for i in range(self.lst_dup_schedules.Items.Count)
            if (i < len(self._dup_schedules)
                and self.lst_dup_schedules.Items[i].IsSelected)
        ]
        if not selected_items:
            MessageBox.Show("Please select at least one schedule to duplicate.",
                            "No Selection", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        # Copy count
        count_str = (self.txt_dup_count.Text or "1").strip()
        try:
            copy_count = int(count_str)
            if copy_count < 1:
                raise ValueError
        except ValueError:
            MessageBox.Show("Copy Count must be a positive integer.", "Invalid Input",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        # Base name
        base_name = (self.txt_dup_basename.Text or "").strip()

        # Duplication method
        method_map = {
            0: DB.ViewDuplicateOption.Duplicate,
            1: DB.ViewDuplicateOption.AsDependent,
            2: DB.ViewDuplicateOption.AsIndependent,
        }
        dup_option = method_map.get(self.cmb_dup_method.SelectedIndex,
                                    DB.ViewDuplicateOption.Duplicate)

        # View template (index 0 = "(None)")
        tmpl_idx  = self.cmb_dup_template.SelectedIndex
        tmpl_obj  = None
        if 0 <= tmpl_idx < len(self._dup_templates):
            _, tmpl_obj = self._dup_templates[tmpl_idx]

        # Clear log
        self.txt_dup_log.Text = ""
        self._set_dup_status("Running...")
        self._set_status("Duplicating schedules...")

        total_created = 0
        total_errors  = 0

        try:
            tg = DB.TransactionGroup(
                self._doc, "T3Lab: Schedule Manager - Duplicate Schedules")
            tg.Start()

            for sch_name, sch in selected_items:
                name_base = base_name if base_name else sch_name + "_Copy"
                self._dup_log("Duplicating '{}' x{}...".format(sch_name, copy_count))

                for i in range(copy_count):
                    t = DB.Transaction(self._doc, "T3Lab: Duplicate Schedule")
                    t.Start()
                    try:
                        new_id  = sch.Duplicate(dup_option)
                        new_sch = self._doc.GetElement(new_id)

                        # Build target name
                        if copy_count == 1 and len(selected_items) == 1:
                            target_name = name_base
                        else:
                            target_name = "{}_{}".format(name_base, i+1)

                        # Ensure uniqueness
                        used_name = target_name
                        suffix    = 1
                        while suffix <= 50:
                            try:
                                new_sch.Name = used_name
                                break
                            except Exception:
                                used_name = "{}_{}".format(target_name, suffix)
                                suffix += 1

                        # Apply view template
                        if tmpl_obj:
                            try:
                                if tmpl_obj.IsValidViewTemplate(new_sch):
                                    new_sch.ViewTemplateId = tmpl_obj.Id
                                else:
                                    self._dup_log(
                                        "  [warn] Template not compatible with '{}'".format(
                                            used_name))
                            except Exception as tex:
                                self._dup_log("  [warn] Template apply failed: {}".format(tex))

                        t.Commit()
                        total_created += 1
                        self._dup_log("  Created: '{}'".format(used_name))

                    except Exception as ex:
                        t.RollBack()
                        total_errors += 1
                        self._dup_log(
                            "  [error] Copy {} of '{}': {}".format(i+1, sch_name, ex))

            tg.Assimilate()

        except Exception as ex:
            self._dup_log("[fatal] {}".format(ex))
            total_errors += 1
            try:
                tg.RollBack()
            except Exception:
                pass

        status_msg = "Done: {} created, {} failed.".format(total_created, total_errors)
        self._set_dup_status(status_msg)
        self._set_status(status_msg)

        if total_errors == 0:
            MessageBox.Show(
                "Duplicated {} schedule(s) successfully.".format(total_created),
                "Done", MessageBoxButton.OK, MessageBoxImage.Information
            )
        else:
            MessageBox.Show(
                "{} schedule(s) created. {} error(s) - see log.".format(
                    total_created, total_errors),
                "Completed with Errors",
                MessageBoxButton.OK, MessageBoxImage.Warning
            )

    # XAML Click= alias for Duplicator run button
    def btn_dup_run_clicked(self, sender, e):
        self._on_dup_run(sender, e)


# ============================================================
# PUBLIC ENTRY POINT
# ============================================================

def show_schedule_manager(script_dir, revit_obj):
    """Open the Schedule Manager window modally."""
    ScheduleManagerWindow(script_dir, revit_obj).ShowDialog()
