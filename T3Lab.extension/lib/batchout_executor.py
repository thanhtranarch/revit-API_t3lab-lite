# -*- coding: utf-8 -*-
"""
Batch Out Executor

Executes batch export operations for sheets in various formats.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Batch Out Executor"

import os
import re


# ─── Format → checkbox attribute name mapping ─────────────────────────────────

_FMT_ATTRS = {
    'pdf': 'export_pdf',
    'dwg': 'export_dwg',
    'dwf': 'export_dwf',
    'dgn': 'export_dgn',
    'nwd': 'export_nwd',
    'ifc': 'export_ifc',
    'img': 'export_images',
    'image': 'export_images',
}

_FMT_SUBFOLDER = {
    'pdf': 'PDF', 'dwg': 'DWG', 'dwf': 'DWF',
    'dgn': 'DGN', 'nwd': 'NWD', 'ifc': 'IFC',
    'img': 'Images', 'image': 'Images',
}


# ─── Public API ───────────────────────────────────────────────────────────────

def configure_batchout_window(window, config):
    """Pre-configure an ExportManagerWindow before ShowDialog().

    Args:
        window: ExportManagerWindow instance (already __init__'d).
        config: dict with keys format, filter, combine, goto_create.
    """
    fmt        = (config.get('format') or 'pdf').lower()
    filter_kw  = (config.get('filter') or '').lower().strip()
    combine    = config.get('combine', False)
    goto_create = config.get('goto_create', True)

    # 1. Select / deselect sheets
    _apply_sheet_filter(window, filter_kw)

    # 2. Enable only the requested format, disable others
    _apply_format(window, fmt, combine)

    # 3. Navigate to Create tab (index 2) if requested
    if goto_create:
        try:
            window.main_tabs.SelectedIndex = 2
        except Exception:
            pass


def direct_export(batchout_mod, config, progress_cb=None):
    """Export sheets without showing any UI.

    Args:
        batchout_mod : The loaded BatchOut script module (from _load_script).
        config       : dict with format, filter, folder, combine.
        progress_cb  : optional callable(message: str) for progress updates.

    Returns:
        (success: bool, exported_count: int, error_msg: str)
    """
    fmt        = (config.get('format') or 'pdf').lower()
    filter_kw  = (config.get('filter') or '').lower().strip()
    combine    = config.get('combine', False)
    folder     = config.get('folder') or os.path.join(
        os.path.expanduser('~'), 'Documents', 'Revit Exports')

    try:
        # Create window WITHOUT showing it — just to access all export logic
        window = batchout_mod.ExportManagerWindow()

        # Configure selections and format
        _apply_sheet_filter(window, filter_kw)
        _apply_format(window, fmt, combine)

        # Collect selected sheets
        selected = [s for s in window.all_sheets if s.IsSelected]
        if not selected:
            msg = u"Không tìm thấy sheet nào{}. Kiểm tra lại filter.".format(
                u" có prefix '{}'".format(filter_kw.upper()) if filter_kw else '')
            if progress_cb:
                progress_cb(msg)
            return False, 0, msg

        if progress_cb:
            progress_cb(u"Tìm thấy {} sheet. Đang xuất {}...".format(
                len(selected), fmt.upper()))

        # Ensure output folder exists
        export_folder = os.path.join(folder, _FMT_SUBFOLDER.get(fmt, fmt.upper()))
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)

        # Dispatch to the correct export method
        count = _run_export_method(window, fmt, selected, export_folder)

        msg = u"Xuất {} file {} thành công!\nThư mục: {}".format(
            count, fmt.upper(), export_folder)
        if progress_cb:
            progress_cb(msg)
        return True, count, msg

    except Exception as ex:
        msg = u"Lỗi khi xuất: {}".format(ex)
        if progress_cb:
            progress_cb(msg)
        return False, 0, msg


# ─── Extract export params from natural language ──────────────────────────────

def parse_export_params(raw_text):
    """Heuristically extract format and filter from a natural-language string.

    Returns:
        dict with 'format' (str) and 'filter' (str).
    """
    cmd = raw_text.lower()

    # Detect format
    fmt = 'pdf'  # sensible default
    for f in ['dwg', 'dwf', 'dgn', 'ifc', 'nwd', 'img', 'image', 'pdf']:
        if f in cmd:
            fmt = f
            break

    # Detect sheet prefix filter
    # Pattern 1: "G sheet" / "G sheets" / "G-sheet"
    m = re.search(r'\b([A-Z])\s*[-–]?\s*(?:sheet|tờ|bản vẽ)', raw_text, re.IGNORECASE)
    if m:
        return {'format': fmt, 'filter': m.group(1).upper()}

    # Pattern 2: standalone uppercase letter at word boundary (e.g. "G" in "toàn bộ G")
    m = re.search(r'\b([A-Z])\b', raw_text)
    if m:
        candidate = m.group(1)
        # Avoid false positives like "PDF" or single chars that are part of words
        if candidate not in ('P', 'D', 'W', 'F', 'I', 'N'):
            return {'format': fmt, 'filter': candidate}

    # Pattern 3: "all" / "toàn bộ" / "tất cả" → no filter
    if any(k in cmd for k in ['tat ca', 'toan bo', 'all', 'het', 'tất cả', 'toàn bộ', 'hết']):
        return {'format': fmt, 'filter': ''}

    return {'format': fmt, 'filter': ''}


# ─── Internal helpers ─────────────────────────────────────────────────────────

def _apply_sheet_filter(window, filter_kw):
    """Select sheets matching filter_kw; select all if filter_kw is empty."""
    kw = filter_kw.lower().strip()

    for item in window.all_sheets:
        label = (item.SheetNumber + ' ' + item.SheetName).lower()
        if not kw:
            item.IsSelected = True
        else:
            item.IsSelected = label.startswith(kw) or (' ' + kw) in label

    # Sync to filtered_sheets (make sure filtered_sheets reflects all_sheets)
    window.filtered_sheets = list(window.all_sheets)

    # Refresh ListView and counter if the window is already initialised
    try:
        window.sheets_listview.ItemsSource = None
        window.sheets_listview.ItemsSource = window.filtered_sheets
    except Exception:
        pass
    try:
        window.update_selection_count()
    except Exception:
        pass


def _apply_format(window, fmt, combine=False):
    """Enable only the specified export format; disable all others."""
    fmt = fmt.lower()
    for name, attr in _FMT_ATTRS.items():
        try:
            cb = getattr(window, attr)
            cb.IsChecked = (name == fmt or (fmt == 'image' and name == 'img'))
        except Exception:
            pass

    # Handle combine PDF option
    if fmt == 'pdf':
        try:
            window.combine_pdf.IsChecked = combine
        except Exception:
            pass


def _run_export_method(window, fmt, selected_items, output_folder):
    """Call the appropriate export_to_* method on the window."""
    method_map = {
        'pdf':   'export_to_pdf',
        'dwg':   'export_to_dwg',
        'dwf':   'export_to_dwf',
        'dgn':   'export_to_dgn',
        'nwd':   'export_to_nwd',
        'ifc':   'export_to_ifc',
        'img':   'export_to_images',
        'image': 'export_to_images',
    }
    method_name = method_map.get(fmt, 'export_to_pdf')
    method = getattr(window, method_name)
    return method(selected_items, output_folder) or 0
