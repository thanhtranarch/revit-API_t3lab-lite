#!/usr/bin/env python3
"""
Generate SVG icons for all T3Lab Revit extension tools.
Icon style matches the T3Lab UI Showcase design system:
  - viewBox="0 0 24 24"
  - fill="none"
  - stroke="#18181B" (Ink color)
  - stroke-width="1.7"
  - stroke-linecap="round"
  - stroke-linejoin="round"
  - 96x96 px rendered size (Revit large button icon)
"""

import os
import subprocess
import sys

BASE = r"c:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\T3Lab.extension\T3Lab.tab"
SVG_OUT = r"c:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\scripts\svg_icons"

os.makedirs(SVG_OUT, exist_ok=True)

# -------------------------------------------------
# SVG helper
# -------------------------------------------------
SW = "1.7"  # default stroke-width
SW2 = "2"   # bolder variant

def svg(paths_or_elements: str, bg_fill: str = "none", size: int = 24) -> str:
    """Wrap path data in a full SVG element."""
    return f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {size} {size}" fill="{bg_fill}" stroke="#18181B" stroke-width="{SW}" stroke-linecap="round" stroke-linejoin="round">
{paths_or_elements}
</svg>'''

def save(name: str, content: str, subfolder: str = ""):
    dest = os.path.join(SVG_OUT, subfolder) if subfolder else SVG_OUT
    os.makedirs(dest, exist_ok=True)
    with open(os.path.join(dest, f"{name}.svg"), "w", encoding="utf-8") as f:
        f.write(content)
    print(f"  saved: {name}.svg")

# -------------------------------------------------
# Icon definitions – each returns SVG string
# -------------------------------------------------

def icon_ui_showcase():
    # Layered squares with a subtle sparkle – represents design components
    return svg('''  <rect x="3" y="3" width="8" height="8" rx="1.5"/>
  <rect x="13" y="3" width="8" height="8" rx="1.5"/>
  <rect x="3" y="13" width="8" height="8" rx="1.5"/>
  <rect x="13" y="13" width="8" height="8" rx="1.5"/>''')

def icon_auto_work():
    # Play button / lightning bolt for automation
    return svg('''  <path d="M13 2L4.5 13.5H11L11 22L19.5 10.5H13L13 2Z"/>''')

# ---- Annotation panel ----

def icon_dwg_manager():
    # CAD/drawing file with layers
    return svg('''  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <line x1="8" y1="13" x2="16" y2="13"/>
  <line x1="8" y1="17" x2="16" y2="17"/>
  <line x1="8" y1="9" x2="10" y2="9"/>''')

def icon_auto_dimension():
    # Ruler with tick marks
    return svg('''  <line x1="3" y1="12" x2="21" y2="12"/>
  <line x1="3" y1="9" x2="3" y2="15"/>
  <line x1="21" y1="9" x2="21" y2="15"/>
  <line x1="9" y1="10" x2="9" y2="14"/>
  <line x1="15" y1="10" x2="15" y2="14"/>
  <path d="M6.5 7.5L9 5l2.5 2.5" stroke-width="1.5"/>
  <path d="M6.5 16.5L9 19l2.5-2.5" stroke-width="1.5"/>''')

def icon_reset_overrides():
    # Reset/refresh arrow with a paint bucket
    return svg('''  <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/>
  <path d="M3 3v5h5"/>''')

def icon_snap_dimension():
    # Magnet + dimension lines
    return svg('''  <path d="M5 5v6a7 7 0 0 0 14 0V5"/>
  <line x1="5" y1="2" x2="5" y2="5"/>
  <line x1="19" y1="2" x2="19" y2="5"/>
  <line x1="8" y1="17" x2="16" y2="17"/>''')

def icon_smart_align():
    # Alignment guides with elements
    return svg('''  <line x1="12" y1="2" x2="12" y2="22"/>
  <rect x="14" y="5" width="6" height="4" rx="1"/>
  <rect x="4" y="10" width="6" height="4" rx="1"/>
  <rect x="14" y="15" width="6" height="4" rx="1"/>''')

def icon_smart_selection():
    # Cursor with sparkle
    return svg('''  <path d="M4 4l7 18 3.5-7.5L22 11z"/>
  <line x1="14.5" y1="14.5" x2="20" y2="20"/>''')

def icon_material_select():
    # Material swatch / texture
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="2"/>
  <path d="M3 9h18"/>
  <path d="M9 9v12"/>
  <path d="M15 9v12"/>''')

def icon_quick_element():
    # Lightning bolt cursor
    return svg('''  <path d="M4 4l7 18 3.5-7.5L22 11z"/>
  <path d="M13 2L8 13h6l-1 9 7-11h-6z" stroke-width="1.4"/>''')

def icon_select_linked():
    # Chain link with selection box
    return svg('''  <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>
  <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>''')

def icon_deselect_grouped():
    # Group box with an X
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="2" stroke-dasharray="4 2"/>
  <line x1="9" y1="9" x2="15" y2="15"/>
  <line x1="15" y1="9" x2="9" y2="15"/>''')

def icon_select_all_inplace():
    # Elements inside a boundary
    return svg('''  <rect x="2" y="2" width="20" height="20" rx="2" stroke-dasharray="4 2"/>
  <circle cx="8" cy="12" r="2" fill="#18181B" stroke="none"/>
  <circle cx="16" cy="12" r="2" fill="#18181B" stroke="none"/>
  <circle cx="12" cy="7" r="2" fill="#18181B" stroke="none"/>''')

def icon_select_by_category():
    # Filter funnel with category grid
    return svg('''  <polygon points="3 4 10 12 10 20 14 18 14 12 21 4"/>''')

def icon_select_on_sheets():
    # Sheet with selection cursor
    return svg('''  <rect x="3" y="2" width="14" height="18" rx="1.5"/>
  <path d="M17 6h4v14a2 2 0 0 1-2 2H7"/>
  <path d="M14 14l4 4"/>
  <path d="M11 11l4 4-2.5.5-.5-2.5z"/>''')

def icon_select_similar():
    # Mirror/copy elements
    return svg('''  <rect x="2" y="8" width="6" height="8" rx="1"/>
  <rect x="16" y="8" width="6" height="8" rx="1"/>
  <line x1="12" y1="4" x2="12" y2="20" stroke-dasharray="3 2"/>
  <path d="M9 10l2 2-2 2"/>
  <path d="M15 10l-2 2 2 2"/>''')

def icon_annotation_manager():
    # Tags with a list
    return svg('''  <path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z"/>
  <line x1="7" y1="7" x2="7.01" y2="7"/>''')

def icon_text_tag_tools():
    # Text with a tag
    return svg('''  <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
  <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>''')

def icon_copy_anno():
    # Copy icon with annotation mark
    return svg('''  <rect x="9" y="9" width="13" height="13" rx="2"/>
  <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
  <line x1="15" y1="13" x2="18" y2="13"/>
  <line x1="15" y1="16" x2="18" y2="16"/>''')

def icon_dim_text():
    # Dimension with text
    return svg('''  <line x1="3" y1="19" x2="21" y2="19"/>
  <line x1="3" y1="16" x2="3" y2="22"/>
  <line x1="21" y1="16" x2="21" y2="22"/>
  <path d="M8 12h3m2 0h3"/>
  <path d="M10 9l2 3-2 3"/>''')

def icon_renumber_spline():
    # Numbers along a curve
    return svg('''  <path d="M4 17c3-3 6-9 10-9s6 6 6 9" fill="none"/>
  <circle cx="5" cy="15" r="1.5" fill="#18181B" stroke="none"/>
  <circle cx="12" cy="9" r="1.5" fill="#18181B" stroke="none"/>
  <circle cx="19" cy="17" r="1.5" fill="#18181B" stroke="none"/>
  <text x="3" y="8" font-size="5" fill="#18181B" stroke="none" font-family="sans-serif" font-weight="600">1</text>
  <text x="10.5" y="7" font-size="5" fill="#18181B" stroke="none" font-family="sans-serif" font-weight="600">2</text>
  <text x="18" y="14.5" font-size="5" fill="#18181B" stroke="none" font-family="sans-serif" font-weight="600">3</text>''')

def icon_tag_checker():
    # Tag with checkmark
    return svg('''  <path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z"/>
  <line x1="7" y1="7" x2="7.01" y2="7"/>
  <path d="M9 13l2 2 4-4"/>''')

def icon_upper_all():
    # Uppercase A letters
    return svg('''  <path d="M4 18L8.5 6l4.5 12"/>
  <line x1="5.5" y1="14" x2="12" y2="14"/>
  <path d="M15 10h6"/>
  <path d="M15 6h6"/>
  <path d="M15 14h4"/>''')

# ---- Data panel ----

def icon_bcf_reader():
    # Chat bubble with file/issue
    return svg('''  <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
  <line x1="9" y1="10" x2="15" y2="10"/>
  <line x1="9" y1="13" x2="13" y2="13"/>''')

def icon_contains_manager():
    # Nested containers
    return svg('''  <rect x="2" y="2" width="20" height="20" rx="3"/>
  <rect x="6" y="6" width="12" height="12" rx="2"/>
  <rect x="10" y="10" width="4" height="4" rx="1" fill="#18181B" stroke="none"/>''')

def icon_ifc_sg():
    # IFC cube with connection nodes
    return svg('''  <path d="M12 2l8.5 4.9v9.9L12 22 3.5 16.8V6.9z"/>
  <line x1="12" y1="2" x2="12" y2="22"/>
  <line x1="3.5" y1="6.9" x2="20.5" y2="6.9"/>''')

def icon_subtype_definer():
    # Tree / hierarchy with sub-nodes
    return svg('''  <circle cx="12" cy="5" r="2"/>
  <line x1="12" y1="7" x2="12" y2="11"/>
  <line x1="12" y1="11" x2="7" y2="14"/>
  <line x1="12" y1="11" x2="17" y2="14"/>
  <circle cx="7" cy="16" r="2"/>
  <circle cx="17" cy="16" r="2"/>''')

def icon_ifc_checker():
    # Cube with checkmark overlay
    return svg('''  <path d="M12 2l8.5 4.9v9.9L12 22 3.5 16.8V6.9z"/>
  <path d="M8 12l3 3 5-5"/>''')

def icon_ifc_assign():
    # Cube with assignment arrow
    return svg('''  <path d="M12 2l8.5 4.9v9.9L12 22 3.5 16.8V6.9z"/>
  <path d="M9 12h6"/>
  <path d="M13 10l2 2-2 2"/>''')

def icon_param_loader():
    # Parameter upload arrow
    return svg('''  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
  <polyline points="17 8 12 3 7 8"/>
  <line x1="12" y1="3" x2="12" y2="15"/>
  <line x1="8" y1="19" x2="16" y2="19"/>''')

def icon_parameter_manager():
    # Grid/table with parameter rows
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="2"/>
  <line x1="3" y1="9" x2="21" y2="9"/>
  <line x1="9" y1="9" x2="9" y2="21"/>
  <line x1="15" y1="9" x2="15" y2="21"/>''')

def icon_room_foundation():
    # Floor plan footprint
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="1"/>
  <path d="M3 9h18"/>
  <path d="M3 15h18"/>
  <path d="M9 3v18"/>
  <path d="M15 3v18"/>''')

def icon_foundation_volume():
    # 3D box/foundation
    return svg('''  <path d="M2 17L12 22l10-5V7L12 2 2 7v10z"/>
  <path d="M12 22V12"/>
  <path d="M2 7l10 5 10-5"/>''')

def icon_room_data():
    # Room outline with data dots
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="1"/>
  <path d="M3 9h18"/>
  <circle cx="9" cy="15" r="1.5" fill="#18181B" stroke="none"/>
  <circle cx="15" cy="15" r="1.5" fill="#18181B" stroke="none"/>''')

def icon_schedule_link():
    # Table with link icon
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="2"/>
  <line x1="3" y1="9" x2="21" y2="9"/>
  <line x1="3" y1="15" x2="21" y2="15"/>
  <path d="M10 12.5a2 2 0 1 0 0-1H14a2 2 0 1 0 0 1H10z"/>''')

def icon_schedule_manager():
    # Spreadsheet / table with settings
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="2"/>
  <line x1="3" y1="9" x2="21" y2="9"/>
  <line x1="3" y1="15" x2="21" y2="15"/>
  <line x1="9" y1="9" x2="9" y2="21"/>''')

def icon_schedule_copy():
    # Copy table
    return svg('''  <rect x="9" y="9" width="12" height="12" rx="2"/>
  <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
  <line x1="12" y1="12" x2="18" y2="12"/>
  <line x1="12" y1="15" x2="18" y2="15"/>
  <line x1="12" y1="18" x2="15" y2="18"/>''')

def icon_text_to_element():
    # Text cursor with arrow to box
    return svg('''  <path d="M4 7V4h16v3"/>
  <path d="M9 20h6"/>
  <line x1="12" y1="4" x2="12" y2="20"/>
  <path d="M16 14l4 4-4 4"/>
  <rect x="2" y="13" width="10" height="8" rx="1"/>''')

def icon_transfer_para():
    # Two elements with arrow between
    return svg('''  <rect x="2" y="4" width="7" height="16" rx="1"/>
  <rect x="15" y="4" width="7" height="16" rx="1"/>
  <path d="M9 9h6m0 0l-2-2m2 2l-2 2"/>
  <path d="M15 15H9m0 0l2-2m-2 2l2 2"/>''')

def icon_values_filled_region():
    # Filled region with value label
    return svg('''  <path d="M3 19L9 5l5 9 3-4 4 9H3z" fill="#18181B" stroke="none" opacity="0.15"/>
  <path d="M3 19L9 5l5 9 3-4 4 9H3z"/>
  <line x1="7" y1="13" x2="13" y2="13"/>''')

# ---- Project panel ----

def icon_cad_to_family():
    # CAD lines to component box
    return svg('''  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <path d="M8 13h8"/>
  <path d="M8 17h5"/>
  <path d="M17 13l2 2-2 2"/>''')

def icon_cad_to_bim():
    # CAD document converting to 3D box
    return svg('''  <path d="M9 3H5a2 2 0 0 0-2 2v4"/>
  <path d="M7 21h10a2 2 0 0 0 2-2v-4"/>
  <path d="M13 3h4a2 2 0 0 1 2 2v4"/>
  <path d="M11 21H7a2 2 0 0 1-2-2v-4"/>
  <path d="M12 8l5 3-5 3-5-3 5-3z"/>
  <line x1="12" y1="11" x2="12" y2="17"/>
  <line x1="7" y1="11" x2="7" y2="14"/>
  <line x1="17" y1="11" x2="17" y2="14"/>''')

def icon_cad_to_elements():
    # Lines to geometric elements
    return svg('''  <path d="M3 6l3-3 3 3"/>
  <line x1="6" y1="3" x2="6" y2="15"/>
  <path d="M14 8l-2-2 2-2"/>
  <path d="M21 8V4H17"/>
  <rect x="13" y="13" width="8" height="8" rx="1"/>''')

def icon_door_threshold():
    # Door with threshold line
    return svg('''  <path d="M5 2h14v20H5z"/>
  <path d="M14 12a2 2 0 1 0-4 0"/>
  <line x1="2" y1="22" x2="22" y2="22"/>
  <line x1="9" y1="2" x2="9" y2="22"/>''')

def icon_image_to_drafting():
    # Image frame to drawing
    return svg('''  <rect x="3" y="3" width="18" height="14" rx="2"/>
  <circle cx="8.5" cy="8.5" r="1.5"/>
  <polyline points="21 15 16 10 5 21"/>
  <line x1="8" y1="20" x2="16" y2="20"/>
  <path d="M8 20l2-2h4l2 2"/>''')

def icon_point_cloud():
    # Scattered dots with a box
    return svg('''  <circle cx="6" cy="6" r="1" fill="#18181B" stroke="none"/>
  <circle cx="12" cy="4" r="1" fill="#18181B" stroke="none"/>
  <circle cx="18" cy="7" r="1" fill="#18181B" stroke="none"/>
  <circle cx="5" cy="13" r="1" fill="#18181B" stroke="none"/>
  <circle cx="14" cy="11" r="1" fill="#18181B" stroke="none"/>
  <circle cx="19" cy="15" r="1" fill="#18181B" stroke="none"/>
  <circle cx="8" cy="18" r="1" fill="#18181B" stroke="none"/>
  <rect x="4" y="4" width="16" height="16" rx="2" fill="none" stroke-dasharray="3 2"/>''')

def icon_room_to_floor():
    # Room outline with floor slab
    return svg('''  <rect x="3" y="3" width="18" height="12" rx="1"/>
  <rect x="2" y="17" width="20" height="5" rx="1" fill="#18181B" stroke="none" opacity="0.2"/>
  <rect x="2" y="17" width="20" height="5" rx="1"/>
  <path d="M9 3v12"/>
  <path d="M15 3v12"/>''')

def icon_datum_tools():
    # Gridlines (level + column grid)
    return svg('''  <line x1="3" y1="7" x2="21" y2="7"/>
  <line x1="3" y1="17" x2="21" y2="17"/>
  <line x1="7" y1="3" x2="7" y2="21"/>
  <line x1="17" y1="3" x2="17" y2="21"/>
  <circle cx="7" cy="7" r="2" fill="#18181B" stroke="none"/>
  <circle cx="17" cy="17" r="2" fill="#18181B" stroke="none"/>''')

def icon_datum_manager():
    # Datum lines with list manager
    return svg('''  <line x1="2" y1="5" x2="22" y2="5"/>
  <line x1="2" y1="12" x2="22" y2="12"/>
  <line x1="2" y1="19" x2="22" y2="19"/>
  <circle cx="5" cy="5" r="2" fill="#18181B" stroke="none"/>
  <circle cx="5" cy="12" r="2" fill="#18181B" stroke="none"/>
  <circle cx="5" cy="19" r="2" fill="#18181B" stroke="none"/>''')

def icon_property_line():
    # Survey boundary polygon
    return svg('''  <polygon points="12 3 20 8 18 18 6 20 3 10"/>
  <circle cx="12" cy="3" r="1.5" fill="#18181B" stroke="none"/>
  <circle cx="20" cy="8" r="1.5" fill="#18181B" stroke="none"/>
  <circle cx="18" cy="18" r="1.5" fill="#18181B" stroke="none"/>
  <circle cx="6" cy="20" r="1.5" fill="#18181B" stroke="none"/>
  <circle cx="3" cy="10" r="1.5" fill="#18181B" stroke="none"/>''')

def icon_element_adjust():
    # Element with adjustment arrows
    return svg('''  <rect x="5" y="8" width="14" height="8" rx="1.5"/>
  <path d="M3 12H1m22 0h-2"/>
  <path d="M12 6V2m0 20v-4"/>''')

def icon_split_elements():
    # Element being split by a line
    return svg('''  <path d="M3 12h18"/>
  <path d="M3 6h8a2 2 0 0 1 2 2v1"/>
  <path d="M3 18h8a2 2 0 0 0 2-2v-1"/>
  <path d="M21 6h-4a2 2 0 0 0-2 2v1"/>
  <path d="M21 18h-4a2 2 0 0 1-2-2v-1"/>''')

def icon_wall_cut_profile():
    # Wall with profile cut
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="1"/>
  <path d="M3 10h18"/>
  <path d="M3 10l4-4"/>
  <path d="M10 10l4-4"/>
  <path d="M17 10l3-3"/>''')

def icon_wall_adjust_base():
    # Wall with adjustable base arrow
    return svg('''  <rect x="6" y="4" width="12" height="14" rx="1"/>
  <path d="M4 20h16"/>
  <path d="M9 20v-2"/>
  <path d="M15 20v-2"/>
  <path d="M12 18V20"/>''')

def icon_family_batch_creator():
    # Multiple family boxes with + icon
    return svg('''  <rect x="2" y="12" width="8" height="8" rx="1"/>
  <rect x="9" y="2" width="13" height="9" rx="1"/>
  <path d="M14 5v6"/>
  <path d="M11 8h6"/>
  <line x1="2" y1="5" x2="7" y2="5"/>
  <path d="M2 2l5 6H2z" fill="#18181B" stroke="none" opacity="0.15"/>''')

def icon_family_manager():
    # Family tree / component library
    return svg('''  <rect x="3" y="3" width="7" height="7" rx="1"/>
  <rect x="14" y="3" width="7" height="7" rx="1"/>
  <rect x="8" y="14" width="8" height="7" rx="1"/>
  <line x1="6.5" y1="10" x2="6.5" y2="14"/>
  <line x1="17.5" y1="10" x2="17.5" y2="14"/>
  <line x1="6.5" y1="14" x2="17.5" y2="14"/>
  <line x1="12" y1="14" x2="12" y2="17"/>''')

def icon_json_to_family():
    # JSON braces to component
    return svg('''  <path d="M8 3H7a2 2 0 0 0-2 2v5a2 2 0 0 1-2 2 2 2 0 0 1 2 2v5c0 1.1.9 2 2 2h1"/>
  <path d="M16 3h1a2 2 0 0 1 2 2v5a2 2 0 0 0 2 2 2 2 0 0 0-2 2v5a2 2 0 0 1-2 2h-1"/>''')

def icon_tile_layout():
    # Grid tile layout
    return svg('''  <rect x="2" y="2" width="9" height="9" rx="1"/>
  <rect x="13" y="2" width="9" height="9" rx="1"/>
  <rect x="2" y="13" width="9" height="9" rx="1"/>
  <rect x="13" y="13" width="9" height="9" rx="1"/>''')

def icon_auto_join():
    # Two elements joining/merging
    return svg('''  <path d="M8 6H4a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h4"/>
  <path d="M16 6h4a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2h-4"/>
  <path d="M8 12h8"/>
  <path d="M13 9l3 3-3 3"/>''')

def icon_workset_manager():
    # Workset layers
    return svg('''  <ellipse cx="12" cy="6" rx="9" ry="3"/>
  <path d="M3 6v4c0 1.66 4.03 3 9 3s9-1.34 9-3V6"/>
  <path d="M3 10v4c0 1.66 4.03 3 9 3s9-1.34 9-3v-4"/>
  <path d="M3 14v4c0 1.66 4.03 3 9 3s9-1.34 9-3v-4"/>''')

def icon_central_file():
    # Central file / sync hub
    return svg('''  <circle cx="12" cy="12" r="3"/>
  <line x1="12" y1="3" x2="12" y2="9"/>
  <line x1="12" y1="15" x2="12" y2="21"/>
  <line x1="3" y1="12" x2="9" y2="12"/>
  <line x1="15" y1="12" x2="21" y2="12"/>
  <line x1="5.64" y1="5.64" x2="7.76" y2="7.76"/>
  <line x1="16.24" y1="16.24" x2="18.36" y2="18.36"/>
  <line x1="5.64" y1="18.36" x2="7.76" y2="16.24"/>
  <line x1="16.24" y1="7.76" x2="18.36" y2="5.64"/>''')

# ---- Settings panel ----

def icon_advanced_purge():
    # Trash with advanced filter
    return svg('''  <polyline points="3 6 5 6 21 6"/>
  <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
  <path d="M10 11v6"/>
  <path d="M14 11v6"/>
  <path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
  <path d="M17 10l-5 5"/>
  <path d="M12 10l5 5"/>''')

def icon_cleanup_manager():
    # Broom with sparkle
    return svg('''  <path d="M3 21l9-9"/>
  <path d="M12.22 6.22L16 2.5a1.1 1.1 0 0 1 1.56 0l3.94 3.94a1.1 1.1 0 0 1 0 1.56L17.78 11.78"/>
  <path d="M8 16l-5 5"/>
  <path d="M16 21l-3-3 1.5-1.5 3 3z"/>''')

def icon_health_check():
    # Heart with check/pulse line
    return svg('''  <path d="M22 12h-4l-3 9L9 3l-3 9H2"/>''')

def icon_inplace_model():
    # In-place element with edit pencil
    return svg('''  <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
  <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>''')

def icon_location_manager():
    # Pin with map grid
    return svg('''  <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/>
  <circle cx="12" cy="10" r="3"/>''')

def icon_material_list():
    # Material sample swatches
    return svg('''  <rect x="3" y="3" width="5" height="5" rx="1"/>
  <rect x="10" y="3" width="5" height="5" rx="1"/>
  <rect x="17" y="3" width="4" height="5" rx="1"/>
  <line x1="3" y1="12" x2="21" y2="12"/>
  <line x1="3" y1="16" x2="15" y2="16"/>
  <line x1="3" y1="20" x2="9" y2="20"/>''')

def icon_model_auditor():
    # Magnifying glass over a model
    return svg('''  <circle cx="10" cy="10" r="7"/>
  <line x1="15.5" y1="15.5" x2="21" y2="21"/>
  <path d="M7 10h6"/>
  <path d="M10 7v6"/>''')

def icon_model_checker():
    # Model with check shield
    return svg('''  <path d="M12 2L3 7v6c0 5.25 3.75 10.15 9 11.25C17.25 23.15 21 18.25 21 13V7l-9-5z"/>
  <path d="M9 12l2 2 4-4"/>''')

def icon_color_splasher():
    # Paint palette
    return svg('''  <path d="M12 2a10 10 0 0 1 10 10c0 3-2 5-5 5a2 2 0 0 0-2 2 2 2 0 0 1-2 2H9a7 7 0 0 1-7-7c0-6.63 5.37-12 10-12z"/>
  <circle cx="8.5" cy="9.5" r="1" fill="#18181B" stroke="none"/>
  <circle cx="12.5" cy="7.5" r="1" fill="#18181B" stroke="none"/>
  <circle cx="16.5" cy="10.5" r="1" fill="#18181B" stroke="none"/>''')

def icon_style_manager():
    # Line styles preview
    return svg('''  <line x1="3" y1="6" x2="21" y2="6"/>
  <line x1="3" y1="12" x2="21" y2="12" stroke-dasharray="4 2"/>
  <line x1="3" y1="18" x2="21" y2="18" stroke-dasharray="1 2"/>
  <path d="M19 4v4"/>
  <path d="M19 10v4"/>
  <path d="M19 16v4"/>''')

def icon_smart_delete():
    # Trash with lightning bolt
    return svg('''  <polyline points="3 6 5 6 21 6"/>
  <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
  <path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
  <path d="M13 3L11 10h4l-2 7"/>''')

def icon_smart_purge():
    # Funnel with sparkle
    return svg('''  <polygon points="3 4 10 12 10 20 14 18 14 12 21 4"/>
  <path d="M15 5l1.5 1.5"/>
  <path d="M17 8l1.5 1.5"/>
  <path d="M14 8l1.5-1.5"/>''')

def icon_warning_manage():
    # Warning triangle with list
    return svg('''  <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
  <line x1="12" y1="9" x2="12" y2="13"/>
  <line x1="12" y1="17" x2="12.01" y2="17"/>''')

# ---- Support panel ----

def icon_t3lab_assistant():
    # AI / brain with chat bubble
    return svg('''  <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
  <path d="M9 10h.01"/>
  <path d="M15 10h.01"/>
  <path d="M9.5 14.5c.83.5 2.17.5 3 0"/>''')

def icon_mcp_control():
    # Connected nodes / API
    return svg('''  <circle cx="12" cy="12" r="3"/>
  <circle cx="4" cy="6" r="2"/>
  <circle cx="20" cy="6" r="2"/>
  <circle cx="4" cy="18" r="2"/>
  <circle cx="20" cy="18" r="2"/>
  <line x1="6" y1="6" x2="9" y2="10"/>
  <line x1="18" y1="6" x2="15" y2="10"/>
  <line x1="6" y1="18" x2="9" y2="14"/>
  <line x1="18" y1="18" x2="15" y2="14"/>''')

def icon_feedback():
    # Speech bubble with heart
    return svg('''  <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
  <path d="M12 8.5c0-1.1.9-2 2-2s2 .9 2 2-2 3.5-2 3.5S12 9.6 12 8.5z"/>''')

def icon_cloud_links():
    # Cloud with external link arrow
    return svg('''  <path d="M18 10h-1.26A8 8 0 1 0 9 20h9a5 5 0 0 0 0-10z"/>
  <path d="M14 14l4 4-4 4" transform="translate(-2 -4) scale(0.7)"/>''')

def icon_acc_platform():
    # Cloud + building
    return svg('''  <path d="M17 8h1a4 4 0 0 1 0 8H6a5 5 0 0 1-.81-9.94A7 7 0 0 1 17 8z"/>
  <line x1="10" y1="21" x2="10" y2="16"/>
  <line x1="14" y1="21" x2="14" y2="16"/>
  <line x1="8" y1="21" x2="16" y2="21"/>''')

def icon_b360_health():
    # Cloud with heart pulse
    return svg('''  <path d="M17 8h1a4 4 0 0 1 0 8H6a5 5 0 0 1 0-10h.39"/>
  <path d="M6 15h2l2-4 2 6 2-3 1.5 1.5"/>''')

def icon_bluebeam_health():
    # Document health check
    return svg('''  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <path d="M8 13h2l2-4 2 6 1-2h1"/>''')

def icon_bg_theme():
    # Color/theme half circle
    return svg('''  <circle cx="12" cy="12" r="9"/>
  <path d="M12 3v18" fill="none"/>
  <path d="M12 3a9 9 0 0 1 0 18" fill="#18181B" stroke="none" opacity="0.15"/>
  <path d="M12 3a9 9 0 0 1 0 18"/>''')

def icon_ribbon_names():
    # Ribbon/menu bar
    return svg('''  <rect x="2" y="4" width="20" height="4" rx="1"/>
  <line x1="5" y1="11" x2="5" y2="21"/>
  <line x1="19" y1="11" x2="19" y2="21"/>
  <line x1="12" y1="11" x2="12" y2="17"/>
  <line x1="3" y1="14" x2="21" y2="14"/>''')

def icon_tab_manager():
    # Tabs / panels
    return svg('''  <rect x="2" y="7" width="20" height="15" rx="1"/>
  <path d="M2 7V5a1 1 0 0 1 1-1h5a1 1 0 0 1 1 1v2"/>
  <path d="M10 7V5a1 1 0 0 1 1-1h5a1 1 0 0 1 1 1v2"/>''')

# ---- ViewsSheets panel ----

def icon_batch_out():
    # Export to multiple formats
    return svg('''  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
  <polyline points="7 10 12 15 17 10"/>
  <line x1="12" y1="15" x2="12" y2="3"/>
  <line x1="3" y1="20" x2="7" y2="20"/>
  <line x1="10" y1="20" x2="14" y2="20"/>
  <line x1="17" y1="20" x2="21" y2="20"/>''')

def icon_create_room_plan():
    # Room + plan view
    return svg('''  <rect x="3" y="3" width="18" height="11" rx="1"/>
  <path d="M9" y1="3" x2="9" y2="14"/>
  <rect x="5" y="17" width="14" height="4" rx="1"/>
  <line x1="9" y1="14" x2="9" y2="17"/>
  <path d="M9 3v11"/>
  <path d="M15 3v11"/>''')

def icon_sheet_hub():
    # Sheet stack with hub
    return svg('''  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <circle cx="12" cy="14" r="3"/>
  <line x1="12" y1="11" x2="12" y2="9"/>
  <line x1="12" y1="17" x2="12" y2="19"/>''')

def icon_linked_element_box():
    # Linked boxes
    return svg('''  <rect x="2" y="2" width="9" height="9" rx="1"/>
  <rect x="13" y="13" width="9" height="9" rx="1"/>
  <path d="M10 6h4a2 2 0 0 1 2 2v2"/>
  <path d="M8 18H6a2 2 0 0 1-2-2v-2"/>''')

def icon_sheet_renumber():
    # Sheet with numbered list
    return svg('''  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <path d="M9 10h2v4"/>
  <path d="M8 14h4"/>
  <path d="M13 10h3l-2 2 2 2h-3"/>''')

def icon_sheet_manager():
    # Sheet stack with settings
    return svg('''  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <line x1="8" y1="13" x2="16" y2="13"/>
  <line x1="8" y1="17" x2="14" y2="17"/>''')

def icon_view_hub():
    # Camera/view with hub
    return svg('''  <path d="M23 7l-7 5 7 5V7z"/>
  <rect x="1" y="5" width="15" height="14" rx="2"/>
  <circle cx="8" cy="12" r="3"/>''')

def icon_view_manager():
    # View list with eye
    return svg('''  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
  <circle cx="12" cy="12" r="3"/>
  <line x1="3" y1="20" x2="21" y2="20"/>
  <line x1="5" y1="17" x2="19" y2="17"/>''')

def icon_view_template():
    # Template/stamp on view
    return svg('''  <rect x="3" y="3" width="18" height="18" rx="2"/>
  <path d="M3 9h18"/>
  <path d="M9 21V9"/>
  <path d="M12 14l2 2 4-4"/>''')

# -------------------------------------------------
# Map of icon_name -> (function, relative_path_in_tab)
# -------------------------------------------------
ICONS = {
    # Standard panel
    "Standard/UIStandard": (icon_ui_showcase, "Standard.panel/UIStandard.pushbutton"),
    "Standard/AutoWork": (icon_auto_work, "Standard.panel/AutoWork.pushbutton"),

    # Annotation panel
    "Annotation/DWGManager": (icon_dwg_manager, "Annotation.panel/DWG Management.pushbutton"),
    "Annotation/AutoDimension": (icon_auto_dimension, "Annotation.panel/Graphics.stack/AutoDimension.pushbutton"),
    "Annotation/ResetOverrides": (icon_reset_overrides, "Annotation.panel/Graphics.stack/ResetOverrides.pushbutton"),
    "Annotation/SnapDimension": (icon_snap_dimension, "Annotation.panel/Graphics.stack/SnapDimension.pushbutton"),
    "Annotation/SmartAlign": (icon_smart_align, "Annotation.panel/SmartAlign.pushbutton"),
    "Annotation/SmartSelection": (icon_smart_selection, "Annotation.panel/SmartSelection.pulldown"),
    "Annotation/MaterialSelect": (icon_material_select, "Annotation.panel/SmartSelection.pulldown/Material Select Element.pushbutton"),
    "Annotation/QuickElement": (icon_quick_element, "Annotation.panel/SmartSelection.pulldown/QuickElement.pushbutton"),
    "Annotation/SelectLinked": (icon_select_linked, "Annotation.panel/SmartSelection.pulldown/Select Linked.pushbutton"),
    "Annotation/Select": (icon_smart_selection, "Annotation.panel/SmartSelection.pulldown/Select.pulldown"),
    "Annotation/DeselectGrouped": (icon_deselect_grouped, "Annotation.panel/SmartSelection.pulldown/Select.pulldown/DeselectGroupedElements.pushbutton"),
    "Annotation/SelectAllInPlace": (icon_select_all_inplace, "Annotation.panel/SmartSelection.pulldown/Select.pulldown/SelectAllInPlaceElements.pushbutton"),
    "Annotation/SelectByCategory": (icon_select_by_category, "Annotation.panel/SmartSelection.pulldown/Select.pulldown/SelectByCategory.pushbutton"),
    "Annotation/AnnotationManager": (icon_annotation_manager, "Annotation.panel/Text.stack/AnnotationManager.pushbutton"),
    "Annotation/TextTagTools": (icon_text_tag_tools, "Annotation.panel/Text.stack/TextTagTools.pulldown"),
    "Annotation/CopyAnno": (icon_copy_anno, "Annotation.panel/Text.stack/TextTagTools.pulldown/Copy Annotation.pushbutton"),
    "Annotation/DimText": (icon_dim_text, "Annotation.panel/Text.stack/TextTagTools.pulldown/DimText.pushbutton"),
    "Annotation/RenumberSpline": (icon_renumber_spline, "Annotation.panel/Text.stack/TextTagTools.pulldown/Renumber by spline.pushbutton"),
    "Annotation/TagChecker": (icon_tag_checker, "Annotation.panel/Text.stack/TextTagTools.pulldown/Tag Checker.pushbutton"),
    "Annotation/UpperAll": (icon_upper_all, "Annotation.panel/Text.stack/TextTagTools.pulldown/UpperAll.pushbutton"),

    # Data panel
    "Data/BCFReader": (icon_bcf_reader, "Data.panel/Data.stack/BCF Reader.pushbutton"),
    "Data/ContainsManager": (icon_contains_manager, "Data.panel/Data.stack/ContainsManager.pushbutton"),
    "Data/IFCSG": (icon_ifc_sg, "Data.panel/IFC-SG.pulldown"),
    "Data/SubtypeDefiner": (icon_subtype_definer, "Data.panel/IFC-SG.pulldown/IFCSG Subtype Definer.pushbutton"),
    "Data/IFCSGChecker": (icon_ifc_checker, "Data.panel/IFC-SG.pulldown/IFCSGChecker.pushbutton"),
    "Data/IFCSGAssign": (icon_ifc_assign, "Data.panel/IFC-SG.pulldown/IFCSG_Assign.pushbutton"),
    "Data/ParamLoader": (icon_param_loader, "Data.panel/IFC-SG.pulldown/ParameterLoader.pushbutton"),
    "Data/ParameterManager": (icon_parameter_manager, "Data.panel/ParameterManager.pushbutton"),
    "Data/RoomFoundation": (icon_room_foundation, "Data.panel/Room & Foundation.pulldown"),
    "Data/FoundationVolume": (icon_foundation_volume, "Data.panel/Room & Foundation.pulldown/Foundation Volume.pushbutton"),
    "Data/RoomData": (icon_room_data, "Data.panel/Room & Foundation.pulldown/RoomDataCollector.pushbutton"),
    "Data/ScheduleLink": (icon_schedule_link, "Data.panel/ScheduleExportImportPro"),
    "Data/ScheduleManager": (icon_schedule_manager, "Data.panel/ScheduleManager.pushbutton"),
    "Data/ScheduleCopy": (icon_schedule_copy, "Data.panel/Schedule_Copy"),
    "Data/TextToElement": (icon_text_to_element, "Data.panel/Text to element"),
    "Data/TransferPara": (icon_transfer_para, "Data.panel/Transfer Para"),
    "Data/ValuesFilledRegion": (icon_values_filled_region, "Data.panel/Values to Filled Region"),

    # Project panel
    "Project/CADToFamily": (icon_cad_to_family, "Project.panel/CAD To Family"),
    "Project/CADToBIM": (icon_cad_to_bim, "Project.panel/Create.stack/Create Elements.pulldown"),
    "Project/CADToElements": (icon_cad_to_elements, "Project.panel/Create.stack/Create Elements.pulldown/CADToElements.pushbutton"),
    "Project/ImageToDrafting": (icon_image_to_drafting, "Project.panel/Create.stack/Create Elements.pulldown/ImageToDrafting.pushbutton"),
    "Project/PointCloud": (icon_point_cloud, "Project.panel/Create.stack/Create Elements.pulldown/PointCloud.pushbutton"),
    "Project/DatumTools": (icon_datum_tools, "Project.panel/Create.stack/Datum.pulldown"),
    "Project/DatumManager": (icon_datum_manager, "Project.panel/Create.stack/DatumManager.pushbutton"),
    "Project/PropertyLine": (icon_property_line, "Project.panel/Create.stack/PropertyLine.pushbutton"),
    "Project/ElementAdjust": (icon_element_adjust, "Project.panel/ElementAdjust.pulldown"),
    "Project/SplitElements": (icon_split_elements, "Project.panel/ElementAdjust.pulldown/SplitElements.pushbutton"),
    "Project/WallCutProfile": (icon_wall_cut_profile, "Project.panel/ElementAdjust.pulldown/Wall Cut Profile.pushbutton"),
    "Project/WallAdjustBase": (icon_wall_adjust_base, "Project.panel/ElementAdjust.pulldown/Wall_Adjust Base.pushbutton"),
    "Project/FamilyBatchCreator": (icon_family_batch_creator, "Project.panel/FamilyBatchCreator.pushbutton"),
    "Project/FamilyManager": (icon_family_manager, "Project.panel/FamilyManager.pushbutton"),
    "Project/JSONToFamily": (icon_json_to_family, "Project.panel/JSON To Family"),
    "Project/TileLayout": (icon_tile_layout, "Project.panel/Tile Layout.pushbutton"),
    "Project/AutoJoin": (icon_auto_join, "Project.panel/Work.stack/AutoJoin.pushbutton"),
    "Project/WorksetManager": (icon_workset_manager, "Project.panel/Work.stack/WorksetManager.pushbutton"),
    "Project/CentralFile": (icon_central_file, "Project.panel/Workset.stack/CentralFile.pushbutton"),

    # Settings panel
    "Settings/AdvancedPurge": (icon_advanced_purge, "Settings.panel/AdvancedPurge"),
    "Settings/CleanupManager": (icon_cleanup_manager, "Settings.panel/CleanupManager.pushbutton"),
    "Settings/HealthCheck": (icon_health_check, "Settings.panel/HealthCheck"),
    "Settings/InPlaceModel": (icon_inplace_model, "Settings.panel/InPlaceModel"),
    "Settings/LocationManager": (icon_location_manager, "Settings.panel/LocationManager.pushbutton"),
    "Settings/MaterialList": (icon_material_list, "Settings.panel/MaterialList"),
    "Settings/ModelAuditor": (icon_model_auditor, "Settings.panel/ModelAuditor.pushbutton"),
    "Settings/ModelChecker": (icon_model_checker, "Settings.panel/ModelChecker"),
    "Settings/ColorSplasher": (icon_color_splasher, "Settings.panel/Settings.stack/ColorSplasher.pushbutton"),
    "Settings/StyleManager": (icon_style_manager, "Settings.panel/Settings.stack/StyleManager.pushbutton"),
    "Settings/SmartDelete": (icon_smart_delete, "Settings.panel/SmartDelete"),
    "Settings/SmartPurge": (icon_smart_purge, "Settings.panel/SmartPurge"),
    "Settings/WarningManage": (icon_warning_manage, "Settings.panel/Warning"),

    # Support panel
    "Support/T3LabAssistant": (icon_t3lab_assistant, "Support.panel/T3LabAssistant.pushbutton"),
    "Support/MCPControl": (icon_mcp_control, "Support.panel/MCPControl.pushbutton"),
    "Support/Feedback": (icon_feedback, "Support.panel/Feedback.pushbutton"),
    "Support/CloudLinks": (icon_cloud_links, "Support.panel/CloudLinks.stack"),
    "Support/AutodeskForma": (icon_acc_platform, "Support.panel/CloudLinks.stack/Autodesk Forma.urlbutton"),
    "Support/AutodeskHealth": (icon_b360_health, "Support.panel/CloudLinks.stack/Autodesk Health.urlbutton"),
    "Support/BluebeamStatus": (icon_bluebeam_health, "Support.panel/CloudLinks.stack/Bluebeam Status.urlbutton"),
    "Support/BGTheme": (icon_bg_theme, "Support.panel/UI.stack/BG Theme.pushbutton"),
    "Support/RibbonNames": (icon_ribbon_names, "Support.panel/UI.stack/Ribbon Names.pushbutton"),
    "Support/TabManager": (icon_tab_manager, "Support.panel/UI.stack/Tab Manager.pushbutton"),

    # ViewsSheets panel
    "ViewsSheets/BatchOut": (icon_batch_out, "ViewsSheets.panel/BatchOut.pushbutton"),
    "ViewsSheets/CreateRoomPlan": (icon_create_room_plan, "ViewsSheets.panel/Create Room Plan"),
    "ViewsSheets/SheetHub": (icon_sheet_hub, "ViewsSheets.panel/SheetManager.stack/SheetHub.pushbutton"),
    "ViewsSheets/LinkedElementBox": (icon_linked_element_box, "ViewsSheets.panel/SheetManager.stack/Linked Element Box.pushbutton"),
    "ViewsSheets/SheetRenumber": (icon_sheet_renumber, "ViewsSheets.panel/SheetManager.stack/Sheet re-number"),
    "ViewsSheets/SheetManager": (icon_sheet_manager, "ViewsSheets.panel/SheetManager.stack/SheetManager"),
    "ViewsSheets/ViewHub": (icon_view_hub, "ViewsSheets.panel/ViewHub.pushbutton"),
    "ViewsSheets/ViewManager": (icon_view_manager, "ViewsSheets.panel/ViewManager"),
    "ViewsSheets/ViewTemplate": (icon_view_template, "ViewsSheets.panel/ViewTemplate"),
}


def main():
    print(f"\n=== Generating {len(ICONS)} SVG icons ===\n")

    # 1. Save all SVGs to the output folder (organized by panel)
    for key, (fn, _rel_path) in ICONS.items():
        panel = key.split("/")[0]
        name = key.split("/")[1]
        svg_content = fn()
        save(name, svg_content, subfolder=panel)

    # 2. Copy SVGs directly to the extension folders
    print(f"\n=== Deploying SVGs to extension folders ===\n")
    for key, (fn, rel_path) in ICONS.items():
        dest_dir = os.path.join(BASE, rel_path.replace("/", os.sep))
        if os.path.isdir(dest_dir):
            svg_content = fn()
            svg_path = os.path.join(dest_dir, "icon.svg")
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(svg_content)
            print(f"  deployed: {rel_path}/icon.svg")
        else:
            print(f"  SKIP (dir not found): {rel_path}")

    print(f"\n✓ Done! {len(ICONS)} icons generated.")
    print(f"  SVGs saved to: {SVG_OUT}")
    print(f"  SVGs deployed to: {BASE}")


if __name__ == "__main__":
    main()
