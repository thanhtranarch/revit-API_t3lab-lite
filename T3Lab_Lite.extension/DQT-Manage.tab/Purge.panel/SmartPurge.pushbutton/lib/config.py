# -*- coding: utf-8 -*-
"""
Configuration for Smart Purge
Color and setting definitions

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"


class Colors(object):
    """Color definitions (ARGB format)"""
    
    # Header and backgrounds
    HEADER = "#FFF0CC88"           # Yellow-orange header
    WINDOW_BG = "#FFF5F5F5"        # Light gray window background
    PANEL_BG = "#FFFFFFFF"         # White panel background
    BORDER = "#FFCCCCCC"           # Light border
    BACKGROUND = "#FFFFFFFF"       # White background
    PANEL = "#FFF5F5F5"            # Gray panel
    
    # Group colors
    GROUP_EXPANDED = "#FFFFE0B2"   # Light orange when expanded
    GROUP_COLLAPSED = "#FFF5F5F5"  # Gray when collapsed
    
    # Status colors
    SAFE_COLOR = "#FF4CAF50"       # Green checkmark
    WARNING_COLOR = "#FFFFC107"    # Yellow warning
    ERROR_COLOR = "#FFF44336"      # Red error
    SUCCESS = "#FF4CAF50"          # Green success
    WARNING = "#FFFFC107"          # Yellow warning
    ERROR = "#FFF44336"            # Red error
    
    # Button colors - ORANGE/EARTH TONES
    BUTTON_PRIMARY = "#FFFF9800"   # Orange (Scan, Purge)
    BUTTON_SECONDARY = "#FF9E9E9E" # Gray (Cancel, Close)
    BUTTON_SUCCESS = "#FFFFA726"   # Light orange (Preview)
    BUTTON_DANGER = "#FFEF6C00"    # Dark orange (Delete operations)
    
    # Text colors
    TEXT_PRIMARY = "#FF000000"     # Black
    TEXT_SECONDARY = "#FF666666"   # Gray
    TEXT_DISABLED = "#FF999999"    # Light gray
    
    # Special colors
    White = "#FFFFFFFF"
    Black = "#FF000000"
    Transparent = "#00000000"
    
    # Highlight - orange theme
    HIGHLIGHT = "#FFFF9800"        # Orange highlight
    SELECTED = "#FFFFE0B2"         # Light orange selected


class Settings(object):
    """Default settings"""
    
    # Window
    WINDOW_WIDTH = 800
    WINDOW_HEIGHT = 900
    
    # Dry run default
    DRY_RUN_DEFAULT = True
    DEFAULT_DRY_RUN = True
    
    # Delete system types default
    DELETE_SYSTEM_DEFAULT = False
    DEFAULT_DELETE_SYSTEM = False
    
    # Batch settings
    BATCH_SIZE = 50
    SHOW_PROGRESS = True
    
    # Timeouts
    SCAN_TIMEOUT = 300  # seconds
    PURGE_TIMEOUT = 600  # seconds