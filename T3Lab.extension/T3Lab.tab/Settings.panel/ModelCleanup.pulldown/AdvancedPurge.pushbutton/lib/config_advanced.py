# -*- coding: utf-8 -*-
"""
Advanced Purge Configuration
Colors, fonts, and styling for the Advanced Purge tool

⚠️ Red/Orange theme to indicate dangerous operations

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"


class Colors:
    """Color definitions - Red/Orange danger theme"""
    
    # Primary colors - RED/ORANGE for danger
    BUTTON_PRIMARY = "#FFFF5722"      # Deep Orange (danger)
    BUTTON_SECONDARY = "#FFE0E0E0"    # Light Gray
    BUTTON_SUCCESS = "#FFF59E0B"      # Orange (warning)
    BUTTON_DANGER = "#FFEF4444"       # Red (critical danger)
    
    # Background colors
    HEADER = "#FFFFCCBC"              # Light Red/Orange
    BACKGROUND = "#FFFFFFFF"          # White
    PANEL = "#FFFAFAFA"               # Very Light Gray
    
    # Border colors
    BORDER = "#FFBDBDBD"              # Medium Gray
    BORDER_LIGHT = "#FFE0E0E0"        # Light Gray
    
    # Status colors
    WARNING = "#FFFF5722"             # Deep Orange
    ERROR = "#FFEF4444"               # Red
    SUCCESS = "#FF10B981"             # Green
    INFO = "#FF2196F3"                # Blue
    
    # Selection colors
    SELECTED = "#FFFFCCBC"            # Light Red (selected item)
    HIGHLIGHT = "#FFFF5722"           # Deep Orange (highlight)
    
    # Group colors
    GROUP_EXPANDED = "#FFFFCCBC"      # Light Red (expanded group)
    GROUP_COLLAPSED = "#FFFFFFFF"     # White (collapsed group)
    
    # Text colors
    TEXT = "#FF212121"                # Dark Gray (main text)
    TEXT_SECONDARY = "#FF757575"      # Medium Gray (secondary text)
    TEXT_DISABLED = "#FFBDBDBD"       # Light Gray (disabled text)
    
    # Other
    White = "#FFFFFFFF"
    Black = "#FF000000"


class Fonts:
    """Font size definitions - Slightly larger for readability"""
    
    TITLE = 18          # Dialog title (e.g., "Advanced Purge")
    HEADER = 14         # Section headers (e.g., "⚠️ ADVANCED VIEWS")
    NORMAL = 13         # Regular text, buttons, checkboxes
    SMALL = 11          # Small text, tooltips
    COPYRIGHT = 10      # Copyright text


class Icons:
    """Unicode icons for UI elements"""
    
    # Warning icons
    WARNING = u"\u26A0"              # ⚠
    DANGER = u"\u2620"               # ☠
    FIRE = u"\U0001F525"             # 🔥
    
    # Status icons
    CHECK = u"\u2713"                # ✓
    CROSS = u"\u2717"                # ✗
    QUESTION = u"\u003F"             # ?
    INFO = u"\u2139"                 # ℹ
    
    # Group icons
    VIEWS = u"\U0001F441"            # 👁
    WORKSET = u"\u2692"              # ⚒
    MODEL = u"\U0001F527"            # 🔧
    COLLAB = u"\U0001F91D"           # 🤝
    DANGEROUS = u"\u2620"            # ☠
    
    # Action icons
    SCAN = u"\U0001F50D"             # 🔍
    PREVIEW = u"\U0001F441"          # 👁
    EXECUTE = u"\u26A0"              # ⚠
    CLOSE = u"\u274C"                # ✖
    
    # Other icons
    FILTER = u"\U0001F50D"           # 🔍
    SETTINGS = u"\u2699"             # ⚙


class Messages:
    """Common message strings"""
    
    # Warnings
    DANGER_WARNING = "⚠️ This operation is DANGEROUS and cannot be undone!"
    DRY_RUN_REMINDER = "Dry Run is enabled - no changes will be made."
    PREVIEW_REQUIRED = "You must preview the items before executing!"
    
    # Confirmations
    CONFIRM_EXECUTE = "Are you sure you want to execute this operation?"
    CONFIRM_DANGEROUS = "This is a DANGEROUS operation! Are you ABSOLUTELY sure?"
    FINAL_WARNING = "FINAL WARNING: This cannot be undone! Continue?"
    
    # Success messages
    SCAN_COMPLETE = "Scan completed successfully!"
    EXECUTE_COMPLETE = "Operation completed successfully!"
    DRY_RUN_COMPLETE = "Dry run completed - no changes were made."
    
    # Error messages
    NO_SELECTION = "Please select at least one category to scan."
    SCAN_FAILED = "Scan failed. Please check the console for details."
    EXECUTE_FAILED = "Execution failed. Please check the console for details."


class Settings:
    """Default settings"""
    
    # Window dimensions
    WINDOW_WIDTH = 900
    WINDOW_HEIGHT = 950
    
    # Scroll settings
    SCROLL_SPEED = 16
    
    # Safety settings
    DRY_RUN_DEFAULT = True           # Always start with Dry Run enabled
    SHOW_WARNINGS_DEFAULT = True     # Always show warnings
    REQUIRE_PREVIEW = True           # Must preview before execute
    
    # Confirmation levels for dangerous operations
    DANGEROUS_CONFIRMATION_LEVELS = 3  # Triple confirmation required