# -*- coding: utf-8 -*-
"""
Sheet Manager - Configuration
Colors, constants, and settings

Copyright © Dang Quoc Truong (DQT)
"""


class Config(object):
    """Configuration constants"""
    
    # Colors (DQT Brand)
    PRIMARY_COLOR = "#F0CC88"      # Golden
    BACKGROUND_COLOR = "#FEF8E7"   # Light cream
    
    # Window dimensions
    MAIN_WINDOW_WIDTH = 1400
    MAIN_WINDOW_HEIGHT = 800
    
    @staticmethod
    def get_color_argb(hex_color):
        """Convert hex color to ARGB tuple"""
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (255, r, g, b)  # A, R, G, B
