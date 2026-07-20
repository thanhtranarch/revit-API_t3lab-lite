# -*- coding: utf-8 -*-
"""LLMs Setting

Standalone provider / model / API key / connection settings for the
T3Lab LLM router — shared by every AI-powered T3Lab tool.
"""
__title__ = "LLMs\nSetting"
__author__ = "Tran Tien Thanh"

import os
import sys

# Path setup — script.py lives 4 levels below T3Lab.extension/ (inside a .stack bundle)
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from GUI.LLMSettingDialog import show_llm_setting_dialog


def main():
    show_llm_setting_dialog()


if __name__ == '__main__':
    main()
