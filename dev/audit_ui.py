#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
T3Lab UI consistency audit theo chuẩn Lumina (CPython 3, chạy ngoài Revit).

Đối chiếu mọi XAML trong lib/GUI/Tools/ với .claude/rules/ui-design-standard.md:
  1. Palette cũ (Kinetix navy / Terra teal-green) còn sót.
  2. Font cấm (Manrope, Segoe UI cho body text).
  3. Window root thiếu FontFamily Hanken Grotesk/Inter (Variant A).
  4. WindowChrome thiếu attribute hoặc viết 1 dòng.
  5. Copyright block: 1 bản duy nhất, ký tự © literal, FontSize 11, màu #F59E0B,
     nằm ở FOOTER (status bar) phía BÊN TRÁI — không right-align/center,
     không float overlay đè content. Item-template XAML được miễn trừ.
  6. Dot-notation <Grid.RowDefinition/> (crash EMPTYPROPERTYELEMENT).
  7. Glyph chrome sai (Unicode minus/white square thay vì Segoe MDL2).
  8. Thiếu block shared styles.

File UI-locked (DWGManagement, ExportManager, T3LabAssistant) chỉ báo info,
không tính vi phạm.

Usage:
    python3 dev/audit_ui.py           # báo cáo đầy đủ
    python3 dev/audit_ui.py --quiet   # chỉ vi phạm, exit 1 nếu có
"""
import os
import re
import sys
import glob

# Console Windows mặc định (cp1252) không in được tiếng Việt — ép UTF-8.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(REPO, "T3Lab.extension", "lib", "GUI", "Tools")

QUIET = "--quiet" in sys.argv

UI_LOCKED = {"DWGManagement.xaml", "ExportManager.xaml",
             # Chat surface, không phải tool dialog: dùng đúng màu UI của Revit
             # + theme sáng/tối theo Revit (GUI/RevitTheme.py) thay cho Lumina.
             # Xem docs/assistant-revit-ui.md.
             "T3LabAssistant.xaml"}

# Item-template XAML (root là list-row <Border>/<DataTemplate>, không phải cửa sổ)
# — chuẩn miễn trừ copyright cho các file này.
COPYRIGHT_EXEMPT = {"CadtoFloorLayerItem.xaml"}

OLD_PALETTE = {
    "#083D56": "Kinetix navy",
    "#0F766E": "Terra teal primary",
    "#115E59": "Terra teal hover",
    "#0B4F4A": "Terra teal pressed",
    "#F6F8F8": "Terra surface",
    "#E6EDEC": "Terra surface hover",
    "#D9EBE8": "Terra selected tint",
    "#DDE5E7": "Terra border",
    "#C7D2D4": "Terra input border",
    "#15803D": "Terra success",
    "#166534": "Terra success hover",
    "#AEBFBC": "Terra disabled bg",
}


def audit_file(src, base):
    issues = []
    is_window = re.search(r'<\s*Window[\s>]', src) is not None

    for hexv, name in OLD_PALETTE.items():
        n = len(re.findall(re.escape(hexv), src, re.I))
        if n:
            issues.append("palette cũ %s (%s) x%d" % (hexv, name, n))

    if re.search(r'FontFamily="[^"]*Manrope', src):
        issues.append("FontFamily Manrope (cấm)")
    for m in re.finditer(r'FontFamily="([^"]*Segoe UI[^"]*)"', src):
        issues.append("FontFamily Segoe UI (%s)" % m.group(1))

    if is_window:
        wtag = re.search(r'<Window\b[^>]*>', src, re.S)
        if wtag and not re.search(r'FontFamily="(Hanken Grotesk|Inter)[^"]*"', wtag.group(0)):
            issues.append("Window root thiếu FontFamily Hanken Grotesk/Inter")
        chrome = re.search(r'<WindowChrome\b(.*?)/>', src, re.S)
        if chrome:
            if "\n" not in chrome.group(0):
                issues.append("WindowChrome viết 1 dòng (mất corner radius)")
            missing = [a for a in ("CaptionHeight", "ResizeBorderThickness",
                                   "GlassFrameThickness", "CornerRadius",
                                   "UseAeroCaptionButtons") if a not in chrome.group(1)]
            if missing:
                issues.append("WindowChrome thiếu: %s" % ",".join(missing))
        else:
            issues.append("không có WindowChrome")

    if base not in COPYRIGHT_EXEMPT:
        n_cr = len(re.findall(r'© Copyright by T3Lab', src))
        n_ent = len(re.findall(r'&#169;', src))
        if n_ent:
            issues.append("copyright dùng &#169; (phải dùng © literal)")
        total = n_cr + n_ent
        if total == 0:
            issues.append("thiếu copyright")
        elif total > 1:
            issues.append("copyright trùng lặp x%d" % total)
        else:
            crtag = re.search(r'<TextBlock[^>]*Copyright by T3Lab[^>]*/>', src, re.S)
            if crtag:
                t = crtag.group(0)
                if re.search(r'HorizontalAlignment="(Right|Center)"', t):
                    issues.append("copyright phải nằm BÊN TRÁI footer (đang Right/Center-align)")
                if '#F59E0B' not in t:
                    issues.append("copyright không dùng amber #F59E0B")
                if 'FontSize="11"' not in t:
                    issues.append("copyright không dùng FontSize 11")

    if re.search(r'<Grid\.(Column|Row)Definition\b', src):
        issues.append("DOT-NOTATION Grid.Row/ColumnDefinition — crash EMPTYPROPERTYELEMENT")

    if "&#x2212;" in src or "&#8722;" in src:
        issues.append("minimize dùng Unicode minus (phải &#xE921;)")
    if "&#x25A1;" in src:
        issues.append("maximize dùng white square (phải &#xE922;)")

    if "T3LAB SHARED STYLES" not in src:
        issues.append("thiếu block shared styles")

    return issues, ("A" if is_window else "B")


def main():
    failed = False
    files = sorted(glob.glob(os.path.join(TOOLS, "*.xaml")))
    n_bad = 0
    for xaml in files:
        base = os.path.basename(xaml)
        with open(xaml, encoding="utf-8", errors="replace") as fh:
            src = fh.read()
        issues, variant = audit_file(src, base)
        if not issues:
            continue
        n_bad += 1
        locked = base in UI_LOCKED
        if locked:
            if not QUIET:
                print("%s (variant %s) [UI-LOCKED — chỉ tham khảo, KHÔNG sửa]" % (base, variant))
                for i in issues:
                    print("   . " + i)
            continue
        failed = True
        print("%s (variant %s)" % (base, variant))
        for i in issues:
            print("   - " + i)

    print()
    print("UI AUDIT: %d/%d file có vấn đề%s" % (n_bad, len(files),
          " — FAILED" if failed else " (chỉ file UI-locked) — OK"))
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
