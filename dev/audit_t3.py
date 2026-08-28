#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
T3Lab UI audit theo chuẩn T3 (CPython 3, chạy ngoài Revit).

Chuẩn duy nhất: `pyRevit UI Design System/T3LAB_UI_STANDARD.md`
              + `pyRevit UI Design System/T3Lab.Styles.xaml`
Luật migration: `docs/ui-governance/08-design-system-authority.md`

Script này THAY THẾ dev/audit_ui.py (gác chuẩn Lumina đã bỏ 2026-08-28).

── Vì sao không fail toàn bộ file legacy ──────────────────────────────────
Hiện 0/51 file đạt chuẩn T3. Một gate fail 51 file là gate không ai chạy.
Nên script chia file làm hai loại:

  T3-DECLARED — file đã merge T3Lab.Styles.xaml hoặc đã dùng {StaticResource T3.*}.
                Bị soi ĐẦY ĐỦ. Vi phạm ⇒ exit 1.
  LEGACY      — file chưa migrate. Chỉ ĐẾM và liệt kê nợ, KHÔNG làm fail gate.

Nhờ vậy gate xanh hôm nay, và siết dần theo từng file được migrate: file nào
khai T3 là file đó phải đúng T3.

Usage:
    python3 dev/audit_t3.py                 # báo cáo đầy đủ
    python3 dev/audit_t3.py --quiet         # chỉ vi phạm, exit 1 nếu có
    python3 dev/audit_t3.py --legacy        # liệt kê nợ của file legacy
    python3 dev/audit_t3.py --legacy-count  # in đúng 1 số: số file chưa migrate
    python3 dev/audit_t3.py --file <path>   # soi 1 file, luôn soi đầy đủ
"""
import os
import re
import sys
import glob
import xml.etree.ElementTree as ET

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(REPO, "T3Lab.extension", "lib", "GUI", "Tools")
STYLESHEET = os.path.join(REPO, "pyRevit UI Design System", "T3Lab.Styles.xaml")

# Thiết kế đã chốt riêng — ngoài phạm vi routine. Xem CLAUDE.md §UI-Frozen Files.
UI_LOCKED = {"DWGManagement.xaml", "ExportManager.xaml", "T3LabAssistant.xaml"}

# ── Chuẩn T3 (T3LAB_UI_STANDARD.md) ───────────────────────────────────────
FONTS_OK = {"Segoe UI", "Consolas"}
SIZES_OK = {"19", "15", "13", "11.5", "11", "12.5"}
SPACING_OK = {0, 4, 8, 12, 16, 24, 32}
RADIUS_OK = {0, 2, 4, 8}
WINDOW_ATTRS = ("UseLayoutRounding", "SnapsToDevicePixels", "MinWidth", "MinHeight")

# Palette của các hệ đã bị bỏ — dấu hiệu nhận dạng file legacy.
DEAD_PALETTE = {
    "#083D56": "Kinetix navy",
    "#0F766E": "Terra teal", "#115E59": "Terra teal", "#0B4F4A": "Terra teal",
    "#F6F8F8": "Terra", "#E6EDEC": "Terra", "#D9EBE8": "Terra",
    "#DDE5E7": "Terra", "#C7D2D4": "Terra", "#15803D": "Terra",
    "#166534": "Terra", "#AEBFBC": "Terra",
    "#0F172A": "Lumina ink", "#3B82F6": "Lumina accent", "#E2E8F0": "Lumina border",
    "#CBD5E1": "Lumina input border", "#64748B": "Lumina muted",
    "#94A3B8": "Lumina faint", "#F8FAFC": "Lumina surface",
    "#F1F5F9": "Lumina surface hover", "#10B981": "Lumina success",
    "#EF4444": "Lumina danger", "#1E293B": "Lumina primary hover",
}
DEAD_FONTS = ("Hanken Grotesk", "Inter", "Manrope")

HEX_RE = re.compile(r"^#(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$")
NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")

LIST_TAGS = {"DataGrid", "ListBox", "ListView"}


def local(tag):
    """Bỏ namespace khỏi tên tag/attribute."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def stylesheet_keys():
    if not os.path.exists(STYLESHEET):
        return None
    with open(STYLESHEET, encoding="utf-8", errors="replace") as fh:
        src = fh.read()
    return set(re.findall(r'x:Key="([^"]+)"', src))


def spacing_bad(value):
    """Trả về danh sách số không chia hết cho 4 trong Margin/Padding."""
    bad = []
    for n in NUM_RE.findall(value or ""):
        try:
            f = float(n)
        except ValueError:
            continue
        if f != int(f) or abs(int(f)) not in SPACING_OK:
            bad.append(n)
    return bad


def audit(src, base, keys):
    """Trả về (issues, declared_t3, legacy_debt).

    issues       — [(severity, message)] vi phạm chuẩn T3
    declared_t3  — file đã khai dùng T3 hay chưa
    legacy_debt  — [str] nợ migration (chỉ để đếm, không phải vi phạm)
    """
    issues = []
    debt = []

    used_keys = set(re.findall(r"\{StaticResource (T3\.[^}]+)\}", src))
    merges_sheet = "T3Lab.Styles.xaml" in src
    declared = merges_sheet or bool(used_keys)

    # ── Nợ migration (tính cho mọi file, kể cả file đã khai T3) ──────────
    for f in DEAD_FONTS:
        if re.search(r'FontFamily="[^"]*%s' % re.escape(f), src):
            debt.append("font đã bỏ: %s" % f)
    dead_found = {h: n for h, n in DEAD_PALETTE.items()
                  if re.search(re.escape(h), src, re.I)}
    if dead_found:
        debt.append("palette đã bỏ x%d (%s...)" % (
            len(dead_found), ", ".join(sorted(dead_found)[:3])))
    if "T3LAB SHARED STYLES" in src:
        debt.append("còn block shared styles Lumina")
    if "T3Theme" in src:
        debt.append("còn token Revit-native T3Theme*")
    if not merges_sheet:
        debt.append("chưa merge T3Lab.Styles.xaml")

    # ── Parse ────────────────────────────────────────────────────────────
    # Dot-notation bắt bằng regex TRƯỚC khi parse: ElementTree nuốt được nó,
    # nhưng WPF thì crash EMPTYPROPERTYELEMENT lúc mở tool.
    if re.search(r"<Grid\.(Row|Column)Definition\b", src):
        issues.append(("P0", "dot-notation <Grid.Row/ColumnDefinition> "
                             "— crash EMPTYPROPERTYELEMENT"))
    try:
        root = ET.fromstring(src)
    except ET.ParseError as exc:
        issues.append(("P0", "XAML không parse được: %s" % exc))
        return issues, declared, debt

    parent = {c: p for p in root.iter() for c in p}

    def ancestors(el):
        while el in parent:
            el = parent[el]
            yield el

    is_window = local(root.tag) == "Window"
    n_primary = 0
    has_default = has_cancel = False
    has_list = has_empty = False
    n_local_style = 0

    for el in root.iter():
        tag = local(el.tag)
        attrs = {local(k): v for k, v in el.attrib.items()}

        # Style tự chế trong file tool — chuẩn cấm, style mới phải vào stylesheet
        if tag == "Style" and "TargetType" in attrs:
            n_local_style += 1

        for name, val in attrs.items():
            if name == "FontFamily" and val not in FONTS_OK:
                issues.append(("P2", "<%s> FontFamily=\"%s\" — chỉ Segoe UI / Consolas"
                               % (tag, val)))
            elif name == "FontSize" and val not in SIZES_OK:
                issues.append(("P2", "<%s> FontSize=\"%s\" — ngoài thang 7 size"
                               % (tag, val)))
            elif name in ("Margin", "Padding"):
                bad = spacing_bad(val)
                if bad:
                    issues.append(("P3", "<%s> %s=\"%s\" — %s không thuộc 4/8/12/16/24/32"
                                   % (tag, name, val, "/".join(bad))))
            elif name == "CornerRadius":
                for n in NUM_RE.findall(val):
                    if float(n) not in RADIUS_OK:
                        issues.append(("P3", "<%s> CornerRadius=\"%s\" — chỉ 0/2/4/8"
                                       % (tag, val)))
                        break
            elif name == "Height" and tag == "TextBlock":
                issues.append(("P3", "<TextBlock Height=\"%s\"> — không set Height cho TextBlock"
                               % val))
            elif name == "IsDefault" and val == "True":
                has_default = True
            elif name == "IsCancel" and val == "True":
                has_cancel = True

            # Hex cứng ở bất kỳ attribute nào
            if HEX_RE.match(val.strip()):
                issues.append(("P2", "<%s %s=\"%s\"> — hex cứng, phải dùng {StaticResource T3.*}"
                               % (tag, name, val)))

        if tag.endswith(".Effect"):
            issues.append(("P2", "<%s> — chuẩn T3 cấm mọi Effect" % tag))

        if attrs.get("Style", "").strip() == "{StaticResource T3.Button.Primary}":
            n_primary += 1

        if tag in LIST_TAGS:
            has_list = True
            # T3.DataGrid / T3.ListBox / T3.LogBox đã set sẵn HorizontalScrollBarVisibility
            # và virtualization bằng Setter — dùng style đó là đã tuân thủ, đừng bắt
            # khai lại attribute trên element.
            styled = attrs.get("Style", "").strip() in (
                "{StaticResource T3.DataGrid}",
                "{StaticResource T3.ListBox}",
                "{StaticResource T3.LogBox}",
            )
            if not styled:
                if attrs.get("ScrollViewer.HorizontalScrollBarVisibility") != "Disabled" \
                        and attrs.get("HorizontalScrollBarVisibility") != "Disabled":
                    issues.append(("P2", "<%s> thiếu HorizontalScrollBarVisibility=\"Disabled\"" % tag))
                if tag == "DataGrid" and attrs.get("EnableRowVirtualization") != "True":
                    issues.append(("P1", "<DataGrid> thiếu EnableRowVirtualization=\"True\""))
                if tag in ("ListBox", "ListView") \
                        and attrs.get("VirtualizingPanel.IsVirtualizing") != "True":
                    issues.append(("P1", "<%s> thiếu VirtualizingPanel.IsVirtualizing=\"True\"" % tag))
            for anc in ancestors(el):
                if local(anc.tag) == "ScrollViewer":
                    issues.append(("P0", "<%s> bị bọc trong <ScrollViewer> "
                                         "— mất virtualization, treo Revit trên model lớn" % tag))
                    break
            stars = [c for c in el.iter()
                     if local(c.tag).endswith("Column")
                     and c.attrib.get("Width", "").strip() == "*"]
            if len(stars) > 1:
                issues.append(("P2", "<%s> có %d cột Width=\"*\" — chuẩn cho đúng một"
                               % (tag, len(stars))))

        if attrs.get("Style", "").strip() == "{StaticResource T3.Empty}":
            has_empty = True

    # ── Luật cấp file ────────────────────────────────────────────────────
    for k in sorted(used_keys):
        if keys is not None and k not in keys:
            issues.append(("P1", "{StaticResource %s} không có trong T3Lab.Styles.xaml" % k))

    if n_local_style:
        issues.append(("P2", "%d <Style> định nghĩa trong file tool — style mới phải "
                             "thêm vào T3Lab.Styles.xaml" % n_local_style))
    if n_primary > 1:
        issues.append(("P2", "%d nút T3.Button.Primary — chuẩn cho đúng một" % n_primary))
    if has_list and not has_empty:
        issues.append(("P2", "có list/grid nhưng không thấy empty state (T3.Empty)"))

    if is_window:
        wattrs = {local(k): v for k, v in root.attrib.items()}
        missing = [a for a in WINDOW_ATTRS if a not in wattrs]
        if "TextFormattingMode" not in " ".join(root.attrib.keys()):
            missing.append("TextOptions.TextFormattingMode")
        if missing:
            issues.append(("P2", "<Window> thiếu: %s" % ", ".join(missing)))
        if not has_default:
            issues.append(("P2", "không có nút IsDefault=\"True\""))
        if not has_cancel:
            issues.append(("P2", "không có nút IsCancel=\"True\""))
        if not merges_sheet:
            issues.append(("P1", "không merge T3Lab.Styles.xaml"))

    return issues, declared, debt


def main():
    argv = sys.argv[1:]
    quiet = "--quiet" in argv
    show_legacy = "--legacy" in argv
    count_only = "--legacy-count" in argv

    keys = stylesheet_keys()
    if keys is None and not count_only:
        print("CẢNH BÁO: không tìm thấy %s — bỏ qua kiểm tra key." % STYLESHEET)

    if "--file" in argv:
        files = [argv[argv.index("--file") + 1]]
        forced = True
    else:
        files = sorted(glob.glob(os.path.join(TOOLS, "*.xaml")))
        forced = False

    n_t3 = n_legacy = n_locked = 0
    failed = False
    legacy_rows = []
    out = []

    for path in files:
        base = os.path.basename(path)
        if base in UI_LOCKED and not forced:
            n_locked += 1
            continue
        with open(path, encoding="utf-8", errors="replace") as fh:
            src = fh.read()
        issues, declared, debt = audit(src, base, keys)

        if declared or forced:
            n_t3 += 1
            if issues:
                failed = True
                out.append((base, issues))
        else:
            n_legacy += 1
            legacy_rows.append((base, debt))

    if count_only:
        print(n_legacy)
        return 0

    for base, issues in out:
        print("%s" % base)
        for sev, msg in sorted(issues, key=lambda i: i[0]):
            print("   [%s] %s" % (sev, msg))

    if show_legacy and not quiet:
        print()
        print("── LEGACY — chưa migrate, KHÔNG tính là vi phạm ──")
        for base, debt in legacy_rows:
            print("%-38s %s" % (base, " · ".join(debt) if debt else "—"))

    print()
    print("T3 AUDIT: %d file khai T3 (%d file có vi phạm) · %d file legacy · %d UI-locked"
          % (n_t3, len(out), n_legacy, n_locked))
    if n_legacy and not show_legacy:
        print("          chạy --legacy để xem nợ migration của %d file kia" % n_legacy)
    print("          %s" % ("FAILED" if failed else "OK"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
