#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Nhúng (inline) stylesheet T3 vào mọi tool XAML đã khai dùng T3.

  NGUỒN : pyRevit UI Design System/T3Lab.Styles.xaml  (file mẫu để refer, sửa ở đây)
  ĐÍCH  : khối giữa 2 marker trong <Window.Resources> của từng tool XAML

── Vì sao INLINE chứ không MergedDictionaries ──────────────────────────────
Đã thử `<ResourceDictionary Source="../Resources/T3Lab.Styles.xaml"/>` và pyRevit
CHẾT lúc mở tool (2026-08-28, BatchOut):

    System.IO.IOException: Assembly.GetEntryAssembly() returns null.
    Set the Application.ResourceAssembly property or use the
    pack://application:,,,/assemblyname;component/ syntax ...

pyRevit nạp XAML bằng XamlReader trên một stream, nên WPF không có base URI và
không có entry assembly; mọi `Source` tương đối bị resolve thành pack://application
rồi ném IOException. Nạp dictionary từ Python sau đó cũng không cứu được, vì
`{StaticResource}` trong XAML được resolve NGAY LÚC PARSE, trước khi Python chạy.

Inline là cách duy nhất còn lại — và là cách repo này vốn đã dùng (dev/sync_wpf_styles.py
của chuẩn Lumina cũ làm đúng như vậy). Đổi lại: mỗi tool XAML nặng thêm ~1000 dòng
sinh tự động, và KHÔNG được sửa tay khối đó.

Usage:
    python3 dev/sync_t3_styles.py           # nhúng vào mọi file đã khai T3
    python3 dev/sync_t3_styles.py --check   # chỉ kiểm tra, exit 1 nếu lệch
"""
import os
import re
import sys
import glob

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(REPO, "pyRevit UI Design System", "T3Lab.Styles.xaml")
TOOLS = os.path.join(REPO, "T3Lab.extension", "lib", "GUI", "Tools")

BEGIN = ("<!-- ═══ T3 STYLES — SINH TỰ ĐỘNG, ĐỪNG SỬA Ở ĐÂY. "
         "Sửa `pyRevit UI Design System/T3Lab.Styles.xaml` rồi chạy "
         "`python3 dev/sync_t3_styles.py` ═══ -->")
END = "<!-- ═══ HẾT T3 STYLES ═══ -->"
SYS_NS = 'xmlns:sys="clr-namespace:System;assembly=mscorlib"'


def styles_body():
    """Nội dung BÊN TRONG <ResourceDictionary> của stylesheet nguồn."""
    with open(SRC, encoding="utf-8") as fh:
        src = fh.read()
    # Tìm ĐÚNG root element: comment hướng dẫn ở đầu file cũng chứa chuỗi
    # "<ResourceDictionary", nên không được dùng index() thẳng.
    m = re.search(r"^<ResourceDictionary\b", src, re.M)
    if not m:
        raise SystemExit("không thấy root <ResourceDictionary> trong %s" % SRC)
    start = src.index(">", m.start()) + 1
    end = src.rindex("</ResourceDictionary>")
    body = src[start:end].strip("\n")
    return "\n".join(("        " + l) if l.strip() else l for l in body.split("\n"))


def block(body):
    return "        %s\n%s\n        %s" % (BEGIN, body, END)


def targets():
    """Tool XAML đã khai dùng T3 (có marker, hoặc dùng {StaticResource T3.*})."""
    out = []
    for p in sorted(glob.glob(os.path.join(TOOLS, "*.xaml"))):
        with open(p, encoding="utf-8") as fh:
            s = fh.read()
        if BEGIN in s or re.search(r"\{StaticResource T3\.", s):
            out.append(p)
    return out


def apply(path, body, check):
    with open(path, encoding="utf-8") as fh:
        s = fh.read()
    want = block(body)

    if BEGIN in s:
        i = s.index(BEGIN)
        j = s.index(END, i) + len(END)
        cur = s[s.rindex("\n", 0, i) + 1:j]
        if cur == want:
            return "ok"
        if check:
            return "lệch"
        s = s[:s.rindex("\n", 0, i) + 1] + want + s[j:]
    else:
        if check:
            return "thiếu block"
        m = re.search(r"[ \t]*<([A-Za-z]+)\.Resources>[ \t]*\n[ \t]*<ResourceDictionary>[ \t]*\n", s)
        if not m:
            return "KHÔNG THẤY <*.Resources><ResourceDictionary>"
        s = s[:m.end()] + want + "\n" + s[m.end():]

    # sys: namespace bắt buộc — stylesheet khai <sys:Double> cho các size token
    if SYS_NS not in s:
        s = s.replace('xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"',
                      'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"\n        ' + SYS_NS, 1)

    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(s)
    return "đã cập nhật"


def main():
    check = "--check" in sys.argv
    if not os.path.exists(SRC):
        print("KHÔNG THẤY NGUỒN: %s" % SRC)
        return 1

    body = styles_body()
    files = targets()
    if not files:
        print("T3 STYLES: chưa tool nào khai dùng T3 — không có gì để nhúng")
        return 0

    bad = 0
    for p in files:
        r = apply(p, body, check)
        if r != "ok":
            bad += 1
            print("  %-34s %s" % (os.path.basename(p), r))

    print("T3 STYLES: %d file · %d lệch (%d dòng style mỗi file)"
          % (len(files), bad, body.count("\n") + 1))
    if check and bad:
        print("           chạy `python3 dev/sync_t3_styles.py` để đồng bộ")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
