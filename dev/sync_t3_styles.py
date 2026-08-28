#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Đồng bộ stylesheet T3 từ thư mục nguồn sang bản deploy trong extension.

  NGUỒN  : pyRevit UI Design System/T3Lab.Styles.xaml   (file mẫu để refer, sửa ở đây)
  DEPLOY : T3Lab.extension/lib/GUI/Resources/T3Lab.Styles.xaml  (bản pyRevit nạp lúc chạy)

Vì sao cần hai bản: thư mục `pyRevit UI Design System/` nằm NGOÀI `T3Lab.extension/`,
nên khi extension được cài riêng (%APPDATA%\\pyRevit\\Extensions) pyRevit không thấy nó.
Tool XAML merge bản deploy bằng `<ResourceDictionary Source="../Resources/T3Lab.Styles.xaml"/>`.

Usage:
    python3 dev/sync_t3_styles.py           # copy nguồn -> deploy
    python3 dev/sync_t3_styles.py --check   # chỉ kiểm tra, exit 1 nếu lệch
"""
import os
import sys
import shutil

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(REPO, "pyRevit UI Design System", "T3Lab.Styles.xaml")
DST = os.path.join(REPO, "T3Lab.extension", "lib", "GUI", "Resources", "T3Lab.Styles.xaml")

BANNER = ("<!-- BẢN DEPLOY — SINH TỰ ĐỘNG, ĐỪNG SỬA Ở ĐÂY.\n"
          "     Sửa `pyRevit UI Design System/T3Lab.Styles.xaml` rồi chạy "
          "`python3 dev/sync_t3_styles.py`. -->\n")


def read(path):
    with open(path, encoding="utf-8") as fh:
        return fh.read()


def main():
    check = "--check" in sys.argv
    if not os.path.exists(SRC):
        print("KHÔNG THẤY NGUỒN: %s" % SRC)
        return 1

    want = BANNER + read(SRC)
    have = read(DST) if os.path.exists(DST) else None

    if have == want:
        print("T3 STYLES: đồng bộ (%d dòng)" % want.count("\n"))
        return 0

    if check:
        print("T3 STYLES: LỆCH — chạy `python3 dev/sync_t3_styles.py` để đồng bộ")
        return 1

    os.makedirs(os.path.dirname(DST), exist_ok=True)
    with open(DST, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(want)
    print("T3 STYLES: đã cập nhật %s" % os.path.relpath(DST, REPO))
    return 0


if __name__ == "__main__":
    sys.exit(main())
