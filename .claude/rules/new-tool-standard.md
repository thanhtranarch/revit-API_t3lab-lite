# Luật cho MỌI tool / script mới

> Áp dụng cho mọi pushbutton, dialog, XAML **tạo mới từ 2026-08-28**.
> Chuẩn UI duy nhất: `pyRevit UI Design System/T3LAB_UI_STANDARD.md` + `T3Lab.Styles.xaml`.
> File cũ chưa migrate không bị luật này ràng buộc — chúng đi theo hàng đợi ở
> `docs/ui-governance/PRIORITY_QUEUE.md`.

Gate: `python3 dev/audit_t3.py --quiet` phải xanh trước khi commit.

---

## 0 · Trước khi viết dòng đầu tiên

Trả lời 3 câu. Không trả lời được thì chưa viết code.

| Câu hỏi | Lựa chọn |
|---------|----------|
| Tool này là **pattern** nào? | P1 form · P2 selection list · P3 progress & log · P4 results table · P5 confirmation |
| **Size class** nào? | S 420×260–320 (NoResize) · M 560×420–560 · L 1000×620 |
| Nó **sửa model** không? | Có → bắt buộc có P5 confirm + câu trạng thái nói rõ số lượng |

Không có pattern nào vừa → dừng lại, ghi `DESIGN SYSTEM GAP` vào
`docs/ui-governance/PRIORITY_QUEUE.md`, hỏi trước khi tự chế pattern thứ 6.

---

## 1 · Cấu trúc file — đặt đúng chỗ

```
T3Lab.extension/
├── T3Lab.tab/<Panel>.panel/<Tool>.pushbutton/
│   ├── script.py          ← entry point, KHÔNG chứa logic Revit nặng
│   ├── icon.png           ← 32×32
│   └── bundle.yaml        ← title + tooltip
├── lib/GUI/Tools/<Tool>.xaml       ← MỌI file .xaml nằm ở đây, không ngoại lệ
├── lib/GUI/<Tool>Dialog.py         ← class WPF (nếu tool đủ lớn để tách)
└── lib/Snippets/                   ← helper Revit API dùng lại được
```

**Tách bạch bắt buộc:** XAML không biết gì về Revit API; `script.py` / `*Dialog.py`
không hardcode màu, size, margin. Logic Revit dùng lại được thì đẩy vào `lib/Snippets/`.

---

## 2 · XAML — 15 luật, `audit_t3.py` kiểm tra tự động

| # | Luật | Vi phạm |
|---|------|---------|
| 1 | Nhúng stylesheet bằng `dev/sync_t3_styles.py`; **không** `<Style x:Key>` nào tự viết trong file tool (implicit style không key thì được) | P1/P2 |
| 2 | Không hex cứng — mọi màu là `{StaticResource T3.*}` | P2 |
| 3 | Font chỉ `Segoe UI` (text) và `Consolas` (số, ID, log) | P2 |
| 4 | FontSize chỉ `19 · 15 · 13 · 11.5 · 11 · 12.5` | P2 |
| 5 | `Margin`/`Padding` chỉ `4 / 8 / 12 / 16 / 24 / 32` | P3 |
| 6 | `CornerRadius` chỉ `8` window · `4` control · `2` pill · `0` grid row | P3 |
| 7 | `<Window>` có `UseLayoutRounding` · `SnapsToDevicePixels` · `TextOptions.TextFormattingMode="Display"` · `MinWidth` · `MinHeight` | P2 |
| 8 | Có đúng **một** nút `T3.Button.Primary`, ngoài cùng phải footer | P2 |
| 9 | Có nút `IsDefault="True"` **và** nút `IsCancel="True"` | P2 |
| 10 | Mọi list/grid: `HorizontalScrollBarVisibility="Disabled"`; DataGrid thêm `EnableRowVirtualization="True"` | P1/P2 |
| 11 | **Không bọc** `DataGrid`/`ListBox`/`ListView` trong `ScrollViewer` | **P0** |
| 12 | Có empty state (`T3.Empty`) cho mọi list/grid | P2 |
| 13 | **Copyright BẮT BUỘC**: đúng MỘT `<TextBlock Style="{StaticResource T3.Copyright}"/>` ở footer, sát trái, trước câu trạng thái | P2 |
| 15 | Không `x:Key` trùng nhau — WPF crash ngay lúc parse | **P0** |
| 14 | **Mọi chữ hiển thị phải là TIẾNG ANH** — `Text` · `Content` · `Header` · `ToolTip` · empty state · thông báo lỗi | P2 |

Thêm hai thứ `audit_t3.py` cũng bắt: `<Grid.RowDefinition/>` dot-notation (**P0**,
crash `EMPTYPROPERTYELEMENT` lúc mở tool) và mọi `Effect` (P2).

### Khung `<Window>` chuẩn

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:sys="clr-namespace:System;assembly=mscorlib"
        Title="Tool name" Height="420" Width="560" MinHeight="420" MinWidth="560"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display"
        Background="{DynamicResource T3.Canvas}"
        WindowStartupLocation="CenterOwner">
  <Window.Resources>
    <ResourceDictionary>
      <!-- chạy `python3 dev/sync_t3_styles.py` để nhúng stylesheet vào đây -->
    </ResourceDictionary>
  </Window.Resources>
  <!-- nội dung + footer có copyright: xem .claude/skills/xaml-templates.md -->
</Window>
```

> **Nhúng, KHÔNG dùng `MergedDictionaries`.** pyRevit nạp XAML bằng `XamlReader` trên
> một stream nên WPF không có base URI và không có entry assembly; mọi
> `<ResourceDictionary Source="..."/>` tương đối đều ném
> `IOException: Assembly.GetEntryAssembly() returns null` và tool chết ngay lúc mở
> (đã thử với BatchOut 2026-08-28). Nạp dictionary từ Python sau đó cũng không cứu
> được, vì `{StaticResource}` được resolve NGAY LÚC PARSE.
> Chạy `python3 dev/sync_t3_styles.py` để nhúng; `--check` để verify. **Không sửa tay
> khối giữa 2 marker `T3 STYLES`.**

---

## 3 · `script.py` — khung bắt buộc

```python
# -*- coding: utf-8 -*-
"""<Tên tool> — <một câu tool này làm gì>."""
__title__ = 'Tên\nTool'
__author__ = 'T3Lab'

# ── IMPORTS ──────────────────────────────────────────────────────────────
import os
from pyrevit import forms, script

from Autodesk.Revit.DB import Transaction, FilteredElementCollector

# ── CONSTANTS ────────────────────────────────────────────────────────────
doc = __revit__.ActiveUIDocument.Document          # noqa: F821
uidoc = __revit__.ActiveUIDocument                 # noqa: F821
XAML_FILE = os.path.join(os.path.dirname(__file__),
                         '..', '..', '..', 'lib', 'GUI', 'Tools', '<Tool>.xaml')

# ── WINDOW ───────────────────────────────────────────────────────────────
class ToolWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._load()                    # nạp dữ liệu NGAY, không đợi Loaded event

    # ...

# ── ENTRY POINT ──────────────────────────────────────────────────────────
if __name__ == '__main__':
    ToolWindow().ShowDialog()
```

### Luật `script.py`

| # | Luật | Vì sao |
|---|------|--------|
| S1 | Một `Transaction` cho một thao tác người dùng | Ctrl+Z revert đúng **một** bước |
| S2 | `t.Start()` / `t.Commit()` trong `try/except` có `t.RollBack()` | Lỗi giữa chừng không để model dở dang |
| S3 | Không bắt `except:` trần — bắt đúng loại, và **hiện lỗi cho người dùng** | Exception bị nuốt = tool "không làm gì" mà không ai biết |
| S4 | Không `print()` trong luồng bình thường | Rác ra output window |
| S5 | Tác vụ >2 giây phải có progress (pattern P3) | Revit đóng băng im lặng là bug |
| S6 | Không chọn gì / model rỗng → thông báo thân thiện, không stacktrace | |
| S7 | Nạp dữ liệu grid trong `__init__`, không trong `Loaded` | Tránh trang trắng lúc mở |
| S8 | Mọi `x:Name` dùng trong Python phải tồn tại trong XAML | Sai = `AttributeError` lúc runtime |
| S9 | Không thêm dependency mới | IronPython 2.7 + .NET 4.8 + stock WPF |

---

## 4 · Nội dung hiển thị

- Error message nói đủ ba phần: **cái gì sai · ở đâu · làm gì tiếp**. Không mã lỗi trần.
- Warning và success luôn kèm **số lượng**: "Đã đổi tên 34 sheet", không phải "Xong".
- Status **không bao giờ chỉ bằng màu** — chấm màu phải đi kèm chữ (`ok` / `skipped` / `failed`).
- **Ngôn ngữ UI là TIẾNG ANH**, không ngoại lệ — kể cả tool chỉ người Việt dùng.
  Comment trong code và tài liệu thì tiếng Việt vẫn được.
- Thuật ngữ Revit giữ nguyên tiếng Anh: View Template, Workset, Sheet, Family.

---

## 5 · Checklist trước khi commit

```
[ ] python3 dev/audit_t3.py --quiet     → xanh
[ ] python3 dev/audit_tools.py --quiet  → xanh
[ ] Pattern P1–P5 rõ ràng, size class đúng S/M/L
[ ] Mở tool trong Revit: không lỗi, chrome hoạt động, không trang trắng
[ ] Happy path đúng; Ctrl+Z revert đúng một bước
[ ] Edge case (không chọn gì / model rỗng / bấm Cancel) ra câu tiếng người
[ ] Đọc được ở 100% VÀ 125% display scaling, không cắt chữ
[ ] Thao tác phá huỷ có P5 confirm, nút đỏ KHÔNG phải IsDefault
[ ] Footer có đúng một dòng © Copyright by T3Lab, sát trái
[ ] Không còn chữ tiếng Việt nào hiển thị cho người dùng
```

Bốn dòng cuối cần Revit thật. Chưa test thì ghi `NEEDS VERIFICATION`, **không tick**.
