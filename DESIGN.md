# T3Lab UI Design System

> **Chuẩn UI duy nhất nằm ở `pyRevit UI Design System/`.** File này chỉ là biển chỉ đường.

| Cần gì | Đọc file nào |
|--------|--------------|
| Luật đầy đủ (token, thang type, spacing, 5 pattern P1–P5, size class S/M/L) | `pyRevit UI Design System/T3LAB_UI_STANDARD.md` |
| Stylesheet — 106 key `T3.*` | `pyRevit UI Design System/T3Lab.Styles.xaml` |
| Component thật, render được, để copy | `T3Lab.extension/lib/GUI/Tools/UIStandardShowcase.xaml` |
| Luật cho mọi tool/script mới | `.claude/rules/new-tool-standard.md` |
| Snippet lẻ | `.claude/skills/xaml-templates.md` |
| Quản trị & migration | `docs/ui-governance/` (bắt đầu ở `README.md`) |

Gate trước khi commit:

```
python3 dev/sync_t3_styles.py --check
python3 dev/audit_t3.py --quiet
python3 dev/audit_tools.py --quiet
```

## File mẫu — UI Standard Showcase

`UIStandardShowcase.xaml` là **MỘT cửa sổ T3 hợp lệ** (một title bar, một rail, một
footer, một dòng copyright, một nút primary) chứa toàn bộ component, chia 5 nhóm đổi
bằng rail bên trái:

| Rail | Nhóm | Có gì |
|------|------|-------|
| 0 | Foundations | token màu, thang 7 size, spacing 4/8/12/16/24/32, radius 0/2/4/8 |
| 1 | Controls | 4 vai trò nút, TextBox/Mono/Search, ComboBox, CheckBox, RadioButton, Chip, Expander, Pill |
| 2 | Data | summary strip + meter, list header + ListBox, DataGrid + status pill + empty state |
| 3 | Feedback | ProgressBar, log 5 mức severity, tally, 4 callout, 4 dòng status |
| 4 | Patterns | công thức P1–P5 |

Mở nó trong Revit (`T3Lab › Standard › UI Showcase`), tìm component cần dùng, copy
khối tương ứng trong XAML.

---

## Lịch sử

Chuẩn **Lumina** (và trước đó Revit-native, Terra v2, Kinetix) đã bị bỏ ngày
2026-08-28. Bản đặc tả Lumina từng nằm trong chính file này; lấy lại từ git history
nếu cần đối chiếu. Không dùng nó cho code mới.
