# 01 · SCAN — quét hệ thống

Mục tiêu bước SCAN: biết **có gì**, **cái gì vừa đổi**, và **cycle này soi cái nào**.
Không đánh giá, không sửa ở bước này.

---

## 1 · Phạm vi quét

| Loại | Vị trí | Ghi chú |
|------|--------|---------|
| Tool / pushbutton | `T3Lab.extension/T3Lab.tab/*.panel/**/script.py` | 42 script, 7 panel |
| UI window (XAML) | `T3Lab.extension/lib/GUI/Tools/*.xaml` | 54 file |
| Python dialog class | `T3Lab.extension/lib/GUI/*.py` | `*Dialog.py`, `WPF_Base.py`, `forms.py` |
| Shared UI component | `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml` | master Lumina |
| Design System | `pyRevit UI Design System/` | chuẩn cao nhất |
| Helper / snippet | `T3Lab.extension/lib/Snippets/` | message box, progress, validation |
| Icon / ảnh | `T3Lab.extension/**/*.png`, `draft_icons/` | kích thước & độ tương phản |
| Check script | `T3Lab.extension/checks/`, `commands/` | UI dạng output/log |

Ngoài XAML còn phải quét **các mặt UI không nằm trong XAML**:

- `forms.alert` / `TaskDialog` / `MessageBox` — nội dung & giọng văn thông báo.
- `print()` ra pyRevit output window — có phải là UI không chủ đích?
- `ProgressBar` / `forms.ProgressBar` — có feedback cho tác vụ dài không?
- `try/except` nuốt lỗi — người dùng thấy gì khi hỏng?
- Tooltip / `__doc__` / `bundle.yaml` title — nội dung ribbon.

## 2 · Kiểm kê tự động

```bash
# Danh sách XAML + số dòng
wc -l T3Lab.extension/lib/GUI/Tools/*.xaml

# File đã đổi kể từ cycle trước (thay <SHA> bằng last_commit trong BASELINE.md)
git diff --name-only <SHA>..HEAD -- T3Lab.extension/ 'pyRevit UI Design System/'

# Tool mới chưa có trong BASELINE.md
ls T3Lab.extension/lib/GUI/Tools/*.xaml | xargs -n1 basename

# Gate code tĩnh — chạy TRƯỚC khi sửa gì, để biết trạng thái nền
python3 dev/audit_tools.py --quiet

# Đếm legacy còn lại (KHÔNG dùng làm pass/fail — nó gác chuẩn Lumina đã bỏ)
python3 dev/audit_ui.py            # chỉ đọc số, bỏ qua exit code
```

> `dev/audit_ui.py` và `dev/sync_wpf_styles.py` gác chuẩn **Lumina đã bị bỏ**. Chúng
> sẽ báo fail đúng những file vừa migrate sang T3 — đó là kết quả mong đợi, không phải
> lỗi. Xem `08-design-system-authority.md` §8.

> Nếu `audit_tools.py` đã **đỏ sẵn từ trước** khi bạn chưa sửa gì: ghi rõ trong report
> là pre-existing, đừng nhận là do cycle này gây ra, và đưa lên đầu queue.

## 3 · Phân loại mỗi file

Với mỗi XAML, xác định 4 thuộc tính và ghi vào BASELINE:

| Thuộc tính | Giá trị |
|------------|---------|
| **Compliance** | `T3` (đã đạt chuẩn) · `LEGACY-LUMINA` · `LEGACY-REVIT-NATIVE` · `LEGACY-TERRA` (xem `08-design-system-authority.md` §2) |
| **Variant** | `A` (root `<Window>`) · `B` (root `<Grid>`, modal content) · `ITEM` (row template) |
| **Pattern** | `P1` form · `P2` selection list · `P3` progress & log · `P4` results table · `P5` confirmation · `MIXED` · `NONE` |
| **Size class** | `S` 420×260–320 · `M` 560×420–560 · `L` 1000×620 · `OFF` (ngoài chuẩn) |

`MIXED` và `NONE` tự nó là một finding (T3 Standard: *mọi tool phải là một trong
P1–P5*) — nhưng là finding cấp **P3/đề xuất**, vì sửa nó là redesign.

## 4 · Thứ tự ưu tiên audit sâu trong cycle

Chọn **5–8 tool**, theo đúng thứ tự này:

1. **Tool mới** hoặc XAML vừa thay đổi kể từ cycle trước (git diff).
2. **Tool có UI inconsistency** đã biết (còn issue mở trong BASELINE).
3. **Tool có bug** đang mở (P0/P1 trong `PRIORITY_QUEUE.md`).
4. **Tool score thấp** (dưới 70).
5. **Tool lâu chưa audit nhất** (`last_audit` cũ nhất) — đảm bảo mọi tool được xoay vòng.

Nếu số ứng viên vượt 8: lấy 8 đầu, phần còn lại đẩy sang `PRIORITY_QUEUE.md`.
Nếu ít hơn 5: lấp đầy bằng mục 5 (xoay vòng).

## 5 · Đầu ra của bước SCAN

Một bảng ngắn, đưa thẳng vào report:

```
SCAN — cycle N
Tổng: 54 XAML  →  T3-compliant 0 · legacy 51 · LOCKED 3
Thay đổi từ cycle trước: 3 file  (AutoDimension.xaml, ManaSheets.xaml, TagChecker.xaml)
Tool mới: 0
Gate nền: audit_tools ✅   (audit_ui/sync_wpf_styles = legacy, chỉ đếm)
Chọn audit sâu cycle này (6): AutoDimension, ManaSheets, TagChecker,
                              PointCloud (score 64), SplitElements (score 68),
                              RibbonNames (chưa audit từ 2026-07-05)
Chọn migrate cycle này (2):   AutoDimension (P1, M), TagChecker (P4, L)
```
