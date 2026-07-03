# Panel 5 — Standards & Settings (Ngày 10)

> 4 tool nhưng đụng vào **cài đặt toàn model** (coordinates, object styles, workset).
> ⚠️ BẮT BUỘC test trên **bản copy** của model, không dùng file thật. Cần thêm 1 file workshared (tạo local copy) cho ManaWorkset.

## Tool 1/4 — ManaLoca ⚠️ nguy hiểm nhất panel
Chain: self-contained (871 loc) → `ManaLoca.xaml` · thao tác site/shared coordinates

- [ ] **Chỉ chạy trên file copy**
- [ ] Đọc/hiển thị đúng Project Base Point, Survey Point, shared site hiện tại
- [ ] Đổi 1 giá trị toạ độ → model dịch chuyển đúng hướng/đơn vị
- [ ] Ctrl+Z revert được thay đổi toạ độ
- [ ] File có nhiều shared site → list đủ, switch đúng
- [ ] Huỷ giữa chừng → không để model ở trạng thái toạ độ dở dang
- [ ] Ghi chú:

## Tool 2/4 — ManaStyles
Chain: launcher → `ManaStylesDialog.py` (2091 loc) → `ManaStyles.xaml`

- [ ] List object styles/line weights/line patterns đúng với Manage > Object Styles
- [ ] Đổi 1 style → áp dụng toàn model, Ctrl+Z OK
- [ ] Category không cho sửa (system) → disable hoặc thông báo
- [ ] Import/export style config (nếu có) → file đúng format
- [ ] Ghi chú:

## Tool 3/4 — ManaWorkset
Chain: self-contained (609 loc, 3 transaction) → `ManaWorkset.xaml`

- [ ] **File thường (không workshared)** → thông báo thân thiện yêu cầu enable worksharing, KHÔNG stacktrace (đây là edge case hay quên nhất)
- [ ] File workshared: list workset đúng
- [ ] Tạo workset mới → xuất hiện trong Revit
- [ ] Rename workset → OK; workset default (Workset1, Shared Levels and Grids) → xử lý đúng giới hạn Revit
- [ ] Element thuộc workset của user khác (chưa checkout) → thông báo ownership
- [ ] Ghi chú:

## Tool 4/4 — ModelAuditor
Chain: launcher → `ModelAuditorDialog.py` (1828 loc) → `ModelAuditor.xaml`

- [ ] Model nhỏ → report đầy đủ các hạng mục audit, số liệu đúng (đối chiếu tay 1–2 mục: số warning, số family in-place...)
- [ ] Model lớn → đo thời gian chạy, UI không treo (có progress?)
- [ ] Export report (nếu có) → file ra đúng
- [ ] Model rỗng → report không crash
- [ ] Ghi chú:

## Chốt ngày 10
- [ ] Cập nhật bảng + README · Commit: `fix(panel-standards): <mô tả>`

| Tool | Trạng thái | Ghi chú |
|------|-----------|---------|
| ManaLoca | ⬜ | |
| ManaStyles | ⬜ | |
| ManaWorkset | ⬜ | Gray-gutter fix (xem Phát sinh) đã sửa 2026-07-03, chưa test functional |
| ModelAuditor | ⬜ | |

## Phát sinh

- **2026-07-03 — ManaWorkset: gray gutter quanh vùng nội dung** (yêu cầu UI trực tiếp từ user, không phải bug): `ManaWorkset.xaml` có `TabControl x:Name="tab_control"` (Grid.Row="1"/Col="1") với `Background="Transparent"` → nền xám cửa sổ `#E4E4E7` lộ ra qua mọi khe hở giữa các card trắng bên trong (cùng lớp bug đã gặp ở ManaSelect/ManaDWG/PDFImport, xem `panel-1-annotation-select.md` Phát sinh). Đã sửa theo đúng pattern chuẩn: bọc toàn bộ `TabControl` trong 1 `Border Background="#F8FAFC"` full-bleed (không margin) phủ kín ô Grid.Row=1/Col=1, không còn khe hở nào lộ màu xám gốc. Verify: XML well-formed, `audit_tools.py --quiet` clean, `sync_wpf_styles.py --check` 53/53 khớp. Chưa được user xác nhận trực quan trong Revit.
- **2026-07-03 (tiếp) — ManaWorkset: user báo tiếp "remove all gutter or gap" sau lần fix nền F8FAFC đầu tiên** — lần fix trước chỉ đổi màu (xám → F8FAFC) chứ chưa bỏ khoảng hở kết cấu: cả 3 tab (Worksets/Bulk Tools/3D Views) đều dùng pattern "floating card" — card nội dung chính có `Margin` (18px các cạnh) + `CornerRadius="20"` nổi giữa nền, để lộ viền F8FAFC quanh card dù không còn xám. Áp dụng đúng pattern "full-screen" đã dùng cho DWGManagement/ManaSelect trước đây (xem `panel-1-annotation-select.md` Phát sinh, mục "user yêu cầu áp dụng cách full-screen"): bỏ `Margin`/`CornerRadius` của card chính ở cả 3 tab, card giờ phủ kín edge-to-edge, chỉ giữ `BorderThickness="1"` (viền mảnh, không bo góc, không margin) — Tab 1 "Worksets card" (`Margin="18,14,18,18"` → `"0"`, `CornerRadius="20"` → `"0"`), Tab 2 "Bulk Tools" (`Grid Margin="18"` → bỏ, `CornerRadius="20"` → `"0"`), Tab 3 "3D Views" (tương tự Tab 2). Giữ nguyên `Padding` nội bộ của card (đó là khoảng cách nội dung bên trong, không phải gutter lộ nền ngoài) và banner cảnh báo "Enable Worksharing" ở Tab 1 (alert box nhỏ, có margin riêng theo đúng chuẩn thiết kế, không phải là "cái bảng" đang bị phàn nàn). Verify: XML well-formed, `audit_tools.py --quiet` clean, `sync_wpf_styles.py --check` 53/53 khớp. Chưa được user xác nhận trực quan trong Revit.
