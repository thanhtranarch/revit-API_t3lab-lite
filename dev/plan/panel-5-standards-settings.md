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
| ManaWorkset | ⬜ | |
| ModelAuditor | ⬜ | |

## Phát sinh

_(ghi lỗi mới tại đây)_
