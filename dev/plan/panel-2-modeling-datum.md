# Panel 2 — Modeling & Datum (Ngày 6–7)

> 14 tool — panel lớn nhất, chia 2 ngày. Model test: file có room, wall (thẳng + cong), floor, door, level, grid; chuẩn bị sẵn 1 file DWG mẫu, 1 ảnh PNG đen trắng, 1 file point (nếu có).

---

## Ngày 6 — Nhóm Create (7 tool)

### Tool 1/7 — CADToElements
Chain: launcher → `CADToElementsDialog.py` (3143 loc) → `CADToElements.xaml` + 4 XAML con (CADtoBeam / CadtoFloor / CadtoFloorLayerItem / CadtoWall)

- [ ] Mở cửa sổ chính + từng flow con (Wall / Floor / Beam) không lỗi XAML
- [ ] Import DWG mẫu → convert Wall theo layer → wall tạo đúng vị trí/đơn vị
- [ ] Convert Floor (có layer item list — CadtoFloorLayerItem variant B load đúng)
- [ ] Convert Beam
- [ ] DWG không có layer khớp → thông báo, không crash
- [ ] Ctrl+Z sau mỗi convert → revert 1 bước
- [ ] Ghi chú:

### Tool 2/7 — DoorThreshold
Chain: self-contained (356 loc, **3 bare except**) → `DoorThreshold.xaml`

- [ ] Trước khi test: tạm thêm traceback log vào 3 nhánh `except:` để lỗi không bị nuốt
- [ ] Door trên wall thường → threshold tạo đúng
- [ ] Door trên curtain wall / wall nghiêng → hành vi hợp lý
- [ ] Thiếu family threshold trong model → thông báo rõ ràng
- [ ] Ghi chú:

### Tool 3/7 — ImageToDrafting
Chain: self-contained (1585 loc) → `ImageToDrafting.xaml` · **compile C# Vectorizer lúc runtime**

- [ ] Ảnh PNG đen trắng nhỏ (~500px) → mode Centerline → drafting lines tạo ra
- [ ] Mode Outline (không cần compile C#) → hoạt động
- [ ] Xác nhận compile C# thành công trên máy triển khai (không chỉ máy dev) — nếu fail phải ra alert "Failed to compile C# vectorizer" chứ không crash
- [ ] Ảnh màu / ảnh lớn → threshold tự động hợp lý, không treo Revit
- [ ] Ghi chú:

### Tool 4/7 — Text to Element
Chain: launcher → `TextToElementDialog.py` (466 loc) → `TextToElement.xaml`

- [ ] Text note hợp lệ → element tạo đúng
- [ ] Text rác / rỗng → validate, không crash
- [ ] Ghi chú:

### Tool 5/7 — PointCloud
Chain: self-contained (1502 loc, 6 transaction) → `PointCloud.xaml`

- [ ] File điểm nhỏ hợp lệ → import/xử lý đúng
- [ ] File sai format → thông báo lỗi rõ
- [ ] Cancel giữa chừng → transaction rollback sạch (kiểm tra Undo stack không có rác)
- [ ] Ghi chú:

### Tool 6/7 — RoomToFloor
Chain: self-contained (383 loc) → `RoomToFloor.xaml`

- [ ] Room chữ nhật → floor đúng boundary, offset mm đúng
- [ ] Room có cột giữa (island) → floor có opening (test cả Revit <2022 nếu dùng: 2 nhánh code theo REVIT_VERSION)
- [ ] Room chưa placed / area = 0 → skip đếm vào error_count, không crash
- [ ] Toàn bộ batch = 1 TransactionGroup → Ctrl+Z 1 lần revert hết
- [ ] Ghi chú:

### Tool 7/7 — Tile Layout
Chain: self-contained (2450 loc) → `TileLayout.xaml`

- [ ] Layout tile trên floor nhỏ → kết quả đúng kích thước/góc
- [ ] Đổi origin/rotation → cập nhật đúng
- [ ] ⚠️ Script không gọi Transaction trực tiếp — xác định wrapper nào ghi model (revit.Transaction? helper?) và Ctrl+Z hoạt động đúng
- [ ] Ghi chú:

### Chốt ngày 6
- [ ] Cập nhật bảng trạng thái + README · Commit nếu có fix

---

## Ngày 7 — Nhóm Adjust / Family / Datum (7 tool)

### Tool 8/14 — PropertyLine
Chain: launcher → `PropertyLineDialog.py` (1727 loc) → `PropertyLine.xaml`

- [ ] Nhập bảng toạ độ hợp lệ → property line tạo đúng
- [ ] Sai đơn vị / toạ độ không khép kín → validate rõ ràng
- [ ] Ghi chú:

### Tool 9/14 — FamiGen
Chain: launcher → `FamiGenDialog.py` (2360 loc) → `FamiGen.xaml` · tạo family document riêng

- [ ] Gen 1 family đơn giản → load vào project OK
- [ ] Huỷ giữa chừng → family document tạm được **đóng sạch** (kiểm tra không còn doc mở ngầm)
- [ ] Template family thiếu → thông báo
- [ ] Ghi chú:

### Tool 10/14 — ManaFami
Chain: launcher → `ManaFamiDialog.py` (1428 loc) → `ManaFami.xaml`

- [ ] Batch rename family → đúng, Ctrl+Z OK
- [ ] Family đang in-use / editable=false (workshared) → xử lý đẹp
- [ ] Purge family không dùng (nếu có chức năng) → xác nhận trước khi xoá
- [ ] Ghi chú:

### Tool 11/14 — Wall_Adjust Base
Chain: self-contained (311 loc, **4 bare except**), không XAML riêng

- [ ] Tạm thêm traceback log vào 4 nhánh except
- [ ] Wall thường → adjust base đúng offset
- [ ] Wall bị attach base/top → hành vi đúng hoặc thông báo
- [ ] Ghi chú:

### Tool 12/14 — SplitElements
Chain: launcher → `SplitElementsDialog.py` (61 loc) → `SplitElements.xaml` (logic có thể nằm ở Snippets)

- [ ] Split wall theo level → đúng số đoạn, join/param giữ hợp lý
- [ ] Element không hỗ trợ split → thông báo
- [ ] Ghi chú:

### Tool 13/14 — AutoJoin
Chain: self-contained (708 loc, 1 transaction) → `AutoJoin.xaml`

- [ ] Model nhỏ nhiều wall-floor giao nhau → join hàng loạt đúng
- [ ] Cặp element không join được (geometry fail) → skip + báo số lượng, không dừng cả batch
- [ ] Đo thời gian trên model vừa (~500 element) — có progress/không treo UI
- [ ] Ghi chú:

### Tool 14/14 — Wall Cut Profile ⚠️ rủi ro cao nhất panel
Chain: self-contained (1624 loc, **37 bare except**, 8 transaction + 2 TransactionGroup)

- [ ] Trước test: chèn traceback log vào các nhánh except quan trọng (đặc biệt quanh L705, L1008, L1016 — vùng commit/rollback)
- [ ] Wall thẳng + opening đơn giản → profile cut đúng
- [ ] Wall cong → hoạt động hoặc thông báo không hỗ trợ
- [ ] Profile đè lên nhau → không tạo geometry hỏng
- [ ] Ctrl+Z 1 lần revert cả group (TransactionGroup Assimilate)
- [ ] Test cả flow "Place Opening Families" (L1212+): symbol chưa active → tự activate
- [ ] Ghi chú:

### Chốt ngày 7
- [ ] Cập nhật bảng + README · Gỡ các traceback log tạm trước khi commit
- [ ] Commit: `fix(panel-modeling): <mô tả>`

## Bảng trạng thái Panel 2

| # | Tool | Ngày | Trạng thái | Ghi chú |
|---|------|------|-----------|---------|
| 1 | CADToElements | 6 | ⬜ | |
| 2 | DoorThreshold | 6 | ⬜ | |
| 3 | ImageToDrafting | 6 | ⬜ | |
| 4 | Text to Element | 6 | ⬜ | |
| 5 | PointCloud | 6 | ⬜ | |
| 6 | RoomToFloor | 6 | ⬜ | |
| 7 | Tile Layout | 6 | ⬜ | |
| 8 | PropertyLine | 7 | ⬜ | |
| 9 | FamiGen | 7 | ⬜ | |
| 10 | ManaFami | 7 | ⬜ | |
| 11 | Wall_Adjust Base | 7 | ⬜ | |
| 12 | SplitElements | 7 | ⬜ | |
| 13 | AutoJoin | 7 | ⬜ | |
| 14 | Wall Cut Profile | 7 | ⬜ | |

## Phát sinh

_(ghi lỗi mới tại đây)_
