# Panel 4 — Data & IFC-SG (Ngày 9)

> 6 tool. Chuẩn bị: file BCF 2.1 mẫu (kèm 1 file BCF hỏng để test), model có foundation giao nhau, room + element, schedule, shared parameter file .txt.

## Tool 1/6 — IFC-SG
Chain: launcher → `IFCSGDialog.py` (2236 loc) → `IFCSG.xaml` + `SubtypeDefinerColMap.xaml` · config: `IFC-SG.pushbutton/configs/`

- [ ] Mở suite, load config mặc định từ `configs/` → không lỗi parse JSON
- [ ] Apply mapping param IFC-SG lên model mẫu → param ghi đúng element
- [ ] Config thiếu field / JSON hỏng → thông báo rõ file nào lỗi
- [ ] Mở SubtypeDefiner column mapping dialog → hoạt động
- [ ] Ctrl+Z revert việc ghi param
- [ ] Ghi chú:

## Tool 2/6 — Foundation Volume
Chain: self-contained (410 loc, 1 transaction) → `FoundationVolume.xaml`

- [ ] Foundation rời nhau → volume đúng (đối chiếu tay 1 khối đơn giản)
- [ ] Foundation giao nhau → boolean không tính trùng phần giao
- [ ] Element không có solid geometry → skip, không crash
- [ ] Ghi chú:

## Tool 3/6 — ManaContains
Chain: launcher → `ManaContainsDialog.py` (2204 loc) → `ManaContains.xaml`

- [ ] Element nằm trong room → nhận diện đúng
- [ ] Element ngoài room / trên biên → hành vi nhất quán
- [ ] Element trong **link model** → hoạt động hoặc thông báo giới hạn
- [ ] Model lớn → đo thời gian, có progress không treo
- [x] Ghi chú: **2026-07-03 (theo screenshot user)** — pyRevit output console bị flood bởi `print("[ManaContains] Skipped spatial element {}: {}"...)` lặp lại liên tục (mỗi room/space bị lỗi in ra 1 dòng). Đã xoá cả 3 chỗ `print()` trong file (2 chỗ "Skipped spatial element" ở `_t1_load_sp`/`_t2_load_spatial`, 1 chỗ "Failed to load" ở `_safe_call`) — except block vẫn skip phần tử lỗi như cũ, chỉ không in ra console nữa. **⚠️ Chưa điều tra nguyên nhân gốc**: log cho thấy rất nhiều spatial element liên tiếp bị skip cùng lỗi "Index was outside the bounds of the array" khi khởi tạo `Tab1SpatialItem`/`Tab2SpatialData` — nghi ngờ đây là bug thật (có thể liên quan Revit 2026 API) khiến danh sách room/space trong tool bị thiếu dữ liệu đáng kể, không chỉ là spam console. Cần điều tra riêng ở GĐ3 khi test functional với model thật.

## Tool 4/6 — ManaSched
Chain: launcher → `ManaSchedDialog.py` (1601 loc) → `ManaSched.xaml`

- [ ] List schedule đúng
- [ ] Sửa field/filter/sort qua tool → schedule cập nhật, Ctrl+Z OK
- [ ] Key schedule + schedule thường → cả 2 loại xử lý đúng
- [ ] Schedule đang mở làm active view → không conflict
- [ ] Ghi chú:

## Tool 5/6 — ManaPara
Chain: launcher → `ManaParaDialog.py` (2548 loc) → `ManaPara.xaml`

- [ ] Add shared param từ file .txt → binding đúng category
- [ ] Remove param → confirm trước khi xoá
- [ ] File shared param bị lock/không tồn tại → thông báo rõ
- [ ] Param ghi vào element instance vs type → đúng target
- [ ] Ghi chú:

## Tool 6/6 — BCF Reader
Chain: self-contained (3041 loc, 4 transaction) → `BCFReader.xaml`

- [ ] Mở file BCF 2.1 mẫu → list topic/comment/viewpoint đúng, hiển thị unicode tiếng Việt OK
- [ ] Apply viewpoint → camera view Revit khớp hướng nhìn BCF
- [ ] BCF hỏng (zip lỗi / XML thiếu) → thông báo, không crash
- [ ] Element highlight theo GUID trong BCF → đúng element (test cả GUID không tồn tại trong model)
- [ ] Ghi chú:

## Chốt ngày 9
- [ ] Cập nhật bảng + README · Commit: `fix(panel-data-ifcsg): <mô tả>`

| Tool | Trạng thái | Ghi chú |
|------|-----------|---------|
| IFC-SG | ⬜ | |
| Foundation Volume | ⬜ | |
| ManaContains | ⬜ | |
| ManaSched | ⬜ | |
| ManaPara | ⬜ | |
| BCF Reader | ⬜ | |

## Phát sinh

_(ghi lỗi mới tại đây)_
