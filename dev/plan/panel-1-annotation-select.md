# Panel 1 — Annotation & Select (Ngày 7)

> 4 tool. Model test: file nhỏ có wall, grid, dimension, text note, tag, 1 DWG import + 1 DWG link.
> Quy trình chung: xem "Định nghĩa xong" trong `dev/plan/README.md`. Bật debug: giữ CTRL khi bấm nút.

## Tool 1/4 — AutoDimension

Chain: `script.py` (24 loc) → `lib/GUI/AutoDimensionDialog.py` (1721 loc) → `AutoDimension.xaml`

- [ ] Mở cửa sổ, chrome OK
- [ ] Auto dim wall trên floor plan → dim tạo đúng, Ctrl+Z revert 1 bước
- [ ] Auto dim grid trên plan + trên section
- [ ] View không hỗ trợ (3D, schedule) → thông báo thân thiện
- [ ] Không chọn gì / model không có element khớp → không stacktrace
- [ ] Ghi chú:

## Tool 2/4 — ManaDWG

Chain: `script.py` (432 loc, self-contained) → `DWGManagement.xaml` (**UI-LOCKED** — chỉ debug logic, không đụng XAML)

- [ ] Mở cửa sổ, list đúng cả DWG **import** lẫn DWG **link**, phân biệt rõ
- [ ] Xoá 1 DWG import → biến mất khỏi model + khỏi list; Ctrl+Z khôi phục
- [ ] Thao tác trên DWG link (unload/remove tuỳ chức năng)
- [ ] Model không có DWG → list rỗng, không lỗi
- [ ] Kiểm tra 2 transaction trong script gom đúng thao tác
- [ ] Ghi chú:

## Tool 3/4 — ManaSelect

Chain: `script.py` (29 loc) → `lib/GUI/ManaSelectDialog.py` (660 loc) → `ManaSelect.xaml` · dialog phụ: `QuickElementDialog.py` → `QuickElement.xaml`

- [ ] Mở cửa sổ, chrome OK
- [ ] Pre-select hỗn hợp element → filter theo category/type hoạt động
- [ ] Selection rỗng khi mở tool → xử lý đẹp
- [ ] Mở được QuickElement dialog phụ từ trong tool (nếu có nút) → không XamlParseException
- [ ] Apply selection → selection trong Revit đổi đúng
- [ ] Ghi chú:

## Tool 4/4 — ManaAnno

Chain: `script.py` (36 loc) → `lib/GUI/ManaAnnoDialog.py` (1416 loc) → `ManaAnno.xaml` · dialog phụ: `DimTextDialog.py` → `DimText.xaml`, `TagCheckerDialog.py` → `TagChecker.xaml`

- [ ] Mở cửa sổ, chrome OK
- [ ] Batch đổi type text/tag/dim trên view test → đúng, Ctrl+Z OK
- [ ] Mở DimText dialog phụ → hoạt động
- [ ] Mở TagChecker dialog phụ → hoạt động
- [ ] View có view template khoá annotation → thông báo hợp lý, không crash
- [ ] Ghi chú:

## Chốt ngày 7
- [ ] Cập nhật trạng thái 4 tool vào bảng dưới + `dev/plan/README.md`
- [ ] Lỗi tìm thấy → mục Phát sinh + quyết định sửa ngay (nhỏ) hay tạo việc riêng (lớn)
- [ ] Commit (nếu có sửa code): `fix(panel-annotation): <mô tả>`

| Tool | Trạng thái | Ghi chú |
|------|-----------|---------|
| AutoDimension | ⬜ | |
| ManaDWG | ⬜ | |
| ManaSelect | ⬜ | |
| ManaAnno | ⬜ | |

## Phát sinh

_(ghi lỗi mới tại đây)_
