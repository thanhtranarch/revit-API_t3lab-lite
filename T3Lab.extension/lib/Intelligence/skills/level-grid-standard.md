---
name: level-grid-standard
description: Level and grid creation rules - naming, elevations, axes
triggers: tao level, them level, tao grid, luoi truc, cao do, them tang, tang moi, truc moi, create level, create grid, new level, add grid, grid line, cao do tang, danh so truc
agents: revit_action, revit_data, modeling
tools: list_levels, create_level, create_grid, get_elements_by_level
---
# Quy ước Level & Grid (T3Lab)

## Level (cao độ)
- Tên: `T01`, `T02`... cho tầng; `T01M` cho lửng; `MAI`, `HAM 1` cho đặc biệt. Thống nhất theo tên có sẵn của dự án — gọi `list_levels` xem trước.
- Cao độ nhập theo mét, đúng cốt kiến trúc (FFL). Khi tạo hàng loạt tầng điển hình: nêu bảng tên + cao độ để người dùng duyệt trước khi tạo.
- KHÔNG tạo level trùng cao độ level đã có; chênh < 50mm coi như trùng — cảnh báo thay vì tạo.

## Grid (lưới trục)
- Trục số `1, 2, 3...` theo một phương; trục chữ `A, B, C...` phương còn lại (bỏ I, O dễ nhầm 1, 0).
- Grid tạo bằng `create_grid` phải chạy phủ hết phạm vi công trình; khoảng cách trục nêu rõ trong kế hoạch trước khi tạo.
- Sau khi tạo xong: liệt kê lại toàn bộ grid/level mới để người dùng kiểm tra, nhắc pin + đưa vào workset `01_GRID_LEVEL` (nếu model worksharing).

## Kiểm tra nhanh
- `get_elements_by_level` để xem tầng nào đang trống/bất thường khi người dùng nghi ngờ model gán sai tầng.
