---
name: schedule-qto
description: Schedules and quantity takeoff straight from the Revit model
triggers: schedule, boc khoi luong, khoi luong, thong ke, bang thong ke, quantity, takeoff, boc tach, material quantities, khoi luong vat lieu, khoi luong be tong, qto, bill of quantities, boq
agents: revit_data, revit_action, general
tools: create_schedule, get_schedule_data, get_material_quantities, analyze_model_statistics, export_room_data
---
# Schedule & Bóc khối lượng (T3Lab)

## Nguyên tắc
- Khối lượng lấy TRỰC TIẾP từ model (`get_schedule_data`, `get_material_quantities`) — không ước lượng tay. Model thiếu gì phải nói rõ "model chưa đủ dữ liệu mục X".
- MỌI con số thống kê dùng field do tool tính sẵn: `row_count` + `column_totals` (get_schedule_data), `total_count` (ai_element_filter), `element_counts` (analyze_model_statistics). TUYỆT ĐỐI không tự đếm/cộng từ danh sách `rows`/`elements` — danh sách có thể đã bị cắt bớt trước khi tới model.
- Đơn vị mặc định: dài m, diện tích m², thể tích m³, khối lượng kg; làm tròn 2 chữ số.

## Tạo schedule
- `create_schedule` theo category người dùng cần; cột tối thiểu: Family & Type, Level, Count + trường số lượng (Length/Area/Volume tùy loại).
- Tên schedule: `QTO_<CATEGORY>_<PHẠM VI>`, ví dụ `QTO_WALL_TANG 1-5`.

## Bóc khối lượng vật liệu
1. `get_material_quantities` theo category/phạm vi yêu cầu.
2. Tổng hợp theo: Vật liệu → Category → Tầng. Xuất bảng markdown: | Vật liệu | Category | Số lượng | Đơn vị |.
3. Chênh lệch bất thường (ví dụ khối lượng 0, vật liệu "By Category") → liệt kê để người dùng kiểm tra model.

## Đối chiếu nhanh
- `analyze_model_statistics` để lấy tổng số element từng category làm số kiểm tra chéo với schedule.
- Nhắc rõ: khối lượng model là khối lượng THIẾT KẾ — chưa gồm hao hụt thi công.
