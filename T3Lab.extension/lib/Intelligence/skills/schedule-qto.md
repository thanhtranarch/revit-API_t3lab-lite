---
name: schedule-qto
description: Schedules and quantity takeoff straight from the Revit model
triggers: schedule, boc khoi luong, khoi luong, thong ke, bang thong ke, quantity, takeoff, boc tach, material quantities, khoi luong vat lieu, khoi luong be tong, qto, bill of quantities, boq
agents: revit_data, revit_action, general
tools: create_schedule, get_schedule_data, get_material_quantities, analyze_model_statistics, export_room_data
---
# Schedule & Bóc khối lượng (T3Lab)

## Chọn đúng tool (QUAN TRỌNG)
- **Thống kê / bóc khối lượng theo CATEGORY** (sàn, tường, cột… — user chưa chỉ định schedule có sẵn):
  - Số lượng element → `analyze_model_statistics` (đọc `element_counts`).
  - Diện tích / thể tích / khối lượng → `get_material_quantities(category=...)`. Category hợp lệ: Walls, Floors, Roofs, Ceilings. "Sàn" → `Floors`.
- **`get_schedule_data` CHỈ đọc một ViewSchedule ĐÃ CÓ** qua `schedule_name` hoặc `schedule_id`. Tool này **KHÔNG có tham số `category`** — đừng gọi `get_schedule_data(category=...)`. Muốn đọc một category dạng bảng thì `create_schedule(category=...)` trước, rồi mới `get_schedule_data(schedule_name=...)`.

## Trình bày kết quả thống kê (BẢNG — QUAN TRỌNG)
- "Có bao nhiêu X" → trả lời một câu tổng (vd "Model có 128 tường").
- "Thống kê / thong ke / breakdown X" → **LUÔN xuất bảng markdown pipe**, nhóm theo **Loại (name) × Tầng (level)** từ `ai_element_filter(category=..., limit=1000)`:

  | Loại | Tầng | Số lượng |
  | --- | --- | --- |
  | ... | ... | ... |
  | **Tổng** |  | **<total_count>** |

  Bắt buộc có dòng phân cách `| --- | --- | --- |` (chat mới render thành bảng viền), và dòng **Tổng** cuối.
- Dòng **Tổng** lấy đúng `total_count` do tool trả — KHÔNG tự cộng danh sách.
- Chỉ nhóm từ danh sách `elements` khi `truncated=false`. Nếu `truncated=true`, gọi lại với `offset` (cùng `limit`) tới khi `offset+count = total_count` rồi mới lập bảng, hoặc nói rõ bảng còn thiếu.
- `analyze_model_statistics` chỉ có tổng theo category (không có type/level) → dùng cho câu hỏi tổng quan, KHÔNG đủ để lập bảng breakdown theo tầng.

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
