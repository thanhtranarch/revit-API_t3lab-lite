---
name: cobie-handover
description: Prepare COBie handover data - map Revit parameters to COBie fields
triggers: cobie, du lieu tai san, asset data, ban giao du lieu, facility management, fm data, du lieu van hanh, van hanh bao tri, asset register, ma tai san
agents: general, revit_data, revit_action, qa_check
tools: get_all_parameters, set_parameter, bulk_set_parameter, create_project_parameter, create_schedule, get_schedule_data, export_room_data, ai_element_filter
---
# Chuẩn bị dữ liệu COBie (T3Lab)

COBie = bộ dữ liệu có cấu trúc bàn giao cho đội vận hành (O&M), thay cho hồ sơ giấy rời rạc.

## Các bảng COBie chính lấy từ Revit
| Bảng | Nguồn trong Revit |
|------|-------------------|
| Facility | Project Information |
| Floor | Levels |
| Space | Rooms (Number, Name, Area) |
| Type | Family Type + parameter loại thiết bị |
| Component | Instance thiết bị có Mark riêng |
| System | MEP systems |

## Parameter tối thiểu cần điền cho thiết bị (Type/Component)
- `Mark` duy nhất từng instance; `Type Mark` từng type.
- Nhà sản xuất, model number, ngày lắp đặt, thời hạn bảo hành, đơn vị bảo trì.
- Thiếu trường nào → `create_project_parameter` bổ sung theo tên chuẩn EIR của dự án (hỏi người dùng EIR yêu cầu trường gì trước khi tạo).

## Quy trình kiểm tra dữ liệu
1. `get_all_parameters` trên nhóm thiết bị mẫu — soi trường trống.
2. `create_schedule` theo category thiết bị với các cột COBie → `get_schedule_data` rà ô trống, Mark trùng.
3. Điền hàng loạt giá trị chung (nhà thầu, ngày bàn giao) bằng `bulk_set_parameter` — báo số lượng trước khi ghi.
4. Space data: `export_room_data` đối chiếu số phòng ↔ danh mục Space.

## Nguyên tắc
- Mark trùng hoặc rỗng là lỗi CHẶN bàn giao — báo danh sách element id.
- Không bịa dữ liệu nhà sản xuất/bảo hành — trường chưa có thông tin để trống và ghi chú "chờ nhà thầu".
