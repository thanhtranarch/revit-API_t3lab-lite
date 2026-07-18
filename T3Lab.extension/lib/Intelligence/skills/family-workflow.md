---
name: family-workflow
description: Quy trinh load family, chon type va dat family vao model
triggers: load family, tai family, dat family, chen family, family type, chon family, place family
agents: revit_action, revit_data
tools: load_family, get_available_family_types, create_point_based_element, create_line_based_element, list_levels
---
# Quy trình Family (T3Lab)

## Trước khi load
1. `get_available_family_types` kiểm tra family/type đã có trong model chưa — tránh load trùng tạo bản `family 1`.
2. Family mới phải rõ nguồn (thư viện T3Lab / người dùng cung cấp đường dẫn). Không load family lạ không rõ xuất xứ.

## Load & kiểm tra
- `load_family` từ đường dẫn người dùng chỉ định; nếu model đã có family cùng tên, hỏi trước: ghi đè type hay giữ bản cũ.
- Sau khi load: liệt kê các type có trong family để người dùng chọn đúng type cần dùng.

## Đặt vào model
- Element theo điểm (bàn ghế, thiết bị, cột...): `create_point_based_element` — nêu rõ level (`list_levels` nếu chưa chắc), tọa độ/phòng đặt và số lượng.
- Element theo đường (tường, dầm, ống...): `create_line_based_element` với điểm đầu–cuối rõ ràng.
- Đặt hàng loạt (> 10 instance): trình kế hoạch số lượng + vị trí trước, chờ xác nhận.

## Chuẩn tên family T3Lab
- `T3_<Category>_<TênMôTả>`, type đặt theo kích thước thật: `600x600`, `D200`.
- Family sai chuẩn tên → đề xuất đổi tên ngay lúc load để thư viện dự án sạch.
