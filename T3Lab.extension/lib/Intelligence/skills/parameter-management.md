---
name: parameter-management
description: Quan ly parameter - doc, gan hang loat, tao project parameter dung chuan
triggers: parameter, tham so, gan parameter, dien parameter, sua parameter, project parameter, shared parameter
agents: revit_action, revit_data
tools: get_all_parameters, get_parameter, set_parameter, bulk_set_parameter, create_project_parameter, ai_element_filter
---
# Quản lý Parameter (T3Lab)

## Đọc trước khi sửa
1. `get_all_parameters` / `get_parameter` xem giá trị hiện tại và parameter là instance hay type, read-only hay không.
2. Parameter type ảnh hưởng MỌI instance cùng type — luôn nêu rõ điều này trước khi sửa.

## Gán hàng loạt
- Gom đối tượng bằng `ai_element_filter` theo điều kiện người dùng mô tả, báo SỐ LƯỢNG element sẽ bị sửa + giá trị cũ→mới, rồi mới `bulk_set_parameter`.
- Giá trị nhập thống nhất: không dấu cách thừa, đơn vị theo project (mm cho chiều dài).
- Sau khi gán: kiểm tra xác suất 3–5 element bằng `get_parameter` và báo kết quả.

## Tạo project parameter
- `create_project_parameter` khi dự án cần trường dữ liệu mới (mã phòng, mã cấu kiện, giai đoạn...).
- Đặt tên: `T3_<NhomChucNang>_<TenThamSo>`, ví dụ `T3_QLDA_MaCauKien`. Nêu rõ: categories áp dụng, kiểu dữ liệu, instance/type — chờ xác nhận rồi mới tạo.
- Không tạo trùng tên parameter có sẵn — kiểm tra bằng `get_all_parameters` trước.

## Báo cáo
- Bảng: | Parameter | Phạm vi | Giá trị | Số element | Trạng thái |.
