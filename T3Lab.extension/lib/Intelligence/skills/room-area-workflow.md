---
name: room-area-workflow
description: Quy trinh tao room, tag room, xuat dien tich va tao san tu room
triggers: tao room, tag room, dien tich phong, bang dien tich, room to floor, tao san tu room, xuat room
agents: revit_action, revit_data
tools: create_room, tag_all_rooms, export_room_data, store_room_data, room_to_floor, create_schedule, list_levels
---
# Quy trình Room & Diện tích (T3Lab)

## Tạo & đặt tên room
- Tên room tiếng Việt viết hoa chữ đầu, thống nhất toàn dự án: `Phòng ngủ 1`, `WC`, `Bếp + Ăn`, `Lô gia`.
- Tạo room theo tầng (`create_room` với level rõ ràng); phòng chưa rõ chức năng đặt `Phòng` + số, KHÔNG để "Room".
- Sau khi tạo: `tag_all_rooms` cho các view mặt bằng liên quan.

## Bảng diện tích
- Dùng `create_schedule` (category Rooms) với cột: Number, Name, Level, Area.
- Xuất số liệu: `export_room_data`; cần lưu để so sánh giữa các phương án thì `store_room_data` trước khi thay đổi lớn.
- Khi báo diện tích: làm tròn 1 chữ số thập phân, kèm tổng theo tầng và tổng toàn dự án.

## Room → Floor
- `room_to_floor` chỉ chạy sau khi room đã khép kín (không lỗi boundary). Nêu rõ floor type sẽ dùng và số room ảnh hưởng trước khi tạo.
- Sàn tạo ra cần kiểm tra lại cao độ + offset so với cốt hoàn thiện.

## Cảnh báo thường gặp
- Room "not enclosed" / "redundant" → liệt kê phòng lỗi kèm tầng để người dùng sửa boundary, không bỏ qua im lặng.
