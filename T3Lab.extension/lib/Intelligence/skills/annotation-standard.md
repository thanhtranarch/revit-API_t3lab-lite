---
name: annotation-standard
description: Dimension, tag and text note conventions on drawings
triggers: dim, kich thuoc, ghi kich thuoc, danh tag, tag tuong, tag elements, ghi chu, text note, annotation, dimension, chuoi dim, tag room, tag phong, general notes
agents: revit_action, general
tools: create_dimension, tag_elements, tag_all_walls, tag_all_rooms, create_text_note
---
# Quy ước Dim / Tag / Ghi chú (T3Lab)

## Dim (kích thước)
- Thứ tự chuỗi dim mặt bằng từ ngoài vào: (1) tổng, (2) trục–trục, (3) mảng tường/lỗ mở chi tiết.
- Dim bám grid và mặt kết cấu, KHÔNG dim vào mặt hoàn thiện trừ bản vẽ hoàn thiện.
- `create_dimension` cho từng chuỗi; sau khi tạo báo số chuỗi dim đã thêm theo view.

## Tag
- Tag phải lấy dữ liệu từ element (không text đè tay). Tường dùng `tag_all_walls`, room dùng `tag_all_rooms`, còn lại `tag_elements` theo category.
- Tag không được đè lên nhau hoặc đè lên đối tượng; nếu dày đặc, đề xuất tag theo cụm/điển hình.
- Thiếu thông tin trong tag (type name rỗng, mark trùng) → báo danh sách element lỗi thay vì tag bừa.

## Text note
- `create_text_note` chỉ dùng cho ghi chú chung (general notes) hoặc chú thích không lấy được từ dữ liệu model. Nội dung tiếng Việt có dấu, ngắn gọn, chữ HOA cho tiêu đề.
- Không dùng text để "sửa" thông tin sai lệch với model — phải sửa model.

## Sau khi thao tác
- Tóm tắt: view nào, bao nhiêu dim/tag/note đã thêm.
