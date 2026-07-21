---
name: image-to-model
description: Phân tích ảnh/bản vẽ công trình rồi dựng model Revit theo quy trình levels → grids → walls → floors → openings
triggers: dung model, dung cong trinh, dung nha, dung lai, build model, build from image, image to model, pdf to model, model tu anh, model tu ban ve, phan tich anh, phan tich ban ve, tao model, dung toa nha, cong trinh nha, phong ngu, can ho, phong khach
agents: revit_action, modeling, general
tools: list_levels, create_level, create_grid, place_wall, create_line_based_element, create_point_based_element, create_surface_based_element, create_room, get_available_family_types, analyze_model_statistics, select_elements
---
# Dựng model từ ảnh / bản vẽ (T3Lab)

## Giai đoạn A — Phân tích (khi có ảnh/PDF đính kèm hoặc đã phân tích ở lượt trước)
Trích xuất thành BẢNG có cấu trúc, mỗi con số ghi rõ **ĐO ĐƯỢC** (có dim/tỉ lệ trên bản vẽ) hay **ƯỚC LƯỢNG**:
- Levels: số tầng + cao độ từng tầng (m). Không có dữ liệu → đề xuất mặc định tầng cao 3.3 m.
- Lưới kết cấu: số nhịp + khoảng cách (m) nếu nhìn thấy.
- Footprint: rộng × sâu tổng thể (m), hình dạng.
- Tường/cửa/vật liệu theo từng mặt đứng.
Trình bảng cho user **xác nhận/sửa trước khi dựng**. Nếu không đọc được nội dung ảnh (model không có vision) → nói thẳng, xin kích thước chính hoặc PDF có text — không được bịa.

## Giai đoạn A' — Dựng từ mô tả text (không có ảnh/bản vẽ)
User chỉ mô tả chương trình phòng ("nhà 2 phòng ngủ, 1 khách, 1 bếp, 1 wc"):
1. Đề xuất bảng layout mặc định để user duyệt: diện tích từng phòng (ngủ ~12 m², khách ~20 m², bếp ~9 m², wc ~4 m²), footprint chữ nhật gọn, 1 tầng cao 3.3 m — kèm sơ đồ bố trí bằng chữ (phòng nào cạnh phòng nào).
2. User xác nhận/sửa xong → sang Giai đoạn B, dựng tường bao + tường ngăn đúng theo layout đã duyệt.

## Giai đoạn B — Dựng (sau khi user xác nhận số liệu)
Đơn vị mọi tool là MÉT. Dựng theo đúng thứ tự, mỗi bước báo số element + id đã tạo:
1. `list_levels` xem level sẵn có → `create_level` cho từng tầng còn thiếu (elevation = cao độ m, name = "L1", "L2"...).
2. `create_grid` cho từng trục (line start/end theo footprint).
3. Tường bao + tường ngăn từng tầng: `place_wall` (start, end, height = tầng cao, level). Batch theo tầng, tầng nào xong báo tầng đó.
4. Sàn/mái: `create_surface_based_element` với boundary points theo footprint (floor cho mỗi tầng, roof trên cùng).
5. Cửa đi/cửa sổ/cột: `get_available_family_types` để biết family có sẵn → `create_point_based_element` theo vị trí trên tường/lưới. Family thiếu → liệt kê cho user load thêm, không đoán tên family.
6. Kiểm tra chéo: `analyze_model_statistics` so số element từng category với bảng phân tích; lệch thì báo rõ.

## Quy tắc
- HÀNH ĐỘNG NGAY bằng tool call trong cùng lượt trả lời. Viết "Tôi sẽ tạo..." mà không gọi tool nào là SAI — mỗi lượt phải có ít nhất một tool call cho tới khi dựng xong (hoặc một câu hỏi cụ thể nếu thiếu số liệu).
- KHÔNG dùng `create_text_note` để "dựng" — text note là ghi chú, không phải hình khối. Nhà ở dựng bằng level/wall/floor/roof; `create_room` chỉ đặt tên phòng SAU khi đã có tường bao quanh.
- Nếu lượt trước ĐÃ có bảng phân tích trong hội thoại: dùng lại, không phân tích lại, không hỏi lại từ đầu.
- Thiếu kích thước nào → hỏi đúng kích thước đó (một câu gọn); user không biết → đề xuất giá trị mặc định hợp lý và chờ đồng ý.
- "Dựng model" là DỰNG BẰNG TOOL trong Revit — không phải mở tool window nào khác; không gọi `open_t3lab_tool` trong workflow này.
- Model phức tạp (cong, tham số hóa): dựng phần khối chính trước, phần chi tiết liệt kê để user quyết định làm tay hay dùng tool chuyên (Point Cloud to Model, CAD to Elements).
