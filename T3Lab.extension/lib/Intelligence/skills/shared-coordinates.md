---
name: shared-coordinates
description: Model positioning - shared coordinates, base and survey points
triggers: shared coordinates, toa do chung, dinh vi model, survey point, base point, origin model, lech toa do, model bi lech
agents: knowledge, revit_data, general
tools: get_revit_context, revit_get_project_info, list_open_documents
---
# Định vị model & Shared Coordinates (T3Lab)

## Ba gốc tọa độ trong Revit — phân biệt rõ
- **Internal Origin**: gốc nội bộ cố định của file, không di chuyển được. Model nên nằm trong bán kính ~16km quanh gốc này (xa hơn → lỗi đồ họa, hình học kém chính xác).
- **Project Base Point**: gốc trục hồ sơ của công trình (thường đặt tại giao 2 trục chính, cốt ±0.000).
- **Survey Point**: điểm mốc trắc đạc thực tế (tọa độ quốc gia VN2000 hoặc mốc dự án).

## Quy tắc khi dự án nhiều model link
1. MỘT model chủ (thường Kiến trúc hoặc model định vị riêng) giữ shared coordinates gốc; các model khác **Acquire Coordinates** từ model chủ — KHÔNG ai tự dịch model bằng tay.
2. Link Revit luôn bằng **Auto - By Shared Coordinates**; nếu Revit báo chưa có shared coordinates thì dừng lại xử lý định vị trước, không "Center to Center rồi kéo".
3. Không bao giờ di chuyển model đã phối hợp — mọi thay đổi định vị phải thông báo tất cả bộ môn và ghi vào BEP.
4. Góc **True North** vs **Project North**: hồ sơ vẽ theo Project North; xoay True North chỉ ở thiết lập tọa độ, không xoay model.

## Chẩn đoán "model bị lệch"
1. `get_revit_context` / `revit_get_project_info` xem thông tin định vị dự án hiện tại; `list_open_documents` xem các model liên quan.
2. Hỏi người dùng: lệch giữa link nào với link nào, lệch phương ngang hay cao độ.
3. Nguyên nhân phổ biến: link kiểu Center-to-Center, model chủ đổi shared coordinates mà không báo, hoặc 2 file dùng 2 mốc survey khác nhau → chỉ ra cách xác minh từng trường hợp, KHÔNG đề xuất kéo model bằng tay.
