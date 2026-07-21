---
name: clash-coordination
description: Multidiscipline coordination and clash detection workflow
triggers: clash, va cham, kiem tra va cham, phoi hop bo mon, clash detection, federated model, coordination
agents: revit_data, revit_action, general
tools: get_model_warnings, ai_element_filter, select_elements, color_elements, create_view, join_geometry, list_open_documents
---
# Phối hợp liên bộ môn & Clash (T3Lab)

## Nguyên tắc phối hợp
- Clash chỉ có ý nghĩa từ LOD 300 trở lên (hình học chính xác). Model chưa đạt → nói rõ, đừng chạy cho có.
- Trước khi xét va chạm giữa các model: kiểm tra các model link đã ĐÚNG shared coordinates chưa (lệch origin là nguồn "clash ảo" số 1).
- Chu kỳ phối hợp: chốt theo BEP (thường 1–2 tuần/lần); mỗi vòng có biên bản clash cũ đã xử lý → còn lại.

## Quy trình trong T3Lab Assistant
1. `list_open_documents` xác nhận các model bộ môn đang mở/link.
2. Trong nội bộ 1 model: `get_model_warnings` lọc nhóm overlap/join — đây là "clash nội bộ".
3. Khoanh vùng nghi vấn: `ai_element_filter` gom element 2 hệ tại khu vực người dùng chỉ định → `color_elements` tô 2 màu khác nhau, `create_view` dựng view 3D khu vực đó cho người dùng soi.
4. Va chạm cứng nhỏ do modeling → `join_geometry` xử lý; va chạm thiết kế (ống xuyên dầm, cao độ trần thiếu) → KHÔNG tự sửa, lập danh sách chuyển bộ môn liên quan.

## Báo cáo clash
- | # | Vị trí (tầng/trục) | Hệ A | Hệ B | Mức độ | Đề xuất | Bộ môn xử lý |
- Mức độ: Cứng (giao nhau vật lý) / Mềm (thiếu khoảng thao tác, bảo trì) / Ảo (do link lệch, duplicate).
- Nhắc người dùng: kiểm tra chéo bằng Navisworks/BIM Collab cho model liên kết lớn — Revit warnings không thay thế clash detection chuyên dụng.
