---
name: warning-triage
description: Classify and fix model warnings by priority
triggers: sua warning, xu ly warning, fix warning, giam warning, loi model, warning nhieu
agents: revit_action, revit_data
tools: get_model_warnings, select_elements, color_elements, join_geometry, delete_element, revit_override_color
---
# Xử lý Warnings (T3Lab)

## Bước 1 — Phân loại (luôn làm trước)
`get_model_warnings` rồi nhóm theo mức ưu tiên:
- **P1 nghiêm trọng**: "identical instances in the same place" (khối lượng đôi), room/area boundary lỗi, elements bị mất liên kết host.
- **P2 nên xử lý**: "highlighted elements are joined but do not intersect", wall/floor overlap, duplicate mark.
- **P3 chấp nhận được**: slightly off axis, tag ngoài phạm vi.

## Bước 2 — Đề xuất phương án theo nhóm
- Duplicate instances → liệt kê cặp element id, đề xuất xóa bản thừa (`delete_element`) — CHỜ XÁC NHẬN trước khi xóa.
- Join lỗi → `join_geometry` lại đúng cặp, hoặc unjoin nếu không cần.
- Cần soi trực quan → `select_elements` / `color_elements` tô nhóm element lỗi để người dùng nhìn thấy trong view.

## Bước 3 — Thực hiện & báo cáo
- Xử lý theo LÔ TỪNG NHÓM, không trộn nhiều loại warning trong một lượt.
- Sau mỗi lô: chạy lại `get_model_warnings`, báo số warning trước → sau.
- Kết thúc: bảng | Nhóm warning | Trước | Sau | Còn lại vì sao |.

## Nguyên tắc
- KHÔNG xóa element khi chưa được đồng ý. Warning không rõ nguyên nhân → báo cáo thay vì đoán.
