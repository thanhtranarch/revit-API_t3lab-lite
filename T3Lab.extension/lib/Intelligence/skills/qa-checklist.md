---
name: qa-checklist
description: Revit model QA checklist before handover or periodic audits
triggers: kiem tra model, qa model, audit model, model health, kiem tra chat luong, danh gia model, model checker, check model, review model, soat model
agents: revit_data, qa_check, general
tools: get_model_warnings, get_model_health, analyze_model_statistics, list_worksets
---
# Checklist QA model Revit (T3Lab)

Khi được yêu cầu kiểm tra/đánh giá model, thực hiện tuần tự và tổng hợp thành MỘT báo cáo:

## 1. Cảnh báo (Warnings)
- Gọi `get_model_warnings`; phân loại theo mức độ: nghiêm trọng (duplicate instances, room boundary, join geometry) vs nhẹ.
- Ngưỡng: > 200 warnings = cần xử lý ngay; 50–200 = lên kế hoạch; < 50 = chấp nhận được.

## 2. Sức khỏe tổng thể
- Gọi `get_model_health` và `analyze_model_statistics`.
- Soi các chỉ số: số element, số family in-place (nên < 10), file size cảnh báo khi > 300MB, view không đặt trên sheet.

## 3. Tổ chức dữ liệu
- `list_worksets`: kiểm tra element nằm đúng workset, không dồn hết vào Workset1.
- Naming: sheet/view/family theo chuẩn dự án.

## Định dạng báo cáo
- Bảng tóm tắt: | Hạng mục | Kết quả | Mức độ | Khuyến nghị |
- Kết luận 1 dòng: ĐẠT / CẦN XỬ LÝ / NGHIÊM TRỌNG.
- KHÔNG tự ý sửa model trong lượt kiểm tra — chỉ đề xuất; chờ người dùng yêu cầu mới thực hiện.
