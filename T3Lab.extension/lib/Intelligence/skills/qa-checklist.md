---
name: qa-checklist
description: Revit model QA checklist before handover or periodic audits
triggers: kiem tra model, qa model, audit model, model health, kiem tra chat luong, danh gia model, model checker, check model, review model, soat model
agents: revit_data, qa_check, general
tools: get_model_warnings, get_model_health, analyze_model_statistics, list_worksets
standard: project
---
# Checklist QA model Revit (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

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
