---
name: bim-handover-checklist
description: End-of-stage / as-built model handover checklist
triggers: ban giao model, as built, as-built, hoan cong, handover model, ban giao bim, dong ho so, giao nop model, chuan bi ban giao, final handover, nghiem thu model
agents: revit_data, revit_action, qa_check, general
tools: audit_model, get_model_health, get_model_warnings, purge_unused, revit_list_sheets, revit_list_views, export_sheets_pdf, delete_element
standard: project
---
# Checklist bàn giao model (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

Chạy TUẦN TỰ khi người dùng yêu cầu chuẩn bị bàn giao model / hồ sơ hoàn công. Mỗi bước báo kết quả rồi mới sang bước sau.

## 1. Sức khỏe model
- `get_model_health` + `get_model_warnings`: warnings phải về ngưỡng chấp nhận (< 50, không còn nhóm P1 — xem skill `warning-triage`).
- `audit_model` chốt tình trạng cuối.

## 2. Dọn sạch
- `purge_unused` (chờ xác nhận — nhắc lưu file trước).
- `revit_list_views`: view WIP/không trên sheet → duyệt danh sách xóa. Model bàn giao KHÔNG chứa view rác, không chứa CAD import tồn đọng.

## 3. Hồ sơ & dữ liệu
- `revit_list_sheets`: đủ bộ sheet, đúng revision phát hành cuối; số/tên theo `sheet-naming-standard`.
- Dữ liệu tài sản đã điền theo yêu cầu EIR (skill `cobie-handover`) nếu dự án yêu cầu FM data.
- Tên file giao nộp theo `iso19650-naming`; trạng thái/mã suitability ghi ở CDE.

## 4. Định vị & liên kết
- Shared coordinates còn nguyên (skill `shared-coordinates`); link RVT/CAD đúng đường dẫn tương đối hoặc gỡ theo yêu cầu bên nhận.

## 5. Xuất bàn giao
- `export_sheets_pdf` bộ hồ sơ cuối; lưu vào thư mục dự án theo ngày để có vết bàn giao.
- As-built: model phải phản ánh thực tế thi công (LOD 500) — hạng mục chưa cập nhật theo hiện trạng phải liệt kê rõ, không bàn giao "nhầm" model thiết kế thành hoàn công.

## Báo cáo tổng
| Hạng mục | Trạng thái | Ghi chú | → Kết luận: SẴN SÀNG / CÒN VIỆC (liệt kê).
