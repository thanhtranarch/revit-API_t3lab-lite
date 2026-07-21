---
name: bim-handover-checklist
description: End-of-stage / as-built model handover checklist
triggers: ban giao model, as built, as-built, hoan cong, handover model, ban giao bim, dong ho so
agents: revit_data, revit_action, general
tools: audit_model, get_model_health, get_model_warnings, purge_unused, revit_list_sheets, revit_list_views, export_sheets_pdf
---
# Checklist bàn giao model (T3Lab)

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
