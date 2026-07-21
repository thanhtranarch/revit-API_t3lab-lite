---
name: model-cleanup
description: Clean up the model - purge, file size, junk views and families
triggers: don dep model, don model, purge, lam sach model, giam dung luong, file nang, model nang, purge unused, xoa view rac, view rac, don file, model cham, file cham
agents: revit_action, revit_data, qa_check
tools: purge_unused, analyze_model_statistics, get_model_health, revit_list_views, audit_model, delete_element
---
# Dọn dẹp Model (T3Lab)

## Chẩn đoán trước (bắt buộc)
1. `get_model_health` + `analyze_model_statistics`: file size, số family, số view, số element, in-place family.
2. `revit_list_views`: đếm view không đặt trên sheet (view rác là nguyên nhân nặng file phổ biến nhất).
3. Báo bảng hiện trạng trước khi đụng vào model.

## Thứ tự dọn (mỗi bước chờ xác nhận)
1. **Purge unused** — `purge_unused` (family/type/material không dùng). Nêu rõ: thao tác không undo chọn lọc được, nên lưu file trước.
2. **View rác** — liệt kê view không trên sheet, không template, không prefix `WIP_` → người dùng duyệt danh sách rồi mới xóa từng lô bằng `delete_element`.
3. **Link & import** — liệt kê DWG import (không phải link) tồn đọng; khuyến nghị xóa import nội thất/tạm.
4. **Audit** — `audit_model` chốt lại sau khi dọn.

## Ngưỡng khuyến nghị
- File > 300MB: cảnh báo đỏ; 150–300MB: nên dọn; family in-place > 10: cần thay bằng family chuẩn.

## Báo cáo kết thúc
- | Hạng mục | Trước | Sau | Tiết kiệm |, kèm khuyến nghị duy trì (chu kỳ purge 2 tuần/lần).
