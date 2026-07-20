---
name: sheet-set-management
description: Batch-create sheet sets and place views on sheets
triggers: tao bo sheet, dat view len sheet, len sheet, place view, bo ban ve, xep sheet, dan trang
agents: revit_action, revit_data
tools: create_sheet, add_view_to_sheet, place_views_on_sheets, revit_list_sheets, revit_list_views, duplicate_view
---
# Quản lý bộ Sheet (T3Lab)

## Lập kế hoạch trước khi tạo
1. `revit_list_sheets` + `revit_list_views`: nắm sheet đã có và view sẵn sàng lên sheet.
2. Trình bảng kế hoạch cho người dùng duyệt TRƯỚC khi tạo hàng loạt:
   | Số sheet | Tên sheet | View sẽ đặt | Title block |
3. Số/tên sheet tuân theo skill `sheet-naming-standard` (khối-số 3 chữ số, nhóm theo trăm).

## Tạo & đặt view
- Tạo sheet: `create_sheet` theo bảng đã duyệt.
- Đặt 1 view: `add_view_to_sheet`; đặt loạt view cùng loại lên loạt sheet: `place_views_on_sheets`.
- View đã nằm trên sheet khác thì KHÔNG đặt lại được — cần bản sao thì `duplicate_view` (as dependent cho mặt bằng chia khu).

## Chuẩn trình bày
- Mỗi sheet một mặt bằng chính; chi tiết phụ xếp lưới từ trái→phải, trên→dưới.
- Tỷ lệ view phải chốt TRƯỚC khi đặt lên sheet; đổi tỷ lệ sau khi dàn trang sẽ vỡ bố cục.
- View trên sheet bắt buộc có template (skill `view-template-standard`).

## Sau khi hoàn thành
- `revit_list_sheets` xác nhận lại, báo: số sheet tạo mới, số view đã đặt, view nào KHÔNG đặt được (kèm lý do — thường là đã thuộc sheet khác).
