---
name: sheet-naming-standard
description: T3Lab sheet naming and numbering rules for creating or renaming sheets
triggers: ten sheet, so sheet, sheet number, doi ten sheet, tao sheet, dat ten sheet, sheet name, rename sheet, danh so sheet, sheet moi
agents: revit_action, general
tools: revit_list_sheets, rename_element, create_sheet
---
# Quy tắc đặt tên & đánh số sheet (T3Lab)

Khi tạo mới hoặc đổi tên/đổi số sheet, áp dụng đúng các quy tắc sau:

## Số sheet (Sheet Number)
- Định dạng: `<Khối>-<Số 3 chữ số>`, ví dụ `A-101`, `S-201`, `M-301`.
- Tiền tố khối: `A` (Kiến trúc), `S` (Kết cấu), `M` (Cơ), `E` (Điện), `P` (Nước), `G` (Chung).
- Số thứ tự tăng dần trong cùng nhóm; nhóm theo trăm: 0xx tổng thể, 1xx mặt bằng, 2xx mặt đứng/cắt, 3xx chi tiết, 4xx schedule.
- Không dùng khoảng trắng hoặc ký tự đặc biệt ngoài dấu gạch nối.

## Tên sheet (Sheet Name)
- Viết HOA toàn bộ, tiếng Việt không dấu hoặc tiếng Anh — thống nhất theo dự án hiện tại (nhìn các sheet có sẵn để chọn).
- Cấu trúc: `<LOẠI BẢN VẼ> - <PHẠM VI>`, ví dụ `MAT BANG TANG 1`, `FLOOR PLAN - LEVEL 2`.
- Không lặp lại số sheet trong tên.

## Quy trình bắt buộc
1. Gọi `revit_list_sheets` trước để xem quy ước hiện hữu của dự án và tránh trùng số.
2. Nếu số sheet đề xuất đã tồn tại, tăng số thứ tự thay vì ghi đè.
3. Sau khi tạo/đổi tên, báo lại danh sách sheet đã thay đổi (số cũ → số mới).
