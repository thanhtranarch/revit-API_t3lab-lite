---
name: sheet-naming-standard
description: T3Lab sheet naming and numbering rules for creating or renaming sheets
triggers: ten sheet, so sheet, sheet number, doi ten sheet, tao sheet, dat ten sheet, sheet name, rename sheet, danh so sheet, sheet moi
agents: revit_action, general
tools: revit_list_sheets, rename_element, create_sheet
standard: project
---
# Quy tắc đặt tên & đánh số sheet (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

Khi tạo mới hoặc đổi tên/đổi số sheet, áp dụng đúng các quy tắc sau:

## Số sheet (Sheet Number) — MẪU THAM KHẢO
- Định dạng: `<Khối>-<Số 3 chữ số>`, ví dụ `A-101`, `S-201`, `M-301`.
- Tiền tố khối: `A` (Kiến trúc), `S` (Kết cấu), `M` (Cơ), `E` (Điện), `P` (Nước), `G` (Chung).
- Số thứ tự tăng dần trong cùng nhóm; nhóm theo trăm: 0xx tổng thể, 1xx mặt bằng, 2xx mặt đứng/cắt, 3xx chi tiết, 4xx schedule.
- Không dùng khoảng trắng hoặc ký tự đặc biệt ngoài dấu gạch nối.

## Tên sheet (Sheet Name)
- Viết HOA toàn bộ, tiếng Việt không dấu hoặc tiếng Anh — thống nhất theo dự án hiện tại (nhìn các sheet có sẵn để chọn).
- Cấu trúc: `<LOẠI BẢN VẼ> - <PHẠM VI>`, ví dụ `MAT BANG TANG 1`, `FLOOR PLAN - LEVEL 2`.
- Không lặp lại số sheet trong tên.

## Quy trình bắt buộc
1. **Tra BEP/tiêu chuẩn dự án trước**; không có thì gọi `revit_list_sheets` để theo quy ước hiện hữu của model (nói rõ là suy từ model) và tránh trùng số.
2. Nếu số sheet đề xuất đã tồn tại, tăng số thứ tự thay vì ghi đè.
3. Sau khi tạo/đổi tên, báo lại danh sách sheet đã thay đổi (số cũ → số mới).
