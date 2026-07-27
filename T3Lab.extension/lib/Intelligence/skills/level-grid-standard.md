---
name: level-grid-standard
description: Level and grid creation rules - naming, elevations, axes
triggers: tao level, them level, tao grid, luoi truc, cao do, them tang, tang moi, truc moi, create level, create grid, new level, add grid, grid line, cao do tang, danh so truc
agents: revit_action, revit_data, modeling
tools: list_levels, create_level, create_grid, get_elements_by_level
standard: project
---
# Quy ước Level & Grid (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

## Level (cao độ) — MẪU THAM KHẢO
- Tên: `T01`, `T02`... cho tầng; `T01M` cho lửng; `MAI`, `HAM 1` cho đặc biệt. Thống nhất theo tên có sẵn của dự án — gọi `list_levels` xem trước.
- Cao độ nhập theo mét, đúng cốt kiến trúc (FFL). Khi tạo hàng loạt tầng điển hình: nêu bảng tên + cao độ để người dùng duyệt trước khi tạo.
- KHÔNG tạo level trùng cao độ level đã có; chênh < 50mm coi như trùng — cảnh báo thay vì tạo.

## Grid (lưới trục)
- Trục số `1, 2, 3...` theo một phương; trục chữ `A, B, C...` phương còn lại (bỏ I, O dễ nhầm 1, 0).
- Grid tạo bằng `create_grid` phải chạy phủ hết phạm vi công trình; khoảng cách trục nêu rõ trong kế hoạch trước khi tạo.
- Sau khi tạo xong: liệt kê lại toàn bộ grid/level mới để người dùng kiểm tra, nhắc pin + đưa vào workset `01_GRID_LEVEL` (nếu model worksharing).

## Kiểm tra nhanh
- `get_elements_by_level` để xem tầng nào đang trống/bất thường khi người dùng nghi ngờ model gán sai tầng.
