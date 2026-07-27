---
name: workset-standard
description: BIM workset naming and allocation rules for worksharing
triggers: workset, tao workset, phan workset, chia workset, chuyen workset, workset nao, workset1, phan bo workset, worksharing, gan workset
agents: revit_action, revit_data, general
tools: list_worksets, create_workset, set_element_workset, ai_element_filter, select_elements
standard: project
---
# Quy ước Workset (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

## Đặt tên workset — MẪU THAM KHẢO
- Định dạng: `<BỘ MÔN>_<NHÓM>`, viết HOA, không dấu. Ví dụ: `AR_NOI THAT`, `ST_KET CAU CHINH`, `MEP_ONG GIO`.
- Workset bắt buộc của mọi dự án worksharing:
  - `00_LINK_RVT` — mọi Revit link
  - `00_LINK_CAD` — mọi DWG link
  - `01_GRID_LEVEL` — grid, level, scope box (khóa lại sau khi ổn định)
  - Mỗi bộ môn tối thiểu 1 workset riêng
- KHÔNG để element rơi vào `Workset1` — dấu hiệu model chưa được tổ chức.

## Quy trình khi được yêu cầu tổ chức workset
1. **Tra BEP/tiêu chuẩn dự án trước** để lấy đúng danh mục workset; rồi `list_worksets` để xem hiện trạng.
2. Thiếu workset chuẩn nào → `create_workset` bổ sung theo tên ở trên.
3. Phân bổ element: dùng `ai_element_filter` gom element theo category/bộ môn, rồi `set_element_workset` theo lô.
4. Báo cáo bảng: | Workset | Số element | Ghi chú |.

## Lưu ý an toàn
- Đổi workset hàng loạt là thao tác lớn — nêu số lượng element trước khi thực hiện.
- Link RVT/CAD luôn về workset `00_LINK_*` để tắt/bật nhanh khi mở model.
