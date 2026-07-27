---
name: view-template-standard
description: View naming rules and consistent view template application
triggers: view template, ap template, dat ten view, ten view, tao view moi, nhan ban view, duplicate view, apply template, rename view, view name, don view, view rac
agents: revit_action, revit_data, general
tools: revit_list_views, create_view, duplicate_view, apply_view_template, create_view_filter, set_active_view
standard: project
---
# Quy ước View & View Template (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

## Đặt tên view — MẪU THAM KHẢO
- Cấu trúc: `<GIAI ĐOẠN>_<LOẠI>_<TẦNG/PHẠM VI>_<MÔ TẢ>`, ví dụ:
  - `HSTK_MB_T01_KIEN TRUC` (mặt bằng tầng 1 hồ sơ thiết kế)
  - `WIP_3D_COORD_MEP` (view 3D phối hợp — Working)
- View làm việc cá nhân prefix `WIP_`, view lên sheet prefix theo giai đoạn (`HSTK_`, `CONCEPT_`...). View không đúng chuẩn coi như view rác.

## Áp template
1. `revit_list_views` xem view + template hiện có của dự án trước.
2. View đặt lên sheet BẮT BUỘC gán view template tương ứng loại bản vẽ — dùng `apply_view_template`.
3. Tạo view mới (`create_view` / `duplicate_view`): đặt tên đúng chuẩn ngay lúc tạo, gán template ngay sau đó.
4. Cần lọc/tô màu theo điều kiện → `create_view_filter` thay vì override thủ công từng element.

## Quy trình dọn view
- Liệt kê view không nằm trên sheet và không có prefix `WIP_` → đề xuất danh sách xóa, KHÔNG tự xóa.
- Báo cáo: | View | Template | Trên sheet? | Đề xuất |.
