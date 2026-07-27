---
name: export-standard
description: PDF/DWG export procedure and deliverable file naming
triggers: xuat pdf, in pdf, export pdf, xuat dwg, export dwg, xuat ban ve, xuat ho so, in ho so, xuat ifc, export image, xuat anh, xuat hinh, print pdf, dwg version
agents: revit_action, revit_data, export
tools: revit_list_sheets, export_sheets_pdf, export_dwg, export_image
standard: project
---
# Quy trình xuất hồ sơ PDF / DWG (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

## Trước khi xuất
1. `revit_list_sheets` — chốt danh sách sheet sẽ xuất, báo lại số lượng cho người dùng.
2. Kiểm tra sheet có revision/số phát hành chưa; sheet thiếu tên/số → cảnh báo trước.

## Đặt tên file xuất — MẪU THAM KHẢO (theo quy định giao nộp của dự án)
- PDF từng sheet: `<SỐ SHEET>_<TÊN SHEET>.pdf`, ví dụ `A-101_MAT BANG TANG 1.pdf`.
- Bộ hồ sơ gộp: `<MÃ DỰ ÁN>_<GIAI ĐOẠN>_<NGÀY YYYYMMDD>.pdf`.
- DWG: cùng tên với sheet; mỗi sheet một file, không dồn model space.

## Thực hiện
- PDF: `export_sheets_pdf` — mặc định khổ giấy theo title block của sheet, màu theo chuẩn hồ sơ (đen trắng trừ khi người dùng yêu cầu màu).
- DWG: `export_dwg` — báo rõ version DWG (mặc định 2018 nếu bên nhận không yêu cầu khác).
- Hình minh họa/3D: `export_image` với view được chỉ định.

## Sau khi xuất
- Báo cáo: số file, thư mục đích, file lỗi (nếu có). Đề xuất lưu vào thư mục dự án theo ngày để có vết bàn giao.
