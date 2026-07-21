---
name: export-standard
description: PDF/DWG export procedure and deliverable file naming
triggers: xuat pdf, in pdf, export pdf, xuat dwg, export dwg, xuat ban ve, xuat ho so, in ho so, xuat ifc, export image, xuat anh, xuat hinh, print pdf, dwg version
agents: revit_action, revit_data, export
tools: revit_list_sheets, export_sheets_pdf, export_dwg, export_image
---
# Quy trình xuất hồ sơ PDF / DWG (T3Lab)

## Trước khi xuất
1. `revit_list_sheets` — chốt danh sách sheet sẽ xuất, báo lại số lượng cho người dùng.
2. Kiểm tra sheet có revision/số phát hành chưa; sheet thiếu tên/số → cảnh báo trước.

## Đặt tên file xuất
- PDF từng sheet: `<SỐ SHEET>_<TÊN SHEET>.pdf`, ví dụ `A-101_MAT BANG TANG 1.pdf`.
- Bộ hồ sơ gộp: `<MÃ DỰ ÁN>_<GIAI ĐOẠN>_<NGÀY YYYYMMDD>.pdf`.
- DWG: cùng tên với sheet; mỗi sheet một file, không dồn model space.

## Thực hiện
- PDF: `export_sheets_pdf` — mặc định khổ giấy theo title block của sheet, màu theo chuẩn hồ sơ (đen trắng trừ khi người dùng yêu cầu màu).
- DWG: `export_dwg` — báo rõ version DWG (mặc định 2018 nếu bên nhận không yêu cầu khác).
- Hình minh họa/3D: `export_image` với view được chỉ định.

## Sau khi xuất
- Báo cáo: số file, thư mục đích, file lỗi (nếu có). Đề xuất lưu vào thư mục dự án theo ngày để có vết bàn giao.
