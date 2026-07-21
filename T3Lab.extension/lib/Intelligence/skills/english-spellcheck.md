---
name: english-spellcheck
description: Check & fix English spelling/wording across ALL model text (notes, sheets, views, rooms, levels, project info...)
triggers: chinh ta, spelling, spellcheck, spell check, check tieng anh, kiem tra tieng anh, typo, loi chinh ta, sai chinh ta, review text
agents: revit_data, revit_action, qa_check, general
tools: ai_element_filter, revit_list_sheets, revit_list_views, revit_get_selected_elements, revit_get_project_info, set_parameter, rename_element
---
# Check chính tả tiếng Anh trong model (T3Lab)

## Phạm vi mặc định: TOÀN DỰ ÁN — KHÔNG HỎI
Mọi yêu cầu check/kiểm tra (kể cả gọi qua `/english-spellcheck` không kèm nội dung) mặc định quét **toàn bộ dự án NGAY LẬP TỨC** — không hỏi phạm vi, không tự giới hạn vào current view, không cần selection hay active view. Chỉ giới hạn khi user NÓI RÕ ("trong view này", "sheet này", "phần đang chọn"). Đã gọi tool quét rồi thì phân tích luôn kết quả — không quét xong rồi quay lại hỏi phạm vi.

## Nguồn quét — TOÀN BỘ text trong model, không chỉ TextNote
Quét đủ các nguồn sau (batch các call read trong một lượt khi có thể):
1. **TextNotes**: `ai_element_filter` với `category="TextNotes"` (không truyền parameter_name) — trường `text` là nội dung cần kiểm tra. Quét toàn dự án, gồm cả text trong legend/drafting view.
2. **Tên sheet + số sheet**: `revit_list_sheets`.
3. **Tên view**: `revit_list_views`.
4. **Tên room**: `ai_element_filter` với `category="Rooms"` (field `name`).
5. **Tên level / grid**: `ai_element_filter` với `category="Levels"` rồi `category="Grids"`.
6. **Thông tin dự án**: `revit_get_project_info` (project name, client, address, status).
- User đang chọn sẵn phần tử ("đoạn text này"): dùng `revit_get_selected_elements` và CHỈ kiểm tra phần được chọn.
- `total_count` > `count` nghĩa là danh sách bị cắt — tăng `limit` để quét đủ, hoặc báo rõ "đã kiểm tra X/Y".

## Cách kiểm tra
- Tự soát chính tả + ngữ pháp tiếng Anh trên từng chuỗi; chỉ báo lỗi THẬT (sai chính tả, sai từ, lặp từ), không sửa văn phong.
- KHÔNG coi là lỗi: viết tắt ngành (TYP, UNO, EQ, FFL, SSL, RC, DN, C/W, ACMV, GRC, NTS...), mã bản vẽ/sheet code (A201, S-101...), kích thước và đơn vị (50"x60", 100mm, Ø, @), tên riêng/hãng.
- CHỮ HOA TOÀN BỘ là quy ước bản vẽ, không phải lỗi.

## Báo cáo
Bảng markdown, chỉ liệt kê mục có lỗi:

| ID | Nguồn (TextNote/Sheet/View/Room/Level/Project) | Hiện tại | Đề xuất | Lý do |

Kết thúc bằng tổng kết theo số liệu tool trả về (`total_count`, số dòng list): "đã kiểm tra X TextNote, Y sheet, Z view, R room, L level/grid + project info — N lỗi". Không có lỗi → nói rõ đã quét đủ các nguồn, không thấy lỗi.

## Sửa (CHỈ sau khi user xác nhận)
- TextNote: `set_parameter` với `parameter_name="Text"`, `value` = câu đã sửa.
- Tên sheet/view/room/level/grid: `rename_element`.
- Thông tin dự án (project name/client/address): `set_parameter` trên project info nếu tool cho phép, không được thì hướng dẫn user sửa trong Project Information.
- Sửa xong tổng kết số mục đã sửa + danh sách ID.
