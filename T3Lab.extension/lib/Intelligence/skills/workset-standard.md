---
name: workset-standard
description: Quy uoc dat ten va phan bo workset chuan BIM khi lam viec nhom (worksharing)
triggers: workset, tao workset, phan workset, chia workset, chuyen workset, workset nao
agents: revit_action, revit_data, general
tools: list_worksets, create_workset, set_element_workset, ai_element_filter, select_elements
---
# Quy ước Workset (T3Lab)

## Đặt tên workset
- Định dạng: `<BỘ MÔN>_<NHÓM>`, viết HOA, không dấu. Ví dụ: `AR_NOI THAT`, `ST_KET CAU CHINH`, `MEP_ONG GIO`.
- Workset bắt buộc của mọi dự án worksharing:
  - `00_LINK_RVT` — mọi Revit link
  - `00_LINK_CAD` — mọi DWG link
  - `01_GRID_LEVEL` — grid, level, scope box (khóa lại sau khi ổn định)
  - Mỗi bộ môn tối thiểu 1 workset riêng
- KHÔNG để element rơi vào `Workset1` — dấu hiệu model chưa được tổ chức.

## Quy trình khi được yêu cầu tổ chức workset
1. `list_worksets` để xem hiện trạng trước.
2. Thiếu workset chuẩn nào → `create_workset` bổ sung theo tên ở trên.
3. Phân bổ element: dùng `ai_element_filter` gom element theo category/bộ môn, rồi `set_element_workset` theo lô.
4. Báo cáo bảng: | Workset | Số element | Ghi chú |.

## Lưu ý an toàn
- Đổi workset hàng loạt là thao tác lớn — nêu số lượng element trước khi thực hiện.
- Link RVT/CAD luôn về workset `00_LINK_*` để tắt/bật nhanh khi mở model.
