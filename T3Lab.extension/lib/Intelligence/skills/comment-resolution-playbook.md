---
name: comment-resolution-playbook
description: Playbook for resolving PDF markup comments traced back to Revit models
triggers: cmt, comment, markup, bluebeam, hoan thien cmt, xu ly comment, ghi chu ban ve, resolve comments, tra loi comment, phan hoi comment, review comment
agents: comment, revit_action, multi_doc
tools: revit_list_sheets, list_open_documents, switch_active_document, create_text_note, set_parameter
---
# Playbook xử lý comment bản vẽ (T3Lab)

Quy trình chuẩn khi nhận PDF bản vẽ có comment/markup review:

## 1. Phân loại từng comment
- **Sửa model** (di chuyển/đổi kích thước/thêm bớt element): cần thao tác trong Revit.
- **Sửa annotation** (dim, tag, text, ký hiệu): thao tác trên view/sheet.
- **Thông tin/trả lời** (câu hỏi, xác nhận): chỉ cần phản hồi, không sửa model.
- **Ngoài phạm vi** (thuộc bộ môn khác, thiếu thông tin): ghi chú lại, không tự xử lý.

## 2. Truy vết vị trí
- Xác định số sheet từ khung tên bản vẽ; đối chiếu `revit_list_sheets` của các model đang mở (`list_open_documents`).
- Nếu sheet thuộc model khác đang mở → đề xuất `switch_active_document` trước.
- Nếu model chưa mở → báo người dùng mở model, KHÔNG đoán.

## 3. Đề xuất phương án
- Mỗi comment: 1 đề xuất ngắn gọn (hành động + đối tượng + vị trí), kèm mức tự tin.
- Comment mơ hồ → hỏi lại, không suy diễn.

## 4. Thực hiện (chỉ khi người dùng xác nhận)
- Sửa từng comment một, báo kết quả theo số thứ tự comment.
- Ghi vết trong model: `create_text_note` trên sheet tương ứng dạng `[CMT-<n>] <nội dung> — <trạng thái>`, hoặc điền tham số Comments của element liên quan qua `set_parameter`.
- Không gộp nhiều comment vào một thao tác nếu không cùng loại.

## 5. Tổng kết
- Bảng: | # | Comment | Sheet | Phương án | Trạng thái |
- Trạng thái: ĐÃ XỬ LÝ / ĐÃ GHI CHÚ / CHỜ THÔNG TIN / NGOÀI PHẠM VI.
