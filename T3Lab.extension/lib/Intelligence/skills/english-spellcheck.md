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

## Nguồn quét — MỌI NƠI chứa text người viết, không chỉ TextNote
Đường quét (`check_spelling` → `collect_spellcheck_text`) gom TẤT CẢ nguồn text trong 1 lượt trên main thread:
1. **TextNotes** (gồm text trong legend/drafting view, và text đặt trực tiếp trên sheet).
2. **Title block**: mọi tham số chuỗi của title block instance (drawing title, ghi chú, revision text, project name/client/address hiển thị trên khung tên...).
3. **Project Information**: TẤT CẢ tham số chuỗi (không chỉ name/client/address/status — gồm cả tham số dự án tuỳ biến).
4. **Revision**: trường Description.
5. **Sheet**: tên + số sheet. **View**: tên + "Title on Sheet".
6. **Room / Level / Grid**: tên.
7. **Dimension**: text override (ValueOverride/Above/Below/Prefix/Suffix — nơi hay có "VERFIY ON SITE").
8. **Model Text** (chữ 3D) và **tên Schedule**.
- Muốn thêm nguồn: sửa nhánh `collect_spellcheck_text` trong `lib/core/server.py`.
- User đang chọn sẵn phần tử ("đoạn text này"): dùng `revit_get_selected_elements` và CHỈ kiểm tra phần được chọn.
- "trong view này" → `view_only=True` (chỉ TextNote/Dimension/ModelText của active view).

## Cơ chế phát hiện lỗi — KHÔNG phụ thuộc LLM cục bộ
Việc quét (`/english-spellcheck` không kèm "fix") do **engine từ điển tất định** đảm nhiệm (`lib/Services/spell_dictionary.py` + `spell_checker.check_deterministic`), KHÔNG đưa từng batch cho LLM tự soát nữa. Lý do: model cục bộ nhỏ (llama3.1:8b, qwen...) thường trả `NO_ERRORS` cho lỗi hiển nhiên (CEMTITIOUS, ACYLIC, WATERPROFFING...) nên báo "sạch" trong khi Claude-qua-MCP bắt hết. Engine dùng từ điển tần suất tiếng Anh (`data/en_freq.json.gz`, ~160k từ) trộn với thuật ngữ ngành (`data/construction_terms.txt`) + sửa lỗi theo khoảng cách chỉnh sửa (Norvig), ưu tiên thuật ngữ ngành khi hoà (SETING→SETTING chứ không phải SEEING). Lỗi dùng-sai-từ mà cả hai đều đúng chính tả (to advice→to advise) do `data/confusables.json` xử lý. Engine chạy offline, không cần AI provider. Chỉ khi thiếu asset từ điển mới quay lại đường LLM cũ. Sửa từ điển: chạy `python3 dev/build_spell_dictionary.py`; test: `python3 dev/test_spell_checker.py`.

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
