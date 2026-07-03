# Panel 6 — Support + Standard (Ngày 11)

> 9 tool + 3 urlbutton. Nhóm này đụng vào UI Revit (ribbon/tabs/theme), network và server MCP — nhiều tool phụ thuộc máy/version Revit hơn là model.

## Tool 1/9 — MCPControl
Chain: launcher → `MCPControlDialog.py` (348 loc) → `MCPControl.xaml` · điều khiển `lib/core/server.py` (port 48884)

- [ ] Start server → status hiển thị đúng; `curl http://localhost:48884/mcp` (hoặc endpoint health) phản hồi
- [ ] Stop server → port nhả ra thật (kiểm tra bằng netstat)
- [ ] Start khi port 48884 đang bị chiếm → thông báo rõ, không crash
- [ ] Token file `%APPDATA%\T3LabAI\mcp_token.txt` được tạo (bridge.py cần nó)
- [ ] Ghi chú:

## Tool 2/9 — Feedback
Chain: self-contained (232 loc) → `Feedback.xaml` · gửi ra ngoài qua network

- [ ] Gửi feedback nội dung tiếng Việt có dấu → nhận được đúng encoding
- [ ] Không có mạng / endpoint chết → thông báo thân thiện, không treo UI
- [ ] Field rỗng → validate
- [ ] Ghi chú:

## Tool 3/9 — PDF import
Chain: launcher → `PDFImportDialog.py` → `PDFImport.xaml`

- [ ] Revit 2022+: import 1 PDF → hiện đúng trên view
- [ ] File không phải PDF / PDF có mật khẩu → thông báo
- [ ] Kiểm tra API import PDF theo version Revit đang chạy (API đổi giữa các version)
- [ ] Ghi chú:

## Tool 4/9 — T3LabAssistant ⚠️ tool lớn nhất codebase (3594 loc)
Chain: self-contained → `T3LabAssistant.xaml` · import động (`imp` → `importlib` fallback) · gọi `core.server`

- [ ] Mở UI chat không lỗi (XAML 2100+ dòng)
- [ ] **Xác nhận nhánh import động**: trên IronPython 2.7 `importlib.util` không tồn tại → code phải rơi vào nhánh `imp` (script.py:165-184). Thêm log tạm xác nhận nhánh nào chạy
- [ ] Server MCP chưa chạy → assistant thông báo, không crash
- [ ] Server đang chạy → gửi 1 lệnh đơn giản, nhận phản hồi
- [ ] Load module BatchOut động (`_load_batchout_mod`) → hoạt động
- [ ] Đóng cửa sổ khi đang chờ phản hồi → không treo Revit
- [ ] Ghi chú:

## Tool 5/9 — ManaTabs
Chain: launcher → `ManaTabsDialog.py` → `ManaTabs.xaml` (**2 bare except**, thao tác Windows API lên UI Revit)

- [ ] Revit mở nhiều document/tab → list + thao tác tab đúng
- [ ] Test đúng version Revit triển khai (Windows API dò cửa sổ dễ vỡ theo version)
- [ ] 1 document duy nhất → không lỗi
- [ ] Ghi chú:

## Tool 6/9 — Ribbon Names
Chain: launcher → `RibbonNamesDialog.py` → `RibbonNames.xaml` (sửa text ribbon qua AdWindows.dll)

- [ ] Đổi tên 1 nút ribbon → hiển thị ngay
- [ ] Restart Revit → kiểm tra persist hay reset (ghi nhận hành vi mong muốn)
- [ ] Tab/panel không tồn tại → không crash
- [ ] Ghi chú:

## Tool 7/9 — BG Theme
Chain: launcher → `BGThemeDialog.py` → `BGTheme.xaml`

- [ ] Đổi background/theme → áp dụng ngay
- [ ] Revit đang ở dark mode → tool hoạt động đúng cả 2 theme
- [ ] Revert về mặc định → OK
- [ ] Ghi chú:

## Tool 8/9 — UIShowcase (Standard.panel)
Chain: launcher → `UIShowcaseDialog.py` (129 loc) → `UIStandardShowcase.xaml`

- [ ] Sau fix F3 (GĐ1): mở từ repo checkout → OK
- [ ] **Mở từ extension deploy rời repo** (copy T3Lab.extension sang máy khác/%APPDATA%) → phải vẫn mở được (fallback path hoạt động)
- [ ] Showcase hiển thị đủ component chuẩn Lumina (dùng làm reference visual cho GĐ2 ngày 4)
- [ ] Ghi chú:

## Tool 9/9 — AutoWork (Standard.panel)
Chain: self-contained (353 loc) → `AutoWork.xaml` · không có Transaction trong script

- [ ] Chạy happy path → xác định tool có ghi model không; nếu có ghi mà không qua Transaction → phải fail (Revit chặn) → xác nhận flow thật
- [ ] Chức năng chính hoạt động đúng mô tả
- [ ] Ghi chú:

## CloudLinks.stack (3 urlbutton — 2 phút)
- [ ] Autodesk Health / Autodesk Forma / Bluebeam Status → mở đúng URL trong browser

## Chốt ngày 11
- [ ] Gỡ log tạm · cập nhật bảng + README
- [ ] Commit: `fix(panel-support): <mô tả>`

| Tool | Trạng thái | Ghi chú |
|------|-----------|---------|
| MCPControl | ⬜ | Bug path (xem Phát sinh) đã sửa 2026-07-03, chưa test functional |
| Feedback | ⬜ | |
| PDF import | ⬜ | |
| T3LabAssistant | ⬜ | |
| ManaTabs | ⬜ | |
| Ribbon Names | ⬜ | |
| BG Theme | ⬜ | |
| UIShowcase | ⬜ | |
| AutoWork | ⬜ | |
| CloudLinks ×3 | ⬜ | |

## Phát sinh

- **2026-07-03 — MCPControl không mở được**: cùng bug path như ManaSheets/ManaViews (xem `panel-3-views-sheets.md` Phát sinh) — `EXT_DIR` đi lên thừa 1 cấp, vượt quá `T3Lab.extension`. Đã sửa, verify path trỏ đúng `MCPControl.xaml`. Chưa được user xác nhận mở thử trong Revit.
