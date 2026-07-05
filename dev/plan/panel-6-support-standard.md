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
- [x] Gỡ log tạm · cập nhật bảng + README
- [ ] Commit: `fix(panel-support): <mô tả>`

| Tool | Trạng thái | Ghi chú |
|------|-----------|---------|
| MCPControl | ✅ | Bug path đã sửa 2026-07-03 · user xác nhận chung 2026-07-05: hoạt động tốt |
| Feedback | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| PDF import | 🔄 | Gray-gutter 2026-07-03. **2026-07-05 F11 — 2 vòng sửa**: (v1) grid kẹt "Loading views…" + crash do hoãn nạp sang `ContentRendered`/`Loaded` → nạp đồng bộ trong `__init__`. (v2) user báo vẫn **mở chậm + crash**: DataGrid tắt virtualization (`EnableRowVirtualization="False"`) render toàn bộ row → bật lại virtualization + `ViewItem` dùng `INotifyPropertyChanged` + chuyển selection sang setter (bỏ event checkbox + `Items.Refresh`). Chờ user reload pyRevit + test lại (xem Phát sinh) |
| T3LabAssistant | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt (smoke test sâu agentic-upgrade vẫn theo roadmap riêng) |
| ManaTabs | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| Ribbon Names | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| BG Theme | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| UIShowcase | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| AutoWork | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| CloudLinks ×3 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |

> 2026-07-05 — user xác nhận tổng quát: **"các tool hiện tại đều đã hoạt động tốt"** (đóng
> luôn Ngày 11 theo xác nhận này — không có phiên review tĩnh riêng cho panel này). Mở lại
> nếu phát hiện lỗi khi dùng thực tế.

## Phát sinh

- **2026-07-03 — MCPControl không mở được**: cùng bug path như ManaSheets/ManaViews (xem `panel-3-views-sheets.md` Phát sinh) — `EXT_DIR` đi lên thừa 1 cấp, vượt quá `T3Lab.extension`. Đã sửa, verify path trỏ đúng `MCPControl.xaml`. Chưa được user xác nhận mở thử trong Revit.
- **2026-07-03 — PDF import: gray gutter quanh 2 panel nội dung** (yêu cầu UI trực tiếp từ user, không phải bug): `PDFImport.xaml` có background cửa sổ `#E4E4E7` lộ ra thành viền/khe hở quanh 2 panel trắng ("PDF File" settings bên trái, "Target Views" bên phải) do content `Grid Grid.Row="1"` có `Margin="18,18,18,10"` và khoảng cách `Margin="0,0,14,0"` giữa 2 panel — cùng lớp bug đã gặp ở ManaSelect/ManaDWG (xem `panel-1-annotation-select.md` Phát sinh). Giao cho `@ui-agent` sửa theo đúng pattern đã dùng ở `ManaViews.xaml`: bọc content trong `Border Background="#F8FAFC"` full-bleed (bỏ Margin ngoài), 2 panel giờ `BorderThickness="0"`/`CornerRadius="0"` sát cạnh cửa sổ, khe hở giữa 2 panel thay bằng 1 đường viền mảnh `#E2E8F0` thay vì khoảng gray. Đã verify `sync_wpf_styles.py --check` (0 lệch) + `audit_ui.py --quiet` (0/54 vấn đề) + XML well-formed. Chưa được user xác nhận trực quan trong Revit.
- **2026-07-05 — PDF import: mở tool không hiện thông tin + crash file (bug thật, user báo — F11)**: Đây là bug functional thật, KHÔNG phải UI — phơi ra rằng xác nhận chung "các tool hoạt động tốt" (2026-07-05) chưa thực sự test functional tool này. Triệu chứng: mở tool thì grid Target Views kẹt ở overlay "Loading views…" (không hiện view nào) và làm crash/đóng file Revit. Truy vết trong `PDFImportDialog.py`:
  - **Root cause 1 (không hiện thông tin)**: bản refactor `0a2a886` hoãn việc nạp grid sang event `ContentRendered` (`_on_content_rendered` → `_load_views`). `ContentRendered` **không chạy đáng tin dưới modal `ShowDialog()` trong Revit** — khi không bắn, `_load_views()` không bao giờ chạy → `pnl_loading` không bao giờ bị collapse → grid kẹt "Loading views…".
  - **Root cause 2 (crash)**: bind `ItemsSource` + mutate `ObservableCollection` được đặt trong event `Loaded`/`ContentRendered` — tức SAU khi cửa sổ bắt đầu render. Mutate collection đang bind DataGrid (non-virtualized) giữa lúc layout có thể **hard-crash host Revit** (lỗi tầng WPF/native, không phải exception Python bắt được).
  - **Fix**: dời việc tạo OC + bind + `_load_views()` về thẳng `__init__`, chạy **đồng bộ trước `ShowDialog()`** (lúc cửa sổ chưa render) — đúng pattern đã kiểm chứng ở `ManaViewsDialog`/`ManaContainsDialog`/`ManaSheetsDialog` (nạp trong `__init__`, mutate 1 OC bound in-place). Xóa `_on_loaded`/`_on_content_rendered`/`_views_loaded`. Guard `mode_changed` (`if self._oc is None`) vẫn giữ vì handler này bắn lúc parse XAML trước khi `_oc` được tạo.
  - Verify tĩnh: `ast.parse` OK · `audit_tools`/`audit_ui`/`sync_wpf_styles --check` đều sạch.
  - **Vòng 2 (2026-07-05) — user báo lại: vẫn mở CHẬM + crash file**: root cause còn lại là DataGrid `grid_views` để `EnableRowVirtualization="False"` → model nhiều view thì WPF dựng toàn bộ row cùng lúc (mỗi row có checkbox template + nhiều Path/Border) → chậm + hard-crash. So với `ManaViews` (chạy tốt với nhiều view): nó KHÔNG tắt virtualization và checkbox KHÔNG có event, chỉ bind `IsSelected`. PDFImport phải tắt virtualization vì checkbox có handler `view_selection_changed` gọi `Items.Refresh()` mỗi lần tick (kỵ virtualization). Sửa triệt để: (1) bật lại virtualization (`VirtualizingPanel.IsVirtualizing=True` + `VirtualizationMode=Recycling`, bỏ `EnableRowVirtualization="False"`); (2) `ViewItem` implement `INotifyPropertyChanged` (giống `RenumberItem` của ManaSheets) nên PAGE pill + checkbox tự cập nhật, bỏ mọi `Items.Refresh()`; (3) bỏ event `Checked/Unchecked` XAML, chuyển tính-lại-page + single-select vào **setter `IsSelected`** (`_on_item_toggle`) — setter chỉ chạy khi user thực sự tick (hướng ghi-nguồn của binding), KHÔNG chạy khi row được realize lúc cuộn → an toàn với virtualization. Static clean (`ast.parse` + XML well-formed + 3 audit). **Chưa smoke test Revit** — cần user reload pyRevit + mở lại: kỳ vọng mở nhanh kể cả model nhiều view, không crash; tick/single-select (All→One)/page-number vẫn đúng.
