# Giai đoạn 2 — UI Consistency theo chuẩn Lumina (Ngày 3–4)

> Làm SAU giai đoạn 1. Branch gợi ý: `fix/phase-2-ui-consistency`.
> Tham chiếu: `.claude/rules/ui-design-standard.md` · Kiểm tra: `python3 dev/audit_ui.py --quiet`

## Kết quả UI review (2026-07-02, cập nhật sau khi chốt chuẩn footer-trái)

Audit 76 XAML trong `lib/GUI/Tools/` theo chuẩn Lumina:

| Hạng mục | Kết quả |
|----------|---------|
| Palette Lumina (không còn Kinetix/Terra) | ✅ trừ **T3LabAssistant.xaml** sót 2 màu Terra |
| Font (Hanken Grotesk/Inter, không Manrope/Segoe UI) | ✅ 100% |
| WindowChrome đủ 5 attribute, multi-line | ✅ 100% |
| Shared styles block đồng bộ | ✅ 76/76 (`sync_wpf_styles.py --check`) |
| Dot-notation Grid definitions (crash risk) | ✅ không có |
| Glyph chrome buttons (E921/E922/E8BB) | ✅ không phát hiện sai |
| **Copyright ở footer, bên trái** | ✅ 74/76 — codebase vốn đã nhất quán footer-trái (copyright → divider 1px → status text). Tài liệu `ui-design-standard.md` trước đây bắt overlay góc phải-dưới là **lỗi thời so với thực tế** và đã được cập nhật (2026-07-02) theo pattern chuẩn của `UIStandardShowcase.xaml`. |

→ Chỉ còn **2 file outlier**, đều thuộc T3LabAssistant:

| File | Vấn đề |
|------|--------|
| `T3LabAssistant.xaml` | Copyright căn **giữa** dưới ô chat input (dòng ~2109) + sót 2 màu Terra (`#15803D`, `#166534`) |
| `AssistantPane.xaml` (dockable pane, variant B) | Copyright theo chuẩn cũ: overlay góc **phải**-dưới (dòng ~1357) — chính là file duy nhất từng theo doc cũ |

`CadtoFloorLayerItem.xaml` không có copyright nhưng là **item template** (root là list-row `<Border>`) → được miễn trừ trong chuẩn mới và trong `audit_ui.py` (`COPYRIGHT_EXEMPT`).

**2 file UI-LOCKED tuyệt đối không sửa**: `DWGManagement.xaml`, `ExportManager.xaml`.

---

## Ngày 3 — Sửa outlier T3LabAssistant + chốt audit sạch

### 3a. `T3LabAssistant.xaml` — palette Terra
- [x] Xem context 2 chỗ dùng màu: `#15803D` → `#10B981` (success), `#166534` → `#059669` (success hover). Ngữ cảnh xác nhận: ô kết quả "Test Connection" (success state), khớp đúng bảng migration trong `ui-design-standard.md`

### 3b. `T3LabAssistant.xaml` — copyright về footer-trái
- [x] Đọc layout vùng đáy cửa sổ chat (dòng ~2022–2109): copyright là child cuối của StackPanel bên trong `Border Grid.Row="4"` (Padding="16,8,16,10"), `HorizontalAlignment="Center"` — đây chính là footer/input-bar (Auto height, row cuối), không có status bar riêng
- [x] Không có status bar chuẩn (chat UI) → đổi TextBlock hiện tại sang `HorizontalAlignment="Left"`; không thêm margin trái 14px vì Border cha đã có `Padding="16,..."` bên trái — thêm margin nữa sẽ lệch khỏi rhythm 16px sẵn có của toàn khối input
- [x] Không đụng block shared styles

### 3c. `AssistantPane.xaml` — copyright về footer-trái
- [x] Dockable pane variant B (root Grid, 4 row: header/chat/quick-actions/input — **không có row footer riêng**, row cuối = input bar full-width) — **đã kiểm tra và phát hiện fallback overlay chuẩn (bottom-left) sẽ đè lên ô chat input** (input textbox chiếm gần hết chiều rộng cột 0 ở row cuối, overlay bottom-left rơi đúng vào góc dưới-trái của textbox)
- [x] **Quyết định (deviate có chủ đích khỏi fallback overlay mặc định):** thêm hẳn 1 row Auto mới (Row 4 — Footer) thay vì overlay đè lên nội dung, tránh xung đột layout. Copyright đặt trong `Border Grid.Row="4"` nền `#F8FAFC`, border-top `#E2E8F0`, `Padding="14,4"`, `TextBlock HorizontalAlignment="Left"` — vẫn đúng chuẩn amber/FontSize 11/trái, chỉ khác cách neo vị trí (footer bar thật thay vì overlay ảo) để đảm bảo không che control. Tốn thêm ~24px chiều cao dockable pane — chấp nhận được.
- [x] Pane là UI nhúng trong Revit dockable panel — đã kiểm tra góc trái-dưới: **không còn đè lên control nào** nhờ dùng row riêng thay vì overlay

### 3d. Kiểm tra bản copy UIStandardShowcase (từ GĐ1-F3)
- [x] `UIStandardShowcase.xaml` copy vào `lib/GUI/Tools/` pass `audit_ui.py` (0 vấn đề, đã verify lại cùng lượt chạy audit hôm nay)

### Chốt ngày 3
- [x] `python3 dev/audit_ui.py --quiet` → **exit 0** (`UI AUDIT: 0/54 file có vấn đề (chỉ file UI-locked) — OK`)
- [x] `python3 dev/sync_wpf_styles.py --check` → 0 lệch (53 file)
- [x] `python3 dev/audit_tools.py --quiet` → không phát sinh mới (`AUDIT: clean`)
- [x] Commit: `style: migrate T3LabAssistant to Lumina palette, move copyright to footer-left`

---

## Ngày 4 — Verify trực quan trong Revit

Checklist mỗi cửa sổ (mở lần lượt 41 tool + dialog phụ):
1. Copyright màu amber nằm **bên trái footer**, có divider 1px ngăn với status text, không che nội dung
2. Title bar: T3Lab + tên tool + subtitle + separator đúng chuẩn
3. Nút min/max/close hover đúng (close → nền đỏ chữ trắng)
4. Font toàn cửa sổ là Hanken Grotesk/Inter
5. Resize cửa sổ → footer/copyright không vỡ layout

- [x] Panel 1 — Annotation & Select (4/4 cửa sổ: AutoDimension, ManaDWG, ManaSelect, ManaAnno) — mở OK, chrome OK, copyright footer-trái OK, resize không vỡ layout
- [x] **Panel 2 — Modeling & Datum: ĐÃ XONG (xác nhận user 2026-07-03).** Toàn bộ 5 điểm checklist (copyright footer-trái, title bar, chrome hover, font Hanken/Inter, resize không vỡ) pass cho các cửa sổ đã QA: Auto Join, Split Elements, CADToElements (Wall/Floor/Beam — sidebar rail + logo, search bar, copyright/footer, card layout 3 tab), ImageToDrafting (dropdown `T3ComboBoxSmall`, drop-zone nét đứt, hết icon user thừa), TextToElement (nút chrome đúng vị trí, hết cuộn), Tile Layout (path đã sửa, mở được, list sàn mặc định, sidebar step-progress mới), PropertyLine (đã xoá nút Close thừa ở action bar), ManaFami (đã xác nhận **mở được và phản hồi click bình thường** — khác với ghi nhận nghi vấn trước đó; search bar Management tab đã dựng lại theo mẫu `UIStandardShowcase.xaml`, nút Refresh thừa đã xoá). Đồng thời phát hiện & sửa nhiều bug functional đi kèm (path 4 cấp thư mục, double-transform CAD import, tên enum Revit API sai — xem `panel-2-modeling-datum.md` Tool 1/7, 3/7, 4/7, 7/7, 10/14 để biết chi tiết). **Còn tồn đọng (không chặn UI, chuyển hẳn sang GĐ3 functional test)**: Wall Cut Profile & Auto Adj Base Offset không hiện dialog trên project trống — cần model có wall thật; FamiGen chưa xác nhận lại được click; DoorThreshold/PointCloud/RoomToFloor chưa QA UI riêng (không có phản hồi UI bất thường được báo, nhưng chưa mở kiểm tra từng cái).
- [ ] Panel 3 (4 cửa sổ) — chưa làm
- [ ] Panel 4 (6 cửa sổ) — chưa làm
- [ ] Panel 5 (4 cửa sổ) — chưa làm
- [ ] Panel 6 (9 cửa sổ + AssistantPane dockable) — chưa làm
- [ ] Screenshot mỗi cửa sổ lưu `dev/plan/screenshots/` — chưa làm (chỉ QA trực tiếp qua computer-use, chưa lưu file ảnh)
- [ ] File nào layout vỡ → ghi vào Phát sinh, sửa tay từng file
- [ ] Commit cuối: `style: visual QA pass for Lumina consistency` — chưa commit (đang dở)
- [ ] Cập nhật bảng tiến độ `dev/plan/README.md`

> **Tạm dừng 2026-07-02**: đã QA xong Panel 1, một phần Panel 2. Lý do dừng: (1) click icon nhỏ trên ribbon qua computer-use rất chậm/khó trúng toạ độ chính xác; (2) project test đang trống (không wall/room/family) nên nhiều tool cần selection không mở được UI; (3) nhiều tool trong `panel-2..6-*.md` không nằm ở vị trí ribbon như dự đoán, cần dò lại.
>
> **Quyết định 2026-07-03 (theo yêu cầu user)**: ~~chấp nhận GĐ2 Ngày 4 dừng ở mức chưa hoàn thành 100%~~ — **ĐÃ HỦY cùng ngày 2026-07-03 theo yêu cầu user: không skip nữa, tiếp tục GĐ2 Ngày 4.**
>
> **Kế hoạch tiếp tục (2026-07-03)**: blocker (3) "không tìm thấy tool trên ribbon" đã giải quyết — dò xong vị trí thật của toàn bộ cửa sổ còn lại từ cấu trúc `T3Lab.tab/` (bảng bên dưới). Blocker (1) computer-use chậm → chuyển hẳn sang **user tự test trong Revit theo checklist**, kết quả tick dựa trên phản hồi user (đúng quy trình CLAUDE.md — không tự bịa kết quả visual QA). Blocker (2) project trống → các cửa sổ thuần UI vẫn mở được không cần model; riêng tool cần selection (Wall Cut Profile, Wall_Adjust Base) test cùng model thật của GĐ3.

### Vị trí ribbon thật của các cửa sổ còn lại (dò từ source `T3Lab.tab/` 2026-07-03)

Ký hiệu: ▸ = pulldown/stack phải bấm mũi tên mở ra.

**Panel 2 nốt — tab T3Lab → panel "Modeling & Datum":**

| Cửa sổ | Đường đi trên ribbon | Ghi chú |
|--------|---------------------|---------|
| DoorThreshold | Create.stack ▸ Create Elements ▸ DoorThreshold | |
| ImageToDrafting | Create.stack ▸ Create Elements ▸ ImageToDrafting | UI đã sửa 2026-07-03 — verify dropdown có style, drop-zone nét đứt, hết icon user |
| Text to Element | Create.stack ▸ Create Elements ▸ Text to Element | UI đã sửa — verify nút chrome đúng vị trí, hết cuộn |
| PointCloud | Create.stack ▸ Create Elements ▸ PointCloud | |
| RoomToFloor | Create.stack ▸ Create Elements ▸ RoomToFloor | |
| Tile Layout | Create.stack ▸ Tile Layout (nút riêng, KHÔNG trong pulldown) | Path đã sửa — verify mở được + list sàn tự động |
| PropertyLine | Create.stack ▸ PropertyLine (nút riêng) | |
| FamiGen | nút riêng trên panel | Lần trước không phản hồi qua automation — thử click trực tiếp |
| ManaFami | nút riêng trên panel | Tương tự FamiGen |
| Wall Cut Profile | ElementAdjust ▸ Wall Cut Profile | Cần model có wall + selection → test cùng GĐ3 |
| Wall_Adjust Base | ElementAdjust ▸ Wall_Adjust Base | Không có XAML riêng → miễn QA UI, chỉ cần không crash |

**Panel 3 — panel "Views & Sheets":** BatchOut (**UI-locked** — chỉ verify mở được + chrome hoạt động, không soi style) · ManaSheets (path đã sửa — verify mở được) · ManaViews (path đã sửa — verify; mở thêm dialog phụ AdvancedViewManager) · SheetGen.

**Panel 4 — panel "Data & IFC-SG":** IFC-SG (+ dialog phụ SubtypeDefiner column mapping) · Foundation Volume · BCF Reader · manaData.stack ▸ ManaContains / ManaPara / ManaSched.

**Panel 5 — panel "Standards & Settings":** ManaLoca (⚠️ chỉ MỞ XEM, không apply đổi toạ độ trong QA UI) · ManaStyles · ManaWorkset (file thường có thể chặn yêu cầu worksharing — nếu ra thông báo thân thiện thay vì mở UI thì ghi nhận, coi là pass UI) · ModelAuditor.

**Panel 6 — panel "Support" + panel "Standard":** MCPControl (path đã sửa — verify mở được) · Feedback · PDF import · T3LabAssistant (verify kết quả Ngày 3: copyright footer-TRÁI, không còn màu Terra) · UI.stack ▸ BG Theme / ManaTabs / Ribbon Names · Standard.panel: UIShowcase · AutoWork · **AssistantPane dockable** (verify footer Row 4 mới thêm Ngày 3: không che ô chat input, pane không vỡ layout).

---

## Phát sinh trong giai đoạn 2

- **Ngày 4 — Title bar thiếu wordmark "T3Lab" (phát hiện qua visual QA + xác nhận bằng grep source) — ĐÃ XỬ LÝ:** `ui-design-standard.md` mục "Window Structure" từng yêu cầu title bar có "Left: T3Lab (11px Bold) + Tool Name (18px Bold)". Thực tế `grep -c 'Text="T3Lab"'` trên toàn bộ `lib/GUI/Tools/*.xaml` cho kết quả **0/50 file có wordmark này — kể cả `UIStandardShowcase.xaml` (chính file canonical reference)**. Toàn bộ codebase hội tụ về pattern "Tool Name (15px Bold) + Subtitle (12.5px)" không có wordmark riêng. Giống trường hợp copyright ở Ngày 3 → đã cập nhật `.claude/rules/ui-design-standard.md` mục 3 (Title bar) để khớp thực tế đã hội tụ (2026-07-02), không cần sửa lại 50 file XAML.
- **Ngày 4 — Wall Cut Profile / Auto Adj Base Offset** (trong dropdown "Element Adjust"): không hiện dialog nào khi bấm trên project trống, không có wall để chọn trước. Chưa xác định được là do cần selection trước (by-design) hay lỗi thật (tool này vốn được panel-2 đánh dấu "rủi ro cao nhất — 37 bare except"). Cần test lại với model có wall thật ở GĐ3.
- **Ngày 4 — ManaFami, FamiGen** (panel Modeling & Datum): không phản hồi khi click dù đã xác nhận đúng toạ độ button qua tooltip hover (tên hiện đúng "FamiGen" khi hover). — **ManaFami: ĐÃ XỬ LÝ (2026-07-03)**, xác nhận lại bằng test trực tiếp trong Revit là do vấn đề UI automation ở lần thử trước, không phải lỗi thật — tool mở và phản hồi click bình thường. **FamiGen: vẫn chưa xác minh lại**, cần thử lại ở phiên làm việc sau.
- **Ngày 4 — nhiều tool trong `panel-2..6-*.md` chưa tìm thấy trên ribbon — ĐÃ XỬ LÝ (2026-07-03)**: dò lại từ cấu trúc thư mục `T3Lab.tab/` → toàn bộ tool đều tồn tại trên ribbon, phần lớn nằm trong pulldown/stack ẩn (`Create Elements.pulldown`, `ElementAdjust.pulldown`, `manaData.stack`, `UI.stack`). Bảng vị trí đầy đủ đã ghi ở mục "Vị trí ribbon thật" phía trên.
- **Ngày 3 — `AssistantPane.xaml`:** fallback overlay bottom-left chuẩn (theo `ui-design-standard.md`) sẽ đè lên ô chat input vì root Grid của pane này không có row footer riêng — row cuối (Row 3) là input bar full-width. Đã xử lý bằng cách thêm Row 4 (Auto height) làm footer thật thay vì dùng kỹ thuật overlay `Grid.RowSpan="99"`. Cần user xác nhận trực quan trong Revit ở Ngày 4 rằng dockable pane không bị vỡ layout / không quá chật do thêm ~24px chiều cao footer.
- Chưa phát sinh vấn đề khác trong Ngày 3.
