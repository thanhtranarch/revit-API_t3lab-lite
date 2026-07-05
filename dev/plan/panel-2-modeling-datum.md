# Panel 2 — Modeling & Datum (Ngày 6–7)

> 14 tool — panel lớn nhất, chia 2 ngày. Model test: file có room, wall (thẳng + cong), floor, door, level, grid; chuẩn bị sẵn 1 file DWG mẫu, 1 ảnh PNG đen trắng, 1 file point (nếu có).

---

## Ngày 6 — Nhóm Create (7 tool)

### Tool 1/7 — CADToElements
Chain: launcher → `CADToElementsDialog.py` (3143+ loc) → `CADToElements.xaml` + 4 XAML con (CADtoBeam / CadtoFloor / CadtoFloorLayerItem / CadtoWall)

- [x] Mở cửa sổ chính + từng flow con (Wall / Floor / Beam) không lỗi XAML — UI đã sửa lại toàn bộ (sidebar rail + logo, search bar pill, layout card đồng bộ 3 tab, chiều cao cửa sổ) trong phiên debug 2026-07-03
- [x] Import DWG mẫu → convert Wall theo layer → wall tạo đúng vị trí/đơn vị — user xác nhận "tool ok r" sau khi sửa
- [x] Convert Floor (có layer item list — CadtoFloorLayerItem variant B load đúng)
- [x] Convert Beam
- [ ] DWG không có layer khớp → thông báo, không crash — chưa test riêng edge case này
- [ ] Ctrl+Z sau mỗi convert → revert 1 bước — chưa test riêng
- [x] Ghi chú: **Loạt bug functional phát hiện + sửa trong phiên 2026-07-03** (ngoài phạm vi checklist gốc, ghi lại để tham khảo):
  1. `extract_lines_from_cad()` gọi sai 4 tham số / mong 4 giá trị trả về trong khi hàm chỉ nhận 3 tham số, 1 giá trị — sửa lại đúng chuỗi `extract_lines_from_cad()` → `merge_collinear_lines()` → `find_parallel_pairs()`.
  2. `get_cad_layers_wall()`/`get_cad_layer_geometry_floor()` chỉ liệt kê layer có geometry phát hiện được (dò 1 lớp GeometryInstance, bỏ sót layer rỗng/lồng block) — đổi Wall sang liệt kê toàn bộ qua `Category.SubCategories` (giống Beam vốn đã đúng); Floor giữ quét geometry nhưng bổ sung layer rỗng vào danh sách.
  3. `extract_lines_from_cad()` áp `Transform` 2 lần (double-transform) cho block không lồng — gây tường/dầm bị lệch vị trí khi CAD import có transform khác Identity (xoay/lệch theo toạ độ site) — bỏ áp transform thủ công, chỉ đệ quy gọi `GetInstanceGeometry()`.
  4. `create_walls_auto()` không kiểm tra status trả về của `Transaction.Commit()` — nếu Revit tự rollback mà không throw exception, hàm vẫn báo "Created: N" giả — đã thêm kiểm tra status, báo trung thực.
  5. Beam: `DB.BuiltInParameter.STRUCTURAL_BEAM_Z_OFFSET_VALUE` không tồn tại (tên enum sai) → crash ngay dầm đầu tiên, rollback toàn bộ, 0 dầm được tạo — sửa thành `Z_OFFSET_VALUE` (đúng tên chuẩn Revit API), sửa ở cả bản mới và class legacy CadtoBeam (dòng ~2198).
  6. Thêm tính năng mới **Create Mode: Wall/Beam thật ↔ Part (DirectShape)** cho cả Wall và Beam (trước đó chỉ Floor có), dùng chung `create_part_from_loop()` + helper mới `build_rect_loop_from_centerline()`.
  7. Beam layer picker đổi từ ComboBox chọn 1 layer → bảng checklist đa chọn (Select All/Clear/search) giống Wall/Floor, `_run_beam()` gộp kết quả nhiều layer vào 1 dialog.
  8. UI: xoá hàng thống kê Floor (4 ô Total/Selected/Closed Loops/Created — không dùng), xoá box "Tip" thừa ở Beam, đồng bộ `MinHeight` card settings giữa 3 tab, tắt scrollbar khu vực content tổng (chỉ còn scrollbar trong bảng CAD Layers).
  - **Còn nợ, chưa xác nhận trong Revit**: Ctrl+Z revert 1 bước; edge case DWG không có layer khớp; Floor's `get_cad_layer_geometry_floor()` vẫn có bug đệ quy 1-lớp giống Wall từng có (đã tách thành task riêng `task_5c67d3b1`, chưa làm).

### Tool 2/7 — DoorThreshold
Chain: self-contained (356 loc, **3 bare except**) → `DoorThreshold.xaml`

- [x] Trước khi test: tạm thêm traceback log vào 3 nhánh `except:` để lỗi không bị nuốt — thay bằng `logger.error` + gom `error_messages` hiện trực tiếp trong TaskDialog kết quả (không cần mở log riêng)
- [x] Door trên wall thường → threshold tạo đúng — user xác nhận "đã xong" 2026-07-04 sau 2 vòng fix (xem Ghi chú)
- [ ] Door trên curtain wall / wall nghiêng → hành vi hợp lý — chưa test
- [ ] Thiếu family threshold trong model → thông báo rõ ràng — chưa test
- [x] Ghi chú: **Bug fix 2026-07-04** (phát hiện qua smoke test user, không nằm trong checklist gốc):
  1. **Sai kích thước ban đầu**: threshold không khớp đúng bề rộng cửa/độ dày tường khi tường có nhiều lớp (compound) hoặc bị tường khác join vào — do dùng `wall.Width` (giá trị nominal cố định của loại tường) thay vì đo thực tế tại vị trí cửa. Sửa: thêm `_get_wall_thickness_at_point()` — chiếu điểm cửa lên mặt ngoài/trong của tường (`HostObjectUtils.GetSideFaces` + `Face.Project`), đo khoảng cách thật tại đúng vị trí đó, fallback về `wall.Width` nếu không lấy được mặt (ví dụ Stacked Wall/Curtain Wall).
  2. **UI**: theo yêu cầu user, bỏ hẳn 2 ô nhập tay "Width Extension (mm)" / "Depth Extension (mm)" — vì kích thước giờ tự động khớp chính xác, không cần chỉnh tay. Chỉ còn "Height Offset (mm)".
  3. **Bug thứ 2 lộ ra sau khi test thật**: `Successfully created 0 thresholds, 4 errors` — nguyên nhân là `BuiltInParameter.DOOR_WIDTH` và `Symbol.FAMILY_WIDTH_PARAM` đều trả về `0` cho family cửa đang dùng (width không map vào 2 param chuẩn này). Sửa: thêm `_get_door_width_ft()` — thử `DOOR_WIDTH` → `FAMILY_WIDTH_PARAM` → `LookupParameter("Width")` (instance rồi symbol) → cuối cùng fallback đo trực tiếp bounding box của instance chiếu lên trục `HandOrientation` (đúng bất kể family map tham số kiểu gì).
  4. User xác nhận kết quả đúng sau khi áp cả 2 fix (width + thickness) — nhưng **mới test 1 kịch bản** (cửa trên tường thường/compound). Curtain wall, wall nghiêng, model thiếu family threshold chưa qua kiểm.

### Tool 3/7 — ImageToDrafting
Chain: self-contained (1585 loc) → `ImageToDrafting.xaml` · **compile C# Vectorizer lúc runtime**

- [ ] Ảnh PNG đen trắng nhỏ (~500px) → mode Centerline → drafting lines tạo ra
- [ ] Mode Outline (không cần compile C#) → hoạt động
- [ ] Xác nhận compile C# thành công trên máy triển khai (không chỉ máy dev) — nếu fail phải ra alert "Failed to compile C# vectorizer" chứ không crash
- [ ] Ảnh màu / ảnh lớn → threshold tự động hợp lý, không treo Revit
- [x] Ghi chú: **UI fix 2026-07-03** (chưa test functional ở trên, chỉ sửa UI): dropdown "Tracing Mode"/"Threshold Mode" thiếu style (`ComboBox` trần) → gán `T3ComboBoxSmall`; làm lại khu vực Drag & Drop (border xanh đặc `#3B82F6` → nét đứt xám chuẩn drop-zone, icon emoji 📎 → vector Path); xoá icon "user profile" thừa ở title bar; rút gọn subtitle.

### Tool 4/7 — Text to Element
Chain: launcher → `TextToElementDialog.py` (466 loc) → `TextToElement.xaml`

- [ ] Text note hợp lệ → element tạo đúng
- [ ] Text rác / rỗng → validate, không crash
- [x] Ghi chú: **UI fix 2026-07-03** (chưa test functional ở trên, chỉ sửa UI): nút min/max/close bị sai vị trí (`VerticalAlignment="Top"` không margin → dính sát góc bo tròn cửa sổ), sửa về `Center` + `Margin="0,0,16,0"` đúng chuẩn canonical; nội dung vượt quá chiều cao cửa sổ (620px, cần ~870px) luôn phải cuộn → tăng `Height` 620→880, `MinHeight` 520→750, ẩn scrollbar (`Hidden`, vẫn cuộn được bằng chuột).

### Tool 5/7 — PointCloud
Chain: self-contained (1502 loc, 6 transaction) → `PointCloud.xaml`

- [ ] File điểm nhỏ hợp lệ → import/xử lý đúng
- [ ] File sai format → thông báo lỗi rõ
- [ ] Cancel giữa chừng → transaction rollback sạch (kiểm tra Undo stack không có rác)
- [ ] Ghi chú:

### Tool 6/7 — RoomToFloor
Chain: self-contained (383 loc) → `RoomToFloor.xaml`

- [ ] Room chữ nhật → floor đúng boundary, offset mm đúng
- [ ] Room có cột giữa (island) → floor có opening (test cả Revit <2022 nếu dùng: 2 nhánh code theo REVIT_VERSION)
- [ ] Room chưa placed / area = 0 → skip đếm vào error_count, không crash
- [ ] Toàn bộ batch = 1 TransactionGroup → Ctrl+Z 1 lần revert hết
- [ ] Ghi chú:

### Tool 7/7 — Tile Layout
Chain: self-contained (2450 loc) → `TileLayout.xaml`

- [ ] Layout tile trên floor nhỏ → kết quả đúng kích thước/góc
- [ ] Đổi origin/rotation → cập nhật đúng
- [ ] ⚠️ Script không gọi Transaction trực tiếp — xác định wrapper nào ghi model (revit.Transaction? helper?) và Ctrl+Z hoạt động đúng
- [x] Ghi chú: **Bug path + UX fix 2026-07-03**:
  1. **Tool không mở được** — `EXT_DIR` trong `script.py` chỉ đi lên 3 cấp thư mục (`os.path.dirname` x3) thay vì 4 cấp cần thiết (`T3Lab.extension/T3Lab.tab/Modeling & Datum.panel/Create.stack/Tile Layout.pushbutton/`), dừng ở `T3Lab.tab` thay vì `T3Lab.extension` → `IOError: Could not find a part of the path ...T3Lab.tab\lib\GUI\Tools\TileLayout.xaml`. Đã sửa thành x4. **Quét toàn bộ codebase phát hiện thêm 3 tool khác cũng bị lỗi tương tự (chiều ngược lại — thừa 1 cấp): MCPControl, ManaSheets, ManaViews — đã sửa cả 3 (xem Support.panel/Views & Sheets.panel).**
  2. **UX**: trước đây `run()` luôn ép pick floor (hoặc bắt buộc có sẵn selection) ngay khi mở tool, chưa mở được wizard nếu chưa pick gì. Tách thành `_initial_floors()` (mở tool lần đầu → tự động liệt kê toàn bộ Floor trong model qua `FilteredElementCollector`, không cần pick) và `_pick_floors_from_revit()` (chỉ dùng cho nút "Pick Floors in Model" — vẫn pick thủ công như cũ).
  3. **UI**: xoá icon "user profile" thừa ở title bar, chỉnh margin nút min/max/close về 16px chuẩn.
  - **Còn nợ, chưa test functional**: layout tile kích thước/góc, đổi origin/rotation, Ctrl+Z qua transaction wrapper.
  4. **UI (2026-07-03, theo yêu cầu user)**: chuyển step-progress từ thanh ngang (Row 1, 3 vòng tròn số) sang **sidebar trái 66px** — theo đúng pattern sidebar rail của BatchOut (`ExportManager.xaml`, UI-locked, chỉ dùng làm tham chiếu đọc) và CADToElements: logo 42×42 trên cùng, 3 icon `ToggleButton` (`T3SidebarToggle`, đã có sẵn trong shared style) xếp dọc thay 3 vòng tròn ngang. Xoá style `StepCircle`/`StepLabel` không còn dùng. Code-behind: thêm `self._max_step` theo dõi bước xa nhất đã tới, `step_nav_clicked()` cho phép bấm quay lại bước đã hoàn thành (không cho nhảy vượt bước chưa tới), `_refresh_step_ui()` viết lại để đồng bộ `IsChecked`/`IsEnabled` 3 toggle thay vì set màu tay theo circle cũ. Đã pass `audit_tools`/`audit_ui`/`sync_wpf_styles --check` + xml well-formed + python ast parse. **Chưa test trực quan trong Revit.**

### Chốt ngày 6
- [ ] Cập nhật bảng trạng thái + README · Commit nếu có fix

---

## Ngày 7 — Nhóm Adjust / Family / Datum (7 tool)

### Tool 8/14 — PropertyLine
Chain: launcher → `PropertyLineDialog.py` (1727 loc) → `PropertyLine.xaml`

- [ ] Nhập bảng toạ độ hợp lệ → property line tạo đúng
- [ ] Sai đơn vị / toạ độ không khép kín → validate rõ ràng
- [ ] Ghi chú:

### Tool 9/14 — FamiGen
Chain: launcher → `FamiGenDialog.py` (2360 loc) → `FamiGen.xaml` · tạo family document riêng

- [ ] Gen 1 family đơn giản → load vào project OK
- [ ] Huỷ giữa chừng → family document tạm được **đóng sạch** (kiểm tra không còn doc mở ngầm)
- [ ] Template family thiếu → thông báo
- [x] Ghi chú: **Cải tiến sub-mode "From JSON" 2026-07-04** (ngoài checklist smoke-test gốc, chỉ động vào `SYSTEM_PROMPT.md` — parser `FamiGenDialog.py` không đổi):
  - **WS2 de-bias**: gỡ toàn bộ ví dụ nghiêng về đèn (Lotus lamp → Ceiling Medallion; wall-lamp arm → generic tube; part-inventory wall lantern → wall-mounted faucet). Prompt phải dùng được cho **mọi** loại family theo yêu cầu user.
  - **WS3**: thêm recipe ống gấp khúc/vuông góc = các đoạn Extrusion-of-Circle theo trục chồng ~2mm ở khuỷu (thay Sweep cong dễ gãy vì parser chỉ hỗ trợ mặt phẳng X/Y/Z).
  - **+7 ví dụ** (E8–E14) → tổng **14 ví dụ**, đủ 4 form type + 6 curve type, trải Furniture/Generic/Casework/Door. Đã validate 14/14 block JSON hợp lệ.
  - Roadmap + backlog (WS4 cảnh báo part rời, WS5 mặt phẳng bất kỳ) + **kế hoạch tách prompt theo category**: `dev/plan/famigen-from-json-roadmap.md`.
  - **WS1 + WS6 (2026-07-04, cùng ngày)**: user test lại → kết quả **tệ hơn** (mọi part dựng được nhưng rời hoàn toàn — riser lơ lửng, arm không chạm chao). Sửa prompt: bắt buộc 1 anchor part + key `"attaches_to"` từng part + key root `"_plan"` (bảng inventory/height/connection — parser bỏ qua key lạ, chỉ để ép AI suy luận); kiểm tra overlap 3 trục bằng số; cấm part mồ côi; cấm copy tọa độ literal từ ví dụ (nghi bleed-through tọa độ Example 7). Bài học: **build được ≠ ráp được**.
  - **Chia prompt theo category (2026-07-04)**: thêm combo "Family type" ở action bar tab From JSON (`json_category_combo`) — chọn loại **trước**, bấm Copy Prompt sẽ copy `core chung + overlay của loại đó`. Core = `SYSTEM_PROMPT.md`; 13 overlay ở `FamiGen.pushbutton/prompts/<slug>.md` (mỗi category 1 file: inventory/convention/dimensions/pitfalls riêng). Code: `_PROMPTS_DIR`, `_init_json_panel()`, `_category_slug()`/`_overlay_path()`, `copy_prompt_clicked()` ghép core+overlay (thiếu overlay → fallback core-only). Đã pass audit_tools/audit_ui/sync + XAML well-formed + 13/13 overlay map đúng.
  - **Prompt v2 tự chứa + soft-form + case library (2026-07-05)**: `SYSTEM_PROMPT.md` đã xóa — còn **7 file prompt tự chứa** trong `prompts/` (bỏ door/window/columns/electrical/site/entourage); combo chỉ hiện category có file. Core chung (byte-identical cả 7 file, sửa bằng script + verify equality) thêm section **SOFT-FORM RECIPES** — 7 công thức tạo độ phồng (capsule extrusion, channel-tufting đè 8–12mm, bo góc Arc3P, vòm/cầu Revolution, Blend thuôn mềm, mép cuộn, Spline organic) + khái niệm **bulge depth** + key `_plan.shapes`. furniture.md thêm Upholstery playbook + ví dụ **sofa channel-tufted 11 part** (verify tĩnh pass: loop kín, đúng plane, overlap 3 trục ≥1mm). Mỗi category thêm "Soft-form applications" + **Object case library** ~9–11 case thực tế (bộ phận → form/recipe → mm → anchor). Sửa cảnh báo lỗi thời trong lighting (Cylinder chéo trục = parser dựng bằng NewRevolution quanh trục của nó, không phải sweep engine). Chi tiết + checklist test: `dev/plan/famigen-from-json-roadmap.md`.
  - **Chưa test functional trong Revit** (sofa example 11 part, Cylinder chéo trục, 2–3 case end-to-end photo→AI→JSON; cộng checklist gốc: gen family, chọn loại → copy → gen, đóng doc khi huỷ, thiếu template) — cần user reload pyRevit rồi test lại.

### Tool 10/14 — ManaFami
Chain: launcher → `ManaFamiDialog.py` (1428 loc) → `ManaFami.xaml`

- [ ] Batch rename family → đúng, Ctrl+Z OK
- [ ] Family đang in-use / editable=false (workshared) → xử lý đẹp
- [ ] Purge family không dùng (nếu có chức năng) → xác nhận trước khi xoá
- [x] Ghi chú: **UI + bug fix 2026-07-03 (theo báo cáo user: field trống/rỗng ở tab Management)**: xoá nút "Refresh DataGrid" (`btn_refresh`, icon-only, tab Management toolbar) theo yêu cầu user — không cần thiết vì `window_loaded()` đã tự gọi `_refresh_data()` khi mở cửa sổ; xoá luôn handler chết `refresh_click`. Đồng thời phát hiện & sửa lỗi tiềm ẩn qua đọc code: `window_loaded()` trước đây gói TOÀN BỘ init (Loader tab + Management tab) trong 1 khối `try/except` duy nhất — nếu bước scan folder của Loader tab lỗi, `_load_worksets()`/`_refresh_data()` của Management tab **không bao giờ chạy**, khiến DataGrid trống mà không có thông báo (chỉ log ngầm). Đã tách thành 2 khối try/except độc lập + hiện lỗi ra `status_text` thay vì nuốt âm thầm. **Chưa xác nhận đây có phải nguyên nhân gốc của "field trống" hay không** — cần user mở lại ManaFami trong Revit, nếu Management tab vẫn trống thì đọc dòng status bar (giờ sẽ hiện lỗi cụ thể thay vì im lặng) và báo lại. **UI (2026-07-03, theo screenshot user)**: search bar tab Management (`tb_search`) trước đó là icon Segoe MDL2 tách rời + textbox `T3InputFieldSmall` riêng (radius 8) → nhìn lệch/méo thành oval to. Đã dựng lại thành 1 pill Border duy nhất (White, border `#DCDCE0`, CornerRadius = nửa chiều cao) chứa icon kính lúp vector + TextBox không viền bên trong, đúng theo mẫu "Search Bar Simulator" của `UIStandardShowcase.xaml`. Search bar tab Loader (`txt_search`) chưa đụng tới (không nằm trong screenshot user gửi).

### Tool 11/14 — Wall_Adjust Base
Chain: self-contained (311 loc, **4 bare except**), không XAML riêng

- [ ] Tạm thêm traceback log vào 4 nhánh except
- [ ] Wall thường → adjust base đúng offset
- [ ] Wall bị attach base/top → hành vi đúng hoặc thông báo
- [ ] Ghi chú:

### Tool 12/14 — SplitElements
Chain: launcher → `SplitElementsDialog.py` (61 loc) → `SplitElements.xaml` (logic có thể nằm ở Snippets)

- [ ] Split wall theo level → đúng số đoạn, join/param giữ hợp lý
- [ ] Element không hỗ trợ split → thông báo
- [ ] Ghi chú:

### Tool 13/14 — AutoJoin
Chain: self-contained (708 loc, 1 transaction) → `AutoJoin.xaml`

- [ ] Model nhỏ nhiều wall-floor giao nhau → join hàng loạt đúng
- [ ] Cặp element không join được (geometry fail) → skip + báo số lượng, không dừng cả batch
- [ ] Đo thời gian trên model vừa (~500 element) — có progress/không treo UI
- [ ] Ghi chú:

### Tool 14/14 — Wall Cut Profile ⚠️ rủi ro cao nhất panel
Chain: self-contained (1624 loc, **37 bare except**, 8 transaction + 2 TransactionGroup)

- [ ] Trước test: chèn traceback log vào các nhánh except quan trọng (đặc biệt quanh L705, L1008, L1016 — vùng commit/rollback)
- [ ] Wall thẳng + opening đơn giản → profile cut đúng
- [ ] Wall cong → hoạt động hoặc thông báo không hỗ trợ
- [ ] Profile đè lên nhau → không tạo geometry hỏng
- [ ] Ctrl+Z 1 lần revert cả group (TransactionGroup Assimilate)
- [ ] Test cả flow "Place Opening Families" (L1212+): symbol chưa active → tự activate
- [ ] Ghi chú:

### Chốt ngày 7
- [ ] Cập nhật bảng + README · Gỡ các traceback log tạm trước khi commit
- [ ] Commit: `fix(panel-modeling): <mô tả>`

## Bảng trạng thái Panel 2

| # | Tool | Ngày | Trạng thái | Ghi chú |
|---|------|------|-----------|---------|
| 1 | CADToElements | 6 | ✅ | Wall/Floor/Beam user xác nhận OK (2026-07-03) · user xác nhận chung 2026-07-05: hoạt động tốt |
| 2 | DoorThreshold | 6 | ✅ | Width=0 + thickness sai (compound/joined wall) sửa xong, user xác nhận OK 2026-07-04 · user xác nhận chung 2026-07-05: hoạt động tốt (curtain wall/wall nghiêng vẫn ngoài scope) |
| 3 | ImageToDrafting | 6 | ✅ | UI sửa xong 2026-07-03 · user xác nhận chung 2026-07-05: hoạt động tốt |
| 4 | Text to Element | 6 | ✅ | UI sửa xong 2026-07-03 · user xác nhận chung 2026-07-05: hoạt động tốt |
| 5 | PointCloud | 6 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| 6 | RoomToFloor | 6 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| 7 | Tile Layout | 6 | ✅ | Bug path + list sàn mặc định sửa 2026-07-03 · user xác nhận chung 2026-07-05: hoạt động tốt |
| 8 | PropertyLine | 7 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| 9 | FamiGen | 7 | ✅ | Tool hoạt động tốt (user xác nhận chung 2026-07-05). Prompt v2 2026-07-05 (7 file tự chứa + SOFT-FORM RECIPES + case library + ví dụ sofa 11 part) — 3 probe chi tiết (sofa smoke test · Cylinder chéo · case end-to-end) vẫn theo dõi ở `famigen-from-json-roadmap.md` |
| 10 | ManaFami | 7 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| 11 | Wall_Adjust Base | 7 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| 12 | SplitElements | 7 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| 13 | AutoJoin | 7 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |
| 14 | Wall Cut Profile | 7 | ✅ | User xác nhận chung 2026-07-05: hoạt động tốt |

> 2026-07-05 — user xác nhận tổng quát: **"các tool hiện tại đều đã hoạt động tốt"**. Các mục
> checklist chi tiết chưa có báo cáo riêng từng mục — coi là pass theo xác nhận chung; mở lại
> nếu phát hiện lỗi khi dùng thực tế.

## Phát sinh

_(ghi lỗi mới tại đây)_
