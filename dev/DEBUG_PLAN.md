# T3Lab — Review & Kế hoạch Debug từng Tool

> Cập nhật: 2026-07-05 (chốt GĐ4) · Phạm vi: toàn bộ 41 pushbutton trong `T3Lab.extension/T3Lab.tab/`
> Công cụ đi kèm: `python3 dev/audit_tools.py` (audit tĩnh) · `python3 dev/audit_ui.py` (audit UI Lumina) · `python3 dev/sync_wpf_styles.py --check` (đồng bộ style)
>
> **📅 Kế hoạch thực hiện theo ngày (14 ngày, chia theo panel): xem [`dev/plan/README.md`](plan/README.md)** — file này là bản ghi review gốc; checklist thực hiện hằng ngày nằm trong `dev/plan/`.
>
> **✅ TRẠNG THÁI CUỐI — 2026-07-05 (GĐ4)**: Kế hoạch 12 ngày hoàn tất — GĐ1/GĐ2 xong từ
> 2026-07-03, GĐ3 (41/41 tool) đóng theo xác nhận chung của user 2026-07-05 *"các tool hiện
> tại đều đã hoạt động tốt"*, GĐ4 chốt cùng ngày. 4 roadmap độc lập đã đóng & xóa file (kết
> cục ghi ở `dev/plan/README.md`). Regression cuối: `audit_tools` clean · `audit_ui` 0/54 ·
> `sync_wpf_styles --check` 53/53 0 lệch. **Issue còn lại & review GĐ4: xem mục 7.**

---

## 1. Tổng quan kiến trúc

Codebase có **2 kiểu tool**:

| Kiểu | Mô tả | Số lượng |
|------|-------|----------|
| **Launcher** | `script.py` mỏng (15–56 dòng) → gọi `lib/GUI/<Tool>Dialog.py` (logic + WPF) → load XAML trong `lib/GUI/Tools/` | ~20 tool |
| **Self-contained** | Toàn bộ logic nằm trong `script.py` của pushbutton (122–3764 dòng), XAML vẫn ở `lib/GUI/Tools/` | ~18 tool |
| **Không UI / urlbutton** | PDF import, CloudLinks... | 3 |

### Kết quả audit tĩnh đã chạy (2026-07-02)

| Check | Kết quả |
|-------|---------|
| Syntax IronPython 2.7 (f-string, annotation, nonlocal) trên toàn bộ lib + tab | ✅ Sạch (trừ `lib/core/bridge.py` — chủ đích chạy CPython 3) |
| Chuỗi launcher → Dialog module → entry function → XAML | ✅ 19/19 launcher resolve đúng |
| Style block chia sẻ (`sync_wpf_styles.py --check`) | ✅ 76 file, 0 lệch |
| Transaction: `Start()` có `Commit()`/`Assimilate()` tương ứng | ✅ Không phát hiện transaction treo |
| Handler XAML ↔ method Python | ⚠️ 2 file lỗi thật (mục 2) |
| XAML mồ côi (không Python nào tham chiếu) | ⚠️ 12 file (mục 2) |

---

## 2. Phát hiện (Findings)

### ✅ F1 — `FindReplace.xaml` thiếu handler `button_run` (ĐÃ SỬA — GĐ1 Ngày 1)
- `lib/GUI/Tools/FindReplace.xaml:1181` khai báo `Click="button_run"` nhưng `lib/GUI/FindReplace.py` không có method này (WPF_Base cũng không).
- Vì class dùng `wpf.LoadComponent` → **XamlParseException ngay khi mở cửa sổ** nếu được gọi.
- Hiện không tool nào khởi tạo `FindReplace`, nhưng `lib/GUI/forms.py` vẫn import class này — bất kỳ tool mới nào dùng lại sẽ crash.
- **Hành động**: thêm `def button_run(self, sender, e)` hoặc xoá module nếu xác nhận không dùng.

### ✅ F2 — `CreateFromRooms.xaml` thiếu 4 handler (ĐÃ XỬ LÝ — GĐ1 Ngày 1)
- Thiếu: `NumberValidationTextBox`, `UIe_ItemChecked`, `button_run`, `text_filter_updated` trong `lib/GUI/CreateFromRooms.py`.
- Không nơi nào import `CreateFromRooms` → dead code, nhưng sẽ crash khi mở nếu được dùng lại.
- **Hành động**: xoá hoặc hoàn thiện handler trước khi tái sử dụng.

### ✅ F3 — UIShowcase phụ thuộc cấu trúc repo (ĐÃ SỬA — GĐ1 Ngày 1)
- `lib/GUI/UIShowcaseDialog.py:33` load XAML từ `[repo]/.claude/standard/UIStandardShowcase.xaml` — **ngoài** thư mục `T3Lab.extension`.
- Nếu deploy extension rời khỏi repo (copy `T3Lab.extension` vào `%APPDATA%\pyRevit\Extensions`), tool hỏng.
- **Hành động**: copy XAML vào `lib/GUI/Tools/` hoặc fallback path.

### ✅ F4 — 12 XAML mồ côi trong `lib/GUI/Tools/` (ĐÃ DỌN — GĐ1 Ngày 2, cùng F9)
`AutoClick`, `BulkFamilyExport`, `ConvertGridline`, `ConvertLevel`, `ExportManagerTest`(*), `InPlaceModelMain`, `JSONtoFamily`, `LocationManager_backup`, `ModelChecker`, `ParameterLoaderCatPicker/Main/Mapper`, `TransferPara` — không file Python nào tham chiếu.
- (*) `ExportManagerTest.xaml` đang được CLAUDE.md dẫn làm ví dụ wizard nav — giữ lại nhưng đánh dấu.
- `DatumManagerDialog.py` (dead code) còn trỏ tới các pushbutton `Datum.pulldown/...` **không tồn tại**.
- **Hành động**: dọn dẹp hoặc chuyển vào thư mục `_archive/`.

### 🟡 F5 — Bare `except:` nuốt lỗi hàng loạt (GIỮ NGUYÊN — tech debt, xem mục 7)
| File | Số bare except |
|------|----------------|
| `Wall Cut Profile/script.py` | 37 |
| `BatchOut/script.py` | 30 |
| `Wall_Adjust Base/script.py` | 4 |
| `DoorThreshold/script.py` | 3 |
| `ManaTabs/script.py` | 2 |

Khi debug các tool này, lỗi thật sẽ bị nuốt im lặng → phải tạm thêm logger vào các nhánh except trước khi test.

### 🟡 F6 — BatchOut: module "Intelligence" degrade im lặng (GIỮ NGUYÊN — tech debt, xem mục 7)
- `BatchOut/script.py:62-67`: import `api_learner`/`api_updater` bọc trong bare except → nếu fail, `HAS_API_LEARNER=False` không có log. Cần in cảnh báo khi debug.

### ✅ F7 — Tài liệu ui-design-standard lỗi thời về vị trí copyright (ĐÃ SỬA DOC 2026-07-02)
- Kiểm tra thực tế: **74/76 XAML và cả file chuẩn `UIStandardShowcase.xaml` đã đặt copyright ở footer (status bar) BÊN TRÁI** từ trước (copyright → divider 1px `#DCDCE0` → status text). Tài liệu `ui-design-standard.md` cũ bắt overlay góc phải-dưới (`Grid.RowSpan="99"`) mới là thứ sai so với codebase.
- **Đã cập nhật** `.claude/rules/ui-design-standard.md`: chuẩn chính thức = copyright ở footer, left-most element, theo snippet của `UIStandardShowcase.xaml`; Variant B không có footer dùng fallback overlay **trái**-dưới; item-template XAML (`CadtoFloorLayerItem.xaml`) được miễn trừ. `dev/audit_ui.py` đã sửa theo rule mới.
- Việc còn lại (2 outlier, đều thuộc T3LabAssistant — gộp xử lý với F8 trong GĐ2 Ngày 3):
  - `T3LabAssistant.xaml:~2109` — copyright đang căn **giữa** dưới ô chat input
  - `AssistantPane.xaml:~1357` — copyright theo chuẩn cũ (overlay góc **phải**-dưới); đây là file duy nhất từng theo doc cũ
- Các hạng mục UI khác đều đạt: palette Lumina (trừ F8), font, WindowChrome, glyph, shared styles, không dot-notation. KHÔNG sửa 2 file UI-locked (`DWGManagement.xaml`, `ExportManager.xaml`).

### ✅ F8 — T3LabAssistant.xaml sót palette Terra cũ (ĐÃ SỬA — GĐ2 Ngày 3)
- `#15803D` (Terra success) ×1, `#166534` (Terra success hover) ×1 → migrate sang `#10B981` / `#059669`.

### ✅ F9 — Mở rộng F4: thêm dead code phát hiện sau (ĐÃ DỌN — GĐ1 Ngày 2)
- XAML mồ côi thêm: `InPlaceModelRename.xaml`, `InPlaceModelSettings.xaml`, `ParaSync.xaml` (tổng 15 file).
- 6 dialog trung gian **không pushbutton nào gọi tới**: `DatumManagerDialog.py` (+`DatumManager.xaml`), `FamilyBatchCreatorDialog.py`, `OpeningAssignValuesDialog.py`, `SheetHubDialog.py`, `ViewHubDialog.py`, `ViewTemplateDialog.py` (+2 XAML).
- Danh sách archive đầy đủ: `dev/plan/phase-1-cleanup-fixes.md`.

### ✅ F11 — PDF import: mở tool không hiện thông tin + crash file (PHÁT HIỆN & SỬA — 2026-07-05, user báo)
- Bug functional thật (không phải UI): mở tool thì grid "Target Views" kẹt overlay "Loading
  views…" (không view nào hiện) và crash/đóng file Revit. Phơi ra rằng xác nhận chung "các
  tool hoạt động tốt" (2026-07-05) chưa thực sự test functional PDF import.
- **Nguyên nhân kép** trong `PDFImportDialog.py`:
  1. *Không hiện thông tin* — refactor `0a2a886` hoãn nạp grid sang event `ContentRendered`,
     mà event này **không chạy đáng tin dưới modal `ShowDialog()`** trong Revit → `_load_views()`
     không chạy → `pnl_loading` không collapse → kẹt "Loading views…".
  2. *Crash* — bind `ItemsSource` + mutate `ObservableCollection` đặt trong `Loaded`/`ContentRendered`
     (SAU khi cửa sổ render); mutate collection bound DataGrid non-virtualized giữa layout có thể
     **hard-crash host** (lỗi tầng WPF/native, không phải exception Python bắt được).
- **Fix**: dời tạo OC + bind + `_load_views()` về `__init__`, chạy **đồng bộ trước `ShowDialog()`**
  (cửa sổ chưa render) — đúng pattern kiểm chứng ở `ManaViews`/`ManaContains`/`ManaSheets`. Xóa
  `_on_loaded`/`_on_content_rendered`/`_views_loaded`; guard `mode_changed` giữ nguyên. Static
  clean (`ast.parse` + 3 audit). **Chưa smoke test Revit** — cần user reload pyRevit + mở lại.
- **Bài học**: xác nhận "tất cả tool OK" theo lô KHÔNG thay được smoke test functional từng tool;
  các tool chỉ mới sửa UI (chưa test chức năng) vẫn có thể còn bug mở-cửa-sổ như thế này.

### ✅ F10 — 2 điểm ghi JSON `ensure_ascii=False` trên file bytes-mode (PHÁT HIỆN & SỬA — review GĐ4, 2026-07-05)
- `lib/Intelligence/t3lab_assistant.py::save_learned_patterns` và `lib/config/user_profile.py::save`
  còn dùng `open(path, 'w')` + `json.dump(..., ensure_ascii=False)` — đúng failure class đã biết
  (từng làm hỏng chat history / tool_registry.json): IronPython 2.7 ném `UnicodeEncodeError` giữa
  chừng khi dữ liệu có tiếng Việt → file bị truncate 0 byte, lỗi bị `except: pass` / `return False`
  nuốt im lặng. `user_profile.py` **chắc chắn dính** vì mặc định `_DEFAULT_NAME = u"Thạnh"`.
- Đã sửa cả 2 theo pattern chuẩn của codebase: `json.dumps(..., ensure_ascii=True)` → decode →
  `io.open(path, 'w', encoding='utf-8')` ghi một lần. Các điểm ghi khác đã an toàn từ trước
  (history/registry trong `T3LabAssistant/script.py`, `tool_discovery.py` dùng cùng pattern;
  `IFCSGDialog.py` dùng `codecs.open`; body HTTP trong `llm_provider.py` encode UTF-8 đúng).

---

## 3. Quy trình smoke-test chuẩn (áp dụng cho mọi tool)

Chạy trong Revit (2022+ khuyến nghị, có model test nhỏ):

1. **Mở tool** từ ribbon → cửa sổ hiện, không XamlParseException (bắt lỗi handler/XAML).
2. **Chrome window**: kéo title bar, Minimize / Maximize / Close hoạt động; resize grip ở góc.
3. **Happy path**: chạy chức năng chính với input hợp lệ → kiểm tra kết quả trong model.
4. **Undo test**: 1 lần Ctrl+Z = revert đúng 1 thao tác tool (transaction/TransactionGroup gom đúng).
5. **Empty/edge case**: không chọn gì, model rỗng, huỷ giữa chừng → phải có thông báo thân thiện, không stacktrace đỏ.
6. **Log check**: bật pyRevit debug (`CTRL+click` nút hoặc `pyrevit env`) xem exception bị nuốt.
7. **Doc khác nhau**: chạy trên file workshared (nếu tool đụng workset/sheet) và file thường.

Ghi kết quả vào bảng mục 4 (cột Trạng thái: ⬜ chưa test / ✅ pass / ❌ fail + issue).

---

## 4. Kế hoạch debug từng tool (theo panel)

### 4.1 Annotation & Select.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 1 | AutoDimension | launcher → `AutoDimensionDialog.py` | 1721 | Logic dim theo reference dễ vỡ giữa các version Revit | Tạo dim trên wall/grid ở plan + section; test view không hỗ trợ | ✅ |
| 2 | ManaDWG | self + `DWGManagement.xaml` (UI-locked) | 432 | Xoá/relink DWG import vs link | List đúng DWG import & link; xoá 1 import; Undo | ✅ |
| 3 | ManaSelect | launcher → `ManaSelectDialog.py` | 660 | Filter selection theo category/param | Chọn hỗn hợp element rồi lọc; selection rỗng | ✅ |
| 4 | ManaAnno | launcher → `ManaAnnoDialog.py` | 1416 | Thao tác hàng loạt trên text/tag/dim | Batch đổi type annotation; view template ảnh hưởng | ✅ |

### 4.2 Modeling & Datum.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 5 | CADToElements | launcher → `CADToElementsDialog.py` (5 XAML con) | 3143 | Đọc geometry DWG; 3 flow (Wall/Floor/Beam) | Import DWG mẫu → convert từng loại; DWG không có layer khớp | ✅ |
| 6 | DoorThreshold | self | 356 | 3 bare except; family threshold phải có sẵn | Cửa trên wall thường + curtain wall; thiếu family | ✅ |
| 7 | ImageToDrafting | self | 1585 | **Compile C# runtime (Vectorizer)** — phụ thuộc .NET compiler trong Revit | Ảnh PNG đen trắng nhỏ; test cả 2 mode (centerline/outline); máy không compile được | ✅ |
| 8 | Text to Element | launcher → `TextToElementDialog.py` | 466 | Parse text note thành element | Text hợp lệ + text rác | ✅ |
| 9 | PointCloud | self | 1502 | 6 transaction; file scan lớn | File điểm nhỏ; format sai; cancel giữa chừng | ✅ |
| 10 | RoomToFloor | self | 383 | Boundary phức tạp (room có island/arc) | Room chữ nhật, room có cột giữa (opening), room chưa placed | ✅ |
| 11 | Tile Layout | self | 2450 | Không dùng Transaction trực tiếp trong script (kiểm tra wrapper) — logic layout tính toán nhiều | Layout trên floor nhỏ; đổi origin/góc; Undo | ✅ |
| 12 | PropertyLine | launcher → `PropertyLineDialog.py` | 1727 | Tạo property line từ tọa độ/bảng | Nhập bảng tọa độ hợp lệ + sai đơn vị | ✅ |
| 13 | FamiGen | launcher → `FamiGenDialog.py` | 2360 | Tạo family qua Document khác (family doc) — dễ leak doc | Gen 1 family, load vào project; huỷ giữa chừng xem family doc có đóng | ✅ |
| 14 | ManaFami | launcher → `ManaFamiDialog.py` | 1428 | Rename/purge family hàng loạt | Batch rename; family in-use | ✅ |
| 15 | Wall_Adjust Base | self | 311 | 4 bare except | Wall thường + wall bị attach top/base | ✅ |
| 16 | SplitElements | launcher → `SplitElementsDialog.py` | 61 | Dialog rất mỏng — logic có thể ở Snippets | Split wall/line theo level | ✅ |
| 17 | AutoJoin | self | 708 | JoinGeometry hàng loạt → chậm/fail từng cặp | Model nhỏ nhiều wall-floor giao nhau; phần tử không join được | ✅ |
| 18 | Wall Cut Profile | self | 1624 | **37 bare except**; 8 transaction; edit sketch profile | Wall thẳng + wall cong; profile đè lên nhau; Undo group | ✅ |

### 4.3 Views & Sheets.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 19 | ManaSheets | launcher → `ManaSheetsDialog.py` | 583 | Batch sửa sheet | Rename/renumber; sheet đang mở | ✅ |
| 20 | SheetGen | self | 1262 | 7 transaction, 0 rollback — lỗi giữa chừng để lại state dở | Gen sheet từ template; titleblock thiếu; fail giữa batch | ✅ |
| 21 | ManaViews | launcher → `ManaViewsDialog.py` | 689 | Xoá/đổi view hàng loạt | View đang active; view trên sheet | ✅ |
| 22 | BatchOut | self + `ExportManager.xaml` (UI-locked) | 3764 | **30 bare except; F6 (Intelligence degrade im lặng)**; export PDF/DWG/NWC phụ thuộc máy | Export 2 sheet PDF + DWG; máy không có PDF printer; đường dẫn output không tồn tại; kiểm tra `HAS_API_LEARNER` | ✅ |

### 4.4 Data & IFC-SG.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 23 | IFC-SG | launcher → `IFCSGDialog.py` + `configs/` | 2236 | Mapping param IFC-SG từ config JSON | Config mặc định; config thiếu field; apply lên model mẫu | ✅ |
| 24 | Foundation Volume | self | 410 | Tính volume qua geometry boolean | Foundation giao nhau; element không có solid | ✅ |
| 25 | ManaContains | launcher → `ManaContainsDialog.py` | 2204 | Logic spatial containment nặng | Element trong/ngoài room; link model | ✅ |
| 26 | ManaSched | launcher → `ManaSchedDialog.py` | 1601 | Sửa schedule field/filter qua API | Schedule thường + key schedule | ✅ |
| 27 | ManaPara | launcher → `ManaParaDialog.py` | 2548 | Shared param file; binding category | Add/remove shared param; file .txt bị lock | ✅ |
| 28 | BCF Reader | self | 3041 | Parse ZIP/XML BCF; 4 transaction; viewpoint camera | BCF 2.1 mẫu; BCF hỏng; apply viewpoint | ✅ |

### 4.5 Standards & Settings.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 29 | ManaLoca | self + `ManaLoca.xaml` | 871 | Site/shared coordinates — thao tác nguy hiểm với model thật | **Chỉ test trên file copy**; đổi project base point; Undo | ✅ |
| 30 | ManaStyles | launcher → `ManaStylesDialog.py` | 2091 | Object styles/line weights toàn model | Đổi 1 style; template lock | ✅ |
| 31 | ManaWorkset | self | 609 | Chỉ chạy được trên file workshared | File thường (phải báo lỗi đẹp); tạo/rename workset | ✅ |
| 32 | ModelAuditor | launcher → `ModelAuditorDialog.py` | 1828 | Quét toàn model → chậm trên model lớn | Model nhỏ ra report đúng; model lớn đo thời gian | ✅ |

### 4.6 Support.panel & Standard.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 33 | MCPControl | launcher → `MCPControlDialog.py` | 348 | Start/stop server `core/server.py`; port 48884 | Bật/tắt server; port bị chiếm | ✅ |
| 34 | Feedback | self | 232 | Gửi feedback ra ngoài (network) | Không có mạng; nội dung unicode tiếng Việt | ✅ |
| 35 | PDF import | self, không UI | 33 | Phụ thuộc API import PDF theo version Revit | Revit 2022+ import 1 PDF | ✅ |
| 36 | T3LabAssistant | self | 3594 | **Tool lớn nhất**; import động (`imp`/`importlib` có guard); gọi `core.server`; UI chat | Mở UI; server chưa chạy; gửi 1 lệnh; kiểm tra fallback import trên IronPython (`importlib.util` không tồn tại → phải rơi vào nhánh `imp`) | ✅ |
| 37 | ManaTabs | self | 147 | Thao tác UI Revit qua Windows API — dễ vỡ theo version | Revit nhiều tab; 2 bare except | ✅ |
| 38 | Ribbon Names | self | 154 | Sửa text ribbon qua AdWindows.dll | Đổi tên; restart Revit xem persist | ✅ |
| 39 | BG Theme | self | 122 | Đổi background/theme qua UIApplication settings | Toggle theme; Revit dark mode | ✅ |
| 40 | UIShowcase | launcher → `UIShowcaseDialog.py` | 129 | **F3: XAML nằm ngoài extension** | Mở từ repo checkout OK; mở từ extension deploy rời → fail | ✅ |
| 41 | AutoWork | self + `AutoWork.xaml` | 353 | Không có Transaction trong script — kiểm tra có ghi model không | Chạy happy path; xem có cần transaction | ✅ |

---

## 5. Thứ tự ưu tiên (ĐÃ HOÀN THÀNH TOÀN BỘ — giữ làm tài liệu quy trình)

**Đợt 1 — Sửa lỗi đã xác nhận (không cần Revit):**
1. F1: thêm `button_run` vào `FindReplace.py` (hoặc xoá module + bỏ import trong `forms.py`).
2. F2: xử lý `CreateFromRooms` (xoá hoặc hoàn thiện).
3. F3: đưa `UIStandardShowcase.xaml` vào trong extension.
4. F4: dọn 12 XAML mồ côi + `DatumManagerDialog.py`.

**Đợt 2 — Smoke test tool rủi ro cao (cần Revit):**
BatchOut (22) → Wall Cut Profile (18) → T3LabAssistant (36) → BCF Reader (28) → CADToElements (5) → ImageToDrafting (7) → SheetGen (20) → ManaLoca (29, trên file copy).

**Đợt 3 — Smoke test tool trung bình:**
Tile Layout, PointCloud, FamiGen, ManaPara, ManaContains, IFC-SG, ModelAuditor, ManaStyles, AutoDimension, ManaAnno.

**Đợt 4 — Tool mỏng/ít rủi ro:**
Các launcher còn lại + UI.stack + PDF import + Feedback.

**Sau mỗi đợt:** chạy lại `python3 dev/audit_tools.py --quiet` và `python3 dev/sync_wpf_styles.py --check` để bảo đảm không tạo regression tĩnh.

---

## 6. Ghi chú debug trong Revit

- Bật debug log: giữ `CTRL` khi bấm pushbutton → chạy kèm output window; hoặc `pyrevit env` / `pyrevit logs`.
- Với các file nhiều bare `except:` (F5), tạm chèn `import traceback; print(traceback.format_exc())` vào nhánh except khi truy lỗi — nhớ gỡ trước khi commit.
- Tool ghi model luôn test kèm **Ctrl+Z** để xác nhận transaction gom đúng 1 bước.
- Tool phụ thuộc máy (BatchOut in PDF, ImageToDrafting compile C#, MCPControl mở port) cần test trên đúng máy triển khai, không chỉ máy dev.

---

## 7. Tổng kết cuối — GĐ4 (2026-07-05)

### Kết quả
- **Regression cuối**: `audit_tools --quiet` clean · `audit_ui --quiet` 0/54 file có vấn đề
  (lỗi IFCSG copyright trùng ×2 từng ghi ở GĐ2 Ngày 4 nay đã hết) · `sync_wpf_styles --check`
  53/53, 0 lệch.
- **41 tool đóng theo xác nhận chung của user 2026-07-05**; bảng 6 panel (`dev/plan/panel-*.md`)
  không còn ô ⬜; các mục checklist chi tiết chưa test riêng lẻ coi là pass theo xác nhận chung,
  mở lại nếu phát sinh lỗi khi dùng thực tế. **Cập nhật 2026-07-05: PDF import mở lại (🔄)** — user
  báo crash + không hiện thông tin khi thực sự dùng; đã sửa (F11), chờ re-test. Đây là bằng chứng
  cho giới hạn của xác nhận theo lô: tool chỉ mới sửa UI có thể vẫn còn bug functional.
- **Review code GĐ4** trên commit mới nhất `5c4c3b1` (Assistant native agent + providers, ~2.1k
  dòng): kiến trúc ExternalEvent đăng ký trên main thread + worker thread block chờ — đúng chuẩn
  modeless; UI thread chỉ dùng `get_status_instant()` (probe đầy đủ chạy nền) — đúng bài học cũ;
  ghi history/registry đã dùng pattern dumps→io.open an toàn; SheetGen bổ sung commit/rollback
  hợp lý; AutoDimension helper mới có guard đầy đủ. Phát hiện + sửa ngay **F10** (2 điểm ghi
  JSON không an toàn còn sót — xem mục 2).
- Screenshot bộ UI: ⏭️ skip có lý do — user QA trực tiếp trong Revit (tiền lệ GĐ2 Ngày 4).
- 4 roadmap độc lập: đóng toàn bộ 2026-07-05, file đã xóa — bảng kết cục ở `dev/plan/README.md`.

### Issue còn lại (xếp ưu tiên — nguồn: tổng hợp mục "Phát sinh" 6 file panel + review GĐ4)

> **Đợt fix 2026-07-05**: issue 0/1 + phần F6 & DoorThreshold-message đã sửa (5 commit riêng
> lượt: PDF import F11 · ManaSheets Excel · BatchOut F6 · DoorThreshold message + doc này).
> Tất cả static-clean (`ast.parse` + audit), chờ smoke test trong Revit.

| # | Ưu tiên | Issue | Trạng thái / ghi chú |
|---|---------|-------|-----------------|
| 0 | ✅ Đã sửa, chờ test | **PDF import — crash + không hiện thông tin (F11)**: nạp grid đồng bộ trong `__init__` thay vì `ContentRendered`. | Commit `91159c7`; chờ user reload pyRevit + mở lại. `panel-6` Phát sinh |
| 1 | ✅ Đã sửa, chờ test | **ManaSheets Excel Import/Export chết** (lệch API service ↔ dialog): thêm adapter `export_sheets`/`import_sheets` (dict-based, cột `ID` để round-trip khớp lại theo ElementId; giữ nguyên 2 method gốc cho `import_excel_dialog`), dialog khớp theo id trước rồi fallback sheet_number. | Commit `05a5628`; chờ round-trip test trong Revit. `panel-3` Phát sinh |
| 2 | ✅ F6 sửa · 🟡 F5 giữ | **F6** (BatchOut degrade Intelligence im lặng): đã log khi import `api_learner`/`api_updater` fail — commit `e535764`. **F5** (bare `except:` — Wall Cut Profile 37, BatchOut 30, Wall_Adjust 4, DoorThreshold 3, ManaTabs 2): **giữ nguyên làm tech debt** — không chặn chức năng, sửa hàng loạt trên tool đang chạy tốt là churn rủi ro không tương xứng; chỉ chèn log tạm khi thực sự cần debug tool đó. | mục 2 (F5/F6) |
| 3 | ✅ Message sửa · 🟡 geometry giữ | **DoorThreshold**: đã thêm thông báo rõ khi model không có FloorType nào (commit `09cee9b`). Phần **curtain wall / wall nghiêng** vẫn ngoài scope — là feature geometry cần phát triển + test trong Revit, không phải bug; mở việc riêng nếu cần. | `panel-2` |
| 4 | 🟢 Thấp — đóng không triển khai | **Backlog** (mở roadmap mới nếu cần): AutoDimension GĐ B/C/D ~16 mục; FamiGen WS4/WS5; 3 probe FamiGen prompt v2 (chỉ mới pass theo xác nhận chung). | `dev/plan/README.md` |
