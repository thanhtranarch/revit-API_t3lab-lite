# T3Lab — Review & Kế hoạch Debug từng Tool

> Cập nhật: 2026-07-02 · Phạm vi: toàn bộ 41 pushbutton trong `T3Lab.extension/T3Lab.tab/`
> Công cụ đi kèm: `python3 dev/audit_tools.py` (audit tĩnh) · `python3 dev/sync_wpf_styles.py --check` (đồng bộ style)

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

### 🔴 F1 — `FindReplace.xaml` thiếu handler `button_run`
- `lib/GUI/Tools/FindReplace.xaml:1181` khai báo `Click="button_run"` nhưng `lib/GUI/FindReplace.py` không có method này (WPF_Base cũng không).
- Vì class dùng `wpf.LoadComponent` → **XamlParseException ngay khi mở cửa sổ** nếu được gọi.
- Hiện không tool nào khởi tạo `FindReplace`, nhưng `lib/GUI/forms.py` vẫn import class này — bất kỳ tool mới nào dùng lại sẽ crash.
- **Hành động**: thêm `def button_run(self, sender, e)` hoặc xoá module nếu xác nhận không dùng.

### 🔴 F2 — `CreateFromRooms.xaml` thiếu 4 handler
- Thiếu: `NumberValidationTextBox`, `UIe_ItemChecked`, `button_run`, `text_filter_updated` trong `lib/GUI/CreateFromRooms.py`.
- Không nơi nào import `CreateFromRooms` → dead code, nhưng sẽ crash khi mở nếu được dùng lại.
- **Hành động**: xoá hoặc hoàn thiện handler trước khi tái sử dụng.

### 🟠 F3 — UIShowcase phụ thuộc cấu trúc repo
- `lib/GUI/UIShowcaseDialog.py:33` load XAML từ `[repo]/.claude/standard/UIStandardShowcase.xaml` — **ngoài** thư mục `T3Lab.extension`.
- Nếu deploy extension rời khỏi repo (copy `T3Lab.extension` vào `%APPDATA%\pyRevit\Extensions`), tool hỏng.
- **Hành động**: copy XAML vào `lib/GUI/Tools/` hoặc fallback path.

### 🟠 F4 — 12 XAML mồ côi trong `lib/GUI/Tools/`
`AutoClick`, `BulkFamilyExport`, `ConvertGridline`, `ConvertLevel`, `ExportManagerTest`(*), `InPlaceModelMain`, `JSONtoFamily`, `LocationManager_backup`, `ModelChecker`, `ParameterLoaderCatPicker/Main/Mapper`, `TransferPara` — không file Python nào tham chiếu.
- (*) `ExportManagerTest.xaml` đang được CLAUDE.md dẫn làm ví dụ wizard nav — giữ lại nhưng đánh dấu.
- `DatumManagerDialog.py` (dead code) còn trỏ tới các pushbutton `Datum.pulldown/...` **không tồn tại**.
- **Hành động**: dọn dẹp hoặc chuyển vào thư mục `_archive/`.

### 🟡 F5 — Bare `except:` nuốt lỗi hàng loạt
| File | Số bare except |
|------|----------------|
| `Wall Cut Profile/script.py` | 37 |
| `BatchOut/script.py` | 30 |
| `Wall_Adjust Base/script.py` | 4 |
| `DoorThreshold/script.py` | 3 |
| `ManaTabs/script.py` | 2 |

Khi debug các tool này, lỗi thật sẽ bị nuốt im lặng → phải tạm thêm logger vào các nhánh except trước khi test.

### 🟡 F6 — BatchOut: module "Intelligence" degrade im lặng
- `BatchOut/script.py:62-67`: import `api_learner`/`api_updater` bọc trong bare except → nếu fail, `HAS_API_LEARNER=False` không có log. Cần in cảnh báo khi debug.

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
| 1 | AutoDimension | launcher → `AutoDimensionDialog.py` | 1721 | Logic dim theo reference dễ vỡ giữa các version Revit | Tạo dim trên wall/grid ở plan + section; test view không hỗ trợ | ⬜ |
| 2 | ManaDWG | self + `DWGManagement.xaml` (UI-locked) | 432 | Xoá/relink DWG import vs link | List đúng DWG import & link; xoá 1 import; Undo | ⬜ |
| 3 | ManaSelect | launcher → `ManaSelectDialog.py` | 660 | Filter selection theo category/param | Chọn hỗn hợp element rồi lọc; selection rỗng | ⬜ |
| 4 | ManaAnno | launcher → `ManaAnnoDialog.py` | 1416 | Thao tác hàng loạt trên text/tag/dim | Batch đổi type annotation; view template ảnh hưởng | ⬜ |

### 4.2 Modeling & Datum.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 5 | CADToElements | launcher → `CADToElementsDialog.py` (5 XAML con) | 3143 | Đọc geometry DWG; 3 flow (Wall/Floor/Beam) | Import DWG mẫu → convert từng loại; DWG không có layer khớp | ⬜ |
| 6 | DoorThreshold | self | 356 | 3 bare except; family threshold phải có sẵn | Cửa trên wall thường + curtain wall; thiếu family | ⬜ |
| 7 | ImageToDrafting | self | 1585 | **Compile C# runtime (Vectorizer)** — phụ thuộc .NET compiler trong Revit | Ảnh PNG đen trắng nhỏ; test cả 2 mode (centerline/outline); máy không compile được | ⬜ |
| 8 | Text to Element | launcher → `TextToElementDialog.py` | 466 | Parse text note thành element | Text hợp lệ + text rác | ⬜ |
| 9 | PointCloud | self | 1502 | 6 transaction; file scan lớn | File điểm nhỏ; format sai; cancel giữa chừng | ⬜ |
| 10 | RoomToFloor | self | 383 | Boundary phức tạp (room có island/arc) | Room chữ nhật, room có cột giữa (opening), room chưa placed | ⬜ |
| 11 | Tile Layout | self | 2450 | Không dùng Transaction trực tiếp trong script (kiểm tra wrapper) — logic layout tính toán nhiều | Layout trên floor nhỏ; đổi origin/góc; Undo | ⬜ |
| 12 | PropertyLine | launcher → `PropertyLineDialog.py` | 1727 | Tạo property line từ tọa độ/bảng | Nhập bảng tọa độ hợp lệ + sai đơn vị | ⬜ |
| 13 | FamiGen | launcher → `FamiGenDialog.py` | 2360 | Tạo family qua Document khác (family doc) — dễ leak doc | Gen 1 family, load vào project; huỷ giữa chừng xem family doc có đóng | ⬜ |
| 14 | ManaFami | launcher → `ManaFamiDialog.py` | 1428 | Rename/purge family hàng loạt | Batch rename; family in-use | ⬜ |
| 15 | Wall_Adjust Base | self | 311 | 4 bare except | Wall thường + wall bị attach top/base | ⬜ |
| 16 | SplitElements | launcher → `SplitElementsDialog.py` | 61 | Dialog rất mỏng — logic có thể ở Snippets | Split wall/line theo level | ⬜ |
| 17 | AutoJoin | self | 708 | JoinGeometry hàng loạt → chậm/fail từng cặp | Model nhỏ nhiều wall-floor giao nhau; phần tử không join được | ⬜ |
| 18 | Wall Cut Profile | self | 1624 | **37 bare except**; 8 transaction; edit sketch profile | Wall thẳng + wall cong; profile đè lên nhau; Undo group | ⬜ |

### 4.3 Views & Sheets.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 19 | ManaSheets | launcher → `ManaSheetsDialog.py` | 583 | Batch sửa sheet | Rename/renumber; sheet đang mở | ⬜ |
| 20 | SheetGen | self | 1262 | 7 transaction, 0 rollback — lỗi giữa chừng để lại state dở | Gen sheet từ template; titleblock thiếu; fail giữa batch | ⬜ |
| 21 | ManaViews | launcher → `ManaViewsDialog.py` | 689 | Xoá/đổi view hàng loạt | View đang active; view trên sheet | ⬜ |
| 22 | BatchOut | self + `ExportManager.xaml` (UI-locked) | 3764 | **30 bare except; F6 (Intelligence degrade im lặng)**; export PDF/DWG/NWC phụ thuộc máy | Export 2 sheet PDF + DWG; máy không có PDF printer; đường dẫn output không tồn tại; kiểm tra `HAS_API_LEARNER` | ⬜ |

### 4.4 Data & IFC-SG.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 23 | IFC-SG | launcher → `IFCSGDialog.py` + `configs/` | 2236 | Mapping param IFC-SG từ config JSON | Config mặc định; config thiếu field; apply lên model mẫu | ⬜ |
| 24 | Foundation Volume | self | 410 | Tính volume qua geometry boolean | Foundation giao nhau; element không có solid | ⬜ |
| 25 | ManaContains | launcher → `ManaContainsDialog.py` | 2204 | Logic spatial containment nặng | Element trong/ngoài room; link model | ⬜ |
| 26 | ManaSched | launcher → `ManaSchedDialog.py` | 1601 | Sửa schedule field/filter qua API | Schedule thường + key schedule | ⬜ |
| 27 | ManaPara | launcher → `ManaParaDialog.py` | 2548 | Shared param file; binding category | Add/remove shared param; file .txt bị lock | ⬜ |
| 28 | BCF Reader | self | 3041 | Parse ZIP/XML BCF; 4 transaction; viewpoint camera | BCF 2.1 mẫu; BCF hỏng; apply viewpoint | ⬜ |

### 4.5 Standards & Settings.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 29 | ManaLoca | self + `ManaLoca.xaml` | 871 | Site/shared coordinates — thao tác nguy hiểm với model thật | **Chỉ test trên file copy**; đổi project base point; Undo | ⬜ |
| 30 | ManaStyles | launcher → `ManaStylesDialog.py` | 2091 | Object styles/line weights toàn model | Đổi 1 style; template lock | ⬜ |
| 31 | ManaWorkset | self | 609 | Chỉ chạy được trên file workshared | File thường (phải báo lỗi đẹp); tạo/rename workset | ⬜ |
| 32 | ModelAuditor | launcher → `ModelAuditorDialog.py` | 1828 | Quét toàn model → chậm trên model lớn | Model nhỏ ra report đúng; model lớn đo thời gian | ⬜ |

### 4.6 Support.panel & Standard.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 33 | MCPControl | launcher → `MCPControlDialog.py` | 348 | Start/stop server `core/server.py`; port 48884 | Bật/tắt server; port bị chiếm | ⬜ |
| 34 | Feedback | self | 232 | Gửi feedback ra ngoài (network) | Không có mạng; nội dung unicode tiếng Việt | ⬜ |
| 35 | PDF import | self, không UI | 33 | Phụ thuộc API import PDF theo version Revit | Revit 2022+ import 1 PDF | ⬜ |
| 36 | T3LabAssistant | self | 3594 | **Tool lớn nhất**; import động (`imp`/`importlib` có guard); gọi `core.server`; UI chat | Mở UI; server chưa chạy; gửi 1 lệnh; kiểm tra fallback import trên IronPython (`importlib.util` không tồn tại → phải rơi vào nhánh `imp`) | ⬜ |
| 37 | ManaTabs | self | 147 | Thao tác UI Revit qua Windows API — dễ vỡ theo version | Revit nhiều tab; 2 bare except | ⬜ |
| 38 | Ribbon Names | self | 154 | Sửa text ribbon qua AdWindows.dll | Đổi tên; restart Revit xem persist | ⬜ |
| 39 | BG Theme | self | 122 | Đổi background/theme qua UIApplication settings | Toggle theme; Revit dark mode | ⬜ |
| 40 | UIShowcase | launcher → `UIShowcaseDialog.py` | 129 | **F3: XAML nằm ngoài extension** | Mở từ repo checkout OK; mở từ extension deploy rời → fail | ⬜ |
| 41 | AutoWork | self + `AutoWork.xaml` | 353 | Không có Transaction trong script — kiểm tra có ghi model không | Chạy happy path; xem có cần transaction | ⬜ |

---

## 5. Thứ tự ưu tiên

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
