# Point Cloud to Model — Nghiên cứu phát triển toàn diện & chống crash

> Ngày lập: 2026-08-14 · Tool: `Modeling & Datum.panel/Create.stack/Create Elements.pulldown/PointCloud.pushbutton/script.py` (2.296 loc)
> Test đi kèm: `dev/test_pointcloud_detect.py` (24 assert, chạy không cần Revit)

---

## 1. Kiến trúc hiện tại

Pipeline 4 tầng, mỗi tầng là một nguồn crash riêng biệt:

```
[Revit/ReCap engine]  GetPoints ──► native, KHÔNG catch được
        │
[Extraction]          tiled box filter → list[(x,y,z)] model space
        │
[Analyzer]            pure Python, 7 detector độc lập
        │
[ElementBuilder]      Transaction/element, bọc 1 TransactionGroup
        │
[WPF Window]          ShowDialog trong Execute + vòng lặp pick
```

Điểm mạnh đã có: extraction tiled + kẹp `averageDistance`, `PointCollection`
copy+Dispose ngay, spatial index 1 m, mỗi detector cô lập trong `run_stage`,
vòng lặp close/reopen cho pick.

---

## 2. Phân tầng rủi ro crash

| # | Tầng | Mã lỗi | Trạng thái |
|---|------|--------|-----------|
| 1 | Native ReCap engine | `0xc0000005` / `0xe0434352` trong `AdskRcPointCloudEngine.dll` | ⚠️ Giảm nhẹ, KHÔNG diệt được |
| 2 | Gọi API ngoài context | `0xe0434352` | ✅ Đã sửa kiến trúc (2026-07-17) |
| 3 | Bão dialog cảnh báo khi generate | Treo → user kill Revit | ✅ **Sửa trong session này** |
| 4 | UI thread starvation | "Not Responding" → user kill Revit | ❌ **Rủi ro còn lại lớn nhất** |
| 5 | Null document | `NoneType` khi mở từ Assistant pane | ✅ **Sửa trong session này** |

### Tầng 1 — engine ReCap (ngoài tầm can thiệp)
Đã chứng minh tool-independent: ProSheets crash y hệt trên cùng dataset.
Đối sách hiện có: tiled 2×2–6×6, kẹp `averageDistance` 1.5–100 mm, `Use View
Crop` (đọc `view.CropBox` thuần API, không vào selection mode), TaskDialog cảnh
báo save trước khi Analyze trên Revit ≤2023.
**Khuyến nghị:** chạy scan-to-BIM trên Revit 2026. Cân nhắc hard-gate ≤2023
sang chế độ chỉ-đọc-crop thay vì chỉ cảnh báo.

### Tầng 2 — API context (đã sửa, cần bảo vệ khỏi hồi quy)
Root cause: `self.Hide()` trên cửa sổ đang `ShowDialog()` làm ShowDialog
return → `Execute` kết thúc → cửa sổ zombie → mọi click sau chạy API ngoài
context. Đã thay bằng: nút pick set `pick_request` rồi `Close()`, main loop
pick trong context và mở lại cửa sổ với state.
**Rủi ro hồi quy:** bất kỳ ai thêm lại `Hide()` cho một nút pick mới sẽ tái
tạo nguyên crash. Cần một guard tĩnh (xem §5, mục A3).

### Tầng 3 — bão dialog cảnh báo ✅ đã sửa
`IFailuresPreprocessor` và `FailureProcessingResult` **được import từ đầu
nhưng chưa bao giờ được implement** — trong khi 8 tool khác trong repo đều có
`WarningSwallower`/`SilentFailurePreprocessor`. Hình học từ scan sinh cảnh báo
gần như trên mọi element (tường lệch trục, tường trùng nhau ở góc chưa vát,
sàn/trần trùng tường, opening không nằm trọn trong host). Mỗi cảnh báo là một
modal dialog chặn cả batch, trong khi TransactionGroup đang mở và hàng trăm
element đang chờ → user thấy Revit treo và kill process.

### Tầng 4 — UI thread ❌ chưa sửa, ưu tiên số 1
`_run_analysis()` và `build_all()` chạy **trên UI thread**, chỉ pump
dispatcher ở `DispatcherPriority.Render` để repaint status bar. Với 50.000
điểm và 7 detector Python thuần, Windows sẽ dán nhãn "Not Responding" — và
theo đúng lịch sử tool này, phản xạ của user là kill Revit. **Không có nút
Cancel.** Đây là "crash" phổ biến nhất còn lại theo cảm nhận người dùng, dù
về mặt kỹ thuật không phải crash.

---

## 3. Đã sửa trong session này

| Sửa | Vị trí | Tác động |
|-----|--------|----------|
| `WarningSwallower` + `_start(t)` gắn vào cả 6 transaction | `ElementBuilder` | Diệt tầng 3 |
| `resolve_doc()` + `_resolve_uidoc()` + guard ở `__main__` | Section "DEFINE VARIABLES" | Diệt tầng 5 |
| Cache `Level` (`_all_levels`) | `PointCloudAnalyzer._best_level` | Bỏ 1 collector sweep/element |
| Stair đếm trung thực (`skipped`) | `build_all` + `btn_generate_clicked` | Không còn báo "N Stairs created" khi `build_stair_marker` là stub trả None |
| **Band-merge tolerance theo bề dày tường** | `detect_walls` | Xem §3.1 — bug thật, test bắt được |

### 3.1 Bug band-merge (mới phát hiện qua test)

Bộ gộp bin cũ: `if band and k - band[-1] > 2` → chỉ gộp qua khe ≤ 2 bin
(100 mm). Nhưng **một bức tường quét từ hai phía cho HAI mặt điểm dày với
LÕI RỖNG ở giữa** (không có điểm bên trong khối xây). Tường 200 mm → hai mặt
cách nhau 4 bin → bị tách thành **hai tường song song**.

Hệ quả kép, nghiêm trọng hơn số lượng:
- mỗi nửa band đo bề dày ≈ 0 → rơi vào clamp sàn 75 mm → `_best_wall_type(75mm)`
  **luôn chọn wall type mỏng nhất project cho MỌI tường**;
- mỗi tường ảo lại sinh bản sao opening của chính nó → cửa/cửa sổ nhân đôi.

Guard `900 mm` bề rộng band ở ngay dưới cho thấy ý định ban đầu **là** gộp cả
hai mặt vào một band — tolerance 2 bin mâu thuẫn với chính nó.

Sửa: `max_gap_bins = max(2, int(max_thickness / bin_n))` (800 mm / 50 mm = 16 bin),
guard 900 mm giữ nguyên để vẫn loại smear chéo.

Kết quả test synthetic (phòng 8×5×3 m, tường 200 mm):

| | Trước | Sau |
|---|---|---|
| Số tường | 8 | **4** ✅ |
| Bề dày đo được | 75 mm (clamp) | **200 mm** ✅ (đúng giá trị mô hình) |
| Cửa / cửa sổ | 2 / 2 (nhân đôi) | **1 / 1** ✅ |

---

## 4. Khoảng trống chất lượng dựng hình (xếp theo tác động)

### G1 — Sàn/trần là hình chữ nhật bao ⚠️ nặng nhất
`detect_horizontal_surfaces` lấy `min/max` X/Y của slice → luôn ra chữ nhật.
Mặt bằng hình L, lõm, hay bất kỳ hình nào không phải chữ nhật đều sai; sàn
tràn ra ngoài ranh nhà.
**Hướng sửa:** đã có sẵn lưới occupancy 500 mm (`_footprint_cells`) — trace
biên lưới đó (marching squares / cell-edge walk), đơn giản hóa bằng
Douglas–Peucker, rồi snap các cạnh về đường tường đã detect.

### G2 — Góc tường không gặp nhau
Mỗi tường tạo độc lập từ centroid + length percentile → các đầu tường không
giao nhau, nhìn trong Revit là hở góc. Không có trim/extend, không
`JoinGeometry`.
**Hướng sửa:** dựng graph giao điểm của các centerline, extend/trim endpoint
về giao điểm, rồi join. (`ElementTransformUtils` đã import sẵn, chưa dùng.)

### G3 — Chiều cao tường & cao độ tầng
`self._wall_height_ft = min(max(z_range, 2.2m), 6.0m)` — **một giá trị cho
TẤT CẢ tường**. `base_z` đã được detect và lưu trong `_data` (detect_openings
có dùng) nhưng `build_wall` bỏ qua: curve đặt ở `z=0.0`, offset `0.0`. Scan
nhiều tầng sẽ dồn hết tường về một level nếu không có Level khớp tên.

### G4 — Chỉ tường được match type
`build_column`/`build_opening` dùng `symbols[0]` — symbol đầu tiên tìm thấy,
tùy ý. Kích thước đã đo (`w×d` của cột, `width×height` của opening) bị bỏ
hoàn toàn. Cột 300×300 detect đúng có thể ra 600×750.

### G5 — Mái: một mặt phẳng least-squares
`detect_roof` fit **một** mặt phẳng trên 15% điểm cao nhất. Mái 2 dốc (gable)
triệt tiêu nhau → slope ≈ 0 → bị loại, hoặc ra một mặt vô nghĩa. `z_ft` dùng
`top_z_thresh` (một ngưỡng, không phải cao độ diềm mái thật).
**Hướng sửa:** RANSAC phân đoạn nhiều mặt phẳng.

### G6 — Cầu thang
Detect chỉ dựa trên z-histogram toàn cục, **không khoanh vùng XY** → scan
nhiều tầng cách đều nhau có thể false-positive. Tạo element là stub.
Hiện đã báo trung thực (§3), nhưng muốn dựng thật cần cluster XY + Architecture
stairs API, hoặc tối thiểu một marker DirectShape/model line.

### G7 — Không có preview trước khi generate
User không thấy được sẽ dựng ra cái gì trước khi ghi vào model. `DetailLine`
đã import sẵn, chưa dùng — vẽ overlay centerline các tường/opening detect được
sẽ cho phép loại bỏ detection sai trước khi commit.

### G8 — UI không expose tham số
`_default_settings()` hardcode 20 tham số (8 toggle detector + ngưỡng). XAML
chỉ có `cmb_density`, hai radio vùng và `btn_use_crop`. Toàn bộ khả năng tinh
chỉnh đã có trong code nhưng người dùng không chạm tới được.

### G9 — Dải quét tường cố định 1,2 m
`scan_z = fz + 1200mm ± 300mm`. Vách lửng, vách chỉ tồn tại trên 1,5 m, hay
clerestory đều bị bỏ qua. Đây cũng là dải dày đặc nội thất nhất — lý do phải
có tall-evidence guard.

---

## 5. Lộ trình

### Phase A — Crash & tin cậy (ưu tiên cao nhất)
- **A1** ✅ WarningSwallower (xong)
- **A2** ✅ resolve_doc + guard (xong)
- **A3** Guard hồi quy tầng 2: thêm assert tĩnh vào `dev/audit_tools.py` —
  cấm `self.Hide()` trong file này; test rằng mọi handler `btn_pick_*` đều
  set `pick_request` và gọi `Close()`
- **A4** **Cancel + chunked analysis** (mục quan trọng nhất còn lại): chia
  `analyzer.run()` theo detector và pump dispatcher giữa các stage với
  `DispatcherPriority.Background`, thêm cờ `_cancelled` và nút Cancel.
  ⚠️ Không dùng background thread cho phần chạm API; phần Analyzer là Python
  thuần nên an toàn để chunk. Xem `ProgressPauseMixin` đã có trong repo.
- **A5** Hard-gate Revit ≤2023: chỉ cho `Use View Crop`, chặn PickBox/PickObject

### Phase B — Hình học đúng
- **B1** G1 footprint polygon (marching squares + Douglas–Peucker + snap tường)
- **B2** G2 wall join graph
- **B3** G3 per-wall height + base offset từ `base_z`, tạo Level còn thiếu

### Phase C — Type intelligence
- **C1** G4 match FamilySymbol theo kích thước đo được cho cột
- **C2** G4 match + set width/height cho cửa/cửa sổ

### Phase D — Nâng cao
- **D1** G7 preview overlay (DetailLine) — nên làm sớm, ROI cao
- **D2** G8 settings UI
- **D3** G5 RANSAC mái
- **D4** G6 cầu thang thật
- **D5** G9 multi-band wall scan

---

## 6. Kiểm thử

### Đã có: `dev/test_pointcloud_detect.py` — 24 assert, không cần Revit
Harness exec **Section 1–5** của script với stub Revit (`FilteredElementCollector`,
`Grid`, `Level`, `ISelectionFilter`…) — Section 6–7 mới cần Revit thật. Trước
session này, 2.296 loc detection math **không có test nào được commit**, dù
plan notes có ghi các lần chạy synthetic ("test 11 case pass") — nghĩa là mọi
lần tinh chỉnh ngưỡng từ đó đến nay đều không được bảo vệ. Bug §3.1 chính là
bằng chứng.

Case bao phủ:

| Case | Nội dung | Bảo vệ điều gì |
|------|----------|----------------|
| C1 | Phòng 8×5×3 m, cửa 900×2100, cửa sổ 1200×1500 | Đếm đúng 4 tường / 1 sàn / 1 trần / 1 cửa / 1 cửa sổ, không có cột & mái ảo |
| C2 | Kích thước tường | Length ~5 m ×2, ~8 m ×2, thickness 100–500 mm |
| C3 | Kệ 4,0×0,6 m cao 1,8 m | Anti-furniture (tall-evidence guard) |
| C4 | Ống đứng 150 mm | Anti-MEP (min cross-section) |
| C5 | Cột thật 300×300 mm | Không loại oan cột thật |
| C6 | Mặt bàn 4,0×1,5 m @750 mm | Anti-furniture (30% footprint guard) |
| C7 | Phòng kín không lỗ | Không sinh cửa/cửa sổ ảo |
| C8 | 1 điểm / collinear / trùng nhau | Không raise trên input suy biến |
| C9 | Unit helper | Round-trip mm↔ft |

```bash
python3 dev/test_pointcloud_detect.py
```

### Cần Revit thật (chưa verify được ở đây)
Toàn bộ Section 6–7 chỉ kiểm chứng được trong Revit. Checklist cho user:

- [ ] Mở tool từ ribbon (Revit 2026) → cửa sổ hiện, `status_count` báo "Revit 2026"
- [ ] Mở tool khi **không có project nào** → TaskDialog báo rõ, không traceback
- [ ] Mở tool từ **Assistant pane** → mở được hoặc báo rõ (trước đây chết ở `__init__`)
- [ ] Analyze cloud thật → không dialog cảnh báo nào chặn giữa batch Generate
- [ ] Generate → **Ctrl+Z một lần** revert sạch toàn bộ (TransactionGroup)
- [ ] Kiểm tra bề dày tường sinh ra khớp wall type thật, **không còn toàn bộ là type mỏng nhất** (§3.1)
- [ ] Nếu có Stair được detect → dialog nói rõ "detected but NOT created"
- [ ] Cloud lớn (>20k điểm): đo thời gian Analyze, xác nhận mức "Not Responding" → dữ liệu đầu vào cho A4

---

## 7. Kết luận

Phần **chống crash** đã tương đối vững sau session này: 3 trong 5 tầng rủi ro
đã đóng (tầng 2 từ 2026-07, tầng 3 và 5 hôm nay). Tầng 1 nằm ngoài tầm can
thiệp — kỷ luật vận hành (Revit 2026) là câu trả lời. **Tầng 4 (UI thread +
thiếu Cancel) là việc kỹ thuật quan trọng nhất còn lại** và nên là hạng mục
tiếp theo.

Phần **dựng hình** thì khác: các bộ lọc anti-noise (MEP, nội thất, occlusion)
đã rất chín, nhưng khâu **xuất ra Revit** vẫn thô — chữ nhật bao thay vì
footprint thật (G1), góc tường hở (G2), một chiều cao cho mọi tường (G3),
type chọn tùy ý (G4). Bug §3.1 cho thấy giá trị của việc có test: một dòng
tolerance sai đã khiến **mọi tường trong mọi model** ra sai bề dày và sai
type, mà không ai phát hiện qua nhiều vòng tinh chỉnh.
