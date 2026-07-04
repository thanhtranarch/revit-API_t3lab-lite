# AutoDimension — Roadmap cải thiện (hoàn thiện bản vẽ)

> Chain: `AutoDimension.pushbutton/script.py` → `lib/GUI/AutoDimensionDialog.py` (v2.2.0) → `lib/GUI/Tools/AutoDimension.xaml`
> Mục tiêu: tool tự hoàn thiện dimension cho bản vẽ (plan + section/elevation) đạt chất lượng gần bản vẽ tay, chạy ổn định trên mọi model.

## Kiến trúc hiện tại (6 phase / view)

| Phase | Nội dung | Offset |
|-------|----------|--------|
| 1 | Grid-to-grid overall chains (X + Y, mirror khi Both sides) | L3 |
| 2 | Column strings — chuỗi theo hàng/cột: grid biên → mặt cột → grid biên | L2 |
| 3 | Windows + doors — chuỗi theo hàng/cột, kẹp grid biên | L1 |
| 4 | Wall chains — mặt ext/int/both + grid kẹp | L2 |
| 5 | Lifts (MechanicalEquipment) — chuỗi theo hàng/cột | L1 |
| 6 | Facade chains — NORTH/SOUTH/EAST/WEST: grid + mặt tường + cửa | facade_off |
| S | **Section/Elevation**: chuỗi grid ngang (đỉnh crop) + chuỗi level dọc (mép trái crop) | L3 |

## GĐ A — Sửa lỗi nền tảng ✅ (đã làm 2026-07-04, v2.2.0 — chưa smoke test Revit)

- [x] **Modal thay vì modeless (lỗi nghiêm trọng nhất)**: cửa sổ mở `show(modal=False)` nhưng `run_clicked` gọi `Transaction.Start()` trực tiếp trong WPF handler → ngoài Revit API context → `InvalidOperationException "outside of API context"`. Repo đã có tiền lệ: TagCheckerDialog thử ExternalEvent bị native crash phải revert về modal. → Đổi `show(modal=True)`. (Nếu sau này cần modeless thật, xem GĐ D.)
- [x] **Extent đặt dim dùng bounding box thay vì centroid**: trước đây `x_min…y_max` là min/max *centroid* nên dim tổng (y_max + L3) đè lên nửa ngoài của tường/cột biên. Giờ dùng `bb.Min/Max` của mọi element + vị trí trục grid.
- [x] **Chuỗi dim cột theo hàng/cột**: bỏ vòng lặp per-column (mỗi cột 1 dim 2-ref, N cột cùng hàng = N chuỗi chồng nhau tại cùng offset). Giờ group theo `GROUPING_TOL` (500mm) → 1 chuỗi/hàng: grid biên – mặt L – mặt R từng cột – grid biên (offset + bề rộng + offset như bản vẽ kết cấu tay). Cột hỏng ref chỉ rơi khỏi chuỗi, không chặn cột khác. Bỏ mirror per-row (trước mirror mọi hàng về cùng 1 đường → chồng đống).
- [x] **Bỏ qua tường xiên**: tường không thẳng trục (dùng `AXIS_TOLERANCE`) bị ép vào chuỗi X/Y → dim sai vị trí hoặc NewDimension fail ngầm. Giờ skip có chủ đích.
- [x] **Section/Elevation hoạt động thật**: trước đây `_is_valid_view` cho phép Section/Elevation nhưng toàn bộ toán học là mặt bằng XY → luôn 0 dim + diagnostic khó hiểu. Thêm `_dim_section_view`: tính trong mặt phẳng view qua `CropBox.Transform` (BasisX = phải, BasisY = lên) — chuỗi grid ngang gần đỉnh crop + chuỗi level dọc gần mép trái (gate bằng checkbox Grids).

## GĐ B — Chất lượng bản vẽ (ưu tiên kế tiếp)

- [ ] **Dọn dim cũ khi chạy lại (idempotent)**: hiện chạy 2 lần = 2 lớp dim chồng nhau. Đánh dấu dim do tool tạo (ExtensibleStorage schema hoặc so khớp tên transaction trong lịch sử không khả thi → dùng schema/param) + checkbox "Replace existing T3Lab dims".
- [ ] **Cửa/cửa sổ dim về mặt tường host**: chuỗi L1 hiện kẹp bằng grid biên; chuẩn bản vẽ kiến trúc là jamb → mặt kết thúc tường/góc tường gần nhất. Lấy `HostObjectUtils.GetSideFaces` của host wall + face ends.
- [ ] **Wall core-layer refs thật**: `_collect_wall_core_refs` hiện trả mặt finish ext/int (đoạn narrow-to-core là no-op). Dùng `CompoundStructure.GetCoreBoundaryLayerIndex` + offset để chọn đúng mặt lõi khi user chọn chế độ core.
- [ ] **Quản lý chồng lấn offset (stacking manager)**: đăng ký các "làn" dim đã chiếm theo từng phía (top/bottom/left/right); chuỗi mới tự đẩy ra làn kế tiếp thay vì đè nhau (L1/L2/L3 chỉ là 3 làn tĩnh).
- [ ] **Min-segment tự xử lý**: hiện chỉ đếm và cảnh báo segment < ngưỡng; nâng cấp: tự bỏ ref tạo segment ngắn (merge) hoặc kéo text ra ngoài bằng `TextPosition`.
- [ ] **EQ toggle**: option bật `AreSegmentsEqual` / hiển thị EQ cho chuỗi grid đều nhịp.

## GĐ C — Mở rộng phạm vi

- [ ] **Công trình xoay/trục xiên**: toàn bộ logic giả định trục thẳng X/Y. Tổng quát hoá theo hướng grid chủ đạo (lấy direction phổ biến nhất của grid làm BasisX cục bộ) — mọi phép chiếu qua dot-product như `_dim_section_view` đã làm.
- [ ] **Section/Elevation mở rộng**: thêm mặt tường, mặt sàn/dầm, cửa sổ (sill/head) vào chuỗi dọc; checkbox "Levels" riêng thay vì gate chung với Grids.
- [ ] **Grid nhiều đoạn / grid cong**: hiện lấy `GetEndPoint(0)` — grid arc hoặc multi-segment cho vị trí sai. Lọc + cảnh báo hoặc lấy giao điểm với crop.
- [ ] **Scope box / crop filter**: chỉ dim element trong scope box được chọn (model lớn nhiều khối).
- [ ] **Element trong link**: dim tham chiếu linked walls/grids qua `Reference.CreateLinkReference`.

## GĐ D — UX & vận hành

- [ ] **Lưu preset** qua `script.get_config()` (offset, hướng, category, dim type) — nhớ giữa các phiên.
- [ ] **Dry-run / preview**: nút "Preview" đếm số chuỗi sẽ tạo per phase per view trước khi commit.
- [ ] **Progress bar per view** (đúng chuẩn Lumina: Height 8, `#3B82F6`) khi chạy nhiều view.
- [ ] **Báo cáo kết quả dạng bảng**: view × phase × số chuỗi tạo/fail thay vì 1 TaskDialog gộp.
- [ ] **Modeless đúng cách (nếu cần)**: chuyển Run sang `IExternalEventHandler` — lưu ý tiền lệ TagChecker native crash; cần test kỹ trên Revit thật trước khi đổi.

## Smoke test Revit (bắt buộc trước khi tick GĐ A hoàn tất)

1. Mở floor plan có grid + tường + cột + cửa → Run mặc định: dim tổng grid nằm **ngoài** công trình; mỗi hàng/cột cột đúng **1** chuỗi (grid–mặt cột–grid); Ctrl+Z revert 1 bước.
2. Model có tường xiên → tường xiên không có dim (bị skip), không stacktrace.
3. Mở **section** → check Grids → Run: chuỗi grid ngang gần đỉnh crop + chuỗi level dọc mép trái. Elevation tương tự.
4. View 3D/schedule → thông báo thân thiện, không crash.
5. Chạy 2 lần liên tiếp → xác nhận dim chồng (đúng hạn chế đã biết, chờ GĐ B idempotent).
6. Cửa sổ mở dạng **modal** — xác nhận Run không còn lỗi "outside of API context".
7. **UI 2 cột mới** (2026-07-04): cửa sổ 780×700 mở đúng, không lỗi parse XAML; toggle "Choose specific plan views" hiện list; bật Checking Mode ẩn cột phải (Style/Placement) + wall faces; bật Facade mới hiện 2 ô offset/tolerance; đổi Offset mode sang 3-Level hiện L1/L2/L3; resize nhỏ nhất 720×560 không vỡ layout.
