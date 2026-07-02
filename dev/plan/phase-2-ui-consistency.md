# Giai đoạn 2 — UI Consistency theo chuẩn Lumina (Ngày 3–6)

> Làm SAU giai đoạn 1 (dead code đã archive — không tốn công sửa UI cho file sắp bỏ).
> Branch gợi ý: `fix/phase-2-ui-consistency`. Tham chiếu: `.claude/rules/ui-design-standard.md`.
> Kiểm tra bằng: `python3 dev/audit_ui.py --quiet`

## Kết quả UI review (2026-07-02)

Audit 76 XAML trong `lib/GUI/Tools/` theo chuẩn Lumina:

| Hạng mục | Kết quả |
|----------|---------|
| Palette Lumina (không còn Kinetix/Terra) | ✅ trừ **T3LabAssistant.xaml** sót 2 màu Terra (`#15803D`, `#166534`) |
| Font (Hanken Grotesk/Inter, không Manrope/Segoe UI) | ✅ 100% |
| WindowChrome đủ 5 attribute, multi-line | ✅ 100% |
| Shared styles block đồng bộ | ✅ 76/76 (`sync_wpf_styles.py --check`) |
| Dot-notation Grid definitions (crash risk) | ✅ không có |
| Glyph chrome buttons (E921/E922/E8BB) | ✅ không phát hiện glyph sai |
| **Copyright block đúng snippet chuẩn** | ❌ **75/76 file sai**: nhúng trong status bar/StackPanel, thiếu `Grid.RowSpan="99"` + `Grid.ColumnSpan="99"`, không phải overlay con trực tiếp của root Grid |

→ Vấn đề UI duy nhất mang tính hệ thống là **copyright block**. Đây chính là deviation "Copyright block embedded inside templates instead of direct root Grid child" mà standard đã cảnh báo.

**2 file UI-LOCKED tuyệt đối không sửa** (kể cả copyright): `DWGManagement.xaml`, `ExportManager.xaml`.

## Snippet chuẩn (verbatim — chèn ngay trước `</Grid>` đóng root Grid)

```xml
<!-- Copyright added automatically -->
<TextBlock Text="© Copyright by T3Lab" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8" Foreground="#F59E0B" FontSize="11" IsHitTestVisible="False" Panel.ZIndex="999" Grid.RowSpan="99" Grid.ColumnSpan="99"/>
```

---

## Ngày 3 — Công cụ + Panel 1 & 3 + palette

### 3a. Viết script chuẩn hoá copyright
- [ ] Tạo `dev/fix_copyright.py` (CPython 3):
  1. Nhận danh sách file (hoặc `--panel` filter)
  2. Xoá mọi TextBlock chứa `Copyright by T3Lab` hiện có (kể cả wrapper StackPanel nếu chỉ chứa copyright)
  3. Chèn snippet chuẩn ngay trước thẻ `</Grid>` **cuối cùng đóng root Grid** (Variant A: Grid con đầu tiên của Window; Variant B: Grid root)
  4. Bỏ qua `DWGManagement.xaml`, `ExportManager.xaml` (hard-code UI_LOCKED)
  5. Có `--dry-run` in diff trước khi ghi
- [ ] Chạy `--dry-run` trên 2–3 file, đọc diff bằng mắt trước khi chạy thật
- [ ] ⚠️ Sau khi sửa từng file: `python3 dev/sync_wpf_styles.py --check` phải vẫn pass (không được đụng vào block styles)

### 3b. Sửa palette T3LabAssistant.xaml
- [ ] `#15803D` → `#10B981` (success)
- [ ] `#166534` → `#059669` (success hover)
- [ ] Xem context từng chỗ thay — nếu là màu nút Success thì đúng map trên; nếu là màu khác thì đối chiếu bảng migration trong `ui-design-standard.md`

### 3c. Copyright Panel 1 — Annotation & Select (6 file)
- [ ] `AutoDimension.xaml`
- [ ] `ManaSelect.xaml`
- [ ] `QuickElement.xaml` (dialog phụ của ManaSelect)
- [ ] `ManaAnno.xaml`
- [ ] `DimText.xaml` (dialog phụ của ManaAnno)
- [ ] `TagChecker.xaml` (dialog phụ của ManaAnno)
- ~~`DWGManagement.xaml`~~ — UI-LOCKED, bỏ qua

### 3d. Copyright Panel 3 — Views & Sheets (6 file)
- [ ] `ManaSheets.xaml`
- [ ] `SheetGen.xaml`
- [ ] `ManaViews.xaml`
- [ ] `AdvancedViewManager.xaml` (dialog phụ của ManaViews)
- [ ] `AdvancedViewManagerBatchRename.xaml` (nt)
- [ ] `ParameterSelector.xaml` (variant B — dialog phụ của BatchOut; root là Grid)
- ~~`ExportManager.xaml`~~ — UI-LOCKED, bỏ qua

### Chốt ngày 3
- [ ] `python3 dev/audit_ui.py` — Panel 1+3 + T3LabAssistant palette hết flag
- [ ] Commit: `style: normalize copyright overlay (panel 1+3), migrate T3LabAssistant palette to Lumina`

---

## Ngày 4 — Copyright Panel 2 — Modeling & Datum (16 file, nhiều nhất)

- [ ] `CADToElements.xaml`
- [ ] `CADtoBeam.xaml`
- [ ] `CadtoFloor.xaml`
- [ ] `CadtoFloorLayerItem.xaml` (**variant B** — root Grid, cẩn thận vị trí chèn)
- [ ] `CadtoWall.xaml`
- [ ] `DoorThreshold.xaml`
- [ ] `ImageToDrafting.xaml`
- [ ] `TextToElement.xaml`
- [ ] `PointCloud.xaml`
- [ ] `RoomToFloor.xaml`
- [ ] `TileLayout.xaml`
- [ ] `PropertyLine.xaml`
- [ ] `FamiGen.xaml`
- [ ] `ManaFami.xaml`
- [ ] `SplitElements.xaml`
- [ ] `AutoJoin.xaml`

### Chốt ngày 4
- [ ] `python3 dev/audit_ui.py` — Panel 2 hết flag · `sync_wpf_styles.py --check` pass
- [ ] Commit: `style: normalize copyright overlay (panel 2)`

---

## Ngày 5 — Copyright Panel 4 + 5 + 6 + file dùng chung (~20 file)

### Panel 4 — Data & IFC-SG (7 file)
- [ ] `IFCSG.xaml`
- [ ] `SubtypeDefinerColMap.xaml`
- [ ] `FoundationVolume.xaml`
- [ ] `ManaContains.xaml`
- [ ] `ManaSched.xaml`
- [ ] `ManaPara.xaml`
- [ ] `BCFReader.xaml`

### Panel 5 — Standards & Settings (4 file)
- [ ] `ManaLoca.xaml`
- [ ] `ManaStyles.xaml`
- [ ] `ManaWorkset.xaml`
- [ ] `ModelAuditor.xaml`

### Panel 6 — Support + Standard (8 file)
- [ ] `MCPControl.xaml`
- [ ] `Feedback.xaml`
- [ ] `PDFImport.xaml`
- [ ] `T3LabAssistant.xaml`
- [ ] `ManaTabs.xaml`
- [ ] `RibbonNames.xaml`
- [ ] `BGTheme.xaml`
- [ ] `AutoWork.xaml`
- [ ] `UIStandardShowcase.xaml` (bản copy vào Tools từ GĐ1-F3 — nếu bản gốc đã chuẩn thì skip)

### File dùng chung / variant B còn lại
- [ ] `FindReplace.xaml`
- [ ] `SelectFromDict.xaml`
- [ ] `FamilyLoader.xaml` (variant B — kiểm tra, có thể đã chuẩn)
- [ ] `FamilyLoaderCloud.xaml` (variant B — kiểm tra, có thể đã chuẩn)
- [ ] `ExportManagerTest.xaml` (reference file — sửa luôn cho chuẩn vì là ví dụ mẫu)

### Chốt ngày 5
- [ ] `python3 dev/audit_ui.py --quiet` → **exit 0** (chỉ còn info của 2 file UI-locked)
- [ ] `python3 dev/sync_wpf_styles.py --check` → 0 lệch
- [ ] Commit: `style: normalize copyright overlay (panel 4-6 + shared dialogs)`

---

## Ngày 6 — Verify trực quan trong Revit

> Copyright overlay dùng `Grid.RowSpan/ColumnSpan=99` — phải NHÌN từng cửa sổ để chắc chắn nó nằm góc phải-dưới, không đè lên nút bấm/status bar.

Checklist mỗi cửa sổ (mở lần lượt 41 tool):
1. Copyright màu amber nằm góc phải-dưới, không che nội dung
2. Title bar: T3Lab + tên tool + subtitle + separator đúng chuẩn
3. Nút min/max/close hover đúng (close → nền đỏ chữ trắng)
4. Font toàn cửa sổ là Hanken Grotesk/Inter
5. Resize cửa sổ → copyright bám theo góc, layout không vỡ

- [ ] Panel 1 (4 cửa sổ + dialog phụ)
- [ ] Panel 2 (14 cửa sổ + dialog phụ)
- [ ] Panel 3 (4 cửa sổ)
- [ ] Panel 4 (6 cửa sổ)
- [ ] Panel 5 (4 cửa sổ)
- [ ] Panel 6 (9 cửa sổ)
- [ ] Screenshot mỗi cửa sổ lưu `dev/plan/screenshots/` (đặt tên `<panel>-<tool>.png`)
- [ ] File nào layout vỡ → ghi vào Phát sinh, sửa tay từng file (không chạy lại script hàng loạt)
- [ ] Commit cuối: `style: visual QA pass for Lumina consistency`
- [ ] Cập nhật bảng tiến độ `dev/plan/README.md`

---

## Phát sinh trong giai đoạn 2

_(ghi tại đây: file bị vỡ layout sau fix, màu sai ngữ cảnh, v.v.)_
