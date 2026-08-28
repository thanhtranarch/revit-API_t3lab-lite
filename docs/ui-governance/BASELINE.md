# BASELINE — trạng thái & điểm UI/UX từng tool

> **Cycle hiện tại: 0 (bootstrap).** Bảng dưới là kiểm kê, **chưa có điểm** — không tool
> nào được audit sâu. Cycle 1 sẽ chấm 5–8 tool đầu tiên theo `01-scan.md` §4.
> Ô `—` nghĩa là **chưa đo**, không phải điểm 0. Không được điền số vào đây nếu chưa
> thực sự chấm theo `03-scoring.md`.

| Trường | Giá trị |
|--------|---------|
| Cycle | 0 |
| Ngày | 2026-08-28 |
| Overall Score | — (chưa có tool nào được chấm) |
| Tổng XAML | 54 |
| Đạt chuẩn T3 | **0** |
| Còn LEGACY | **51** |
| LOCKED (ngoài phạm vi) | 3 |
| Gate `dev/audit_tools.py` | chưa chạy trong cycle này |

**Chú thích cột** — xem `06-baseline-tracking.md` §1.
`State`: `T3` đạt chuẩn · `LEGACY` chưa migrate · `LOCKED` ngoài phạm vi routine.
`Variant`: `A` root `<Window>` · `B` root `<Grid>` (modal content / item template).
`Hex`: số màu hex cứng riêng biệt trong file — chỉ số nợ migration thô, càng cao càng xa chuẩn T3.

| Tool | File | State | Variant | Pattern | Size | Hex | Prev | Score | Δ | Last audit | Open | Fixed |
|------|------|-------|---------|---------|------|-----|------|-------|---|-----------|------|-------|
| AdvancedViewManager | `Tools/AdvancedViewManager.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| AdvancedViewManagerBatchRename | `Tools/AdvancedViewManagerBatchRename.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| AutoDimension | `Tools/AutoDimension.xaml` | LEGACY | A | — | — | 40 | — | — | — | — | — | 0 |
| AutoJoin | `Tools/AutoJoin.xaml` | LEGACY | A | — | — | 43 | — | — | — | — | — | 0 |
| AutoWork | `Tools/AutoWork.xaml` | LEGACY | A | — | — | 43 | — | — | — | — | — | 0 |
| BCFReader | `Tools/BCFReader.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| BGTheme | `Tools/BGTheme.xaml` | LEGACY | A | — | — | 46 | — | — | — | — | — | 0 |
| CADToElements | `Tools/CADToElements.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| CADtoBeam | `Tools/CADtoBeam.xaml` | LEGACY | A | — | — | 39 | — | — | — | — | — | 0 |
| CadtoFloor | `Tools/CadtoFloor.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| CadtoFloorLayerItem | `Tools/CadtoFloorLayerItem.xaml` | LEGACY | B | — | — | 38 | — | — | — | — | — | 0 |
| CadtoWall | `Tools/CadtoWall.xaml` | LEGACY | A | — | — | 41 | — | — | — | — | — | 0 |
| DWGManagement | `Tools/DWGManagement.xaml` | LOCKED | A | — | — | 44 | — | — | — | — | — | 0 |
| DimText | `Tools/DimText.xaml` | LEGACY | A | — | — | 43 | — | — | — | — | — | 0 |
| DoorThreshold | `Tools/DoorThreshold.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| ExportManager | `Tools/ExportManager.xaml` | LOCKED | A | — | — | 43 | — | — | — | — | — | 0 |
| ExportManagerTest | `Tools/ExportManagerTest.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| FamiGen | `Tools/FamiGen.xaml` | LEGACY | A | — | — | 45 | — | — | — | — | — | 0 |
| Feedback | `Tools/Feedback.xaml` | LEGACY | A | — | — | 38 | — | — | — | — | — | 0 |
| FindReplace | `Tools/FindReplace.xaml` | LEGACY | A | — | — | 41 | — | — | — | — | — | 0 |
| FoundationVolume | `Tools/FoundationVolume.xaml` | LEGACY | A | — | — | 40 | — | — | — | — | — | 0 |
| IFCSG | `Tools/IFCSG.xaml` | LEGACY | A | — | — | 46 | — | — | — | — | — | 0 |
| ImageToDrafting | `Tools/ImageToDrafting.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| LLMSetting | `Tools/LLMSetting.xaml` | LEGACY | A | — | — | 46 | — | — | — | — | — | 0 |
| MCPControl | `Tools/MCPControl.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| ManaAnno | `Tools/ManaAnno.xaml` | LEGACY | A | — | — | 53 | — | — | — | — | — | 0 |
| ManaContains | `Tools/ManaContains.xaml` | LEGACY | A | — | — | 39 | — | — | — | — | — | 0 |
| ManaFami | `Tools/ManaFami.xaml` | LEGACY | A | — | — | 46 | — | — | — | — | — | 0 |
| ManaLoca | `Tools/ManaLoca.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| ManaPara | `Tools/ManaPara.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| ManaSched | `Tools/ManaSched.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| ManaSelect | `Tools/ManaSelect.xaml` | LEGACY | A | — | — | 39 | — | — | — | — | — | 0 |
| ManaSheets | `Tools/ManaSheets.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| ManaStyles | `Tools/ManaStyles.xaml` | LEGACY | A | — | — | 41 | — | — | — | — | — | 0 |
| ManaTabs | `Tools/ManaTabs.xaml` | LEGACY | A | — | — | 38 | — | — | — | — | — | 0 |
| ManaViews | `Tools/ManaViews.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| ManaWorkset | `Tools/ManaWorkset.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| ModelAuditor | `Tools/ModelAuditor.xaml` | LEGACY | A | — | — | 52 | — | — | — | — | — | 0 |
| PDFImport | `Tools/PDFImport.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| ParameterSelector | `Tools/ParameterSelector.xaml` | LEGACY | B | — | — | 41 | — | — | — | — | — | 0 |
| PointCloud | `Tools/PointCloud.xaml` | LEGACY | A | — | — | 50 | — | — | — | — | — | 0 |
| PropertyLine | `Tools/PropertyLine.xaml` | LEGACY | A | — | — | 58 | — | — | — | — | — | 0 |
| QuickElement | `Tools/QuickElement.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| RibbonNames | `Tools/RibbonNames.xaml` | LEGACY | A | — | — | 41 | — | — | — | — | — | 0 |
| RoomToFloor | `Tools/RoomToFloor.xaml` | LEGACY | A | — | — | 44 | — | — | — | — | — | 0 |
| SelectFromDict | `Tools/SelectFromDict.xaml` | LEGACY | A | — | — | 41 | — | — | — | — | — | 0 |
| SheetGen | `Tools/SheetGen.xaml` | LEGACY | A | — | — | 42 | — | — | — | — | — | 0 |
| SplitElements | `Tools/SplitElements.xaml` | LEGACY | A | — | — | 39 | — | — | — | — | — | 0 |
| SubtypeDefinerColMap | `Tools/SubtypeDefinerColMap.xaml` | LEGACY | A | — | — | 40 | — | — | — | — | — | 0 |
| T3LabAssistant | `Tools/T3LabAssistant.xaml` | LOCKED | A | — | — | 57 | — | — | — | — | — | 0 |
| TagChecker | `Tools/TagChecker.xaml` | LEGACY | A | — | — | 43 | — | — | — | — | — | 0 |
| TextToElement | `Tools/TextToElement.xaml` | LEGACY | A | — | — | 41 | — | — | — | — | — | 0 |
| TileLayout | `Tools/TileLayout.xaml` | LEGACY | A | — | — | 41 | — | — | — | — | — | 0 |
| UIStandardShowcase | `Tools/UIStandardShowcase.xaml` | LEGACY | A | — | — | 46 | — | — | — | — | — | 0 |

## Ghi chú kiểm kê cycle 0

- **0/51 file đạt chuẩn T3.** Toàn bộ tool đang dùng hệ Lumina đã bị bỏ (hex cứng
  38–58 màu/file, `FontFamily="Hanken Grotesk"`, block `T3LAB SHARED STYLES`).
  Đây là điểm xuất phát, không phải lỗi của cycle nào.
- **`UIStandardShowcase.xaml` không còn là file tham chiếu.** Nó minh hoạ chuẩn cũ.
  File tham chiếu duy nhất là `pyRevit UI Design System/T3Lab.Styles.xaml`.
- **`CadtoFloorLayerItem.xaml` và `ParameterSelector.xaml` là variant B** (root `<Grid>`)
  — không có title bar/footer riêng, chấm điểm bỏ qua các tiêu chí đó.
- Cột `Pattern` và `Size` để trống cho tới khi tool được audit sâu lần đầu — phân loại
  pattern đòi hỏi đọc cả XAML lẫn script.py, không đoán từ tên file.
