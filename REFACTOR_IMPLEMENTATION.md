# T3Lab Extension — Refactor Implementation Plan

> **Mục tiêu:** Giảm số lượng pushbutton trên ribbon bằng cách **gộp các tool trùng workflow/UI** vào những cửa sổ hợp nhất, **giữ nguyên 100% chức năng hiện có**.
>
> **Hiện trạng:** 105 pushbuttons / 20 pulldowns / 13 stacks / 7 panels → **mục tiêu ~76 pushbuttons** (giảm ~28%).

---

## Mục lục
1. [Nguyên tắc consolidation](#1-nguyên-tắc-consolidation)
2. [Bản đồ consolidation tổng thể](#2-bản-đồ-consolidation-tổng-thể)
3. [Lộ trình 3 phase](#3-lộ-trình-3-phase)
4. [Những gì KHÔNG gộp](#4-những-gì-không-gộp)
5. [Pattern implementation chuẩn](#5-pattern-implementation-chuẩn)
6. [Checklist tổng & tiêu chí nghiệm thu](#6-checklist-tổng--tiêu-chí-nghiệm-thu)

---

## 1. Nguyên tắc consolidation

- **Mỗi tool cũ → 1 tab / 1 mode** trong cửa sổ hợp nhất. Logic Revit API giữ nguyên 1:1, chỉ gộp vỏ UI + entry point.
- **1 pushbutton mới thay nhiều pushbutton cũ** → giảm nút trên ribbon, KHÔNG giảm tính năng.
- Dùng **TabControl** (hoặc top-nav RadioButton theo Lumina design) cho cửa sổ đa chức năng.
- **Tiền lệ có sẵn:** `AnnotationManager` đã gộp Dimension Manager + Text Note Manager vào 1 cửa sổ duy nhất — dùng làm khuôn mẫu.
- Tuân thủ `.claude/rules/ui-design-standard.md` (Lumina design system) cho mọi UI mới/sửa.
- Mỗi cửa sổ hợp nhất phải có khối shared styles được sync bằng `python dev/sync_wpf_styles.py`.

---

## 2. Bản đồ consolidation tổng thể

| Nhóm hợp nhất | Gộp từ | → | Cơ chế | Effort | Risk |
|---|---|---|---|---|---|
| **Style Manager** | Hatching + Line Style Edit + Line Pattern | 3→1 | Tabs: Fill / Line Style / Line Pattern | Thấp | Thấp |
| **Select Similar** | 6 biến thể SelectSimilar (Cat/Fam/Type × View/Model) | 6→1 | 1 dialog + toggle In View/In Model | Thấp | Thấp |
| **Select on Sheets** | SelectOnSheets_DWG + SelectOnSheets_TitleBlocks | 2→1 | radio nguồn | Thấp | Thấp |
| **Schedule Manager** | ScheduleExportImportPro + Schedule_Copy | 2→1 | Tabs: Export-Import / Duplicate | Thấp | Thấp |
| **Workset Manager** | CentralFile + Workset | 2→1 | Tabs: Create Central / Manage | Thấp | Thấp |
| **Family Manager** | Family Management + Load Family | 2→1 | Tabs: Manage / Load | Thấp | Thấp |
| **Datum Manager** | Save/Restore/RestoreAll Grids + Align Grid + ConvertGrid + Align Level + ConvertLevel | 7→1 | Tabs: Grids / Levels (Align · 2D-3D Swap · Save-Restore) | TB | TB |
| **Split Elements** | Wall_Split + Column_Split + Floor_Split | 3→1 | auto-detect category từ selection | TB | TB |
| **Cleanup Manager** | SmartPurge + AdvancedPurge + SmartDelete | 3→1 | Tabs: Purge (toggle Advanced) / Delete | TB | TB (destructive) |
| **CAD → Elements** | Beam + CadtoWall + CadtoFloor | 3→1 | chọn loại element đích | TB | TB |
| **IFC-SG Assign** | Auto Assign + Manual Assign | 2→1 | mode Auto / Manual | Cao | TB |
| **View & Sheet Hub** | ViewManager + ViewTemplate + Create Room Plan / SheetManager + Sheet re-number | 5→2 | 2 cửa sổ Tabs (Views / Sheets) | Cao | TB |

**Tổng giảm dự kiến: ~26 pushbutton (105 → ~79).**

---

## 3. Lộ trình 3 phase

### Phase 1 — Quick wins (risk thấp, làm trước)
- Style Manager (3→1)
- Select Similar (6→1) + Select on Sheets (2→1)
- Schedule Manager (2→1)
- Workset Manager (2→1)
- Family Manager (2→1)

### Phase 2 — Domain merges (risk trung bình)
- Datum Manager (7→1)
- Split Elements (3→1)
- Cleanup Manager (3→1)
- CAD → Elements (3→1)

### Phase 3 — Managers lớn (effort cao)
- View & Sheet Hub (5→2)
- IFC-SG Assign (2→1)

---

## 4. Những gì KHÔNG gộp

- **AI Connection panel (giữ nguyên toàn bộ):** T3Lab Assistant, Settings, StartMCP, StopMCP — mỗi cái có flow riêng biệt, không gộp.
- **SmartAlign / Distribute (8 nút)** — one-click trực tiếp, gộp vào cửa sổ làm chậm UX.
- Creator độc lập: RoomToFloor, DoorThreshold, ImageToDrafting, PointCloud, PropertyLine, JSON To Family, CAD To Family, Wall Cut Profile, Wall Adjust Base.
- IFC-SG standalone khác workflow: IFCSGChecker, Subtype Definer, ParameterLoader.
- Singletons: BatchOut, AutoJoin, AutoDimension, Snap Dimension, ColorSplasher, ParaManager, BCF Reader, ContainsManager, BG Theme, Ribbon Names, Tab Manager, Feedback, AutoWork, UIStandard, AnnotationManager, Reset Overrides, Linked Element Box, Material List, ModelChecker, HealthCheck, InPlaceModel, LocationManager, Warning, các Parameter Tools, RoomDataCollector, Foundation Volume.

---

## 5. Pattern implementation chuẩn

Mỗi nhóm gộp theo **6 bước**:

1. **Tạo XAML hợp nhất** trong `lib/GUI/Tools/<NewName>.xaml` theo Variant A (Lumina). Dùng `TabControl` (hoặc top-nav RadioButton) — mỗi tool cũ = 1 tab.
2. **Copy khối shared styles** rồi chạy `python dev/sync_wpf_styles.py` để đồng bộ.
3. **Tạo pushbutton mới** `<NewName>.pushbutton/script.py` + `icon.png`; mỗi tab nạp logic của tool cũ (import module hoặc copy hàm core). Header/imports/window-control theo BatchOut frame (`@script-frame-agent`).
4. **Di chuyển logic Revit API** nguyên vẹn từ các script cũ vào — giữ transaction/collector y hệt, không đổi hành vi.
5. **Xóa các pushbutton cũ** sau khi xác nhận tab mới chạy đúng.
6. **QA** bằng `@qa-agent` + `@ui-police-agent` (audit UI consistency).

### Ghi chú riêng từng nhóm

| Nhóm | Lưu ý implementation |
|---|---|
| Style Manager | 3 tool đều là grid-manager (list → rename → delete) gần giống nhau → share nhiều code. Đặt vào `Style Editors.stack`. |
| Select Similar | Không cần WPF nặng: 2 RadioGroup (Category/Family/Type · View/Model) + nút Run. Có thể dùng `forms.CommandSwitchWindow`. |
| Cleanup Manager | AdvancedPurge có cảnh báo "dangerous" → giữ warning UI rõ ràng, không ẩn sau tab thường. |
| Split Elements | Auto-detect category từ selection; fallback dropdown chọn Wall/Column/Floor. |
| Datum Manager | Gộp xuyên panel (Grids đang ở Annotation, Align/Convert ở Project) → đặt 1 chỗ, cập nhật cả 2 panel. |
| View & Sheet Hub | Tách 2 cửa sổ (Views / Sheets) vì DataGrid lớn, gộp 5→1 sẽ quá tải. |
| IFC-SG Assign | Domain phức tạp → để cuối; chỉ gộp Auto + Manual, giữ Checker/Definer/Loader. |

---

## 6. Checklist tổng & tiêu chí nghiệm thu

### Tiến độ
- [ ] **Phase 1**
  - [ ] Style Manager (3→1)
  - [ ] Select Similar (6→1) + Select on Sheets (2→1)
  - [ ] Schedule Manager (2→1)
  - [ ] Workset Manager (2→1)
  - [ ] Family Manager (2→1)
- [ ] **Phase 2**
  - [ ] Datum Manager (7→1)
  - [ ] Split Elements (3→1)
  - [ ] Cleanup Manager (3→1)
  - [ ] CAD → Elements (3→1)
- [ ] **Phase 3**
  - [ ] View & Sheet Hub (5→2)
  - [ ] IFC-SG Assign (2→1)

### Tiêu chí nghiệm thu mỗi nhóm
- Mọi chức năng của tool cũ vẫn truy cập được trong cửa sổ mới (không mất tính năng).
- XAML pass `python dev/sync_wpf_styles.py --check`.
- UI đúng Lumina design (BulkFamilyExport.xaml là chuẩn); pass `@ui-police-agent`.
- Tất cả `x:Name` trong code-behind khớp XAML.
- Pushbutton cũ đã xóa; ribbon hiển thị đúng số nút mới.
- Logic Revit API (transaction/collector) chạy đúng như trước.

### Kết quả cuối
**105 → ~79 pushbuttons** (giảm ~26), giữ nguyên 100% chức năng.
