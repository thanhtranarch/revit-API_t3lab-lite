# PRIORITY QUEUE — hàng đợi cho cycle kế tiếp

> Cập nhật cuối: 2026-08-28 (cycle 0 — bootstrap + dọn dẹp)
> Thứ tự xếp hàng theo `04-severity-and-issues.md` §4. Cycle sau vào là chạy từ trên xuống.

---

## P1 — HIGH

*(trống — Q1 đã xong 2026-08-28)*

---

## P2 — MEDIUM

### Q2 · Cycle 1: audit sâu 6 tool đầu tiên

Chưa tool nào có điểm. Chọn theo `01-scan.md` §4 — vì chưa có lịch sử, dùng tiêu chí 5
(xoay vòng) và bắt đầu bằng nhóm tool đơn giản để hiệu chỉnh rubric:

`SplitElements` (1170 dòng, nhỏ nhất) · `RibbonNames` · `Feedback` · `SelectFromDict`
· `ManaTabs` · `SubtypeDefinerColMap`

Lý do bắt đầu từ file nhỏ: hiệu chỉnh rubric trên file dễ trước, tránh chấm sai hàng
loạt trên file 2000 dòng rồi phải chấm lại.

### Q3 · Cycle 1–2: migrate 2 file đầu tiên sang T3

Chọn sau khi Q2 xong (cần biết pattern của từng tool trước). Ứng viên ưu tiên: tool
nhỏ, một pattern rõ ràng, ít binding — để quy trình migration ở
`08-design-system-authority.md` §6.2 được thử ở chi phí thấp nhất.

### Q4 · Quét toàn bộ 51 file cho 3 lỗi P0 tiềm ẩn

Chạy ngay ở cycle 1, không cần audit sâu:

```bash
grep -rln '<Grid\.\(Row\|Column\)Definition' T3Lab.extension/lib/GUI/Tools/
grep -rln 'DropShadowEffect\|<.*\.Effect>'   T3Lab.extension/lib/GUI/Tools/
grep -rlzoP '(?s)<ScrollViewer.{0,400}?<(DataGrid|ListBox|ListView)' T3Lab.extension/lib/GUI/Tools/
```

Lệnh 1 tìm crash `EMPTYPROPERTYELEMENT`; lệnh 3 tìm nguyên nhân treo Revit phổ biến
nhất (list bị bọc `ScrollViewer` → mất virtualization). Cả hai là **P0** nếu có kết quả.

Vì sao vẫn cần grep dù đã có `audit_t3.py`: script chỉ soi đầy đủ file **đã khai T3**.
51 file legacy không được soi, nên 2 lỗi P0 này phải quét thủ công cho tới khi chúng
được migrate.

---

## DESIGN SYSTEM GAPS — chờ user duyệt

### GAP #2 · Standard định nghĩa size class S/M/L nhưng stylesheet không cung cấp cách áp

`T3LAB_UI_STANDARD.md` quy định `S 420×260–320` · `M 560×420–560` · `L 1000×620`, nhưng
`T3Lab.Styles.xaml` không có `Style` nào `TargetType="Window"`. Mỗi tool sẽ tự đặt
`Width`/`Height`/`MinWidth` bằng tay → chính là loại inconsistency mà Design System sinh
ra để chống.

**Đề xuất:** thêm `T3.Window.S` / `T3.Window.M` / `T3.Window.L` vào `T3Lab.Styles.xaml`.

---

### GAP #5 · Không có indeterminate progress

`T3.ProgressBar` dùng `ControlTemplate` tự viết, không có storyboard cho
`IsIndeterminate`. Đặt `IsIndeterminate="True"` chỉ ra một thanh track rỗng đứng im —
tệ hơn không có gì, vì nó trông như tool đã treo.

Phát hiện 2026-08-28 khi dựng lại showcase (cycle 0f); mẫu vật indeterminate đã bị gỡ
khỏi showcase và thay bằng callout nói rõ T3 không có nó.

**Đề xuất:** giữ nguyên như hiện tại. Chuẩn T3 (P3) vốn đã yêu cầu **đếm việc trước rồi
báo `n / total`** — một thanh chỉ quay không cho người dùng biết có nên chờ hay không.
Chỉ thêm component `T3.ProgressBar.Indeterminate` khi có một tool thật không thể đếm
trước được số việc, và khi đó nó vào stylesheet, không phải mỗi tool tự chế.

---

## Đã hoàn thành

### ✅ GAP #4 · Chrome wizard & cửa sổ — 2026-08-28

Đóng bằng hướng 1 (thêm component), theo quyết định của user. 8 component vào
`T3Lab.Styles.xaml`: `T3.WinCtrl` · `T3.WinClose` · `T3.TabItem.Hidden` ·
`T3.Rail.Tile` · `T3.Rail.Logo` · `T3.Pill` · `T3.ComboBox.Toggle` ·
`T3.ScrollBar.Thumb`. Stylesheet 84 → 92 key.

Đo lại khi bắt tay vào thì gap nhỏ hơn ước tính ban đầu: trong 26 `<Style>` bị báo,
**15 là định nghĩa chết** (không ai tham chiếu) và **5 là implicit style** hợp lệ —
chỉ 6 cái thật sự cần nâng thành component. Xoá 15 style chết gỡ luôn phần lớn vi
phạm `CornerRadius`.

Hai chỗ chỉnh trong gate, cả hai là lỗi của gate chứ không phải nới luật:
- `<Style>` không có `x:Key` là override cục bộ, stock WPF không có cách khác —
  chỉ style **có** `x:Key` mới là component đặt sai chỗ.
- `PENDING_GAP` nay **rỗng**: BatchOut 0 vi phạm, 0 waiver.

### ✅ GAP #1 · Vị trí deploy `T3Lab.Styles.xaml` — 2026-08-28

Chốt: thư mục `pyRevit UI Design System/` là **file mẫu để refer**, không phải artifact
runtime. `python3 dev/sync_t3_styles.py` **nhúng** nội dung stylesheet vào
`<Window.Resources>` của từng tool XAML, giữa hai marker `T3 STYLES`.

**Bản `MergedDictionaries` đã thử và THẤT BẠI** (2026-08-28): pyRevit nạp XAML bằng
`XamlReader` không có base URI → `IOException: Assembly.GetEntryAssembly() returns null`
→ BatchOut chết ngay lúc mở. Chi tiết + hai cái bẫy khi nhúng:
`08-design-system-authority.md` §5.

### ✅ GAP #3 · Copyright — 2026-08-28

Chốt: **BẮT BUỘC** trong UI của mọi tool. Đã thêm token `T3.Copyright #F59E0B` + style
`T3.Copyright` vào stylesheet, luật 11 vào `T3LAB_UI_STANDARD.md`, luật 13 vào
`.claude/rules/new-tool-standard.md`, và `audit_t3.py` enforce trên **mọi** file kể cả
legacy. Hiện 53/53 file có đúng một dòng (`CadtoFloorLayerItem.xaml` là item-template,
miễn trừ theo chuẩn). Đã bổ sung cho `T3LabAssistant.xaml` — file duy nhất còn thiếu.

### ✅ Q1 · Thay `dev/audit_ui.py` bằng `dev/audit_t3.py` — 2026-08-28

`dev/audit_t3.py` enforce luật T3, chia file làm **T3-DECLARED** (soi đầy đủ, vi phạm ⇒
`exit 1`) và **LEGACY** (chỉ đếm, không fail gate) — nên gate xanh ngay hôm nay dù
0/51 file đạt chuẩn, và siết dần theo từng file được migrate. Cờ: `--quiet` ·
`--legacy` · `--legacy-count` · `--file`.

Đã xoá cùng đợt: `dev/audit_ui.py`, `dev/sync_wpf_styles.py`,
`.claude/rules/ui-design-standard.md`, `.claude/standard/UIStandardShowcase.xaml`,
`.claude/docs/wpf-window-templates.md`, `.claude/agents/ui-police-agent.md`,
`.codex/agents/ui-police-agent.toml`, và 6 script migration one-shot
(`fix_ui.py`, `migrate_xaml_colors.py`, `add_sidebar_nav.py`, `export_palette.py`,
`mute_colors.js`, `refine_palette.js`).

Đã thêm: `.claude/rules/new-tool-standard.md` — luật cho **mọi tool/script mới**.
Đã viết lại theo T3: `.claude/agents/ui-agent.md`, `.claude/skills/xaml-templates.md`,
`.codex/agents/ui-agent.toml`.

**Kiểm chứng:** dựng 1 XAML cố ý sai → bắt đủ 18 vi phạm, `exit 1`; dựng 1 XAML từ
đúng snippet trong `xaml-templates.md` → `exit 0`. Tài liệu và gate khớp nhau.
