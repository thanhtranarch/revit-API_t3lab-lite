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

### GAP #1 · Vị trí deploy của `T3Lab.Styles.xaml` chưa chốt

`T3Lab.Styles.xaml` đang nằm ở `pyRevit UI Design System/` — **ngoài** `T3Lab.extension/`,
nên pyRevit không nạp được nó lúc runtime. Mọi tool migrate sang T3 đều cần
`<ResourceDictionary Source="..."/>` trỏ tới một đường dẫn có thật trong extension.

**Đề xuất:** deploy vào `T3Lab.extension/lib/GUI/Resources/T3Lab.Styles.xaml`, giữ bản
ở `pyRevit UI Design System/` làm nguồn, và thêm một bước đồng bộ (`dev/sync_t3_styles.py`,
chưa viết). Không migrate file nào trước khi gap này được chốt — nếu không, `Source` sẽ
trỏ vào chỗ không tồn tại và tool chết lúc mở.

Snippet trong `.claude/rules/new-tool-standard.md` và `.claude/skills/xaml-templates.md`
đang dùng `Source="../Resources/T3Lab.Styles.xaml"` theo đúng đề xuất này — chốt xong
thì hai file đó đã đúng sẵn.

**Chặn:** Q3.

### GAP #2 · Standard định nghĩa size class S/M/L nhưng stylesheet không cung cấp cách áp

`T3LAB_UI_STANDARD.md` quy định `S 420×260–320` · `M 560×420–560` · `L 1000×620`, nhưng
`T3Lab.Styles.xaml` không có `Style` nào `TargetType="Window"`. Mỗi tool sẽ tự đặt
`Width`/`Height`/`MinWidth` bằng tay → chính là loại inconsistency mà Design System sinh
ra để chống.

**Đề xuất:** thêm `T3.Window.S` / `T3.Window.M` / `T3.Window.L` vào `T3Lab.Styles.xaml`.

### GAP #3 · Copyright `© Copyright by T3Lab` không có trong chuẩn T3

Chuẩn cũ bắt buộc mỗi cửa sổ có dòng copyright amber `#F59E0B` ở footer trái, và cả 54
file hiện đều có. `T3LAB_UI_STANDARD.md` **không nhắc tới copyright**, đồng thời cấm
mọi màu ngoài token — `#F59E0B` không phải token T3.

**Câu hỏi cho user:** copyright giữ hay bỏ?
- Giữ → cần thêm token + quy tắc vị trí vào `T3LAB_UI_STANDARD.md`.
- Bỏ → 54 file sẽ mất dòng copyright khi migrate.

Không tự quyết. Cho tới khi có trả lời: **giữ nguyên copyright như hiện trạng** khi
migrate, và ghi là ngoại lệ đã biết.

---

## Đã hoàn thành

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
