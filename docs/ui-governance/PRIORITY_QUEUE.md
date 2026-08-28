# PRIORITY QUEUE — hàng đợi cho cycle kế tiếp

> Cập nhật cuối: 2026-08-28 (cycle 0 — bootstrap)
> Thứ tự xếp hàng theo `04-severity-and-issues.md` §4. Cycle sau vào là chạy từ trên xuống.

---

## P1 — HIGH

### Q1 · Thay `dev/audit_ui.py` bằng `dev/audit_t3.py`

**Vấn đề:** `dev/audit_ui.py` đang enforce chuẩn Lumina **đã bị bỏ** — nó cấm
`FontFamily="Segoe UI"`, bắt buộc `Hanken Grotesk` ở root `<Window>`, và bắt buộc block
`T3LAB SHARED STYLES`. Cả ba đều **ngược** với chuẩn T3. Hệ quả: file đầu tiên được
migrate sang T3 sẽ làm script này báo fail, và gate đỏ sẽ chặn công việc đúng đắn.

**Việc cần làm:** viết `dev/audit_t3.py` enforce luật T3
(`08-design-system-authority.md` §4): Segoe UI/Consolas · 7 size · token `T3.*` không
hex cứng · spacing chia 4 · một primary · `IsDefault`+`IsCancel` · empty state ·
`HorizontalScrollBarVisibility="Disabled"` · một cột `*` · `UseLayoutRounding` +
`SnapsToDevicePixels` + `TextFormattingMode` · `MinWidth`/`MinHeight` · không `Effect` ·
không dot-notation · không `ScrollViewer` bọc list. Script cần cờ `--legacy-count` để
đếm file chưa migrate, và **không** fail vì file legacy (chỉ fail vì file đã khai T3 mà
sai T3).

**Trạng thái:** `MANUAL REVIEW REQUIRED` — sửa gate là quyết định của user.
**Cho tới khi làm xong:** không dùng `audit_ui.py`/`sync_wpf_styles.py --check` làm
pass/fail; chỉ `dev/audit_tools.py --quiet` là gate thật.

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

---

## DESIGN SYSTEM GAPS — chờ user duyệt

### GAP #1 · Vị trí deploy của `T3Lab.Styles.xaml` chưa chốt

`T3Lab.Styles.xaml` đang nằm ở `pyRevit UI Design System/` — **ngoài** `T3Lab.extension/`,
nên pyRevit không nạp được nó lúc runtime. Mọi tool migrate sang T3 đều cần
`<ResourceDictionary Source="..."/>` trỏ tới một đường dẫn có thật trong extension.

**Đề xuất:** deploy vào `T3Lab.extension/lib/GUI/Resources/T3Lab.Styles.xaml`, giữ bản
ở `pyRevit UI Design System/` làm nguồn, và thêm bước đồng bộ (tương tự
`dev/sync_wpf_styles.py`). Không migrate file nào trước khi gap này được chốt — nếu
không, `Source` sẽ trỏ vào chỗ không tồn tại và tool chết lúc mở.

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

*(trống — cycle 0 chưa thực hiện việc nào)*
