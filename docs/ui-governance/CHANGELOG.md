# CHANGE HISTORY — lịch sử các cycle

> Entry mới nhất ở **trên cùng**. Mỗi cycle một entry, đủ 11 trường
> (xem `06-baseline-tracking.md` §8). Không sửa entry cũ — sai thì thêm entry đính chính.

---

## Cycle 0d — 2026-08-28 — 8 COMPONENT CHROME · GAP #4 ĐÓNG

```
DATE:                   2026-08-28
CYCLE:                  0d
TOOLS AUDITED:          1 (ExportManager / BatchOut — dọn nốt)
ISSUES FOUND:           12 waiver tồn từ cycle 0c + 1 lỗi over-broad của gate
ISSUES FIXED:           13 — PENDING_GAP nay RỖNG
ISSUES REMAINING:       0
SCORE CHANGES:          — (chưa chấm điểm)
REGRESSIONS:            —
DESIGN SYSTEM GAPS:     #4 ĐÓNG · #2 vẫn mở
NEXT PRIORITIES:        QA BatchOut trong Revit · Q4 quét P0 · Q2 audit 6 tool đầu
COMMIT:                 <điền sha sau khi commit>
```

User duyệt hướng 1 của GAP #4. Thêm 8 component vào `T3Lab.Styles.xaml`
(84 → 92 key): `T3.WinCtrl` · `T3.WinClose` · `T3.TabItem.Hidden` · `T3.Rail.Tile`
· `T3.Rail.Logo` · `T3.Pill` · `T3.ComboBox.Toggle` · `T3.ScrollBar.Thumb`.

**Đo lại thì gap nhỏ hơn ước tính.** Trong 26 `<Style>` bị báo ở cycle 0c: 15 là
**định nghĩa chết** (sót lại sau đợt swap style trước, không ai tham chiếu — xoá 420
dòng), 5 là **implicit style** hợp lệ, chỉ 6 cái thật sự cần nâng thành component.
Xoá 15 style chết gỡ luôn 8/11 vi phạm `CornerRadius`.

**Hai chỗ chỉnh trong gate — cả hai là lỗi của gate, không phải nới luật:**
1. Gate đếm mọi `<Style>` là "component đặt sai chỗ". Sai: style **không** có `x:Key`
   là override chỉ có hiệu lực trong container chứa nó (vd `ItemContainerStyle`),
   stock WPF không có cách nào khác. Nay chỉ style **có** `x:Key` mới bị tính.
2. Regex tìm style của tôi giả định `x:Key` đứng trước `TargetType` → bỏ sót
   `ScrollBarThumbStyle`. Đã dùng regex không phụ thuộc thứ tự attribute.

**BatchOut: 0 vi phạm, 0 waiver.** `PENDING_GAP` rỗng. 1112 dòng (từ 2083).
Kiểm chứng: 14 `x:Name` mất so với main đều là PART_/template nội bộ của các style đã
nâng lên stylesheet, **không cái nào được script.py dùng**. Ba gate xanh.

**Vẫn CHƯA verify trong Revit.**

---

## Cycle 0c — 2026-08-28 — COPYRIGHT · TIẾNG ANH · MIGRATE BATCHOUT

```
DATE:                   2026-08-28
CYCLE:                  0c
TOOLS AUDITED:          1 (ExportManager / BatchOut — migrate đầu tiên sang T3)
ISSUES FOUND:           27 trên ExportManager + 2 bug trong chính audit_t3.py
ISSUES FIXED:           15 (+2 bug gate)
ISSUES REMAINING:       12 — hoãn có kiểm soát theo DESIGN SYSTEM GAP #4
SCORE CHANGES:          — (chưa chấm điểm; cycle 1 sẽ chấm)
REGRESSIONS:            —
SYSTEM-WIDE PATTERNS:   BatchOut vốn đã dùng ~60% palette T3 — nó là tiền thân
                        của Design System, nên migrate rẻ hơn dự kiến.
DESIGN SYSTEM GAPS:     #1 ĐÓNG · #3 ĐÓNG · #2 vẫn mở · #4 MỚI (chrome wizard)
NEXT PRIORITIES:        QA BatchOut trong Revit · quyết GAP #4 · Q4 quét P0 ·
                        Q2 audit sâu 6 tool đầu
COMMIT:                 <điền sha sau khi commit>
```

**Quyết định của user, đã áp vào chuẩn:**
- `pyRevit UI Design System/` là **file mẫu để refer**, không phải artifact runtime →
  GAP #1 đóng: bản deploy ở `lib/GUI/Resources/T3Lab.Styles.xaml`, sinh bằng
  `dev/sync_t3_styles.py`.
- **Copyright BẮT BUỘC** trong UI mọi tool → GAP #3 đóng: thêm token + style
  `T3.Copyright`, luật 11 vào standard, luật 13 vào new-tool-standard, enforce trên
  **mọi** file kể cả legacy. 53/53 file đạt.
- **Toàn bộ UI phải tiếng Anh** → luật 12 vào standard, luật 14 vào new-tool-standard,
  `audit_t3.py` dò dấu tiếng Việt trong `Text`/`Content`/`Header`/`ToolTip`.

**BatchOut (`ExportManager.xaml`) — gỡ khoá và migrate sang T3:**
gỡ bộ brush riêng 19 key → token T3 · toàn bộ hex → `{StaticResource T3.*}` ·
Hanken Grotesk/Inter → Segoe UI · 19 FontSize về thang 7 size · 70 Margin/Padding về
bội số 4 · bo góc window 22→8 · xoá `DropShadowEffect` · merge stylesheet deploy ·
thay 18 style cục bộ bằng `T3.*` và xoá 10 định nghĩa · `T3.Copyright` ở footer ·
`IsDefault`/`IsCancel` · empty state (tiếng Anh) · virtualization + tắt scroll ngang
cho 2 `ListView`.

**Kiểm chứng cấu trúc:** 136/136 `x:Name` và 53/53 event handler còn nguyên (chỉ mất
`x:Name` nội bộ của ControlTemplate đi theo style đã thay — không tên nào được script.py
dùng). 0 `{StaticResource}` mồ côi. XML parse sạch. Ba gate xanh.

**CHƯA verify trong Revit** — xem checklist NEEDS VERIFICATION trong report.

**Hai bug tự tìm thấy trong `audit_t3.py` và đã sửa:**
1. Style `T3.Copyright` tự cấp `Text` nên file dùng style không chứa chuỗi literal →
   check báo "thiếu copyright" nhầm. Nay đếm cả hai dạng.
2. `Segoe MDL2 Assets` bị coi là font sai — nó là font **glyph** cho chrome cửa sổ
   (min/max/close), mọi window tự vẽ caption đều cần. Nay được cho phép.

**Cơ chế `PENDING_GAP`:** 12 vi phạm còn lại của BatchOut (26 `<Style>` cục bộ + 11
`CornerRadius` tròn/pill) là chrome wizard mà T3 chưa có từ vựng. Khai báo trong
`dev/audit_t3.py`, **luôn in ra** mỗi lần chạy kèm lý do, không làm đỏ gate. Hoãn thì
được, giấu thì không.

---

## Cycle 0b — 2026-08-28 — DỌN DẸP + GATE T3

```
DATE:                   2026-08-28
CYCLE:                  0b (dọn dẹp luật cũ + dựng gate T3; vẫn chưa audit tool nào)
TOOLS AUDITED:          0
ISSUES FOUND:           1 (false positive của chính audit_t3.py — xem dưới)
ISSUES FIXED:           1
ISSUES REMAINING:       0
SCORE CHANGES:          — (chưa có điểm nào)
REGRESSIONS:            —
SYSTEM-WIDE PATTERNS:   —
DESIGN SYSTEM GAPS:     #1 và #3 vẫn mở; #2 vẫn mở. Q1 đã đóng.
NEXT PRIORITIES:        Q4 quét 3 lỗi P0 tiềm ẩn · Q2 audit sâu 6 tool đầu ·
                        Q3 migrate 2 file (bị GAP #1 chặn)
COMMIT:                 <điền sha sau khi commit>
```

**Thêm `dev/audit_t3.py`** — gate của chuẩn T3, thay `dev/audit_ui.py`. Chia file làm
**T3-DECLARED** (soi đầy đủ, vi phạm ⇒ `exit 1`) và **LEGACY** (chỉ đếm). Nhờ vậy gate
xanh ngay hôm nay dù 0/51 file đạt chuẩn, và siết dần theo từng file được migrate.

**Thêm `.claude/rules/new-tool-standard.md`** — luật cho **mọi tool/script mới**:
bố cục file, 12 luật XAML, khung `script.py` với 9 luật (transaction, rollback, không
nuốt exception, nạp grid trong `__init__`…), checklist trước khi commit.

**Xoá 13 file của chuẩn đã bỏ:** `.claude/rules/ui-design-standard.md` ·
`.claude/standard/UIStandardShowcase.xaml` · `.claude/docs/wpf-window-templates.md` ·
`.claude/agents/ui-police-agent.md` · `.codex/agents/ui-police-agent.toml` ·
`dev/audit_ui.py` · `dev/sync_wpf_styles.py` · và 6 script migration one-shot
(`fix_ui.py`, `migrate_xaml_colors.py`, `add_sidebar_nav.py`, `export_palette.py`,
`mute_colors.js`, `refine_palette.js`). Không file code sống nào phụ thuộc chúng —
đã kiểm tra trước khi xoá; nội dung còn trong lịch sử git.

**Viết lại theo T3:** `.claude/agents/ui-agent.md` · `.claude/skills/xaml-templates.md`
· `.codex/agents/ui-agent.toml` · `.claude/CLAUDE.md` · `AGENTS.md` · `README.md` ·
`.claude/settings.local.json` (bỏ 12 permission trỏ script đã xoá).

**Issue tự tìm thấy và đã sửa:** bản đầu của `audit_t3.py` báo false positive
`thiếu HorizontalScrollBarVisibility` / `thiếu EnableRowVirtualization` cho `DataGrid`
dùng `Style="{StaticResource T3.DataGrid}"` — style đó đã set sẵn hai property bằng
`Setter`. Đã sửa: nhận diện `T3.DataGrid` / `T3.ListBox` / `T3.LogBox` là đã tuân thủ.

**Kiểm chứng:** XAML cố ý sai → bắt đủ 18 vi phạm, `exit 1`. XAML dựng từ đúng snippet
trong `xaml-templates.md` → `exit 0`. Tài liệu và gate khớp nhau.

---

## Cycle 0 — 2026-08-28 — BOOTSTRAP

```
DATE:                   2026-08-28
CYCLE:                  0 (bootstrap — thiết lập hệ thống governance, chưa audit)
TOOLS AUDITED:          0
ISSUES FOUND:           0 issue cấp tool
ISSUES FIXED:           0
ISSUES REMAINING:       0
SCORE CHANGES:          — (chưa có điểm nào)
REGRESSIONS:            —
SYSTEM-WIDE PATTERNS:   51/51 tool đang dùng hệ UI đã bị bỏ (Lumina) — toàn bộ là
                        legacy debt. 38–58 màu hex cứng mỗi file.
DESIGN SYSTEM GAPS:     #1 vị trí deploy T3Lab.Styles.xaml chưa chốt (CHẶN mọi migration)
                        #2 thiếu Style cho <Window> theo size class S/M/L
                        #3 copyright © T3Lab không có trong chuẩn T3 — giữ hay bỏ?
NEXT PRIORITIES:        Q1 viết dev/audit_t3.py (P1) · Q4 quét 3 lỗi P0 tiềm ẩn ·
                        Q2 audit sâu 6 tool đầu · Q3 migrate 2 file (bị GAP #1 chặn)
COMMIT:                 <điền sha sau khi commit>
```

**Nội dung cycle 0 — thiết lập, không sửa tool nào:**

- Chốt `pyRevit UI Design System/` là **chuẩn duy nhất**; bỏ Lumina, Revit-native,
  Terra v2, Kinetix.
- Dựng bộ tài liệu governance `docs/ui-governance/` (01→08 + BASELINE + CHANGELOG +
  PRIORITY_QUEUE + ROUTINE_PROMPT).
- Dựng agent `.claude/agents/ui-governance-agent.md`.
- Kiểm kê 54 XAML vào `BASELINE.md` — **không chấm điểm**, vì chấm điểm mà chưa audit
  là bịa số.

**Phát hiện quan trọng nhất của cycle 0:** `dev/audit_ui.py` gác chuẩn đã bị bỏ và sẽ
báo fail đúng những file được migrate cho đúng. Phải thay trước khi migration bắt đầu —
xem `PRIORITY_QUEUE.md` Q1.
