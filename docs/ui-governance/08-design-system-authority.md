# 08 · Design System Authority — một chuẩn duy nhất

> **Đọc file này trước khi sửa bất kỳ XAML nào.**

---

## 1 · Chuẩn duy nhất

```
pyRevit UI Design System/
├── T3LAB_UI_STANDARD.md    ← luật: token, type, spacing, 5 pattern, luật bố cục
└── T3Lab.Styles.xaml       ← 82 resource key T3.* — nguồn duy nhất của màu/size/style
```

**Đây là nguồn chuẩn cao nhất và là nguồn chuẩn duy nhất.**
Mọi hệ UI khác trong repo đã bị **bỏ**. Không có "track", không có ngoại lệ theo file,
không có chuẩn song song.

## 2 · Những gì đã bị bỏ (LEGACY — không còn hiệu lực)

| Hệ đã bỏ | Đặc điểm nhận dạng | Xử lý khi gặp |
|----------|--------------------|----------------|
| **Lumina** | `FontFamily="Hanken Grotesk"` / `"Inter"`, nền `White`, hex cứng `#0F172A` `#3B82F6` `#E2E8F0` `#CBD5E1` `#64748B` `#94A3B8` `#F8FAFC` `#10B981` `#EF4444`, block `T3LAB SHARED STYLES` | Ghi là **LEGACY DEBT**, migrate sang `T3.*` theo hàng đợi |
| **Revit-native** | `{DynamicResource T3Theme*}`, `RevitTheme.py`, Segoe UI 12, `WindowChrome CaptionHeight="56"` | Như trên |
| **Terra v2** | `#0F766E` `#115E59` `#0B4F4A` `#F6F8F8` `#E6EDEC` `#D9EBE8` `#DDE5E7` `#C7D2D4` `#15803D` `#166534` `#AEBFBC` | Như trên, ưu tiên cao hơn (chết lâu nhất) |
| **Kinetix** | `#083D56` | Như trên |

Các file luật cũ (`.claude/rules/ui-design-standard.md` và mọi mô tả Lumina trong
agent definitions) **không được dùng làm căn cứ chấm điểm nữa**. Chúng chỉ còn giá trị
tra cứu để *nhận dạng* code cũ.

## 3 · Thứ tự thẩm quyền

```
1. Ràng buộc kỹ thuật & an toàn Revit  (IronPython 2.7, .NET 4.8, virtualization, transaction)
2. pyRevit UI Design System/           ← chuẩn thiết kế duy nhất
3. Sở thích cá nhân                    ← không bao giờ
```

Chỉ mục 1 được phép thắng mục 2, và chỉ khi có lý do kỹ thuật viết ra được. Ví dụ hợp
lệ: một list phải bật virtualization dù Design System không nói tới — vì không bật thì
treo Revit. Ví dụ **không** hợp lệ: "để nguyên màu cũ cho đỡ khác biệt".

---

## 4 · Tóm tắt luật T3 (bản tra nhanh — bản đầy đủ ở `T3LAB_UI_STANDARD.md`)

### 4.1 · Ràng buộc kỹ thuật

- Chỉ stock WPF. Không thư viện thứ ba (MahApps, HandyControl, custom DLL).
- IronPython 2.7 / .NET 4.8 / `pyrevit.forms.WPFWindow`.
- **Không effect**: không gradient, không blur, không transition, không custom scrollbar.
- Đọc được ở 100% **và** 125% display scaling.

### 4.2 · Token — không tool nào được tự định nghĩa màu/size/margin

| Nhóm | Token |
|------|-------|
| Chữ | `T3.Ink #18181B` · `T3.Text #27272A` · `T3.TextSecondary #52525B` · `T3.TextMuted #71717A` · `T3.TextDisabled #9A9AA2` |
| Viền | `T3.Border #DCDCE0` · `T3.BorderStrong #A1A1AA` |
| Nền | `T3.Surface #FFFFFF` · `T3.SurfaceSunken #F4F4F6` · `T3.Canvas #E4E4E7` · `T3.RowAlt #FAFAFB` · `T3.RowRule #F1F1F4` |
| Success | `Fill #EAF8F0` · `Accent #22A85C` · `Text #157038` |
| Warning | `Fill #FFF6E6` · `Accent #F5CE5A` · `Text #8A6308` |
| Danger | `Fill #FDECEC` · `Accent #F87171` · `Text #D23B3B` · `Border #F6D5D5` |
| Progress | `Track #ECECEF` · `Fill #C2410C` |

> Cam `#C2410C` **CHỈ** dùng cho progress / trạng thái đang chạy. Không làm button,
> không làm border, không làm brand.

### 4.3 · Type — Segoe UI, đúng 7 size

`Display 19 SemiBold` · `Title 15 SemiBold` · `Body 13 Regular` · `BodyStrong 13 SemiBold`
· `Caption 11.5` · `Label 11 SemiBold uppercase` · `Mono 12.5 Consolas`

Không size ngoài danh sách. Không font khác Segoe UI / Consolas.

### 4.4 · Spacing · chiều cao · bo góc · size window

- Spacing: chỉ `4 / 8 / 12 / 16 / 24 / 32`. Mọi margin chia hết cho 4.
- Chiều cao: row `26` · control `28` · action button `30` · title bar `40` · footer `48`.
- Bo góc: window `8` · control `4` · pill `2` · grid row `0`.
- Size window: `S 420×260–320` (NoResize) · `M 560×420–560` · `L 1000×620`.

### 4.5 · 10 luật bố cục

| # | Luật |
|---|------|
| L1 | Một left rail duy nhất: 16px từ mép window, 12px trong panel lồng |
| L2 | Label nằm **TRÊN** control (4px); nhóm field cách nhau 12px. Không label cột bên trái |
| L3 | Mỗi grid chỉ **một** cột `*` (là Name). Cột khác fix px: ID 90 · Category 140 · Status 150–170 · số 70 |
| L4 | `HorizontalScrollBarVisibility="Disabled"` mọi grid/list. Không đủ chỗ thì **bỏ bớt cột** |
| L5 | Số căn phải (Consolas) · text căn trái · Element ID căn trái Consolas |
| L6 | Footer cố định: trái = dot + câu trạng thái; phải = ghost huỷ → secondary → secondary → **MỘT** primary. Gap 8 |
| L7 | Panel lồng tối đa 2 cấp. Chia section bằng label uppercase + `Separator`, **không** bằng card |
| L8 | `UseLayoutRounding="True"` · `SnapsToDevicePixels="True"` · `TextOptions.TextFormattingMode="Display"` trên Window. Luôn có `MinWidth`/`MinHeight`. Không set `Height` cho `TextBlock`. Không fix `Width` cho text dịch |
| L9 | Mọi window có `IsDefault` **và** `IsCancel`. Focus = viền 1px `T3.Ink`, không dotted rectangle |
| L10 | Mọi list/grid có **empty state**: TextBlock canh giữa, `T3.TextDisabled`, nói thiếu gì và làm gì tiếp |

### 4.6 · 5 pattern — mọi tool phải là một trong số này

| Pattern | Size | Đặc trưng bắt buộc |
|---------|------|--------------------|
| **P1** Parameter input form | M | Form một cột · `Expander` cho Advanced · callout hệ quả **có số lượng** trên footer |
| **P2** Element selection list | M | Filter pinned · `ListBox` virtualized · All/None ghost ở footer · primary **mang số đếm** |
| **P3** Progress & log | M | Phase + `n / total` + item hiện tại · bar 8px cam · log `ListBox` Consolas màu theo severity **kèm chữ** (`ok`/`skipped`/`failed`) · dải tally · footer "đang chạy, đừng đóng Revit" · xong thì Cancel → Close, dot xanh |
| **P4** Results table | L | Dải summary 52px (số + label, chia bằng rule 1px) · chip filter · `DataGrid` row 26 · status pill nền tint + dot + chữ · primary trả người dùng về Revit ("Select in Revit") |
| **P5** Confirmation | S | Headline là **câu hỏi có số** ("Delete 34 view templates?") · nút phá huỷ đỏ mang tên thao tác + số · nút đó **KHÔNG** `IsDefault` (Cancel mới là) · bản success dùng lại vỏ này ở thì quá khứ |

### 4.7 · Checklist review (dán vào PR)

Merge `T3Lab.Styles.xaml`, không tự định nghĩa brush · đúng một trong P1–P5 · size S/M/L
· Segoe UI 13, không size lạ · margin chia hết 4 · đúng một primary, ngoài cùng phải
· có `IsDefault` + `IsCancel` · grid một cột `*`, tắt scroll ngang · có empty state
· chụp màn hình 100% và 125% không cắt chữ · thao tác phá huỷ có P5 · status không bao
giờ chỉ bằng màu.

### 4.8 · Khi viết tool mới

Hỏi tool thuộc pattern nào, rồi sinh XAML dùng `{StaticResource T3.*}` — không hardcode
hex/size/margin. **Không tạo style mới trong file tool**; style mới phải thêm vào
`T3Lab.Styles.xaml` (và việc thêm style là `MANUAL REVIEW REQUIRED`).

---

## 5 · Cách merge stylesheet vào một tool

```xml
<Window x:Class="..." FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display"
        Background="{StaticResource T3.Canvas}"
        MinWidth="560" MinHeight="420">
  <Window.Resources>
    <ResourceDictionary>
      <ResourceDictionary.MergedDictionaries>
        <ResourceDictionary Source="../../lib/T3Lab.Styles.xaml"/>
      </ResourceDictionary.MergedDictionaries>
    </ResourceDictionary>
  </Window.Resources>
  ...
</Window>
```

> `Source` là đường dẫn tương đối từ file XAML tới bản `T3Lab.Styles.xaml` được deploy
> trong extension. Vị trí deploy chính thức của file style **chưa được chốt**
> (`pyRevit UI Design System/` là thư mục nguồn, không nằm trong `T3Lab.extension/`) —
> xem `DESIGN SYSTEM GAP #1` trong `PRIORITY_QUEUE.md`.

---

## 6 · Migration — luật vàng của giai đoạn chuyển tiếp

Hiện trạng: **0/54 XAML đạt chuẩn T3**. Toàn bộ là legacy debt. Điều đó là bình thường
và không phải lỗi của cycle nào — nhiệm vụ của routine là kéo con số đó xuống, đều đặn,
không phải trong một đêm.

### 6.1 · Hai loại công việc

| Loại | Nội dung | Auto-fix? |
|------|----------|-----------|
| **In-place fix** | Sửa lỗi T3 mà không đổi hệ màu/font của file: spacing chia 4, một primary, `IsDefault`/`IsCancel`, empty state, tắt scroll ngang, virtualization, bỏ `Effect`, dot-notation crash, status có chữ | ✅ Có — đây là công việc thường ngày |
| **Full migration** | Chuyển hẳn một file sang `T3.*` + Segoe UI 13 + pattern P1–P5 + size class | ⚠️ Có, nhưng **tối đa 2 file/cycle**, mỗi file phải verify riêng |

Lý do tách: in-place fix rủi ro thấp, làm được hàng loạt. Full migration đụng toàn bộ
visual của một cửa sổ nên **bắt buộc có QA trong Revit** trước khi đánh `FIXED`.

### 6.2 · Quy trình full migration một file

1. Xác định pattern (P1–P5) và size class (S/M/L). Không rõ → `NEEDS REVIEW`, dừng.
2. Ghi lại mọi `x:Name` đang có → script.py bám vào chúng. **Không đổi tên, không xoá.**
3. Ghi lại mọi `Click=` / `SelectionChanged=` / binding → phải còn nguyên sau migrate.
4. Thay style: merge `T3Lab.Styles.xaml`, xoá mọi brush/style tự định nghĩa, xoá block
   `T3LAB SHARED STYLES` của Lumina.
5. Áp 10 luật bố cục (§4.5) + yêu cầu riêng của pattern (§4.6).
6. Verify theo `05-improvement-and-safety.md` — bao gồm **đối chiếu `x:Name`** trước/sau.
7. Đánh `NEEDS VERIFICATION` cho tới khi user xác nhận đã mở trong Revit.

### 6.3 · Không bao giờ

- Đổi hoặc xoá `x:Name` mà script.py đang dùng.
- Đổi tên event handler.
- Xoá một control chỉ vì Design System không có chỗ cho nó → `NEEDS REVIEW`.
- Migrate file UI-locked (§7).
- Migrate quá 2 file trong một cycle.

## 7 · File UI-locked — bỏ qua hoàn toàn

Ba file dưới đây có thiết kế đã chốt riêng và **không thuộc phạm vi routine**:

- `T3Lab.extension/lib/GUI/Tools/DWGManagement.xaml`
- `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml`
- `T3Lab.extension/lib/GUI/Tools/T3LabAssistant.xaml`

Liệt kê trong BASELINE với trạng thái `LOCKED`, điểm `—`, không tính vào Overall Score.
Việc gỡ khoá là quyết định của user, không phải của routine.

## 8 · Regression gate — trạng thái sau khi bỏ luật cũ

| Script | Chuẩn nó gác | Còn dùng? |
|--------|--------------|-----------|
| `dev/audit_tools.py` | Code tĩnh (import, syntax, dead code) — trung lập với Design System | ✅ Giữ, chạy mỗi cycle |
| `dev/audit_ui.py` | **Lumina** — cấm Segoe UI, bắt buộc Hanken Grotesk + block shared styles | ❌ **Đã lỗi thời.** Nó sẽ báo fail đúng những file đã migrate sang T3. Không dùng làm gate nữa |
| `dev/sync_wpf_styles.py --check` | Đồng bộ block shared styles **Lumina** | ❌ Lỗi thời với file đã migrate; vẫn đúng với file legacy chưa migrate |

**Việc cần làm (P1, đã nằm ở đầu `PRIORITY_QUEUE.md`):** viết `dev/audit_t3.py` enforce
luật T3 (§4), rồi rút `audit_ui.py` + `sync_wpf_styles.py` khỏi gate khi file legacy
cuối cùng được migrate. Cho tới lúc đó:

- Chạy `audit_ui.py` để **đếm số file legacy còn lại**, không để pass/fail chặn công việc.
- File đã migrate sang T3 mà `audit_ui.py` báo lỗi Lumina → **kết quả đúng như dự kiến**,
  ghi vào report là *expected legacy-gate noise*, không sửa ngược lại.
