# Giai đoạn 2 — UI Consistency theo chuẩn Lumina (Ngày 3–4)

> Làm SAU giai đoạn 1. Branch gợi ý: `fix/phase-2-ui-consistency`.
> Tham chiếu: `.claude/rules/ui-design-standard.md` · Kiểm tra: `python3 dev/audit_ui.py --quiet`

## Kết quả UI review (2026-07-02, cập nhật sau khi chốt chuẩn footer-trái)

Audit 76 XAML trong `lib/GUI/Tools/` theo chuẩn Lumina:

| Hạng mục | Kết quả |
|----------|---------|
| Palette Lumina (không còn Kinetix/Terra) | ✅ trừ **T3LabAssistant.xaml** sót 2 màu Terra |
| Font (Hanken Grotesk/Inter, không Manrope/Segoe UI) | ✅ 100% |
| WindowChrome đủ 5 attribute, multi-line | ✅ 100% |
| Shared styles block đồng bộ | ✅ 76/76 (`sync_wpf_styles.py --check`) |
| Dot-notation Grid definitions (crash risk) | ✅ không có |
| Glyph chrome buttons (E921/E922/E8BB) | ✅ không phát hiện sai |
| **Copyright ở footer, bên trái** | ✅ 74/76 — codebase vốn đã nhất quán footer-trái (copyright → divider 1px → status text). Tài liệu `ui-design-standard.md` trước đây bắt overlay góc phải-dưới là **lỗi thời so với thực tế** và đã được cập nhật (2026-07-02) theo pattern chuẩn của `UIStandardShowcase.xaml`. |

→ Chỉ còn **2 file outlier**, đều thuộc T3LabAssistant:

| File | Vấn đề |
|------|--------|
| `T3LabAssistant.xaml` | Copyright căn **giữa** dưới ô chat input (dòng ~2109) + sót 2 màu Terra (`#15803D`, `#166534`) |
| `AssistantPane.xaml` (dockable pane, variant B) | Copyright theo chuẩn cũ: overlay góc **phải**-dưới (dòng ~1357) — chính là file duy nhất từng theo doc cũ |

`CadtoFloorLayerItem.xaml` không có copyright nhưng là **item template** (root là list-row `<Border>`) → được miễn trừ trong chuẩn mới và trong `audit_ui.py` (`COPYRIGHT_EXEMPT`).

**2 file UI-LOCKED tuyệt đối không sửa**: `DWGManagement.xaml`, `ExportManager.xaml`.

---

## Ngày 3 — Sửa outlier T3LabAssistant + chốt audit sạch

### 3a. `T3LabAssistant.xaml` — palette Terra
- [ ] Xem context 2 chỗ dùng màu: `#15803D` → `#10B981` (success), `#166534` → `#059669` (success hover). Nếu ngữ cảnh không phải nút Success thì đối chiếu bảng migration trong `ui-design-standard.md`

### 3b. `T3LabAssistant.xaml` — copyright về footer-trái
- [ ] Đọc layout vùng đáy cửa sổ chat (quanh dòng 2100–2115): copyright đang là child cuối của StackPanel chứa input, `HorizontalAlignment="Center"`
- [ ] Nếu cửa sổ **có status bar** dạng chuẩn → chuyển copyright vào đó theo snippet chuẩn (copyright → divider `#DCDCE0` → status text)
- [ ] Nếu **không có status bar** (chat UI full-height) → hoặc thêm status bar chuẩn (khuyến nghị, đồng bộ toàn codebase), hoặc tối thiểu đổi TextBlock hiện tại sang `HorizontalAlignment="Left"` + margin trái 14px
- [ ] Không đụng block shared styles

### 3c. `AssistantPane.xaml` — copyright về footer-trái
- [ ] Dockable pane variant B (root Grid, không status bar) → áp dụng **fallback Variant B** trong chuẩn mới: TextBlock cuối root Grid với `HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="14,0,0,8" IsHitTestVisible="False" Panel.ZIndex="999" Grid.RowSpan="99" Grid.ColumnSpan="99"`
- [ ] Pane là UI nhúng trong Revit dockable panel — kiểm tra góc trái-dưới không đè lên control nào

### 3d. Kiểm tra bản copy UIStandardShowcase (từ GĐ1-F3)
- [ ] `UIStandardShowcase.xaml` copy vào `lib/GUI/Tools/` phải pass `audit_ui.py` (bản gốc `.claude/standard/` đã đúng footer-trái)

### Chốt ngày 3
- [ ] `python3 dev/audit_ui.py --quiet` → **exit 0**
- [ ] `python3 dev/sync_wpf_styles.py --check` → 0 lệch
- [ ] `python3 dev/audit_tools.py --quiet` → không phát sinh mới
- [ ] Commit: `style: migrate T3LabAssistant to Lumina palette, move copyright to footer-left`

---

## Ngày 4 — Verify trực quan trong Revit

Checklist mỗi cửa sổ (mở lần lượt 41 tool + dialog phụ):
1. Copyright màu amber nằm **bên trái footer**, có divider 1px ngăn với status text, không che nội dung
2. Title bar: T3Lab + tên tool + subtitle + separator đúng chuẩn
3. Nút min/max/close hover đúng (close → nền đỏ chữ trắng)
4. Font toàn cửa sổ là Hanken Grotesk/Inter
5. Resize cửa sổ → footer/copyright không vỡ layout

- [ ] Panel 1 (4 cửa sổ + dialog phụ)
- [ ] Panel 2 (14 cửa sổ + dialog phụ)
- [ ] Panel 3 (4 cửa sổ)
- [ ] Panel 4 (6 cửa sổ)
- [ ] Panel 5 (4 cửa sổ)
- [ ] Panel 6 (9 cửa sổ + AssistantPane dockable)
- [ ] Screenshot mỗi cửa sổ lưu `dev/plan/screenshots/` (đặt tên `<panel>-<tool>.png`)
- [ ] File nào layout vỡ → ghi vào Phát sinh, sửa tay từng file
- [ ] Commit cuối: `style: visual QA pass for Lumina consistency`
- [ ] Cập nhật bảng tiến độ `dev/plan/README.md`

---

## Phát sinh trong giai đoạn 2

_(ghi tại đây: file bị vỡ layout sau fix, màu sai ngữ cảnh, v.v.)_
