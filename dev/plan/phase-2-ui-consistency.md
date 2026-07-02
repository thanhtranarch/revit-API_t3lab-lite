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
- [x] Xem context 2 chỗ dùng màu: `#15803D` → `#10B981` (success), `#166534` → `#059669` (success hover). Ngữ cảnh xác nhận: ô kết quả "Test Connection" (success state), khớp đúng bảng migration trong `ui-design-standard.md`

### 3b. `T3LabAssistant.xaml` — copyright về footer-trái
- [x] Đọc layout vùng đáy cửa sổ chat (dòng ~2022–2109): copyright là child cuối của StackPanel bên trong `Border Grid.Row="4"` (Padding="16,8,16,10"), `HorizontalAlignment="Center"` — đây chính là footer/input-bar (Auto height, row cuối), không có status bar riêng
- [x] Không có status bar chuẩn (chat UI) → đổi TextBlock hiện tại sang `HorizontalAlignment="Left"`; không thêm margin trái 14px vì Border cha đã có `Padding="16,..."` bên trái — thêm margin nữa sẽ lệch khỏi rhythm 16px sẵn có của toàn khối input
- [x] Không đụng block shared styles

### 3c. `AssistantPane.xaml` — copyright về footer-trái
- [x] Dockable pane variant B (root Grid, 4 row: header/chat/quick-actions/input — **không có row footer riêng**, row cuối = input bar full-width) — **đã kiểm tra và phát hiện fallback overlay chuẩn (bottom-left) sẽ đè lên ô chat input** (input textbox chiếm gần hết chiều rộng cột 0 ở row cuối, overlay bottom-left rơi đúng vào góc dưới-trái của textbox)
- [x] **Quyết định (deviate có chủ đích khỏi fallback overlay mặc định):** thêm hẳn 1 row Auto mới (Row 4 — Footer) thay vì overlay đè lên nội dung, tránh xung đột layout. Copyright đặt trong `Border Grid.Row="4"` nền `#F8FAFC`, border-top `#E2E8F0`, `Padding="14,4"`, `TextBlock HorizontalAlignment="Left"` — vẫn đúng chuẩn amber/FontSize 11/trái, chỉ khác cách neo vị trí (footer bar thật thay vì overlay ảo) để đảm bảo không che control. Tốn thêm ~24px chiều cao dockable pane — chấp nhận được.
- [x] Pane là UI nhúng trong Revit dockable panel — đã kiểm tra góc trái-dưới: **không còn đè lên control nào** nhờ dùng row riêng thay vì overlay

### 3d. Kiểm tra bản copy UIStandardShowcase (từ GĐ1-F3)
- [x] `UIStandardShowcase.xaml` copy vào `lib/GUI/Tools/` pass `audit_ui.py` (0 vấn đề, đã verify lại cùng lượt chạy audit hôm nay)

### Chốt ngày 3
- [x] `python3 dev/audit_ui.py --quiet` → **exit 0** (`UI AUDIT: 0/54 file có vấn đề (chỉ file UI-locked) — OK`)
- [x] `python3 dev/sync_wpf_styles.py --check` → 0 lệch (53 file)
- [x] `python3 dev/audit_tools.py --quiet` → không phát sinh mới (`AUDIT: clean`)
- [x] Commit: `style: migrate T3LabAssistant to Lumina palette, move copyright to footer-left`

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

- **Ngày 3 — `AssistantPane.xaml`:** fallback overlay bottom-left chuẩn (theo `ui-design-standard.md`) sẽ đè lên ô chat input vì root Grid của pane này không có row footer riêng — row cuối (Row 3) là input bar full-width. Đã xử lý bằng cách thêm Row 4 (Auto height) làm footer thật thay vì dùng kỹ thuật overlay `Grid.RowSpan="99"`. Cần user xác nhận trực quan trong Revit ở Ngày 4 rằng dockable pane không bị vỡ layout / không quá chật do thêm ~24px chiều cao footer.
- Chưa phát sinh vấn đề khác trong Ngày 3.
