# T3Lab Assistant — giao diện hoà làm 1 với Revit

> Cập nhật 2026-07-30. File liên quan:
> `T3Lab.extension/lib/GUI/RevitTheme.py` ·
> `T3Lab.extension/lib/GUI/Tools/T3LabAssistant.xaml` ·
> `T3Lab.extension/lib/GUI/AssistantPaneControl.py` ·
> `T3Lab.extension/T3Lab.tab/Support.panel/T3LabAssistant.pushbutton/script.py`

Mục tiêu: Assistant không còn là một cửa sổ lạ dán đè lên Revit, mà đọc như
**một panel của chính Revit**.

## 1. Bảng màu "paper" (cách chia màu kiểu chat)

Assistant là **chat surface**, không phải tool dialog — nên nó dùng bảng màu
riêng thay cho Lumina (`.claude/rules/ui-design-standard.md`). Đây là **ngoại lệ
có chủ đích**, đã ghi vào `UI_LOCKED` của `dev/audit_ui.py` và bảng UI-Frozen
trong `.claude/CLAUDE.md` để các đợt sweep UI sau không "sửa" ngược lại.

Cách chia màu (chỉ **một** bên hội thoại được tô nền):

| Thành phần | Vai trò |
|---|---|
| Nền chat | giấy ấm `#FAF9F5` — cả khung chat lẫn khung soạn tin dùng chung, không có đường cắt |
| Tin của **user** | bong bóng xám ấm `#F0EEE6`, bo 14px, canh phải |
| Trả lời của **assistant** | **không bong bóng, không viền, không avatar** — chữ nằm thẳng trên nền giấy |
| Accent | coral `#D97757` — nút gửi, chip đang chọn, số lượng selection |
| Copyright | vẫn amber `#F59E0B` theo chuẩn Lumina (audit kiểm tra) |

Không token màu nào được gõ thẳng vào chỗ dùng: tất cả đi qua
`GUI/RevitTheme.py`. XAML dùng `{DynamicResource T3Theme<Token>}`, code Python
dùng `_tb('Token')` / `_trgb('Token')` / `_theme_color('Token')`.

## 2. Theme sáng/tối theo Revit

`RevitTheme.current_theme()` đọc `UIThemeManager.CurrentTheme` (Revit 2024+;
host cũ hơn không có dark mode nên trả `light`). `RevitTheme.apply(window)` ghi
28 token vào `Window.Resources`, nên mọi chỗ bind `DynamicResource` đổi màu ngay.

- Áp dụng khi mở cửa sổ **và** mỗi lần cửa sổ được activate (user đổi theme
  giữa chừng trong Options ▸ User Interface, hoặc bằng tool BG Theme).
- **Giới hạn đã biết:** các tin nhắn *đã render* giữ nguyên màu lúc render.
  Chúng theo theme mới khi có tin mới, hoặc toàn bộ sau "New conversation" /
  mở lại pane. Render lại transcript giữa chừng có thể đụng vào bong bóng đang
  stream nên không làm.
- Card xác nhận thao tác phá huỷ (`#FEF2F2`) cố tình giữ nguyên đỏ sáng ở cả
  hai theme — nó phải nổi bật.

## 3. Dockable pane: đúng chỗ, đúng viền

- `AssistantPaneControl._apply_initial_dock_state()` đặt
  `InitialState.DockPosition = Right` và `TabBehind = ProjectBrowser`. Trước đây
  không set `InitialState` nên Revit mặc định thả pane **nổi giữa màn hình** —
  đúng cái ấn tượng "app rời" cần bỏ. Revit chỉ dùng `InitialState` ở lần hiện
  pane đầu tiên; sau đó tôn trọng vị trí user tự dock.
- Khi docked, `_flatten_chrome_for_pane()` bỏ bo góc 22px + viền 1.5px của
  `root_chrome` để Assistant lấp kín pane, không còn hở nền Revit ở 4 góc.
  Cụm nút min/max của cửa sổ nổi cũng ẩn (pane đã có nút riêng của Revit).

## 4. Layout hẹp

Pane dock cạnh Project Browser thường chỉ 300–340px. Dưới `COMPACT_WIDTH = 400`
(`_apply_compact_layout`): padding co lại, và footer bỏ dòng disclaimer + dòng
gợi ý phím tắt — **giữ copyright** (bắt buộc theo chuẩn).

Việc trả lời bỏ bong bóng + bỏ avatar cũng trả lại toàn bộ bề ngang cho nội
dung, đây là thay đổi đáng kể nhất cho pane hẹp.

## 5. Hàng icon dưới mỗi câu trả lời

Style `MsgActionBtn` (XAML) + `_make_message_actions()` (script.py). Icon line-art
xám, không nền; hover mới hiện nền bo tròn:

| Icon | Hành vi |
|---|---|
| Copy | chép câu trả lời vào clipboard; icon đổi thành dấu ✓ xanh 1.2s rồi trả lại |
| Try again | chạy lại đúng prompt đã sinh ra câu trả lời đó |
| 👍 / 👎 | chốt đánh giá (bấm lại để bỏ), ghi vào activity log |
| ‹ n / m › | chuyển qua lại giữa các phiên bản trả lời |

"Try again" **không** ghi đè câu trả lời cũ và **không** lặp lại tin nhắn của
user: câu trả lời mới được lưu thành **một version khác của cùng hàng đó**, nên
cặp mũi tên `‹ n/m ›` mới xuất hiện (giống `3/3` trong UI tham chiếu). Trạng thái
này nằm ở `row.Tag` (dict) — IronPython không cho gắn attribute lên object CLR
nên `Tag` là chỗ duy nhất mang được.

> Icon "read aloud" (loa) trong UI tham chiếu **không** làm: chưa có TTS trong
> extension, và một nút không chạy còn tệ hơn không có nút.

Avatar T3Lab xoay khi đang xử lý vẫn còn — nhưng chuyển vào **dòng typing**:
một mark động cho mỗi lượt, đặt đúng chỗ câu trả lời sắp hiện, thay vì mỗi bong
bóng một avatar.

## 6. Scrollbar mỏng kiểu Revit

`RevitThinScrollBar` + `RevitThinThumb`: không nút mũi tên, không nền — chỉ một
thumb bo tròn rộng 5px (7px khi hover/kéo), vùng bấm 8px.

Áp bằng **nested `<ScrollViewer.Resources>`** trên từng surface cuộn, không sửa
block shared styles: block đó khai báo `<Style TargetType="{x:Type ScrollBar}">`
*implicit* (không `x:Key`), thêm cái thứ hai cùng dictionary sẽ lỗi trùng key —
và sửa file master sẽ kéo theo 52 tool window khác.
