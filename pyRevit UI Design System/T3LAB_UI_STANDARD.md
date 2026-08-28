# T3Lab pyRevit UI Standard — project instructions

Mọi dialog WPF/XAML sinh ra trong project này PHẢI tuân theo standard dưới đây.
Bản đặc tả đầy đủ (có mock + lý do từng quyết định): `T3Lab pyRevit UI Standard.dc.html`
File dùng thật: `dist/T3Lab.Styles.xaml` · `dist/T3LAB_UI_STANDARD.md`

## Ràng buộc kỹ thuật
- Chỉ stock WPF: Grid, StackPanel, DockPanel, UniformGrid, ListBox, ListView, ComboBox,
  CheckBox, RadioButton, TextBox, DataGrid, ProgressBar, Button, Expander, TabControl, Separator.
- Không thư viện thứ ba (MahApps, HandyControl, custom DLL).
- IronPython 2.7 / .NET 4.8 / `pyrevit.forms.WPFWindow`.
- Không effect: không gradient, không blur, không transition, không custom scrollbar.
- Phải đọc được ở 100% và 125% display scaling.

## Token — không tool nào được tự định nghĩa màu/size/margin
Màu: `T3.Ink #18181B` · `T3.Text #27272A` · `T3.TextSecondary #52525B` · `T3.TextMuted #71717A`
· `T3.TextDisabled #9A9AA2` · `T3.Border #DCDCE0` · `T3.BorderStrong #A1A1AA`
· `T3.SurfaceSunken #F4F4F6` · `T3.Canvas #E4E4E7` · `T3.Surface #FFFFFF`
Status (Fill / Accent / Text): Success `#EAF8F0 #22A85C #157038` · Warning `#FFF6E6 #F5CE5A #8A6308`
· Danger `#FDECEC #F87171 #D23B3B` · Progress track `#ECECEF`, fill `#C2410C`.
Cam `#C2410C` CHỈ dùng cho progress / trạng thái đang chạy. Không dùng làm button, border, brand.
Amber `T3.Copyright #F59E0B` CHỈ dùng cho dòng copyright ở footer — xem luật 11.

Type — Segoe UI, đúng 7 size: Display 19 SemiBold · Title 15 SemiBold · Body 13 Regular
· BodyStrong 13 SemiBold · Caption 11.5 · Label 11 SemiBold uppercase · Mono 12.5 Consolas.
Không dùng size ngoài danh sách. Không dùng font khác Segoe UI / Consolas.

Spacing: chỉ 4 / 8 / 12 / 16 / 24 / 32. Mọi margin chia hết cho 4.
Chiều cao: row 26 · control 28 · action button 30 · title bar 40 · footer 48.
Bo góc: window 8 · control 4 · pill 2 · grid row 0.
Kích thước cửa sổ: S 420×260–320 (NoResize) · M 560×420–560 · L 1000×620.

## Luật bố cục
1. Một left rail duy nhất: 16px từ mép window, 12px trong panel lồng.
2. Label nằm TRÊN control (4px), nhóm field cách nhau 12px. Không label cột bên trái.
3. Mỗi grid chỉ một cột `*` (là Name). Cột khác fix px: ID 90, Category 140, Status 150–170, số 70.
4. `HorizontalScrollBarVisibility="Disabled"` mọi grid/list. Không đủ chỗ thì bỏ bớt cột.
5. Số căn phải (Consolas), text căn trái, Element ID căn trái Consolas.
6. Footer cố định: trái = dot + câu trạng thái; phải = ghost huỷ → secondary → secondary → MỘT primary. Gap 8.
7. Panel lồng tối đa 2 cấp. Chia section bằng label uppercase + `Separator`, không bằng card.
8. `UseLayoutRounding="True"`, `SnapsToDevicePixels="True"`, `TextOptions.TextFormattingMode="Display"`
   trên Window. Luôn có MinWidth/MinHeight. Không set Height cho TextBlock. Không fix Width cho text dịch.
9. Mọi window có `IsDefault` và `IsCancel`. Focus = viền 1px `T3.Ink`, không dùng dotted rectangle.
10. Mọi list/grid có empty state: TextBlock canh giữa, `T3.TextDisabled`, nói thiếu gì và làm gì tiếp.
11. **Copyright BẮT BUỘC.** Mọi cửa sổ tool có ĐÚNG MỘT dòng `© Copyright by T3Lab`,
    dùng `{StaticResource T3.Copyright}`, đặt ở footer **sát trái**, đứng trước câu
    trạng thái và cách nó bằng một `Border` dọc 1×14 `T3.Border`, margin `12,0`.
    Ký tự `©` phải là ký tự thật — không dùng `&#169;`. Không right-align, không
    canh giữa, không thả nổi đè lên nội dung, không lặp lại.
    Miễn trừ duy nhất: XAML item-template (root là `<Border>`/`<DataTemplate>` của một
    dòng list, không phải cửa sổ) — copyright thuộc về cửa sổ chứa nó.
12. **Ngôn ngữ UI là TIẾNG ANH.** Mọi chữ người dùng đọc — label, nút, header, tooltip,
    empty state, thông báo lỗi/cảnh báo/thành công — viết bằng tiếng Anh, không ngoại lệ.
    Comment trong XAML/code và tài liệu nội bộ thì không bị ràng buộc.

## 5 pattern — mọi tool phải là một trong số này
- **P1 Parameter input form** (M) — form một cột, Expander cho Advanced, callout hệ quả có số lượng trên footer.
- **P2 Element selection list** (M) — filter pinned, ListBox virtualized, All/None ghost ở footer, primary mang số đếm.
- **P3 Progress & log** (M) — phase + `n / total` + item hiện tại, bar 8px cam, log ListBox Consolas
  màu theo severity + chữ (`ok` / `skipped` / `failed`), dải tally, footer "đang chạy, đừng đóng Revit".
  Xong thì Cancel → Close, dot chuyển xanh.
- **P4 Results table** (L) — dải summary 52px (số + label, chia bằng rule 1px), chip filter, DataGrid 26px row,
  status pill nền tint + dot + chữ, primary trả người dùng về Revit ("Select in Revit").
- **P5 Confirmation** (S) — headline là câu hỏi có số ("Delete 34 view templates?"), nút phá huỷ màu đỏ
  mang tên thao tác + số, và KHÔNG phải `IsDefault` (Cancel mới là). Bản success dùng lại vỏ này ở thì quá khứ.

## Chrome cửa sổ & wizard — 8 component (thêm 2026-08-28)

Tool nhiều bước có sidebar rail không khớp P1–P5. Thay vì tự chế pattern thứ 6,
dùng 8 component này; chúng mang sẵn giá trị hình dạng riêng nên file tool không
bao giờ phải tự viết số lạ:

`T3.WinCtrl` nút minimize/maximize · `T3.WinClose` nút đóng (hover đỏ) ·
`T3.TabItem.Hidden` tab ẩn cho wizard điều khiển bằng code-behind ·
`T3.Rail.Tile` ô rail 42×42 (ToggleButton) · `T3.Rail.Logo` ô logo 42×42 ·
`T3.Pill` dải ngang 36px · `T3.ComboBox.Toggle` nút mở của combo nhỏ ·
`T3.ScrollBar.Thumb` thumb thanh cuộn.

Implicit style (`<Style TargetType="X">` **không** có `x:Key`) được phép trong file
tool: đó là override chỉ có hiệu lực trong đúng container chứa nó, và stock WPF không
có cách nào khác. Style **có** `x:Key` thì không — nó là component mới, phải vào
`T3Lab.Styles.xaml`.

## Checklist review (dán vào PR)
Merge `T3Lab.Styles.xaml`, không tự định nghĩa brush · đúng một trong P1–P5 · size S/M/L
· có đúng một dòng copyright ở footer trái · mọi chữ hiển thị là tiếng Anh
· Segoe UI 13, không size lạ · margin chia hết 4 · đúng một primary, ngoài cùng phải
· có `IsDefault` + `IsCancel` · grid một cột `*`, tắt scroll ngang · có empty state
· chụp màn hình 100% và 125% không cắt chữ · thao tác phá huỷ có P5 · status không bao giờ chỉ bằng màu.

## Khi được nhờ viết tool mới
Hỏi tool thuộc pattern nào, rồi sinh XAML dùng `{StaticResource T3.*}` — không hardcode hex/size/margin.
Không tạo style mới trong file tool; style mới phải thêm vào `T3Lab.Styles.xaml`.
