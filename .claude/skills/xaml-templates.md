# XAML snippets — chuẩn T3

Mọi snippet dưới đây dùng key có thật trong `pyRevit UI Design System/T3Lab.Styles.xaml`.
Luật đầy đủ: `pyRevit UI Design System/T3LAB_UI_STANDARD.md` ·
luật cho tool mới: `.claude/rules/new-tool-standard.md`.

Copy nguyên si. **Không sửa giá trị, không thêm hex, không định nghĩa `<Style>` trong
file tool.** Cần style chưa có → ghi `DESIGN SYSTEM GAP`, hỏi trước.

---

## Khung `<Window>`

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Tên tool" Width="560" Height="420" MinWidth="560" MinHeight="420"
        FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display"
        Background="{StaticResource T3.Canvas}"
        WindowStartupLocation="CenterOwner">
  <Window.Resources>
    <ResourceDictionary>
      <ResourceDictionary.MergedDictionaries>
        <ResourceDictionary Source="../Resources/T3Lab.Styles.xaml"/>
      </ResourceDictionary.MergedDictionaries>
    </ResourceDictionary>
  </Window.Resources>

  <DockPanel>
    <Border DockPanel.Dock="Top"    Style="{StaticResource T3.TitleBar}">…</Border>
    <Border DockPanel.Dock="Bottom" Style="{StaticResource T3.FooterBar}">…</Border>
    <Grid Margin="16">…</Grid>
  </DockPanel>
</Window>
```

## Title bar (40px)

```xml
<Border Style="{StaticResource T3.TitleBar}">
  <TextBlock Text="Rename Sheets" Style="{StaticResource T3.Title}"
             VerticalAlignment="Center"/>
</Border>
```

## Footer (48px) — trạng thái trái · nút phải

Thứ tự cố định: ghost huỷ → secondary → secondary → **MỘT** primary. Gap 8.

```xml
<Border Style="{StaticResource T3.FooterBar}">
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
      <Ellipse Width="6" Height="6" Margin="0,0,8,0" VerticalAlignment="Center"
               Fill="{StaticResource T3.Success.Accent}"/>
      <TextBlock x:Name="status_text" Text="Sẵn sàng — 34 sheet được chọn"
                 Style="{StaticResource T3.Body.Secondary}" VerticalAlignment="Center"/>
    </StackPanel>

    <StackPanel Grid.Column="1" Orientation="Horizontal">
      <Button Content="Cancel" IsCancel="True"
              Style="{StaticResource T3.Button.Ghost}" Margin="0,0,8,0"/>
      <Button Content="Preview" Style="{StaticResource T3.Button.Secondary}" Margin="0,0,8,0"/>
      <Button x:Name="btn_apply" Content="Rename 34 sheets" IsDefault="True"
              Style="{StaticResource T3.Button.Primary}"/>
    </StackPanel>
  </Grid>
</Border>
```

> Primary luôn **mang số đếm** khi thao tác có số lượng. "Rename 34 sheets", không phải "OK".

## Field — label TRÊN control, cách 4px; nhóm field cách nhau 12px

```xml
<StackPanel Margin="0,0,0,12">
  <TextBlock Text="Tiền tố" Style="{StaticResource T3.Label}" Margin="0,0,0,4"/>
  <TextBox x:Name="tb_prefix" Style="{StaticResource T3.TextBox}"/>
</StackPanel>
```

Ô nhập số / ID dùng `T3.TextBox.Mono`.

## ComboBox · CheckBox · RadioButton

```xml
<ComboBox x:Name="cb_level" Style="{StaticResource T3.ComboBox}"
          ItemContainerStyle="{StaticResource T3.ComboBoxItem}"/>

<CheckBox x:Name="chk_include" Content="Bao gồm sheet đã ẩn"
          Style="{StaticResource T3.CheckBox}"/>

<RadioButton x:Name="rb_all" Content="Toàn bộ model" GroupName="scope"
             Style="{StaticResource T3.RadioButton}"/>
```

## Section — label uppercase + `Separator`, KHÔNG dùng card

```xml
<TextBlock Text="PHẠM VI" Style="{StaticResource T3.Label}" Margin="0,0,0,4"/>
<Separator Style="{StaticResource T3.Rule}" Margin="0,0,0,12"/>
```

## Callout — hệ quả, luôn kèm số lượng

```xml
<Border Style="{StaticResource T3.Callout.Warning}">
  <TextBlock Style="{StaticResource T3.Body}" TextWrapping="Wrap"
             Text="12 sheet đang bị khoá bởi người khác và sẽ bị bỏ qua."/>
</Border>
```

Bốn biến thể: `T3.Callout` (trung tính) · `.Success` · `.Warning` · `.Danger`.

## ListBox (P2 — selection list)

`T3.ListBox` đã bật virtualization và tắt scroll ngang bằng Setter — **không khai lại**.

```xml
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
  </Grid.RowDefinitions>

  <TextBox x:Name="tb_filter" Grid.Row="0" Margin="0,0,0,8"
           Style="{StaticResource T3.TextBox}"/>

  <Grid Grid.Row="1">
    <ListBox x:Name="lst_items" Style="{StaticResource T3.ListBox}"
             SelectionMode="Extended"/>
    <TextBlock x:Name="lst_empty" Style="{StaticResource T3.Empty}"
               Visibility="Collapsed"
               Text="Không có sheet nào khớp bộ lọc.&#x0a;Xoá bớt từ khoá để xem toàn bộ."/>
  </Grid>
</Grid>
```

> **Empty state là bắt buộc.** Nói *thiếu gì* và *làm gì tiếp* — không phải "No data".
> **Không bao giờ** bọc `ListBox` trong `ScrollViewer`: mất virtualization, treo Revit.

## DataGrid (P4 — results table)

`T3.DataGrid` đã set row 26, virtualization, tắt scroll ngang, header/row/cell style.

```xml
<DataGrid x:Name="grid_results" Style="{StaticResource T3.DataGrid}">
  <DataGrid.Columns>
    <DataGridTextColumn Header="ID" Width="90" Binding="{Binding Id}"
                        ElementStyle="{StaticResource T3.Cell.Mono}"/>
    <DataGridTextColumn Header="Tên" Width="*" Binding="{Binding Name}"/>
    <DataGridTextColumn Header="Category" Width="140" Binding="{Binding Category}"/>
    <DataGridTextColumn Header="Số lượng" Width="70" Binding="{Binding Count}"
                        ElementStyle="{StaticResource T3.Cell.Number}"/>
    <DataGridTemplateColumn Header="Trạng thái" Width="150"
                            CellTemplate="{StaticResource T3.StatusPill}"/>
  </DataGrid.Columns>
</DataGrid>
```

Luật cột: **đúng một** cột `Width="*"` (là Name). Còn lại fix px —
`ID 90` · `Category 140` · `Status 150–170` · số `70`. Không đủ chỗ thì **bỏ bớt cột**,
không bật scroll ngang.

`T3.StatusPill` cần row VM có 2 property: `StatusText` (string) và
`Severity` (`"Success"` / `"Warning"` / `"Danger"`).

## Progress & log (P3)

```xml
<StackPanel>
  <TextBlock x:Name="lbl_phase" Style="{StaticResource T3.BodyStrong}"
             Text="Đang xuất — 12 / 34"/>
  <TextBlock x:Name="lbl_item" Style="{StaticResource T3.Caption}" Margin="0,4,0,8"
             Text="A-101 — Ground Floor Plan"/>

  <ProgressBar x:Name="bar" Style="{StaticResource T3.ProgressBar}"
               Minimum="0" Maximum="34" Value="12" Margin="0,0,0,12"/>

  <ListBox x:Name="log" Style="{StaticResource T3.LogBox}" Height="180"/>
</StackPanel>
```

Log ghi màu theo severity **và kèm chữ**: `ok` / `skipped` / `failed`. Màu một mình
không bao giờ là trạng thái.

## Expander — chỉ cho Advanced

```xml
<Expander Header="Tuỳ chọn nâng cao" Style="{StaticResource T3.Expander}">
  <StackPanel Margin="0,8,0,0">…</StackPanel>
</Expander>
```

## Confirmation (P5) — size S

```xml
<StackPanel Margin="16">
  <TextBlock Style="{StaticResource T3.Display}" TextWrapping="Wrap"
             Text="Xoá 34 view template?"/>
  <TextBlock Style="{StaticResource T3.Body.Secondary}" Margin="0,8,0,0" TextWrapping="Wrap"
             Text="Thao tác này không thể hoàn tác bằng Ctrl+Z sau khi đóng file."/>
</StackPanel>

<!-- footer -->
<Button Content="Cancel" IsCancel="True" IsDefault="True"
        Style="{StaticResource T3.Button.Secondary}" Margin="0,0,8,0"/>
<Button x:Name="btn_delete" Content="Xoá 34 template"
        Style="{StaticResource T3.Button.Danger}"/>
```

> Ở P5, **Cancel** là `IsDefault`, không phải nút phá huỷ. Nút phá huỷ mang **tên thao
> tác + số**, không phải "OK" hay "Yes".

---

## Bảng tra style key

| Cần gì | Key |
|--------|-----|
| Chữ | `T3.Display` `T3.Title` `T3.Body` `T3.Body.Secondary` `T3.BodyStrong` `T3.Caption` `T3.Label` `T3.Mono` |
| Empty state | `T3.Empty` |
| Cell trong grid | `T3.Cell.Mono` `T3.Cell.Number` |
| Nút | `T3.Button.Primary` `.Secondary` `.Ghost` `.Danger` |
| Nhập liệu | `T3.TextBox` `T3.TextBox.Mono` `T3.ComboBox` `T3.ComboBoxItem` `T3.CheckBox` `T3.RadioButton` |
| List / grid | `T3.ListBox` `T3.ListBoxItem` `T3.LogBox` `T3.DataGrid` `T3.DataGridColumnHeader` `T3.DataGridRow` `T3.DataGridCell` |
| Trạng thái | `T3.StatusPill` `T3.ProgressBar` |
| Khung | `T3.TitleBar` `T3.FooterBar` `T3.Panel` `T3.Rule` `T3.Expander` |
| Callout | `T3.Callout` `.Success` `.Warning` `.Danger` |

Màu/size/spacing chỉ dùng khi thật sự cần đặt trực tiếp: `T3.Ink` `T3.Text`
`T3.TextSecondary` `T3.TextMuted` `T3.TextDisabled` `T3.Border` `T3.BorderStrong`
`T3.Surface` `T3.SurfaceSunken` `T3.Canvas` `T3.RowAlt` `T3.RowRule` ·
`T3.Success.*` `T3.Warning.*` `T3.Danger.*` `T3.Progress.*` ·
`T3.Size.*` `T3.H.*` `T3.Pad.*` `T3.R.*`.

## Kiểm tra

```bash
python3 dev/audit_t3.py --file T3Lab.extension/lib/GUI/Tools/<Tool>.xaml
```
