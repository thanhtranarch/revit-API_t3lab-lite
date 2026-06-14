# T3Lab UI Design System — Lumina Architecture System

This document is the official design specification for T3Lab Revit extension tool windows. It connects our codebase to design systems (like Google's Stitch) and outlines visual styles, color tokens, typography, layouts, and components.

> [!IMPORTANT]
> **Definitive Master Reference UI (Source of Truth)**:
> The official **UI Standard Showcase** window (defined in [UIStandardShowcase.xaml](file:///d:/01. T3Lab/02 Revit Tools/t3lab-revit-api/T3Lab.extension/lib/GUI/Tools/UIStandardShowcase.xaml) and launched via [UIStandardShowcaseDialog.py](file:///d:/01. T3Lab/02 Revit Tools/t3lab-revit-api/T3Lab.extension/lib/GUI/UIStandardShowcaseDialog.py)) is the master visual benchmark for the entire T3Lab extension. All existing and future tool designs, layout structures, spacing definitions, and component styling must align exactly with this showcase window.

---

## Design Core Principles
1. **Developer-Utility Aesthetic**: Premium dark mode accents, highly structured layout with minimal distractions.
2. **Visual Hierarchy**: Consistent, clear visual cards, statistics displays, and structured controls.
3. **No Fluff**: No custom logos or empty placeholders. Maximum space allocated for user data.
4. **Benchmark Reference**: The [UI Standard Showcase](file:///d:/01. T3Lab/02 Revit Tools/t3lab-revit-api/T3Lab.extension/lib/GUI/Tools/UIStandardShowcase.xaml) acts as the living style guide. Any styling modifications, new controls, or design improvements must be prototyped and approved in the Showcase before being applied to other extension tools.

---

## Design Tokens

### 1. Colors (Lumina Palette)
All tool windows must use the following unified color scheme:

| Token Name | Hex Code | WPF Usage |
|:---|:---|:---|
| `Primary` (Slate) | `#0F172A` | Window headers, primary buttons, chrome control icons |
| `Accent` (Blue) | `#3B82F6` | Selection highlights, input focus borders, primary action triggers |
| `Success` (Green) | `#10B981` | Confirmations, completions, success actions (Hover: `#059669`) |
| `Danger` (Red) | `#EF4444` | Deletion, cancellations, destructive actions (Hover: `#DC2626`) |
| `Warning` (Amber) | `#F59E0B` | Highlight banners, copyright notices, warnings |
| `Ink Text` | `#0F172A` | Major titles, headings, active text fields |
| `Muted Text` | `#64748B` | Subtitles, disabled text labels, inactive tabs |
| `Faint Text` | `#94A3B8` | Input placeholder text, disabled options |
| `Bg Light` | `#F8FAFC` | Main surface backgrounds, input background, status bar bg |
| `Border` | `#E2E8F0` | Divider lines, grid separators, border bounds |
| `Input Border` | `#CBD5E1` | TextBox and ComboBox borders |

### 2. Typography
- **Primary Font Family**: `Inter` (applied globally to the `<Window>` or `<Grid>` root).
- **Secondary Icon Font**: `Segoe MDL2 Assets` (strictly for title bar minimize/maximize/close icons).
- **Weights**:
  - Titles & Primary Buttons: `SemiBold` (WPF: `FontWeights.SemiBold` or `FontWeights.Bold`)
  - Labels & Body: `Medium` / `Normal`

---

## Window Structure

### 1. Window Dimensions & Chrome
- Standard Tool Window: `Height="680"`, `Width="1100"`.
- Multiline `<WindowChrome>` for consistent border corner radii (`CornerRadius="8"`, `CaptionHeight="64"`, `GlassFrameThickness="0"`, `ResizeBorderThickness="5"`).

### 2. Header Title Bar (Height: 64px)
- **Left**: Organization Name (`T3Lab` - 11px Bold slate `#0F172A`) + Tool Title (18px Bold `#0F172A`).
- **Below**: Subtle separator (1px `#E2E8F0`) + descriptive subtitle (10px Italic `#64748B`).
- **Right**: Minimise (`&#xE921;`), Maximise (`&#xE922;`), and Close (`&#xE8BB;`) buttons (Segoe MDL2 Assets, FontSize 10).

### 3. Status Bar (Height: 32px)
- **Background**: `#F8FAFC`, **Top Border**: 1px `#E2E8F0`, **Padding**: `14,8`.
- Accommodates operation summary count and a collapsible progress bar if running long tasks.

### 4. Copyright Notice
- Floating copyright overlay block strictly placed as the **last child of the root layout Grid**:
  ```xml
  <TextBlock Text="© Copyright by T3Lab" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8" Foreground="#F59E0B" FontSize="11" IsHitTestVisible="False" Panel.ZIndex="999"/>
  ```

---

## Component Guidelines

### 1. Buttons

| Style Key | Target Color | Text Color | Text Style | Radius |
|:---|:---|:---|:---|:---|
| `PrimaryButton` | `#0F172A` | `#FFFFFF` | SemiBold | `6` |
| `SecondaryButton` | `Transparent` (Border: `#0F172A`) | `#0F172A` | SemiBold | `6` |
| `SuccessButton` | `#10B981` | `#FFFFFF` | SemiBold | `6` |
| `DangerButton` | `#EF4444` | `#FFFFFF` | SemiBold | `6` |
| `AccentButton` | `#3B82F6` | `#FFFFFF` | SemiBold | `6` |

### 2. Inputs (TextBox & ComboBox)
- **Height**: `28px`
- **Padding**: `6,4`
- **Border**: `#CBD5E1`
- **Focus Border**: `#3B82F6` (BorderThickness: 2)

### 3. DataGrid
- Alternating row backgrounds (`#FFFFFF` and `#F8FAFC`).
- Gridlines visible horizontally (`#F1F5F9`).
- Selected Row Highlight: Surface press (`#E2E8F0`).

### 4. Banners & Expanders
- **Banners**: Light background (e.g. `#F0F9FF` for Info, `#ECFDF5` for Success) with matching borders, thin corner radius (`4`), and Segoe MDL2 iconography.
- **Expanders**: Bordered container (`#E2E8F0`) with padding `8` for grouped optional parameters.
