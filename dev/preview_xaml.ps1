# Render a tool XAML to PNG without Revit.
#
# Visual QA used to require opening Revit. This loads the XAML into WPF exactly
# as pyRevit does, applies the palette from GUI/RevitTheme.py, and writes a
# light and a dark PNG. It already caught real defects (a collapsed star column,
# an empty checked check box) before anyone opened the host.
#
#   python3 dev/export_palette.py                     # writes palette.json
#   powershell -STA -File dev/preview_xaml.ps1 `
#       -Xaml .claude/standard/UIStandardShowcase.xaml `
#       -Palette palette.json -OutDir .
#
# Caveats: this renders the CLIENT AREA only — the native title bar is drawn by
# Windows and is not in the image. Event handlers are stripped before parsing,
# so nothing is interactive. Off-screen layout needs several measure/arrange
# passes before DataGrid star columns settle; that loop is below and is a
# property of the harness, not of the XAML.

param(
  [string]$Xaml,
  [string]$Palette,
  [string]$OutDir,
  [int]$W = 820,
  [int]$H = 560
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$ErrorActionPreference = 'Stop'

$src = Get-Content -Raw -Encoding UTF8 $Xaml

# XamlReader cannot resolve Click="handler" without a code-behind type.
$src = [regex]::Replace($src, '\s(Click|MouseDoubleClick|SelectionChanged)="[^"]*"', '')

$pal = Get-Content -Raw -Encoding UTF8 $Palette | ConvertFrom-Json

function Render-Theme([string]$theme, [string]$outPath) {
  $reader = New-Object System.IO.StringReader($src)
  $xr     = [System.Xml.XmlReader]::Create($reader)
  $win    = [Windows.Markup.XamlReader]::Load($xr)

  # sample rows, same shape as ShowcaseItem in UIShowcaseDialog.py
  $rows = @(
    [pscustomobject]@{id='104231'; name='01_Plan_Mat bang dinh vi cot'; category='Floor Plan';   status='Compliant'}
    [pscustomobject]@{id='104232'; name='02_Plan_Mat bang kich thuoc cot'; category='Floor Plan'; status='Compliant'}
    [pscustomobject]@{id='104235'; name='Detail_Chi tiet noi thep cot'; category='Detail View';  status='Compliant'}
    [pscustomobject]@{id='104240'; name='3D_Phoi canh tong the'; category='3D View';             status='Non-Compliant (Template missing)'}
    [pscustomobject]@{id='104245'; name='Section_Mat cat dung doc nha'; category='Section';      status='Compliant'}
    [pscustomobject]@{id='104250'; name='Elevation_Mat dung truc A-D'; category='Elevation';     status='Compliant'}
    [pscustomobject]@{id='104255'; name='Schedule_Thong ke cot thep dam'; category='Schedule';   status='Compliant'}
    [pscustomobject]@{id='104265'; name='Drafting_Chi tiet cau tao se no'; category='Drafting';  status='Non-Compliant (Naming standard)'}
    [pscustomobject]@{id='104272'; name='Elevation_Mat dung truc E-H'; category='Elevation';     status='Needs Review'}
  )
  $grid = $win.FindName('sample_grid')
  if ($grid) { $grid.ItemsSource = $rows }

  $dark = ($theme -eq 'dark')
  $chk = $win.FindName('chk_dark_preview')
  if ($chk) { $chk.IsChecked = $dark }
  $st = $win.FindName('status_text')
  if ($st) { $st.Text = 'Ready - palette 24/62 from Revit' }
  # a couple of rows selected, so the selection tint is visible in the shot
  if ($grid) { $grid.SelectedItems.Add($rows[3]) | Out-Null }

  # detach the content so it can be rendered off-screen, and carry the
  # window's ResourceDictionary with it
  $content = $win.Content
  $win.Content = $null
  $shell = New-Object Windows.Controls.Border
  $shell.Resources = $win.Resources
  $shell.Child = $content

  # this is what RevitTheme.apply() does at runtime
  $table = if ($dark) { $pal.dark } else { $pal.light }
  $bc = New-Object Windows.Media.BrushConverter
  foreach ($p in $table.PSObject.Properties) {
    $b = [Windows.Media.Brush]$bc.ConvertFromString([string]$p.Value)
    $b.Freeze()
    $shell.Resources[('T3Theme' + $p.Name)] = $b
  }
  $shell.Background = [Windows.Media.Brush]$bc.ConvertFromString([string]$table.DialogBg)

  # mirror RevitTheme._fallback_check_glyph(): outside Revit there is no
  # CheckBoxCheckedImage to borrow, so the drawn stand-in is what shows.
  $grp = New-Object Windows.Media.DrawingGroup
  $sq  = New-Object Windows.Media.GeometryDrawing
  $sq.Brush = [Windows.Media.Brush]$bc.ConvertFromString([string]$table.CheckFill)
  $sq.Geometry = New-Object Windows.Media.RectangleGeometry((New-Object Windows.Rect(0,0,16,16)), 3, 3)
  $grp.Children.Add($sq) | Out-Null
  $fig = New-Object Windows.Media.PathFigure
  $fig.StartPoint = New-Object Windows.Point(3.6, 8.4)
  $fig.Segments.Add((New-Object Windows.Media.LineSegment((New-Object Windows.Point(6.7, 11.4)), $true))) | Out-Null
  $fig.Segments.Add((New-Object Windows.Media.LineSegment((New-Object Windows.Point(12.3, 5.2)), $true))) | Out-Null
  $pg = New-Object Windows.Media.PathGeometry
  $pg.Figures.Add($fig) | Out-Null
  $pen = New-Object Windows.Media.Pen(([Windows.Media.Brush]$bc.ConvertFromString([string]$table.OnAccent)), 1.9)
  $pen.StartLineCap = [Windows.Media.PenLineCap]::Round
  $pen.EndLineCap   = [Windows.Media.PenLineCap]::Round
  $tk = New-Object Windows.Media.GeometryDrawing
  $tk.Geometry = $pg
  $tk.Pen = $pen
  $grp.Children.Add($tk) | Out-Null
  $img = New-Object Windows.Media.DrawingImage($grp)
  $img.Freeze()
  $shell.Resources['T3ThemeCheckedGlyph'] = [Windows.Media.ImageSource]$img
  # mirror RevitTheme._progress_fill(): the Aero green vertical gradient
  $gs = New-Object Windows.Media.GradientStopCollection
  $gs.Add((New-Object Windows.Media.GradientStop(([Windows.Media.Color][Windows.Media.ColorConverter]::ConvertFromString([string]$table.ProgressHi)),  0.0))) | Out-Null
  $gs.Add((New-Object Windows.Media.GradientStop(([Windows.Media.Color][Windows.Media.ColorConverter]::ConvertFromString([string]$table.ProgressMid)), 0.45))) | Out-Null
  $gs.Add((New-Object Windows.Media.GradientStop(([Windows.Media.Color][Windows.Media.ColorConverter]::ConvertFromString([string]$table.ProgressLo)),  0.5))) | Out-Null
  $gs.Add((New-Object Windows.Media.GradientStop(([Windows.Media.Color][Windows.Media.ColorConverter]::ConvertFromString([string]$table.ProgressMid)), 1.0))) | Out-Null
  $lg = New-Object Windows.Media.LinearGradientBrush($gs, (New-Object Windows.Point(0,0)), (New-Object Windows.Point(0,1)))
  $lg.Freeze()
  $shell.Resources['T3ThemeProgressFill'] = [Windows.Media.Brush]$lg

  [Windows.Documents.TextElement]::SetFontFamily($shell, (New-Object Windows.Media.FontFamily 'Segoe UI'))
  [Windows.Documents.TextElement]::SetFontSize($shell, 12)
  [Windows.Documents.TextElement]::SetForeground($shell,
      ([Windows.Media.Brush]$bc.ConvertFromString([string]$table.Ink)))

  $size = New-Object Windows.Size($W, $H)
  $rect = New-Object Windows.Rect(0, 0, $W, $H)
  # DataGrid star columns settle against the ScrollViewer viewport, which is only
  # known after an arrange pass. One measure/arrange is not enough off-screen.
  for ($i = 0; $i -lt 4; $i++) {
    $shell.Measure($size)
    $shell.Arrange($rect)
    $shell.UpdateLayout()
    [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [Windows.Threading.DispatcherPriority]::Background, [action]{}) | Out-Null
  }

  $scale = 2
  $rtb = New-Object Windows.Media.Imaging.RenderTargetBitmap(
      ($W * $scale), ($H * $scale), (96 * $scale), (96 * $scale),
      [Windows.Media.PixelFormats]::Pbgra32)
  $rtb.Render($shell)

  $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
  $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
  $fs = [System.IO.File]::Create($outPath)
  $enc.Save($fs)
  $fs.Close()
  Write-Output ("rendered {0} -> {1}" -f $theme, $outPath)
}

Render-Theme 'light' (Join-Path $OutDir 'showcase_light.png')
Render-Theme 'dark'  (Join-Path $OutDir 'showcase_dark.png')







