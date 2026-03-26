# -*- coding: utf-8 -*-
"""
Annotation Manager
------------------
Unified tool combining Dimension and Text Note management:
  - Find elements by keyword → jump to view
  - Delete selected instances / types
  - Auto-rename all types based on their properties

Author: T3Lab (Tran Tien Thanh)
"""

__title__  = "Annotation\nManager"
__author__ = "T3Lab"

import re
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

from System.Windows import Visibility
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

doc   = revit.doc
uidoc = revit.uidoc

# ============================================================
# SHARED COLOR TABLE
# ============================================================
_DIM_COLORS = {
    (255,128,128):"Light Coral",(255,255,128):"Light Yellow",(128,255,128):"Pale Green",
    (128,255,255):"Pale Cyan",(128,128,255):"Light Slate Blue",(255,128,255):"Orchid",
    (255,0,0):"Red",(255,255,0):"Yellow",(0,255,0):"Lime",(0,255,255):"Cyan",
    (0,0,255):"Blue",(255,0,255):"Magenta",(128,64,64):"Brown",(255,192,128):"Light Salmon",
    (128,255,192):"Aquamarine",(192,192,255):"Lavender",(192,128,255):"Medium Orchid",
    (128,0,0):"Maroon",(255,128,0):"Orange",(0,128,0):"Green",(0,128,128):"Teal",
    (0,0,128):"Navy",(128,0,128):"Purple",(128,64,0):"Saddle Brown",(192,128,64):"Peru",
    (0,128,64):"Dark Sea Green",(0,128,192):"Steel Blue",(64,128,255):"Dodger Blue",
    (128,0,192):"Dark Orchid",(0,0,0):"Black",(128,128,0):"Olive",(128,128,128):"Gray128",
    (0,192,192):"Medium Turquoise",(192,192,192):"Silver",(255,255,255):"White",
    (70,70,70):"Gray70",(128,0,64):"Dark Raspberry",(77,77,77):"Gray77",
}
_TXT_COLORS = {
    (255,0,0):"Red",(0,255,0):"Lime",(0,0,255):"Blue",(255,255,0):"Yellow",
    (0,255,255):"Cyan",(255,0,255):"Magenta",(0,0,0):"Black",(255,255,255):"White",
    (128,128,128):"Gray",(128,0,0):"Maroon",(0,128,0):"Green",(0,0,128):"Navy",
    (128,128,0):"Olive",(0,128,128):"Teal",(128,0,128):"Purple",(255,128,0):"Orange",
    (128,128,255):"LightBlue",(192,192,192):"Silver",
}

def _rgb(color_int):
    return (color_int & 255, (color_int >> 8) & 255, (color_int >> 16) & 255)

def _sanitize(v):
    if not v:
        return "N/A"
    return re.sub(r'[\\/:?"<>|=]', '', v).strip() or "N/A"

def _mm(param):
    return "{:.2f}mm".format(round(param.AsDouble() * 304.8, 2))


# ============================================================
# DIMENSION RENAME HELPERS
# ============================================================
def _dim_name(dt, origin):
    def gp(bip):
        try: return dt.get_Parameter(bip)
        except: return None

    discipline = "STR" if "STR" in origin.upper() else "ARC"
    p = gp(BuiltInParameter.TEXT_SIZE)
    size  = _mm(p) if p else "N/A"
    p = gp(BuiltInParameter.TEXT_FONT)
    font  = p.AsString() if p else "N/A"
    p = gp(BuiltInParameter.DIM_TEXT_BACKGROUND)
    bg    = p.AsValueString() if p else "N/A"
    p = gp(BuiltInParameter.LINE_COLOR)
    color = _DIM_COLORS.get(_rgb(p.AsInteger()), "RGB") if p else "N/A"
    p = gp(BuiltInParameter.DIM_PREFIX)
    pref  = _sanitize(p.AsString()) if p else "N/A"
    p = gp(BuiltInParameter.DIM_STYLE_CENTERLINE_SYMBOL)
    ctr   = "Center" if (p and p.AsElementId() != ElementId.InvalidElementId) else "N/A"
    p = gp(BuiltInParameter.SPOT_ELEV_IND_ELEVATION)
    elev  = _sanitize(p.AsString()) if p else "N/A"
    p = gp(BuiltInParameter.SPOT_ELEV_IND_TOP)
    top   = _sanitize(p.AsString()) if p else "N/A"
    p = gp(BuiltInParameter.SPOT_ELEV_IND_BOTTOM)
    bot   = _sanitize(p.AsString()) if p else "N/A"

    parts = ["LB", discipline, size, font, bg]
    if color != "Black": parts.append(color)
    if ctr  != "N/A":   parts.append(ctr)
    if pref != "N/A":   parts.append(pref)
    if elev != "N/A":
        parts.append(elev)
    else:
        if top != "N/A": parts.append(top)
        if bot != "N/A": parts.append(bot)
    return "_".join(parts)


# ============================================================
# TEXTNOTE RENAME HELPERS
# ============================================================
def _txt_name(tt, origin):
    def gp(bip):
        try: return tt.get_Parameter(bip)
        except: return None

    discipline = "STR" if "STR" in origin.upper() else "ARC"
    p = gp(BuiltInParameter.TEXT_SIZE)
    size   = _mm(p) if p else "N/A"
    p = gp(BuiltInParameter.TEXT_FONT)
    font   = p.AsString().replace(" ", "") if p else "N/A"
    p = gp(BuiltInParameter.TEXT_BACKGROUND)
    bg     = ("Opaque" if p.AsInteger() == 0 else "Transparent") if p else "N/A"
    p = gp(BuiltInParameter.TEXT_WIDTH_SCALE)
    factor = str(round(p.AsDouble(), 2)) if p else "N/A"
    p = gp(BuiltInParameter.LINE_COLOR)
    color  = _TXT_COLORS.get(_rgb(p.AsInteger()), "RGB") if p else "N/A"
    p = gp(BuiltInParameter.TEXT_BOX_VISIBILITY)
    border = p and p.AsInteger() == 1
    p = gp(BuiltInParameter.TEXT_STYLE_BOLD)
    bold   = p and p.AsInteger() == 1
    p = gp(BuiltInParameter.TEXT_STYLE_UNDERLINE)
    uline  = p and p.AsInteger() == 1
    p = gp(BuiltInParameter.TEXT_STYLE_ITALIC)
    italic = p and p.AsInteger() == 1

    parts = ["LB", discipline, size, font, bg, factor]
    if color != "Black": parts.append(color)
    if border:  parts.append("Border")
    if bold:    parts.append("B")
    if uline:   parts.append("U")
    if italic:  parts.append("I")
    return "_".join(parts)


# ============================================================
# XAML
# ============================================================
XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Annotation Manager - T3Lab"
        Width="700" Height="640"
        WindowStartupLocation="CenterScreen"
        Background="White">
  <Window.Resources>
    <Style x:Key="ModeOn" TargetType="Button">
      <Setter Property="Background"   Value="#3498DB"/>
      <Setter Property="Foreground"   Value="White"/>
      <Setter Property="FontWeight"   Value="Bold"/>
      <Setter Property="FontSize"     Value="13"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"      Value="22,9"/>
      <Setter Property="Cursor"       Value="Hand"/>
    </Style>
    <Style x:Key="ModeOff" TargetType="Button">
      <Setter Property="Background"   Value="#ECF0F1"/>
      <Setter Property="Foreground"   Value="#7F8C8D"/>
      <Setter Property="FontWeight"   Value="Normal"/>
      <Setter Property="FontSize"     Value="13"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"      Value="22,9"/>
      <Setter Property="Cursor"       Value="Hand"/>
    </Style>
    <Style x:Key="BtnBlue" TargetType="Button">
      <Setter Property="Background"   Value="#3498DB"/>
      <Setter Property="Foreground"   Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"      Value="12,6"/>
      <Setter Property="FontSize"     Value="12"/>
      <Setter Property="Cursor"       Value="Hand"/>
    </Style>
    <Style x:Key="BtnRed" TargetType="Button">
      <Setter Property="Background"   Value="#E74C3C"/>
      <Setter Property="Foreground"   Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"      Value="12,6"/>
      <Setter Property="FontSize"     Value="12"/>
      <Setter Property="Cursor"       Value="Hand"/>
    </Style>
    <Style x:Key="BtnOrange" TargetType="Button">
      <Setter Property="Background"   Value="#E67E22"/>
      <Setter Property="Foreground"   Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding"      Value="12,6"/>
      <Setter Property="FontSize"     Value="12"/>
      <Setter Property="Cursor"       Value="Hand"/>
    </Style>
    <Style x:Key="BtnGray" TargetType="Button">
      <Setter Property="Background"   Value="#ECF0F1"/>
      <Setter Property="Foreground"   Value="#2C3E50"/>
      <Setter Property="BorderBrush"  Value="#BDC3C7"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"      Value="10,5"/>
      <Setter Property="FontSize"     Value="12"/>
      <Setter Property="Cursor"       Value="Hand"/>
    </Style>
  </Window.Resources>

  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <StackPanel Grid.Row="0" Margin="0,0,0,10">
      <TextBlock Text="Annotation Manager" FontSize="20" FontWeight="Bold" Foreground="#2C3E50"/>
      <TextBlock x:Name="lbl_subtitle"
                 Text="Dimension: Find · Remove · Rename Types"
                 FontSize="11" Foreground="#7F8C8D" Margin="0,2,0,0"/>
      <Separator Margin="0,8,0,0"/>
    </StackPanel>

    <!-- Mode toggle -->
    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,12">
      <Button x:Name="btn_dim"  Content="📐  Dimension"  Style="{StaticResource ModeOn}"
              Click="switch_dim"/>
      <Button x:Name="btn_txt"  Content="📝  Text Note"  Style="{StaticResource ModeOff}"
              Margin="5,0,0,0" Click="switch_txt"/>
    </StackPanel>

    <!-- ═══════════════ DIMENSION PANEL ═══════════════ -->
    <Grid x:Name="pnl_dim" Grid.Row="2">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Search row -->
      <Grid Grid.Row="0" Margin="0,0,0,6">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="Type name:" VerticalAlignment="Center"
                   FontWeight="Medium" Margin="0,0,8,0"/>
        <TextBox  x:Name="dim_kw"  Grid.Column="1" Height="30"
                  VerticalContentAlignment="Center" Padding="6,0"
                  BorderBrush="#BDC3C7" FontFamily="Consolas"/>
        <Button   Grid.Column="2" Content="Search" Height="30" Margin="6,0,0,0"
                  Style="{StaticResource BtnBlue}" Click="dim_search"/>
        <TextBlock x:Name="dim_count" Grid.Column="3"
                   VerticalAlignment="Center" Margin="10,0,0,0"
                   FontSize="11" Foreground="#7F8C8D"/>
      </Grid>

      <!-- Result list -->
      <Border Grid.Row="1" BorderBrush="#BDC3C7" BorderThickness="1">
        <ListBox x:Name="lb_dim" SelectionMode="Extended"
                 FontFamily="Consolas" FontSize="11"
                 ScrollViewer.HorizontalScrollBarVisibility="Auto"/>
      </Border>

      <!-- Action buttons -->
      <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,8">
        <Button Content="Jump to View"    Style="{StaticResource BtnBlue}"
                Margin="0,0,6,0" Click="dim_jump"/>
        <Button Content="Delete Selected" Style="{StaticResource BtnRed}"
                Margin="0,0,6,0" Click="dim_delete"/>
        <Button Content="Select All"      Style="{StaticResource BtnGray}"
                Margin="0,0,6,0" Click="dim_select_all"/>
        <Button Content="Clear"           Style="{StaticResource BtnGray}"
                Click="dim_clear_sel"/>
      </StackPanel>

      <!-- Rename section -->
      <Border Grid.Row="3" Background="#F8F9FA" BorderBrush="#D5DBDB" BorderThickness="1" Padding="10,8">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0">
            <TextBlock Text="Auto-Rename All Dimension Types" FontWeight="Bold" Foreground="#2C3E50"/>
            <TextBlock FontSize="10" Foreground="#7F8C8D" TextWrapping="Wrap"
                       Text="Generates: LB_[discipline]_[size]_[font]_[bg]_[color]_[prefix]_[indicators]"/>
          </StackPanel>
          <Button Grid.Column="1" Content="Rename All" Width="100"
                  Style="{StaticResource BtnOrange}" Click="dim_rename_all" VerticalAlignment="Center"/>
        </Grid>
      </Border>
    </Grid>

    <!-- ═══════════════ TEXTNOTE PANEL ═══════════════ -->
    <Grid x:Name="pnl_txt" Grid.Row="2" Visibility="Collapsed">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Sub-mode -->
      <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
        <RadioButton x:Name="rb_notes" Content="Find Notes  (search by text content)"
                     IsChecked="True" Margin="0,0,20,0" FontSize="12"
                     Checked="txt_submode"/>
        <RadioButton x:Name="rb_types" Content="Find Types  (search by type name)"
                     FontSize="12" Checked="txt_submode"/>
      </StackPanel>

      <!-- Search row -->
      <Grid Grid.Row="1" Margin="0,0,0,6">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="txt_lbl" Grid.Column="0" Text="Content:"
                   VerticalAlignment="Center" FontWeight="Medium" Margin="0,0,8,0"/>
        <TextBox  x:Name="txt_kw" Grid.Column="1" Height="30"
                  VerticalContentAlignment="Center" Padding="6,0"
                  BorderBrush="#BDC3C7" FontFamily="Consolas"/>
        <Button   Grid.Column="2" Content="Search" Height="30" Margin="6,0,0,0"
                  Style="{StaticResource BtnBlue}" Click="txt_search"/>
        <TextBlock x:Name="txt_count" Grid.Column="3"
                   VerticalAlignment="Center" Margin="10,0,0,0"
                   FontSize="11" Foreground="#7F8C8D"/>
      </Grid>

      <!-- Result list -->
      <Border Grid.Row="2" BorderBrush="#BDC3C7" BorderThickness="1">
        <ListBox x:Name="lb_txt" SelectionMode="Extended"
                 FontFamily="Consolas" FontSize="11"
                 ScrollViewer.HorizontalScrollBarVisibility="Auto"/>
      </Border>

      <!-- Action buttons -->
      <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,8,0,8">
        <Button x:Name="btn_txt_jump" Content="Jump to View" Style="{StaticResource BtnBlue}"
                Margin="0,0,6,0" Click="txt_jump"/>
        <Button Content="Delete Selected" Style="{StaticResource BtnRed}"
                Margin="0,0,6,0" Click="txt_delete"/>
        <Button Content="Select All"      Style="{StaticResource BtnGray}"
                Margin="0,0,6,0" Click="txt_select_all"/>
        <Button Content="Clear"           Style="{StaticResource BtnGray}"
                Click="txt_clear_sel"/>
      </StackPanel>

      <!-- Rename section -->
      <Border Grid.Row="4" Background="#F8F9FA" BorderBrush="#D5DBDB" BorderThickness="1" Padding="10,8">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0">
            <TextBlock Text="Auto-Rename All Text Note Types" FontWeight="Bold" Foreground="#2C3E50"/>
            <TextBlock FontSize="10" Foreground="#7F8C8D" TextWrapping="Wrap"
                       Text="Generates: LB_[discipline]_[size]_[font]_[bg]_[factor]_[color]_[border]_[B/U/I]"/>
          </StackPanel>
          <Button Grid.Column="1" Content="Rename All" Width="100"
                  Style="{StaticResource BtnOrange}" Click="txt_rename_all" VerticalAlignment="Center"/>
        </Grid>
      </Border>
    </Grid>

    <!-- Status bar -->
    <TextBlock x:Name="status" Grid.Row="3" Margin="0,8,0,0"
               FontSize="11" Foreground="#7F8C8D" Text="Ready."/>
  </Grid>
</Window>
"""


# ============================================================
# WINDOW CLASS
# ============================================================
class AnnotationManagerWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML, literal_string=True)
        self._mode     = "dim"    # "dim" | "txt"
        self._submode  = "notes"  # "notes" | "types"
        # Parallel item stores: index matches ListBox row
        self._dim_items  = []   # list of DimResultItem
        self._txt_items  = []   # list of TextNoteItem or TextNoteTypeItem

    # ── helpers ────────────────────────────────────────────────
    def _status(self, msg):
        self.status.Text = msg

    def _lb_selected_indices(self, lb):
        """Return list of indices of selected rows in a ListBox."""
        selected = []
        for i in range(lb.Items.Count):
            lbi = lb.ItemContainerGenerator.ContainerFromIndex(i)
            if lbi and lbi.IsSelected:
                selected.append(i)
        return selected

    # ── Mode switching ──────────────────────────────────────────
    def switch_dim(self, sender, args):
        self._mode = "dim"
        self.pnl_dim.Visibility = Visibility.Visible
        self.pnl_txt.Visibility = Visibility.Collapsed
        self.btn_dim.Style = self.FindResource("ModeOn")
        self.btn_txt.Style = self.FindResource("ModeOff")
        self.lbl_subtitle.Text = "Dimension: Find · Remove · Rename Types"
        self._status("Dimension mode.")

    def switch_txt(self, sender, args):
        self._mode = "txt"
        self.pnl_dim.Visibility = Visibility.Collapsed
        self.pnl_txt.Visibility = Visibility.Visible
        self.btn_dim.Style = self.FindResource("ModeOff")
        self.btn_txt.Style = self.FindResource("ModeOn")
        self.lbl_subtitle.Text = "Text Note: Find Notes · Find Types · Rename Types"
        self._status("Text Note mode.")

    # ── TextNote sub-mode ───────────────────────────────────────
    def txt_submode(self, sender, args):
        if self.rb_notes.IsChecked:
            self._submode = "notes"
            self.txt_lbl.Text       = "Content:"
            self.btn_txt_jump.IsEnabled = True
        else:
            self._submode = "types"
            self.txt_lbl.Text       = "Type name:"
            self.btn_txt_jump.IsEnabled = False
        self.lb_txt.Items.Clear()
        self._txt_items = []
        self.txt_count.Text = ""
        self._status("Sub-mode: {}.".format("Find Notes" if self._submode == "notes" else "Find Types"))

    # ── DIMENSION operations ────────────────────────────────────

    def dim_search(self, sender, args):
        kw = self.dim_kw.Text.strip().lower()
        if not kw:
            self._status("Enter a keyword.")
            return
        dims = FilteredElementCollector(doc).OfClass(Dimension)\
               .WhereElementIsNotElementType().ToElements()
        self._dim_items = []
        self.lb_dim.Items.Clear()
        for d in dims:
            if kw in (d.Name or "").lower():
                view = doc.GetElement(d.OwnerViewId)
                if view:
                    label = "{}  |  {}".format(d.Name or "<unnamed>", view.Name)
                    self._dim_items.append(d)
                    self.lb_dim.Items.Add(label)
        n = len(self._dim_items)
        self.dim_count.Text = "{} found".format(n)
        self._status("Found {} dimension(s) matching '{}'.".format(n, kw))

    def dim_jump(self, sender, args):
        idxs = self._lb_selected_indices(self.lb_dim)
        if not idxs:
            self._status("Select a dimension first.")
            return
        d = self._dim_items[idxs[0]]
        view = doc.GetElement(d.OwnerViewId)
        if view:
            uidoc.ActiveView = view
            uidoc.ShowElements(d.Id)
            self._status("Jumped to view '{}' — dimension '{}'.".format(view.Name, d.Name or ""))

    def dim_delete(self, sender, args):
        idxs = self._lb_selected_indices(self.lb_dim)
        if not idxs:
            self._status("Nothing selected.")
            return
        t = Transaction(doc, "Delete Selected Dimensions")
        t.Start()
        for i in sorted(idxs, reverse=True):
            try:
                doc.Delete(self._dim_items[i].Id)
            except Exception as e:
                self._status("Error deleting: {}".format(e))
        t.Commit()
        # remove from list (reverse order to keep indices valid)
        for i in sorted(idxs, reverse=True):
            self.lb_dim.Items.RemoveAt(i)
            del self._dim_items[i]
        n = len(idxs)
        self.dim_count.Text = "{} found".format(len(self._dim_items))
        self._status("Deleted {} dimension(s).".format(n))

    def dim_select_all(self, sender, args):
        self.lb_dim.SelectAll()

    def dim_clear_sel(self, sender, args):
        self.lb_dim.UnselectAll()

    def dim_rename_all(self, sender, args):
        from pyrevit import forms as pf
        if not pf.alert("Auto-rename ALL DimensionTypes in this document?\nThis cannot be undone.",
                        title="Confirm Rename", yes=True, no=True):
            return
        t = Transaction(doc, "Rename Dimension Types")
        t.Start()
        count = 0
        try:
            for dt in FilteredElementCollector(doc).OfClass(DimensionType)\
                      .WhereElementIsElementType().ToElements():
                try:
                    origin = dt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    dt.Name = _dim_name(dt, origin)
                    count += 1
                except Exception as e:
                    pass
        finally:
            t.Commit()
        self._status("Renamed {} DimensionType(s).".format(count))

    # ── TEXTNOTE operations ─────────────────────────────────────

    def txt_search(self, sender, args):
        kw = self.txt_kw.Text.strip().lower()
        if not kw:
            self._status("Enter a keyword.")
            return
        self._txt_items = []
        self.lb_txt.Items.Clear()

        if self._submode == "notes":
            notes = FilteredElementCollector(doc).OfClass(TextNote)\
                    .WhereElementIsNotElementType().ToElements()
            for tn in notes:
                if kw in (tn.Text or "").lower():
                    view = doc.GetElement(tn.ViewId)
                    if view:
                        preview = (tn.Text or "")[:60].replace("\n", " ").replace("\r", "")
                        label = "{}  |  {}".format(preview, view.Name)
                        self._txt_items.append(tn)
                        self.lb_txt.Items.Add(label)
        else:  # types
            types = FilteredElementCollector(doc).OfClass(TextNoteType)\
                    .WhereElementIsElementType().ToElements()
            for tt in types:
                name = tt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() or ""
                if kw in name.lower():
                    label = "{}  [ID: {}]".format(name, tt.Id)
                    self._txt_items.append(tt)
                    self.lb_txt.Items.Add(label)

        n = len(self._txt_items)
        self.txt_count.Text = "{} found".format(n)
        self._status("Found {} {}.".format(n, "note(s)" if self._submode == "notes" else "type(s)"))

    def txt_jump(self, sender, args):
        if self._submode != "notes":
            self._status("Jump to View is only available in Find Notes mode.")
            return
        idxs = self._lb_selected_indices(self.lb_txt)
        if not idxs:
            self._status("Select a text note first.")
            return
        tn = self._txt_items[idxs[0]]
        view = doc.GetElement(tn.ViewId)
        if view:
            uidoc.ActiveView = view
            uidoc.ShowElements(tn.Id)
            preview = (tn.Text or "")[:40].replace("\n", " ")
            self._status("Jumped to view '{}' — note: '{}'.".format(view.Name, preview))

    def txt_delete(self, sender, args):
        idxs = self._lb_selected_indices(self.lb_txt)
        if not idxs:
            self._status("Nothing selected.")
            return
        label = "note instance(s)" if self._submode == "notes" else "TextNoteType(s)"
        t = Transaction(doc, "Delete Selected Text {}".format(label))
        t.Start()
        errors = 0
        for i in sorted(idxs, reverse=True):
            try:
                doc.Delete(self._txt_items[i].Id)
            except Exception:
                errors += 1
        t.Commit()
        for i in sorted(idxs, reverse=True):
            self.lb_txt.Items.RemoveAt(i)
            del self._txt_items[i]
        n = len(idxs)
        self.txt_count.Text = "{} found".format(len(self._txt_items))
        msg = "Deleted {} {}.".format(n - errors, label)
        if errors:
            msg += "  ({} could not be deleted — may be in use.)".format(errors)
        self._status(msg)

    def txt_select_all(self, sender, args):
        self.lb_txt.SelectAll()

    def txt_clear_sel(self, sender, args):
        self.lb_txt.UnselectAll()

    def txt_rename_all(self, sender, args):
        from pyrevit import forms as pf
        if not pf.alert("Auto-rename ALL TextNoteTypes in this document?\nThis cannot be undone.",
                        title="Confirm Rename", yes=True, no=True):
            return
        t = Transaction(doc, "Rename TextNote Types")
        t.Start()
        count = 0
        try:
            for tt in FilteredElementCollector(doc).OfClass(TextNoteType)\
                      .WhereElementIsElementType().ToElements():
                try:
                    origin = tt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    tt.Name = _txt_name(tt, origin)
                    count += 1
                except Exception:
                    pass
        finally:
            t.Commit()
        self._status("Renamed {} TextNoteType(s).".format(count))


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    win = AnnotationManagerWindow()
    win.show_dialog()
