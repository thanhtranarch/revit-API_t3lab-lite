# -*- coding: utf-8 -*-
"""In-Place Model Manager v1.6 - Full Features
Author: Dang Quoc Truong (DQT)
"""
__title__ = "In-Place\nManager"
__author__ = "DQT"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from pyrevit import revit, forms, script
from pyrevit.forms import WPFWindow
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from System.Collections.Generic import List
import System
import re, datetime, codecs, os

# ============================================================================
# CONFIGURATION
# ============================================================================
class Config:
    patterns = [r'^[A-Z]{2,4}_.*', r'^IP_.*', r'^.*_InPlace$']
    warn_faces = 100
    err_faces = 500

# ============================================================================
# DATA MODEL
# ============================================================================
class ItemData(object):
    def __init__(self):
        self.element_id = 0
        self.name = ""
        self.family_name = ""
        self.category_name = ""
        self.face_count = 0
        self.complexity_status = "OK"
        self.issue_text = "OK"
        self.workset = "-"
        self.level = "-"
        self.has_issues = False
        self.element = None

def get_items(doc):
    items = []
    for el in FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType():
        try:
            if el.Symbol and el.Symbol.Family and el.Symbol.Family.IsInPlace:
                d = ItemData()
                d.element = el
                d.element_id = el.Id.IntegerValue
                d.name = el.Name if el.Name else "<No Name>"
                d.family_name = el.Symbol.Family.Name
                d.category_name = el.Category.Name if el.Category else "N/A"
                
                # Get Workset info
                d.workset = "-"
                try:
                    if doc.IsWorkshared:
                        ws_id = el.WorksetId
                        if ws_id and ws_id.IntegerValue > 0:
                            wst = doc.GetWorksetTable()
                            ws = wst.GetWorkset(ws_id)
                            if ws:
                                d.workset = ws.Name
                except: 
                    pass
                
                # Level
                d.level = "-"
                try:
                    if el.LevelId and el.LevelId.IntegerValue != -1:
                        lv = doc.GetElement(el.LevelId)
                        d.level = lv.Name if lv else "-"
                except: pass
                
                # Face count
                d.face_count = 0
                try:
                    opt = Options()
                    opt.DetailLevel = ViewDetailLevel.Coarse
                    geom = el.get_Geometry(opt)
                    if geom:
                        for g in geom:
                            if hasattr(g, 'Faces') and g.Faces:
                                d.face_count += g.Faces.Size
                except: pass
                
                # Complexity status
                if d.face_count >= Config.err_faces:
                    d.complexity_status = "Error"
                elif d.face_count >= Config.warn_faces:
                    d.complexity_status = "Warning"
                
                # Issues
                issues = []
                if d.name == "<No Name>":
                    issues.append("No Name")
                elif not any(re.match(p, d.name) for p in Config.patterns):
                    issues.append("Bad Name")
                if d.complexity_status != "OK":
                    issues.append("Complex")
                
                d.issue_text = ", ".join(issues) if issues else "OK"
                d.has_issues = len(issues) > 0
                items.append(d)
        except:
            pass
    return items

# ============================================================================
# HTML REPORT
# ============================================================================
def generate_html_report(items, doc_title):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    total = len(items)
    issues = len([i for i in items if i.has_issues])
    
    cats = {}
    for i in items:
        c = i.category_name
        if c not in cats:
            cats[c] = {"count": 0, "issues": 0, "faces": 0}
        cats[c]["count"] += 1
        cats[c]["faces"] += i.face_count
        if i.has_issues:
            cats[c]["issues"] += 1
    
    cat_rows = ""
    for c in sorted(cats.keys()):
        s = cats[c]
        cat_rows += "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(c, s["count"], s["faces"], s["issues"])
    
    item_rows = ""
    for i in items:
        color = "#4CAF50" if i.complexity_status == "OK" else ("#FF9800" if i.complexity_status == "Warning" else "#f44336")
        iss = '<span style="color:#4CAF50">OK</span>' if not i.has_issues else '<span style="color:#f44336">{}</span>'.format(i.issue_text)
        item_rows += "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td style='color:{}'>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
            i.element_id, i.family_name, i.name, i.category_name, i.face_count, color, i.complexity_status, iss, i.level, i.workset)
    
    return """<!DOCTYPE html><html><head><meta charset="UTF-8"><title>In-Place Report - {}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f5f5f5;padding:20px;color:#333}}
.container{{max-width:1400px;margin:0 auto}}
.header{{background:linear-gradient(135deg,#F0CC88,#E5B85C);padding:25px;border-radius:10px;margin-bottom:20px}}
.header h1{{font-size:24px;margin-bottom:5px}}
.header .meta{{font-size:12px;opacity:0.8}}
.header .author{{font-size:11px;margin-top:8px;padding-top:8px;border-top:1px solid rgba(255,255,255,0.3)}}
.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin-bottom:20px}}
.card{{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}}
.card .label{{font-size:11px;color:#666;text-transform:uppercase}}
.card .value{{font-size:32px;font-weight:bold;margin-top:5px}}
.card.ok .value{{color:#4CAF50}}
.card.err .value{{color:#f44336}}
.section{{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1);margin-bottom:20px}}
.section h2{{font-size:16px;margin-bottom:15px;padding-bottom:10px;border-bottom:2px solid #F0CC88}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th,td{{padding:10px;text-align:left;border-bottom:1px solid #eee}}
th{{background:#F0CC88;font-weight:600}}
tr:hover{{background:#fafafa}}
.footer{{text-align:center;padding:20px;color:#888;font-size:11px;background:#fff;border-radius:8px;margin-top:20px}}
.footer .copyright{{font-weight:600;color:#666;margin-bottom:5px}}
</style></head>
<body><div class="container">
<div class="header">
<h1>In-Place Model Report</h1>
<div class="meta">{} | Generated: {}</div>
<div class="author">Created by: <strong>Dang Quoc Truong (DQT)</strong> | In-Place Manager v1.6</div>
</div>
<div class="cards">
<div class="card"><div class="label">Total In-Place</div><div class="value">{}</div></div>
<div class="card ok"><div class="label">OK</div><div class="value">{}</div></div>
<div class="card err"><div class="label">With Issues</div><div class="value">{}</div></div>
<div class="card"><div class="label">Categories</div><div class="value">{}</div></div>
</div>
<div class="section">
<h2>By Category</h2>
<table><tr><th>Category</th><th>Count</th><th>Total Faces</th><th>Issues</th></tr>{}</table>
</div>
<div class="section">
<h2>All Items ({})</h2>
<table><tr><th>ID</th><th>Family Name</th><th>Type</th><th>Category</th><th>Faces</th><th>Status</th><th>Issues</th><th>Level</th><th>Workset</th></tr>{}</table>
</div>
<div class="footer">
<div class="copyright">&copy; 2024 Dang Quoc Truong (DQT) - All Rights Reserved</div>
<div>In-Place Manager v1.6 | Generated: {}</div>
</div>
</div></body></html>""".format(doc_title, doc_title, ts, total, total-issues, issues, len(cats), cat_rows, total, item_rows, ts)

# ============================================================================
# XAML
# ============================================================================
MAIN_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="In-Place Manager v1.6 - DQT" 
        Height="750" Width="1150" 
        WindowStartupLocation="CenterScreen"
        Background="#FEF8E7">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#F0CC88" CornerRadius="5" Padding="12,8" Margin="0,0,0,10">
            <Grid>
                <StackPanel>
                    <TextBlock Text="In-Place Model Manager v1.6" FontSize="17" FontWeight="Bold"/>
                    <TextBlock Text="by Dang Quoc Truong (DQT)" FontSize="10" Foreground="#5D4E37" Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="btnSettings" Content="⚙ Settings" Padding="10,4" Background="White" Margin="0,0,5,0"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Summary Cards -->
        <Grid Grid.Row="1" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="White" BorderBrush="#D4B87A" BorderThickness="1" CornerRadius="4" Padding="10,6" Margin="0,0,4,0">
                <StackPanel><TextBlock Text="TOTAL" FontSize="9" Foreground="#666"/><TextBlock x:Name="txtTotal" Text="0" FontSize="22" FontWeight="Bold"/></StackPanel>
            </Border>
            <Border Grid.Column="1" Background="White" BorderBrush="#D4B87A" BorderThickness="1" CornerRadius="4" Padding="10,6" Margin="4,0">
                <StackPanel><TextBlock Text="SELECTED" FontSize="9" Foreground="#666"/><TextBlock x:Name="txtSelected" Text="0" FontSize="22" FontWeight="Bold" Foreground="#E5B85C"/></StackPanel>
            </Border>
            <Border Grid.Column="2" Background="White" BorderBrush="#D4B87A" BorderThickness="1" CornerRadius="4" Padding="10,6" Margin="4,0">
                <StackPanel><TextBlock Text="CATEGORIES" FontSize="9" Foreground="#666"/><TextBlock x:Name="txtCategories" Text="0" FontSize="22" FontWeight="Bold" Foreground="#4CAF50"/></StackPanel>
            </Border>
            <Border Grid.Column="3" Background="White" BorderBrush="#FF6B6B" BorderThickness="1" CornerRadius="4" Padding="10,6" Margin="4,0,0,0">
                <StackPanel><TextBlock Text="ISSUES" FontSize="9" Foreground="#FF6B6B"/><TextBlock x:Name="txtIssues" Text="0" FontSize="22" FontWeight="Bold" Foreground="#FF6B6B"/></StackPanel>
            </Border>
        </Grid>
        
        <!-- Content -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="180"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Left Panel -->
            <Border Grid.Column="0" Background="White" BorderBrush="#D4B87A" BorderThickness="1" CornerRadius="4" Padding="8" Margin="0,0,8,0">
                <StackPanel>
                    <TextBlock Text="SEARCH" FontSize="9" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <TextBox x:Name="txtSearch" Padding="6,4" Margin="0,0,0,10"/>
                    <TextBlock Text="FILTER" FontSize="9" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <ComboBox x:Name="cmbFilter" Padding="6,4" Margin="0,0,0,10" SelectedIndex="0">
                        <ComboBoxItem Content="All"/>
                        <ComboBoxItem Content="With Issues"/>
                        <ComboBoxItem Content="No Name"/>
                        <ComboBoxItem Content="Complex"/>
                        <ComboBoxItem Content="Bad Name"/>
                        <ComboBoxItem Content="OK Only"/>
                    </ComboBox>
                    <TextBlock Text="CATEGORY" FontSize="9" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <ListBox x:Name="lstCategories" Height="280"/>
                </StackPanel>
            </Border>
            
            <!-- DataGrid -->
            <DataGrid Grid.Column="1" x:Name="dataGrid" 
                      AutoGenerateColumns="False" IsReadOnly="True"
                      SelectionMode="Extended" SelectionUnit="FullRow"
                      CanUserSortColumns="True"
                      Background="White" BorderBrush="#D4B87A"
                      GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#EEE"
                      RowBackground="White" AlternatingRowBackground="#FFFDF5">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="ID" Binding="{Binding element_id}" Width="70" SortMemberPath="element_id"/>
                    <DataGridTextColumn Header="Family Name" Binding="{Binding family_name}" Width="*" SortMemberPath="family_name"/>
                    <DataGridTextColumn Header="Type" Binding="{Binding name}" Width="100" SortMemberPath="name"/>
                    <DataGridTextColumn Header="Category" Binding="{Binding category_name}" Width="90" SortMemberPath="category_name"/>
                    <DataGridTextColumn Header="Faces" Binding="{Binding face_count}" Width="50" SortMemberPath="face_count"/>
                    <DataGridTextColumn Header="Status" Binding="{Binding complexity_status}" Width="60" SortMemberPath="complexity_status"/>
                    <DataGridTextColumn Header="Issues" Binding="{Binding issue_text}" Width="80" SortMemberPath="issue_text"/>
                    <DataGridTextColumn Header="Level" Binding="{Binding level}" Width="80" SortMemberPath="level"/>
                    <DataGridTextColumn Header="Workset" Binding="{Binding workset}" Width="80" SortMemberPath="workset"/>
                </DataGrid.Columns>
            </DataGrid>
        </Grid>
        
        <!-- Action Buttons -->
        <Border Grid.Row="3" Background="White" BorderBrush="#D4B87A" BorderThickness="1" CornerRadius="4" Padding="8" Margin="0,10,0,0">
            <Grid>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                    <Button x:Name="btnSelectAll" Content="Select All" Padding="10,5" Margin="2" Background="White"/>
                    <Button x:Name="btnClear" Content="Clear" Padding="10,5" Margin="2" Background="White"/>
                    <Button x:Name="btnSelectIssues" Content="Select Issues" Padding="10,5" Margin="2" Background="White"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                    <Button x:Name="btnZoom" Content="Zoom To" Padding="10,5" Margin="2" Background="#F0CC88"/>
                    <Button x:Name="btnSelect" Content="Select" Padding="10,5" Margin="2" Background="#F0CC88"/>
                    <Button x:Name="btnIsolate" Content="Isolate" Padding="10,5" Margin="2" Background="#F0CC88"/>
                    <Button x:Name="btnRename" Content="Rename" Padding="10,5" Margin="2" Background="#F0CC88"/>
                    <Button x:Name="btnExportFamily" Content="To Family" Padding="10,5" Margin="2" Background="#4CAF50" Foreground="White" ToolTip="Convert to Loadable Family"/>
                    <Button x:Name="btnExportCSV" Content="CSV" Padding="10,5" Margin="2" Background="#F0CC88"/>
                    <Button x:Name="btnExportHTML" Content="Report" Padding="10,5" Margin="2" Background="#F0CC88"/>
                    <Button x:Name="btnDelete" Content="Delete" Padding="10,5" Margin="2" Background="#FF6B6B" Foreground="White"/>
                </StackPanel>
                <Button x:Name="btnRefresh" Content="Refresh" Padding="10,5" Margin="2" Background="White" HorizontalAlignment="Right"/>
            </Grid>
        </Border>
        
        <!-- Footer with Copyright -->
        <Grid Grid.Row="4" Margin="0,8,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0">
                <TextBlock Text="Ctrl+Click: multi-select | Double-click: zoom to element" FontSize="10" Foreground="#888"/>
                <Button x:Name="btnClose" Content="Close" Padding="15,5" Background="White" HorizontalAlignment="Right"/>
            </Grid>
            <Border Grid.Row="1" Background="#F0CC88" CornerRadius="3" Padding="8,5" Margin="0,8,0,0">
                <TextBlock Text="© 2024 Dang Quoc Truong (DQT) - All Rights Reserved" 
                           FontSize="10" FontWeight="SemiBold" HorizontalAlignment="Center" Foreground="#5D4E37"/>
            </Border>
        </Grid>
    </Grid>
</Window>
"""

SETTINGS_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Settings" Height="350" Width="400" 
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        Background="#FEF8E7">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="Settings" FontSize="18" FontWeight="Bold" Margin="0,0,0,15"/>
        
        <StackPanel Grid.Row="1">
            <Border Background="White" BorderBrush="#D4B87A" BorderThickness="1" CornerRadius="4" Padding="12" Margin="0,0,0,10">
                <StackPanel>
                    <TextBlock Text="COMPLEXITY THRESHOLDS" FontSize="10" FontWeight="SemiBold" Foreground="#666" Margin="0,0,0,10"/>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                        <TextBlock Text="Warning (faces):" Width="100" VerticalAlignment="Center"/>
                        <TextBox x:Name="txtWarn" Width="80" Padding="6,4"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Error (faces):" Width="100" VerticalAlignment="Center"/>
                        <TextBox x:Name="txtErr" Width="80" Padding="6,4"/>
                    </StackPanel>
                </StackPanel>
            </Border>
            
            <Border Background="White" BorderBrush="#D4B87A" BorderThickness="1" CornerRadius="4" Padding="12">
                <StackPanel>
                    <TextBlock Text="NAMING PATTERNS (regex, one per line)" FontSize="10" FontWeight="SemiBold" Foreground="#666" Margin="0,0,0,10"/>
                    <TextBox x:Name="txtPatterns" Height="80" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Padding="6,4"/>
                </StackPanel>
            </Border>
        </StackPanel>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnReset" Content="Reset" Width="70" Padding="8,6" Margin="0,0,8,0" Background="White"/>
            <Button x:Name="btnCancel" Content="Cancel" Width="70" Padding="8,6" Margin="0,0,8,0" Background="White"/>
            <Button x:Name="btnSave" Content="Save" Width="70" Padding="8,6" Background="#F0CC88" FontWeight="SemiBold"/>
        </StackPanel>
    </Grid>
</Window>
"""

RENAME_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Batch Rename" Height="400" Width="450" 
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        Background="#FEF8E7">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="Batch Rename" FontSize="16" FontWeight="Bold" Margin="0,0,0,5"/>
        <TextBlock Grid.Row="1" x:Name="txtCount" Text="Selected: 0 items" Foreground="#666" Margin="0,0,0,10"/>
        
        <Border Grid.Row="2" Background="White" BorderBrush="#D4B87A" BorderThickness="1" CornerRadius="4" Padding="10" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="RENAME ALL TO:" FontSize="10" FontWeight="SemiBold" Foreground="#666" Margin="0,0,0,5"/>
                <TextBox x:Name="txtNewName" Padding="6,4" ToolTip="Leave empty to use options below"/>
            </StackPanel>
        </Border>
        
        <TextBlock Grid.Row="3" Text="─── OR USE OPTIONS BELOW ───" HorizontalAlignment="Center" Foreground="#999" Margin="0,0,0,10"/>
        
        <Grid Grid.Row="4" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="60"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="60"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Prefix:" VerticalAlignment="Center"/>
            <TextBox Grid.Column="1" x:Name="txtPrefix" Padding="6,4" Margin="0,0,10,0"/>
            <TextBlock Grid.Column="2" Text="Suffix:" VerticalAlignment="Center"/>
            <TextBox Grid.Column="3" x:Name="txtSuffix" Padding="6,4"/>
        </Grid>
        
        <Grid Grid.Row="5" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="60"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="30"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Find:" VerticalAlignment="Center"/>
            <TextBox Grid.Column="1" x:Name="txtFind" Padding="6,4"/>
            <TextBlock Grid.Column="2" Text="→" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="14"/>
            <TextBox Grid.Column="3" x:Name="txtReplace" Padding="6,4" ToolTip="Replace with (can be empty)"/>
        </Grid>
        
        <Border Grid.Row="6" Background="White" BorderBrush="#D4B87A" BorderThickness="1" Padding="10" Margin="0,5,0,0">
            <StackPanel>
                <TextBlock Text="Preview (first item):" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <TextBlock x:Name="txtOldName" Text="" Foreground="#666" TextWrapping="Wrap"/>
                <TextBlock Text="↓" HorizontalAlignment="Center" Foreground="#999" Margin="0,3"/>
                <TextBlock x:Name="txtNewPreview" Text="" Foreground="#4CAF50" FontWeight="SemiBold" TextWrapping="Wrap"/>
            </StackPanel>
        </Border>
        
        <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnCancel" Content="Cancel" Width="70" Padding="8,6" Margin="0,0,8,0" Background="White"/>
            <Button x:Name="btnApply" Content="Apply" Width="70" Padding="8,6" Background="#F0CC88" FontWeight="SemiBold"/>
        </StackPanel>
    </Grid>
</Window>
"""

# ============================================================================
# SETTINGS DIALOG
# ============================================================================
class SettingsDialog(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, SETTINGS_XAML, literal_string=True)
        self.result = False
        
        self.txtWarn.Text = str(Config.warn_faces)
        self.txtErr.Text = str(Config.err_faces)
        self.txtPatterns.Text = "\n".join(Config.patterns)
        
        self.btnReset.Click += self.on_reset
        self.btnCancel.Click += self.on_cancel
        self.btnSave.Click += self.on_save
    
    def on_reset(self, s, e):
        self.txtWarn.Text = "100"
        self.txtErr.Text = "500"
        self.txtPatterns.Text = "^[A-Z]{2,4}_.*\n^IP_.*\n^.*_InPlace$"
    
    def on_cancel(self, s, e):
        self.Close()
    
    def on_save(self, s, e):
        try:
            Config.warn_faces = int(self.txtWarn.Text)
            Config.err_faces = int(self.txtErr.Text)
            Config.patterns = [p.strip() for p in self.txtPatterns.Text.split("\n") if p.strip()]
            self.result = True
            self.Close()
        except Exception as ex:
            forms.alert("Invalid input: {}".format(str(ex)), title="Error")

# ============================================================================
# RENAME DIALOG
# ============================================================================
class RenameDialog(WPFWindow):
    def __init__(self, items):
        WPFWindow.__init__(self, RENAME_XAML, literal_string=True)
        self.items = items
        self.result = False
        self.first_name = items[0].family_name if items else ""
        
        self.txtCount.Text = "Selected: {} item(s)".format(len(items))
        self.txtOldName.Text = self.first_name
        self.txtNewPreview.Text = self.first_name
        
        self.txtNewName.TextChanged += self.update_preview
        self.txtPrefix.TextChanged += self.update_preview
        self.txtSuffix.TextChanged += self.update_preview
        self.txtFind.TextChanged += self.update_preview
        self.txtReplace.TextChanged += self.update_preview
        
        self.btnCancel.Click += self.on_cancel
        self.btnApply.Click += self.on_apply
    
    def update_preview(self, sender, args):
        new = self.get_new_name(self.first_name)
        self.txtOldName.Text = self.first_name
        if new != self.first_name:
            self.txtNewPreview.Text = new
        else:
            self.txtNewPreview.Text = "(No changes)"
    
    def on_cancel(self, s, e):
        self.Close()
    
    def on_apply(self, s, e):
        has_new_name = self.txtNewName.Text.strip() if self.txtNewName.Text else ""
        has_prefix = self.txtPrefix.Text if self.txtPrefix.Text else ""
        has_suffix = self.txtSuffix.Text if self.txtSuffix.Text else ""
        has_find = self.txtFind.Text if self.txtFind.Text else ""
        
        if not has_new_name and not has_prefix and not has_suffix and not has_find:
            forms.alert("Enter at least one option", title="Info")
            return
        
        self.result = True
        self.Close()
    
    def get_new_name(self, name):
        new_name = self.txtNewName.Text.strip() if self.txtNewName.Text else ""
        if new_name:
            return new_name
        
        result = name
        find_text = self.txtFind.Text if self.txtFind.Text else ""
        replace_text = self.txtReplace.Text if self.txtReplace.Text else ""
        if find_text:
            result = result.replace(find_text, replace_text)
        
        prefix = self.txtPrefix.Text if self.txtPrefix.Text else ""
        if prefix:
            result = prefix + result
        
        suffix = self.txtSuffix.Text if self.txtSuffix.Text else ""
        if suffix:
            result = result + suffix
        
        return result

# ============================================================================
# MAIN WINDOW
# ============================================================================
class InPlaceManagerWindow(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, MAIN_XAML, literal_string=True)
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.items = []
        self.filtered = []
        self.cats = {}
        
        # Events
        self.txtSearch.TextChanged += self.on_filter
        self.cmbFilter.SelectionChanged += self.on_filter
        self.lstCategories.SelectionChanged += self.on_filter
        self.dataGrid.SelectionChanged += self.on_selection
        self.dataGrid.MouseDoubleClick += self.on_double_click
        
        self.btnSettings.Click += self.show_settings
        self.btnSelectAll.Click += self.select_all
        self.btnClear.Click += self.select_none
        self.btnSelectIssues.Click += self.select_issues
        self.btnZoom.Click += self.zoom_to
        self.btnSelect.Click += self.select_in_model
        self.btnIsolate.Click += self.isolate
        self.btnRename.Click += self.rename
        self.btnExportFamily.Click += self.export_to_family
        self.btnExportCSV.Click += self.export_csv
        self.btnExportHTML.Click += self.export_html
        self.btnDelete.Click += self.delete
        self.btnRefresh.Click += self.refresh
        self.btnClose.Click += self.close_window
        
        self.load_data()
        self.update_ui()
    
    def load_data(self):
        self.items = get_items(self.doc)
        self.filtered = list(self.items)
        self.cats = {}
        for item in self.items:
            c = item.category_name
            self.cats[c] = self.cats.get(c, 0) + 1
    
    def update_ui(self):
        self.txtTotal.Text = str(len(self.items))
        self.txtSelected.Text = "0"
        self.txtCategories.Text = str(len(self.cats))
        self.txtIssues.Text = str(len([i for i in self.items if i.has_issues]))
        
        self.lstCategories.Items.Clear()
        self.lstCategories.Items.Add("All ({})".format(len(self.items)))
        for c in sorted(self.cats.keys()):
            self.lstCategories.Items.Add("{} ({})".format(c, self.cats[c]))
        self.lstCategories.SelectedIndex = 0
        self.update_grid()
    
    def update_grid(self):
        self.dataGrid.Items.Clear()
        for item in self.filtered:
            self.dataGrid.Items.Add(item)
    
    def on_filter(self, s, e):
        search = self.txtSearch.Text.lower().strip() if self.txtSearch.Text else ""
        cat = None
        if self.lstCategories.SelectedIndex > 0:
            cat = str(self.lstCategories.SelectedItem).rsplit(" (", 1)[0]
        fi = self.cmbFilter.SelectedIndex
        
        self.filtered = []
        for item in self.items:
            if cat and item.category_name != cat: continue
            if fi == 1 and not item.has_issues: continue
            if fi == 2 and "No Name" not in item.issue_text: continue
            if fi == 3 and "Complex" not in item.issue_text: continue
            if fi == 4 and "Bad Name" not in item.issue_text: continue
            if fi == 5 and item.has_issues: continue
            if search and search not in "{} {} {}".format(item.name, item.family_name, item.category_name).lower(): continue
            self.filtered.append(item)
        self.update_grid()
    
    def on_selection(self, s, e):
        self.txtSelected.Text = str(self.dataGrid.SelectedItems.Count)
    
    def on_double_click(self, s, e):
        if self.dataGrid.SelectedItems.Count == 1:
            item = self.dataGrid.SelectedItem
            ids = List[ElementId]()
            ids.Add(ElementId(item.element_id))
            try:
                self.uidoc.ShowElements(ids)
                self.uidoc.Selection.SetElementIds(ids)
            except: pass
    
    def show_settings(self, s, e):
        dlg = SettingsDialog()
        dlg.ShowDialog()
        if dlg.result:
            self.load_data()
            self.update_ui()
    
    def select_all(self, s, e):
        self.dataGrid.SelectAll()
    
    def select_none(self, s, e):
        self.dataGrid.UnselectAll()
    
    def select_issues(self, s, e):
        self.dataGrid.UnselectAll()
        for item in self.filtered:
            if item.has_issues:
                self.dataGrid.SelectedItems.Add(item)
    
    def zoom_to(self, s, e):
        if self.dataGrid.SelectedItems.Count == 0:
            forms.alert("Select at least one item", title="Info")
            return
        ids = List[ElementId]()
        for item in self.dataGrid.SelectedItems:
            ids.Add(ElementId(item.element_id))
        try:
            self.uidoc.ShowElements(ids)
            self.uidoc.Selection.SetElementIds(ids)
        except Exception as ex:
            forms.alert(str(ex), title="Error")
    
    def select_in_model(self, s, e):
        if self.dataGrid.SelectedItems.Count == 0: return
        ids = List[ElementId]()
        for item in self.dataGrid.SelectedItems:
            ids.Add(ElementId(item.element_id))
        try:
            self.uidoc.Selection.SetElementIds(ids)
        except: pass
    
    def isolate(self, s, e):
        if self.dataGrid.SelectedItems.Count == 0: return
        ids = List[ElementId]()
        for item in self.dataGrid.SelectedItems:
            ids.Add(ElementId(item.element_id))
        try:
            with revit.Transaction("Isolate In-Place"):
                self.doc.ActiveView.IsolateElementsTemporary(ids)
        except: pass
    
    def rename(self, s, e):
        if self.dataGrid.SelectedItems.Count == 0:
            forms.alert("Select at least one item", title="Info")
            return
        
        selected = [item for item in self.dataGrid.SelectedItems]
        dlg = RenameDialog(selected)
        dlg.ShowDialog()
        
        if dlg.result:
            count = 0
            errors = []
            try:
                with revit.Transaction("Batch Rename"):
                    for item in selected:
                        try:
                            new_name = dlg.get_new_name(item.family_name)
                            if new_name != item.family_name:
                                el = self.doc.GetElement(ElementId(item.element_id))
                                if el and el.Symbol and el.Symbol.Family:
                                    el.Symbol.Family.Name = new_name
                                    count += 1
                        except Exception as ex:
                            errors.append("{}: {}".format(item.family_name, str(ex)))
                
                if errors:
                    forms.alert("Renamed {} item(s)\n\nErrors:\n{}".format(count, "\n".join(errors[:5])), title="Result")
                else:
                    forms.alert("Renamed {} family(s)".format(count), title="Success")
                self.load_data()
                self.update_ui()
            except Exception as ex:
                forms.alert(str(ex), title="Error")
    
    def delete(self, s, e):
        if self.dataGrid.SelectedItems.Count == 0:
            forms.alert("Select at least one item", title="Info")
            return
        
        count = self.dataGrid.SelectedItems.Count
        if not forms.alert("Delete {} item(s)?\n\nThis action CANNOT be undone!".format(count), 
                          title="Confirm Delete", yes=True, no=True):
            return
        
        deleted = 0
        try:
            with revit.Transaction("Delete In-Place"):
                for item in list(self.dataGrid.SelectedItems):
                    try:
                        self.doc.Delete(ElementId(item.element_id))
                        deleted += 1
                    except: pass
            forms.alert("Deleted {} item(s)".format(deleted), title="Success")
            self.load_data()
            self.update_ui()
        except Exception as ex:
            forms.alert(str(ex), title="Error")
    
    def export_csv(self, s, e):
        current_items = [item for item in self.dataGrid.Items]
        
        if not current_items:
            forms.alert("No data to export", title="Info")
            return
        
        from System.Windows.Forms import SaveFileDialog, DialogResult
        dlg = SaveFileDialog()
        dlg.Filter = "CSV Files|*.csv"
        dlg.FileName = "InPlace_{}.csv".format(datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        
        if dlg.ShowDialog() == DialogResult.OK:
            try:
                with codecs.open(dlg.FileName, 'w', 'utf-8-sig') as f:
                    f.write("ID,Family Name,Type,Category,Faces,Status,Issues,Level,Workset\n")
                    for item in current_items:
                        f.write("{},{},{},{},{},{},{},{},{}\n".format(
                            item.element_id, item.family_name, item.name,
                            item.category_name, item.face_count, item.complexity_status,
                            item.issue_text, item.level, item.workset
                        ))
                forms.alert("Exported {} items".format(len(current_items)), title="Success")
            except Exception as ex:
                forms.alert(str(ex), title="Error")
    
    def export_html(self, s, e):
        current_items = [item for item in self.dataGrid.Items]
        
        if not current_items:
            forms.alert("No data to export", title="Info")
            return
        
        from System.Windows.Forms import SaveFileDialog, DialogResult
        dlg = SaveFileDialog()
        dlg.Filter = "HTML Files|*.html"
        dlg.FileName = "InPlace_Report_{}.html".format(datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        
        if dlg.ShowDialog() == DialogResult.OK:
            try:
                title = self.doc.Title if self.doc.Title else "Untitled"
                html = generate_html_report(current_items, title)
                with codecs.open(dlg.FileName, 'w', 'utf-8') as f:
                    f.write(html)
                forms.alert("Report saved!", title="Success")
                os.startfile(dlg.FileName)
            except Exception as ex:
                forms.alert(str(ex), title="Error")
    
    # ========================================================================
    # EXPORT TO FAMILY - with Auto Load & Place
    # ========================================================================
    def export_to_family(self, s, e):
        """Export In-Place to Loadable Family with option to auto-place"""
        if self.dataGrid.SelectedItems.Count == 0:
            forms.alert("Select at least one item", title="Info")
            return
        
        selected = [item for item in self.dataGrid.SelectedItems]
        
        # Ask user for export mode
        options = ["Export .rfa only", "Export + Load + Place"]
        selected_option = forms.CommandSwitchWindow.show(options, message="Select export mode:")
        
        if not selected_option:
            return
        
        auto_place = (selected_option == options[1])
        
        # Select folder
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        folder_dlg = FolderBrowserDialog()
        folder_dlg.Description = "Select folder to save Family files (.rfa)"
        
        if folder_dlg.ShowDialog() != DialogResult.OK:
            return
        
        save_path = folder_dlg.SelectedPath
        
        exported = 0
        placed = 0
        errors = []
        
        for item in selected:
            fam_doc = None
            try:
                # Get element
                el = self.doc.GetElement(ElementId(item.element_id))
                if not el:
                    errors.append("{}: Element not found".format(item.family_name))
                    continue
                
                # Get geometry
                opt = Options()
                opt.ComputeReferences = True
                opt.DetailLevel = ViewDetailLevel.Fine
                geom = el.get_Geometry(opt)
                
                if not geom:
                    errors.append("{}: No geometry found".format(item.family_name))
                    continue
                
                # Get bounding box for positioning
                bb = geom.GetBoundingBox()
                if not bb:
                    errors.append("{}: Cannot get bounding box".format(item.family_name))
                    continue
                
                # Store original position
                family_origin = bb.Min
                
                # Collect solids with subcategory info
                solids_dict = {}
                self._collect_geometry_with_subcat(geom, bb.Min, solids_dict)
                
                if not solids_dict:
                    errors.append("{}: No solid geometry found".format(item.family_name))
                    continue
                
                # Find template
                template_path = self._find_family_template(el.Category.Id.IntegerValue if el.Category else -2000151)
                
                if not template_path:
                    errors.append("{}: Cannot find family template".format(item.family_name))
                    continue
                
                # Create new family document
                app = self.doc.Application
                fam_doc = app.NewFamilyDocument(template_path)
                
                if not fam_doc:
                    errors.append("{}: Cannot create family document".format(item.family_name))
                    continue
                
                # Generate unique family name
                project_num = self.doc.ProjectInformation.Number if self.doc.ProjectInformation.Number else "000"
                base_fam_name = "{}_{}".format(project_num, item.family_name)
                base_fam_name = re.sub(r'[<>:"/\\|?*]', '_', base_fam_name).strip()
                fam_name = self._get_unique_family_name(base_fam_name)
                
                # Add geometry to family
                with Transaction(fam_doc, "Add Geometry") as t:
                    t.Start()
                    
                    parent_cat = fam_doc.OwnerFamily.FamilyCategory
                    
                    for solid, subcat_name in solids_dict.items():
                        try:
                            if solid.Volume > 0:
                                copied_geo = FreeFormElement.Create(fam_doc, solid)
                                
                                # Assign subcategory if exists
                                if subcat_name and copied_geo:
                                    try:
                                        cats = fam_doc.Settings.Categories
                                        if cats.Contains(subcat_name):
                                            for cat in fam_doc.Settings.Categories:
                                                if cat.Name == subcat_name:
                                                    copied_geo.Subcategory = cat
                                                    break
                                        else:
                                            new_subcat = cats.NewSubcategory(parent_cat, subcat_name)
                                            copied_geo.Subcategory = new_subcat
                                    except:
                                        pass
                        except:
                            pass
                    
                    t.Commit()
                
                # Save family
                safe_name = re.sub(r'[<>:"/\\|?*]', '_', fam_name)
                file_path = os.path.join(save_path, "{}.rfa".format(safe_name))
                
                counter = 1
                while os.path.exists(file_path):
                    file_path = os.path.join(save_path, "{}_{}.rfa".format(safe_name, counter))
                    counter += 1
                
                save_options = SaveAsOptions()
                save_options.OverwriteExistingFile = True
                fam_doc.SaveAs(file_path, save_options)
                fam_doc.Close(False)
                fam_doc = None
                
                exported += 1
                
                # Auto load and place
                if auto_place:
                    try:
                        with revit.Transaction("Load and Place Family"):
                            # Load family
                            self.doc.LoadFamily(file_path)
                            self.doc.Regenerate()
                            
                            # Find family symbol
                            fam_symbol = None
                            symbols = FilteredElementCollector(self.doc).OfClass(FamilySymbol).ToElements()
                            
                            for sym in symbols:
                                try:
                                    sym_fam_name = sym.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
                                    if sym_fam_name and fam_name in sym_fam_name:
                                        fam_symbol = sym
                                        break
                                except:
                                    pass
                            
                            if fam_symbol:
                                if not fam_symbol.IsActive:
                                    fam_symbol.Activate()
                                    self.doc.Regenerate()
                                
                                # Place at original position
                                str_type = StructuralType.NonStructural
                                new_instance = self.doc.Create.NewFamilyInstance(
                                    family_origin, 
                                    fam_symbol, 
                                    str_type
                                )
                                placed += 1
                    except Exception as place_ex:
                        errors.append("{}: Load error - {}".format(item.family_name, str(place_ex)))
                
            except Exception as ex:
                errors.append("{}: {}".format(item.family_name, str(ex)))
            finally:
                if fam_doc:
                    try:
                        fam_doc.Close(False)
                    except:
                        pass
        
        # Show result
        if auto_place:
            msg = "Exported: {}\nLoaded & Placed: {}".format(exported, placed)
        else:
            msg = "Exported: {} family(s)".format(exported)
        
        if errors:
            error_msg = "\n".join(errors[:5])
            if len(errors) > 5:
                error_msg += "\n... and {} more".format(len(errors) - 5)
            forms.alert("{}\n\nSome errors:\n{}".format(msg, error_msg), title="Export Result")
        else:
            forms.alert("{}\n\nSaved to: {}".format(msg, save_path), title="Success")
    
    def _collect_geometry_with_subcat(self, geom, ref_point, solids_dict):
        """Collect solids with subcategory info, transformed to origin"""
        transform = Transform.CreateTranslation(XYZ(-ref_point.X, -ref_point.Y, -ref_point.Z))
        
        for g in geom:
            if isinstance(g, Solid) and g.Volume > 0:
                try:
                    new_solid = SolidUtils.CreateTransformed(g, transform)
                    subcat_name = self._get_subcat_name(g)
                    solids_dict[new_solid] = subcat_name
                except:
                    solids_dict[g] = None
            elif isinstance(g, GeometryInstance):
                inst_geom = g.GetInstanceGeometry()
                if inst_geom:
                    for ig in inst_geom:
                        if isinstance(ig, Solid) and ig.Volume > 0:
                            try:
                                new_solid = SolidUtils.CreateTransformed(ig, transform)
                                subcat_name = self._get_subcat_name(ig)
                                solids_dict[new_solid] = subcat_name
                            except:
                                solids_dict[ig] = None
    
    def _get_subcat_name(self, geometry):
        """Get subcategory name from geometry"""
        try:
            subcat_id = geometry.GraphicsStyleId
            if subcat_id and subcat_id.IntegerValue != -1:
                style = self.doc.GetElement(subcat_id)
                if style:
                    return style.Name
        except:
            pass
        return None
    
    def _find_family_template(self, category_id):
        """Find appropriate family template"""
        app = self.doc.Application
        template_path = app.FamilyTemplatePath
        
        generic_templates = [
            "Metric Generic Model.rft",
            "Generic Model.rft",
            "Metric Generic Model face based.rft",
            "Generic Model face based.rft"
        ]
        
        for tname in generic_templates:
            tpath = os.path.join(template_path, tname)
            if os.path.exists(tpath):
                return tpath
            tpath = os.path.join(template_path, "English", tname)
            if os.path.exists(tpath):
                return tpath
        
        return None
    
    def _get_unique_family_name(self, base_name):
        """Generate unique family name"""
        fam_name = base_name
        
        existing = FilteredElementCollector(self.doc).OfClass(Family).ToElements()
        existing_names = [f.Name for f in existing]
        
        counter = 1
        while fam_name in existing_names:
            fam_name = "{}_Copy{}".format(base_name, counter)
            counter += 1
        
        return fam_name
    
    def refresh(self, s, e):
        self.load_data()
        self.update_ui()
        forms.alert("Found {} In-Place models".format(len(self.items)), title="Refresh")
    
    def close_window(self, s, e):
        self.Close()


# ============================================================================
# MAIN
# ============================================================================
if __name__ == "__main__":
    try:
        if not revit.doc:
            forms.alert("Please open a project first.", title="In-Place Manager")
        else:
            InPlaceManagerWindow().ShowDialog()
    except Exception as ex:
        forms.alert("Error: {}".format(str(ex)), title="In-Place Manager Error")