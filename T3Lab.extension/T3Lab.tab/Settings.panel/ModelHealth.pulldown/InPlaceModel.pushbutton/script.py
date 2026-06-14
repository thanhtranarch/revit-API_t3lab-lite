# -*- coding: utf-8 -*-
"""In-Place Model Manager v1.7 - Full Features
Author: Dang Quoc Truong (DQT)
Fixed: Revit 2026 compatibility, Sweep/Curve geometry support
"""
__title__ = "In-Place\nManager"
__author__ = "DQT"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from pyrevit import revit, forms, script, HOST_APP, DB
from pyrevit.forms import WPFWindow
from pyrevit.compat import get_elementid_value_func
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from System.Collections.Generic import List
import System
import re, datetime, codecs, os, tempfile

def get_xaml_path(filename):
    current_dir = os.path.dirname(__file__)
    while current_dir and not current_dir.endswith('T3Lab.extension'):
        parent = os.path.dirname(current_dir)
        if parent == current_dir:
            break
        current_dir = parent
    return os.path.join(current_dir, "lib", "GUI", "Tools", filename)


# ============================================================================
# REVIT 2026 COMPATIBILITY - Use pyrevit.compat
# ============================================================================
get_elementid_value = get_elementid_value_func()

def _eid_int(eid):
    """Get integer value from ElementId - compatible with Revit 2024-2026"""
    if eid is None:
        return -1
    return get_elementid_value(eid)

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
                d.element_id = _eid_int(el.Id)
                d.name = el.Name if el.Name else "<No Name>"
                d.family_name = el.Symbol.Family.Name
                d.category_name = el.Category.Name if el.Category else "N/A"
                
                # Get Workset info
                d.workset = "-"
                try:
                    if doc.IsWorkshared:
                        ws_id = el.WorksetId
                        if ws_id and _eid_int(ws_id) > 0:
                            wst = doc.GetWorksetTable()
                            ws = wst.GetWorkset(ws_id)
                            if ws:
                                d.workset = ws.Name
                except: 
                    pass
                
                # Level
                d.level = "-"
                try:
                    if el.LevelId and _eid_int(el.LevelId) != -1:
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
        color = "#10B981" if i.complexity_status == "OK" else ("#F59E0B" if i.complexity_status == "Warning" else "#EF4444")
        iss = '<span style="color:#10B981">OK</span>' if not i.has_issues else '<span style="color:#EF4444">{}</span>'.format(i.issue_text)
        item_rows += "<tr><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td>{}</td><td style='color:{}'>{}</td><td>{}</td><td>{}</td><td>{}</td></tr>".format(
            i.element_id, i.family_name, i.name, i.category_name, i.face_count, color, i.complexity_status, iss, i.level, i.workset)
    
    return """<!DOCTYPE html><html><head><meta charset="UTF-8"><title>In-Place Report - {}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f5f5f5;padding:20px;color:#333}}
.container{{max-width:1400px;margin:0 auto}}
.header{{background:linear-gradient(135deg,#0F172A,#E5B85C);padding:25px;border-radius:10px;margin-bottom:20px}}
.header h1{{font-size:24px;margin-bottom:5px}}
.header .meta{{font-size:12px;opacity:0.8}}
.header .author{{font-size:11px;margin-top:8px;padding-top:8px;border-top:1px solid rgba(255,255,255,0.3)}}
.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin-bottom:20px}}
.card{{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1)}}
.card .label{{font-size:11px;color:#666;text-transform:uppercase}}
.card .value{{font-size:32px;font-weight:bold;margin-top:5px}}
.card.ok .value{{color:#10B981}}
.card.err .value{{color:#EF4444}}
.section{{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1);margin-bottom:20px}}
.section h2{{font-size:16px;margin-bottom:15px;padding-bottom:10px;border-bottom:2px solid #0F172A}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th,td{{padding:10px;text-align:left;border-bottom:1px solid #eee}}
th{{background:#0F172A;font-weight:600}}
tr:hover{{background:#fafafa}}
.footer{{text-align:center;padding:20px;color:#888;font-size:11px;background:#fff;border-radius:8px;margin-top:20px}}
.footer .copyright{{font-weight:600;color:#666;margin-bottom:5px}}
</style></head>
<body><div class="container">
<div class="header">
<h1>In-Place Model Report</h1>
<div class="meta">{} | Generated: {}</div>
<div class="author">Created by: <strong>Dang Quoc Truong (DQT)</strong> | In-Place Manager v1.7</div>
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
<div>In-Place Manager v1.7 | Generated: {}</div>
</div>
</div></body></html>""".format(doc_title, doc_title, ts, total, total-issues, issues, len(cats), cat_rows, total, item_rows, ts)

# ============================================================================
# XAML
# ============================================================================






# ============================================================================
# SETTINGS DIALOG
# ============================================================================
class SettingsDialog(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, get_xaml_path('InPlaceModelSettings.xaml'))
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
        WPFWindow.__init__(self, get_xaml_path('InPlaceModelRename.xaml'))
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
        WPFWindow.__init__(self, get_xaml_path('InPlaceModelMain.xaml'))
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
    # EXPORT TO FAMILY - Full rewrite based on reference script
    # ========================================================================
    def export_to_family(self, s, e):
        """Export In-Place to Loadable Family with option to auto-place
        Based on pychilizer InPlace to Loadable script
        """
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
        
        # Select folder for saving
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
            new_family_doc = None
            try:
                # Get element
                el = self.doc.GetElement(ElementId(item.element_id))
                if not el:
                    errors.append("{}: Element not found".format(item.family_name))
                    continue
                
                # Get geometry
                geo_element = el.get_Geometry(Options())
                if not geo_element:
                    errors.append("{}: No geometry found".format(item.family_name))
                    continue
                
                # Get bounding box for positioning
                bb = geo_element.GetBoundingBox()
                if not bb:
                    errors.append("{}: Cannot get bounding box".format(item.family_name))
                    continue
                
                # Store original position
                family_origin = bb.Min
                
                # Collect solids and curves with subcategory info
                solids_dict = {}  # solid -> subcategory_name
                curves_list = []
                
                # Create inverse transform
                def inverted_transform(ref_point):
                    transform = Transform.CreateTranslation(ref_point)
                    return transform.Inverse
                
                inv_transform = inverted_transform(bb.Min)
                
                # Iterate through GeometryInstance
                for instance_geo in geo_element:
                    if isinstance(instance_geo, GeometryInstance):
                        geometry_element = instance_geo.GetInstanceGeometry()
                        for geometry in geometry_element:
                            # Collect solids
                            if isinstance(geometry, Solid) and geometry.Volume > 0:
                                try:
                                    new_solid = SolidUtils.CreateTransformed(geometry, inv_transform)
                                    subcat_name = self._get_subcat_name(geometry)
                                    solids_dict[new_solid] = subcat_name
                                except:
                                    solids_dict[geometry] = None
                            # Collect curves
                            elif isinstance(geometry, Curve):
                                try:
                                    new_curve = geometry.CreateTransformed(inv_transform)
                                    curves_list.append(new_curve)
                                except:
                                    pass
                    # Also check if instance_geo itself is a Solid
                    elif isinstance(instance_geo, Solid) and instance_geo.Volume > 0:
                        try:
                            new_solid = SolidUtils.CreateTransformed(instance_geo, inv_transform)
                            subcat_name = self._get_subcat_name(instance_geo)
                            solids_dict[new_solid] = subcat_name
                        except:
                            solids_dict[instance_geo] = None
                
                if not solids_dict:
                    errors.append("{}: No solid geometry found".format(item.family_name))
                    continue
                
                # Find template based on category
                el_cat_id = _eid_int(el.Category.Id) if el.Category else -2000151
                template_path = self._find_family_template(el_cat_id)
                
                if not template_path:
                    errors.append("{}: Cannot find family template".format(item.family_name))
                    continue
                
                # Create new family document
                app = self.doc.Application
                new_family_doc = app.NewFamilyDocument(template_path)
                
                if not new_family_doc:
                    errors.append("{}: Cannot create family document".format(item.family_name))
                    continue
                
                # Generate unique family name
                project_num = self.doc.ProjectInformation.Number if self.doc.ProjectInformation.Number else "000"
                base_fam_name = "{}_{}".format(project_num, item.family_name)
                base_fam_name = re.sub(r'[<>:"/\\|?*]', '_', base_fam_name).strip()
                fam_name = self._get_unique_family_name(base_fam_name)
                
                # Save family first (required before adding geometry in some cases)
                fam_path = os.path.join(save_path, "{}.rfa".format(fam_name))
                counter = 1
                while os.path.exists(fam_path):
                    fam_path = os.path.join(save_path, "{}_{}.rfa".format(fam_name, counter))
                    counter += 1
                
                saveas_opt = SaveAsOptions()
                saveas_opt.OverwriteExistingFile = True
                new_family_doc.SaveAs(fam_path, saveas_opt)
                
                # Add geometry to family
                with Transaction(new_family_doc, "Copy Geometry") as t:
                    t.Start()
                    
                    parent_cat = new_family_doc.OwnerFamily.FamilyCategory
                    
                    for solid, subcat_name in solids_dict.items():
                        try:
                            if solid.Volume > 0:
                                copied_geo = FreeFormElement.Create(new_family_doc, solid)
                                
                                # Assign subcategory if exists
                                if subcat_name and copied_geo:
                                    try:
                                        cats = new_family_doc.Settings.Categories
                                        if cats.Contains(subcat_name):
                                            for cat in new_family_doc.Settings.Categories:
                                                if cat.Name == subcat_name:
                                                    copied_geo.Subcategory = cat
                                                    break
                                        else:
                                            new_subcat = cats.NewSubcategory(parent_cat, subcat_name)
                                            copied_geo.Subcategory = new_subcat
                                    except:
                                        pass
                        except Exception as geo_ex:
                            pass
                    
                    # Add curves (model lines) - for sweep paths etc.
                    for curve in curves_list:
                        try:
                            if isinstance(curve, Line):
                                # Create sketch plane for line
                                p1 = curve.Evaluate(0, True)
                                tangent = curve.ComputeDerivatives(0, True).BasisX
                                normal = tangent.CrossProduct(p1)
                                if normal.GetLength() > 0.001:
                                    plane = Plane.CreateByNormalAndOrigin(normal, p1)
                                    sk_plane = SketchPlane.Create(new_family_doc, plane)
                                    new_family_doc.FamilyCreate.NewModelCurve(curve, sk_plane)
                            else:
                                # For arcs, ellipses, etc.
                                deriv = curve.ComputeDerivatives(0.5, True)
                                if deriv.BasisZ and deriv.BasisZ.GetLength() > 0.001:
                                    normal = deriv.BasisZ.Normalize()
                                    p1 = curve.Evaluate(0, True)
                                    plane = Plane.CreateByNormalAndOrigin(normal, p1)
                                    sk_plane = SketchPlane.Create(new_family_doc, plane)
                                    new_family_doc.FamilyCreate.NewModelCurve(curve, sk_plane)
                        except:
                            pass
                    
                    new_family_doc.Regenerate()
                    t.Commit()
                
                # Save and close family
                save_opt = SaveOptions()
                new_family_doc.Save(save_opt)
                new_family_doc.Close(False)
                new_family_doc = None
                
                exported += 1
                
                # Auto load and place
                if auto_place:
                    try:
                        with revit.Transaction("Load and Place Family"):
                            # Load family
                            loaded = revit.db.create.load_family(fam_path, doc=self.doc)
                            self.doc.Regenerate()
                            
                            # Find family symbol
                            fam_symbol = None
                            symbols = FilteredElementCollector(self.doc).OfClass(FamilySymbol).WhereElementIsElementType().ToElements()
                            
                            for sym in symbols:
                                try:
                                    type_name = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                                    if type_name and type_name.strip() == fam_name:
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
                            else:
                                errors.append("{}: Family loaded but symbol not found".format(item.family_name))
                    except Exception as place_ex:
                        errors.append("{}: Load/Place error - {}".format(item.family_name, str(place_ex)))
                
            except Exception as ex:
                errors.append("{}: {}".format(item.family_name, str(ex)))
            finally:
                if new_family_doc:
                    try:
                        new_family_doc.Close(False)
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
    
    def _get_subcat_name(self, geometry):
        """Get subcategory name from geometry"""
        try:
            subcat_id = geometry.GraphicsStyleId
            if subcat_id and str(subcat_id) != "-1":
                style = self.doc.GetElement(subcat_id)
                if style:
                    return style.Name
        except:
            pass
        return None
    
    def _find_family_template(self, category_id):
        """Find appropriate family template based on category"""
        app = self.doc.Application
        template_path = app.FamilyTemplatePath
        
        # Try generic templates first
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
            # Try English subfolder
            tpath = os.path.join(template_path, "English", tname)
            if os.path.exists(tpath):
                return tpath
        
        return None
    
    def _get_unique_family_name(self, base_name):
        """Generate unique family name that doesn't exist in project"""
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