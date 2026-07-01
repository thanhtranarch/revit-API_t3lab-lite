# -*- coding: utf-8 -*-
"""ManaSelect — unified controller for smart selection tools."""

import os
import sys
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import System
import System.Windows
from System.Windows import WindowState, Visibility
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSheet,
    ImportInstance,
    BuiltInCategory,
    FamilyInstance,
    ElementId
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit, DB, script

# Add current folder to sys.path to find GUI dialog classes
sys.path.append(os.path.dirname(__file__))
# Add parent of Selection folder to find Selection
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

import QuickElementDialog

# Import dqt selection logic
from Selection.dqt_select import core as dqt_core
from Selection.dqt_select import compat as dqt_compat

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ManaSelect.xaml')


class ManaSelectWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)
        self.uidoc = revit.uidoc
        self.doc = revit.doc

        # Initialize and nest sub-panels
        self._init_sub_panels()

        # Connect navigation (left icon-rail sidebar toggles)
        self.nav_toggle_quick_select.Click += self._on_nav_toggle_clicked
        self.nav_toggle_select_similar.Click += self._on_nav_toggle_clicked
        self.nav_toggle_select_sheets.Click += self._on_nav_toggle_clicked
        self.nav_toggle_quick_actions.Click += self._on_nav_toggle_clicked

        # Connect run actions for tabs
        self.btn_run_select_similar.Click += self._on_run_select_similar
        self.btn_run_select_sheets.Click += self._on_run_select_sheets

        # Connect sidebar buttons
        self.btn_select_by_cat.Click += self._on_select_by_category
        self.btn_select_linked.Click += self._on_select_linked
        self.btn_select_inplace.Click += self._on_select_inplace
        self.btn_deselect_grouped.Click += self._on_deselect_grouped
        self.btn_material_select.Click += self._on_material_select

        # Chrome actions
        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close_chrome.Click += self._close_chrome

    def _init_sub_panels(self):
        """Loads Quick Select window grid and collapses its header to avoid duplicates."""
        try:
            self._quick_select_win = QuickElementDialog.QuickSelectWindow()
            quick_select_border = self._quick_select_win.Content
            # Content is Border, its Child is Grid
            quick_select_grid = quick_select_border.Child
            self._quick_select_win.Content = None
            self.grid_quick_select.Children.Add(quick_select_border)

            # Collapse sub-tool header and footer to unify status bar
            quick_select_grid.RowDefinitions[0].Height = System.Windows.GridLength(0)
            quick_select_grid.RowDefinitions[4].Height = System.Windows.GridLength(0)

            # Re-wire close
            self._quick_select_win.Close = self.Close
        except Exception as ex:
            print("Error loading Quick Select panel: {}".format(ex))

    def _on_nav_toggle_clicked(self, sender, e):
        """Switch active TabControl index, sync rail toggle state and update status bar text."""
        if sender == self.nav_toggle_quick_select:
            index = 0
            self.status_text.Text = "Quick Select — Query elements by categories, parameters and text filters"
        elif sender == self.nav_toggle_select_similar:
            index = 1
            self.status_text.Text = "Select Similar — Match Type, Family or Category of current selection"
        elif sender == self.nav_toggle_select_sheets:
            index = 2
            self.status_text.Text = "Select on Sheets — Find CAD imports or title blocks across drawings"
        elif sender == self.nav_toggle_quick_actions:
            index = 3
            self.status_text.Text = "Quick Actions — One-click viewport filters and cleanup tools"
        else:
            return

        self.main_tab_control.SelectedIndex = index
        self.nav_toggle_quick_select.IsChecked = (index == 0)
        self.nav_toggle_select_similar.IsChecked = (index == 1)
        self.nav_toggle_select_sheets.IsChecked = (index == 2)
        self.nav_toggle_quick_actions.IsChecked = (index == 3)

    # =========================================================================
    # TAB 2: SELECT SIMILAR
    # =========================================================================
    def _on_run_select_similar(self, sender, e):
        """Execute select similar based on UI configurations."""
        if not self.uidoc.Selection.GetElementIds():
            forms.alert("Please select at least one seed element in the model first.", title="Select Similar")
            return

        scope = 'view' if self.rb_similar_scope_view.IsChecked else 'model'
        try:
            if self.rb_similar_mode_type.IsChecked:
                dqt_core.select_similar_type(mode=scope)
            elif self.rb_similar_mode_family.IsChecked:
                dqt_core.select_similar_family(mode=scope)
            else:
                dqt_core.select_similar_category(mode=scope)
        except Exception as ex:
            forms.alert("Error running Select Similar: {}".format(ex), title="Error")

    # =========================================================================
    # TAB 3: SELECT ON SHEETS
    # =========================================================================
    def _on_run_select_sheets(self, sender, e):
        """Execute selection on sheets based on target selection."""
        use_dwg = bool(self.rb_sheet_target_dwg.IsChecked)
        if use_dwg:
            sheets = self._get_target_sheets('Select DWGs', 'DQT - On Sheets: CAD Imports')
            if sheets:
                self._select_dwgs(sheets)
        else:
            sheets = self._get_target_sheets('Select Title Blocks', 'DQT - On Sheets: Title Blocks')
            if sheets:
                self._select_title_blocks(sheets)

    def _get_target_sheets(self, button_name, alert_title):
        sel_ids = self.uidoc.Selection.GetElementIds()
        sheets = [self.doc.GetElement(i) for i in sel_ids
                  if isinstance(self.doc.GetElement(i), ViewSheet)]
        if sheets:
            return sheets

        all_sheets = FilteredElementCollector(self.doc).OfClass(ViewSheet).ToElements()
        if not all_sheets:
            forms.alert('There are no sheets in this model.', title=alert_title)
            return None

        sheet_map = {'{} - {}'.format(s.SheetNumber, s.Name): s for s in all_sheets}
        chosen = forms.SelectFromList.show(
            sorted(sheet_map.keys()),
            title='DQT - Pick Sheets',
            button_name=button_name,
            multiselect=True,
        )
        if not chosen:
            return None
        return [sheet_map[c] for c in chosen]

    def _select_dwgs(self, sheets):
        sheet_ids = set(dqt_compat.eid_int(s.Id) for s in sheets)
        all_imports = (FilteredElementCollector(self.doc)
                       .OfClass(ImportInstance)
                       .WhereElementIsNotElementType()
                       .ToElements())

        dwg_ids = [imp.Id for imp in all_imports
                   if dqt_compat.eid_int(imp.OwnerViewId) in sheet_ids]

        if dwg_ids:
            self.uidoc.Selection.SetElementIds(dqt_compat.to_element_id_list(dwg_ids))
            dqt_compat.notify('Selected {} DWG(s) on {} sheet(s).'.format(
                       len(dwg_ids), len(sheets)),
                    title='DQT - On Sheets: CAD Imports')
        else:
            forms.alert('No DWGs found on the selected sheets.',
                        title='DQT - On Sheets: CAD Imports')

    def _select_title_blocks(self, sheets):
        sheet_ids = set(dqt_compat.eid_int(s.Id) for s in sheets)
        all_tb = (FilteredElementCollector(self.doc)
                  .OfCategory(BuiltInCategory.OST_TitleBlocks)
                  .WhereElementIsNotElementType()
                  .ToElements())

        tb_ids = [tb.Id for tb in all_tb
                  if dqt_compat.eid_int(tb.OwnerViewId) in sheet_ids]

        if tb_ids:
            self.uidoc.Selection.SetElementIds(dqt_compat.to_element_id_list(tb_ids))
            dqt_compat.notify('Selected {} title block(s) on {} sheet(s).'.format(
                       len(tb_ids), len(sheets)),
                    title='DQT - On Sheets: Title Blocks')
        else:
            forms.alert('No title blocks found on the selected sheets.',
                        title='DQT - On Sheets: Title Blocks')

    # =========================================================================
    # SIDEBAR: QUICK ACTIONS
    # =========================================================================
    def _on_select_by_category(self, sender, e):
        self.Hide()
        try:
            self._run_select_by_category()
        finally:
            self.Show()

    def _run_select_by_category(self):
        def is_valid_category(cat):
            try:
                if cat is None: return False
                if not cat.CanAddSubcategory and cat.CategoryType == DB.CategoryType.Internal: return False
                app = self.doc.Application
                rvt_year = int(app.VersionNumber)
                if rvt_year > 2022:
                    try:
                        if cat.BuiltInCategory == BuiltInCategory.INVALID: return False
                    except AttributeError:
                        pass
                return True
            except Exception:
                return False

        class CategorySelectionFilter(ISelectionFilter):
            def __init__(self, allowed_cat_int_ids):
                self.allowed = set(allowed_cat_int_ids)
            def AllowElement(self, element):
                try:
                    if element.Category is None: return False
                    return dqt_compat.eid_int(element.Category.Id) in self.allowed
                except Exception:
                    return False
            def AllowReference(self, reference, position):
                return False

        all_cats = [c for c in self.doc.Settings.Categories if is_valid_category(c)]
        for extra_bic in (BuiltInCategory.OST_Grids, BuiltInCategory.OST_Levels, BuiltInCategory.OST_Viewports):
            try:
                extra = self.doc.Settings.Categories.get_Item(extra_bic)
                if extra is not None and extra not in all_cats:
                    all_cats.append(extra)
            except Exception:
                pass

        cat_map = {c.Name: c for c in all_cats}
        names = sorted(cat_map.keys())

        chosen = forms.SelectFromList.show(
            names,
            title='DQT - Select By Category',
            button_name='Select',
            multiselect=True,
        )
        if not chosen:
            return

        allowed_ids = [dqt_compat.eid_int(cat_map[n].Id) for n in chosen]
        sel_filter = CategorySelectionFilter(allowed_ids)

        try:
            refs = self.uidoc.Selection.PickObjects(
                ObjectType.Element, sel_filter,
                'Pick elements (only chosen categories) then click Finish')
        except OperationCanceledException:
            refs = []

        picked_ids = [self.doc.GetElement(r).Id for r in refs if self.doc.GetElement(r)]
        if picked_ids:
            self.uidoc.Selection.SetElementIds(dqt_compat.to_element_id_list(picked_ids))
            dqt_compat.notify('Selected {} element(s).'.format(len(picked_ids)),
                        title='DQT - Select By Category')

    def _on_select_linked(self, sender, e):
        self.Hide()
        try:
            self._run_select_linked()
        except Exception as ex:
            forms.alert("Error selecting linked elements: {}".format(ex), title="Error")
        finally:
            self.Show()

    def _run_select_linked(self):
        from Autodesk.Revit.DB import RevitLinkInstance, Reference, ViewType
        
        def is_view_supported(view):
            if view is None or view.IsTemplate: return False
            unsupported = {ViewType.Schedule, ViewType.ProjectBrowser, ViewType.SystemBrowser, ViewType.Internal, ViewType.Undefined}
            return view.ViewType not in unsupported

        active_view = self.doc.ActiveView
        if not is_view_supported(active_view):
            forms.alert("The active view does not support element selection.\nPlease open a Plan, Section, Elevation, 3D, Drafting or Sheet view.", title="Select All Linked")
            return

        all_links = list(FilteredElementCollector(self.doc).OfClass(RevitLinkInstance).ToElements())
        if not all_links:
            forms.alert("No Revit Links found in this project.", title="Select All Linked")
            return

        label_to_link = {}
        labels = []
        for li in all_links:
            try:
                ldoc = li.GetLinkDocument()
                doc_title = ldoc.Title if ldoc else "<Not Loaded>"
            except Exception:
                doc_title = "<Not Loaded>"
             
            label = "{}  |  {}".format(li.Name, doc_title)
            labels.append(label)
            label_to_link[label] = li

        chosen_label = forms.SelectFromList.show(
            labels,
            title="Select Revit Link",
            button_name="Select All Elements",
            multiselect=False,
            width=520,
            height=420,
        )
        if not chosen_label:
            return

        link_inst = label_to_link[chosen_label]
        link_doc = link_inst.GetLinkDocument()
        link_title = link_doc.Title if link_doc else link_inst.Name

        if link_doc is None:
            forms.alert("The selected link is not loaded. Please reload it first.", title="Select All Linked")
            return

        refs = List[Reference]()
        total_seen = 0
        fail_count = 0

        for el in FilteredElementCollector(link_doc).WhereElementIsNotElementType():
            if el.Category is None:
                continue
            try:
                if el.ViewSpecific: continue
            except Exception:
                pass
            total_seen += 1
            try:
                r = Reference(el).CreateLinkReference(link_inst)
                if r is not None:
                     refs.Add(r)
            except Exception:
                fail_count += 1

        if refs.Count == 0:
            forms.alert("Could not build any selectable references for the {} linked elements found.".format(total_seen), title="Select All Linked")
            return

        try:
            self.uidoc.Selection.SetReferences(refs)
            self.uidoc.RefreshActiveView()
            lines = [
                 "Done.", "",
                 "Link:     {}".format(link_title),
                 "View:     {}".format(active_view.Name),
                 "Selected: {} linked elements".format(refs.Count)
            ]
            if fail_count > 0:
                 lines.append("Skipped:  {} (no selectable Reference)".format(fail_count))
            forms.alert("\n".join(lines), title="Select All Linked")
        except Exception as ex:
             forms.alert("Failed to apply selection: {}".format(ex), title="Select All Linked")

    def _on_select_inplace(self, sender, e):
        self.Hide()
        try:
            self._run_select_inplace()
        finally:
            self.Show()

    def _run_select_inplace(self):
        scope = forms.CommandSwitchWindow.show(
            ['Active View', 'Whole Model'],
            message='Where do you want to select In-Place elements?',
        )
        if not scope:
            return

        view_id = self.doc.ActiveView.Id if scope == 'Active View' else None
        if view_id is not None:
             collector = FilteredElementCollector(self.doc, view_id)
        else:
             collector = FilteredElementCollector(self.doc)

        instances = collector.OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements()
        result = []
        for fi in instances:
             try:
                 if fi.Symbol and fi.Symbol.Family and fi.Symbol.Family.IsInPlace:
                     result.append(fi)
             except Exception:
                 continue

        if not result:
             forms.alert('No In-Place elements found in {}.'.format(scope.lower()), title='Select In-Place')
             return

        ids = [e.Id for e in result]
        self.uidoc.Selection.SetElementIds(dqt_compat.to_element_id_list(ids))
        dqt_compat.notify('Selected {} In-Place element(s) in {}.'.format(len(ids), scope.lower()), title='Select In-Place')

    def _on_deselect_grouped(self, sender, e):
        selected_ids = self.uidoc.Selection.GetElementIds()
        if not selected_ids or selected_ids.Count == 0:
            forms.alert('Nothing is selected. Select some elements first.', title='Deselect Grouped')
            return

        kept = []
        removed = 0
        for eid in selected_ids:
             elem = self.doc.GetElement(eid)
             if elem is None:
                 continue
             if dqt_compat.eid_int(elem.GroupId) == dqt_compat.INVALID_ELEMENT_ID_INT:
                 kept.append(elem.Id)
             else:
                 removed += 1

        self.uidoc.Selection.SetElementIds(dqt_compat.to_element_id_list(kept))
        dqt_compat.notify('Removed {} grouped element(s). {} kept.'.format(removed, len(kept)), title='Deselect Grouped')

    def _on_material_select(self, sender, e):
        self.Hide()
        try:
            self._run_material_select()
        finally:
            self.Show()

    def _run_material_select(self):
        def get_all_materials():
             try:
                 collector = FilteredElementCollector(self.doc)
                 materials = collector.OfClass(DB.Material).ToElements()
                 return [mat for mat in materials if mat and mat.IsValidObject]
             except Exception as ex:
                 print("Error getting materials: {}".format(ex))
                 return []

        def get_material_info(material):
             try:
                 name = material.Name if material.Name else "Unnamed"
                 category = material.MaterialCategory if hasattr(material, 'MaterialCategory') and material.MaterialCategory else "Unknown"
                 return name, category
             except Exception:
                 return "Unnamed", "Unknown"

        def _element_uses_material(element, material_int):
             try:
                 if hasattr(element, 'GetMaterialIds'):
                     for include_paint in (True, False):
                         try:
                             for mid in element.GetMaterialIds(include_paint):
                                 if mid and dqt_compat.eid_int(mid) == material_int:
                                     return True
                         except Exception:
                             pass
                 try:
                     type_id = element.GetTypeId()
                     if type_id and dqt_compat.eid_int(type_id) > 0:
                         etype = self.doc.GetElement(type_id)
                         if etype is not None:
                             if hasattr(etype, 'GetMaterialIds'):
                                 try:
                                     for mid in etype.GetMaterialIds(False):
                                         if mid and dqt_compat.eid_int(mid) == material_int:
                                             return True
                                 except Exception:
                                     pass
                             if hasattr(etype, 'GetCompoundStructure'):
                                 try:
                                     cs = etype.GetCompoundStructure()
                                     if cs is not None:
                                         for layer in cs.GetLayers():
                                             mid = layer.MaterialId
                                             if mid and dqt_compat.eid_int(mid) == material_int:
                                                 return True
                                 except Exception:
                                     pass
                 except Exception:
                     pass
             except Exception:
                 pass
             return False

        def find_elements_by_material(material):
             if not material or not material.IsValidObject:
                 return []
             elements_found = []
             material_int = dqt_compat.eid_int(material.Id)
             try:
                 categories_to_check = [
                     BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Doors, BuiltInCategory.OST_Windows,
                     BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_StructuralColumns, BuiltInCategory.OST_Ceilings,
                     BuiltInCategory.OST_Roofs, BuiltInCategory.OST_StructuralFoundation, BuiltInCategory.OST_GenericModel
                 ]
                 seen_ids = set()
                 for category in categories_to_check:
                     try:
                         collector = FilteredElementCollector(self.doc)
                         elements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements()
                         for element in elements:
                             try:
                                 eid = dqt_compat.eid_int(element.Id)
                                 if eid in seen_ids: continue
                                 if _element_uses_material(element, material_int):
                                     elements_found.append(element)
                                     seen_ids.add(eid)
                             except Exception:
                                 continue
                     except Exception:
                         continue
                 return elements_found
             except Exception as ex:
                 print("Error finding elements: {}".format(ex))
                 return []

        choice = forms.CommandSwitchWindow.show(
             ['Create Material Report', 'Find Elements by Material'],
             message='Select function:'
        )
        if not choice:
             return

        if choice == 'Create Material Report':
             materials = get_all_materials()
             if not materials:
                 forms.alert("No materials found")
                 return
             from collections import defaultdict
             category_count = defaultdict(int)
             for material in materials:
                 name, category = get_material_info(material)
                 category_count[category] += 1

             out = script.get_output()
             out.print_md("# **MATERIAL STATISTICS**")
             out.print_md("---")
             out.print_md("**Total Materials:** {}".format(len(materials)))
             out.print_md("**Project:** {}".format(self.doc.Title))
             out.print_md("## **STATISTICS BY CATEGORY**")
             for category, count in sorted(category_count.items()):
                 out.print_md("- **{}:** {}".format(category, count))

             out.print_md("## **MATERIAL LIST**")
             out.print_md("---")
             materials_by_category = defaultdict(list)
             for material in materials:
                 name, category = get_material_info(material)
                 materials_by_category[category].append((name, material))

             for category in sorted(materials_by_category.keys()):
                 out.print_md("### **{}**".format(category))
                 materials_list = sorted(materials_by_category[category], key=lambda x: x[0])
                 for idx, (name, material) in enumerate(materials_list, 1):
                     out.print_md("{}. **{}**".format(idx, name))
                 out.print_md("---")
             forms.alert("Report completed! Check output window.")

        elif choice == 'Find Elements by Material':
             materials = get_all_materials()
             if not materials:
                 forms.alert("No materials found in project!")
                 return
             class MaterialChoice(object):
                 def __init__(self, material):
                     self.material = material
                     self.name, self.category = get_material_info(material)
                     self.display_name = "{} ({})".format(self.name, self.category)
             material_choices = [MaterialChoice(mat) for mat in materials]
             material_choices.sort(key=lambda x: x.display_name)

             selected_choice = forms.SelectFromList.show(
                 material_choices,
                 title="Select Material to Find Elements",
                 button_name='Find Elements',
                 name_attr='display_name'
             )
             if not selected_choice: return

             selected_material = selected_choice.material
             material_name = selected_choice.name
             elements = find_elements_by_material(selected_material)

             if not elements:
                 forms.alert("No elements found using material '{}'".format(material_name))
                 return

             element_choices = []
             for i, element in enumerate(elements):
                 try:
                     try: elem_name = element.Name if element.Name else "Unnamed"
                     except Exception: elem_name = "Unnamed"
                     elem_type = element.GetType().Name
                     elem_category = element.Category.Name if element.Category else "Unknown"
                     display_text = "{}. {} - {} - {} [ID: {}]".format(
                         i + 1, elem_name, elem_type, elem_category, dqt_compat.eid_int(element.Id))
                     element_choices.append(display_text)
                 except Exception:
                     element_choices.append("{}. Unknown Element".format(i + 1))

             forms.SelectFromList.show(
                 element_choices,
                 title="Elements using '{}' (Found: {})".format(material_name, len(elements)),
                 button_name='Close'
             )

             result = forms.alert(
                 "Found {} elements using material '{}'.\n\nSelect these elements in Revit?".format(
                     len(elements), material_name),
                 yes=True, no=True
             )
             if result:
                 id_list = List[ElementId]()
                 for element in elements:
                     id_list.Add(element.Id)
                 self.uidoc.Selection.SetElementIds(id_list)
                 forms.alert("Selected {} elements in Revit!".format(len(elements)))

    # =========================================================================
    # WINDOW CHROME
    # =========================================================================
    def _minimize(self, sender, e):
        self.WindowState = WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()


def show_dialog():
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)
    window = ManaSelectWindow()
    window.ShowDialog()
