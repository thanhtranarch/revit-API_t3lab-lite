# -*- coding: utf-8 -*-
"""Select All Linked Elements

Select every element of a chosen Revit Link in the active view.

Workflow:
    1. Pick a Revit Link.
    2. Tool builds Reference objects for every model element in that link
       and sets them as the active selection in the active view.
    3. User can then: right-click > Unhide in View > Elements (or Tab,
       Filter, Isolate, etc.)

Author: Dang Quoc Truong - DQT
Copyright by Dang Quoc Truong - DQT (c) 2026
"""

__title__ = 'Select\nLinked'
__author__ = 'Dang Quoc Truong - DQT'
__doc__ = 'Select every element of a chosen Revit Link in the active view.'

# ----------------------------------------------------------------------------
# Imports
# ----------------------------------------------------------------------------
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    Reference,
    ViewType,
)

from pyrevit import forms
from pyrevit import script

from System.Collections.Generic import List

# ----------------------------------------------------------------------------
# Globals
# ----------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView
output = script.get_output()


# ----------------------------------------------------------------------------
# Step 1: Validate active view
# ----------------------------------------------------------------------------
def is_view_supported(view):
    if view is None or view.IsTemplate:
        return False
    unsupported = {
        ViewType.Schedule,
        ViewType.ProjectBrowser,
        ViewType.SystemBrowser,
        ViewType.Internal,
        ViewType.Undefined,
    }
    return view.ViewType not in unsupported


if not is_view_supported(active_view):
    forms.alert(
        "The active view does not support element selection.\n\n"
        "Please open a Plan, Section, Elevation, 3D, Drafting or Sheet view.",
        title="Select All Linked",
        exitscript=True,
    )


# ----------------------------------------------------------------------------
# Step 2: Pick a link
# ----------------------------------------------------------------------------
all_links = list(
    FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
)

if not all_links:
    forms.alert("No Revit Links found in this project.",
                title="Select All Linked", exitscript=True)


# Build label-to-instance map (string-based to avoid TemplateListItem issues)
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
    script.exit()

link_inst = label_to_link[chosen_label]
link_doc = link_inst.GetLinkDocument()
link_title = link_doc.Title if link_doc else link_inst.Name

if link_doc is None:
    forms.alert(
        "The selected link is not loaded. Please reload it first "
        "(Manage > Manage Links) and try again.",
        title="Select All Linked", exitscript=True,
    )


# ----------------------------------------------------------------------------
# Step 3: Build References for every model element in the linked document
# ----------------------------------------------------------------------------
view = active_view

refs = List[Reference]()
total_seen = 0
fail_count = 0

for el in FilteredElementCollector(link_doc).WhereElementIsNotElementType():
    if el.Category is None:
        continue
    try:
        if el.ViewSpecific:
            continue
    except Exception:
        pass
    total_seen += 1
    try:
        r = Reference(el).CreateLinkReference(link_inst)
        if r is not None:
            refs.Add(r)
    except Exception:
        # Some element types (constraints, internal markers, ...) cannot
        # produce a host-side Reference. Skip them silently.
        fail_count += 1


if refs.Count == 0:
    forms.alert(
        "Could not build any selectable references for the {} linked "
        "elements found.".format(total_seen),
        title="Select All Linked", exitscript=True,
    )


# ----------------------------------------------------------------------------
# Step 4: Apply selection
# ----------------------------------------------------------------------------
try:
    uidoc.Selection.SetReferences(refs)
except Exception as ex:
    forms.alert(
        "Failed to apply selection: {}".format(ex),
        title="Select All Linked", exitscript=True,
    )

try:
    uidoc.RefreshActiveView()
except Exception:
    pass


# ----------------------------------------------------------------------------
# Step 5: Report
# ----------------------------------------------------------------------------
lines = []
lines.append("Done.")
lines.append("")
lines.append("Link:     {}".format(link_title))
lines.append("View:     {}".format(view.Name))
lines.append("Selected: {} linked elements".format(refs.Count))
if fail_count > 0:
    lines.append("Skipped:  {} (no selectable Reference)".format(fail_count))
lines.append("")
lines.append("The selection is now active in the view. You can:")
lines.append("  - Right-click > Unhide in View > Elements")
lines.append("  - Right-click > Hide in View > Elements")
lines.append("  - Filter / Tab / Isolate / etc.")

forms.alert("\n".join(lines), title="Select All Linked")

output.print_md("## Select All Linked - Done")
output.print_md("**Link:** `{}`".format(link_title))
output.print_md("**View:** `{}`".format(view.Name))
output.print_md("**Selected:** {} linked elements".format(refs.Count))
if fail_count > 0:
    output.print_md("**Skipped:** {}".format(fail_count))
output.print_md("")
output.print_md("---")
output.print_md("**Dang Quoc Truong - DQT (c) 2026**")