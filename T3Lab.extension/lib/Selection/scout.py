# -*- coding: utf-8 -*-
"""
Background Scout
----------------
Collects current Revit context for the AI Agent to avoid redundant questions.

Author: Tran Tien Thanh
"""

from pyrevit import revit, DB


def _eid_value(element_id):
    """Return the integer value of an ElementId, version-safe.

    Revit 2024+ replaced ElementId.IntegerValue with ElementId.Value (Int64).
    Falls back to IntegerValue for older Revit versions.
    """
    try:
        return element_id.Value          # Revit 2024+
    except AttributeError:
        return element_id.IntegerValue   # Revit 2023 and earlier


class ContextScout:
    """Specialized module for rapid context gathering from the active Revit session."""

    @staticmethod
    def get_active_context():
        """Returns a dictionary containing the current state of the Revit document."""
        doc = revit.doc
        if not doc:
            return {"error": "No active document"}

        # 1. Project Information
        proj_info = doc.ProjectInformation
        
        # 2. View Context
        active_view = doc.ActiveView
        scale_val = "1/{}".format(active_view.Scale) if (active_view and hasattr(active_view, "Scale")) else "Unknown"
        discipline_val = str(active_view.Discipline) if (active_view and hasattr(active_view, "Discipline")) else "Unknown"
        
        # 3. Selection Context
        uidoc = revit.uidoc
        selection_ids = []
        selection_details = []
        if uidoc:
            sel_ids = uidoc.Selection.GetElementIds()
            selection_ids = [_eid_value(e) for e in sel_ids]
            for eid in sel_ids:
                if len(selection_details) >= 5:
                    break
                elem = doc.GetElement(eid)
                if elem:
                    cat_name = elem.Category.Name if elem.Category else "Unknown"
                    elem_name = elem.Name if hasattr(elem, "Name") else str(elem)
                    selection_details.append({
                        "id": _eid_value(eid),
                        "name": elem_name,
                        "category": cat_name
                    })
        
        # Heuristic for Region
        address = (proj_info.Address or "").lower()
        title = (doc.Title or "").lower()
        region = "Unknown"
        if any(kw in address or kw in title for kw in ["singapore", "sgp", "jurong", "changi"]):
            region = "Singapore"
        elif any(kw in address or kw in title for kw in ["vietnam", "vn", "hà nội", "hcm", "việt nam"]):
            region = "Vietnam"

        context = {
            "project": {
                "title": doc.Title,
                "name": proj_info.Name,
                "number": proj_info.Number,
                "region": region
            },
            "active_view": {
                "name": active_view.Name if active_view else "None",
                "type": str(active_view.ViewType) if active_view else "None",
                "id": _eid_value(active_view.Id) if active_view else 0,
                "scale": scale_val,
                "discipline": discipline_val
            },
            "selection": {
                "count": len(selection_ids),
                "ids": selection_ids[:50],  # Cap at 50 IDs to avoid massive JSON
                "details": selection_details
            },
            "revit": {
                "version": doc.Application.VersionNumber,
                "language": str(doc.Application.Language)
            }
        }
        
        return context

    @staticmethod
    def get_context_summary_for_ai():
        """Returns a concise string summary for inclusion in AI prompts."""
        ctx = ContextScout.get_active_context()
        if "error" in ctx: return "No Revit document is currently open."
        
        summary = (
            "Current Context:\n"
            "- Project: {title} ({region})\n"
            "- Active View: {view_name} ({view_type}, Scale: {scale}, Discipline: {discipline})\n"
            "- Selected Elements: {sel_count} items\n"
        ).format(
            title=ctx["project"]["title"],
            region=ctx["project"]["region"],
            view_name=ctx["active_view"]["name"],
            view_type=ctx["active_view"]["type"],
            scale=ctx["active_view"]["scale"],
            discipline=ctx["active_view"]["discipline"],
            sel_count=ctx["selection"]["count"]
        )
        
        if ctx["selection"]["details"]:
            summary += "Selected items details:\n"
            for d in ctx["selection"]["details"]:
                summary += "  * {} ({}) [ID: {}]\n".format(d["name"], d["category"], d["id"])
                
        return summary
