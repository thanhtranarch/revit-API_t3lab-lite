# -*- coding: utf-8 -*-
"""
Background Scout
----------------
Collects current Revit context for the AI Agent to avoid redundant questions.

Author: Tran Tien Thanh
"""

from pyrevit import revit, DB

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
        
        # 3. Selection Context
        uidoc = revit.uidoc
        selection_ids = [e.IntegerValue for e in uidoc.Selection.GetElementIds()]
        
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
                "name": active_view.Name,
                "type": str(active_view.ViewType),
                "id": active_view.Id.IntegerValue
            },
            "selection": {
                "count": len(selection_ids),
                "ids": selection_ids[:50]  # Cap at 50 IDs to avoid massive JSON
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
            "- Active View: {view_name} ({view_type})\n"
            "- Selected Elements: {sel_count} items\n"
        ).format(
            title=ctx["project"]["title"],
            region=ctx["project"]["region"],
            view_name=ctx["active_view"]["name"],
            view_type=ctx["active_view"]["type"],
            sel_count=ctx["selection"]["count"]
        )
        return summary
