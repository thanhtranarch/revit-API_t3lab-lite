# -*- coding: utf-8 -*-
"""
Tool Registry
-------------
Centralized registry for T3Lab tools accessible by the AI Agent.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

import os

class ToolRegistry:
    def __init__(self, extension_path=None):
        if extension_path is None:
            # Assume we are in lib/core/registry.py
            # Go up 3 levels to reach T3Lab.extension
            self.base_path = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        else:
            self.base_path = extension_path
            
        self.tab_path = os.path.join(self.base_path, "T3Lab.tab")
        self.tools = self._initialize_tools()

    def _initialize_tools(self):
        """Returns a dictionary of available tools with their metadata."""
        return {
            "cad_to_beam": {
                "name": "CAD to Beam",
                "description": "Converts CAD lines to Revit beams with AI dimension detection from nearby TextNotes.",
                "rel_path": "Project.panel/Create.stack/Beam.pushbutton/script.py"
            },
            "align_positions": {
                "name": "Align Positions",
                "description": "Snaps element positions to clean multiples of 5 or 10 mm relative to a reference element (Grid, Wall, or Column).",
                "rel_path": "Project.panel/AlignPositions.pushbutton/script.py"
            },
            "annotation_manager": {
                "name": "Annotation Manager",
                "description": "Unified tool for managing Dimensions and Text Notes — find, delete, and auto-rename types and instances.",
                "rel_path": "Annotation.panel/AnnotationManager.pushbutton/script.py"
            },
            "workset_manager": {
                "name": "Workset Manager",
                "description": "List, rename, and manage user worksets; remove unused worksets via a checklist.",
                "rel_path": "Project.panel/Workset.stack/Workset.pushbutton/script.py"
            },
            "batch_export": {
                "name": "BatchOut",
                "description": "Batch export sheets to PDF, DWG, NWD, and IFC formats with custom naming and revision tracking.",
                "rel_path": "Export.panel/BatchOut.pushbutton/script.py"
            },
            "load_family": {
                "name": "Load Family",
                "description": "Browse and load Revit families from local disk or Cloud (Vercel API).",
                "rel_path": "Project.panel/FamilyWork.stack/LoadFamily.pushbutton/script.py"
            },
            "room_to_area": {
                "name": "Room to Area",
                "description": "Automatically convert room boundaries to area boundaries in the active area plan.",
                "rel_path": "Project.panel/Areas.stack/Room to Area.pushbutton/script.py"
            },
            "create_plan_views": {
                "name": "Create Plan Views",
                "description": "Batch-generate individual floor plan views for each room with custom naming and template assignment.",
                "rel_path": "Project.panel/Create.stack/Create Room Plan.pushbutton/script.py"
            }
        }

    def get_tool(self, tool_id):
        """Returns tool metadata by ID."""
        return self.tools.get(tool_id)

    def get_all_tools(self):
        """Returns all registered tools."""
        return self.tools

    def get_script_path(self, tool_id):
        """Returns the absolute path to a tool's script."""
        tool = self.get_tool(tool_id)
        if tool:
            return os.path.join(self.tab_path, tool["rel_path"].replace("/", os.sep))
        return None

    def list_tools_for_ai(self):
        """Returns a simplified list of tools for LLM consumption."""
        return [
            {"id": tid, "name": t["name"], "description": t["description"]}
            for tid, t in self.tools.items()
        ]
