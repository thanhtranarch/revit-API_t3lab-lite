# -*- coding: utf-8 -*-
"""
T3Lab Agent Registry

Defines all available intents (T3Lab tools + MCP Revit commands) and builds
the system prompt that tells AI models exactly which APIs are available.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "T3Lab Agent"

# ─── Intent catalogue ──────────────────────────────────────────────────────────
# Format: intent_name → (category, description, example_phrase)
# Categories: Export | Tools | Revit Query | Revit Create | Revit Modify | Chat

AVAILABLE_INTENTS = {
    # ── T3Lab Export ──────────────────────────────────────────────────────────
    "export_direct":            ("Export",  "Export sheets to PDF/DWG/DWF directly",    u"xuất pdf G sheet"),
    "open_batchout_configured": ("Export",  "Open BatchOut with pre-set config",        u"mở batchout G sheet pdf"),
    "open_batchout":            ("Export",  "Open BatchOut tool",                       u"mở batchout"),

    # ── T3Lab Tools ───────────────────────────────────────────────────────────
    "open_parasync":        ("Tools", "Open ParaSync — sync parameters",           u"mở parasync"),
    "open_loadfamily":      ("Tools", "Open Load Family tool",                     u"load family"),
    "open_loadfamily_cloud":("Tools", "Open Load Family Cloud",                    u"load family cloud"),
    "open_projectname":     ("Tools", "Open Project Name manager",                 u"project name"),
    "open_workset":         ("Tools", "Open Workset manager",                      u"mở workset"),
    "open_dimtext":         ("Tools", "Edit dimension text override",               u"dimension text"),
    "open_upperdimtext":    ("Tools", "Edit upper dimension text",                  u"upper dim text"),
    "open_resetoverrides":  ("Tools", "Reset all graphic overrides",                u"reset overrides"),
    "open_grids":           ("Tools", "Open Grid tool",                             u"mở grids"),

    # ── MCP Revit Query (requires revit-mcp server) ───────────────────────────
    "revit_get_view_info":    ("Revit Query", "Get current active view info",          u"thông tin view hiện tại"),
    "revit_get_elements":     ("Revit Query", "List all elements in current view",     u"danh sách element trong view"),
    "revit_get_selected":     ("Revit Query", "Get info on currently selected elements",u"element đang chọn"),
    "revit_filter_category":  ("Revit Query", "Filter elements by category (OST_Walls etc)", u"lọc tất cả tường"),
    "revit_list_families":    ("Revit Query", "List available family types for a category", u"family types của cửa"),
    "revit_material_qty":     ("Revit Query", "Calculate material quantities",         u"khối lượng vật liệu tường"),
    "revit_model_stats":      ("Revit Query", "Get overall model statistics",          u"thống kê toàn bộ model"),
    "revit_export_rooms":     ("Revit Query", "Export all rooms with area data",        u"xuất danh sách phòng"),
    "revit_query_data":       ("Revit Query", "Query stored project data (extensible storage)", u"truy vấn dữ liệu project"),

    # ── MCP Revit Create (requires revit-mcp server) ──────────────────────────
    "revit_create_wall":      ("Revit Create", "Create a wall along a line",           u"tạo tường/wall mới"),
    "revit_create_floor":     ("Revit Create", "Create a floor element",               u"tạo sàn/floor mới"),
    "revit_create_roof":      ("Revit Create", "Create a roof element",                u"tạo mái/roof"),
    "revit_create_room":      ("Revit Create", "Create a room at a point",             u"tạo phòng/room"),
    "revit_create_grid":      ("Revit Create", "Create a grid line",                   u"tạo lưới/grid"),
    "revit_create_level":     ("Revit Create", "Create a new level",                   u"tạo tầng/level"),
    "revit_place_door":       ("Revit Create", "Place a door in a wall",               u"đặt cửa/door"),
    "revit_place_furniture":  ("Revit Create", "Place a furniture family instance",    u"đặt nội thất"),
    "revit_create_beam":      ("Revit Create", "Create a structural beam",             u"tạo dầm/beam"),

    # ── MCP Revit Modify (requires revit-mcp server) ──────────────────────────
    "revit_select_category":  ("Revit Modify", "Select all elements of a category",   u"chọn tất cả tường"),
    "revit_override_color":   ("Revit Modify", "Override element color in view",       u"đổi màu tường thành xanh"),
    "revit_highlight":        ("Revit Modify", "Highlight elements with color",        u"highlight elements đỏ"),
    "revit_hide_elements":    ("Revit Modify", "Hide selected elements in view",       u"ẩn element trong view"),
    "revit_isolate":          ("Revit Modify", "Isolate selected elements in view",    u"isolate element"),
    "revit_reset_overrides":  ("Revit Modify", "Reset all graphic overrides in view",  u"reset graphic overrides"),
    "revit_delete_elements":  ("Revit Modify", "Delete selected Revit elements",       u"xoá element đã chọn"),
    "revit_tag_walls":        ("Revit Modify", "Auto-tag all walls in current view",   u"tag tất cả tường"),
    "revit_tag_rooms":        ("Revit Modify", "Auto-tag all rooms in current view",   u"tag tất cả phòng"),
    "revit_run_csharp":       ("Revit Modify", "Execute a C# code snippet in Revit",  u"chạy code C#"),
    "revit_store_data":       ("Revit Modify", "Store custom data in project (extensible storage)", u"lưu dữ liệu vào project"),
    "revit_get_stored_data":  ("Revit Modify", "Retrieve stored project data",         u"lấy dữ liệu đã lưu"),

    # ── Conversation ─────────────────────────────────────────────────────────
    "help":    ("Chat", "Answer a T3Lab help question",         u"batchout là gì?"),
    "greet":   ("Chat", "Reply to a greeting",                  u"hello / xin chào"),
    "chat":    ("Chat", "General conversation",                 u"câu hỏi chung"),
    "unknown": ("Chat", "Cannot understand — ask for clarification", u""),
}

# ─── Agent descriptors (for settings UI) ─────────────────────────────────────

AGENTS = [
    {
        "id":          "export",
        "name":        "Export Agent",
        "icon":        u"↗",
        "description": "Handles PDF, DWG, DWF exports and BatchOut operations",
        "intents":     ["export_direct", "open_batchout", "open_batchout_configured"],
    },
    {
        "id":          "tools",
        "name":        "Tools Agent",
        "icon":        u"⚙",
        "description": "Opens T3Lab tools: ParaSync, Load Family, Workset, etc.",
        "intents":     ["open_parasync", "open_loadfamily", "open_loadfamily_cloud",
                        "open_projectname", "open_workset", "open_dimtext",
                        "open_upperdimtext", "open_resetoverrides", "open_grids"],
    },
    {
        "id":          "revit_query",
        "name":        "Revit Query Agent",
        "icon":        u"?",
        "description": "Queries Revit model data — views, elements, families, stats",
        "intents":     ["revit_get_view_info", "revit_get_elements", "revit_get_selected",
                        "revit_filter_category", "revit_list_families",
                        "revit_material_qty", "revit_model_stats",
                        "revit_export_rooms", "revit_query_data"],
    },
    {
        "id":          "revit_create",
        "name":        "Revit Create Agent",
        "icon":        u"➕",
        "description": "Creates Revit elements — walls, floors, rooms, grids, etc.",
        "intents":     ["revit_create_wall", "revit_create_floor", "revit_create_roof",
                        "revit_create_room", "revit_create_grid", "revit_create_level",
                        "revit_place_door", "revit_place_furniture", "revit_create_beam"],
    },
    {
        "id":          "revit_modify",
        "name":        "Revit Modify Agent",
        "icon":        u"✏",
        "description": "Modifies elements — color, hide, isolate, tag, delete",
        "intents":     ["revit_select_category", "revit_override_color", "revit_highlight",
                        "revit_hide_elements", "revit_isolate", "revit_reset_overrides",
                        "revit_delete_elements", "revit_tag_walls", "revit_tag_rooms",
                        "revit_run_csharp", "revit_store_data", "revit_get_stored_data"],
    },
    {
        "id":          "chat",
        "name":        "Chat Agent",
        "icon":        u"...",
        "description": "General conversation, help, and T3Lab knowledge base",
        "intents":     ["help", "greet", "chat"],
    },
]

# ─── MCP-enabled intents (require revit-mcp server running) ──────────────────
MCP_INTENTS = frozenset(
    k for k, v in AVAILABLE_INTENTS.items()
    if v[0].startswith("Revit ")
)


# ─── System prompt builder ────────────────────────────────────────────────────

def build_system_prompt(revit_context=u""):
    """Return the comprehensive system prompt listing all available APIs.

    revit_context: optional Revit model summary injected into the prompt.
    """
    # Group intents by category
    by_cat = {}
    for intent, (cat, desc, ex) in AVAILABLE_INTENTS.items():
        by_cat.setdefault(cat, []).append((intent, desc, ex))

    lines = []
    for cat in ["Export", "Tools", "Revit Query", "Revit Create", "Revit Modify", "Chat"]:
        entries = by_cat.get(cat, [])
        if not entries:
            continue
        lines.append(u"\n[{}]".format(cat.upper()))
        for intent, desc, ex in entries:
            ex_txt = u"  e.g. \"{}\"".format(ex) if ex else u""
            lines.append(u"  {:36s} # {}{}".format(intent, desc, ex_txt))

    revit_block = u""
    if revit_context:
        revit_block = u"\nRevit Model Context:\n{}\n".format(revit_context)

    prompt = (
        u"You are T3Lab Assistant — an AI integrated into Autodesk Revit via T3Lab pyRevit extension.\n"
        u"You can understand natural language and map it to the available APIs below.\n"
        + revit_block +
        u"\nAVAILABLE APIs (use these as 'intent' values):\n"
        + u"\n".join(lines) +
        u"\n\nOUTPUT FORMAT — JSON only, no other text:\n"
        u'{"intent": "<intent>", "params": {<optional>}, "message": "<short reply same language>"}\n'
        u"\nPARAMS:\n"
        u"  export_direct / open_batchout_configured: format (pdf|dwg|dwf|dgn|ifc|nwd|img), filter (sheet prefix e.g. G/A/S or '' for all), combine (bool)\n"
        u"  revit_filter_category: category (OST_Walls | OST_Doors | OST_Rooms | ...)\n"
        u"  revit_override_color: color (hex or CSS name), category (optional filter)\n"
        u"  revit_create_wall: length, level, wallType (optional)\n"
        u"  revit_create_room / revit_place_*: level, position (optional)\n"
        u"  revit_run_csharp: code (C# snippet string)\n"
        u"  revit_store_data / revit_get_stored_data: key, value (optional)\n"
        u"\nRULES:\n"
        u"  - CRITICAL: Never confuse queries about Revit elements (e.g. columns/cột, walls/tường, doors/cửa, rooms/phòng) with sheet exporting (export_direct). If the user asks about elements (e.g. \"có bao nhiêu cột\"), use revit_model_stats, revit_filter_category, or ask for clarification. Do NOT use export_direct.\n"
        u"  - 'xuất/export' + format → export_direct\n"
        u"  - 'mở batchout' + config → open_batchout_configured\n"
        u"  - 'mở batchout' alone → open_batchout\n"
        u"  - Revit element queries/filters → revit_get_* / revit_filter_*\n"
        u"  - Create Revit elements → revit_create_* / revit_place_*\n"
        u"  - Modify, color, hide, tag → revit_override_* / revit_hide_* / revit_tag_*\n"
        u"  - General conversation → chat\n"
        u"  - SAME language as user (Vietnamese → Vietnamese, English → English)\n"
        u"  - If unsure → unknown\n"
        u"\nEXAMPLES:\n"
        u"  input: xuất pdf G sheet\n"
        u'  output: {"intent":"export_direct","params":{"format":"pdf","filter":"G","combine":false},"message":"Đang xuất G sheet sang PDF..."}\n'
        u"\n"
        u"  input: list all walls in view\n"
        u'  output: {"intent":"revit_filter_category","params":{"category":"OST_Walls"},"message":"Fetching all walls in current view..."}\n'
        u"\n"
        u"  input: create a new wall\n"
        u'  output: {"intent":"revit_create_wall","params":{},"message":"Creating a new wall element..."}\n'
    )
    return prompt


def get_intent_info(intent):
    """Return (category, description) for an intent or (None, None)."""
    entry = AVAILABLE_INTENTS.get(intent)
    if entry:
        return entry[0], entry[1]
    return None, None


def is_mcp_intent(intent):
    """Return True if the intent requires the revit-mcp server."""
    return intent in MCP_INTENTS


def get_apis_text():
    """Return a plain-text list of all available APIs for display in settings."""
    lines = []
    current_cat = None
    for intent, (cat, desc, ex) in sorted(AVAILABLE_INTENTS.items(), key=lambda x: x[1][0]):
        if cat != current_cat:
            lines.append(u"\n── {} ──".format(cat.upper()))
            current_cat = cat
        lines.append(u"  {}  —  {}".format(intent, desc))
    return u"\n".join(lines).strip()
