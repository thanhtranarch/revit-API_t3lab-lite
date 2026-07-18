# -*- coding: utf-8 -*-
"""
SpecialistSpec registry — per-specialist prompt, tool subset and budget.

The multi-agent layer is NOT separate processes: a specialist is just a
parameterization of the existing native agent loop (script.py's
_run_native_agent) — smaller focused prompt + smaller tool catalog +
tighter iteration budget. That is precisely what makes small local
models answer better.

Tool-subset policy:
    revit_data    read-only subset for ALL providers (queries never
                  need writes; a misroute degrades to "I can't modify").
    revit_action  curated subset for LOCAL providers only — cloud keeps
                  the full registry so no capability is lost there.
    general       provider default (exactly today's behavior).

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

# ─── Tool subsets (validated against tool_schema.ESSENTIAL_TOOL_NAMES) ───────

READ_TOOLS = frozenset([
    "get_revit_context", "revit_get_project_info", "revit_get_active_view",
    "get_current_view_info", "get_current_view_elements",
    "revit_get_selected_elements", "revit_get_element_info",
    "get_elements_by_level", "list_levels", "list_worksets",
    "revit_list_views", "revit_list_sheets", "ai_element_filter",
    "get_parameter", "get_all_parameters", "get_model_warnings",
    "get_model_health", "analyze_model_statistics", "get_schedule_data",
    "get_available_family_types",
    "list_open_documents", "switch_active_document",
])

MODIFY_TOOLS = frozenset([
    "set_parameter", "bulk_set_parameter", "rename_element",
    "select_elements", "color_elements", "revit_override_color",
    "set_active_view", "create_text_note", "tag_elements",
    "tag_all_rooms", "tag_all_walls", "move_elements", "delete_element",
    "create_sheet", "add_view_to_sheet", "duplicate_view",
])

EXPORT_TOOLS = frozenset([
    "export_sheets_pdf", "export_dwg", "export_image",
])

ACTION_TOOLS = READ_TOOLS | MODIFY_TOOLS | EXPORT_TOOLS


class SpecialistSpec(object):

    def __init__(self, name, description, tool_names=None,
                 restrict_cloud=False, prompt_intro='', few_shot='',
                 max_iterations=10, max_tokens=1500, allows_writes=True,
                 use_launcher=True):
        self.name = name
        self.description = description
        self.tool_names = tool_names          # None = provider default
        self.restrict_cloud = restrict_cloud  # subset applies to cloud too
        self.prompt_intro = prompt_intro
        self.few_shot = few_shot              # text block, local models only
        self.max_iterations = max_iterations
        self.max_tokens = max_tokens
        self.allows_writes = allows_writes
        self.use_launcher = use_launcher

    def tools_for(self, local):
        """The tool-name subset to apply, or None for provider default."""
        if self.tool_names is None:
            return None
        if not local and not self.restrict_cloud:
            return None
        return self.tool_names


SPECIALISTS = {
    "general": SpecialistSpec(
        "general",
        "Everything else — full capability, today's exact behavior.",
    ),

    "revit_data": SpecialistSpec(
        "revit_data",
        "Read-only model questions: counts, lists, schedules, health.",
        tool_names=READ_TOOLS,
        restrict_cloud=True,
        allows_writes=False,
        use_launcher=False,
        max_iterations=6,
        max_tokens=1200,
        prompt_intro=(
            "## Role: DATA specialist (read-only)\n"
            "This request is a QUERY about the model. Answer it with read "
            "tools only — you have no modify tools in this turn. Get the "
            "numbers, then answer compactly (counts, tables). If the user "
            "actually wants a modification, say so and stop."
        ),
        few_shot=(
            "## Examples\n"
            "User: \"có bao nhiêu bức tường?\" -> call ai_element_filter "
            "(category Walls), reply \"Model có 128 bức tường.\"\n"
            "User: \"liệt kê các sheet A\" -> call revit_list_sheets, reply "
            "with a compact table of matching numbers + names."
        ),
    ),

    "revit_action": SpecialistSpec(
        "revit_action",
        "Model modifications: rename, set parameters, tag, move, export.",
        tool_names=ACTION_TOOLS,
        restrict_cloud=False,     # cloud keeps the full registry
        max_iterations=10,
        max_tokens=1500,
        prompt_intro=(
            "## Role: ACTION specialist\n"
            "This request modifies the model. Work in this order: (1) read "
            "tools to resolve the exact target element ids, (2) apply the "
            "change, (3) report WHAT changed — counts + element ids. Never "
            "guess ids; if the target is ambiguous, ask ONE clarifying "
            "question instead of acting."
        ),
        few_shot=(
            "## Examples\n"
            "User: \"đổi tên sheet A-101 thành A-102\" -> revit_list_sheets "
            "to find the id -> rename_element -> reply \"Đã đổi tên sheet "
            "A-101 → A-102 (id 12345).\""
        ),
    ),

    # Pipelines — no AgentLoop tools; listed for dispatcher completeness.
    "knowledge": SpecialistSpec(
        "knowledge",
        "Questions answered from indexed documents (RAG with citations).",
        tool_names=frozenset(),
        use_launcher=False,
        allows_writes=False,
    ),
    "comment": SpecialistSpec(
        "comment",
        "PDF markup-comment resolution workflow.",
        tool_names=frozenset(),
        use_launcher=False,
    ),
}


def get_spec(name):
    """Spec by name; unknown names return the general spec."""
    return SPECIALISTS.get(name) or SPECIALISTS["general"]


def build_specialist_prompt(spec, revit_context, project_instructions='',
                            skills_block='', local=False):
    """Base agent prompt + specialist role + (local-only) few-shot +
    project instructions + active-skill block."""
    from Intelligence.agent_loop import build_agent_system_prompt
    parts = [build_agent_system_prompt(revit_context)]
    if spec is not None and spec.prompt_intro:
        parts.append(spec.prompt_intro)
    if local and spec is not None and spec.few_shot:
        parts.append(spec.few_shot)
    if project_instructions:
        parts.append("## Project instructions\n" + project_instructions)
    if skills_block:
        parts.append(skills_block)
    return "\n\n".join(parts)
