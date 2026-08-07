# -*- coding: utf-8 -*-
"""
opus_teacher — DISTILLATION enricher: Opus 5 (teacher) writes gold answers into
the self-study SFT corpus so a LOCAL Qwen (student) can be fine-tuned toward
Opus-level quality on T3Lab / Revit tasks.

Unlike every other enricher, this one is ALLOWED to spend Claude API tokens: it
calls the Claude provider DIRECTLY (not the router's active provider), so it can
distil from Opus even while the assistant itself runs on local Qwen. Because it
costs money it is doubly gated — the study engine only schedules it when the
host reports a teacher is available (agents.opus_teacher on + a Claude key), and
it caps how many examples it generates per idle cycle.

Candidate prompts come from two sources, both cheap and OFFLINE to select:
  1) the office's REAL commands (learned_patterns) whose stored answer is a terse
     routing ack — exactly the turns a richer Opus answer improves most;
  2) a curated bilingual seed list of core Revit/BIM questions (mirrors the
     Intelligence/skills/*.md playbooks) so distillation has signal from day one.

Already-taught prompts are skipped by folding the corpus's existing user turns,
so re-running on a later idle cycle costs nothing until there is something new.

Pure selection helpers take plain dicts/lists (CPython-3 unit-testable); run()
wires them to the real stores + Claude provider + dataset. Never raises.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Opus Teacher"

SOURCE  = 'opus_teacher'
QUALITY = 'teacher'

# Distillation is token-bearing, so keep each idle cycle cheap; the round-robin
# schedule comes back to this enricher on later cycles for the rest.
MAX_PER_CYCLE = 3

# A stored learned-pattern answer at or below this many characters is a routing
# ack ("Đang mở BatchOut...") — a real office command that deserves a richer
# teacher answer, not a target to train on as-is.
_TERSE_MAX = 40

# Room for a full expert answer (tables, steps) without runaway token spend.
_ANSWER_TOKENS = 700

# The teacher answers as the assistant SHOULD in chat — concise, practical,
# grounded, same language as the user. Deliberately NOT the JSON-intent system
# prompt (build_system_prompt): we are distilling helpful conversational
# knowledge, not routing labels.
TEACHER_SYSTEM_PROMPT = (
    u"You are T3Lab Assistant, an expert Autodesk Revit and BIM coordination "
    u"assistant for a Vietnamese AEC office. Answer the user's question "
    u"concisely and practically, grounded in real Revit workflows and common "
    u"BIM standards (ISO 19650, LOD, worksets, view templates, schedules/QTO). "
    u"Prefer concrete steps and specifics over generalities. Use plain text (no "
    u"Markdown headings). Reply in the SAME language the user wrote in "
    u"(Vietnamese or English)."
)

# Curated bilingual seed questions — one per core skill playbook. Distillation
# has signal even before the office has generated any learned patterns.
SEED_PROMPTS = [
    u"Giải thích các mức LOD 100–500 trong Revit và khi nào dùng mức nào?",
    u"How do I name sheets and models to ISO 19650? Give the field order.",
    u"Quy trình xử lý warning trong Revit nên ưu tiên loại nào trước?",
    u"What is a good workset strategy for a multi-discipline Revit model?",
    u"Làm sao lập schedule khối lượng (QTO) chính xác cho tường và sàn?",
    u"How should view templates be structured for a coordinated project?",
    u"Shared coordinates vs project base point — khi nào dùng cái nào?",
    u"What is the recommended level and grid setup workflow in Revit?",
    u"Quy trình tạo và quản lý family chuẩn trong dự án gồm những bước nào?",
    u"How do I run clash coordination between architecture and MEP models?",
    u"Chuẩn đặt tên annotation và dimension trong bản vẽ nên theo nguyên tắc gì?",
    u"What should a model cleanup checklist cover before an issue?",
    u"Cách quản lý shared parameters và project parameters khác nhau thế nào?",
    u"What belongs in a BIM Execution Plan (BEP) for a mid-size project?",
    u"Quy trình bàn giao COBie/BIM handover cần chuẩn bị những gì?",
    u"How do I set up sheet sets and a revision workflow in Revit?",
]


# ─── Pure helpers (CPython-3 unit-testable) ──────────────────────────────────────

def _fold(text):
    """Cheap dedup key: lowercase + collapsed whitespace. Never raises."""
    try:
        return u" ".join(u"{}".format(text or u"").split()).lower()
    except Exception:
        return u""


def terse_pattern_prompts(patterns, terse_max=_TERSE_MAX):
    """User phrasings from learned_patterns whose stored answer is terse/empty.

    These are the office's real commands (last_raw) that currently map to a
    short routing ack (last_message) — the turns an Opus answer improves most.
    Returns a list of prompt strings, most-used first.
    """
    scored = []
    for _key, data in (patterns or {}).items():
        if not isinstance(data, dict):
            continue
        user = u"{}".format(data.get('last_raw') or u"").strip()
        if not user:
            continue
        answer = u"{}".format(data.get('last_message') or u"").strip()
        if len(answer) > terse_max:
            continue          # already has a rich stored answer
        scored.append((int(data.get('hits', 0) or 0), user))
    scored.sort(key=lambda t: t[0], reverse=True)
    return [u for _h, u in scored]


def select_candidates(patterns, seen_users, max_n=MAX_PER_CYCLE,
                      seeds=None):
    """Ordered, de-duplicated prompts to distil this cycle, capped at max_n.

    Real office commands (patterns) come first — they are what the office
    actually needs — then the curated seeds fill any remaining budget. Anything
    already present in the corpus (folded user turns in `seen_users`) is skipped
    so a later cycle never re-teaches the same prompt.
    """
    seen = set(seen_users or [])
    out = []
    for prompt in terse_pattern_prompts(patterns) + list(
            seeds if seeds is not None else SEED_PROMPTS):
        key = _fold(prompt)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(prompt)
        if len(out) >= max_n:
            break
    return out


# ─── Orchestration (injectable for tests) ────────────────────────────────────────

def run(chat_fn=None, add_example=None, load_patterns=None, iter_examples=None,
        system_prompt=None, revit_version=None, max_new=MAX_PER_CYCLE,
        **_ignored):
    """Distil up to `max_new` gold answers from the teacher into the SFT corpus.

    chat_fn(user_prompt) -> answer text | None is the ONLY teacher dependency;
    loop.py wires it to the Claude provider pinned to an Opus model. All other
    I/O is injected (defaults wired to the real modules) so tests drive it with
    plain dicts + a fake chat_fn. Returns {status, added}. Never raises.
    """
    try:
        if add_example is None:
            from Intelligence.learning import dataset as _ds
            add_example = _ds.add_example
        if iter_examples is None:
            from Intelligence.learning import dataset as _ds2
            iter_examples = _ds2.iter_examples
        if load_patterns is None:
            from Intelligence import t3lab_assistant as _ta
            load_patterns = _ta.load_learned_patterns
    except Exception as ex:
        return {'status': u'unavailable: {}'.format(ex), 'added': 0}

    if chat_fn is None:
        return {'status': 'no teacher chat_fn', 'added': 0}

    sys_prompt = system_prompt or TEACHER_SYSTEM_PROMPT

    # Fold the user turns already in the corpus so we never re-teach a prompt.
    seen_users = set()
    try:
        for obj in iter_examples():
            for m in (obj.get('messages') or []):
                if isinstance(m, dict) and m.get('role') == 'user':
                    seen_users.add(_fold(m.get('content')))
    except Exception:
        pass

    try:
        patterns = load_patterns() or {}
    except Exception:
        patterns = {}

    candidates = select_candidates(patterns, seen_users, max_n=max_new)
    if not candidates:
        return {'status': 'ok', 'added': 0, 'note': 'nothing new to teach'}

    added = 0
    for prompt in candidates:
        try:
            answer = chat_fn(prompt)
        except Exception as ex:
            return {'status': u'teacher call failed: {}'.format(ex),
                    'added': added}
        answer = u"{}".format(answer or u"").strip()
        if not answer:
            continue          # teacher unavailable/empty — stop wasting the cycle
        try:
            ok, note = add_example(
                [{'role': 'system', 'content': sys_prompt},
                 {'role': 'user', 'content': prompt},
                 {'role': 'assistant', 'content': answer}],
                source=SOURCE, quality=QUALITY, revit_version=revit_version)
            if ok and note == u'added':
                added += 1
        except Exception:
            continue

    return {'status': 'ok', 'added': added}
