# -*- coding: utf-8 -*-
"""
Spell-check service for Revit Text Notes.

Collects every TextNote in the model (or the active view only), de-duplicates
repeated texts, batches them into LLM-sized chunks and builds a focused
proofreading prompt. Pure logic — no WPF here; the assistant window drives
the LLM calls and renders the report.

Author: Tran Tien Thanh
"""

import re


# ─── Collection ───────────────────────────────────────────────────────────────

def collect_text_notes(doc, view_only=False):
    """Return [{id, text, view}] for every non-empty TextNote.

    Notes with no alphabetic content (pure numbers / symbols / dims) are
    skipped — there is nothing to proofread in them.
    """
    # Imported here (not module level) so the pure-text helpers below stay
    # testable outside Revit.
    from Autodesk.Revit import DB
    from Snippets._compat import eid_value

    notes = []
    try:
        if view_only and doc.ActiveView is not None:
            collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
        else:
            collector = DB.FilteredElementCollector(doc)
        collector = collector.OfClass(DB.TextNote).WhereElementIsNotElementType()
        for tn in collector:
            try:
                text = (tn.Text or u"").strip()
                if not text:
                    continue
                if not any(c.isalpha() for c in text):
                    continue
                view_name = u""
                try:
                    ov = doc.GetElement(tn.OwnerViewId)
                    if ov is not None:
                        view_name = ov.Name
                except Exception:
                    pass
                notes.append({
                    "id":   eid_value(tn.Id),
                    "text": re.sub(r'[\r\n]+', u' / ', text),
                    "view": view_name,
                })
            except Exception:
                continue
    except Exception:
        pass
    return notes


def dedupe_notes(notes):
    """Group identical texts → [{text, ids, views}], original order kept.

    Drawings repeat the same annotation dozens of times; proofreading each
    copy separately wastes tokens and floods the report.
    """
    seen, order = {}, []
    for n in notes:
        key = n["text"].strip().lower()
        if key not in seen:
            seen[key] = {"text": n["text"], "ids": [], "views": []}
            order.append(key)
        seen[key]["ids"].append(n["id"])
        if n["view"] and n["view"] not in seen[key]["views"]:
            seen[key]["views"].append(n["view"])
    return [seen[k] for k in order]


def build_batches(unique_notes, max_notes=25, max_chars=3500):
    """Split unique notes into chunks a small local model can handle."""
    batches, cur, cur_chars = [], [], 0
    for n in unique_notes:
        n_len = len(n["text"]) + 12
        if cur and (len(cur) >= max_notes or cur_chars + n_len > max_chars):
            batches.append(cur)
            cur, cur_chars = [], 0
        cur.append(n)
        cur_chars += n_len
    if cur:
        batches.append(cur)
    return batches


# ─── Prompts ──────────────────────────────────────────────────────────────────

def build_system_prompt(viet):
    """Focused proofreader instruction — plain text output, strict format."""
    p = (
        u"You are a meticulous proofreader for architectural construction-"
        u"drawing annotations. The user sends a numbered list of Revit Text "
        u"Notes. Find ENGLISH spelling mistakes and clear grammar errors ONLY.\n"
        u"\n"
        u"NOT errors — never flag these:\n"
        u"- Construction abbreviations/codes: FFL, SFL, SSL, RC, CONC, GRC, "
        u"ALUM, GALV, M&E, TYP, UNO, EQ, CL, DIA, THK, NTS, RWP, FFH, DN, UP, "
        u"W/, C/W, U/S, O/A, dims like 100THK, level/door/sheet codes.\n"
        u"- ALL-CAPS text (normal on drawings) and capitalisation style.\n"
        u"- Proper nouns, project names, product/material brand names.\n"
        u"- Style preferences or synonyms. Report REAL mistakes only.\n"
        u"\n"
        u"Output format — one line per faulty note, nothing else, no preamble:\n"
        u'#<n>: "<wrong fragment>" -> "<correction>"\n'
        u'Several errors in one note: separate with "; " on the same line.\n'
        u"If no note has any error, reply exactly: NO_ERRORS"
    )
    if viet:
        p += (u"\nGiải thích (nếu cần) viết bằng tiếng Việt, nhưng giữ nguyên "
              u"phần văn bản gốc và phần sửa bằng tiếng Anh.")
    return p


def build_batch_query(batch):
    """Numbered plain-text list the model proofreads."""
    lines = []
    for i, n in enumerate(batch):
        lines.append(u"#{}: {}".format(i + 1, n["text"]))
    return u"\n".join(lines)


# ─── Response parsing / report ────────────────────────────────────────────────

_FINDING_RE = re.compile(r'^[-*\s]*#?\s*(\d+)\s*[:.\)\-]\s*(.+)$')


def parse_findings(resp, batch):
    """Map '#<n>: ...' lines back to the batch notes → [{note, issue}]."""
    findings = []
    if not resp:
        return findings
    for line in resp.splitlines():
        line = line.strip()
        if not line or u"NO_ERRORS" in line.upper():
            continue
        m = _FINDING_RE.match(line)
        if not m:
            continue
        idx = int(m.group(1)) - 1
        if idx < 0 or idx >= len(batch):
            continue
        issue = m.group(2).strip()
        if issue:
            findings.append({"note": batch[idx], "issue": issue})
    return findings


def _short(text, limit=90):
    return text if len(text) <= limit else text[:limit - 1] + u"…"


def format_report(findings, total, uniq, failed_batches, viet, view_only):
    """Assemble the final markdown report shown in the chat bubble."""
    lines = []
    if viet:
        scope = u" trong view hiện tại" if view_only else u" trong toàn bộ dự án"
        lines.append(u"🔍 **Kết quả kiểm tra chính tả Text Note**")
        lines.append(u"")
        lines.append(u"Đã quét **{}** Text Note ({} nội dung khác nhau){}.".format(
            total, uniq, scope))
    else:
        scope = u" in the active view" if view_only else u" in the whole project"
        lines.append(u"🔍 **Text Note spell-check results**")
        lines.append(u"")
        lines.append(u"Scanned **{}** Text Notes ({} unique texts){}.".format(
            total, uniq, scope))
    lines.append(u"")

    if not findings:
        lines.append(u"✅ Không phát hiện lỗi chính tả tiếng Anh nào." if viet
                     else u"✅ No English spelling errors found.")
    else:
        lines.append(u"**{} lỗi phát hiện:**".format(len(findings)) if viet
                     else u"**{} issue(s) found:**".format(len(findings)))
        for i, f in enumerate(findings):
            n = f["note"]
            ids = n["ids"]
            id_txt = u", ".join(u"{}".format(x) for x in ids[:4])
            if len(ids) > 4:
                id_txt += (u" +{} vị trí khác".format(len(ids) - 4) if viet
                           else u" +{} more".format(len(ids) - 4))
            lines.append(u"")
            lines.append(u"{}. {}".format(i + 1, f["issue"]))
            lines.append(u'   - "{}"'.format(_short(n["text"])))
            lines.append(u"   - ID: {}".format(id_txt))

    if failed_batches:
        lines.append(u"")
        lines.append(u"⚠️ {} nhóm không kiểm tra được (AI không phản hồi) — thử lại sau.".format(failed_batches)
                     if viet else
                     u"⚠️ {} batch(es) could not be checked (no AI response) — try again.".format(failed_batches))
    return u"\n".join(lines)
