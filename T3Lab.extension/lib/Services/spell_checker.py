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


# ─── Deterministic detection (dictionary engine — no LLM) ─────────────────────

def check_deterministic(unique_notes):
    """Proofread deduped notes with the bundled dictionary engine.

    This is the primary detector: it does NOT depend on any AI provider, so it
    catches obvious typos (CEMTITIOUS, ACYLIC, WATERPROFFING...) that small
    local models silently pass as NO_ERRORS. Returns findings in the same
    shape the LLM path produces, plus the structured ``items`` each finding is
    built from so the caller can aggregate a table:

        [{"note": <unique note>,
          "issue": u'"CEMTITIOUS" -> "CEMENTITIOUS"',
          "items": [{"wrong", "right", "kind", "reason"}, ...]}]

    Returns None when the dictionary asset is unavailable, so the caller can
    fall back to the LLM path instead of reporting a falsely clean model.
    """
    try:
        from Services import spell_dictionary as SD
    except ImportError:
        import spell_dictionary as SD          # test / non-package import
    if not SD.is_available():
        return None
    out = []
    for n in unique_notes:
        items = SD.check_text(n["text"])
        if items:
            out.append({"note": n, "issue": SD.format_issue(items),
                        "items": items})
    return out


def aggregate_findings(findings):
    """Group per-note findings by the (wrong -> right) pair for a compact
    table: one row per distinct correction, with every element id / view it
    occurs in. Mirrors how Claude presents the scan (# · Wrong · Correct ·
    Qty · Element IDs · View)."""
    rows, order = {}, []
    for f in findings:
        note = f["note"]
        # a finding always carries structured items from check_deterministic;
        # the LLM path (issue string only) degrades to a single opaque row.
        items = f.get("items") or [{"wrong": f.get("issue", u""),
                                    "right": u"", "kind": "spelling",
                                    "reason": u""}]
        for it in items:
            key = (it["wrong"].lower(), it["right"].lower())
            if key not in rows:
                rows[key] = {"wrong": it["wrong"], "right": it["right"],
                             "kind": it.get("kind", "spelling"),
                             "reason": it.get("reason", u""),
                             "ids": [], "views": []}
                order.append(key)
            row = rows[key]
            for _id in note.get("ids", []):
                if _id is not None and _id not in row["ids"]:
                    row["ids"].append(_id)
            for _v in note.get("views", []):
                if _v and _v not in row["views"]:
                    row["views"].append(_v)
    return [rows[k] for k in order]


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


def format_report(findings, total, uniq, failed_batches, viet, view_only):
    """Assemble the final markdown report shown in the chat bubble."""
    lines = []
    if viet:
        scope = (u" trong view hiện tại" if view_only
                 else u" trong toàn bộ dự án (Text Note, title block, thông tin "
                      u"dự án, revision, tên sheet/view/room/level, chữ kích "
                      u"thước, model text, tên schedule)")
        lines.append(u"🔍 **Kết quả kiểm tra chính tả tiếng Anh**")
        lines.append(u"")
        lines.append(u"Đã quét **{}** mục text ({} nội dung khác nhau){}.".format(
            total, uniq, scope))
    else:
        scope = (u" in the active view" if view_only
                 else u" in the whole project (Text Notes, title blocks, "
                      u"Project Information, revisions, sheet/view/room/level "
                      u"names, dimension text, model text, schedule names)")
        lines.append(u"🔍 **English spell-check results**")
        lines.append(u"")
        lines.append(u"Scanned **{}** text items ({} unique texts){}.".format(
            total, uniq, scope))
    lines.append(u"")

    if not findings:
        lines.append(u"✅ Không phát hiện lỗi chính tả tiếng Anh nào." if viet
                     else u"✅ No English spelling errors found.")
    else:
        rows = aggregate_findings(findings)
        lines.append(u"Tìm được **{} lỗi** ({} vị trí):".format(
            len(rows), sum(len(r["ids"]) for r in rows)) if viet else
            u"Found **{} issue(s)** ({} instances):".format(
                len(rows), sum(len(r["ids"]) for r in rows)))
        lines.append(u"")
        lines.append(u"| # | Sai | Đúng | SL | Element IDs | View |" if viet
                     else u"| # | Wrong | Correct | Qty | Element IDs | View |")
        lines.append(u"|---|------|------|----|-------------|------|")
        for i, r in enumerate(rows):
            ids = r["ids"]
            id_txt = u", ".join(u"{}".format(x) for x in ids[:6])
            if len(ids) > 6:
                id_txt += u" +{}".format(len(ids) - 6)
            view_txt = u" / ".join(r["views"][:3]) if r["views"] else u"—"
            if len(r["views"]) > 3:
                view_txt += u" …"
            right = r["right"]
            if r.get("reason"):
                right = u"{} ({})".format(right, r["reason"]) if right \
                    else r["reason"]
            lines.append(u"| {} | {} | {} | {} | {} | {} |".format(
                i + 1, r["wrong"], right, len(ids), id_txt, view_txt))

    if failed_batches:
        lines.append(u"")
        lines.append(u"⚠️ {} nhóm không kiểm tra được (AI không phản hồi) — thử lại sau.".format(failed_batches)
                     if viet else
                     u"⚠️ {} batch(es) could not be checked (no AI response) — try again.".format(failed_batches))
    return u"\n".join(lines)
