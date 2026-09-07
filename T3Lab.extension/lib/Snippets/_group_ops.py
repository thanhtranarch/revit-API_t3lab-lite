# -*- coding: utf-8 -*-
"""
_group_ops.py
=============
Revit API helpers for managing **groups** — model groups, detail groups and
attached detail groups:

* collect every group type with its instances, members and worksets
* rename group types in batch (find/replace, prefix, suffix, case, cleanup)
* move group instances (and optionally their members) to another workset
* audit and clean up: unused types, name problems, ungroup, purge

Kept free of any WPF/UI reference so it can be reused by other tools.

Part of T3Lab Extension.
"""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Element,
    ElementId,
    FilteredElementCollector,
    FilteredWorksetCollector,
    Group,
    GroupType,
    Transaction,
    WorksetKind,
)

import System


# ── KINDS ────────────────────────────────────────────────────────────────────

KIND_MODEL = "Model"
KIND_DETAIL = "Detail"
KIND_ATTACHED = "Attached"

KINDS = (KIND_MODEL, KIND_DETAIL, KIND_ATTACHED)

# Revit rejects these in any element name.
ILLEGAL_CHARS = "\\:{}[]|;<>?`~"

CASE_KEEP = "Keep as is"
CASE_UPPER = "UPPERCASE"
CASE_LOWER = "lowercase"
CASE_TITLE = "Title Case"

CASE_MODES = (CASE_KEEP, CASE_UPPER, CASE_LOWER, CASE_TITLE)


# ── VERSION SHIMS ────────────────────────────────────────────────────────────

def eid_int(element_id):
    """Integer value of an ElementId, valid on Revit 2023 - 2026+."""
    if element_id is None:
        return -1
    try:
        return element_id.Value          # Revit 2024+
    except AttributeError:
        return element_id.IntegerValue   # Revit 2023


def new_element_id(value):
    """Build an ElementId from a plain int.

    A bare Python int makes the ElementId constructor ambiguous under pythonnet
    (BuiltInParameter / BuiltInCategory / Int64 all match), so the value is
    widened explicitly: Int64 on Revit 2024+, Int32 on Revit 2023.
    """
    try:
        return ElementId(System.Int64(int(value)))
    except Exception:
        pass
    try:
        return ElementId(System.Int32(int(value)))
    except Exception:
        return ElementId(int(value))


def workset_int(workset_id):
    """Integer value of a WorksetId across Revit versions."""
    if workset_id is None:
        return -1
    try:
        return workset_id.IntegerValue
    except AttributeError:
        pass
    try:
        return workset_id.Value
    except AttributeError:
        return -1


def element_name(element):
    """Best-effort name of an element, never raising."""
    if element is None:
        return ""
    try:
        name = Element.Name.__get__(element)
        if name:
            return name
    except Exception:
        pass
    try:
        return element.Name or ""
    except Exception:
        return ""


# ── NAME RULES ───────────────────────────────────────────────────────────────

def illegal_chars_in(name):
    """The forbidden characters present in name, in order, without repeats."""
    found = []
    for char in name or "":
        if char in ILLEGAL_CHARS and char not in found:
            found.append(char)
    return found


def clean_name(name):
    """Strip forbidden characters, collapse repeated spaces, trim the ends."""
    text = name or ""
    for char in ILLEGAL_CHARS:
        text = text.replace(char, "")
    while "  " in text:
        text = text.replace("  ", " ")
    return text.strip()


def apply_case(name, mode):
    """Apply one of the CASE_MODES to name."""
    if not name:
        return name
    if mode == CASE_UPPER:
        return name.upper()
    if mode == CASE_LOWER:
        return name.lower()
    if mode == CASE_TITLE:
        return " ".join(w[:1].upper() + w[1:].lower() if w else w
                        for w in name.split(" "))
    return name


def build_new_name(current, find="", replace="", match_case=False,
                   prefix="", suffix="", case_mode=CASE_KEEP, cleanup=False):
    """Run the rename rules over one name and return the proposed result."""
    text = current or ""

    if find:
        if match_case:
            text = text.replace(find, replace or "")
        else:
            lowered = text.lower()
            needle = find.lower()
            out = []
            i = 0
            while True:
                hit = lowered.find(needle, i)
                if hit < 0:
                    out.append(text[i:])
                    break
                out.append(text[i:hit])
                out.append(replace or "")
                i = hit + len(needle)
            text = "".join(out)

    if prefix:
        text = prefix + text
    if suffix:
        text = text + suffix

    text = apply_case(text, case_mode)

    if cleanup:
        text = clean_name(text)
    return text


def name_problems(name):
    """Human-readable problems with a group name — empty list when it is fine."""
    issues = []
    text = name or ""
    if not text.strip():
        issues.append("Empty name")
        return issues
    bad = illegal_chars_in(text)
    if bad:
        issues.append("Illegal characters %s" % " ".join(bad))
    if text != text.strip():
        issues.append("Leading/trailing spaces")
    if "  " in text:
        issues.append("Double spaces")
    return issues


# ── RECORDS ──────────────────────────────────────────────────────────────────

class GroupTypeRecord(object):
    """One group type in the document, with everything the UI needs to show."""

    def __init__(self, group_type, kind, instances, doc):
        self.group_type = group_type
        self.type_id = eid_int(group_type.Id)
        self.name = element_name(group_type)
        self.kind = kind
        self.instances = list(instances)
        self.instance_count = len(self.instances)
        self.member_count = 0
        self.workset_names = []
        self.attached_count = 0
        self._read_details(doc)

    def _read_details(self, doc):
        first = self.instances[0] if self.instances else None
        if first is not None:
            try:
                self.member_count = len(list(first.GetMemberIds()))
            except Exception:
                self.member_count = 0
            if self.kind == KIND_MODEL:
                try:
                    self.attached_count = len(list(
                        first.GetAvailableAttachedDetailGroupTypeIds()))
                except Exception:
                    self.attached_count = 0

        seen = []
        for inst in self.instances:
            ws_name = instance_workset_name(doc, inst)
            if ws_name and ws_name not in seen:
                seen.append(ws_name)
        self.workset_names = seen

    @property
    def workset_summary(self):
        if not self.workset_names:
            return "—"
        if len(self.workset_names) == 1:
            return self.workset_names[0]
        return "%d worksets" % len(self.workset_names)

    def audit_issues(self):
        """Cleanup findings for this group type, as a list of short phrases."""
        issues = list(name_problems(self.name))
        if self.instance_count == 0:
            issues.append("Unused")
        elif self.instance_count == 1:
            issues.append("Single instance")
        if len(self.workset_names) > 1:
            issues.append("Mixed worksets")
        if self.instance_count and not self.member_count:
            issues.append("No members")
        return issues


# ── COLLECTION ───────────────────────────────────────────────────────────────

def _type_ids_of(doc, built_in_category):
    """Ids of the group types belonging to one group category."""
    ids = set()
    try:
        collector = FilteredElementCollector(doc) \
            .OfCategory(built_in_category) \
            .WhereElementIsElementType()
        for element_id in collector.ToElementIds():
            ids.add(eid_int(element_id))
    except Exception:
        pass
    return ids


def instance_workset_name(doc, element):
    """Name of the workset an element sits on, or an empty string."""
    if doc is None or element is None or not doc.IsWorkshared:
        return ""
    try:
        table = doc.GetWorksetTable()
        workset = table.GetWorkset(element.WorksetId)
        return workset.Name if workset else ""
    except Exception:
        return ""


def collect_group_types(doc):
    """Every group type in doc as GroupTypeRecord, sorted by kind then name."""
    if doc is None:
        return []

    model_ids = _type_ids_of(doc, BuiltInCategory.OST_IOSModelGroups)
    detail_ids = _type_ids_of(doc, BuiltInCategory.OST_IOSDetailGroups)
    attached_ids = _type_ids_of(doc, BuiltInCategory.OST_IOSAttachedDetailGroups)

    instances_by_type = {}
    try:
        for group in FilteredElementCollector(doc) \
                .OfClass(Group) \
                .WhereElementIsNotElementType() \
                .ToElements():
            try:
                key = eid_int(group.GroupType.Id)
            except Exception:
                continue
            instances_by_type.setdefault(key, []).append(group)
    except Exception:
        pass

    records = []
    for group_type in FilteredElementCollector(doc).OfClass(GroupType).ToElements():
        key = eid_int(group_type.Id)
        if key in detail_ids:
            kind = KIND_DETAIL
        elif key in attached_ids:
            kind = KIND_ATTACHED
        elif key in model_ids:
            kind = KIND_MODEL
        else:
            continue
        records.append(GroupTypeRecord(
            group_type, kind, instances_by_type.get(key, []), doc))

    order = {KIND_MODEL: 0, KIND_DETAIL: 1, KIND_ATTACHED: 2}
    records.sort(key=lambda r: (order.get(r.kind, 9), r.name.lower()))
    return records


def user_worksets(doc):
    """User worksets of doc as [(id_int, name)], sorted by name."""
    if doc is None or not doc.IsWorkshared:
        return []
    out = []
    try:
        for workset in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset):
            out.append((workset_int(workset.Id), workset.Name))
    except Exception:
        return []
    out.sort(key=lambda pair: pair[1].lower())
    return out


# ── RENAME ───────────────────────────────────────────────────────────────────

def rename_group_types(doc, pairs, progress=None):
    """Rename group types in one transaction.

    pairs      — [(GroupTypeRecord, new_name)]
    progress   — optional callable(index, total, label)
    Returns    — [(record, ok, message)]
    """
    results = []
    if doc is None or not pairs:
        return results

    total = len(pairs)
    transaction = Transaction(doc, "T3Lab — Rename Groups")
    transaction.Start()
    try:
        for index, (record, new_name) in enumerate(pairs, 1):
            if progress:
                progress(index, total, record.name)

            wanted = (new_name or "").strip()
            if not wanted:
                results.append((record, False, "Empty name"))
                continue
            if wanted == record.name:
                results.append((record, False, "Unchanged"))
                continue
            bad = illegal_chars_in(wanted)
            if bad:
                results.append((record, False, "Illegal %s" % " ".join(bad)))
                continue

            try:
                record.group_type.Name = wanted
                record.name = wanted
                results.append((record, True, "Renamed"))
            except Exception as exc:
                results.append((record, False, _short_error(exc, "Rename failed")))
        transaction.Commit()
    except Exception:
        if transaction.HasStarted() and not transaction.HasEnded():
            transaction.RollBack()
        raise
    return results


# ── WORKSET ──────────────────────────────────────────────────────────────────

def _set_workset(element, workset_value):
    """Move one element to a workset. Returns (ok, message)."""
    try:
        param = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
    except Exception:
        param = None
    if param is None:
        return False, "No workset parameter"
    if param.IsReadOnly:
        return False, "Workset is read-only"
    try:
        if param.AsInteger() == workset_value:
            return True, "Already there"
    except Exception:
        pass
    try:
        param.Set(System.Int32(int(workset_value)))
        return True, "Moved"
    except Exception as exc:
        return False, _short_error(exc, "Set failed")


def apply_workset(doc, records, workset_value, include_members=False, progress=None):
    """Move every instance of the given group types onto one workset.

    records         — [GroupTypeRecord]
    workset_value   — integer id of the target workset
    include_members — also move the elements inside each group instance
    progress        — optional callable(index, total, label)
    Returns         — [(record, moved, skipped, failed, message)]
    """
    results = []
    if doc is None or not records:
        return results

    total = len(records)
    transaction = Transaction(doc, "T3Lab — Set Group Workset")
    transaction.Start()
    try:
        for index, record in enumerate(records, 1):
            if progress:
                progress(index, total, record.name)

            moved = skipped = failed = 0
            message = ""

            if not record.instances:
                results.append((record, 0, 0, 0, "No instance"))
                continue

            targets = []
            for instance in record.instances:
                targets.append(instance)
                if include_members:
                    try:
                        for member_id in instance.GetMemberIds():
                            member = doc.GetElement(member_id)
                            if member is not None:
                                targets.append(member)
                    except Exception:
                        pass

            for element in targets:
                ok, note = _set_workset(element, workset_value)
                if ok and note == "Moved":
                    moved += 1
                elif ok:
                    skipped += 1
                else:
                    failed += 1
                    message = message or note

            if failed:
                message = "%d failed — %s" % (failed, message)
            elif moved:
                message = "Moved %d element%s" % (moved, "" if moved == 1 else "s")
            else:
                message = "Already there"
            results.append((record, moved, skipped, failed, message))
        transaction.Commit()
    except Exception:
        if transaction.HasStarted() and not transaction.HasEnded():
            transaction.RollBack()
        raise
    return results


# ── CLEANUP ──────────────────────────────────────────────────────────────────

def purge_group_types(doc, records, progress=None):
    """Delete the given group types. Only unused types delete cleanly.

    Returns [(record, ok, message)].
    """
    results = []
    if doc is None or not records:
        return results

    total = len(records)
    transaction = Transaction(doc, "T3Lab — Purge Group Types")
    transaction.Start()
    try:
        for index, record in enumerate(records, 1):
            if progress:
                progress(index, total, record.name)
            if record.instance_count:
                results.append((record, False, "Still placed"))
                continue
            try:
                doc.Delete(record.group_type.Id)
                results.append((record, True, "Purged"))
            except Exception as exc:
                results.append((record, False, _short_error(exc, "Delete failed")))
        transaction.Commit()
    except Exception:
        if transaction.HasStarted() and not transaction.HasEnded():
            transaction.RollBack()
        raise
    return results


def ungroup_instances(doc, records, progress=None):
    """Ungroup every instance of the given group types.

    Returns [(record, ungrouped, failed, message)].
    """
    results = []
    if doc is None or not records:
        return results

    total = len(records)
    transaction = Transaction(doc, "T3Lab — Ungroup Groups")
    transaction.Start()
    try:
        for index, record in enumerate(records, 1):
            if progress:
                progress(index, total, record.name)

            if not record.instances:
                results.append((record, 0, 0, "No instance"))
                continue

            ungrouped = failed = 0
            message = ""
            for instance in record.instances:
                try:
                    instance.UngroupMembers()
                    ungrouped += 1
                except Exception as exc:
                    failed += 1
                    message = message or _short_error(exc, "Ungroup failed")
            if failed:
                message = "%d failed — %s" % (failed, message)
            else:
                message = "Ungrouped %d" % ungrouped
            results.append((record, ungrouped, failed, message))
        transaction.Commit()
    except Exception:
        if transaction.HasStarted() and not transaction.HasEnded():
            transaction.RollBack()
        raise
    return results


# ── HELPERS ──────────────────────────────────────────────────────────────────

def _short_error(exc, fallback):
    """First line of an exception message, trimmed for a status cell."""
    try:
        text = str(exc).strip().splitlines()[0]
    except Exception:
        text = ""
    text = text.replace("Autodesk.Revit.Exceptions.", "")
    return text[:60] if text else fallback


def duplicate_names(records):
    """Names shared by more than one group type, compared case-insensitively."""
    counts = {}
    for record in records:
        key = (record.name or "").strip().lower()
        counts[key] = counts.get(key, 0) + 1
    return set(key for key, count in counts.items() if count > 1 and key)
