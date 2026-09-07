# -*- coding: utf-8 -*-
"""
Tests for Group Manager logic: the Revit-free naming helpers in
Snippets/_group_ops.py (rename rules, cleanup, name validation) and the
duplicate detection the Rename tab relies on.
Run: python dev/test_group_manager.py
"""
import os
import sys
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB_DIR = os.path.join(REPO, 'T3Lab.extension', 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)


def _load_group_ops_module():
    """Exec Snippets/_group_ops.py with the Revit and .NET imports stubbed out.

    The module is only importable inside Revit, so the Revit-free helpers are
    exercised against the real shipped source rather than a copy.
    """
    import types

    class _AnyMeta(type):
        """Any attribute lookup yields another permissive placeholder type."""
        def __getattr__(cls, item):
            return _make_any(item)

    def _make_any(name):
        return _AnyMeta(str(name), (), {})

    def _stub(name):
        mod = types.ModuleType(name)
        mod.__getattr__ = _make_any
        return mod

    saved = {}
    stubs = {
        'Autodesk': _stub('Autodesk'),
        'Autodesk.Revit': _stub('Autodesk.Revit'),
        'Autodesk.Revit.DB': _stub('Autodesk.Revit.DB'),
        'System': _stub('System'),
    }
    stubs['Autodesk'].Revit = stubs['Autodesk.Revit']
    stubs['Autodesk.Revit'].DB = stubs['Autodesk.Revit.DB']
    for name, mod in stubs.items():
        saved[name] = sys.modules.get(name)
        sys.modules[name] = mod

    try:
        source_path = os.path.join(LIB_DIR, 'Snippets', '_group_ops.py')
        with open(source_path, 'r', encoding='utf-8') as handle:
            source = handle.read()
        module = types.ModuleType('t3_group_ops_under_test')
        module.__file__ = source_path
        exec(compile(source, source_path, 'exec'), module.__dict__)
        return module
    finally:
        for name, previous in saved.items():
            if previous is None:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = previous


class _FakeRecord(object):
    def __init__(self, name):
        self.name = name


class TestRenameRules(unittest.TestCase):
    """The rename rules that drive the NEW NAME preview column."""

    @classmethod
    def setUpClass(cls):
        cls.ops = _load_group_ops_module()

    def test_no_rule_keeps_the_name(self):
        self.assertEqual(self.ops.build_new_name("Bathroom Pod"), "Bathroom Pod")

    def test_find_replace_is_case_insensitive_by_default(self):
        self.assertEqual(
            self.ops.build_new_name("BATHROOM Pod", find="bathroom", replace="Bath"),
            "Bath Pod")

    def test_find_replace_honours_match_case(self):
        self.assertEqual(
            self.ops.build_new_name("BATHROOM Pod", find="bathroom", replace="Bath",
                                    match_case=True),
            "BATHROOM Pod")

    def test_find_replace_hits_every_occurrence(self):
        self.assertEqual(
            self.ops.build_new_name("AR-AR-AR", find="AR", replace="ST"),
            "ST-ST-ST")

    def test_find_with_empty_replace_deletes(self):
        self.assertEqual(
            self.ops.build_new_name("OLD_Bathroom", find="OLD_"), "Bathroom")

    def test_prefix_and_suffix(self):
        self.assertEqual(
            self.ops.build_new_name("Pod", prefix="AR-", suffix="-L06"),
            "AR-Pod-L06")

    def test_case_modes(self):
        self.assertEqual(
            self.ops.build_new_name("bath pod", case_mode=self.ops.CASE_UPPER),
            "BATH POD")
        self.assertEqual(
            self.ops.build_new_name("Bath POD", case_mode=self.ops.CASE_LOWER),
            "bath pod")
        self.assertEqual(
            self.ops.build_new_name("bath POD", case_mode=self.ops.CASE_TITLE),
            "Bath Pod")

    def test_cleanup_strips_illegal_characters_and_spaces(self):
        self.assertEqual(
            self.ops.build_new_name("  Bath{Pod}  x  ", cleanup=True), "BathPod x")

    def test_rules_apply_in_order_find_prefix_suffix_case_cleanup(self):
        result = self.ops.build_new_name(
            "old pod", find="old", replace="new", prefix="ar-", suffix=" {x}",
            case_mode=self.ops.CASE_UPPER, cleanup=True)
        self.assertEqual(result, "AR-NEW POD X")


class TestNameValidation(unittest.TestCase):
    """Name problems surfaced in the STATUS and FINDINGS columns."""

    @classmethod
    def setUpClass(cls):
        cls.ops = _load_group_ops_module()

    def test_clean_name_is_idempotent(self):
        once = self.ops.clean_name("  A  B|C  ")
        self.assertEqual(once, "A BC")
        self.assertEqual(self.ops.clean_name(once), once)

    def test_illegal_chars_are_reported_once_each(self):
        self.assertEqual(self.ops.illegal_chars_in("a{b}{c}"), ["{", "}"])

    def test_illegal_chars_empty_for_a_good_name(self):
        self.assertEqual(self.ops.illegal_chars_in("AR-Bathroom Pod 01"), [])

    def test_name_problems_flags_empty(self):
        self.assertEqual(self.ops.name_problems("   "), ["Empty name"])

    def test_name_problems_flags_spaces_and_illegal(self):
        problems = self.ops.name_problems(" Bath|Pod  01 ")
        self.assertTrue(any("Illegal" in p for p in problems))
        self.assertIn("Leading/trailing spaces", problems)
        self.assertIn("Double spaces", problems)

    def test_name_problems_empty_for_a_good_name(self):
        self.assertEqual(self.ops.name_problems("AR-Bathroom Pod 01"), [])


class TestDuplicateDetection(unittest.TestCase):
    """Two group types may not end up sharing a name."""

    @classmethod
    def setUpClass(cls):
        cls.ops = _load_group_ops_module()

    def test_duplicates_ignore_case_and_padding(self):
        records = [_FakeRecord("Bath Pod"), _FakeRecord("bath pod "),
                   _FakeRecord("Kitchen Pod")]
        self.assertEqual(self.ops.duplicate_names(records), {"bath pod"})

    def test_no_duplicates_returns_empty(self):
        records = [_FakeRecord("A"), _FakeRecord("B")]
        self.assertEqual(self.ops.duplicate_names(records), set())


if __name__ == '__main__':
    unittest.main()
