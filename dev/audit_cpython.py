#! python3
# -*- coding: utf-8 -*-
"""audit_cpython.py — cong tac vien kiem tra migration IronPython -> CPython.

Moi check o day tuong ung MOT loi that da lam chet tool trong Revit, khong phai
suy doan. Chay:  python3 dev/audit_cpython.py [--quiet]

  C1  P0  CLR generic nhan Python builtin        -> TypeError: type(s) expected
  C2  P0  revit.doc / revit.uidoc o module scope -> AttributeError luc import
  C3  P0  pyrevit.forms.<api> chua duoc shim     -> PyRevitCPythonNotSupported
  C4  P0  __namespace__ trong pushbutton script  -> Duplicate type name (click 2)
  C5  P1  DB.<Enum> qua pyrevit.DB               -> attribute missing tren stub
  C6  P1  print() ngoai except                   -> ScriptIO has no attribute write
"""

import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXT = os.path.join(ROOT, 'T3Lab.extension')

# pyrevit.forms API da duoc _cpython_bootstrap shim -> an toan
SHIMMED = {'alert', 'pick_file', 'save_file', 'pick_folder', 'WPFWindow',
           'SelectFromList', 'CommandSwitchWindow', 'ask_for_string',
           'ProgressBar', 'WarningBar', 'MessageBox', 'toast',
           'Reactive', 'IRONPY', 'Visibility', '_cpy'}

GENERIC = re.compile(
    r'\b(List|Array|Dictionary|ObservableCollection|IList|HashSet|IEnumerable|'
    r'Queue|Stack|KeyValuePair)\[\s*(object|str|int|float|bool|bytes|dict|list|tuple)\s*[,\]]')
MODULE_DOC = re.compile(
    r'^(doc|uidoc|app)\s*=\s*(revit\.(doc|uidoc|app)'
    r'|__revit__\.(ActiveUIDocument|Application)'
    r'|__revit__\.ActiveUIDocument\.Document)\s*$')
FORMS_API = re.compile(r'\bforms\.([A-Za-z_][A-Za-z0-9_]*)')
DB_ENUM = re.compile(r'\bDB\.(BuiltInCategory|BuiltInParameter|UnitType|SpecTypeId)\b')
PRINT_CALL = re.compile(r'^\s*print\s*\(')


def iter_py():
    for base, dirs, files in os.walk(EXT):
        dirs[:] = [d for d in dirs if d not in ('_archive', '__pycache__')]
        for fn in files:
            if fn.endswith('.py'):
                yield os.path.join(base, fn)


def rel(path):
    return os.path.relpath(path, ROOT).replace('\\', '/')


def audit():
    findings = []

    def add(sev, code, path, line, msg):
        findings.append((sev, code, rel(path), line, msg))

    for path in iter_py():
        try:
            src = io.open(path, encoding='utf-8').read()
        except Exception:
            continue
        lines = src.split('\n')
        is_pushbutton = path.replace('\\', '/').endswith('.pushbutton/script.py')

        for i, line in enumerate(lines, 1):
            stripped = line.strip()
            if stripped.startswith('#'):
                continue

            m = GENERIC.search(line)
            if m:
                add('P0', 'C1', path, i,
                    '%s[%s] - dung System.Object / CLR type' % (m.group(1), m.group(2)))

            if MODULE_DOC.match(line):
                add('P0', 'C2', path, i,
                    '%s o module scope - boc try/except, dung Snippets._host.resolve_doc()'
                    % stripped)

            for name in FORMS_API.findall(line):
                if name not in SHIMMED:
                    add('P0', 'C3', path, i,
                        'forms.%s chua shim - dung GUI.T3Dialog hoac them vao '
                        '_cpython_bootstrap.install_forms_shim()' % name)

            if is_pushbutton and '__namespace__' in line and '=' in line:
                # An giai neu da boc "if sys.version_info[0] < 3:" ngay tren:
                # CPython khong dat __namespace__ -> pythonnet tu sinh ten duy nhat.
                guarded = any('version_info[0] < 3' in prev
                              for prev in lines[max(0, i - 3):i - 1])
                if not guarded:
                    add('P0', 'C4', path, i,
                        '__namespace__ trong script.py chay lai moi click -> '
                        'Duplicate type name; boc "if sys.version_info[0] < 3:"')

            if 'script.get_output()' in line and '=' in line:
                # An giai neu da di qua _cpython_bootstrap.safe_output()
                if not any('safe_output' in prev for prev in lines[max(0, i - 5):i]):
                    add('P0', 'C7', path, i,
                        'script.get_output() truc tiep -> ScriptOutput.GetDefault '
                        'khong ton tai; dung _cpython_bootstrap.safe_output()')

            m = DB_ENUM.search(line)
            if m:
                add('P1', 'C5', path, i,
                    'DB.%s - import truc tiep tu Autodesk.Revit.DB' % m.group(1))

            if PRINT_CALL.match(line):
                add('P1', 'C6', path, i, 'print() - bo hoac dung logger')

    return findings


def main():
    quiet = '--quiet' in sys.argv
    findings = audit()
    p0 = [f for f in findings if f[0] == 'P0']
    p1 = [f for f in findings if f[0] == 'P1']

    if not quiet:
        for code in ('C1', 'C2', 'C3', 'C4', 'C5', 'C6'):
            rows = [f for f in findings if f[1] == code]
            if not rows:
                continue
            print('\n%s  (%d)' % (code, len(rows)))
            shown = {}
            for sev, _c, path, line, msg in rows:
                shown.setdefault(path, []).append((line, msg))
            for path in sorted(shown):
                hits = shown[path]
                print('  %s' % path)
                for line, msg in hits[:4]:
                    print('      %-5d %s' % (line, msg))
                if len(hits) > 4:
                    print('      ... +%d dong nua' % (len(hits) - 4))

    print('\nCPYTHON AUDIT: %d P0 · %d P1  (%d file quet)'
          % (len(p0), len(p1), len(list(iter_py()))))
    return 1 if p0 else 0


if __name__ == '__main__':
    sys.exit(main())
