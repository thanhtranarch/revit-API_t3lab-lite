# -*- coding: utf-8 -*-
"""
_cpython_bootstrap.py
=====================
Bootstraps Python 3 standard library paths and C-extension paths for CPython
running inside pyRevit. In pyRevit's CPython engine, python312.zip, extracted Lib,
and the engine root (containing .pyd files like _sqlite3.pyd and DLLs like sqlite3.dll)
may not be in sys.path / DLL directory, causing 'No module named configparser',
'No module named csv', or 'No module named _sqlite3'.
This module locates and prepends all required paths to sys.path and DLL loaders.
"""

import os
import sys


def init_cpython_paths():
    """Ensure pyRevit CPython standard library and C extensions are in sys.path and DLL path."""
    sys.dont_write_bytecode = True
    engine_dirs = []
    candidates = []

    # APPDATA & PROGRAMDATA paths
    for env_var in ('APPDATA', 'PROGRAMDATA'):
        base = os.environ.get(env_var, '')
        if base:
            for clone in ('pyRevit-Master', 'pyRevit'):
                ceng = os.path.join(base, clone, 'bin', 'cengines', 'CPY3123')
                if os.path.isdir(ceng):
                    engine_dirs.append(ceng)
                    candidates.append(ceng)                       # .pyd files (_sqlite3.pyd, etc.)
                    candidates.append(os.path.join(ceng, 'Lib'))  # stdlib (.pyc files: csv, json, configparser)
                    candidates.append(os.path.join(ceng, 'python312.zip'))

    # Program Files paths
    for pf in (r'C:\Program Files', r'C:\Program Files (x86)'):
        for clone in ('pyRevit-Master', 'pyRevit'):
            ceng = os.path.join(pf, clone, 'bin', 'cengines', 'CPY3123')
            if os.path.isdir(ceng):
                engine_dirs.append(ceng)
                candidates.append(ceng)
                candidates.append(os.path.join(ceng, 'Lib'))
                candidates.append(os.path.join(ceng, 'python312.zip'))

    # Configure DLL search path for C extensions (e.g. sqlite3.dll, libffi-8.dll, libssl-3.dll)
    for ed in engine_dirs:
        for _dd in (ed, os.path.join(ed, 'Lib')):
            if hasattr(os, 'add_dll_directory') and os.path.isdir(_dd):
                try:
                    os.add_dll_directory(_dd)
                except Exception:
                    pass
        path_env = os.environ.get('PATH', '')
        _lib_d = os.path.join(ed, 'Lib')
        if ed not in path_env:
            os.environ['PATH'] = ed + os.pathsep + _lib_d + os.pathsep + path_env

    # Also make sure this extension's lib directory is in sys.path
    ext_lib = os.path.dirname(__file__)
    if ext_lib not in sys.path:
        sys.path.insert(0, ext_lib)

    # Prepend existing candidates to sys.path
    for c in reversed(candidates):
        if os.path.exists(c) and c not in sys.path:
            sys.path.insert(0, c)

    # Force reload of GUI.WPF_Base if cached, and monkeypatch forms.WPFWindow
    try:
        import importlib
        for _mod_name in ('GUI.WPF_Base', 'WPF_Base', 'GUI.forms', 'pyrevit.forms._cpy'):
            if _mod_name in sys.modules:
                try:
                    importlib.reload(sys.modules[_mod_name])
                except Exception:
                    pass
        from GUI.WPF_Base import T3WPFWindow
        import pyrevit.forms as _pyrevit_forms
        _pyrevit_forms.WPFWindow = T3WPFWindow
        class _CPythonReactive(object):
            def OnPropertyChanged(self, *args, **kwargs):
                pass

        _pyrevit_forms.Reactive = _CPythonReactive
        try:
            import pyrevit.forms._cpy as _cpy
            _cpy.WPFWindow = T3WPFWindow
            _cpy.Reactive = _CPythonReactive
        except Exception:
            pass
    except Exception:
        pass

    # Scripts call init_cpython_paths() directly, so install the CPython API
    # shims here too rather than only at module import.
    for _installer in (fix_std_streams, install_forms_shim, install_script_shim,
                       enable_safe_engine_shutdown):
        try:
            _installer()
        except Exception:
            pass


class _NullStream(object):
    """Stand-in for a stdout/stderr that has no write() (pyRevit ScriptIO)."""

    def write(self, *_a, **_k):
        return 0

    def flush(self, *_a, **_k):
        pass

    def isatty(self):
        return False


def fix_std_streams():
    """Make `print()` safe.

    Under the CPython engine sys.stdout can be a pyRevit ScriptIO with no
    `write`, so a bare print() raises `'ScriptIO' object has no attribute
    'write'`. That most often fires from inside an `except` block, replacing the
    real error with a bogus one.
    """
    for name in ('stdout', 'stderr'):
        try:
            stream = getattr(sys, name, None)
            usable = False
            if stream is not None:
                try:
                    # Probe by writing. hasattr() is not enough: pyRevit's
                    # ScriptIO can answer attribute lookups with an exception
                    # that is not AttributeError, which hasattr re-raises.
                    stream.write('')
                    usable = True
                except Exception:
                    usable = False
            if not usable:
                setattr(sys, name, _NullStream())
        except Exception:
            pass


def _t3_alert(msg, title=None, sub_msg=None, expanded=None, footer='',
              ok=True, cancel=False, yes=False, no=False, retry=False,
              warn_icon=True, options=None, exitscript=False, **kwargs):
    """CPython-safe stand-in for `pyrevit.forms.alert`, backed by GUI.T3Dialog."""
    caption = title or 'T3Lab'
    details = expanded or sub_msg or footer or None
    answer = True
    try:
        from GUI.T3Dialog import show_info, show_warning, confirm
        if yes or no or cancel:
            answer = bool(confirm(msg, title=caption, details=details))
        elif warn_icon:
            show_warning(msg, title=caption, details=details)
        else:
            show_info(msg, title=caption, details=details)
    except Exception:
        # Last resort: Revit's own dialog, so a message still reaches the user.
        try:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show(caption, msg if not details
                            else u"{}\n\n{}".format(msg, details))
        except Exception:
            pass
    if exitscript:
        sys.exit(0)
    return answer


def _t3_pick_file(file_ext='', files_filter='', init_dir='', title=None,
                  multi_file=False, unc_paths=False, **kwargs):
    from Microsoft.Win32 import OpenFileDialog
    dlg = OpenFileDialog()
    dlg.Filter = files_filter or (
        "{0} files|*.{0}|All files|*.*".format(file_ext) if file_ext else "All files|*.*")
    dlg.Multiselect = bool(multi_file)
    if init_dir:
        dlg.InitialDirectory = init_dir
    if title:
        dlg.Title = title
    if dlg.ShowDialog() != True:  # noqa: E712 - Nullable<bool> from WPF
        return None
    return list(dlg.FileNames) if multi_file else dlg.FileName


def _t3_save_file(file_ext='', files_filter='', init_dir='', default_name='',
                  title=None, **kwargs):
    from Microsoft.Win32 import SaveFileDialog
    dlg = SaveFileDialog()
    dlg.Filter = files_filter or (
        "{0} files|*.{0}|All files|*.*".format(file_ext) if file_ext else "All files|*.*")
    if file_ext:
        dlg.DefaultExt = file_ext
    if default_name:
        dlg.FileName = default_name
    if init_dir:
        dlg.InitialDirectory = init_dir
    if title:
        dlg.Title = title
    return dlg.FileName if dlg.ShowDialog() == True else None  # noqa: E712


def _t3_pick_folder(title=None, owner=None, **kwargs):
    try:
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        dlg = FolderBrowserDialog()
        if title:
            dlg.Description = title
        return dlg.SelectedPath if dlg.ShowDialog() == DialogResult.OK else None
    except Exception:
        return None


def _wpf_list_dialog(items, title='Select', multiselect=False,
                     button_name='Select', name_attr=None, prompt=None):
    """Minimal WPF list picker built in code (no XAML dependency).

    Returns the ORIGINAL objects, not their labels, so callers that pass
    dict keys or Revit elements get their objects back.
    """
    from System.Windows import (Window, WindowStartupLocation, Thickness,
                                GridLength, GridUnitType, HorizontalAlignment)
    from System.Windows.Controls import (Grid, RowDefinition, ListBox, Button,
                                         StackPanel, Orientation, SelectionMode,
                                         TextBlock)

    entries = list(items or [])

    def label(obj):
        if name_attr:
            return u'{}'.format(getattr(obj, name_attr, obj))
        return u'{}'.format(obj)

    win = Window()
    win.Title = title or 'Select'
    win.Width = 420
    win.Height = 480
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    grid = Grid()
    grid.Margin = Thickness(12)
    for height in (GridLength(1, GridUnitType.Auto),
                   GridLength(1, GridUnitType.Star),
                   GridLength(1, GridUnitType.Auto)):
        rd = RowDefinition()
        rd.Height = height
        grid.RowDefinitions.Add(rd)

    caption = TextBlock()
    caption.Text = prompt or title or 'Select'
    caption.Margin = Thickness(0, 0, 0, 8)
    Grid.SetRow(caption, 0)
    grid.Children.Add(caption)

    box = ListBox()
    if multiselect:
        box.SelectionMode = SelectionMode.Extended
    for obj in entries:
        box.Items.Add(label(obj))
    Grid.SetRow(box, 1)
    grid.Children.Add(box)

    bar = StackPanel()
    bar.Orientation = Orientation.Horizontal
    bar.Margin = Thickness(0, 8, 0, 0)
    bar.HorizontalAlignment = HorizontalAlignment.Right

    state = {'ok': False}

    def _accept(_s, _e):
        state['ok'] = True
        win.DialogResult = True

    ok_btn = Button()
    ok_btn.Content = button_name or 'Select'
    ok_btn.Width = 96
    ok_btn.IsDefault = True
    ok_btn.Margin = Thickness(0, 0, 8, 0)
    ok_btn.Click += _accept
    bar.Children.Add(ok_btn)

    cancel_btn = Button()
    cancel_btn.Content = 'Cancel'
    cancel_btn.Width = 96
    cancel_btn.IsCancel = True
    bar.Children.Add(cancel_btn)

    Grid.SetRow(bar, 2)
    grid.Children.Add(bar)
    win.Content = grid
    win.ShowDialog()

    if not state['ok']:
        return [] if multiselect else None
    picked = [entries[i] for i in range(len(entries))
              if box.SelectedItems.Contains(label(entries[i]))]
    if multiselect:
        return picked
    return picked[0] if picked else None


class _T3SelectFromList(object):
    """CPython stand-in for `pyrevit.forms.SelectFromList`."""

    @classmethod
    def show(cls, context, title=None, button_name='Select', multiselect=False,
             name_attr=None, **kwargs):
        try:
            return _wpf_list_dialog(context, title=title or 'Select',
                                    multiselect=multiselect,
                                    button_name=button_name, name_attr=name_attr)
        except Exception:
            return [] if multiselect else None


class _T3CommandSwitchWindow(object):
    """CPython stand-in for `pyrevit.forms.CommandSwitchWindow` (option picker)."""

    @classmethod
    def show(cls, context, message=None, title=None, **kwargs):
        try:
            return _wpf_list_dialog(context, title=title or message or 'Select',
                                    prompt=message, button_name='OK')
        except Exception:
            return None


def _t3_ask_for_string(default='', prompt=None, title=None, **kwargs):
    """CPython stand-in for `pyrevit.forms.ask_for_string`."""
    from System.Windows import (Window, WindowStartupLocation, Thickness,
                                GridLength, GridUnitType, HorizontalAlignment,
                                SizeToContent)
    from System.Windows.Controls import (Grid, RowDefinition, TextBox, Button,
                                         StackPanel, Orientation, TextBlock)
    try:
        win = Window()
        win.Title = title or 'Input'
        win.Width = 420
        win.SizeToContent = SizeToContent.Height
        win.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(12)
        for _ in range(3):
            rd = RowDefinition()
            rd.Height = GridLength(1, GridUnitType.Auto)
            grid.RowDefinitions.Add(rd)

        caption = TextBlock()
        caption.Text = prompt or 'Enter a value:'
        caption.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(caption, 0)
        grid.Children.Add(caption)

        entry = TextBox()
        entry.Text = u'{}'.format(default or '')
        entry.SelectAll()
        Grid.SetRow(entry, 1)
        grid.Children.Add(entry)

        bar = StackPanel()
        bar.Orientation = Orientation.Horizontal
        bar.HorizontalAlignment = HorizontalAlignment.Right
        bar.Margin = Thickness(0, 12, 0, 0)
        state = {'ok': False}

        def _accept(_s, _e):
            state['ok'] = True
            win.DialogResult = True

        ok_btn = Button()
        ok_btn.Content = 'OK'
        ok_btn.Width = 96
        ok_btn.IsDefault = True
        ok_btn.Margin = Thickness(0, 0, 8, 0)
        ok_btn.Click += _accept
        bar.Children.Add(ok_btn)

        cancel_btn = Button()
        cancel_btn.Content = 'Cancel'
        cancel_btn.Width = 96
        cancel_btn.IsCancel = True
        bar.Children.Add(cancel_btn)

        Grid.SetRow(bar, 2)
        grid.Children.Add(bar)
        win.Content = grid
        entry.Focus()
        win.ShowDialog()
        return entry.Text if state['ok'] else None
    except Exception:
        return None


class _T3ProgressBar(object):
    """No-op context manager standing in for `pyrevit.forms.ProgressBar`."""

    def __init__(self, *_args, **_kwargs):
        self.cancelled = False
        self.title = _kwargs.get('title', '')

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        return False

    def update_progress(self, *_a, **_k):
        pass

    def reset(self, *_a, **_k):
        pass


class _T3WarningBar(object):
    """No-op context manager standing in for `pyrevit.forms.WarningBar`."""

    def __init__(self, *_args, **_kwargs):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        return False

    def Close(self):
        pass


def install_forms_shim():
    """Fill in the `pyrevit.forms` API that the CPython engine refuses to serve.

    `pyrevit/forms/__init__.py` defines a module `__getattr__` that raises
    PyRevitCPythonNotSupported for EVERY attribute, so `forms.alert` and friends
    are fatal under CPython — ~390 call sites in this extension. Binding real
    attributes on the module means normal lookup succeeds and `__getattr__` is
    never consulted, so existing call sites work unchanged and even
    already-imported (cached) modules are fixed.

    Only fills genuine gaps: anything the engine already provides is left alone.
    """
    try:
        import pyrevit.forms as _forms
    except Exception:
        return
    if getattr(_forms, 'IRONPY', False):
        return          # IronPython has the real implementations

    def _missing(name):
        try:
            getattr(_forms, name)
            return False
        except Exception:
            return True

    for name, impl in (('alert', _t3_alert),
                       ('pick_file', _t3_pick_file),
                       ('save_file', _t3_save_file),
                       ('pick_folder', _t3_pick_folder),
                       ('SelectFromList', _T3SelectFromList),
                       ('CommandSwitchWindow', _T3CommandSwitchWindow),
                       ('ask_for_string', _t3_ask_for_string),
                       ('ProgressBar', _T3ProgressBar),
                       ('WarningBar', _T3WarningBar),
                       ('MessageBox', _t3_alert),
                       ('toast', lambda *a, **k: None)):
        if _missing(name):
            try:
                setattr(_forms, name, impl)
            except Exception:
                pass


class _NullOutput(object):
    """No-op stand-in for pyRevit's output window.

    The CPython engine has no ScriptOutput (`ScriptOutput.GetDefault` is
    missing), so `script.get_output()` raises and takes the tool with it. Every
    method here is a silent no-op: `print_md`, `print_table`, `close_others`,
    `set_width`, `self_destruct` and friends (71 `output.print_md` calls in this
    extension) become harmless instead of fatal.
    """

    def __getattr__(self, _name):
        def _noop(*_args, **_kwargs):
            return None
        return _noop


def enable_safe_engine_shutdown():
    """Make `Reload pyRevit` work again on Revit 2026 / .NET 8.

    Reload calls ScriptEngineManager.ClearEngines() -> PythonEngine.Shutdown()
    -> RuntimeData.Stash(), and Stash serializes with BinaryFormatter, which
    .NET 8 refuses:

        PlatformNotSupportedException: BinaryFormatter serialization and
        deserialization have been removed.

    sessionmgr._clear_running_engines() only catches AttributeError, so this
    escapes and aborts the whole reload — which means edits under lib/ can never
    be picked up without restarting Revit.

    pythonnet ships `NoopFormatter` for exactly this; pointing RuntimeData at it
    makes Stash a no-op. Discarding the stash is what we want anyway: the point
    of a reload is fresh modules, not restored engine state.

    Static on the loaded Python.Runtime assembly, so one CPython script running
    this fixes Reload for the rest of the Revit session.
    """
    if sys.version_info[0] < 3:
        return
    try:
        import clr
        from Python.Runtime import RuntimeData, NoopFormatter
        if RuntimeData.FormatterType is None:
            RuntimeData.FormatterType = clr.GetClrType(NoopFormatter)
    except Exception:
        pass


def safe_output():
    """The pyRevit output window, or a no-op stand-in — never raises.

    Call sites should use this instead of `script.get_output()` directly: the
    bootstrap runs long before a script imports pyrevit, so patching
    `pyrevit.script` from here is a timing gamble. This is not.
    """
    if sys.version_info[0] >= 3:
        return _NullOutput()        # CPython has no ScriptOutput.GetDefault
    try:
        from pyrevit import script
        return script.get_output()
    except Exception:
        return _NullOutput()


def install_script_shim():
    """Patch `pyrevit.script.get_output` when pyrevit is already loaded.

    Belt-and-braces for library modules; the reliable path is `safe_output()`.
    Deliberately does NOT import pyrevit itself — the bootstrap runs at the very
    top of a script, before pyrevit is ready, and forcing the import there is
    both fragile and a side effect we do not want.
    """
    if sys.version_info[0] < 3:
        return                      # IronPython has the real output window
    for mod_name in ('pyrevit.script', 'pyrevit.output'):
        mod = sys.modules.get(mod_name)
        if mod is None:
            continue                # not imported yet - safe_output() covers it
        try:
            mod.get_output = lambda *_a, **_k: _NullOutput()
        except Exception:
            pass


# Run immediately upon import.
# NOTE: this module lives in lib/, so the persistent CPython engine caches it in
# sys.modules — editing it requires one pyRevit Reload before the new code runs.
init_cpython_paths()
fix_std_streams()
install_forms_shim()
install_script_shim()
enable_safe_engine_shutdown()
