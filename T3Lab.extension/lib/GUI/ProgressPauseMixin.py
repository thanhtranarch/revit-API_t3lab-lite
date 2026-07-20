# -*- coding: utf-8 -*-
"""
ProgressPauseMixin.py
Shared progress bar + Pause/Resume + Stop behaviour for T3Lab batch tools.

Single source of truth for the pattern previously copy-pasted in
AutoJoin/script.py and FamiGenDialog.py.

Usage (window class)::

    from GUI.ProgressPauseMixin import ProgressPauseMixin

    class MyToolWindow(forms.WPFWindow, ProgressPauseMixin):
        # Override PP_* class attrs if the XAML uses different x:Name values.

        def run_clicked(self, sender, e):
            items = collect_items()
            self.begin_progress(len(items), disable=[self.btn_run])
            for i, item in enumerate(items):
                if not self.step_progress(i, "Processing {}...".format(item)):
                    break                      # user pressed Stop
                process(item)                  # one item per step
            cancelled = self.is_cancelled      # read BEFORE end_progress()
            self.end_progress()

XAML requirements (canonical snippet: .claude/skills/xaml-templates.md §13)::

    progress_panel  — container Grid, Visibility="Collapsed" when idle
    pb_run          — ProgressBar (Style="{StaticResource T3ProgressBar}")
    btn_pause       — Click="pause_resume_clicked"
    btn_stop        — Click="stop_clicked"
    status_text     — optional TextBlock for status messages

Pause/Stop semantics:
    - The batch runs on the Revit UI thread inside a MODAL window
      (ShowDialog). step_progress()/_update_progress() pump the WPF
      dispatcher (DispatcherFrame) so the window keeps repainting and the
      Pause/Stop buttons stay clickable.
    - Pause blocks inside step_progress() at an item boundary. An already
      open Transaction stays open while paused — prefer per-item/chunk
      transactions (TransactionGroup) for new tools so pause never holds
      a transaction open.
    - Stop is cooperative: the NEXT step_progress() call returns False —
      the current item always finishes ("finishing current item").
    - Do NOT use this mixin in modeless windows (BatchOut-style
      ExternalEvent tools): never pump the dispatcher inside Execute().
      Pause there is done by not re-queueing the next chunk instead.
"""

try:
    import clr
    clr.AddReference('PresentationFramework')
    clr.AddReference('PresentationCore')
    clr.AddReference('WindowsBase')
    clr.AddReference('System')
    from System import Action                                     # type: ignore
    from System.Windows import Visibility                         # type: ignore
    from System.Windows.Threading import (                        # type: ignore
        DispatcherFrame, DispatcherPriority)
except Exception:
    # Outside Revit/IronPython (e.g. dev audit scripts import-scanning
    # this module with CPython) — the UI methods are never called there.
    Action = Visibility = DispatcherFrame = DispatcherPriority = None


class ProgressPauseMixin(object):
    """Progress bar + Pause/Resume + Stop for modal batch tools."""

    # x:Name lookups — override per window when the XAML differs
    PP_PANEL  = "progress_panel"
    PP_BAR    = "pb_run"
    PP_PAUSE  = "btn_pause"
    PP_STOP   = "btn_stop"
    PP_STATUS = "status_text"

    # Labels / messages — override per tool if needed
    PP_PAUSE_LABEL  = u"⏸  Pause"
    PP_RESUME_LABEL = u"▶  Resume"
    PP_STOP_MSG     = u"Stopping… finishing current item"
    PP_PAUSED_MSG   = u"Paused — click Resume to continue"

    # Optional glyph elements: when the XAML names a TextBlock pair inside
    # btn_pause (Segoe MDL2 icon + label), the mixin toggles their Text
    # instead of replacing Button.Content (canonical AutoJoin.xaml pattern).
    PP_PAUSE_ICON   = "btn_pause_icon"
    PP_PAUSE_TEXT   = "btn_pause_label"
    PP_PAUSE_GLYPH  = u"\uE769"   # MDL2 Pause
    PP_RESUME_GLYPH = u"\uE768"   # MDL2 Play
    PP_PAUSE_PLAIN  = u"Pause"
    PP_RESUME_PLAIN = u"Resume"

    # --------------------------------------------------
    # Internals
    # --------------------------------------------------

    def _pp_el(self, name):
        """XAML element lookup by x:Name; None when the window lacks it."""
        return getattr(self, name, None)

    def _pp_ensure_state(self):
        """Lazy flag init so handlers are safe even before begin_progress()."""
        if getattr(self, "_pause_requested", None) is None:
            self._pause_requested = False
        if getattr(self, "_cancel_requested", None) is None:
            self._cancel_requested = False
        if getattr(self, "_pp_disabled", None) is None:
            self._pp_disabled = []

    def _pp_show_paused(self, paused):
        """Reflect pause state on the Pause/Resume button.

        Prefers named icon/label TextBlocks (PP_PAUSE_ICON / PP_PAUSE_TEXT,
        Segoe MDL2 glyph pattern); falls back to swapping Button.Content
        strings for windows without them (FamiGen pattern).
        """
        try:
            icon  = self._pp_el(self.PP_PAUSE_ICON)
            label = self._pp_el(self.PP_PAUSE_TEXT)
            if icon is not None or label is not None:
                if icon is not None:
                    icon.Text = self.PP_RESUME_GLYPH if paused else self.PP_PAUSE_GLYPH
                if label is not None:
                    label.Text = self.PP_RESUME_PLAIN if paused else self.PP_PAUSE_PLAIN
            else:
                btn = self._pp_el(self.PP_PAUSE)
                if btn is not None:
                    btn.Content = self.PP_RESUME_LABEL if paused else self.PP_PAUSE_LABEL
        except Exception:
            pass

    def _pp_set_status(self, text):
        """Write to the window's status area.

        Prefers a window-defined _update_status(text) (FamiGen pattern);
        falls back to setting <PP_STATUS>.Text directly.
        """
        upd = getattr(self, "_update_status", None)
        if callable(upd):
            try:
                upd(text)
                return
            except Exception:
                pass
        st = self._pp_el(self.PP_STATUS)
        if st is not None:
            try:
                st.Text = text
            except Exception:
                pass

    def _do_events(self):
        """Pump the WPF dispatcher so the window repaints and buttons click."""
        try:
            frame = DispatcherFrame()
            def _stop(f=frame):
                f.Continue = False
            self.Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                Action(_stop))
            self.Dispatcher.PushFrame(frame)
        except Exception:
            pass

    # --------------------------------------------------
    # Low-level API (back-compat with pre-mixin call sites)
    # --------------------------------------------------

    def _update_progress(self, value, maximum=None):
        """Set bar value (and Maximum), show the panel, pump events.

        Blocks here while paused; returns immediately after Stop.
        """
        self._pp_ensure_state()
        try:
            bar = self._pp_el(self.PP_BAR)
            if bar is not None:
                if maximum is not None:
                    bar.Maximum = maximum
                bar.Value = value
            panel = self._pp_el(self.PP_PANEL)
            if panel is not None:
                panel.Visibility = Visibility.Visible
        except Exception:
            pass
        self._do_events()
        while self._pause_requested and not self._cancel_requested:
            self._do_events()

    def _hide_progress(self):
        """Hide the panel and reset flags + Pause/Stop button states."""
        self._pp_ensure_state()
        try:
            panel = self._pp_el(self.PP_PANEL)
            if panel is not None:
                panel.Visibility = Visibility.Collapsed
            bar = self._pp_el(self.PP_BAR)
            if bar is not None:
                bar.Value = 0
            self._cancel_requested = False
            self._pause_requested  = False
            self._pp_show_paused(False)
            btn_pause = self._pp_el(self.PP_PAUSE)
            if btn_pause is not None:
                btn_pause.IsEnabled = True
            btn_stop = self._pp_el(self.PP_STOP)
            if btn_stop is not None:
                btn_stop.IsEnabled = True
        except Exception:
            pass

    # --------------------------------------------------
    # High-level API
    # --------------------------------------------------

    @property
    def is_cancelled(self):
        """True once the user pressed Stop. Reset by end_progress()."""
        self._pp_ensure_state()
        return self._cancel_requested

    def begin_progress(self, maximum=100, disable=None):
        """Reset flags, show the panel at 0/<maximum>, disable action controls.

        `disable` is a list of controls (e.g. the Run button) re-enabled
        automatically by end_progress() — the reentrancy guard while the
        dispatcher is being pumped.
        """
        self._pp_ensure_state()
        self._cancel_requested = False
        self._pause_requested  = False
        self._pp_disabled = []
        for ctrl in (disable or []):
            try:
                if ctrl.IsEnabled:
                    ctrl.IsEnabled = False
                    self._pp_disabled.append(ctrl)
            except Exception:
                pass
        self._update_progress(0, maximum)

    def step_progress(self, value, message=None):
        """Per-item update: bar + optional status message.

        Blocks while paused. Returns False once Stop was pressed —
        use as:  if not self.step_progress(i, msg): break
        """
        if message is not None:
            self._pp_set_status(message)
        self._update_progress(value)
        return not self._cancel_requested

    def end_progress(self):
        """Hide the panel and re-enable controls disabled by begin_progress().

        NOTE: resets the cancel flag — read self.is_cancelled BEFORE this.
        """
        self._pp_ensure_state()
        for ctrl in self._pp_disabled:
            try:
                ctrl.IsEnabled = True
            except Exception:
                pass
        self._pp_disabled = []
        self._hide_progress()

    # --------------------------------------------------
    # XAML event handlers
    # --------------------------------------------------

    def stop_clicked(self, sender, e):
        """Click handler for btn_stop — cooperative cancel."""
        self._pp_ensure_state()
        self._cancel_requested = True
        self._pause_requested  = False
        try:
            btn_stop = self._pp_el(self.PP_STOP)
            if btn_stop is not None:
                btn_stop.IsEnabled = False
        except Exception:
            pass
        self._pp_set_status(self.PP_STOP_MSG)

    def pause_resume_clicked(self, sender, e):
        """Click handler for btn_pause — toggles pause/resume."""
        self._pp_ensure_state()
        if self._pause_requested:
            self._pause_requested = False
            self._pp_show_paused(False)
        else:
            self._pause_requested = True
            self._pp_show_paused(True)
            self._pp_set_status(self.PP_PAUSED_MSG)
