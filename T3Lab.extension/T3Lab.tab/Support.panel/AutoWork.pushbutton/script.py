# -*- coding: utf-8 -*-
"""
Auto Work

Quick Click: automate repetitive clicks at a fixed screen coordinate.
Record & Replay: record any sequence of mouse actions and replay them
                 with exact timing and positions.

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
--------------------------------------------------------
"""

__title__   = "Auto Work"
__author__  = "Tran Tien Thanh"
__version__ = "2.0.0"

# IMPORT LIBRARIES
# ==================================================
import os
import sys
import clr
import time
import ctypes

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')

from System import Action
from System.Windows import WindowState
from System.Windows.Forms import Cursor
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Threading import DispatcherPriority
from System.Threading import Thread, ThreadStart

from pyrevit import forms, script

extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir       = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

logger = script.get_logger()

XAML_PATH = os.path.join(extension_dir, 'lib', 'GUI', 'Tools', 'AutoWork.xaml')

# Mouse event flags
_LDOWN  = 0x0002
_LUP    = 0x0004
_RDOWN  = 0x0008
_RUP    = 0x0010

# Virtual key codes
_VK_LBUTTON = 0x01
_VK_RBUTTON = 0x02


# ╔══════════════════════════════════════════════════════════════════════╗
# ║  ⚠  IRONPYTHON 2.7 RANGE FIX — DO NOT CHANGE BACK TO range()  ⚠  ║
# ║                                                                      ║
# ║  In IronPython 2.7, range() builds a full list in memory.           ║
# ║  When a value comes from a WPF TextBox (System.String), int()        ║
# ║  may return a Python *long* (BigInteger) even for small numbers.    ║
# ║  range(long) → OverflowError: "too many items in the range".        ║
# ║                                                                      ║
# ║  Rules:                                                              ║
# ║   • Use xrange() for ALL for-loops in this file.                    ║
# ║   • Use while loops for counters derived from TextBox inputs.        ║
# ║   • NEVER replace xrange with range, even after a Python 3 upgrade. ║
# ╚══════════════════════════════════════════════════════════════════════╝

def _flush_keys():
    for vk in xrange(8, 256):
        ctypes.windll.user32.GetAsyncKeyState(vk)


def _any_key_pressed():
    for vk in xrange(8, 256):
        if ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000:
            return True
    return False


def _interruptible_sleep(ms):
    """Sleep for ms milliseconds, returns True if aborted by keypress."""
    elapsed = 0
    while elapsed < ms:
        if _any_key_pressed():
            return True
        Thread.Sleep(20)
        elapsed += 20
    return False


# WINDOW CLASS
# ==================================================
class AutoWorkWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_PATH)
        self._recorded_actions = []   # list of (action_type, x, y, delay_ms)
        self._is_recording = False
        self._is_playing   = False
        self._set_status("Ready")

    # ── Chrome ──────────────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self._is_recording = False
        self._is_playing   = False
        self.Close()

    # ── Status bar ──────────────────────────────────────────────────────────

    def _set_status(self, text, error=False):
        self.status_text.Text = text
        if error:
            self.status_text.Foreground = SolidColorBrush(Color.FromRgb(211, 47, 47))
        else:
            self.status_text.Foreground = SolidColorBrush(Color.FromRgb(127, 140, 141))

    # ── Mode selector ────────────────────────────────────────────────────────

    def mode_changed(self, sender, args):
        if self.rb_quick.IsChecked:
            self.main_tabs.SelectedIndex = 0
        else:
            self.main_tabs.SelectedIndex = 1

    # ══════════════════════════════════════════════════════════════════════
    # QUICK CLICK MODE
    # ══════════════════════════════════════════════════════════════════════

    def pick_location_clicked(self, sender, e):
        self._set_status("Minimizing for 3 seconds — move mouse to target!")
        self.WindowState = WindowState.Minimized
        Thread.Sleep(3000)
        pos = Cursor.Position
        self.txt_x.Text = str(pos.X)
        self.txt_y.Text = str(pos.Y)
        self.WindowState = WindowState.Normal
        self._set_status("Location captured: X={}, Y={}".format(pos.X, pos.Y))

    def start_clicked(self, sender, e):
        try:
            x             = int(self.txt_x.Text)
            y             = int(self.txt_y.Text)
            interval      = float(self.txt_interval.Text)
            total_clicks  = int(self.txt_clicks.Text)
        except ValueError:
            self._set_status("Invalid values — enter numbers only.", error=True)
            return

        if total_clicks <= 0:
            self._set_status("Total clicks must be > 0.", error=True)
            return

        self._set_status("Starting in 2 seconds...")
        self.WindowState = WindowState.Minimized
        Thread.Sleep(2000)
        _flush_keys()

        # ⚠ while loop required — see IRONPYTHON 2.7 RANGE FIX block above.
        interval_ms  = int(round(interval * 1000))
        clicks_done  = 0
        while clicks_done < total_clicks:
            if _any_key_pressed():
                self.WindowState = WindowState.Normal
                self._set_status("Stopped by keypress after {} click(s).".format(clicks_done), error=True)
                return

            ctypes.windll.user32.SetCursorPos(x, y)
            ctypes.windll.user32.mouse_event(_LDOWN, 0, 0, 0, 0)
            Thread.Sleep(50)
            ctypes.windll.user32.mouse_event(_LUP, 0, 0, 0, 0)
            clicks_done += 1

            if clicks_done < total_clicks:
                if _interruptible_sleep(interval_ms):
                    self.WindowState = WindowState.Normal
                    self._set_status("Stopped by keypress after {} click(s).".format(clicks_done), error=True)
                    return

        self.WindowState = WindowState.Normal
        self._set_status("Done — {} click(s) completed.".format(clicks_done))

    # ══════════════════════════════════════════════════════════════════════
    # RECORD & REPLAY MODE
    # ══════════════════════════════════════════════════════════════════════

    def start_recording_clicked(self, sender, e):
        self.btn_start_record.IsEnabled = False
        self.btn_play.IsEnabled         = False
        self._set_status("Minimizing for 2 seconds — get ready to perform your actions...")
        self.WindowState = WindowState.Minimized
        Thread.Sleep(2000)
        _flush_keys()

        self._is_recording = True
        t = Thread(ThreadStart(self._record_worker))
        t.IsBackground = True
        t.Start()

    def _record_worker(self):
        recorded      = []
        last_left     = False
        last_right    = False
        last_time     = time.time()

        while self._is_recording:
            if _any_key_pressed():
                self._is_recording = False
                break

            now        = time.time()
            left_down  = bool(ctypes.windll.user32.GetAsyncKeyState(_VK_LBUTTON) & 0x8000)
            right_down = bool(ctypes.windll.user32.GetAsyncKeyState(_VK_RBUTTON) & 0x8000)
            pos        = Cursor.Position

            if left_down and not last_left:
                delay = int((now - last_time) * 1000)
                recorded.append(('LEFT_DOWN', pos.X, pos.Y, delay))
                last_time = now
            elif not left_down and last_left:
                delay = int((now - last_time) * 1000)
                recorded.append(('LEFT_UP', pos.X, pos.Y, delay))
                last_time = now

            if right_down and not last_right:
                delay = int((now - last_time) * 1000)
                recorded.append(('RIGHT_DOWN', pos.X, pos.Y, delay))
                last_time = now
            elif not right_down and last_right:
                delay = int((now - last_time) * 1000)
                recorded.append(('RIGHT_UP', pos.X, pos.Y, delay))
                last_time = now

            last_left  = left_down
            last_right = right_down
            Thread.Sleep(20)

        self._recorded_actions = recorded
        self.Dispatcher.BeginInvoke(
            DispatcherPriority.Normal,
            Action(self._on_record_finished)
        )

    def _on_record_finished(self):
        self._update_record_list()
        self.WindowState            = WindowState.Normal
        n                           = len(self._recorded_actions)
        self.lbl_action_count.Text  = "{} action(s)".format(n)
        self.btn_start_record.IsEnabled = True
        self.btn_play.IsEnabled     = n > 0
        self._set_status("Recording done — {} action(s) captured. Press Play to replay.".format(n))

    def _update_record_list(self):
        self.lst_actions.Items.Clear()
        for i, (atype, x, y, delay) in enumerate(self._recorded_actions):
            line = "{:3d}.  {:<12s}  X:{:4d}  Y:{:4d}  +{:4d}ms".format(
                i + 1, atype, x, y, delay)
            self.lst_actions.Items.Add(line)

    # ── Replay ──────────────────────────────────────────────────────────────

    def play_clicked(self, sender, e):
        if not self._recorded_actions:
            self._set_status("No actions recorded.", error=True)
            return
        try:
            loops = max(1, int(self.txt_loops.Text))
        except ValueError:
            loops = 1

        self.btn_play.IsEnabled         = False
        self.btn_start_record.IsEnabled = False
        self._set_status("Starting replay in 2 seconds...")
        self.WindowState = WindowState.Minimized
        Thread.Sleep(2000)
        _flush_keys()

        actions = list(self._recorded_actions)
        t = Thread(ThreadStart(lambda: self._play_worker(actions, loops)))
        t.IsBackground = True
        t.Start()

    def _play_worker(self, actions, loops):
        aborted     = False
        loops_done  = 0

        for _ in xrange(loops):
            if aborted:
                break
            for (atype, x, y, delay_ms) in actions:
                if _interruptible_sleep(delay_ms):
                    aborted = True
                    break

                ctypes.windll.user32.SetCursorPos(x, y)
                if atype == 'LEFT_DOWN':
                    ctypes.windll.user32.mouse_event(_LDOWN, 0, 0, 0, 0)
                elif atype == 'LEFT_UP':
                    ctypes.windll.user32.mouse_event(_LUP,   0, 0, 0, 0)
                elif atype == 'RIGHT_DOWN':
                    ctypes.windll.user32.mouse_event(_RDOWN, 0, 0, 0, 0)
                elif atype == 'RIGHT_UP':
                    ctypes.windll.user32.mouse_event(_RUP,   0, 0, 0, 0)

            if not aborted:
                loops_done += 1

        is_error = aborted
        msg = ("Replay stopped by keypress after {} loop(s).".format(loops_done)
               if aborted else
               "Replay complete — {} loop(s) done.".format(loops_done))

        def _finish():
            self.WindowState                = WindowState.Normal
            self.btn_play.IsEnabled         = True
            self.btn_start_record.IsEnabled = True
            self._set_status(msg, error=is_error)

        self.Dispatcher.BeginInvoke(DispatcherPriority.Normal, Action(_finish))

    # ── Clear ────────────────────────────────────────────────────────────────

    def clear_clicked(self, sender, e):
        self._recorded_actions     = []
        self.lst_actions.Items.Clear()
        self.lbl_action_count.Text = "0 actions"
        self.btn_play.IsEnabled    = False
        self._set_status("Cleared.")


# MAIN SCRIPT
# ==================================================
if __name__ == '__main__':
    AutoWorkWindow().ShowDialog()
