# -*- coding: utf-8 -*-
"""
AssistantCards — code-built WPF cards for the T3Lab Assistant chat.

Keeps script.py thin: card construction lives here, callbacks stay in
the window class. Zinc styling matches T3LabAssistant.xaml's local
palette. UI THREAD only (WPF imports happen at call time).

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__ = "Assistant Cards"

_ACTION_LABELS = {
    'fix_parameter':  'Sửa parameter',
    'move_element':   'Di chuyển element',
    'fix_tag':        'Sửa tag',
    'fix_dimension':  'Sửa dimension',
    'create_element': 'Tạo element',
    'delete_element': 'Xóa element',
    'place_note':     'Đặt ghi chú',
    'info':           'Chỉ thông tin',
    'manual':         'Xử lý tay',
}

_EXECUTABLE_ACTIONS = frozenset([
    'fix_parameter', 'move_element', 'fix_tag', 'fix_dimension',
    'create_element', 'delete_element', 'place_note',
])


def _brush(r, g, b):
    from System.Windows.Media import SolidColorBrush, Color
    return SolidColorBrush(Color.FromRgb(r, g, b))


def _font():
    from System.Windows.Media import FontFamily
    return FontFamily("Hanken Grotesk")


def _mini_button(label, tooltip, fg=(82, 82, 91)):
    from System.Windows.Controls import Button
    from System.Windows import Thickness
    import System.Windows.Input
    btn = Button()
    btn.Content = label
    btn.FontSize = 10
    btn.FontFamily = _font()
    btn.Foreground = _brush(*fg)
    btn.Background = _brush(244, 244, 246)
    btn.BorderBrush = _brush(230, 230, 234)
    btn.BorderThickness = Thickness(1)
    btn.Padding = Thickness(8, 2, 8, 3)
    btn.Margin = Thickness(0, 0, 4, 0)
    btn.Cursor = System.Windows.Input.Cursors.Hand
    btn.ToolTip = tooltip
    return btn


def build_comment_report_card(report, on_run, on_note, on_skip):
    """The comment-resolution report card.

    Args:
        report: CommentAgent.analyze() result.
        on_run/on_note/on_skip: callbacks receiving (item, status_setter)
            where status_setter(text, ok) updates the row's status label.

    Returns a WPF Border ready to append to the chat panel.
    """
    from System.Windows.Controls import (Border, TextBlock, StackPanel,
                                         Orientation, Grid, ColumnDefinition)
    from System.Windows import Thickness, CornerRadius, TextWrapping, GridLength
    import System.Windows

    card = Border()
    card.Background = _brush(255, 255, 255)
    card.BorderBrush = _brush(230, 230, 234)
    card.BorderThickness = Thickness(1)
    card.CornerRadius = CornerRadius(10)
    card.Padding = Thickness(14, 10, 14, 12)
    card.Margin = Thickness(34, 0, 40, 10)

    root = StackPanel()

    # ── header ──
    title = TextBlock()
    title.Text = "Comment bản vẽ: " + (report.get('pdf_name') or '')
    title.FontSize = 12.5
    title.FontWeight = System.Windows.FontWeights.SemiBold
    title.FontFamily = _font()
    title.Foreground = _brush(24, 24, 27)
    title.TextWrapping = TextWrapping.Wrap
    root.Children.Add(title)

    match = report.get('sheet_match')
    sub = TextBlock()
    if match:
        sub.Text = "Sheet khớp: {} — {}   ·   model đang mở".format(
            match.get('number', ''), match.get('name', ''))
        sub.Foreground = _brush(16, 185, 129)
    elif report.get('needs_switch'):
        sub.Text = ("Không thấy sheet trong model ACTIVE — có model khác "
                    "đang mở, cần switch document trước khi xử lý.")
        sub.Foreground = _brush(245, 158, 11)
    else:
        sub.Text = "Chưa khớp được sheet nào trong model đang mở."
        sub.Foreground = _brush(239, 68, 68)
    sub.FontSize = 10.5
    sub.FontFamily = _font()
    sub.TextWrapping = TextWrapping.Wrap
    sub.Margin = Thickness(0, 2, 0, 0)
    root.Children.Add(sub)

    if report.get('partial_extraction'):
        warn = TextBlock()
        warn.Text = "PDF nén một phần — có thể còn markup chưa đọc được."
        warn.FontSize = 9.5
        warn.FontFamily = _font()
        warn.Foreground = _brush(161, 161, 170)
        warn.Margin = Thickness(0, 2, 0, 0)
        root.Children.Add(warn)

    sep = Border()
    sep.Height = 1
    sep.Background = _brush(240, 240, 243)
    sep.Margin = Thickness(0, 8, 0, 8)
    root.Children.Add(sep)

    # ── items ──
    for item in report.get('items', []):
        root.Children.Add(_build_item_row(item, report,
                                          on_run, on_note, on_skip))

    card.Child = root
    return card


def _build_item_row(item, report, on_run, on_note, on_skip):
    from System.Windows.Controls import (Border, TextBlock, StackPanel,
                                         Orientation)
    from System.Windows import Thickness, CornerRadius, TextWrapping
    import System.Windows

    prop = item.get('proposal') or {}
    row = Border()
    row.Background = _brush(250, 250, 251)
    row.BorderBrush = _brush(236, 236, 239)
    row.BorderThickness = Thickness(1)
    row.CornerRadius = CornerRadius(8)
    row.Padding = Thickness(10, 7, 10, 8)
    row.Margin = Thickness(0, 0, 0, 6)

    box = StackPanel()

    head = TextBlock()
    head.Text = "[{}] trang {} · {} · {}".format(
        item.get('id', ''), item.get('page', ''),
        item.get('subtype', ''), item.get('author') or 'không rõ tác giả')
    head.FontSize = 9.5
    head.FontFamily = _font()
    head.Foreground = _brush(161, 161, 170)
    box.Children.Add(head)

    content = TextBlock()
    content.Text = (item.get('content') or item.get('subject')
                    or '(markup không có nội dung chữ)')
    content.FontSize = 12
    content.FontFamily = _font()
    content.Foreground = _brush(39, 39, 42)
    content.TextWrapping = TextWrapping.Wrap
    content.Margin = Thickness(0, 2, 0, 0)
    box.Children.Add(content)

    proposal = TextBlock()
    action = prop.get('action_type', 'manual')
    label = _ACTION_LABELS.get(action, action)
    desc = prop.get('description') or ''
    proposal.Text = "Đề xuất: [{}] {}".format(label, desc)
    proposal.FontSize = 10.5
    proposal.FontFamily = _font()
    proposal.Foreground = _brush(59, 130, 246)
    proposal.TextWrapping = TextWrapping.Wrap
    proposal.Margin = Thickness(0, 3, 0, 0)
    box.Children.Add(proposal)

    # buttons + status
    bar = StackPanel()
    bar.Orientation = Orientation.Horizontal
    bar.Margin = Thickness(0, 6, 0, 0)

    status = TextBlock()
    status.FontSize = 10
    status.FontFamily = _font()
    status.Foreground = _brush(161, 161, 170)
    status.VerticalAlignment = System.Windows.VerticalAlignment.Center
    status.Margin = Thickness(4, 0, 0, 0)

    def _make_setter(_status=status, _row=row):
        def _set(text, ok):
            try:
                _status.Text = text
                _status.Foreground = (_brush(16, 185, 129) if ok
                                      else _brush(161, 161, 170))
            except Exception:
                pass
        return _set
    setter = _make_setter()

    executable = action in _EXECUTABLE_ACTIONS
    if executable and on_run is not None:
        b_run = _mini_button("▶ Thực hiện", "Chạy phương án qua AI agent",
                             fg=(16, 185, 129))
        b_run.Click += _wrap_cb(on_run, item, setter)
        bar.Children.Add(b_run)
    if on_note is not None:
        b_note = _mini_button("Ghi chú vào sheet",
                              "Đặt TextNote đánh dấu comment này trong model")
        b_note.Click += _wrap_cb(on_note, item, setter)
        bar.Children.Add(b_note)
    if on_skip is not None:
        b_skip = _mini_button("Bỏ qua", "Đánh dấu không xử lý",
                              fg=(161, 161, 170))
        b_skip.Click += _wrap_cb(on_skip, item, setter)
        bar.Children.Add(b_skip)
    bar.Children.Add(status)

    box.Children.Add(bar)
    row.Child = box
    return row


def _wrap_cb(cb, item, setter):
    def _handler(sender, e):
        try:
            cb(item, setter)
        except Exception:
            pass
    return _handler
