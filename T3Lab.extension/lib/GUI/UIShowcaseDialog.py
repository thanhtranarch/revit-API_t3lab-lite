# -*- coding: utf-8 -*-
"""T3 UI Standard Showcase — cửa sổ tham chiếu của Design System.

MỘT cửa sổ, năm nhóm component, đổi bằng rail bên trái. Đây là file mẫu để
copy khi dựng tool mới: mở nhóm cần dùng trong Revit, tìm khối tương ứng trong
`lib/GUI/Tools/UIStandardShowcase.xaml`, copy nguyên khối, đổi chữ.

Class này chỉ làm 3 việc — nạp XAML, nạp dữ liệu mẫu cho bảng, xử lý chrome
cửa sổ (min/max/close + rail). Không có logic Revit nào; showcase không đụng
vào model.

Lịch sử: bản trước còn kéo theo bộ theme sáng/tối của chuẩn "Revit-native" đã
bị bỏ (`RevitTheme.apply`, `chk_dark_preview`, `host_report`…). Chuẩn T3 chỉ có
một bảng màu, khai trong `T3Lab.Styles.xaml` và nhúng thẳng vào XAML, nên toàn
bộ khối đó đã được gỡ — nó bind vào những `x:Name` không còn tồn tại và chỉ
sống sót nhờ `try/except`, đúng thứ luật S3 cấm.
"""

import os

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import (MessageBox, MessageBoxButton, MessageBoxImage,
                            Visibility, WindowState)
from System.Windows.Controls.Primitives import ToggleButton
from System.Collections.ObjectModel import ObservableCollection

from pyrevit import forms

# XAML nằm cùng extension, không phụ thuộc repo checkout — tool vẫn chạy khi
# extension được deploy riêng (vd %APPDATA%\pyRevit\Extensions).
GUI_DIR = os.path.dirname(__file__)                       # [ext]/lib/GUI
XAML_FILE = os.path.join(GUI_DIR, 'Tools', 'UIStandardShowcase.xaml')


class ShowcaseItem(object):
    """Một hàng của bảng trong nhóm Data.

    `T3.StatusPill` cần đúng 2 property: StatusText (chữ) và Severity
    ("Success" / "Warning" / "Danger"). Trạng thái KHÔNG BAO GIỜ chỉ bằng màu —
    chấm màu luôn đi kèm chữ.
    """

    def __init__(self, item_id, name, category, template, status_text,
                 severity, issues):
        self._id = item_id
        self._name = name
        self._category = category
        self._template = template
        self._status_text = status_text
        self._severity = severity
        self._issues = issues

    @property
    def id(self):
        return self._id

    @property
    def name(self):
        return self._name

    @property
    def category(self):
        return self._category

    @property
    def template(self):
        return self._template

    @property
    def issues(self):
        return self._issues

    # T3.StatusPill bind vào đúng 2 tên này
    @property
    def StatusText(self):
        return self._status_text

    @property
    def Severity(self):
        return self._severity


class UIShowcaseWindow(forms.WPFWindow):
    """Cửa sổ showcase. Caption do XAML tự vẽ (WindowStyle="None" + WindowChrome)."""

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._load_sample_data()          # nạp NGAY, không đợi Loaded (luật S7)

    # ── Chrome cửa sổ ────────────────────────────────────────────────────────

    def setup_icon(self):
        """Không gắn icon vào caption — caption là của mình, không có chỗ cho nó."""
        return

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
        """Nút close ở caption, và Close / Done ở footer."""
        self.Close()

    # ── Rail ─────────────────────────────────────────────────────────────────

    def rail_clicked(self, sender, e):
        """Đổi nhóm component. Tag của tile = index của TabItem.

        ToggleButton không tự nhóm như RadioButton, nên phải tự bỏ chọn các tile
        còn lại — và ép tile vừa bấm luôn ở trạng thái checked, kể cả khi người
        dùng bấm lại đúng tile đang mở (nếu không nó tự bỏ chọn và rail trông
        như không có nhóm nào đang mở).
        """
        try:
            index = int(str(sender.Tag))
        except (TypeError, ValueError):
            return

        self.content_tabs.SelectedIndex = index
        for tile in self.rail_tiles.Children:
            if isinstance(tile, ToggleButton):
                tile.IsChecked = (tile is sender)

    # ── Dữ liệu mẫu ──────────────────────────────────────────────────────────

    def _load_sample_data(self):
        """Dữ liệu mẫu cho bảng nhóm Data. Chữ hiển thị là TIẾNG ANH (luật 14)."""
        items = [
            ShowcaseItem("104231", "01_Plan_Column setting-out", "Floor Plan",
                         "ARC_Plan_1-100", "Compliant", "Success", "0"),
            ShowcaseItem("104232", "02_Plan_Column dimensions", "Floor Plan",
                         "ARC_Plan_1-100", "Compliant", "Success", "0"),
            ShowcaseItem("104235", "Detail_Column rebar splice", "Detail View",
                         "— none —", "Template missing", "Danger", "1"),
            ShowcaseItem("104240", "3D_Overall perspective", "3D View",
                         "— none —", "Naming standard", "Danger", "2"),
            ShowcaseItem("104245", "Section_Longitudinal section", "Section",
                         "ARC_Sect_1-50", "Needs review", "Warning", "1"),
            ShowcaseItem("104250", "Elevation_Grid A-D", "Elevation",
                         "ARC_Elev_1-100", "Compliant", "Success", "0"),
            ShowcaseItem("104255", "Schedule_Beam rebar take-off", "Schedule",
                         "n/a", "Compliant", "Success", "0"),
            ShowcaseItem("104261", "Detail_Pad footing M1", "Detail View",
                         "STR_Det_1-25", "Naming standard", "Danger", "1"),
            ShowcaseItem("104266", "Plan_Services ceiling", "Ceiling Plan",
                         "ARC_Plan_1-100", "Compliant", "Success", "0"),
            ShowcaseItem("104270", "Elevation_Grid 1-9", "Elevation",
                         "ARC_Elev_1-100", "Needs review", "Warning", "3"),
        ]
        self.sample_grid.ItemsSource = ObservableCollection[object](items)
        # Empty state chỉ hiện khi thật sự không có hàng nào — đây là cách mọi
        # tool phải làm, không phải để Collapsed vĩnh viễn trong XAML.
        self.grid_empty.Visibility = \
            Visibility.Collapsed if items else Visibility.Visible


def show_ui_standard_showcase():
    """Mở showcase."""
    try:
        UIShowcaseWindow().ShowDialog()
    except Exception as ex:
        import traceback
        traceback.print_exc()
        MessageBox.Show(
            "Could not open the UI Standard Showcase.\n\n{}\n\n"
            "Check that lib/GUI/Tools/UIStandardShowcase.xaml is present and "
            "that the embedded T3 stylesheet is in sync "
            "(python3 dev/sync_t3_styles.py).".format(ex),
            "T3 UI Standard Showcase",
            MessageBoxButton.OK,
            MessageBoxImage.Error)
