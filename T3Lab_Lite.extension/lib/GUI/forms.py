from WPF_Base import my_WPF
from FindReplace import FindReplace
from SelectFromDict import select_from_dict


class ListItem:
    """Helper Class for displaying selected sheets in my custom GUI.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""
    def __init__(self,  Name='Unnamed', element = None, checked = False):
        self.Name       = Name
        self.IsChecked  = checked
        self.element    = element