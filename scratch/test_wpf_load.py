import sys
import os
import io
import clr
clr.AddReference('System.Xml')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, WindowState
import wpf

class DummyWindow(Window):
    def __init__(self, xaml_path):
        print("Initializing DummyWindow with XAML path: {}".format(xaml_path))
        # Use wpf.LoadComponent to load the XAML and wire up event handlers
        wpf.LoadComponent(self, xaml_path)

    def minimize_button_clicked(self, sender, e):
        pass

    def maximize_button_clicked(self, sender, e):
        pass

    def close_button_clicked(self, sender, e):
        pass

def test_load():
    xaml_path = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\lib\GUI\Tools\UIStandardShowcase.xaml"
    try:
        win = DummyWindow(xaml_path)
        print("DummyWindow loaded successfully!")
    except Exception as e:
        print("Error loading DummyWindow:")
        print(e)
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_load()
