import sys
import os
import io
import clr
clr.AddReference('System.Xml')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode, FileAccess, FileShare
from System.Xml import XmlReader

def test_load():
    xaml_path = ".claude/standard/UIStandardShowcase.xaml"
    print("Loading XAML from: {}".format(xaml_path))
    try:
        # We need to open the file and read it
        with io.open(xaml_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # Parse it using XamlReader
        # In IronPython/WPF, we can load from a string using System.IO.StringReader and XmlReader
        from System.IO import StringReader
        from System.Xml import XmlReader
        
        string_reader = StringReader(content)
        xml_reader = XmlReader.Create(string_reader)
        
        obj = XamlReader.Load(xml_reader)
        print("XAML loaded successfully!")
    except Exception as e:
        print("Error loading XAML:")
        print(e)
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_load()
