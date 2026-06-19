import sys
import os
import io

try:
    import clr
    clr.AddReference('System.Xml')
    clr.AddReference('PresentationFramework')
    clr.AddReference('PresentationCore')
    clr.AddReference('WindowsBase')
    from System.Windows.Markup import XamlReader
    from System.IO import StringReader
    from System.Xml import XmlReader
    HAS_CLR = True
except ImportError:
    HAS_CLR = False
    print("pythonnet (clr) is not available in this Python environment.")

def test_load():
    xaml_path = r"T3Lab.extension/lib/GUI/Tools/PropertyLine.xaml"
    print("Loading XAML from: {}".format(xaml_path))
    if not HAS_CLR:
        print("Skipping WPF XAML load test because pythonnet/clr is not available.")
        return
        
    try:
        with io.open(xaml_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        string_reader = StringReader(content)
        xml_reader = XmlReader.Create(string_reader)
        
        obj = XamlReader.Load(xml_reader)
        print("XAML loaded successfully!")
    except Exception as e:
        print("Error loading XAML:")
        print(e)
        if hasattr(e, 'LineNumber'):
            print("Line: {}, Column: {}".format(e.LineNumber, e.LinePosition))

if __name__ == '__main__':
    test_load()
