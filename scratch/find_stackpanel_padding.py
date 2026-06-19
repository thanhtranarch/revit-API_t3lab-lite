import xml.etree.ElementTree as ET

def find_padding():
    path = "T3Lab.extension/lib/GUI/Tools/PropertyLine.xaml"
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    for i, line in enumerate(lines):
        if "StackPanel" in line and "Padding" in line:
            print("Line {}: {}".format(i+1, line.strip()))

if __name__ == '__main__':
    find_padding()
