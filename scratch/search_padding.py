import sys

def show_stackpanel_contexts():
    # Force output encoding to utf-8
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass
        
    path = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\lib\GUI\Tools\UIStandardShowcase.xaml"
    with open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    for i, line in enumerate(lines):
        if 'stackpanel' in line.lower():
            print("--- StackPanel at Line {} ---".format(i+1))
            start = max(0, i-3)
            end = min(len(lines), i+4)
            for j in range(start, end):
                prefix = "-> " if j == i else "   "
                # Safe print with ascii/backslashreplace if needed
                text = lines[j].rstrip()
                try:
                    print("{}{:03d}: {}".format(prefix, j+1, text))
                except UnicodeEncodeError:
                    print("{}{:03d}: {}".format(prefix, j+1, text.encode('ascii', 'backslashreplace').decode('ascii')))

if __name__ == '__main__':
    show_stackpanel_contexts()
