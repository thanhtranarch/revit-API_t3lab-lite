import xml.etree.ElementTree as ET
import sys

def parse_and_check():
    xaml_path = "T3Lab.extension/lib/GUI/Tools/PropertyLine.xaml"
    print("Checking XAML:", xaml_path)
    
    # We want to keep track of line numbers, so we can use a custom parser or read file lines.
    # ET doesn't give line numbers easily in older versions, but we can do a line-by-line search
    # and print elements with '.' in their tag names.
    with open(xaml_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    for i, line in enumerate(lines):
        stripped = line.strip()
        if '<' in stripped:
            # extract tag name
            start = stripped.find('<')
            end = stripped.find(' ', start)
            if end == -1 or (stripped.find('>', start) != -1 and stripped.find('>', start) < end):
                end = stripped.find('>', start)
            tag = stripped[start+1:end].replace('/', '').replace('>', '').strip()
            
            if '.' in tag:
                # This is a property element (e.g., Grid.RowDefinitions, WindowChrome.WindowChrome)
                # Check if it is self-closing
                is_self_closing = '/>' in stripped or '/>' in lines[i] or (i+1 < len(lines) and '</' + tag + '>' in lines[i+1])
                # Wait, if it is closed on the same line or next line without content, it might be considered empty.
                print("Line {:4d}: Tag: {:<30} | Self-closing/Empty?: {}".format(i+1, tag, is_self_closing))

if __name__ == '__main__':
    parse_and_check()
