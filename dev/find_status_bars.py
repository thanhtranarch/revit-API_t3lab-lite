# -*- coding: utf-8 -*-
import os
import glob
import re

REPO_ROOT = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api"
TOOLS_DIR = os.path.join(REPO_ROOT, "T3Lab.extension", "lib", "GUI", "Tools")

def main():
    xaml_files = sorted(glob.glob(os.path.join(TOOLS_DIR, "*.xaml")))
    print("Scanning {} files for status bar patterns...".format(len(xaml_files)))
    
    # We will search for common status control names
    names = ["status", "status_text", "txtStatus", "statusText", "lblStatus", "txt_status", "lbl_status", "status_lbl", "lbl_status_text"]
    
    for path in xaml_files:
        filename = os.path.basename(path)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
            
        found_name = None
        for name in names:
            if 'x:Name="{}"'.format(name) in content or 'Name="{}"'.format(name) in content:
                found_name = name
                break
                
        # Also look for any comment with STATUS BAR or FOOTER
        comment_match = re.search(r'<!--.*?STATUS.*?BAR.*?-->|<!--.*?FOOTER.*?-->', content, re.IGNORECASE)
        comment_text = comment_match.group(0) if comment_match else None
        
        # Let's see if we can find the Grid.Row containing the status/footer
        # usually it is the last row. Let's look for borders near the end of the file.
        # Find all Border tags
        borders = re.findall(r'<Border\s+[^>]*?Grid\.Row="(\d+)"[^>]*?>|<Border\s+[^>]*?>', content)
        
        ascii_comment = comment_text.encode("ascii", errors="replace").decode("ascii") if comment_text else None
        print("{}: Name={!r}, Comment={!r}".format(filename, found_name, ascii_comment))

if __name__ == "__main__":
    main()
