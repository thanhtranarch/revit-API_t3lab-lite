import os
import re
import io

from dev.add_sidebar_nav import extract_tab_radiobutttons, inject_sidebar, add_column_def_to_root_grid

def test():
    # Read the version on disk (which has been modified by fix_ui.py)
    filepath = r'T3Lab.extension\lib\GUI\Tools\FamilyManagement.xaml'
    with io.open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
        
    tabs = extract_tab_radiobutttons(content)
    print("Found tabs:", [t['text'] for t in tabs])
    
    # Run inject_sidebar
    content_new, sidebar_xml = inject_sidebar(content, tabs, 'FamilyManagement.xaml')
    
    # Trace add_column_def_to_root_grid
    resources_end = content_new.rfind('</Window.Resources>')
    if resources_end == -1:
        resources_end = content_new.rfind('</Grid.Resources>')
    print("resources_end:", resources_end)
    
    rest = content_new[resources_end:]
    grid_match = re.search(r'<Grid\b', rest, re.IGNORECASE)
    if grid_match:
        grid_start_abs = resources_end + grid_match.start()
        print("grid_start_abs:", grid_start_abs)
        print("Grid match string:", rest[grid_match.start():grid_match.start()+100])
        
        grid_tag_match = re.match(r'<Grid\b[^>]*>', content_new[grid_start_abs:], re.IGNORECASE)
        if grid_tag_match:
            grid_body_start = grid_start_abs + grid_tag_match.end()
            print("content_new[:grid_body_start] suffix:")
            print(repr(content_new[grid_body_start-200:grid_body_start]))
            
            # Find the first child visual element
            child_match = re.search(r'<[a-zA-Z]+(?:\s|>|/)', content_new[grid_body_start:])
            first_child_idx = grid_body_start + child_match.start() if child_match else len(content_new)
            print("first_child_idx:", first_child_idx)
            if child_match:
                print("first child match:", child_match.group(0))
                print("context around first child:", content_new[first_child_idx:first_child_idx+100])
                
            col_defs_match = re.search(
                r'<Grid\.ColumnDefinitions>.*?</Grid\.ColumnDefinitions>',
                content_new[grid_body_start:],
                re.DOTALL | re.IGNORECASE
            )
            print("col_defs_match found:", col_defs_match is not None)
            if col_defs_match:
                col_defs_start_abs = grid_body_start + col_defs_match.start()
                print("col_defs_start_abs:", col_defs_start_abs)
                print("col_defs_start_abs < first_child_idx:", col_defs_start_abs < first_child_idx)
                
            # Perform the add_column_def_to_root_grid call
            final_content, ok = add_column_def_to_root_grid(content_new, sidebar_xml)
            print("add_column_def_to_root_grid ok:", ok)
            print("Did the opening Border tag survive?", "<Border BorderBrush=" in final_content)
            # Find index of "<Border BorderBrush="
            idx = final_content.find("<Border BorderBrush=")
            if idx != -1:
                print("Border found at index:", idx)
                print("Context around Border:", final_content[idx:idx+200])
            else:
                print("Border NOT found!")

if __name__ == '__main__':
    test()
