import os
import xml.etree.ElementTree as ET

def check_all_xaml():
    tools_dir = r"T3Lab.extension/lib/GUI/Tools"
    xaml_files = [os.path.join(tools_dir, f) for f in os.listdir(tools_dir) if f.endswith('.xaml')]
    
    # Also add standard showcase xaml
    xaml_files.append(r".claude/standard/UIStandardShowcase.xaml")
    
    errors = 0
    parsed = 0
    for path in xaml_files:
        try:
            tree = ET.parse(path)
            parsed += 1
        except Exception as e:
            print(f"ERROR: Failed to parse XML in {path}:")
            print(e)
            errors += 1
            
    print(f"\n--- Validation Complete ---")
    print(f"Successfully parsed XAML files: {parsed}")
    print(f"Failed XAML files: {errors}")
    return errors == 0

if __name__ == "__main__":
    check_all_xaml()
