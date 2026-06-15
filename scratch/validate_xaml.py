import os
import xml.etree.ElementTree as ET

def validate_all():
    dir_path = r'T3Lab.extension\lib\GUI\Tools'
    errors = 0
    for filename in os.listdir(dir_path):
        if filename.endswith('.xaml'):
            filepath = os.path.join(dir_path, filename)
            try:
                # Remove namespaces or just parse
                # Since XAML has default namespace, ET.parse is fine
                ET.parse(filepath)
            except Exception as e:
                print(f"FAILED: {filename} - {e}")
                errors += 1
    if errors == 0:
        print("All XAML files are well-formed XML!")
    else:
        print(f"Found {errors} invalid XAML files.")

if __name__ == '__main__':
    validate_all()
