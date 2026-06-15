import xml.etree.ElementTree as ET

def get_line_number(element, file_path):
    # ElementTree doesn't store line numbers by default, so we do a quick line search
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    tag_name = element.tag.split('}')[-1]
    for i, line in enumerate(lines):
        if f"<{tag_name}" in line or f"<Grid.{tag_name}" in line:
            return i + 1
    return -1

def check_empty_property_elements(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    
    # Registry of namespaces
    namespaces = {node[0]: node[1] for _, node in ET.iterparse(file_path, events=['start-ns'])}
    
    empty_elements = []
    
    def walk(node, path=""):
        name = node.tag.split('}')[-1]
        current_path = f"{path}/{name}"
        
        # Check if tag is a property element (contains '.')
        if '.' in name:
            has_children = len(node) > 0
            has_text = node.text and node.text.strip()
            if not has_children and not has_text:
                empty_elements.append((name, current_path, node))
                
        for child in node:
            walk(child, current_path)
            
    walk(root)
    
    if not empty_elements:
        print("No empty property elements found!")
        return
        
    for name, path, node in empty_elements:
        print(f"Empty Property Element found: '{name}' at path '{path}'")

if __name__ == "__main__":
    check_empty_property_elements(".claude/standard/UIStandardShowcase.xaml")
