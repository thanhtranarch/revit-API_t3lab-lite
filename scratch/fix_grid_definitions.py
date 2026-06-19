import os

def fix_file():
    path = "T3Lab.extension/lib/GUI/Tools/PropertyLine.xaml"
    if not os.path.exists(path):
        print("File not found:", path)
        return
        
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
        
    original = content
    
    # Standard replacement for ColumnDefinition
    content = content.replace("<Grid.ColumnDefinition ", "<ColumnDefinition ")
    content = content.replace("<Grid.ColumnDefinition/>", "<ColumnDefinition/>")
    content = content.replace("<Grid.ColumnDefinition />", "<ColumnDefinition />")
    content = content.replace("</Grid.ColumnDefinition>", "</ColumnDefinition>")
    
    # Standard replacement for RowDefinition
    content = content.replace("<Grid.RowDefinition ", "<RowDefinition ")
    content = content.replace("<Grid.RowDefinition/>", "<RowDefinition/>")
    content = content.replace("<Grid.RowDefinition />", "<RowDefinition />")
    content = content.replace("</Grid.RowDefinition>", "</RowDefinition>")
    
    if content != original:
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        print("Successfully updated XAML definitions.")
    else:
        print("No changes needed or matching tags found.")

if __name__ == '__main__':
    fix_file()
