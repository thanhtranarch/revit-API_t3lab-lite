#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automated UI fixer for T3Lab Revit extension tool windows.
Aligns XAML files with the Lumina Design Standard and injects click handlers in Python scripts.
"""

import os
import glob
import re
import io
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS_DIR = os.path.join(REPO_ROOT, "T3Lab.extension", "lib", "GUI", "Tools")
LIB_DIR = os.path.join(REPO_ROOT, "T3Lab.extension", "lib")

VARIANT_B_FILES = ["FamilyLoader.xaml", "FamilyLoaderCloud.xaml", "ParameterSelector.xaml"]

# ----------------------------------------------------
# XML matching / parsing helpers
# ----------------------------------------------------

def find_matching_end_tag(content, start_idx, tag_name):
    """Find matching closing tag, handling nested tags of the same name."""
    open_pattern = re.compile(r'<' + tag_name + r'\b', re.IGNORECASE)
    close_pattern = re.compile(r'</' + tag_name + r'>', re.IGNORECASE)
    self_closing_pattern = re.compile(r'<[^>]*?/\s*>', re.DOTALL)
    
    pos = start_idx
    depth = 0
    
    # We find next tag boundary
    while pos < len(content):
        # Find next open and close tag
        next_open = open_pattern.search(content, pos)
        next_close = close_pattern.search(content, pos)
        
        if not next_close:
            return -1
            
        if next_open and next_open.start() < next_close.start():
            # Check if this open tag is self-closing
            tag_end = content.find('>', next_open.start())
            if tag_end != -1:
                tag_content = content[next_open.start():tag_end+1]
                if tag_content.strip().endswith('/>'):
                    # Self closing, depth doesn't change
                    pos = tag_end + 1
                    continue
            depth += 1
            pos = next_open.end()
        else:
            depth -= 1
            if depth == 0:
                return next_close.end()
            pos = next_close.end()
    return -1

def extract_attribute(tag_content, attr_name):
    """Extract attribute value from tag string."""
    pattern = re.compile(r'\b' + attr_name + r'\s*=\s*"(.*?)"', re.IGNORECASE)
    match = pattern.search(tag_content)
    return match.group(1) if match else None

# ----------------------------------------------------
# XAML Repair Logic
# ----------------------------------------------------

def fix_xaml_file(path):
    with io.open(path, "r", encoding="utf-8") as f:
        content = f.read()
        
    filename = os.path.basename(path)
    is_variant_b = filename in VARIANT_B_FILES
    original = content
    modified = False
    
    # 1. FontFamily replacement (except Segoe MDL2 Assets / Consolas / JetBrains Mono)
    for font in ["Manrope", "Segoe UI", "Inter"]:
        pattern = re.compile(r'\bFontFamily="' + re.escape(font) + r'"', re.IGNORECASE)
        if pattern.search(content):
            content = pattern.sub('FontFamily="Hanken Grotesk"', content)
            modified = True
    # Restore Segoe MDL2 / Consolas / JetBrains Mono if mistakenly overwritten
    for restore_font in ["Segoe MDL2 Assets", "Consolas", "JetBrains Mono"]:
        bad = 'FontFamily="Hanken Grotesk"'
        orig = f'FontFamily="{restore_font}"'
        # The above regex won't touch these because we only replaced exact font names;
        # this guard is a no-op but kept for safety.

    # Temporarily extract resources blocks to prevent matching Grid.Row="0" inside them
    resources_pattern = re.compile(r'(<(\w+)\.Resources>.*?</\2\.Resources>)', re.DOTALL | re.IGNORECASE)
    resources_blocks = []
    
    while True:
        match = resources_pattern.search(content)
        if not match:
            break
        block = match.group(1)
        placeholder = f"<!-- RESOURCES_PLACEHOLDER_{len(resources_blocks)} -->"
        content = content.replace(block, placeholder)
        resources_blocks.append(block)

    # 2. Remove T3Lab_logo.png references and their empty hosting borders
    # Standard format: <Border Height="60" ...> <Image Source="...T3Lab_logo.png".../> </Border>
    # Let's search for logo image first
    logo_pattern = re.compile(r'<Image\s+[^>]*?Source="[^"]*?T3Lab_logo\.png"[^>]*?>', re.IGNORECASE)
    logo_match = logo_pattern.search(content)
    if logo_match:
        # Trace back to enclosing Border
        preceding = content[:logo_match.start()]
        border_idx = preceding.rfind("<Border")
        if border_idx != -1:
            # Check if it has a matching end tag after the image
            closing_idx = find_matching_end_tag(content, border_idx, "Border")
            if closing_idx != -1 and closing_idx > logo_match.end():
                # Remove the entire border containing the logo
                content = content[:border_idx] + content[closing_idx:]
                modified = True
        else:
            # Just remove the image tag
            content = content[:logo_match.start()] + content[logo_match.end():]
            modified = True

    if not is_variant_b:
        # 3. Root Window attributes: Background="White", ResizeMode="CanResizeWithGrip", FontFamily="Hanken Grotesk"
        window_match = re.search(r'<Window\b([^>]*?)>', content, re.DOTALL)
        if window_match:
            window_attrs = window_match.group(1)
            new_attrs = window_attrs
            
            # Background — Lumina standard is #E4E4E7 (zinc-200 app shell)
            bg = extract_attribute(window_attrs, "Background")
            if not bg:
                new_attrs += '\n        Background="#E4E4E7"'
            elif bg.lower() not in ["#e4e4e7"]:
                new_attrs = re.sub(r'\bBackground="[^"]*"', 'Background="#E4E4E7"', new_attrs)

            # ResizeMode
            rm = extract_attribute(window_attrs, "ResizeMode")
            if not rm:
                new_attrs += '\n        ResizeMode="CanResizeWithGrip"'
            elif rm != "CanResizeWithGrip":
                new_attrs = re.sub(r'\bResizeMode="[^"]*"', 'ResizeMode="CanResizeWithGrip"', new_attrs)

            # FontFamily
            ff = extract_attribute(window_attrs, "FontFamily")
            if not ff:
                new_attrs += '\n        FontFamily="Hanken Grotesk"'
            elif ff != "Hanken Grotesk":
                new_attrs = re.sub(r'\bFontFamily="[^"]*"', 'FontFamily="Hanken Grotesk"', new_attrs)
                
            if new_attrs != window_attrs:
                content = content[:window_match.start(1)] + new_attrs + content[window_match.end(1):]
                modified = True

        # 4. WindowChrome
        chrome_pattern = re.compile(r'<WindowChrome\.WindowChrome\b.*?>(.*?)</WindowChrome\.WindowChrome>', re.DOTALL | re.IGNORECASE)
        chrome_match = chrome_pattern.search(content)
        standard_chrome = """    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                       ResizeBorderThickness="5"
                       GlassFrameThickness="0"
                       CornerRadius="22"
                       UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>"""
    
        if chrome_match:
            # Check if it needs reformatting
            chrome_body = chrome_match.group(0)
            if "CaptionHeight=\"64\"" not in chrome_body or "CornerRadius=\"22\"" not in chrome_body:
                content = chrome_pattern.sub(standard_chrome, content)
                modified = True
        else:
            # Check if there is a single-line Chrome tag
            single_chrome = re.search(r'<WindowChrome\b.*?/>', content, re.IGNORECASE)
            if single_chrome:
                content = content[:single_chrome.start()] + standard_chrome + content[single_chrome.end():]
                modified = True
            else:
                # Insert right after the opening Window tag
                win_tag_end = content.find('>', window_match.start() if window_match else 0)
                if win_tag_end != -1:
                    content = content[:win_tag_end+1] + "\n\n" + standard_chrome + content[win_tag_end+1:]
                    modified = True

        # 5. Title Bar Replacement (Height=64)
        # Search for Title Bar grid by comment or Grid.Row="0"
        title_grid_start = -1
        comment_match = re.search(r'<!--\s*═══\s*TITLE BAR\s*═══\s*-->', content)
        if comment_match:
            # The grid starts after the comment
            grid_search = re.search(r'<Grid\b', content[comment_match.end():])
            if grid_search:
                title_grid_start = comment_match.end() + grid_search.start()
        else:
            # Look for Grid Grid.Row="0" or Border Grid.Row="0"
            grid_row_match = re.search(r'<(Grid|Border)\s+[^>]*?Grid\.Row="0"', content)
            if grid_row_match:
                title_grid_start = grid_row_match.start()

        if title_grid_start != -1:
            # Find the closing tag
            tag_name = "Grid" if content[title_grid_start:].strip().startswith("<Grid") else "Border"
            title_grid_end = find_matching_end_tag(content, title_grid_start, tag_name)
            
            if title_grid_end != -1:
                title_block = content[title_grid_start:title_grid_end]
                
                # Extract original title and subtitle
                win_title = extract_attribute(content, "Title") or "T3Lab Tool"
                tool_name = win_title.replace("T3Lab - ", "").replace("T3Lab -", "").strip()
                
                # Try to extract existing subtitle TextBlock Text attribute
                sub_text = "T3Lab Revit productivity tool"
                sub_match = re.search(r'<TextBlock\s+[^>]*?FontStyle="Italic"[^>]*?Text="([^"]+)"', title_block, re.IGNORECASE)
                if sub_match:
                    sub_text = sub_match.group(1)
                else:
                    # Try other textblocks
                    tb_matches = re.finditer(r'<TextBlock\s+([^>]*?)>', title_block)
                    texts = []
                    for m in tb_matches:
                        attrs = m.group(1)
                        if "Segoe MDL2" in attrs:
                            continue
                        t = extract_attribute(m.group(0), "Text")
                        if t and t not in ["T3Lab", tool_name] and not t.startswith("&#x"):
                            texts.append(t)
                    if texts:
                        sub_text = texts[0]
                
                # Construct standard title bar (Lumina: #F4F4F6 bg, CornerRadius 22)
                std_title_bar = """<!-- ═══ TITLE BAR ═══ -->
        <Grid Grid.Row="0" Height="64" Background="#F4F4F6">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Vertical" Margin="22,0,0,0" VerticalAlignment="Center"
                        WindowChrome.IsHitTestVisibleInChrome="True">
                <TextBlock Text="{tool_name}" FontSize="15" FontWeight="Bold" Foreground="#18181B"/>
                <TextBlock Text="{sub_text}" FontSize="12.5" Foreground="#71717A" Margin="0,2,0,0"/>
            </StackPanel>

            <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,16,0"
                        WindowChrome.IsHitTestVisibleInChrome="True">
                <Button x:Name="btn_minimize" WindowChrome.IsHitTestVisibleInChrome="True"
                        Style="{StaticResource WinCtrlButton}"
                        Click="minimize_button_clicked" ToolTip="Minimize">
                    <TextBlock Text="&#xE921;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
                </Button>
                <Button x:Name="btn_maximize" WindowChrome.IsHitTestVisibleInChrome="True"
                        Style="{StaticResource WinCtrlButton}"
                        Click="maximize_button_clicked" ToolTip="Maximize">
                    <TextBlock Text="&#xE922;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
                </Button>
                <Button x:Name="btn_close" WindowChrome.IsHitTestVisibleInChrome="True"
                        Style="{StaticResource CloseButton}"
                        Click="close_button_clicked" ToolTip="Close">
                    <TextBlock Text="&#xE8BB;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
                </Button>
            </StackPanel>

            <Border Height="1" VerticalAlignment="Bottom" Background="#DCDCE0" Grid.ColumnSpan="3"/>
        </Grid>""".replace("{tool_name}", tool_name).replace("{sub_text}", sub_text)
                
                # Replace the title grid
                replace_start = comment_match.start() if comment_match else title_grid_start
                content = content[:replace_start] + std_title_bar + content[title_grid_end:]
                modified = True

        # 6. Status Bar formatting
        # Look for status bar border
        status_bar_match = re.search(r'(<!--\s*═══\s*STATUS BAR\s*═══\s*-->\s*)?<Border\s+([^>]*?)>\s*(?:<Grid>|<StackPanel>)?\s*<StackPanel>\s*<!--\s*Progress panel', content, re.DOTALL | re.IGNORECASE)
        if not status_bar_match:
            status_bar_match = re.search(r'(<!--\s*═══\s*STATUS BAR\s*═══\s*-->\s*)?<Border\s+([^>]*?)>\s*<Grid>\s*<Grid\.ColumnDefinitions>.*?status_text', content, re.DOTALL | re.IGNORECASE)
        if not status_bar_match:
            # Fallback using status_text
            status_text_idx = content.find("status_text")
            if status_text_idx != -1:
                preceding = content[:status_text_idx]
                border_idx = preceding.rfind("<Border")
                if border_idx != -1:
                    border_tag = preceding[border_idx:preceding.find(">", border_idx)+1]
                    status_bar_match = re.search(r'(<!--\s*═══\s*STATUS BAR\s*═══\s*-->\s*)?<Border\s+([^>]*?)>', border_tag, re.IGNORECASE)
                    if status_bar_match:
                        # Re-evaluate start position in content
                        status_bar_match = re.search(r'<Border\s+([^>]*?)>', content[border_idx:], re.IGNORECASE)
                        # Offset the match group to absolute index
                        class OffsetMatch:
                            def __init__(self, m, offset):
                                self.m = m
                                self.offset = offset
                            def start(self, g=0): return self.m.start(g) + self.offset
                            def end(self, g=0): return self.m.end(g) + self.offset
                            def group(self, g=0): return self.m.group(g)
                        status_bar_match = OffsetMatch(status_bar_match, border_idx)

        if status_bar_match:
            sb_attrs = status_bar_match.group(1) if len(status_bar_match.group()) > 1 else status_bar_match.group(0)
            # Find attributes within <Border ...>
            border_match = re.search(r'<Border\s+([^>]*?)>', status_bar_match.group(0))
            if border_match:
                attrs = border_match.group(1)
                new_attrs = attrs
                
                # Replace properties
                for prop, val in [("Background", "#F8FAFC"), ("BorderBrush", "#E2E8F0"), ("BorderThickness", "0,1,0,0"), ("Padding", "14,8")]:
                    extracted = extract_attribute(attrs, prop)
                    if not extracted:
                        new_attrs += f' {prop}="{val}"'
                    else:
                        new_attrs = re.sub(r'\b' + prop + r'="[^"]*"', f'{prop}="{val}"', new_attrs)
                
                if new_attrs != attrs:
                    # Update inside content
                    start_idx = status_bar_match.start() + border_match.start(1)
                    end_idx = status_bar_match.start() + border_match.end(1)
                    content = content[:start_idx] + new_attrs + content[end_idx:]
                    modified = True

    # 7. Copyright Block — remove old floating copyright, embed in footer status bar
    # Remove any existing floating copyright TextBlock (Panel.ZIndex=999 style)
    old_copyright_patterns = [
        re.compile(r'\s*<\!--\s*Copyright\s+added\s+automatically\s*-->\s*<TextBlock[^>]*?Copyright by T3Lab[^>]*?/>', re.IGNORECASE | re.DOTALL),
        re.compile(r'\s*<TextBlock[^>]*?Panel\.ZIndex\s*=\s*"999"[^>]*?Copyright by T3Lab[^>]*?/>', re.IGNORECASE | re.DOTALL),
        re.compile(r'\s*<TextBlock[^>]*?Copyright by T3Lab[^>]*?Panel\.ZIndex[^>]*?/>', re.IGNORECASE | re.DOTALL),
    ]
    for pat in old_copyright_patterns:
        new_content = pat.sub('', content)
        if new_content != content:
            content = new_content
            modified = True

    # Inject copyright into the last status/footer Border if not already present
    # We look for the last Border that wraps action buttons or a status area
    copyright_embedded = ('Copyright by T3Lab' in content or
                          'Copyright: T3Lab' in content or
                          '&#169;' in content or
                          '\u00a9' in content)
    if not copyright_embedded:
        # Find last <Border ...Background="#F4F4F6"... or #F8FAFC... near </Window>
        # Strategy: inject before the last </StackPanel> that comes before </Border> near </Window>
        # Simpler: append inside the last Grid before </Window>
        footer_copyright_block = """
        <!-- Copyright: T3Lab -->
        <Border CornerRadius="10" Background="Transparent" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Right">
                <TextBlock Text="&#169; 2026 T3Lab" FontSize="10" Foreground="#9A9AA2" HorizontalAlignment="Right"/>
                <TextBlock Text="&#169; Copyright by T3Lab" FontSize="10" Foreground="#F59E0B" Margin="0,2,0,0" HorizontalAlignment="Right"/>
            </StackPanel>
        </Border>"""
        # Find last </Grid> before </Window>
        window_close = content.rfind('</Window>')
        if window_close != -1:
            search_area = content[:window_close]
            last_grid_close = search_area.rfind('</Grid>')
            if last_grid_close != -1:
                content = content[:last_grid_close] + footer_copyright_block + '\n' + content[last_grid_close:]
                modified = True

    # Restore resources blocks FIRST (needed so </Window.Resources> is visible for border wrap)
    for i in range(len(resources_blocks)):
        placeholder = f"<!-- RESOURCES_PLACEHOLDER_{i} -->"
        content = content.replace(placeholder, resources_blocks[i])

    # 8. Outer Border wrapper — add ClipToBounds border if not already wrapped
    # Must run AFTER resources are restored so </Window.Resources> is in content
    outer_border_marker = 'ClipToBounds="True"'
    if outer_border_marker not in content:
        # Strategy: find first non-blank line after </Window.Resources>
        # That line starts the root content element (Grid or Border)
        res_end_idx = content.rfind('</Window.Resources>')
        if res_end_idx == -1:
            res_end_idx = content.rfind('</WindowChrome.WindowChrome>')
        if res_end_idx != -1:
            after_res = content[res_end_idx + len('</Window.Resources>'):]
            # Find first <Grid or <Border (the root layout element)
            root_match = re.search(r'\n(\s*)(<(?:Grid|Border)\b[^>]*>)', after_res)
            if root_match:
                root_tag_name = re.match(r'<(\w+)', root_match.group(2)).group(1)
                root_indent = root_match.group(1)

                # The root element ends just before </Window>
                # Find the last matching </Grid> or </Border> before </Window>
                window_close = content.rfind('</Window>')
                if window_close != -1:
                    search_region = content[:window_close]
                    last_close = search_region.rfind(f'</{root_tag_name}>')
                    if last_close != -1:
                        root_start_abs = res_end_idx + len('</Window.Resources>') + root_match.start(2)
                        root_end_abs = last_close + len(f'</{root_tag_name}>')
                        inner = content[root_start_abs:root_end_abs]
                        outer_open = f'{root_indent}<Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22" ClipToBounds="True" Background="#E4E4E7">\n'
                        outer_close = f'{root_indent}</Border>\n'
                        content = (
                            content[:root_start_abs]
                            + outer_open
                            + inner
                            + '\n' + outer_close
                            + content[root_end_abs:]
                        )
                        modified = True






    if modified or content != original:
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(content)
        print("FIXED XAML: %s" % filename)
        return True
    return False

# ----------------------------------------------------
# Python Click Handlers Injection
# ----------------------------------------------------

def fix_python_file(path):
    with io.open(path, "r", encoding="utf-8") as f:
        content = f.read()
        
    original = content
    modified = False
    
    # 1. Verify and inject WindowState import if needed
    # Check if WindowState is imported
    if "WindowState" not in content:
        # Find System.Windows imports
        import_match = re.search(r'from System\.Windows import\b(.*?)\n', content)
        if import_match:
            imports = import_match.group(1)
            # Add WindowState to existing import line
            new_imports = imports.strip()
            if not new_imports.endswith(')'):
                new_imports = " " + new_imports.lstrip('(').rstrip(',') + ", WindowState"
                if imports.startswith('('):
                    new_imports = " (" + new_imports.strip()
            else:
                new_imports = " " + new_imports.lstrip('(').rstrip(')').rstrip(',').strip() + ", WindowState)"
                if imports.startswith('('):
                    new_imports = " (" + new_imports.strip()
            
            content = content[:import_match.start(1)] + new_imports + content[import_match.end(1):]
            modified = True
        else:
            # Append global import at top
            import_line = "from System.Windows import WindowState\n"
            content = import_line + content
            modified = True

    # 2. Inject click handlers inside the Window classes
    # Look for class declarations that subclass WPFWindow or Window
    class_matches = list(re.finditer(r'class\s+(\w+)\s*\((forms\.WPFWindow|Window|forms\.WPFWindowBase)\):', content))
    
    for cm in class_matches:
        class_name = cm.group(1)
        # Find where this class ends. In Python, class ends when indentation returns to 0
        class_start_idx = cm.end()
        
        # Check if the class already has the click handlers
        has_min = "def minimize_button_clicked" in content[class_start_idx:]
        has_max = "def maximize_button_clicked" in content[class_start_idx:]
        has_close = "def close_button_clicked" in content[class_start_idx:]
        
        if not (has_min and has_max and has_close):
            # We find the next method definition inside the class to find indentation,
            # or scan lines in class body.
            lines = content[class_start_idx:].splitlines(True)
            indentation = "    " # default 4 spaces
            
            # Find indentation of the first def statement in lines
            for line in lines:
                if line.strip().startswith("def "):
                    indentation = line[:line.find("def")]
                    break
            
            # Find where to insert the methods.
            # We can append them at the end of the class body (before the next unindented line, or end of file)
            insert_idx = -1
            curr_pos = class_start_idx
            
            for line in lines:
                # If there's an unindented line (starts with word/no-space, and not a comment or empty)
                # it means the class has ended
                if line.strip() and not line.startswith(" ") and not line.startswith("\t") and not line.startswith("#"):
                    insert_idx = curr_pos
                    break
                curr_pos += len(line)
            
            if insert_idx == -1:
                insert_idx = len(content)
                
            # Construct standard handlers
            handlers = f"""

{indentation}def minimize_button_clicked(self, sender, e):
{indentation}    self.WindowState = WindowState.Minimized

{indentation}def maximize_button_clicked(self, sender, e):
{indentation}    if self.WindowState == WindowState.Maximized:
{indentation}        self.WindowState = WindowState.Normal
{indentation}        self.btn_maximize.ToolTip = "Maximize"
{indentation}    else:
{indentation}        self.WindowState = WindowState.Maximized
{indentation}        self.btn_maximize.ToolTip = "Restore"

{indentation}def close_button_clicked(self, sender, e):
{indentation}    self.Close()
"""
            
            content = content[:insert_idx] + handlers + content[insert_idx:]
            modified = True
            
    if modified or content != original:
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(content)
        print("INJECTED HANDLERS: %s" % os.path.basename(path))
        return True
    return False

# ----------------------------------------------------
# Main execution
# ----------------------------------------------------

def main():
    # 1. Mappings mapping XAML -> Py
    xaml_files = sorted(glob.glob(os.path.join(TOOLS_DIR, "*.xaml")))
    print("Found %d XAML files." % len(xaml_files))
    
    # 2. Fix all XAML files
    fixed_xaml = 0
    for path in xaml_files:
        if fix_xaml_file(path):
            fixed_xaml += 1
            
    # 3. Find loading Python scripts and fix them
    # Build list of all .py files in repo
    py_files = []
    for root, dirs, files in os.walk(os.path.join(REPO_ROOT, "T3Lab.extension")):
        for file in files:
            if file.endswith(".py"):
                py_files.append(os.path.join(root, file))
                
    # Check which py files load our tool xamls (Variant A standard only)
    xaml_basenames = [os.path.basename(p) for p in xaml_files if os.path.basename(p) not in VARIANT_B_FILES]
    
    target_py_files = set()
    for py_path in py_files:
        try:
            with io.open(py_path, "r", encoding="utf-8") as f:
                py_content = f.read()
            for xaml in xaml_basenames:
                if xaml in py_content:
                    target_py_files.add(py_path)
        except Exception:
            pass
            
    print("Found %d Python files referencing standard XAML windows." % len(target_py_files))
    
    fixed_py = 0
    for py_path in sorted(list(target_py_files)):
        if fix_python_file(py_path):
            fixed_py += 1
            
    print("\n--- REPAIR COMPLETE ---")
    print(f"XAML files fixed: {fixed_xaml}")
    print(f"Python files fixed: {fixed_py}")

if __name__ == "__main__":
    main()
