#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
add_sidebar_nav.py
==================
Converts tools that use horizontal RadioButton tab bars (T3TopNavTabStyle)
into a sidebar icon-rail layout matching UIStandardShowcase.

Process:
  1. Detect RadioButton groups using T3TopNavTabStyle in the content body
  2. Extract tab names, x:Names, event handlers
  3. Wrap content in 2-column Grid (66px sidebar + *)
  4. Move RadioButtons to sidebar as T3SidebarToggle ToggleButtons (icon + tooltip)
  5. Remove the old horizontal tab row

Sidebar icons are auto-assigned from a lookup table based on tab name keywords.
"""
import os
import re
import io
import glob

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS_DIR = os.path.join(REPO_ROOT, "T3Lab.extension", "lib", "GUI", "Tools")

# ────────────────────────────────────────────────────────────────────
# Icon library: keyword → SVG path data (all normalised to 24x24 viewBox)
# ────────────────────────────────────────────────────────────────────
ICON_MAP = {
    "view":       "M4 5 h16 v14 h-16 z M4 10 h16 M10 10 v9",
    "sheet":      "M4 5 h16 v14 h-16 z M4 10 h16 M10 10 v9",
    "list":       "M4 8 h16 M4 12 h16 M4 16 h10",
    "export":     "M12 4 V14 M8 10 L12 14 L16 10 M5 18 H19 V21 H5 Z",
    "import":     "M12 20 V10 M8 14 L12 10 L16 14 M5 6 H19 V3 H5 Z",
    "filter":     "M4 8 h9 M17 8 h3 M4 16 h3 M11 16 h9",
    "setting":    "M12 15 A3 3 0 1 0 12 9 A3 3 0 0 0 12 15 M12 4 V6 M12 18 V20 M4.2 6.2 L5.6 7.6 M18.4 17.8 L19.8 19.2 M4 12 H6 M18 12 H20",
    "rename":     "M14 4 L20 10 L8 22 H2 V16 Z M11 7 L17 13",
    "batch":      "M3 5 h5 v5 H3 Z M10 7 h11 M10 12 h11 M10 17 h11 M3 14 h5 v5 H3 Z",
    "rule":       "M4 8 h9 M17 8 h3 M4 16 h3 M11 16 h9",
    "dim":        "M3 20 H21 M7 20 V8 M17 20 V8 M7 8 H17 M7 14 H17",
    "dimension":  "M3 20 H21 M7 20 V8 M17 20 V8 M7 8 H17 M7 14 H17",
    "tag":        "M4 4 L4 13 L12 20 L20 12 L13 4 Z M9 9 A1 1 0 1 0 9 7 A1 1 0 0 0 9 9",
    "text":       "M4 7 h16 M4 12 h16 M4 17 h10 M12 4 H4 V6",
    "grid":       "M4 4 h7 v7 h-7 z M13 4 h7 v7 h-7 z M4 13 h7 v7 h-7 z M13 13 h7 v7 h-7 z",
    "family":     "M12 4 L20 8 L12 12 L4 8 Z M4 12 L12 16 L20 12 M4 16 L12 20 L20 16",
    "link":       "M8 8 H5 A3 3 0 0 0 5 14 H8 M16 8 H19 A3 3 0 0 1 19 14 H16 M8 11 H16",
    "room":       "M3 5 V21 H21 V5 H3 Z M3 10 H21 M10 10 V21",
    "work":       "M12 20 A8 8 0 1 0 12 4 A8 8 0 0 0 12 20 M12 8 V12 L15 14",
    "auto":       "M12 3 L14.5 9 L21 9.5 L16 14 L17.5 21 L12 17.5 L6.5 21 L8 14 L3 9.5 L9.5 9 Z",
    "quick":      "M13 2 L4.5 13.5 H11 L10.5 22 L20 10.5 H13.5 L13 2 Z",
    "property":   "M4 4 h5 v5 H4 Z M4 13 h5 v5 H4 Z M13 4 h7 M13 9 h5 M13 13 h7 M13 18 h5",
    "place":      "M12 2 A5 5 0 0 0 12 12 A5 5 0 0 0 12 2 M12 12 V22 M8 17 H16",
    "check":      "M12 21 A9 9 0 1 0 12 3 A9 9 0 0 0 12 21 M8.4 12.2 L10.8 14.6 L15.6 9.6",
    "default":    "M4 6 h16 M4 10 h12 M4 14 h8 M4 18 h10",
}

def get_icon_for_tab(name_or_content):
    """Pick icon path from name keywords."""
    s = name_or_content.lower()
    for keyword, path in ICON_MAP.items():
        if keyword in s:
            return path
    return ICON_MAP["default"]


def extract_tab_radiobutttons(content):
    """
    Returns list of dicts: {name, content_text, handler, group, style_ref, raw}
    for RadioButtons using T3TopNavTabStyle or NavTab.
    """
    # Pattern: <RadioButton x:Name="..." Content="..." Style="{StaticResource T3TopNavTabStyle}" .../>
    pattern = re.compile(
        r'<RadioButton\s+'
        r'([^/]*?(?:T3TopNavTabStyle|NavTab)[^/]*?)'
        r'(?:/>'
        r'|>\s*(.*?)\s*</RadioButton>)',
        re.DOTALL | re.IGNORECASE
    )
    tabs = []
    for m in pattern.finditer(content):
        raw = m.group(0)
        attrs = m.group(1)
        body = m.group(2) if m.lastindex >= 2 else ""

        name_m = re.search(r'x:Name\s*=\s*"([^"]*)"', attrs)
        content_m = re.search(r'Content\s*=\s*"([^"]*)"', attrs)

        text = ""
        if content_m:
            text = content_m.group(1)
        elif body:
            # Check for <TextBlock Text="..."/>
            tb_text_m = re.search(r'<TextBlock\s+[^>]*?Text\s*=\s*"([^"]*)"', body, re.IGNORECASE)
            if tb_text_m:
                text = tb_text_m.group(1)
            else:
                # Check for <TextBlock>...</TextBlock> or direct text
                tb_body_m = re.search(r'<TextBlock\b[^>]*>(.*?)</TextBlock>', body, re.DOTALL | re.IGNORECASE)
                if tb_body_m:
                    text = tb_body_m.group(1).strip()
                else:
                    text = body.strip()

        text = text.replace("&amp;", "&")

        handler_m = re.search(r'(?:Checked|Click|Command)\s*=\s*"([^"]*)"', attrs)
        group_m = re.search(r'GroupName\s*=\s*"([^"]*)"', attrs)
        checked_m = re.search(r'IsChecked\s*=\s*"([^"]*)"', attrs)
        tab = {
            "raw": raw,
            "name": name_m.group(1) if name_m else "",
            "text": text,
            "handler": handler_m.group(1) if handler_m else "",
            "group": group_m.group(1) if group_m else "NavTabs",
            "is_checked": (checked_m.group(1).lower() == "true") if checked_m else False,
        }
        tabs.append(tab)
    return tabs


def build_sidebar_toggle(tab, idx):
    """Build a T3SidebarToggle ToggleButton for a tab."""
    icon_path = get_icon_for_tab(tab["text"] or tab["name"])
    checked_attr = ' IsChecked="True"' if tab["is_checked"] or idx == 0 else ""
    handler_attr = f' Click="{tab["handler"]}"' if tab["handler"] else ""
    name_attr = f' x:Name="{tab["name"]}"' if tab["name"] else ""

    return f"""
                    <!-- Tab: {tab["text"]} -->
                    <ToggleButton{name_attr} Grid.Row="{idx + 1}"
                                  Style="{{StaticResource T3SidebarToggle}}"
                                  Margin="0,0,0,8"{checked_attr}
                                  ToolTip="{tab["text"]}"{handler_attr}>
                        <Path Data="{icon_path}"
                               Stroke="{{Binding Path=Foreground, RelativeSource={{RelativeSource AncestorType=ToggleButton}}}}"
                               StrokeThickness="1.8" StrokeLineJoin="Round"
                               StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                               Width="17" Height="17" Stretch="Uniform"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </ToggleButton>"""


def inject_sidebar(content, tabs, filename):
    """
    Wrap the root Grid content in a 2-column layout:
    Col 0: 66px sidebar with tabs as T3SidebarToggle
    Col 1: original content (minus horizontal tab row)
    """
    # 1. Remove original RadioButton tab rows from content
    for tab in tabs:
        content = content.replace(tab["raw"], "", 1)

    # 2. Find & remove the horizontal tab container (StackPanel/Grid that held only those RadioButtons)
    #    Strategy: find any StackPanel/Grid/Border that now has only whitespace between its tags
    #    (because we removed all RadioButtons from it). Uses backreference to ensure tag names match.
    empty_containers = re.compile(
        r'(\s*<(StackPanel|Grid|Border)\b[^>]*?>\s*</\2>)',
        re.DOTALL | re.IGNORECASE
    )
    for _ in range(3):  # multiple passes
        content = empty_containers.sub('', content)

    # 3. Build sidebar ToggleButton XML
    row_defs = ""
    for i in range(len(tabs)):
        row_defs += f'                        <RowDefinition Height="Auto"/> <!-- Tab {i+1} -->\n'

    toggle_buttons = ""
    for i, tab in enumerate(tabs):
        toggle_buttons += build_sidebar_toggle(tab, i)

    # Get logo icon hint from filename
    tool_name_hint = os.path.splitext(filename)[0]
    logo_icon = get_icon_for_tab(tool_name_hint)

    sidebar_xml = f"""
            <!-- ════════════════════════════════════════════════════════════ -->
            <!-- COL 0: LEFT SIDEBAR ICON RAIL (all rows)                    -->
            <!-- ════════════════════════════════════════════════════════════ -->
            <Border Grid.Column="0" Grid.RowSpan="99"
                    Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,0,1,0">
                <Grid Margin="0,16,0,16">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/> <!-- Logo -->
{row_defs}                        <RowDefinition Height="*"/>    <!-- Spacer -->
                    </Grid.RowDefinitions>

                    <!-- Logo icon -->
                    <Border Grid.Row="0" Width="42" Height="42" CornerRadius="13"
                            Background="White" BorderBrush="#DCDCE0" BorderThickness="1"
                            HorizontalAlignment="Center" Margin="0,0,0,20">
                        <Path Data="{logo_icon}"
                              Stroke="#18181B" StrokeThickness="1.6" StrokeLineJoin="Round"
                              StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                              Width="16" Height="16" Stretch="Uniform"
                              HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
{toggle_buttons}
                </Grid>
            </Border>
"""
    return content, sidebar_xml


def add_column_def_to_root_grid(content, sidebar_xml):
    """
    Find the root Grid (first Grid after Window.Resources) and add sidebar column.
    """
    # Find root Grid tag
    resources_end = content.rfind('</Window.Resources>')
    if resources_end == -1:
        resources_end = content.rfind('</Grid.Resources>')
    if resources_end == -1:
        resources_end = content.rfind('</WindowChrome.WindowChrome>')
    if resources_end == -1:
        return content, False

    rest = content[resources_end:]
    grid_match = re.search(r'<Grid\b', rest, re.IGNORECASE)
    if not grid_match:
        return content, False

    grid_start_abs = resources_end + grid_match.start()
    
    # Find the end of opening <Grid...> tag
    grid_tag_match = re.match(r'<Grid\b[^>]*>', content[grid_start_abs:], re.IGNORECASE)
    if not grid_tag_match:
        return content, False
        
    grid_body_start = grid_start_abs + grid_tag_match.end()
    
    # Find the first child visual element (any tag without a dot in its name, after opening Grid tag)
    child_match = re.search(r'<[a-zA-Z]+(?:\s|>|/)', content[grid_body_start:])
    first_child_idx = grid_body_start + child_match.start() if child_match else len(content)
    
    # 1. Ensure ColumnDefinitions exist on the root Grid, and prepend our 66px column
    col_defs_match = re.search(
        r'<Grid\.ColumnDefinitions>.*?</Grid\.ColumnDefinitions>',
        content[grid_body_start:],
        re.DOTALL | re.IGNORECASE
    )
    
    col_defs_present = False
    if col_defs_match:
        col_defs_start_abs = grid_body_start + col_defs_match.start()
        if col_defs_start_abs < first_child_idx:
            col_defs_present = True
            
    if col_defs_present:
        # Prepend sidebar column
        existing_cols = col_defs_match.group(0)
        if '66' in existing_cols:
            return content, False # already has it
            
        col_def_start = grid_body_start + col_defs_match.start()
        col_def_end = grid_body_start + col_defs_match.end()
        
        new_col_defs = existing_cols.replace(
            '<Grid.ColumnDefinitions>',
            '<Grid.ColumnDefinitions>\n                <ColumnDefinition Width="66"/>    <!-- Col 0: Sidebar -->'
        )
        
        # Shift all Grid.Column in the rest of the file
        body = content[col_def_end:]
        body = re.sub(r'Grid\.Column="(\d+)"', lambda m: f'Grid.Column="{int(m.group(1))+1}"', body)
        content = content[:col_def_start] + new_col_defs + body
    else:
        # Create new ColumnDefinitions and insert right at grid start
        col_defs = """
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="66"/>    <!-- Col 0: Sidebar -->
                <ColumnDefinition Width="*"/>     <!-- Col 1: Content -->
            </Grid.ColumnDefinitions>
"""
        # Shift Grid.Column in the rest of the file
        body = content[grid_body_start:]
        body = re.sub(r'Grid\.Column="(\d+)"', lambda m: f'Grid.Column="{int(m.group(1))+1}"', body)
        content = content[:grid_body_start] + col_defs + body
        
    # 2. Find insertion point for the sidebar visual child
    # Re-calculate grid_start_abs
    resources_end = content.rfind('</Window.Resources>')
    if resources_end == -1:
        resources_end = content.rfind('</Grid.Resources>')
    if resources_end == -1:
        resources_end = content.rfind('</WindowChrome.WindowChrome>')
    rest = content[resources_end:]
    grid_match = re.search(r'<Grid\b', rest, re.IGNORECASE)
    grid_start_abs = resources_end + grid_match.start()
    
    grid_tag_match = re.match(r'<Grid\b[^>]*>', content[grid_start_abs:], re.IGNORECASE)
    grid_body_start = grid_start_abs + grid_tag_match.end()
    
    col_end_match = re.search(r'</Grid\.ColumnDefinitions>', content[grid_body_start:], re.IGNORECASE)
    row_end_match = re.search(r'</Grid\.RowDefinitions>', content[grid_body_start:], re.IGNORECASE)
    
    col_end = grid_body_start + col_end_match.end() if col_end_match else -1
    row_end = grid_body_start + row_end_match.end() if row_end_match else -1
    
    if col_end == -1 and row_end == -1:
        insert_at = grid_body_start
    else:
        insert_at = max(col_end, row_end)
        
    # Insert sidebar XML
    content = content[:insert_at] + sidebar_xml + content[insert_at:]
    return content, True


def process_file(path):
    filename = os.path.basename(path)
    with io.open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    original = content

    # Skip if already has sidebar
    if 'Grid.Column="0" Grid.RowSpan' in content and 'T3SidebarToggle' in content:
        print(f"SKIP (already has sidebar): {filename}")
        return False

    tabs = extract_tab_radiobutttons(content)
    if not tabs:
        return False  # no tab RadioButtons

    print(f"PROCESSING: {filename} — found {len(tabs)} tabs: {[t['text'] for t in tabs]}")

    content, sidebar_xml = inject_sidebar(content, tabs, filename)
    content, ok = add_column_def_to_root_grid(content, sidebar_xml)

    if not ok:
        print(f"  WARNING: Could not inject sidebar for {filename}")
        return False

    with io.open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"  DONE: sidebar injected.")
    return True


def main():
    target_files = sorted(glob.glob(os.path.join(TOOLS_DIR, "*.xaml")))
    changed = 0
    for path in target_files:
        if process_file(path):
            changed += 1
    print(f"\n--- SIDEBAR NAV COMPLETE: {changed} files updated ---")


if __name__ == "__main__":
    main()
