import os
import io

repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
showcase_path = os.path.join(repo_root, ".claude", "standard", "UIStandardShowcase.xaml")
wpf_styles_path = os.path.join(repo_root, "T3Lab.extension", "lib", "GUI", "Resources", "WPF_styles.xaml")

# Read UIStandardShowcase.xaml
with io.open(showcase_path, "r", encoding="utf-8") as f:
    showcase_content = f.read()

# Extract styles from showcase
start_marker = "<Window.Resources>"
end_marker = "<!-- ═══ END T3LAB SHARED STYLES ═══ -->"

start_idx = showcase_content.find(start_marker)
if start_idx == -1:
    print("Could not find <Window.Resources> in showcase")
    exit(1)
start_idx += len(start_marker)

end_idx = showcase_content.find(end_marker, start_idx)
if end_idx == -1:
    print("Could not find end marker in showcase")
    exit(1)

extracted_styles = showcase_content[start_idx:end_idx].strip()

# Read WPF_styles.xaml
with io.open(wpf_styles_path, "r", encoding="utf-8") as f:
    wpf_content = f.read()

# Find the sync block in WPF_styles.xaml
begin_sync = "<!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT (edit lib/GUI/Resources/WPF_styles.xaml, then run dev/sync_wpf_styles.py) ═══ -->"
end_sync = "<!-- ═══ END T3LAB SHARED STYLES ═══ -->"

begin_sync_idx = wpf_content.find(begin_sync)
if begin_sync_idx == -1:
    print("Could not find begin sync block in WPF_styles.xaml")
    exit(1)

end_sync_idx = wpf_content.find(end_sync, begin_sync_idx)
if end_sync_idx == -1:
    print("Could not find end sync block in WPF_styles.xaml")
    exit(1)

# Construct new content
new_wpf_content = (
    wpf_content[:begin_sync_idx] +
    begin_sync + "\n        " +
    extracted_styles + "\n        " +
    end_sync +
    wpf_content[end_sync_idx + len(end_sync):]
)

with io.open(wpf_styles_path, "w", encoding="utf-8") as f:
    f.write(new_wpf_content)

print("Successfully updated WPF_styles.xaml from UIStandardShowcase.xaml")
