import os
import io
import re

tools_dir = r"T3Lab.extension/lib/GUI/Tools"
xaml_files = [os.path.join(tools_dir, f) for f in os.listdir(tools_dir) if f.endswith('.xaml')]

for path in sorted(xaml_files):
    with io.open(path, "r", encoding="utf-8") as f:
        content = f.read()
    
    # Remove Window.Resources
    content_no_res = re.sub(r'<Window\.Resources>.*?</Window\.Resources>', '', content, flags=re.DOTALL)
    content_no_res = re.sub(r'<WindowChrome\.WindowChrome>.*?</WindowChrome\.WindowChrome>', '', content_no_res, flags=re.DOTALL)
    
    # Find the first element tag under Window
    match = re.search(r'<Window\b.*?>\s*(<(\w+)\b)', content_no_res, flags=re.DOTALL)
    if not match:
        # Try finding the first tag after Window.Resources and WindowChrome
        first_tag_match = re.search(r'</Window\.Resources>\s*<(\w+)\b', content, flags=re.DOTALL)
        if first_tag_match:
            print(f"{os.path.basename(path)}: {first_tag_match.group(1)} (after resources)")
        else:
            print(f"{os.path.basename(path)}: No root tag found")
    else:
        print(f"{os.path.basename(path)}: {match.group(2)}")
