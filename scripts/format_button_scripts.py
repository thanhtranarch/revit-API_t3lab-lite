# -*- coding: utf-8 -*-
"""
Batch Format Button Scripts

Updates all button scripts to the standard format.

Author: Tran Tien Thanh
"""

import os
import re

# Standard header template
HEADER_TEMPLATE = '''# -*- coding: utf-8 -*-
"""
{title}

{description}

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__  = "{title}"
'''

def update_script_format(script_path):
    """Update a single script to standard format"""
    try:
        with open(script_path, 'r') as f:
            content = f.read()

        # Extract title from existing __title__
        title_match = re.search(r'__title__\s*=\s*["\']([^"\']+)["\']', content)
        if not title_match:
            print("Skipping {} - no title found".format(script_path))
            return False

        title = title_match.group(1)

        # Extract description from docstring
        doc_match = re.search(r'"""([^"]+)"""', content)
        description = doc_match.group(1).strip() if doc_match else title

        # Replace ASCII art section headers
        content = re.sub(
            r'#\s*[╦╣║╩╔╗═╠╝]+.*?IMPORTS.*?={40,}',
            '# IMPORT LIBRARIES\n# ==================================================',
            content,
            flags=re.DOTALL
        )

        content = re.sub(
            r'#\s*[╦╣║╩╔╗═╠╝]+.*?MAIN.*?={40,}',
            '# MAIN SCRIPT\n# ==================================================',
            content,
            flags=re.DOTALL
        )

        # Replace header (everything before first import)
        import_match = re.search(r'^(import |from )', content, re.MULTILINE)
        if import_match:
            import_start = import_match.start()
            # Find the section header before imports
            header_end = content.rfind('# IMPORT', 0, import_start)
            if header_end == -1:
                header_end = content.rfind('#', 0, import_start)
            if header_end == -1:
                header_end = import_start

            new_header = HEADER_TEMPLATE.format(
                title=title,
                description=description
            )

            content = new_header + '\n' + content[header_end:]

        # Write back
        with open(script_path, 'w') as f:
            f.write(content)

        print("Updated: {}".format(script_path))
        return True

    except Exception as ex:
        print("Error updating {}: {}".format(script_path, ex))
        return False

def main():
    """Main entry point"""
    # Find all script.py files in T3Lab_Lite.tab
    base_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'T3Lab_Lite.extension', 'T3Lab_Lite.tab')

    count = 0
    for root, dirs, files in os.walk(base_dir):
        if 'script.py' in files:
            script_path = os.path.join(root, 'script.py')
            if update_script_format(script_path):
                count += 1

    print("\n" + "="*60)
    print("Updated {} button scripts".format(count))
    print("="*60)

if __name__ == '__main__':
    main()
