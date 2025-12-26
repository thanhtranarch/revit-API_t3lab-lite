# -*- coding: utf-8 -*-
"""
Fix All Formatting Issues

Batch update all Python files to standard format.

Author: Tran Tien Thanh
"""

import os
import re

def fix_author(content):
    """Fix author attribution"""
    content = re.sub(
        r'__author__\s*=\s*["\']T3Lab["\']',
        '__author__ = "Tran Tien Thanh"',
        content
    )
    return content

def fix_ascii_headers(content):
    """Replace ASCII art section headers"""

    # IMPORTS section
    content = re.sub(
        r'#\s*[╦╣║╩╔╗═╠╝]+.*?IMPORTS.*?={20,}',
        '# IMPORT LIBRARIES\n# ==================================================',
        content,
        flags=re.DOTALL
    )

    # VARIABLES section
    content = re.sub(
        r'#\s*[╦╣║╩╔╗═╠╝]+.*?VARIABLES.*?={20,}',
        '# DEFINE VARIABLES\n# ==================================================',
        content,
        flags=re.DOTALL
    )

    # FUNCTIONS section
    content = re.sub(
        r'#\s*[╦╣║╩╔╗═╠╝]+.*?FUNCTIONS.*?={20,}',
        '# HELPER FUNCTIONS\n# ==================================================',
        content,
        flags=re.DOTALL
    )

    # CLASSES section
    content = re.sub(
        r'#\s*[╦╣║╩╔╗═╠╝]+.*?CLASSES.*?={20,}',
        '# CLASSES\n# ==================================================',
        content,
        flags=re.DOTALL
    )

    # MAIN section
    content = re.sub(
        r'#\s*[╦╣║╩╔╗═╠╝]+.*?MAIN.*?={20,}',
        '# MAIN SCRIPT\n# ==================================================',
        content,
        flags=re.DOTALL
    )

    return content

def add_author_if_missing(content, filename):
    """Add author info if missing from docstring"""
    # Check if already has author info
    if 'Author: Tran Tien Thanh' in content:
        return content

    # Find docstring
    docstring_match = re.search(r'(""".*?""")', content, re.DOTALL)
    if docstring_match:
        docstring = docstring_match.group(1)

        # Add author info before closing """
        new_docstring = docstring.rstrip('"""').rstrip() + '\n\n'
        new_docstring += 'Author: Tran Tien Thanh\n'
        new_docstring += 'Mail: trantienthanh909@gmail.com\n'
        new_docstring += 'Linkedin: linkedin.com/in/sunarch7899/\n'
        new_docstring += '"""'

        content = content.replace(docstring, new_docstring)

    return content

def process_file(filepath):
    """Process a single Python file"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        original = content

        # Apply fixes
        content = fix_author(content)
        content = fix_ascii_headers(content)
        content = add_author_if_missing(content, os.path.basename(filepath))

        # Write back if changed
        if content != original:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        return False

    except Exception as ex:
        print("Error processing {}: {}".format(filepath, ex))
        return False

def main():
    """Main entry point"""
    base_dir = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'T3Lab_Lite.extension'
    )

    updated_count = 0
    total_count = 0

    print("Scanning for Python files...")

    for root, dirs, files in os.walk(base_dir):
        # Skip __pycache__
        if '__pycache__' in root:
            continue

        for file in files:
            if file.endswith('.py'):
                filepath = os.path.join(root, file)
                total_count += 1

                if process_file(filepath):
                    updated_count += 1
                    print("Updated: {}".format(filepath))

    print("\n" + "="*60)
    print("Total files scanned: {}".format(total_count))
    print("Files updated: {}".format(updated_count))
    print("="*60)

if __name__ == '__main__':
    main()
