import os
import re
import io

def parse_simple_yaml(path):
    data = {}
    if not os.path.exists(path):
        return data
        
    current_key = None
    with io.open(path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    for line in lines:
        if line.startswith('  - '):
            if current_key:
                if not isinstance(data[current_key], list):
                    data[current_key] = []
                data[current_key].append(line[4:].strip())
        elif line.startswith('  '):
            if current_key and isinstance(data[current_key], str):
                data[current_key] += '\n' + line[2:].rstrip('\n')
        else:
            match = re.match(r'^([a-zA-Z_]+):\s*(.*)$', line)
            if match:
                key = match.group(1).strip()
                val = match.group(2).strip()
                
                if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                    val = val[1:-1]
                    
                if key == 'title':
                    val = val.replace('\\n', '\n')
                    
                if val == '|' or val == '>':
                    data[key] = ''
                elif val == '':
                    data[key] = []
                else:
                    data[key] = val
                current_key = key
    
    for k in data:
        if isinstance(data[k], str):
            data[k] = data[k].strip()
            
    return data

def get_script_metadata(script_path):
    metadata = {'author': 'T3Lab', 'version': '1.0.0', 'tooltip': '', 'description': ''}
    if not os.path.exists(script_path):
        return metadata
    
    with io.open(script_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
        author_match = re.search(r'__author__\s*=\s*[\'"]([^\'"]+)[\'"]', content)
        if author_match:
            metadata['author'] = author_match.group(1)
            
        version_match = re.search(r'__version__\s*=\s*[\'"]([^\'"]+)[\'"]', content)
        if version_match:
            metadata['version'] = version_match.group(1)
            
        docstring_match = re.search(r'\"\"\"(.*?)\"\"\"', content, re.DOTALL)
        if docstring_match:
            docstring = docstring_match.group(1).strip()
            lines = docstring.split('\n')
            if len(lines) > 0:
                metadata['tooltip'] = lines[0].strip()
                if len(lines) > 1:
                    metadata['description'] = '\n'.join([l.strip() for l in lines[1:] if l.strip() != ''])
                else:
                    metadata['description'] = metadata['tooltip']
                    
    return metadata

def update_pushbutton_bundle(bundle_path):
    script_path = os.path.join(os.path.dirname(bundle_path), 'script.py')
    metadata = get_script_metadata(script_path)
    data = parse_simple_yaml(bundle_path)
        
    new_data = {}
    
    # 1. Title
    if 'title' in data and data['title']:
        new_data['title'] = data['title']
    else:
        new_data['title'] = os.path.basename(os.path.dirname(bundle_path)).split('.')[0]
        
    # 2. Tooltip
    if 'tooltip' in data and data['tooltip']:
        new_data['tooltip'] = data['tooltip']
    else:
        new_data['tooltip'] = metadata['tooltip'] or new_data['title']
        
    # 3. Description
    if 'description' in data and data['description']:
        new_data['description'] = data['description']
    else:
        new_data['description'] = metadata['description'] or new_data['tooltip']
        
    if isinstance(new_data['description'], list):
        new_data['description'] = '\n'.join(new_data['description'])
        
    # 4. Author
    if 'author' in data and data['author'] and data['author'] != []:
        new_data['author'] = data['author']
    else:
        new_data['author'] = metadata['author']
        
    # 5. Version
    if 'version' in data and data['version'] and data['version'] != []:
        new_data['version'] = data['version']
    else:
        new_data['version'] = metadata['version']
        
    # Keep context and hyperlink
    for k in ['context', 'hyperlink']:
        if k in data and data[k] != []:
            new_data[k] = data[k]
        
    write_yaml(bundle_path, new_data)

def update_layout_bundle(bundle_path):
    data = parse_simple_yaml(bundle_path)
    if not data:
        return
        
    new_data = {}
    if 'title' in data and data['title']:
        new_data['title'] = data['title']
    if 'layout' in data and isinstance(data['layout'], list):
        new_data['layout'] = data['layout']
        
    write_yaml(bundle_path, new_data)

def write_yaml(path, data):
    with io.open(path, 'w', encoding='utf-8') as f:
        if 'title' in data:
            title = data['title'].replace('\n', '\\n')
            f.write(u'title: "{}"\n'.format(title))
            
        if 'tooltip' in data:
            t = str(data['tooltip']).replace('"', '\\"')
            f.write(u'tooltip: "{}"\n'.format(t))
            
        if 'description' in data:
            f.write(u'description: |\n')
            for line in str(data['description']).split('\n'):
                f.write(u'  {}\n'.format(line.replace('\r', '')))
                
        if 'author' in data:
            f.write(u'author: "{}"\n'.format(data['author']))
            
        if 'version' in data:
            f.write(u'version: "{}"\n'.format(data['version']))
            
        if 'context' in data:
            if isinstance(data['context'], list):
                f.write(u'context:\n')
                for item in data['context']:
                    f.write(u'  - {}\n'.format(item))
            else:
                f.write(u'context: "{}"\n'.format(data['context']))
                
        if 'hyperlink' in data:
            f.write(u'hyperlink: "{}"\n'.format(data['hyperlink']))
                
        if 'layout' in data:
            f.write(u'layout:\n')
            for item in data['layout']:
                f.write(u'  - {}\n'.format(item))

# Main
root_dir = 'T3Lab.extension'
for dirpath, dirnames, filenames in os.walk(root_dir):
    if 'bundle.yaml' in filenames:
        bundle_path = os.path.join(dirpath, 'bundle.yaml')
        dirname = os.path.basename(dirpath)
        if dirname.endswith('.pushbutton') or dirname.endswith('.urlbutton'):
            update_pushbutton_bundle(bundle_path)
        elif dirname.endswith('.panel') or dirname.endswith('.tab') or dirname.endswith('.pulldown') or dirname.endswith('.stack'):
            update_layout_bundle(bundle_path)

print("All bundles updated!")
