# -*- coding: utf-8 -*-
import re

def check_xaml(filepath):
    print("\nChecking: {}".format(filepath))
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Find x:Key or x:Key=...
    keys = re.findall(r'x:Key\s*=\s*"([^"]+)"', content)
    seen = set()
    dups = []
    for k in keys:
        if k in seen:
            dups.append(k)
        seen.add(k)
        
    print("Total keys found: {}".format(len(keys)))
    print("Duplicates found: {}".format(dups))

if __name__ == '__main__':
    check_xaml('T3Lab.extension/lib/GUI/Tools/ContainsManager.xaml')
    check_xaml('T3Lab.extension/lib/GUI/Tools/FamilyManager.xaml')
    check_xaml('T3Lab.extension/lib/GUI/Tools/ParameterManager.xaml')
    check_xaml('T3Lab.extension/lib/GUI/Tools/ScheduleManager.xaml')
