import os

xaml_path = r"c:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\T3Lab.extension\lib\GUI\Tools\T3LabAssistant.xaml"

with open(xaml_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Make sure we don't replace inside the synced block.
# Actually, the synced block doesn't contain these slate colors because they are specific to Assistant.
replacements = {
    "#0F172A": "#18181B",
    "#1E293B": "#27272A",
    "#64748B": "#71717A",
    "#94A3B8": "#A1A1AA",
    "#E2E8F0": "#E6E6EA",
    "#F8FAFC": "#F4F4F6",
    "#F1F5F9": "#ECECEF",
    "#CBD5E1": "#D4D4DA",
    "#EF4444": "#D23B3B"
}

for old, new in replacements.items():
    content = content.replace(old, new)

with open(xaml_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("UI Colors updated successfully.")
