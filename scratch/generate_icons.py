import os

output_dir = "icons_export"
os.makedirs(output_dir, exist_ok=True)

icons = {
    "T3LabAssistant.svg": """<svg width="64" height="64" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <path d="M32 8 C32 24 24 32 8 32 C24 32 32 40 32 56 C32 40 40 32 56 32 C40 32 32 24 32 8 Z" fill="#3B82F6" opacity="0.8"/>
  <path d="M32 16 C32 26 26 32 16 32 C26 32 32 38 32 48 C32 38 38 32 48 32 C38 32 32 26 32 16 Z" fill="#0F172A"/>
  <circle cx="48" cy="16" r="4" fill="#3B82F6" />
</svg>""",
    "Feedback.svg": """<svg width="64" height="64" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <path d="M12 16 C12 11.58 15.58 8 20 8 L44 8 C48.42 8 52 11.58 52 16 L52 36 C52 40.42 48.42 44 44 44 L28 44 L16 54 L16 44 C13.79 44 12 42.21 12 40 Z" fill="none" stroke="#0F172A" stroke-width="4" stroke-linejoin="round"/>
  <line x1="20" y1="20" x2="36" y2="20" stroke="#0F172A" stroke-width="4" stroke-linecap="round"/>
  <line x1="20" y1="28" x2="44" y2="28" stroke="#0F172A" stroke-width="4" stroke-linecap="round"/>
  <circle cx="52" cy="12" r="6" fill="#F59E0B" />
</svg>""",
    "MCPControl.svg": """<svg width="64" height="64" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <rect x="8" y="12" width="48" height="40" rx="4" fill="none" stroke="#0F172A" stroke-width="4"/>
  <line x1="8" y1="24" x2="56" y2="24" stroke="#0F172A" stroke-width="4"/>
  <polyline points="14,34 20,40 14,46" fill="none" stroke="#3B82F6" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
  <line x1="26" y1="46" x2="36" y2="46" stroke="#3B82F6" stroke-width="4" stroke-linecap="round"/>
  <circle cx="16" cy="18" r="2" fill="#0F172A"/>
  <circle cx="24" cy="18" r="2" fill="#0F172A"/>
</svg>""",
    "PDF_Import.svg": """<svg width="64" height="64" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <path d="M16 8 L36 8 L48 20 L48 44" fill="none" stroke="#0F172A" stroke-width="4" stroke-linejoin="round"/>
  <path d="M16 8 L16 56 C16 58.2 17.8 60 20 60 L44 60" fill="none" stroke="#0F172A" stroke-width="4" stroke-linejoin="round"/>
  <path d="M36 8 L36 20 L48 20" fill="none" stroke="#0F172A" stroke-width="4" stroke-linejoin="round"/>
  <path d="M26 26 L26 40 M20 34 L26 40 L32 34" fill="none" stroke="#0F172A" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
  <rect x="22" y="46" width="36" height="16" rx="4" fill="#EF4444" />
  <text x="40" y="58" font-family="Arial, sans-serif" font-size="12" font-weight="bold" fill="white" text-anchor="middle">PDF</text>
</svg>""",
    "ACC_Platform.svg": """<svg width="64" height="64" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <path d="M16 44 C10 44 8 38 12 34 C12 24 24 20 32 24 C38 16 52 20 52 32 C58 34 54 44 48 44 Z" fill="none" stroke="#0F172A" stroke-width="4" stroke-linejoin="round"/>
  <path d="M32 26 L26 38 M32 26 L38 38 M28 34 L36 34" fill="none" stroke="#3B82F6" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
</svg>""",
    "Cloud_Health.svg": """<svg width="64" height="64" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">
  <path d="M16 44 C10 44 8 38 12 34 C12 24 24 20 32 24 C38 16 52 20 52 32 C58 34 54 44 48 44 Z" fill="none" stroke="#0F172A" stroke-width="4" stroke-linejoin="round"/>
  <circle cx="44" cy="44" r="12" fill="#F8FAFC" />
  <circle cx="44" cy="44" r="9" fill="#10B981" />
  <path d="M40 44 L43 47 L49 41" fill="none" stroke="#FFFFFF" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
</svg>"""
}

for name, content in icons.items():
    with open(os.path.join(output_dir, name), "w", encoding="utf-8") as f:
        f.write(content)

print(f"Created {len(icons)} SVG icons in {os.path.abspath(output_dir)}")
