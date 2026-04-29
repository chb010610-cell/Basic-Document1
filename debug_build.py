# -*- coding: utf-8 -*-
import sys, os, io, json, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r'C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Lib\site-packages')

# Import server module's functions
sys.path.insert(0, r'C:\Users\Administrator\WorkBuddy\20260429103845')

# Read and exec server.py to get ORIGINAL_VALUES
exec(open(r'C:\Users\Administrator\WorkBuddy\20260429103845\server.py', encoding='utf-8').read().split('def main')[0])

# Now test build_replace_map
data = {
    "company_name": "\u6e29\u5dde\u6d4b\u8bd5\u7535\u5668\u6709\u9650\u516c\u53f8",
    "doc_date": "2026-07-01",
    "version_no": "C/0",
    "iqc_person": "\u9648\u8fdb\u6599",
}
replacements = build_replace_map(data)
print(f"Number of replacement pairs: {len(replacements)}")
for old, new in replacements:
    print(f"  '{old}' -> '{new}'")

if len(replacements) == 0:
    print("\n[DEBUG] Checking why no replacements...")
    print(f"  company_name in form: '{data.get('company_name')}'")
    print(f"  ov['company_name']: '{ORIGINAL_VALUES['company_name']}'")
    print(f"  They are equal: {ORIGINAL_VALUES['company_name'] == data.get('company_name', '')}")
