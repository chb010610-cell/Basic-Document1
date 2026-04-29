# -*- coding: utf-8 -*-
"""直接用 server.py 模块做替换测试，绕过 HTTP"""
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r'C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Lib\site-packages')

# Force reimport
if 'server' in sys.modules:
    del sys.modules['server']

# Add cwd
sys.path.insert(0, r'C:\Users\Administrator\WorkBuddy\20260429103845')

import importlib
import server
importlib.reload(server)

print(f"ORIGINAL_VALUES keys: {list(server.ORIGINAL_VALUES.keys())}")
print(f"Has iqc_exam_person: {'iqc_exam_person' in server.ORIGINAL_VALUES}")
print(f"Has header_version: {'header_version' in server.ORIGINAL_VALUES}")
print(f"iqc_exam_person value: {server.ORIGINAL_VALUES.get('iqc_exam_person', 'NOT FOUND')}")

# Test build_replace_map
data = {
    "company_name": "温州测试电器有限公司",
    "doc_date": "2026-07-01",
    "version_no": "C/0",
    "iqc_person": "陈进料",
}
replacements = server.build_replace_map(data)
print(f"\nReplacements count: {len(replacements)}")
for old, new in replacements:
    print(f"  '{old}' -> '{new}'")
