# -*- coding: utf-8 -*-
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r'C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Lib\site-packages')

# Manually copy ORIGINAL_VALUES and build_replace_map from server.py
ORIGINAL_VALUES = {
    "company_name": "\u5b81\u6ce2\u5e02\u6d77\u66ae\u65b0\u827a\u6d01\u5177\u6709\u9650\u516c\u53f8",
    "doc_date": "2025-03-05",
    "version_no": "B/0",
    "header_version": "A/0",
    "iqc_person": "\u6234\u6d77\u80fd",
    "iqc_exam_person": "\u674e\u5fe0\u4f1f",
}

data = {
    "company_name": "\u6e29\u5dde\u6d4b\u8bd5\u7535\u5668\u6709\u9650\u516c\u53f8",
    "doc_date": "2026-07-01",
    "version_no": "C/0",
    "iqc_person": "\u9648\u8fdb\u6599",
}

def build_replace_map(form_data):
    replacements = []
    ov = ORIGINAL_VALUES
    def add(old, new):
        if old and new and old != new:
            replacements.append((old, new))

    new_date = form_data.get("doc_date", "").strip()
    if ov["doc_date"] and new_date and ov["doc_date"] != new_date:
        add(ov["doc_date"], new_date)
        print(f"  Added date replacement")

    new_ver = form_data.get("version_no", "").strip()
    if ov["version_no"] and new_ver and ov["version_no"] != new_ver:
        add(ov["version_no"], new_ver)
        print(f"  Added version_no replacement")
    if ov["header_version"] and new_ver and ov["header_version"] != new_ver:
        add(ov["header_version"], new_ver)
        print(f"  Added header_version replacement")

    add(ov["company_name"], form_data.get("company_name", ""))
    add(ov["iqc_person"], form_data.get("iqc_person", ""))
    add(ov["iqc_exam_person"], form_data.get("iqc_person", ""))

    return replacements

replacements = build_replace_map(data)
print(f"\nTotal replacement pairs: {len(replacements)}")
for old, new in replacements:
    print(f"  '{old}' -> '{new}'")
