# -*- coding: utf-8 -*-
import sys, os, io, json, urllib.request
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

data = {
    "company_name": "ABC",
    "doc_date": "2026-07-01",
    "version_no": "C/0",
    "iqc_person": "XYZ",
}

url = "http://127.0.0.1:8765/api/generate"
req = urllib.request.Request(url, data=json.dumps(data, ensure_ascii=False).encode('utf-8'),
                              headers={'Content-Type': 'application/json; charset=utf-8'})
print("Sending request with simple ASCII names...")
resp = urllib.request.urlopen(req, timeout=60)
result = json.loads(resp.read().decode('utf-8'))
print(f"Success: {result['success']}  Count: {result['count']}")
total = sum(g.get('replacements', 0) for g in result.get('generated', []))
print(f"Total replacements: {total}")
for g in result.get('generated', [])[:3]:
    print(f"  {g['name']}: {g.get('replacements', 0)} repl")
