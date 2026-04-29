import urllib.request, json, os, sys, io
os.environ["PYTHONUTF8"] = "1"
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)

data = {
    "company_name": "宁波市金狼电器有限公司",
    "address": "浙江省余姚市梨洲街道新墅村叶岙路东111号",
    "product": "金属笔",
    "english_abbr": "KL",
    "gm_name": "史迪华",
    "production_mgr": "陈啸",
    "mgr_rep": "黄金金",
    "hr_head": "干映明",
    "sales_head": "张灿灿",
    "purchase_head": "韩曼",
    "tech_head": "潘开东",
    "raw_warehouse": "王娟",
    "finished_warehouse": "章伟勇",
    "mechanic": "张国龙",
    "iqc_person": "阮建良",
    "doc_date": "2025-04-20",
    "version_no": "B/0",
    "license_code": "",
    "supplier": "",
}

body = json.dumps(data).encode("utf-8")
req = urllib.request.Request("http://127.0.0.1:8765/api/generate", data=body, headers={"Content-Type": "application/json"})
try:
    resp = urllib.request.urlopen(req, timeout=120)
    result = json.loads(resp.read().decode())
    print(f"SUCCESS: {result['success']}, count={result['count']}, total_replacements={result.get('total_replacements',0)}")
    for g in result.get("generated", []):
        status = "OK" if "error" not in g else f"ERR:{g.get('error','')}"
        print(f"  {g['name']}: {status} (repl={g.get('replacements',0)})")
except Exception as e:
    print(f"ERROR: {e}")
