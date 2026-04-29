import urllib.request, json, os

data = {
    "company_name": "测试公司有限公司",
    "address": "浙江省杭州市西湖区文三路",
    "product": "电子产品制造",
    "english_abbr": "CS",
    "gm_name": "张总",
    "production_mgr": "李生产",
    "mgr_rep": "王代表",
    "hr_head": "赵人事",
    "sales_head": "钱业务",
    "purchase_head": "孙采购",
    "tech_head": "周技术",
    "raw_warehouse": "吴仓管",
    "finished_warehouse": "郑成仓",
    "mechanic": "冯机修",
    "iqc_person": "陈检验",
    "ipqc1_person": "褚制程",
    "ipqc2_person": "卫制程",
    "ipqc3_person": "蒋制程",
    "fqc_person": "沈成品",
    "doc_date": "2025-06-15",
    "version_no": "C/1",
}

body = json.dumps(data).encode("utf-8")
req = urllib.request.Request("http://127.0.0.1:8765/api/generate", data=body, headers={"Content-Type": "application/json"})
resp = urllib.request.urlopen(req, timeout=60)
result = json.loads(resp.read().decode())

out = r"C:\Users\Administrator\WorkBuddy\20260429103845\test_gen_result.json"
with open(out, "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"SUCCESS: {result['success']}, count={result['count']}, total_replacements={result.get('total_replacements',0)}")
for g in result.get("generated", []):
    status = "OK" if "error" not in g else f"ERR:{g.get('error','')}"
    print(f"  {g['name']}: {status} (repl={g.get('replacements',0)})")
