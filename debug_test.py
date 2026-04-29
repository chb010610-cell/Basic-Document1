import subprocess, sys, os, time
os.environ["PYTHONUTF8"] = "1"

logfile = r"C:\Users\Administrator\WorkBuddy\20260429103845\debug_output.txt"

def log(msg):
    with open(logfile, "a", encoding="utf-8") as f:
        f.write(str(msg) + "\n")

# Start server as subprocess
proc = subprocess.Popen(
    [sys.executable, r"C:\Users\Administrator\WorkBuddy\20260429103845\server.py"],
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    encoding='utf-8',
    errors='replace'
)
time.sleep(3)

# Check if server is alive
if proc.poll() is not None:
    out = proc.stdout.read()
    log(f"SERVER DIED IMMEDIATELY: {out}")
else:
    log("Server seems to be running")

# Send test request
try:
    import urllib.request, json
    data = {"company_name": "测试公司", "doc_date": "2025-06-01", "version_no": "C/1", "iqc_person": "张三"}
    body = json.dumps(data).encode("utf-8")
    req = urllib.request.Request("http://127.0.0.1:8765/api/generate", data=body, headers={"Content-Type": "application/json"})
    resp = urllib.request.urlopen(req, timeout=30)
    result = resp.read().decode()
    log(f"API RESPONSE: {result}")
except Exception as e:
    log(f"API ERROR: {type(e).__name__}: {e}")

# Get all server output
time.sleep(2)
if proc.poll() is None:
    # Server still running - just read what we can
    # We need to kill it to get buffered output
    proc.terminate()
    time.sleep(1)

out = proc.stdout.read()
log(f"=== SERVER OUTPUT ===")
log(out)
log("=== END ===")
