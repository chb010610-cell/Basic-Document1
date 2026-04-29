import sys
sys.path.insert(0, r"C:\Users\Administrator\WorkBuddy\20260429103845")
try:
    from server import *
    print("Import OK, starting server...")
except Exception as e:
    print(f"Import error: {e}")
    import traceback
    traceback.print_exc()