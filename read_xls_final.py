# -*- coding: utf-8 -*-
"""Read import.xls to understand the data structure and find all replaceable values."""
import os, sys
os.environ["PYTHONUTF8"] = "1"

XLS_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "导入.xls")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xls_analysis.txt")

try:
    import xlrd
    wb = xlrd.open_workbook(XLS_PATH)
    lines = []
    lines.append(f"Sheets: {wb.sheet_names()}")
    
    for si in range(wb.nsheets):
        sh = wb.sheet_by_index(si)
        lines.append(f"\n=== Sheet {si}: {sh.name} ({sh.nrows} rows x {sh.ncols} cols) ===")
        for ri in range(sh.nrows):
            row_data = []
            for ci in range(sh.ncols):
                val = sh.cell_value(ri, ci)
                row_data.append(str(val).strip() if val else "")
            # Only print rows that have some content
            if any(v for v in row_data):
                lines.append(f"Row {ri}: {row_data}")
    
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Done. Lines: {len(lines)}")

except Exception as e:
    import traceback
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(f"Error: {e}\n{traceback.format_exc()}")
    print(f"Error: {e}")
