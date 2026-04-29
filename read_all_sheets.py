import xlrd
import json
import sys

wb = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\导入.xls")
print("All sheets:", json.dumps(wb.sheet_names(), ensure_ascii=False))

def fix_enc(val):
    if isinstance(val, str):
        try:
            return val.encode('latin-1').decode('gbk')
        except:
            try:
                return val.encode('latin-1').decode('utf-8')
            except:
                return val
    return val

# Read all sheets - headers and first data row
for sheet_idx in range(wb.nsheets):
    sheet = wb.sheet_by_index(sheet_idx)
    name = fix_enc(sheet.name)
    print(f"\n=== Sheet {sheet_idx}: {name} ===")
    print(f"Rows: {sheet.nrows}, Cols: {sheet.ncols}")
    
    headers = []
    for col_idx in range(sheet.ncols):
        headers.append(fix_enc(sheet.cell_value(0, col_idx)))
    print(f"Headers: {json.dumps(headers, ensure_ascii=False)}")
    
    if sheet.nrows > 1:
        row1 = []
        for col_idx in range(sheet.ncols):
            row1.append(fix_enc(sheet.cell_value(1, col_idx)))
        print(f"Row 1: {json.dumps(row1, ensure_ascii=False)}")