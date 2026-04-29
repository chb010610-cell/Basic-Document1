import xlrd
import json
import sys

wb = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\导入.xls")
sheet = wb.sheet_by_index(0)
print(f"Sheet 0: {wb.sheet_names()[0]}")
print(f"Rows: {sheet.nrows}, Cols: {sheet.ncols}")

# Get headers from row 0
headers = []
for col_idx in range(sheet.ncols):
    val = sheet.cell_value(0, col_idx)
    # Try to fix encoding - the data appears to be UTF-8 read as GBK/GB2312
    try:
        fixed = val.encode('latin-1').decode('gbk')
    except:
        try:
            fixed = val.encode('latin-1').decode('utf-8')
        except:
            fixed = val
    headers.append(fixed)

print("Headers:", json.dumps(headers, ensure_ascii=False))

# Get row 1 data
row1 = []
for col_idx in range(sheet.ncols):
    val = sheet.cell_value(1, col_idx)
    try:
        fixed = val.encode('latin-1').decode('gbk')
    except:
        try:
            fixed = val.encode('latin-1').decode('utf-8')
        except:
            fixed = val
    row1.append(fixed)
    
print("Row 1:", json.dumps(row1, ensure_ascii=False))