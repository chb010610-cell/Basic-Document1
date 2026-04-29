import xlrd
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\导入.xls", encoding_override='gbk')
for sheet_idx in range(wb.nsheets):
    sheet = wb.sheet_by_index(sheet_idx)
    print(f"=== Sheet: {sheet.name} ===")
    print(f"Rows: {sheet.nrows}, Cols: {sheet.ncols}")
    for row_idx in range(min(sheet.nrows, 5)):
        row_data = []
        for col_idx in range(sheet.ncols):
            cell = sheet.cell(row_idx, col_idx)
            row_data.append(str(cell.value))
        print(f"Row {row_idx}: {json.dumps(row_data, ensure_ascii=False)}")