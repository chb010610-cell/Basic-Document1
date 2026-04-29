"""Read all sheets from XLS and print contents"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try:
    import xlrd
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlrd", "--quiet"])
    import xlrd

xls_path = r"C:\Users\Administrator\Desktop\导入.xls"
wb = xlrd.open_workbook(xls_path)

print(f"Total sheets: {wb.nsheets}")
print(f"Sheet names: {wb.sheet_names()}")
print("="*80)

for si in range(wb.nsheets):
    sheet = wb.sheet_by_index(si)
    print(f"\n{'='*80}")
    print(f"Sheet {si}: {sheet.name}")
    print(f"Rows: {sheet.nrows}, Cols: {sheet.ncols}")
    print("-"*80)

    if sheet.nrows > 0:
        # Print headers
        headers = []
        for c in range(sheet.ncols):
            val = sheet.cell_value(0, c)
            try:
                val = val.encode('latin-1').decode('gbk')
            except:
                pass
            headers.append(str(val))
        print("Headers:", headers)

        # Print first data row
        if sheet.nrows > 1:
            print("\nFirst data row:")
            for c in range(sheet.ncols):
                val = sheet.cell_value(1, c)
                try:
                    val = val.encode('latin-1').decode('gbk')
                except:
                    pass
                print(f"  Col{c}: {headers[c]} = {val}")

        # Print all data rows if few
        if sheet.nrows <= 5:
            print(f"\nAll rows:")
            for r in range(sheet.nrows):
                row_data = []
                for c in range(sheet.ncols):
                    val = sheet.cell_value(r, c)
                    try:
                        val = val.encode('latin-1').decode('gbk')
                    except:
                        pass
                    row_data.append(str(val))
                print(f"  Row{r}: {row_data}")
