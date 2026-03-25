"""打印 Checklist 模板中各节的参数名，帮助对照参数提取情况"""
import sys, os
sys.path.insert(0, 'backend')
import openpyxl

xlsx_path = os.path.join('backend', '..', 'Checklist模板.xlsx')
wb = openpyxl.load_workbook(xlsx_path)
ws = wb.active

current_section = '?'
for row in ws.iter_rows(values_only=True):
    cell0 = str(row[0]).strip() if row[0] else ''
    cell1 = str(row[1]).strip() if row[1] else ''
    # Detect section header rows (numbered like "1.", "8.", "15." etc.)
    if cell0 and cell0[0].isdigit() and '.' not in cell0[1:3]:
        # It's a section number row
        current_section = cell0 + (' ' + cell1 if cell1 else '')
        print(f'\n=== {current_section} ===')
    elif cell1:
        print(f'  {cell1}')
