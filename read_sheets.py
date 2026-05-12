import openpyxl
import json

wb = openpyxl.load_workbook('Ramais Atualizado 1601.xlsx')
sheet = wb.active

print('=== PLANILHA ESTRUTURA ===')
for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=100, values_only=True), 1):
    if any(row):
        col1 = str(row[0]).strip() if row[0] else ''
        col2 = str(row[1]).strip() if row[1] else ''
        col3 = str(row[2]).strip() if row[2] else ''
        col4 = str(row[3]).strip() if row[3] else ''
        print(f'{i:2d}: {col1:35s} | {col2:25s} | {col3:10s} | {col4}')
