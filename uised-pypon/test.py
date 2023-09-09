import json

import openpyxl
import string

wb = openpyxl.load_workbook('ari.xlsx')
sheet = wb['Лист1']

for rng in sheet.merged_cells.ranges:
    if 'D68' in rng:
        print(type(rng))
        for row in rng.rows:
            return len(row)