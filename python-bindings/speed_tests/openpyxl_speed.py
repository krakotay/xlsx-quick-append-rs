import openpyxl
import os
base_dir = os.path.dirname(os.path.abspath(__file__))

wb = openpyxl.open(os.path.join(base_dir, "../tests/100mb.xlsx"))
ws = wb['Tablo3']
def generate_table(width: int, height: int):
    table = [[str(i) for i in range(width)] for _i in range(height)]
    return table

data = generate_table(5, 200)
for row in data:
    ws.append(row)
wb.save(os.path.join(base_dir, "100mb_openpyxl.xlsx"))
