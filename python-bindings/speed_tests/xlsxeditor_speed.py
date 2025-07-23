from excelsior import PyXlsxEditor
import os
base_dir = os.path.dirname(os.path.abspath(__file__))

editor = PyXlsxEditor(os.path.join(base_dir, "../tests/100mb.xlsx"), "Tablo3")
def generate_table(width: int, height: int):
    table = [[str(i) for i in range(width)] for _i in range(height)]
    return table
data = generate_table(5, 200)
for row in data:
    editor.append_row(row)
editor.save(os.path.join(base_dir, "100mb_appended_xlsxeditor.xlsx"))
