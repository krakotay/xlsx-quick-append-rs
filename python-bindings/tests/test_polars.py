from xlsx_append_py import PyXlsxScanner
import polars as pl
import os
import tempfile
df = pl.DataFrame(
    {
        "int": [1, 2, 3],
        "float": [1.1, 2.2, 3.3],
        "string": ["a", "b", "c"],
        "bool": [True, False, True], # not yet fully supported, only as string
    }
)
print(df)
base_dir = os.path.dirname(os.path.abspath(__file__))
inp_filename = os.path.join(base_dir, "../../test/test_polars.xlsx")

out_filename = os.path.join(base_dir, "../../test/test_polars_appended.xlsx")

scanner = PyXlsxScanner(inp_filename)
sheets = scanner.get_sheets()
editor = scanner.open_editor(sheets[0])
editor.with_polars(df, None)
editor.with_polars(df, "B15")
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_excel_file:
    tf_path = temp_excel_file.name
    editor.save(tf_path)
print("1. Done")

scanner = PyXlsxScanner(tf_path)
sheets = scanner.get_sheets()
editor = scanner.open_editor(sheets[1])
editor.with_polars(df, None)
editor.with_polars(df, "AH2000")
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_excel_file:
    tf_path = temp_excel_file.name
    editor.save(tf_path)
print("2. Done")

scanner = PyXlsxScanner(tf_path)
sheets = scanner.get_sheets()
editor = scanner.open_editor(sheets[1])
editor.add_worksheet('polars_ws')
editor = scanner.open_editor('polars_ws')
editor.append_table_at([['1', '2', '3'], ['4', '5', '6']], 'B40')
editor.with_polars(df, None)
editor.with_polars(df, "D20")
editor.save(out_filename)
print("3. Done")

