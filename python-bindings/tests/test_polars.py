from xlsx_append_py import scan_excel, PyXlsxEditor
import polars as pl
import os
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

sheets = scan_excel(os.path.join(base_dir, "../../test/test_polars.xlsx"))
editor = PyXlsxEditor(os.path.join(base_dir, "../../test/test_polars.xlsx"), sheets[0])
editor.with_polars(df, None)
editor.with_polars(df, "B15")
editor.save(os.path.join(base_dir, "../../test/test_polars_appended.xlsx"))

print("Done")

editor = PyXlsxEditor(os.path.join(base_dir, "../../test/test_polars.xlsx"), sheets[1])
editor.with_polars(df, None)
editor.with_polars(df, "AH2000")
editor.save(os.path.join(base_dir, "../../test/test_polars_appended.xlsx"))
