from excelsior import Scanner
import polars as pl
import os
df = pl.DataFrame(
    {
        "int": [1, 2, 3],
        "float": [1.1, 2.2, 3.3],
        "string": ["a", "b", "c"],
        "bool": [True, False, True], # not yet fully supported, only as string
        "long_string" : ["123" * 10, "123" * 10, "123" * 10]
    }
)
print(df)
base_dir = os.path.dirname(os.path.abspath(__file__))
inp_filename = os.path.join(base_dir, "../../test/test_polars.xlsx")
out_filename = os.path.join(base_dir, "../../test/test_polars_appended.xlsx")

scanner = Scanner(inp_filename)
editor = scanner.open_editor(scanner.get_sheets()[0])
# editor.with_polars(df, "A1")
editor.add_worksheet('polars_ws').with_polars(df, "B4")
editor.add_worksheet('polars_ws_2').with_polars(df)
editor.save(out_filename)