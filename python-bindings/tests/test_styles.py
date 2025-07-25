from excelsior import Scanner, AlignSpec, HorizAlignment
import os

base_dir = os.path.dirname(os.path.abspath(__file__))
inp_filename = os.path.join(base_dir, "../../test/style_test.xlsx")

out_filename = os.path.join(base_dir, "../../test/style_test_out_py.xlsx")

scanner = Scanner(inp_filename)
editor = scanner.open_editor(scanner.get_sheets()[0])
# .set_fill("B24:B28", "FFFF00")\
editor\
        .set_font("D4:D8", "Arial", 12.0, True, False)\
        .set_border("A1:C3", "thin")\
        .set_font("A1:C3", "Calibri", 10.0, False, True, AlignSpec(horiz=HorizAlignment.Center))\
        .merge_cells("B12:D12")
editor.save(out_filename)
print("Done")
