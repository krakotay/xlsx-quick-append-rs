from excelsior import Scanner
import os

base_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_dir, "../../test/style_test.xlsx")
file_path_out = os.path.join(base_dir, "../../test/style_test_borders_py.xlsx")
scanner = Scanner(file_path)
print(scanner.get_sheets())
editor = scanner.open_editor(scanner.get_sheets()[0])
editor.set_border("D4:G8", "thin").merge_cells('D4:G4').set_font('D4:G4', 'Arial', 12, False, False)
editor.save(file_path_out)
