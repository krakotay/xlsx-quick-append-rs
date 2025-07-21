from xlsx_append_py import PyXlsxScanner
import os
import shutil

# Copy 'source.txt' to 'destination.txt'

base_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(base_dir, "../../test/test_new_ws.xlsx")
shutil.copy(file_path, file_path + "_copy.xlsx")

scanner = PyXlsxScanner(file_path + "_copy.xlsx")
print(scanner.get_sheets())
assert scanner.get_sheets() == ["Sheet1"]



editor = scanner.open_editor("Sheet1")
editor.add_worksheet("NewSheet")
editor.add_worksheet("NewSheetTwo")
scanner = PyXlsxScanner(file_path + "_copy.xlsx")
print(scanner.get_sheets())
assert set(scanner.get_sheets()) == set(["Sheet1", "NewSheet", "NewSheetTwo"])


print('passed')