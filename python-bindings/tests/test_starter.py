from xlsx_append_py import scan_excel, PyXlsxEditor

print(scan_excel("tests/test_sum.xlsx"))
sheets = scan_excel("tests/test_sum.xlsx")
editor = PyXlsxEditor("tests/test_sum.xlsx", sheets[0])
# editor.append_row(["Hello", "World"])
# editor.save("tests/Для_ТЗ_2_appended.xlsx")
editor.append_table_at([["=_xlfn.SUM(A1:B1)"], ['=_xlfn.SUM(A2:B2)']], "C1")
editor.save("tests/test_sum_appended.xlsx")
