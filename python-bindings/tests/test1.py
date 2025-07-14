from xlsx_append_py import scan_excel

assert scan_excel("../../test/test_sum.xlsx") == ["Sheet1"]
