import xlwings as xw
app1 = xw.App(spec='wpsoffice', visible=True, add_book=False)
xlsx_file = r'large-number-bug-reproduction.xlsx'
wb2_result = app1.books.open(xlsx_file, read_only=False)
sht0 = wb2_result.sheets[0]
large_number = sht0.cells(1, 1).options(numbers=int).value  # (rows, column)

print('read number from xlsx: ', large_number)
