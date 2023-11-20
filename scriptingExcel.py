# import numpy

# with open('msft_10K.xls', 'r', encoding='utf-8') as input_file:
#     print(type(input_file))
#     lines = input_file.read()
#     for line in lines:
#         print(line)


import xlrd
book = xlrd.open_workbook("msft_10K.xls")
# print("The number of worksheets is {0}".format(book.nsheets))
# print("Worksheet name(s): {0}".format(book.sheet_names()))

if 'Income Statements' in book.sheet_names():
    print("Inside the I/S")
    idx = book.sheet_names().index('Income Statements')
    IS = book.sheet_by_index(idx)
    year = 2023

    years_row = 4
        
    for i, cell in enumerate(IS.col(0)):
        if IS.cell_value(i, 0) == 'Total revenue':
            for j, cell2 in enumerate(IS.row(i)):
                if IS.cell_value(years_row, j):
                    year = IS.cell_value(years_row, j)
                if IS.cell_type(i, j) == 2:
                    print(f'{year}: {IS.cell_value(i,j)}')
        
        
            


if 'Balance Sheet' in book.sheet_names():
    print("Inside the B/S")
    idx = book.sheet_names().index('Balance Sheet')


if 'Cash Flow Statement' in book.sheet_names():
    print("Inside the C/S")
    idx = book.sheet_names().index('Cash Flow Statement')


# for i in book.sheet_names():
#     print(i)
# sh = book.sheet_by_index(0)
# print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
# for rx in range(sh.nrows):
#     print(sh.row(rx))
