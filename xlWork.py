import xlrd

DIGITS = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']


def find_index(lst, target):
    rowCount = 0
    for row in lst:
        if target in row:
            return rowCount, row.index(target)
        rowCount += 1
    return -1, -1


fileLocation = r'C:\Users\Intern-5\Documents\PDS\2016-2017 Competitive Red (5-7)\Donnelly, Michael'

wb = xlrd.open_workbook('Hat & PDS Test.xlsx')

sheetNames = wb.sheet_names()

for sheet in sheetNames:
    if sheet[0] in DIGITS:
        curSheet = wb.sheet_by_name(sheet)
        # print("sheet name:", sheet, "rows:", curSheet.nrows, "columns:", curSheet.ncols)
        data = [[curSheet.cell_value(r, c) for c in range(9)] for r in range(64, 111)]
        row, col = find_index(data, 'Putting ')
        print(row, col)

# 'Putting' in sheet has space after it, making it "Putting "

####### RANGE DEFINITIONS ###########
# 'Putting " located at (3,2)


"""
print(sheetNames, "\n")
print(type(sheetNames[0]))
print(type(sheetNames[1]))

x = wb.sheet_by_name(sheetNames[0])

for col in range(3):
    for row in range(6):
        print(x.cell(row, col))

print(x.cell(1, 1).value)
print(type(x.cell(1, 1).value))

lst = [5, 7, 7, 4, 'hi']
print(5 in lst)
print(10 in lst)
print(lst.index('hi'))
"""
