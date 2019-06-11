
import xlrd

PDS = xlrd.open_workbook("2017-2018 query.xlsx")
sheet = PDS.sheet_by_index(0)
paths = [sheet.cell_value(col, 7) for col in range(1, 85)]
print(paths)