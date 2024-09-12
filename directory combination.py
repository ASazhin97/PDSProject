
import xlrd


# # for 2017-18
PDS = xlrd.open_workbook("2017-2018 query.xlsx")
sheet = PDS.sheet_by_index(0)
paths = [sheet.cell_value(col, 7) for col in range(1, 85)]
print(paths)



# for 2018-19
PDS = xlrd.open_workbook("2018-19 query.xlsx")
sheet = PDS.sheet_by_index(0)
paths = [sheet.cell_value(col, 7) for col in range(1, 87)]
print(paths)
