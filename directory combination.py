import shutil
import xlrd

PDS = xlrd.open_workbook("All Hat and PDS tests.xlsx")
sheet = PDS.sheet_by_index(0)
paths = [sheet.cell_value(col, 8) for col in range(1, 133)]
print(paths)




