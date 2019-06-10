import xlwt
import xlrd

ath_data = xlwt.Workbook()
sheet1 = ath_data.add_sheet("Master")


# adds athelete's data to ath_data file
def ath_data_add(data):
    row_read = xlrd.open_workbook("ath_data.xls")
    # finds the number of rows currently in the workbook, used to find appropriate spacing for write
    row_count = row_read.sheet_by_index(0).nrows
    # checks to see if player name is present, skips if not present
    if data[0] == '':
        return
    # for loop used to find spacing and index of relevant data in passed through data file
    for index in range(len(data)):
        sheet1.write(row_count, index, data[index])
    ath_data.save('ath_data.xls')
    return
