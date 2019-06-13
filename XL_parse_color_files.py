import xlrd
import xlwt

DIGITS = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
# list to hold the names of each column
TEMPLATE_LIST = ['Name', 'Date', 'Putting: 2 ft.', 'Putting: 3 ft.', 'Putting: 4 ft.',
                 'Putting: 6 ft.', 'Putting: 10 ft.', 'Putting: 20 ft.', 'Putting: 30 ft.',
                 'Putting Total', 'C/P: 3 yard', 'C/P: 5 yard', 'C/P: 10 yard', 'C/P: 20 yard',
                 'C/P: 30 yard', 'C/P: 40 yard', 'C/P: 60 yard', 'C/P: 60 yard (rough)',
                 'C/P: 80 yard', 'C/P: 100 yard', 'C/P: Flop Shot', 'C/P Total',
                 'Bunker: 10 yard', 'Bunker: 25 yard', 'Bunker Total', 'FS: Driver',
                 'FS: 5 wood/hybrid', 'FS: 6 iron', 'FS: Driver - Draw', 'FS: Driver - Fade',
                 'FS: 6 iron - Draw', 'FS: 6 iron - Fade', 'FS: BFC 6 iron', 'FS Total',
                 'Total Strokes', 'Test 1: PDS Points', 'FMS Score: Result', 'FMS Score: PDS Score',
                 'Push - Ups: Result', 'Push - Ups: PDS Score', 'Pull - Ups: PDS Score',
                 'Pull - Ups: Result', 'Horizontal Rows: Result', 'Horizontal Rows: PDS Score',
                 'Seated Chest Pass (ft): Result', 'Seated Chest Pass: PDS Score',
                 'Sit up & Throw (ft): Result', 'Sit up & Throw: PDS Score', 'Plank (sec): Result',
                 'Plank: PDS Score', 'Supine Bridge (sec): Result', 'Supine Bridge: PDS Score',
                 'Vertical Jump (ft): Result', 'Vertical Jump: PDS Score',
                 'Broad Jump (ft): Result', 'Broad Jump: PDS Score', '5-10-5: Result',
                 '5-10-5: PDS Score', 'Test 2: Physical Proficiency Total Score',
                 'Test 2: PDS Points', 'Scoring Average', 'Scoring Average: PDS Score', 'GIR',
                 'GIR: PDS Score', 'FIR', 'FIR: PDS Score', 'Putts per Round',
                 'Putts per Round: PDS Score', 'Putts per GIR', 'Putts per GIR: PDS Score',
                 'Scrambling', 'Scrambling: PDS Score', 'SAM Putt Lab Score',
                 'SAM Putt Lab: PDS Score', 'GPC Short Game Test',
                 'GPC Short Game Test: PDS Score', 'Short Course (Bumpy) Score',
                 'Short Course (Bumpy): PDS Score', 'Test 3: Golf Performance Total Score',
                 'Test 3: PDS Points', 'PDS Score'
                 ]

# format for these lists is that a tuple represents where the data is for the
# following strings relative to the indices of the string. Tuples take form
# (delta col, # to add, start row)
KW1 = [(2, 1, 0), "2'", "3'", "4'", "6'", "10'", "20'", "30'", (1, 1, 0),
       "Putting Overall Score", (2, 1, 13), "3 yard", "5 yard",
       "10 yard", "20 yard", "30 yard", "40 yard", "60 yard", "60 yard rough",
       "80 yard", "100 yard", "Flop Shot", (1, 1, 13), "Wedge Control Overall Score",
       (2, 1, 25), "10 yard", "25 yard", (1, 1, 25), "Bunker Overall Score", (2, 1, 25),
       "Driver ", "5w or Hybrid", "6 Iron", "Driver - Draw", "Driver - Fade",
       "6 Iron - Draw", "6 Iron - Fade", "BFC - 6i", (1, 1, 25), "Full Swing Overall Score",
       "Total Strokes", "GPC Golf Shot Making", (1, 2, 0), "FMS Score", "Push - Ups",
       "Pull - Ups", "Horizontal Rows", "Seated Chest Pass (ft)", "Sit up & Throw (ft)",
       "Plank (sec)", "Supine Bridge (sec)", "Vertical Jump (in)", "Broad Jump (ft)",
       "5-10-5", (1, 1, 11), "Total Score", "Test 2: PDS Points "]

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


def find_index(target, matrix, startRow=0):
    # print("*********************")
    # print("target:", target, "Start Row:", startRow)
    r = startRow
    for row in matrix[startRow:]:
        if target in row:
            c = row.index(target)
            # print('Found:', target, "at r and c:", r, c)
            return r, row.index(target)
        r += 1
    # print("did not find", target)
    return False, False


def add_res(data, matrix, target, dc, quant_to_add, sr):
    r, c = find_index(target, matrix, sr)
    if r:
        for i in range(quant_to_add):
            data.append(matrix[r][c + dc + i])
    else:
        for i in range(quant_to_add):
            data.append("")
    return data


def add_all_data(mat, date):
    data = []
    r, c = find_index("Name:", mat)
    data.append(mat[r][c + 2])  # add name
    data.append(date)
    dc, quant_to_add, sr = 0, 1, 0  # delta row, delta col
    for el in KW1:
        if type(el) is str:
            # call function to add
            data = add_res(data, mat, el, dc, quant_to_add, sr)
        elif type(el) is tuple:
            dc, quant_to_add, sr = el[0], el[1], el[2]
    # now add scoring avg. and other golf stats
    r, c = find_index("Scoring Average", mat)  # find first occurrence of SA
    r, c = find_index("Scoring Average", mat, r + 1)  # find second occurrence
    if r:
        data.append(mat[r][c + 1])
        data.append(mat[r][c + 2])
    else:
        data.append("")
        data.append("")
    # skip missing data
    for i in range(10):
        data.append("")
    # add SAM putt lab and save coordinates
    r, c = find_index("Sam Putt Lab", mat)
    save_r, save_c = r, c
    if r:
        data.append(mat[r][c + 1])
        data.append(mat[r][c + 2])
    else:
        data.append("")
        data.append("")
    for i in range(2):  # skip GPC short game test
        data.append("")
    r, c = find_index("Short Game", mat)
    if r:
        data.append(mat[r][c + 1])
        data.append(mat[r][c + 2])
    else:
        data.append("")
        data.append("")
    data.append(mat[save_r + 1][save_c + 2])  # add test 3 total score
    data.append(mat[save_r + 2][save_c + 2])  # add test 3 PDS points
    r, c = find_index("Short Game", mat)
    if r:
        data.append(mat[r][c + 1])
    else:
        data.append("")
    return data


def print_formatted(matrix):
    """
    helper method for debugging. Prints matrix in nice format
    """
    for row in matrix:
        print(row)


def process_file(fileName):
    """
    Takes a string as the parameter that represents the directory and name of
    the .xlsx file.
    Puts data in the form [name, date, attribute1, a2, ...] for each test
    """
    wb = xlrd.open_workbook(fileName)  # open the spreadsheet
    sheetNames = wb.sheet_names()  # list of the sheet names (dates)

    # loop iterates through all sheets in a document
    for sheet in sheetNames:
        if sheet[0] in DIGITS:  # filter out sheet names that are not dates
            curSheet = wb.sheet_by_name(sheet)
            # get the relevant data in a list of lists called matrix
            matrix = [[curSheet.cell_value(r, c) for c in range(curSheet.ncols)]
                      for r in range(curSheet.nrows)]

            # get list with all relevant data for one test
            data = add_all_data(matrix, sheet.replace('.', '/'))
            ath_data_add(data)  # add data to master spreadsheet


def main():
    # change this
    PDS = xlrd.open_workbook("2018-19 query.xlsx")
    sheet = PDS.sheet_by_index(0)
    paths = [sheet.cell_value(col, 7) for col in range(1, 87)]
    for path in paths:
        print(path)
        process_file(path)
    return


main()
