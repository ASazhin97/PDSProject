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
    for i in range(len(TEMPLATE_LIST)):
        sheet1.write(0, i, TEMPLATE_LIST[i])
    for index in range(len(data)):
        sheet1.write(row_count, index, data[index])
    ath_data.save('ath_data.xls')
    return




def correctFeet(data, toCorrect):
    """
    Helper function to correct the values in indices given in the list
    toCorrect to feet
    Returns data list with corrected values
    """
    for i in toCorrect:
        num = int(data[i])
        w = num // 1  # whole number part
        d = (num % 1) * 100 # decimal part converted to whole number

        # this if statement corrects for values that were input incorrectly
        # which caused > 12 inches to be in inches part
        while d >= 12:
            d = d / 10
        data[i] = round(w + (d / 12), 3)
    return data


def add_data(name, date, matrix):
    data = [name, date]  # add name and date to data list

    # putting
    r, c = 4, 4  # indices for first data point to append (2 ft. putt)
    while r <= 11:
        data.append(matrix[r][c])
        r += 1
    # chipping / pitching
    r = 14
    while r <= 25:
        data.append(matrix[r][c])
        r += 1
    # bunker
    r = 28
    while r <= 30:
        data.append(matrix[r][c])
        r += 1
    # full swing out of 6
    r = 33
    while r <= 35:
        data.append(matrix[r][c])
        r += 1
    # full swing out of 4
    r = 37
    while r <= 40:
        data.append(matrix[r][c])
        r += 1
    # full swing out of 9
    data.append(matrix[42][4])
    data.append(matrix[43][4])
    # total strokes and Test 1: PDS Points
    data.append(matrix[45][4])
    data.append(round(matrix[46][4], 3))

    # Loading Test 2: Phys Proficiency data
    r, c = 4, 7
    while r <= 14:
        while c <= 8:
            if not isinstance(matrix[r][c], str):
                data.append(round(matrix[r][c], 3))
            else:
                data.append(matrix[r][c])
            c += 1
        c = 7
        r += 1
    # total score and Test 2 PDS Points
    data.append(round(matrix[15][8], 3))
    data.append(round(matrix[16][8], 3))

    # Loading Test 3: GPC golf performance
    r, c = 21, 7
    while r <= 29:
        while c <= 8:
            if not isinstance(matrix[r][c], str):
                data.append(round(matrix[r][c], 3))
            else:
                data.append(matrix[r][c])
            c += 1
        c = 7
        r += 1
    # total score and Test 3 PDS Points
    data.append(round(matrix[30][8], 3))
    data.append(round(matrix[31][8], 3))

    # Load PDS Score
    data.append(round(matrix[38][7], 3))

    # correct feet and inches form to feet
    #data = correctFeet(data, [44, 46, 52, 54])
    return data


def process_file(fileName):
    """
    Takes a string as the parameter that represents the directory and name of
    the .xlsx file.
    Puts data in the form [name, date, attribute1, a2, ...] for each test, and
    call csv_write() on each list to add to a master csv file
    """
    wb = xlrd.open_workbook(fileName)  # open the spreadsheet
    sheetNames = wb.sheet_names()  # list of the sheet names (dates)

    # loop iterates through all sheets in a document
    for sheet in sheetNames:
        if sheet[0] in DIGITS:  # filter out sheet names that are not dates
            curSheet = wb.sheet_by_name(sheet)
            date = sheet.replace('.', '/')
            playerName = curSheet.cell_value(1, 13)[12:].strip()
            # get the relevant data in a list of lists called matrix
            matrix = [[curSheet.cell_value(r, c) for c in range(9)] for r in range(64, 111)]
            data = add_data(playerName, date, matrix)  # list with all relevant data for one test
            ath_data_add(data) # add data to master spreadsheet


def main():
    PDS = xlrd.open_workbook("All Hat and PDS tests.xlsx")
    sheet = PDS.sheet_by_index(0)
    paths = [sheet.cell_value(col, 8) for col in range(1, 133)]
    for path in paths:
        print(path)
        process_file(path)
        print("Done!")
    return


main()