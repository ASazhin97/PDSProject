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

KEYWORDS1 = ["2'", "3'", "4'", "6'", "10'", "20'", "30'", "Total", "3 yard", "5 yard",
             "10 yard", "20 yard", "30 yard", "40 yard", "60 yard", "60 yard (rough)",
             "80 yard", "100 yard", "Flop Shot ", "Total", "10 yard", "25 yard", "Total ",
             "Driver", "5 wood/hybrid", "6 Iron", "Driver - Draw", "Driver - Fade",
             "6 Iron - Draw", "6 Iron - Fade", "Ball Flight Laws - 6 Iron", "Total ",
             "Total Strokes"]

KEYWORDS2 = ["FMS Score", "Push - Ups", "Pull - Ups", "Horizontal Rows",
             "Seated Chest Pass (ft)", "Sit up & Throw (ft)", "Plank (sec)",
             "Supine Bridge (sec)", "Vertical Jump (ft)", "Broad Jump (ft)",
             "5-10-5"]
KEYWORDS20 = ["Total Score", "Test 2: PDS Points "]
KEYWORDS3 = ["Scoring Average", "Greens in Regulation %", "Fairways in Regulation %",
             "Putts per Round", "Putts per GIR", "Scrambling %", "Sam Putt Lab",
             "GPC Short Game Test ", "Short Game"]

ath_data = xlwt.Workbook()
sheet1 = ath_data.add_sheet("Master")
ath_data.save("ath_data.xls")


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

def name_from_path(path):
    name_comma_index = None
    for i in range(len(path)):
        if path[i] == ',':
            name_comma_index = i
    for j in range(name_comma_index, 1, -1):
        if path[j] == "\\":
            lastname = path[j+1:name_comma_index]
            break
    for h in range(name_comma_index, len(path)):
        if path[h] == "\\":
            firstname = path[name_comma_index + 2:h]
            break
    fullname = firstname + " " + lastname
    return fullname

def find_index(target, matrix, startRow=0):
    # print("*********************")
    # print("target:", target, "SR:", startRow)
    totRow = startRow
    for row in matrix[startRow:]:
        if target in row:
            if 'Total' in target:
                return totRow, row.index(target) - 1
            c = row.index(target)
            return totRow, row.index(target)
        totRow += 1
    return False, False


def add_data(name, date, matrix):
    """
    Adds data to a list with all the tests in correct positions
    :param name: player name
    :param date: date of test
    :param matrix: matrix representing the spreadsheet to be parsed with just PDS data, no HAT
    :return: 81 index long list of all test data
    """
    data = [name, date]  # add name and date to data list
    r = 0
    # iterate through Test 1: Shot Making scores
    for test in KEYWORDS1:
        r, c = find_index(test, matrix, r)
        if r:  # if r is a number, it is True
            data.append(matrix[r][c + 2])
        else:
            data.append("")
    r, c = find_index("Test 1: PDS Points ", matrix, r)
    if r:
        data.append(matrix[r][c + 1])
    else:
        data.append("")
    r = 0
    # iterate through Test 2: Physical Proficiency
    for test in KEYWORDS2:
        r, c = find_index(test, matrix, r)
        if r:
            data.append(matrix[r][c + 1])
            data.append(matrix[r][c + 2])
        else:
            data.append("")
            data.append("")
    rs = r  # save r for later
    # add Test 2 score and PDS points
    r, c = find_index("Total Score", matrix, r)
    data.append(matrix[r][c + 2])
    r, c = find_index("Test 2: PDS Points ", matrix, r)
    data.append(matrix[r][c + 1])

    # iterate through test 3 golf performance
    for test in KEYWORDS3:
        r, c = find_index(test, matrix, r)
        if r:
            data.append(matrix[r][c + 1])
            data.append(matrix[r][c + 2])
        else:
            data.append("")
            data.append("")
    r, c = find_index("Sam Putt Lab", matrix)  # find index of SPL to work off of
    if r:
        data.append(matrix[r + 1][c + 2])  # add Test 3 total score
        data.append(matrix[r + 2][c + 2])  # add Test 3 PDS points
    else:
        data.append("")
        data.append("")
    r, c = find_index("Player Development Score", matrix)
    if r:
        data.append(matrix[r][c + 1])
    else:
        data.append("")
    return data


def fix_matrix(matrix):
    """
    This function deletes the HAT data from top of matrix
    :param matrix: contains whole spreadsheet as a matrix
    :return: new matrix with just the PDS data, no HAT data
    """
    r, c = find_index("GPC PDS Testing Sheet", matrix)
    return matrix[r:]


def process_file(fileName, playerName):
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
            print('Processing sheet', sheet, ". . .")
            curSheet = wb.sheet_by_name(sheet)
            date = sheet.replace('.', '/')
            # delete letters from date
            newDate = ''
            for char in date:
                if char in DIGITS or char == '/':
                    newDate += char
            date = newDate
            # get the relevant data in a list of lists called matrix
            matrix = [[curSheet.cell_value(r, c) for c in range(11)] for r in range(curSheet.nrows)]
            matrix = fix_matrix(matrix)
            data = add_data(playerName, date, matrix)  # list with all relevant data for one test
            ath_data_add(data) # add data to master spreadsheet. change: uncomment
            print('Successfully processed sheet', sheet)



def main():
    # change these lines

    PDS = xlrd.open_workbook("2017-2018 query.xlsx")
    sheet = PDS.sheet_by_index(0)
    paths = [sheet.cell_value(col, 7) for col in range(1, 85)]

    for path in paths:
        print(path)
        process_file(path, name_from_path(path))
    return


main()

