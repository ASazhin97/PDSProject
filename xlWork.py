import xlrd

DIGITS = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
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
                 'Test 2: PDS Points', 'Scoring Average', 'Scoring Average: PDS Score', 'GIR %',
                 'GIR: PDS Score', 'FIR %', 'FIR: PDS Score', 'Putts per Round',
                 'Putts per Round: PDS Score', 'Putts per GIR', 'Putts per GIR: PDS Score',
                 'Scrambling %', 'Scrambling: PDS Score', 'SAM Putt Lab Score',
                 'SAM Putt Lab: PDS Score', 'GPC Short Game Test',
                 'GPC Short Game Test: PDS Score', 'Short Course (Bumpy) Score',
                 'Short Course (Bumpy): PDS Score', 'Test 3: Golf Performance Total Score'
                 'Test 3: PDS Points', 'PDS Score'
                 ]


def find_index(lst, target):
    """
    Finds index (row, col) of a target value given a list of lists
    Returns: row, col
    """
    rowCount = 0
    for row in lst:
        if target in row:
            return rowCount, row.index(target)
        rowCount += 1
    return -1, -1


def add_data(name, date, matrix):
    data = [name, date]

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
            print(playerName, date)


fileName = r'C:\Users\Intern-5\Documents\PDS\2016-2017 Competitive Red ' \
           r'(5-7)\Donnelly, Michael\Hat _ PDS Test 03 15.xlsx'
process_file(fileName)
