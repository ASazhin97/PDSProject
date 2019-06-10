import xlwt
import xlrd
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

file = xlwt.Workbook()
sheet = file.add_sheet("yes")
for i in range(len(TEMPLATE_LIST)):
    sheet.write(0, i, TEMPLATE_LIST[i])
file.save('headers.xls')