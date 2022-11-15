# uses the xlswriter module. The module overwrites the data in the excel sheet
# import xlsxwriter module
import xlsxwriter

department = "Army"
sent = 34
sentCharge = 0.37
received = 78
receivedCharge = 0.04

workbook = xlsxwriter.Workbook('report.xls')

# By default worksheet names in the spreadsheet will be
# Sheet1, Sheet2 etc., but we can also specify a name.
worksheet = workbook.add_worksheet("Fax_Dept_Nov_2022")

# Some data we want to write to the worksheet.
data = (
	['ankit', 1000, 0.02, 254, 0.03],
	['rahul', 100, 0.01, 457, 0.02],
	['priya', 300, 0.04, 698, 0.04],
	['harshita', 505, 0.02, 657, 0.01],
    ['kev', 187,0.12, 789, 0.45]
)

# Start from the first cell. Rows and
# columns are zero indexed.
row = 6
col = 0

# Iterate over the data and write it out row by row.
for dept, sent, srate, received, rrate in data:
    worksheet.write(row, col, dept)
    worksheet.write(row, col + 1, sent)
    worksheet.write(row, col + 2, srate)
    worksheet.write(row, col + 3, received)
    worksheet.write(row, col + 4, rrate)
    
    row += 1

workbook.close()
