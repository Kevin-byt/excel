import openpyxl

# def create_workbook(path):
#     workbook = openpyxl.Workbook()
#     workbook.save(path)

# create_workbook('data.xlsx')

data = (
	['ankit', 1000, 0.02, 254, 0.03],
	['rahul', 100, 0.01, 457, 0.02],
	['buga', 300, 0.04, 698, 0.04],
	['fest', 505, 0.02, 657, 0.01],
    ['kev', 187,0.12, 789, 0.45]
)

workbook = openpyxl.load_workbook('data.xlsx')
sheet = workbook.active


# sheet["A1"] = "Hello"
# sheet["B1"] = "from"
# sheet["C1"] = "Jungle"

r = 3
for dept, sent, srate, received, rrate in data:
    sheet['A'+str(r)] = dept
    sheet['B'+str(r)] = sent
    sheet['C'+str(r)] = srate
    sheet['D'+str(r)] = received
    sheet['E'+str(r)] = rrate
    
    r += 1

workbook.save('data.xlsx')