import os
import pandas as pd 
import operator

from classes import Paddler, Boat
from openpyxl import load_workbook

# cwd = os.getcwd()
# print(cwd)



# print(os.listdir('.'))

# Assign filename to 'file'
# file = 'List.xlsx'

# Load Spreadsheet
# xl = pd.ExcelFile(file)

# Print the sheet names
# print(xl.sheet_names)

# Load a sheet into a DataFrame by name: df1
# df1 = xl.parse('Sheet1')


wb = load_workbook('./List.xlsx')

# print(wb.get_sheet_names())

sheet = wb.get_sheet_by_name('Sheet1')

sheet.title

# anotherSheet = wb.active

# anotherSheet

numPaddlers = input('Input number of Paddlers: ')

# sheet['A1'].value

# All Paddler Information passed into AllPaddlers
AllPaddlers = []

for i in range (2,numPaddlers + 2):
	name = sheet.cell(row=i, column=1).value
	weight = sheet.cell(row=i, column=2).value
	gender = sheet.cell(row=i, column=3).value
	side = sheet.cell(row=i, column=4).value
	time = sheet.cell(row=i, column=5).value
	notes = sheet.cell(row=i, column=6).value
	newPaddler = Paddler(name, weight, gender, side, time, notes)
	# newPaddler.printPaddler()
	AllPaddlers.append(newPaddler)

# Sorts Paddlers by Trial Times
AllPaddlers.sort(key=operator.attrgetter('timetrial'))

LeftPaddlers = []
RightPaddlers = []
Both = []

# Seperates into Left, Right and Both Side preferences
for i in range (0, numPaddlers):
	if AllPaddlers[i].side == 'L':
		LeftPaddlers.append(AllPaddlers[i])
	elif AllPaddlers[i].side == 'R':
		RightPaddlers.append(AllPaddlers[i])
	elif AllPaddlers[i].side == 'B':
		Both.append(AllPaddlers[i])

while 1:
	SplitRatio = input('Choose option for Gender Split Ratio (Male:Female): \n1. 10:10 \n2. 12:6 ')
	
	if SplitRatio == 1:
		numMales = 10
		numFemales = 10
		break
	elif SplitRatio == 2:
		numMales = 12
		numFemales = 6
		break
	else:
		print('Invalid Selection, Please try again.')

BoatLeft = []
BoatRight = []
BoathBoth = []

		

print('Left Paddlers:')
for i in range (0, len(LeftPaddlers)):
	LeftPaddlers[i].printPaddler()
print('Right Paddlers:')
for i in range (0, len(RightPaddlers)):
	RightPaddlers[i].printPaddler()
print('Both: ')
for i in range (0, len(Both)):
	Both[i].printPaddler()

# print(sheet.cell(row=j, column=i).value)

# writer = pd.ExcelWriter('List.xlsx', engine='xlsxwriter')

# yourData.to_excel(writer, 'Sheet1')

# writer.save()