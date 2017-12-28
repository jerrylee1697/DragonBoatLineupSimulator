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
AllPaddlers.sort(key=operator.attrgetter('TimeTrial'))

LeftPaddlers = []
RightPaddlers = []
Both = []

# Seperates into Left, Right and Both Side preferences
for i in range (0, numPaddlers):
	if AllPaddlers[i].Side == 'L':
		LeftPaddlers.append(AllPaddlers[i])
	elif AllPaddlers[i].Side == 'R':
		RightPaddlers.append(AllPaddlers[i])
	elif AllPaddlers[i].Side == 'B':
		Both.append(AllPaddlers[i])

while 1:
	SplitRatio = input('Choose option for Gender Split Ratio (Male:Female): \n1. 10:10 \n2. 12:6 ')
	
	if SplitRatio == 1:
		numMales = 10
		numFemales = 10
		BoatSize = 20
		break
	elif SplitRatio == 2:
		numMales = 12
		numFemales = 6
		BoatSize = 18
		break
	else:
		print('Invalid Selection, Please try again.')


BoatLeft = []
BoatRight = []
BoatBoth = []

MaleCounter = 0
FemaleCounter = 0
i = 0
while i < len(AllPaddlers):
	if len(BoatRight) + len(BoatLeft) + len(BoatBoth) < numMales + numFemales:
		if AllPaddlers[i].Gender == 'M' and MaleCounter < numMales:
			# print(AllPaddlers[i].printPaddler())
			if AllPaddlers[i].Side == 'R' and len(BoatRight) < BoatSize / 2:
				BoatRight.append(AllPaddlers[i])
				AllPaddlers.remove(AllPaddlers[i])
				MaleCounter += 1
			elif AllPaddlers[i].Side == 'L' and len(BoatLeft) < BoatSize / 2:
				BoatLeft.append(AllPaddlers[i])
				AllPaddlers.remove(AllPaddlers[i])
				MaleCounter += 1
			elif AllPaddlers[i].Side == 'B' and len(BoatBoth) < BoatSize:
				BoatBoth.append(AllPaddlers[i])
				AllPaddlers.remove(AllPaddlers[i])
				MaleCounter += 1
			# print('Male ', MaleCounter, ' Added')
			# print(MaleCounter)
		elif AllPaddlers[i].Gender == 'F' and FemaleCounter < numFemales:
			# print(AllPaddlers[i].printPaddler())
			if AllPaddlers[i].Side == 'R' and len(BoatRight) < BoatSize / 2:
				BoatRight.append(AllPaddlers[i])
				AllPaddlers.remove(AllPaddlers[i])
				FemaleCounter += 1
			elif AllPaddlers[i].Side == 'L' and len(BoatLeft) < BoatSize / 2:
				BoatLeft.append(AllPaddlers[i])
				AllPaddlers.remove(AllPaddlers[i])
				FemaleCounter += 1
			elif AllPaddlers[i].Side == 'B' and len(BoatBoth) < BoatSize:
				BoatBoth.append(AllPaddlers[i])
				AllPaddlers.remove(AllPaddlers[i])
				FemaleCounter += 1
			
			# print('Female ', FemaleCounter, ' Added')
	i += 1
		
# print(numMales + numFemales)
# print(len(BoatBoth) + len(BoatLeft) + len(BoatRight))

print('Left Paddlers:')
for i in range (0, len(BoatLeft)):
	BoatLeft[i].printPaddler()
print('Right Paddlers:')
for i in range (0, len(BoatRight)):
	BoatRight[i].printPaddler()
print('Both: ')
for i in range (0, len(BoatBoth)):
	BoatBoth[i].printPaddler()

# print(sheet.cell(row=j, column=i).value)

# writer = pd.ExcelWriter('List.xlsx', engine='xlsxwriter')

# yourData.to_excel(writer, 'Sheet1')

# writer.save()