import sys

class Paddler:
	def __init__(self, name, weight, gender, side, timetrial, notes):
		self.Name = name
		self.Weight = weight
		self.Gender = gender
		self.Side = side
		self.TimeTrial = timetrial
		self.Notes = notes

	def printPaddler(self):
		print(self.Name , self.Weight, self.Gender, self.Side, self.TimeTrial, self.Notes)
		
	def printPaddler(self, index):
		print(index, self.Name , self.Weight, self.Gender, self.Side, self.TimeTrial, self.Notes)

	def getWeight(self):
		return self.Weight

class Boat:
	def __init__(self, BoatNumber, leftSide, rightSide, subs):
		self.BoatNumber = BoatNumber
		self.LeftSide = leftSide
		self.RightSide = rightSide
		self.Subs = subs

	def printBoat(self):
		print('Boat: ', self.BoatNumber)
		print('Boat Left Side: ')
		for i in range(0, len(self.LeftSide)):
			self.LeftSide[i].printPaddler(i + 1)
		print('Boat Right Side: ')
		for i in range(0, len(self.RightSide)):
			self.RightSide[i].printPaddler(i + 1)
		print('Boat Substitutes: ')
		for i in range(0, len(self.Subs)):
			self.Subs[i].printPaddler(i + 1)
	
	

