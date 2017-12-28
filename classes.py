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
		# sys.stdout.write(self.name)
		# sys.stdout.flush()
		# sys.stdout.write(self.weight)
		# sys.stdout.flush()
		# sys.stdout.write(self.gender)
		# sys.stdout.flush()

class Boat:
	def __init__(self, leftSide, rightSide, subs):
		self.leftSide = leftSide
		self.rightSide = rightSide
		self.subs = subs