import sys

class Paddler:
	def __init__(self, name, weight, gender, side, timetrial, notes):
		self.name = name
		self.weight = weight
		self.gender = gender
		self.side = side
		self.timetrial = timetrial
		self.notes = notes

	def printPaddler(self):
		print(self.name , self.weight, self.gender, self.side, self.timetrial, self.notes)
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