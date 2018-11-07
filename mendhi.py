from openpyxl import load_workbook
import random

zips = {}

wb = load_workbook('ZipCodes.xlsx')
ws = wb.active
for r in range(1,101):
	for c in range(1,11):
		a = str(ws.cell(row=r,column=c).value).split(' ',1)
		a[1] = a[1].split('\n')
		a[1] = a[1][0] + ' ' + a[1][1]
		zips[a[0]] = a[1]

winners = []

wb = load_workbook('Sample More Doc.xlsx')
ws = wb.active

class Contestants:

	def __init__(self):
		self.contestants = []
		self.cheaters = []
		self.suspiciousPeeps = []
		self.createContestants()
		self.findSuspiciousPeople()
		self.deleteSuspiciousPerson()

	def createContestants(self):
		for r in range(2,ws.max_row+1):
			person = []
			for x in range(1,4):
				person.append(str(ws.cell(row=r,column=x).value).upper().strip().split(' ')[0])
			person.append(str(int(ws.cell(row=r,column=4).value)))
			if person in self.cheaters:
				pass
			elif person in self.contestants:
				self.cheaters.append(person)
				self.contestants.remove(person)
			else:
				self.contestants.append(person)

	def printContestants(self):
		print("CONTESTANTS")
		print('-'*20)
		for x in range(len(self.contestants)):
			print(formatPerson(self.contestants[x]))
	
	def printCheaters(self):
			print("CHEATERS")
			print('-'*20)
			for x in range(len(self.cheaters)):
				print(formatPerson(self.cheaters[x]))

	def findSuspiciousPeople(self):
		for person in self.contestants:
			p1 = set(tuple(person))
			for person2 in self.contestants:
				p2 = set(tuple(person2))
				if len(p1.intersection(p2)) == 3:
					self.suspiciousPeeps.append(person)

	def printSuspiciousPeople(self):
		print("SUSPICIOUS PEOPLE")
		print('-'*20)
		for x in range(len(self.suspiciousPeeps)):
			print(formatPerson(self.suspiciousPeeps[x]))

	def deleteSuspiciousPerson(self):
		for person in self.suspiciousPeeps:
			YorN = raw_input('Delete {} from contestant list? (y or n)'.format(formatPerson(person)))				
			if YorN.upper() == 'Y':					
				self.contestants.remove(person)


def formatPerson(person):
		return '{} {} {} {}'.format(person[1],person[0],person[2],person[3])

class Winners:

	def __init__(self,contestantList,numberofWinners):
		self.contestantList = contestantList
		self.winners = []
		self.numberofWinners = numberofWinners
		self.findWinners()
		self.printWinners()

	def findWinners(self):
		for x in range(self.numberofWinners):
			winner = random.choice(self.contestantList)
			self.winners.append(winner)
			self.contestantList.remove(winner)	

	def printWinners(self):
		for x in range(len(self.winners)):
			print('The winner is :{} from {}'.format(formatPerson(self.winners[x]),zips[self.winners[x][3][:3]]))	

contest = Contestants()
contest.printContestants()
print("")
contest.printCheaters()
print("")
contest.printSuspiciousPeople()
print("")
Winners(contest.contestants,2)



