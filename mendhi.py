from openpyxl import load_workbook
from operator import itemgetter
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

wb = load_workbook('Mehndi Lottery.xlsx')
for sheet in wb:
	if sheet.title != 'Sheet1':
		wb.remove(sheet)
ws1 = wb.create_sheet("Winners")
cheaterSheet = wb.create_sheet("Cheaters")
ws = wb["Sheet1"]

class Contestants:

	def __init__(self):
		self.contestants = []
		self.cheaters = []
		self.suspiciousPeeps = []
		self.createContestants()
		self.findSuspiciousPeople()
		self.deleteSuspiciousPerson()
		self.publishCheaters()

	def createContestants(self):
		for r in range(2,ws.max_row+1):
			person = []
			person.append(str(ws.cell(row=r,column=2).value).upper().strip().split(' ')[0])
			person.append(str(ws.cell(row=r,column=3).value).upper().strip().split(' ')[0])
			person.append(str(ws.cell(row=r,column=4).value).strip())
			person.append(str(int(ws.cell(row=r,column=5).value)))
			person.append(str(ws.cell(row=r,column=6).value).upper().strip().split(' ')[0])
			if person in self.cheaters:
				pass
			elif person in self.contestants:
				self.cheaters.append(person)
				self.contestants.remove(person)
			else:
				self.contestants.append(person)
	
	def publishCheaters(self):
		for x in range(len(self.cheaters)):
			cheaterSheet.cell(row=x+1,column=1,value=self.cheaters[x][0])
			cheaterSheet.cell(row=x+1,column=2,value=self.cheaters[x][1])
			cheaterSheet.cell(row=x+1,column=3,value=self.cheaters[x][2])
			cheaterSheet.cell(row=x+1,column=4,value=self.cheaters[x][3])
			cheaterSheet.cell(row=x+1,column=5,value=self.cheaters[x][4])

	def findSuspiciousPeople(self):
		for x in range(len(self.contestants)):
			p1 = {self.contestants[x][1],self.contestants[x][4]}
			for y in range(x+1,len(self.contestants)):
				p2 = {self.contestants[y][1],self.contestants[y][4]}
				if len(p1.intersection(p2)) == 2:
					self.suspiciousPeeps.append(self.contestants[x])
					self.suspiciousPeeps.append(self.contestants[y])
					break

	def deleteSuspiciousPerson(self):
		x = 0
		while x < (len(self.suspiciousPeeps)):
			YorN = input('{}\n{}\nDelete from contestant list? (y or n)'.format(formatPerson(self.suspiciousPeeps[x]),formatPerson(self.suspiciousPeeps[x+1])))				
			if YorN.upper() == 'Y':					
				self.contestants.remove(self.suspiciousPeeps[x])
				self.cheaters.append(self.suspiciousPeeps[x])
				self.contestants.remove(self.suspiciousPeeps[x+1])
				self.cheaters.append(self.suspiciousPeeps[x+1])				
				x += 2
			else:
				x += 2

def formatPerson(person):
		return '{} {} {} {} {}'.format(person[0],person[1],person[4],person[3],person[2])

class Winners:

	def __init__(self,contestantList,numberofWinners):
		self.contestantList = contestantList
		self.winners = []
		self.numberofWinners = numberofWinners
		self.findWinners()
		self.publishWinners()

	def findWinners(self):
		for x in range(self.numberofWinners):
			winner = random.choice(self.contestantList)
			self.winners.append(winner)
			self.contestantList.remove(winner)	

	def publishWinners(self):
		sorted(self.winners,key=itemgetter(1))
		for x in range(len(self.winners)):
			ws1.cell(row=x+1,column=1,value=self.winners[x][0])
			ws1.cell(row=x+1,column=2,value=self.winners[x][1])
			ws1.cell(row=x+1,column=3,value=self.winners[x][2])
			ws1.cell(row=x+1,column=4,value=self.winners[x][3])
			ws1.cell(row=x+1,column=5,value=self.winners[x][4])
			ws1.cell(row=x+1,column=6,value=zips[self.winners[x][3][:3]])

contest = Contestants()
winners = Winners(contest.contestants,694)
wb.save('Mehndi Lottery.xlsx')
#batchgeo for heat map



