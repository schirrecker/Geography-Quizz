import xlrd
import random
import datetime
import os, os.path
import pickle

# -----------------------------------------------------
# openpyxl syntax: 
# wb = openpyxl.Workbook()
# grab the active worksheet: ws = wb.active
# Data can be assigned directly to cells: ws['A1'] = 42
# Rows can also be appended: ws.append([1, 2, 3])
# Python types will automatically be converted
# import datetime
# ws['A2'] = datetime.datetime.now()
# Save the file - wb.save("sample.xlsx")
#-----------------------------------------------------

# ----------------------------------------------------
# xlrd syntax:
# workbook = xlrd.open_workbook("FILE.xlsx")
# worksheet = workbook.sheet_by_name("<NAME OF SHEET>")
# worksheet.cell(row, col).value)
# worksheet.row_values(row)
# worksheet.nrows
# worksheet.ncols
# Row = Country, Country Code, Continent, Capital,
# Population, Area, Coastline, Government, Currency 
# ----------------------------------------------------

# ----------------------------------------------------
# Class definitions
# ----------------------------------------------------
class Player:
    def __init__(self, name, password="password"):
        self.name = name
        self.password = password
        self.points = 0
        self.level = 1
        self.knowledge = {}
        self.CountriesToTest = []
        self.InitKnowledge()
        self.UpdateCountriesToTest()
        
        # self.knowlege = {"country": [-1, -1, -1, -1, -1, -1]}
        # 0  = not asked yet
        # 1  = correct
        # -1 = wrong
         
    def InitKnowledge(self):
        for country in Countries:
            self.knowledge[country.name] = [-1, -1, -1, -1, -1, -1]

    def PrintKnowledge(self):
        for country in Countries:
            if self.knowledge[country.name][0] == 1:    
                print (country.name + " - continent: " + country.continent)
            if self.knowledge[country.name][1] == 1:    
                print (country.name + " - capital: " + country.capital)            
            if self.knowledge[country.name][2] == 1:    
                print (country.name + " - population: " + str(int(country.population)))               
            if self.knowledge[country.name][3] == 1:    
                print (country.name + " - area: " + str(int(country.area)))              
            if self.knowledge[country.name][4] == 1:    
                print (country.name + " - coastline: " + str(int(country.coastline)))
            if self.knowledge[country.name][5] == 1:    
                print (country.name + " - currency: " + country.currency)                

    # test only countries that player doesn't know
    # 0: continent, 1: capital, 2: pop, 3: area, 4: coastline, 5: currency
    def UpdateCountriesToTest(self, level):
        self.CountriesToTest = []
        for country in CountriesByLevel[str(level)]:
            if self.knowledge[country.name][0] != 1 and self.knowledge[country.name][1] != 1:
                self.CountriesToTest.append(country)

    def QuizRound(self, Nb_Questions, level):
        self.UpdateCountriesToTest(level)
        for i in range(Nb_Questions):
            country = random.choice(self.CountriesToTest)
            print ("")
            self.Quiz(country)
            print ("")
            print ("You have " + str(self.points) + " points")

    def Quiz(self, country):
        # Continent
        # ---------
        txt = "In which continent is the country of " + country.name + " ? " + os.linesep
        txt = txt + "Continents are: "+ " ,".join(Continents) + os.linesep
        InputOk = False
        while not InputOk:
            answer = input(txt)
            if answer in Continents:
                InputOk = True
            else:
                InputOk = False
                print ("please enter from the list of continents")    
        if country.CheckContinent(answer):
            print ("This is correct")
            self.points += 1
            self.knowledge[country.name][0] = 1
        else:
            print ("This is wrong")
            self.knowledge[country.name][0] = 0
        print (" The answer was: ", country.continent)

       # Capital
       # -------
        InputOk = False
        while not InputOk:
            answer = input("What is the capital of " + country.name + " ? ")
            if len(answer) > 1:
                InputOk = True
            else:
                InputOk = False
                print ("print enter a valid city")
                
        if country.CheckCapital(answer):
            print ("This is correct")
            self.points += 1
            self.knowledge[country.name][1] = 1
        else:
            print ("This is wrong")
            self.knowledge[country.name][1] = 0
        print ("Capital: ", country.capital)

        answer = int(input("What is the population of " + country.name + " (in millions ?) ")) * 1000000
        if country.CheckPopulation(answer):
            print ("This is correct")
            self.points += 1
            self.knowledge[country.name][2] = 1
        else:
            print ("This is wrong")
            self.knowledge[country.name][2] = 0
        print ("Population: ", int(country.population))

        answer = int(input("What is the area of " + country.name + " ? "))
        if country.CheckArea(answer):
            print ("This is correct")
            self.points += 1
            self.knowledge[country.name][3] = 1
        else:
            print ("This is wrong")
            self.knowledge[country.name][3] = 0
        print ("Area: ", int(country.area))       

        answer = input("Does " + country.name + " have a coastline ? (y or n)")
        if country.CheckCoastline(answer):
            print ("This is correct")
            self.points += 1
            self.knowledge[country.name][4] = 1
        else:
            print ("This is wrong")
            self.knowledge[country.name][4] = 0
        print ("Coastline: ", int(country.coastline))

        answer = input("What is the currency of " + country.name + " ? ")
        if country.CheckCurrency(answer):
            print ("This is correct")
            self.points += 1
            self.knowledge[country.name][5] = 1
        else:
            print ("This is wrong")
            self.knowledge[country.name][5] = 0
        print ("Currency: ", country.currency)       

class Country:
    def __init__(self, name, continent, capital, population,
                 area, coastline,currency, level):
        self.name = name
        self.continent = continent
        self.capital = capital
        self.population = population
        self.population_margin = 0.3
        self.area = area
        self.area_margin = 0.5
        self.coastline = coastline
        self.coastine_margin = .7
        self.currency = currency
        self.level = int(level)

    def __repr__(self):
        txt = self.name + os.linesep
        txt = txt + "Continent: " + self.continent + os.linesep
        txt = txt + "Capital: " + self.capital + os.linesep
        txt = txt + "Population: " + str(int(self.population/1000000)) + "M" + os.linesep
        txt = txt + "Area: " + str(self.area) + os.linesep
        txt = txt + "Coastline: " + str(self.coastline) + os.linesep
        txt = txt + "Currency: " + self.currency + os.linesep
        txt = txt + "Difficulty: " + str(self.level) + os.linesep
        return txt
      
    def CheckCapital(self, capital):
        return stringsMatch(self.capital, capital)

    def CheckContinent(self, continent):
        return stringsMatch(self.continent, continent)

    def CheckCurrency(self, currency):
        return stringsMatch(self.currency, currency)

    def CheckPopulation(self, population):
        return isclose(self.population, population, self.population_margin)
            
    def CheckArea(self, area):
        if self.area < 1000: 
            return area < 1000
        elif self.area < 10000:
            return isclose(self.area, area, 5)      
        elif self.area < 100000:
            return isclose(self.area, area, 3)     
        elif self.area < 500000:
            return isclose(self.area, area, .7)
        else:
            return isclose(self.area, area, self.area_margin)       

    def CheckCoastline(self, coastline):
        coastline = coastline.lower()[0]     # first letter y or n
        if (self.coastline == 0 and coastline == 'n') or (self.coastline != 0 and coastline == 'y'):
            return True
        else:
            return False
        
# ---------------------------------------------------
# Useful functions
# ---------------------------------------------------
def isclose(a, b, margin):
    if a * (1 - margin) < b < a * (1 + margin):
        return True
    else:
        return False

def stringsMatch(a, b):
    if a.lower() == b.lower():
        return True
    else:
        return False

def approx(a, n):
    return int(a/(10**n)) * (10**n)
    

# -------------------------------------------------
# Load country data from file
# -------------------------------------------------
def LoadCountryData():
    Continents = []
    Countries = []
    CountriesByLevel = {"1": [], "2": [], "3": [], "4": [], "5": []}
    for row in range(1, worksheet.nrows):
        country = worksheet.row_values(row)
        Countries.append(Country(
            country[0],               # name
            country[1],               # continent
            country[2],               # capital
            approx(country[3], 3),    # population
            country[4],               # area
            country[5],               # coastline
            country[6],               # currency
            country[8]))              # difficulty level

    # create levels: dictionary = {"level" : [list of countries]}
    # create continents: list of continents
    for country in Countries:
        CountriesByLevel[str(country.level)].append(country)
        Continents.append(country.continent)
    Continents = list(set(Continents))
    return Countries, CountriesByLevel, Continents

# -------------------------------------------------
# Save player data to file
# -------------------------------------------------
def SavePlayerData():
    global Players
    with open('PlayerData', 'wb') as PlayerFile:
        pickle.dump(Players, PlayerFile)

# -------------------------------------------------
# Load player data from file
# -------------------------------------------------
def LoadPlayerData():
    Players = []
    if os.path.isfile('PlayerData'):
        with open('PlayerData', 'rb') as PlayerFile:
            Players = pickle.load(PlayerFile)
            print ("Player data loaded.")
    else:
        print("No existing player data.  This is the first game.")
    return Players

# ----------------------------------------------------
# Load Excel file and its worksheets
# ----------------------------------------------------
workbook = xlrd.open_workbook("Countries.xlsx")
worksheet = workbook.sheet_by_name("Facts")
Nb_Countries = worksheet.nrows
Nb_Facts = worksheet.ncols - 1
    
# -------------------------------------------------
# Initialize global data, define useful functions
# -------------------------------------------------
Countries = []
CountriesByLevel = {"1": [], "2": [], "3": [], "4": [], "5": []}
Players = []
Continents = []
Nb_Questions = 2
     
# -------------------------------------------------
# Load data
# -------------------------------------------------
Countries, CountriesByLevel, Continents = LoadCountryData()
Players = LoadPlayerData()

print (str(worksheet.nrows) + " countries loaded.")
txt = "Players loaded: "
for player in Players:
    txt = txt + player.name + ", " 
print (txt)

def CheckPlayer(name):
    global Players
    exist = False
    for p in Players:
        if name.lower() == p.name.lower():
            print ("Welcome back, " + name + os.linesep)
            exist = True
            return p
    if not exist:          # if it doesn't exist, create it
        answer = input("Would you like to create this user? (y/n) ")
        if answer[0].lower() == "y":
            p = Player(name)
            Players.append(p)
            return p

# -------------------------------------------------
# Start game loop
# -------------------------------------------------

name = str(input("Enter player name: "))
player = CheckPlayer(name)

print ("Your are at level " + str(player.level))
print ("You currently have " + str(player.points) + " points")

quit = False 
while quit == False:
    play = input("Would you like to play ? (y/n)") 
    if play.lower() == "n":
        quit = True
    else: 
        level = input("what level ?")
        player.QuizRound(Nb_Questions, int(level))
        SavePlayerData()

print("Here is your knowlege so far:")
player.PrintKnowledge()


    
