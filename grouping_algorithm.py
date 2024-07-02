from openpyxl import load_workbook

class Grouping_Algorithm:
    def __init__(self):
        pass
    def getWorkbook(self, filename):
        file = "team_data/" + filename + ".xlsx"
        try:
            workbook = load_workbook(filename=file)
        except:
            print("error: could not find file named: " + filename + ".xlsx")
            return False
        sheet = workbook.active
        return sheet
    def putInEvent(self, people, sheet):
        pass
    def sortSeniors(self, sheet):
        dictionary = {}
        counter = 2
        key = "D" + str(counter)
        doneSorting = False
        while doneSorting == False:
            if sheet[key].value == "Sr":
                nameKey = "A" + str(counter)
                eventKey = "C" + str(counter)
                name = sheet[nameKey].value
                events = sheet[eventKey].value
                dictionary[name] = events
            if sheet[key].value == None:
                doneSorting = True
            counter = counter + 1
            key = "D" + str(counter)
        return dictionary
    def sortJuniors(self, sheet):
        dictionary = {}
        counter = 2
        key = "D" + str(counter)
        doneSorting = False
        while doneSorting == False:
            if sheet[key].value == "J":
                nameKey = "A" + str(counter)
                eventKey = "C" + str(counter)
                name = sheet[nameKey].value
                events = sheet[eventKey].value
                dictionary[name] = events
            if sheet[key].value == None:
                doneSorting = True
            counter = counter + 1
            key = "D" + str(counter)
        return dictionary
    def sortSophomores(self, sheet):
        dictionary = {}
        counter = 2
        key = "D" + str(counter)
        doneSorting = False
        while doneSorting == False:
            if sheet[key].value == "S":
                nameKey = "A" + str(counter)
                eventKey = "C" + str(counter)
                name = sheet[nameKey].value
                events = sheet[eventKey].value
                dictionary[name] = events
            if sheet[key].value == None:
                doneSorting = True
            counter = counter + 1
            key = "D" + str(counter)
        return dictionary
    def sortFreshman(self, sheet):
        dictionary = {}
        counter = 2
        key = "D" + str(counter)
        doneSorting = False
        while doneSorting == False:
            if sheet[key].value == "F":
                nameKey = "A" + str(counter)
                eventKey = "C" + str(counter)
                name = sheet[nameKey].value
                events = sheet[eventKey].value
                dictionary[name] = events
            if sheet[key].value == None:
                doneSorting = True
            counter = counter + 1
            key = "D" + str(counter)
        return dictionary
            


def getFileName():
    print("enter in the name of the .xlsx file that you want to generate groups and teams for")
    userInput = input(".xlsx file with team data: ")
    return userInput

filename = getFileName()
algorithm = Grouping_Algorithm()
sheet = algorithm.getWorkbook(filename)
if sheet:
    seniors = algorithm.sortSeniors(sheet=sheet)
    juniors = algorithm.sortJuniors(sheet=sheet)
    sophomores = algorithm.sortSophomores(sheet=sheet)
    freshmen = algorithm.sortFreshman(sheet=sheet)
    
    
#workbook = load_workbook(filename="team_data/Scioly test data.xlsx")


    


