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
    def sortPeople(sheet):
        counter = 2
        key = "C" + str(counter)
        doneSorting = False
        while doneSorting == False:
            pass


def getFileName():
    print("enter in the name of the .xlsx file that you want to generate groups and teams for")
    userInput = input(".xlsx file with team data: ")
    return userInput

filename = getFileName()
algorithm = Grouping_Algorithm()
sheet = algorithm.getWorkbook(filename)
if sheet:
    print(sheet["A1"].value)
    algorithm.sortPeople()
    
#workbook = load_workbook(filename="team_data/Scioly test data.xlsx")


    


