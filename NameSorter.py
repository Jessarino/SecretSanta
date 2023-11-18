try:
    import random
    import os
    import pandas as pd
    import xlsxwriter as xlsx
    from progress.spinner import Spinner as Spinner
except ModuleNotFoundError:
    print("Please install libraries 'os', 'pandas', and 'progress' with 'pip install'")
    exit()

#Save Rani as the first choice

class person:
    def __init__(self, name, email):
        self.name = name
        self.email = email
        self.gto = person
        self.gfrom = person
        self.giftable = True

def RecipientChecker(choice):
    for peep in ppllist:
        if peep.gto == choice:
            return False
        
def giftable(ppllist,choice):
    for peep in ppllist:
        if peep.giftable == True:
            return True
    choice.gto = ppllist[0]
    ppllist[0].gfrom = choice.name
    return False

def loopbreakerchecker(prevchoice, choice):
    if (prevchoice.gto == choice) or (choice.gto == prevchoice):
        return False
    else:
        return True
    
def tablemaker(ppllist):
    tablelist = []
    for peep in ppllist:
        print(peep.name + " is giving a gift to " + peep.gto.name)
        tablelist.append([peep.email, peep.gto.name])
    return tablelist

ppllist = []
namelist = []
for workbook in os.listdir():
    if workbook.endswith('Master.xlsx'):
        df = pd.read_excel(workbook,usecols="D,E,H",index_col=None)
        for indx in df.index:
            if df.iloc[indx][2] == "Yes":
                ppllist.append(person(df.iloc[indx][1],df.iloc[indx][0]))
        prevchoice = ppllist[0]
        ppllist[0].giftable = False
        while giftable(ppllist,prevchoice):   

            choice = random.choice(ppllist)
            if (choice.giftable == True) and (prevchoice.name != choice.name)\
                  and loopbreakerchecker(prevchoice,choice) and (choice != ppllist[0]):
                prevchoice.gto = choice
                prevchoice.gto.giftable = False
                prevchoice.gto.gfrom = prevchoice.name
                prevchoice = choice

tabledf = pd.DataFrame(tablemaker(ppllist))
tabledf.columns=["Addressee", "Giftee"]
tabledf.to_excel("MasterTable.xlsx")
input("Enter to close")
                

        

