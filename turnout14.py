import requests
from datetime import date
from dateutil import parser
import xlsxwriter
counties = ['ALA', 'BAK', 'BAY', 'BRA', 'BRE', 'BRO', 'CAL', 'CHA', 'CIT', 'CLA', 'CLL', 'CLM', 'DAD','DES', 'DIX', 'DUV', 'ESC', 'FLA', 'FRA', 'GAD', 'GIL', 'GLA', 'GUL','HAM', 'HAR', 'HEN', 'HER', 'HIG', 'HIL', 'HOL', 'IND', 'JAC', 'JEF', 'LAF','LAK', 'LEE', 'LEO', 'LEV', 'LIB', 'MAD','MAN', 'MON', 'MRN', 'MRT', 'NAS', 'OKA', 'OKE', 'ORA', 'OSC', 'PAL', 'PAS', 'PIN', 'POL', 'PUT', 'SAN', 'SAR', 'SEM', 'STJ', 'STL', 'SUM', 'SUW', 'TAY', 'UNI', 'VOL', 'WAK', 'WAL', 'WAS']

#if VH/VR files are in downloadable form
#STEP 1) download Voter History and Voter Registration files for same year  
#STEP 2) adjust file names in readVRTxtFile() and readVHFile() for the relevant year
#STEP 3) Select for desired election in selectElection()
#STEP 4) Adjust election date in calcAGE()
#STEP 5) Change Excel sheet output name

#If VR/HR files are in HTML form
#STEP 1) switch scrapeVHURL() for readVHFile() in the createVoterIDString() function and scrapeVRURL() for readVRTxtFile() in the registeredVoters() function
#STEP 2) Adjust URL names in scrapeVHURL() and scrapeVRURL() to reflect correct dates 
#STEPS 3,4 and 5 repeat from above

def run():
    listofCountLists = createFullCountyList()
    writeSheet(listofCountLists)

def createFullCountyList():
    spreadsheet= []
    for county in counties:
        votedIDList = createVoterIDString(county)
        registeredVoterDict = registeredVoters(county)
        final = combineLists2(votedIDList, registeredVoterDict)
        summarized = createDemoSummary(final)
        spreadsheet.append(summarized)
    return spreadsheet

def readVRTxtFile(y):
    x = y + '_20141208.txt' #ADJUST for file name
    with open(x) as file:
        data =file.read()
        rowList = data.split('\n')
        listOfLists = []
        for row in rowList:
            vars = row.split('\t') 
            if len(vars) != 38: 
                rowList.remove(row)
            else:
                listOfLists.append(vars)
    return listOfLists

#iterate through list, if list == key, then add key value to a new list , and return new list
def combineLists2(votedIDList , registeredVoterDict):
    newList = []
    for i in votedIDList:
        if str(i) in registeredVoterDict:
            listValue = registeredVoterDict[str(i)]
            newList.append(listValue)
    return newList


def readVHFile(y): #interchangeable with scrapeURL
    x = y + '_H_20141208.txt' #ADJUST for file name
    with open(x) as file:
        data =file.read()
        splitLines = data.split('\n')
    return splitLines

def createVoterIDString(county):
    a = readVHFile(county) #or scrapeURL
    b = filterDataString(a) # b = a list of lists, with all only five variables- working 12/9
    c = selectElection(b)
    return c

def filterDataString(splitLines): #creates a list of lists, only of readable lists
    listOfLists= []
    for i in splitLines: #breaks each line into a list of each ariable (first = county, second = voterID, etc)
        vars = i.split('\t')
        listOfLists.append(vars) # y = list of every line of file, broken into a list of each category (list of lists)
    readableList = [] #clean of incompleteLines
    for i in listOfLists:
        if len(i) == 5:
            readableList.append(i)
    return readableList

def calcAge(bday):
    if bday == '':
        return 0
    bd = parser.parse(bday).date()
    today = date(2014,11,4) #ADJUST for whatever the election date is  
    y = today.year - bd.year
    if today.month < bd.month or today.month == bd.month and today.day < bd.day:
        y -= 1
    return y

def selectElection(listOfReadableLists):  #returns list of IDs for voters who voted in a specific election
    voteList = voteOnly(listOfReadableLists) 
    genList = []
    for a in voteList:
        if a[3] == 'GEN' and a[2]== '11/04/2014': #ADJUST for the election of choice
            voterId =  a[1]
            genList.append(voterId)
    return genList

def voteOnly(genList): #selects only voters who voted
    newList = []
    for vars in genList:
        if vars[-1]== 'A' or vars[-1]=='E' or vars[-1] == 'Y':
            newList.append(vars)
    return newList


def registeredVoters(county):
    a = readVRTxtFile(county) 
    regVoterDict = {}
    for i in a:
        voterid = i[1] 
        regVoterDict[str(voterid)] = i
    return regVoterDict

def createDemoSummary(i):
    totalRegistered=0 #
    aU30 = 0 # Under 30
    aU40 = 0 #30-40
    aU50 = 0 #40-50
    aU65= 0 #50-65
    aO65 = 0 #65+
    B3 = 0
    H4=0
    W5=0
    RaceOth= 0 #American Indian/Alaska Native, Asian, or - coded as 1
    Dem = 0
    Rep = 0
    Ind = 0 #Independent, or otherparty
    BlackDem = 0 #black DEM
    HisDem = 0 #his DEM
    WhiDem= 0  #white DEM
    BlackRep = 0 #black REP
    HisRep=0 #his REP
    WhiRep = 0 #white REP
    BlackInd = 0 #black IND
    HisInd = 0 #his IND
    WhiInd = 0 #white IND
    PH = 0
    U30Black= 0 #age-race--Black
    U40Black= 0 
    U50Black = 0 
    U65Black = 0
    O65Black = 0
    U30White = 0 #age-race--White
    U40White= 0 
    U50White = 0 
    U65White = 0 
    O65White = 0
    U30His= 0  #age-race--His
    U40His = 0 
    U50His = 0 
    U65His = 0 
    O65His = 0
    U30Dem = 0 #age-party--Dem
    U40Dem = 0 
    U50Dem = 0 
    U65Dem = 0 
    O65Dem = 0
    U30Rep= 0  #age-party--Rep
    U40Rep = 0 
    U50Rep = 0 
    U65Rep = 0 
    O65Rep = 0
    U30Ind = 0  #age-party--Ind
    U40Ind = 0 
    U50Ind = 0 
    U65Ind = 0 
    O65Ind = 0
    U30BDem = 0 #age-race-party-BLACK
    U40BDem = 0 
    U50BDem = 0 
    U65BDem = 0 
    O65BDem = 0
    U30BRep = 0
    U40BRep = 0 
    U50BRep = 0 
    U65BRep = 0 
    O65BRep = 0
    U30BInd = 0 
    U40BInd = 0 
    U50BInd = 0 
    U65BInd = 0 
    O65BInd = 0
    U30HDem = 0 #age-race-party-HIS
    U40HDem = 0 
    U50HDem = 0 
    U65HDem = 0 
    O65HDem = 0
    U30HRep = 0 
    U40HRep = 0 
    U50HRep = 0 
    U65HRep = 0 
    O65HRep = 0
    U30HInd = 0 
    U40HInd = 0 
    U50HInd = 0 
    U65HInd= 0 
    O65HInd = 0
    U30WDem = 0 #age-race-party-WHITE
    U40WDem = 0 
    U50WDem = 0 
    U65WDem = 0 
    O65WDem = 0
    U30WRep = 0 
    U40WRep = 0 
    U50WRep = 0 
    U65WRep = 0 
    O65WRep = 0
    U30WInd = 0
    U40WInd = 0 
    U50WInd = 0 
    U65WInd = 0 
    O65WInd = 0
    OtherAge = 0
    for vars in i: #iterates through rows
        Hispanic = False
        Black = False 
        White = False
        otherAge = False 
        U30 = False
        U40 = False
        U50 = False
        U65 = False
        O65 = False
        totalRegistered += 1 #counts the registered voter
        if len(vars) != 38: #returns the problem Row if its not right
            return len(vars) #eiher will give number or URL
        else:
            countyName = vars[0]
            x = calcAge(vars[21]) #age calculate
            if x >=18 and x <= 29: # age sort
                aU30 += 1
                U30 = True
            elif x > 29 and x <=39:#30-39
                aU40 += 1
                U40= True
            elif x > 39 and x <= 49:#40-49
                aU50 += 1
                U50 = True 
            elif x > 49 and x <65:#50-64
                aU65 += 1
                U65 = True
            elif x >= 65: #over 65
                aO65 +=1
                O65 = True
            else:
                otherAge +=1
                otherAge = True
            if vars[20]== '3': #CALCULATE RACE--> then RACE/AGE#counts black
                B3 += 1
                Black = True
                if U30 == True:
                    U30Black +=1
                elif U40 == True:
                    U40Black +=1
                elif U50 == True:
                    U50Black +=1
                elif U65 == True:
                    U65Black +=1
                else:
                    if O65 == True:
                        O65Black +=1
            elif vars[20]== '4': #hispanic-- counts three age buckets, old, medium old, young
                H4 += 1
                Hispanic= True
                if U30 == True:
                    U30His +=1
                elif U40 == True:
                    U40His +=1
                elif U50 == True:
                    U50His +=1
                elif U65 == True:
                    U65His +=1
                else:
                    if O65 == True:
                        O65His +=1
            elif vars[20]== '5': #white-- counts 3 age buckets of whites
                W5 += 1
                White = True
                if U30 == True:
                    U30White +=1
                elif U40 == True:
                    U40White +=1
                elif U50 == True:
                    U50White +=1
                elif U65 == True:
                    U65White +=1
                else:
                    if O65 == True:
                        O65White +=1
            else: #counts other race
                RaceOth +=1
            #CALCULATE PARTY--> THEN PARTY-AGE--> THEN PARTY-RACE (+ AGE PARTY RACE)
            if vars[23] == 'DEM': #party sort
                #PARTY-AGE (DEM)
                Dem += 1
                if U30 == True:
                    U30Dem +=1
                elif U40 == True:
                    U40Dem +=1
                elif U50 == True:
                    U50Dem +=1
                elif U65 == True:
                    U65Dem +=1
                else:
                    if O65 == True:
                        O65Dem +=1
                if Black == True:
                    BlackDem+=1
                    if U30 == True:
                        U30BDem +=1
                    elif U40 == True:
                        U40BDem +=1
                    elif U50 == True:
                        U50BDem +=1
                    elif U65 == True:
                        U65BDem +=1
                    else:
                        if O65 == True:
                            O65BDem +=1
                if Hispanic == True: #PARTYRACEAGE- HIS DEM
                    HisDem +=1
                    if U30 == True:
                        U30HDem +=1
                    elif U40 == True:
                        U40HDem +=1
                    elif U50 == True:
                        U50HDem +=1
                    elif U65 == True:
                        U65HDem +=1
                    else:
                        if O65 == True:
                            O65HDem +=1
                if White == True: #PARTYRACEAGE- WHITE DEM
                    WhiDem+=1
                    if U30 == True:
                        U30WDem +=1
                    elif U40 == True:
                        U40WDem +=1
                    elif U50 == True:
                        U50WDem +=1
                    elif U65 == True:
                        U65WDem +=1
                    else:
                        if O65 == True:
                            O65WDem +=1
            elif vars[23]== 'REP':
                Rep += 1
                if U30 == True:
                    U30Rep +=1
                elif U40 == True:
                    U40Rep +=1
                elif U50 == True:
                    U50Rep +=1
                elif U65 == True:
                    U65Rep +=1
                else:
                    if O65 == True:
                        O65Rep +=1
                if Black == True:
                    BlackRep  +=1
                    if U30 == True:
                        U30BRep +=1
                    elif U40 == True:
                        U40BRep +=1
                    elif U50 == True:
                        U50BRep +=1
                    elif U65 == True:
                        U65BRep +=1
                    else:
                        if O65 == True:
                            O65BRep +=1
                if Hispanic == True: #PARTY-RACE-AGE (HIS REP)
                    HisRep +=1
                    if U30 == True:
                        U30HRep +=1
                    elif U40 == True:
                        U40HRep +=1
                    elif U50 == True:
                        U50HRep +=1
                    elif U65 == True:
                        U65HRep +=1
                    else:
                        if O65 == True:
                            O65HRep +=1
                elif White == True: #PARTY-RACE-AGE (WHITE REP)
                    WhiRep +=1
                    if U30 == True:
                        U30WRep +=1
                    elif U40 == True:
                        U40WRep +=1
                    elif U50 == True:
                        U50WRep +=1 
                    elif U65 == True:
                        U65WRep +=1 
                    else:
                        if O65 == True:
                            O65WRep +=1
            else:
                Ind += 1
                if U30 == True:
                    U30Ind +=1 
                elif U40 == True:
                    U40Ind +=1
                elif U50 == True:
                    U50Ind +=1
                elif U65 == True:
                    U65Ind +=1
                else:
                    if O65 == True:
                        O65Ind +=1
                if Black == True:
                    BlackInd  +=1
                    if U30 == True:
                        U30BInd +=1
                    elif U40 == True:
                        U40BInd +=1
                    elif U50 == True:
                        U50BInd +=1
                    elif U65 == True:
                        U65BInd +=1
                    else:
                        if O65 == True:
                            O65BInd +=1
                if Hispanic == True: #PARTY-RACE-AGE (HIS IND)
                    HisInd +=1
                    if U30 == True:
                        U30HInd +=1
                    elif U40 == True:
                        U40HInd +=1
                    elif U50 == True:
                        U50HInd +=1
                    elif U65 == True:
                        U65HInd +=1
                    else:
                        if O65 == True:
                            O65HInd +=1
                elif White == True: #PARTY-RACE-AGE (WHITE IND)
                    WhiInd +=1
                    if U30 == True:
                        U30WInd +=1
                    elif U40 == True:
                        U40WInd +=1
                    elif U50 == True:
                        U50WInd +=1
                    elif U65 == True:
                        U65WInd +=1 
                    else:
                        if O65 == True:
                            O65WInd +=1
    rowSum = [countyName, totalRegistered, aU30, aU40, aU50, aU65, aO65, B3, H4, W5, RaceOth, Dem, Rep, Ind,BlackDem, HisDem, WhiDem, BlackRep, HisRep, WhiRep, BlackInd, HisInd, WhiInd, U30Black, U40Black, U50Black, U65Black, O65Black, U30White, U40White, U50White, U65White, O65White, U30His,U40His, U50His, U65His, O65His, U30Dem, U40Dem, U50Dem, U65Dem, O65Dem, U30Rep, U40Rep, U50Rep, U65Rep, O65Rep, U30Ind, U40Ind, U50Ind, U65Ind, O65Ind, U30BDem, U40BDem, U50BDem, U65BDem, O65BDem, U30BRep, U40BRep, U50BRep, U65BRep, O65BRep, U30BInd, U40BInd, U50BInd, U65BInd, O65BInd, U30HDem,U40HDem, U50HDem, U65HDem, O65HDem, U30HRep, U40HRep, U50HRep, U65HRep, O65HRep, U30HInd, U40HInd, U50HInd, U65HInd, O65HInd, U30WDem, U40WDem, U50WDem, U65WDem, O65WDem, U30WRep, U40WRep, U50WRep, U65WRep, O65WRep, U30WInd, U40WInd, U50WInd, U65WInd, O65WInd, OtherAge]
    return rowSum 

def writeSheet(spreadsheet):
    workbook = xlsxwriter.Workbook('2014-Turnout Voters.xlsx') #ADJUST! File Name
    worksheet =  workbook.add_worksheet()
    worksheet.set_column(1, 1, 15)
    bold = workbook.add_format({'bold': 1})
    worksheet.write('B1', 'County', bold) #countyName
    worksheet.write('C1', 'Total Registered', bold) #totalRegistered
    worksheet.write('D1', '18-29,', bold) #B3
    worksheet.write('E1', '30-39', bold) #H4
    worksheet.write('F1', '40-49', bold) #W5
    worksheet.write('G1', '50-65', bold) #RaceOth
    worksheet.write('H1', '65+', bold) #Dem
    worksheet.write('I1', 'Black', bold) #Rep
    worksheet.write('J1', 'Hispanic', bold) #Ind/Other
    worksheet.write('K1', 'White', bold) #Age1
    worksheet.write('L1', 'Other Race', bold) #Age2
    worksheet.write('M1', 'Democrat', bold) #Age3
    worksheet.write('N1', 'Republican', bold) #Age4
    worksheet.write('O1', 'Independent', bold) #Age5
    worksheet.write('P1', 'Democrat, Black ', bold) 
    worksheet.write('Q1', 'Democrat, Hispanic ', bold)
    worksheet.write('R1', 'Democrat, White ', bold)
    worksheet.write('S1', 'Republican, Black ', bold)
    worksheet.write('T1', 'Republican, Hispanic ', bold)
    worksheet.write('U1', 'Republican, White ', bold)
    worksheet.write('V1', 'Independent, Black ', bold)
    worksheet.write('W1', 'Independent, Hispanic', bold)
    worksheet.write('X1', 'Independent, White ', bold)
    worksheet.write('Y1', 'Black, 18-29', bold)
    worksheet.write('Z1', 'Black, 30-39', bold)
    worksheet.write('AA1', 'Black, 40-49', bold)
    worksheet.write('AB1', 'Black, 50-65', bold)
    worksheet.write('AC1', 'Black, 65+', bold)
    worksheet.write('AD1', 'White, 18-29', bold)
    worksheet.write('AE1', 'White, 30-39', bold)
    worksheet.write('AF1', 'White, 40-49', bold)
    worksheet.write('AG1', 'White, 50-65', bold)
    worksheet.write('AH1', 'White, 65+', bold)
    worksheet.write('AI1', 'Hispanic, 18-29', bold)
    worksheet.write('AJ1', 'Hispanic, 30-39', bold)
    worksheet.write('AK1', 'Hispanic, 40-49', bold)
    worksheet.write('AL1', 'Hispanic, 50-65', bold)
    worksheet.write('AM1', 'Hispanic, 65+', bold)
    worksheet.write('AN1', 'Democrat, 18-29', bold)
    worksheet.write('AO1', 'Democrat, 30-39', bold)
    worksheet.write('AP1', 'Democrat, 40-49', bold)
    worksheet.write('AQ1', 'Democrat, 50-65', bold)
    worksheet.write('AR1', 'Democrat, 65+', bold)
    worksheet.write('AS1', 'Republican, 18-29', bold)
    worksheet.write('AT1', 'Republican, 30-39', bold)
    worksheet.write('AU1', 'Republican, 40-49', bold)
    worksheet.write('AV1', 'Republican, 50-59', bold)
    worksheet.write('AW1', 'Republican 65+', bold)
    worksheet.write('AX1', 'Independent, 18-29', bold)
    worksheet.write('AY1', 'Independent, 30-39', bold)
    worksheet.write('AZ1', 'Independent, 40-49', bold)
    worksheet.write('BA1', 'Independent, 50-59', bold)
    worksheet.write('BB1', 'Independent, 65+', bold)
    worksheet.write('BC1', 'Black Democrat, 18-29', bold)
    worksheet.write('BD1', 'Black Democrat, 30-39', bold)
    worksheet.write('BE1', 'Black Democrat, 40-49', bold)
    worksheet.write('BF1', 'Black Democrat, 50-65', bold)
    worksheet.write('BG1', 'Black Democrat, 65+', bold)
    worksheet.write('BH1', 'Black Republican, 18-29', bold)
    worksheet.write('BI1', 'Black Republican, 30-39', bold)
    worksheet.write('BJ1', 'Black Republican, 40-49', bold)
    worksheet.write('BK1', 'Black Republican, 50-59', bold)
    worksheet.write('BL1', 'Black Republican 65+', bold)
    worksheet.write('BM1', 'Black Independent, 18-29', bold)
    worksheet.write('BN1', 'Black Independent, 30-39', bold)
    worksheet.write('BO1', 'Black Independent, 40-49', bold)
    worksheet.write('BP1', 'Black Independent, 50-59', bold)
    worksheet.write('BQ1', 'Black Independent, 65+', bold)
    worksheet.write('BR1', 'Hispanic Democrat, 18-29', bold)
    worksheet.write('BS1', 'Hispanic Democrat, 30-39', bold)
    worksheet.write('BT1', 'Hispanic Democrat, 40-49', bold)
    worksheet.write('BU1', 'Hispanic Democrat, 50-65', bold)
    worksheet.write('BV1', 'Hispanic Democrat, 65+', bold)
    worksheet.write('BW1', 'Hispanic Republican, 18-29', bold)
    worksheet.write('BX1', 'Hispanic Republican, 30-39', bold)
    worksheet.write('BY1', 'Hispanic Republican, 40-49', bold)
    worksheet.write('BZ1', 'Hispanic Republican, 50-59', bold)
    worksheet.write('CA1', 'Hispanic Republican 65+', bold)
    worksheet.write('CB1', 'Hispanic Independent, 18-29', bold)
    worksheet.write('CC1', 'Hispanic Independent, 30-39', bold)
    worksheet.write('CD1', 'Hispanic Independent, 40-49', bold)
    worksheet.write('CE1', 'Hispanic Independent, 50-59', bold)
    worksheet.write('CF1', 'Hispanic Independent, 65+', bold)
    worksheet.write('CG1', 'White Democrat, 18-29', bold)
    worksheet.write('CH1', 'White Democrat, 30-39', bold)
    worksheet.write('CI1', 'White Democrat, 40-49', bold)
    worksheet.write('CJ1', 'White Democrat, 50-65', bold)
    worksheet.write('CK1', 'White Democrat, 65+', bold)
    worksheet.write('CL1', 'White Republican, 18-29', bold)
    worksheet.write('CM1', 'White Republican, 30-39', bold)
    worksheet.write('CN1', 'White Republican, 40-49', bold)
    worksheet.write('CO1', 'White Republican, 50-59', bold)
    worksheet.write('CP1', 'White Republican 65+', bold)
    worksheet.write('CQ1', 'White Independent, 18-29', bold)
    worksheet.write('CR1', 'White Independent, 30-39', bold)
    worksheet.write('CS1', 'White Independent, 40-49', bold)
    worksheet.write('CT1', 'White Independent, 50-59', bold)
    worksheet.write('CU1', 'White Independent, 65+', bold)
    worksheet.write('CV1', 'Other Age Gap', bold)
    row = 1
    col = 1
    for countyName, totalRegistered, aU30, aU40, aU50, aU65, aO65, B3, H4, W5, RaceOth, Dem, Rep, Ind,BlackDem, HisDem, WhiDem, BlackRep, HisRep, WhiRep, BlackInd, HisInd, WhiInd, U30Black, U40Black, U50Black, U65Black, O65Black, U30White, U40White, U50White, U65White, O65White, U30His,U40His, U50His, U65His, O65His, U30Dem, U40Dem, U50Dem, U65Dem, O65Dem, U30Rep, U40Rep, U50Rep, U65Rep, O65Rep, U30Ind, U40Ind, U50Ind, U65Ind, O65Ind, U30BDem, U40BDem, U50BDem, U65BDem, O65BDem, U30BRep, U40BRep, U50BRep, U65BRep, O65BRep, U30BInd, U40BInd, U50BInd, U65BInd, O65BInd, U30HDem,U40HDem, U50HDem, U65HDem, O65HDem, U30HRep, U40HRep, U50HRep, U65HRep, O65HRep, U30HInd, U40HInd, U50HInd, U65HInd, O65HInd, U30WDem, U40WDem, U50WDem, U65WDem, O65WDem, U30WRep, U40WRep, U50WRep, U65WRep, O65WRep, U30WInd, U40WInd, U50WInd, U65WInd, O65WInd, OtherAge in spreadsheet:
        worksheet.write_string  (row, col, countyName)
        worksheet.write_number (row, col +1, totalRegistered)
        worksheet.write_number(row, col+ 2, aU30)
        worksheet.write_number(row, col+ 3, aU40)
        worksheet.write_number(row, col+ 4, aU50)
        worksheet.write_number(row, col+ 5, aU65)
        worksheet.write_number(row, col+ 6, aO65)
        worksheet.write_number(row, col+ 7, B3)
        worksheet.write_number(row, col+ 8, H4)
        worksheet.write_number(row, col+ 9, W5)
        worksheet.write_number(row, col+ 10, RaceOth)
        worksheet.write_number(row, col+ 11, Dem)
        worksheet.write_number(row, col+ 12, Rep)
        worksheet.write_number(row, col+ 13, Ind)
        worksheet.write_number(row, col+ 14, BlackDem)
        worksheet.write_number(row, col+ 15, HisDem)
        worksheet.write_number(row, col+ 16, WhiDem)
        worksheet.write_number(row, col+ 17, BlackRep)
        worksheet.write_number(row, col+ 18, HisRep)
        worksheet.write_number(row, col+ 19, WhiRep)
        worksheet.write_number(row, col+ 20, BlackInd)
        worksheet.write_number(row, col+ 21, HisInd)
        worksheet.write_number(row, col+ 22, WhiInd)
        worksheet.write_number(row, col+ 23, U30Black)
        worksheet.write_number(row, col+ 24, U40Black)
        worksheet.write_number(row, col+ 25, U50Black)
        worksheet.write_number(row, col+ 26, U65Black)
        worksheet.write_number(row, col+ 27, O65Black)
        worksheet.write_number(row, col+ 28, U30White)
        worksheet.write_number(row, col+ 29, U40White)
        worksheet.write_number(row, col+ 30, U50White)
        worksheet.write_number(row, col+ 31, U65White)
        worksheet.write_number(row, col+ 32, O65White)
        worksheet.write_number(row, col+ 33, U30His)
        worksheet.write_number(row, col+ 34, U40His)
        worksheet.write_number(row, col+ 35, U50His)
        worksheet.write_number(row, col+ 36, U65His)
        worksheet.write_number(row, col+ 37, O65His)
        worksheet.write_number(row, col+ 38, U30Dem)
        worksheet.write_number(row, col+ 39, U40Dem)
        worksheet.write_number(row, col+ 40, U50Dem)
        worksheet.write_number(row, col+ 41, U65Dem)
        worksheet.write_number(row, col+ 42, O65Dem)
        worksheet.write_number(row, col+ 43, U30Rep)
        worksheet.write_number(row, col+ 44, U40Rep)
        worksheet.write_number(row, col+ 45, U50Rep)
        worksheet.write_number(row, col+ 46, U65Rep)
        worksheet.write_number(row, col+ 47, O65Rep)
        worksheet.write_number(row, col+ 48, U30Ind)
        worksheet.write_number(row, col+ 49, U40Ind)
        worksheet.write_number(row, col+ 50, U50Ind)
        worksheet.write_number(row, col+ 51, U65Ind)
        worksheet.write_number(row, col+ 52, O65Ind)
        worksheet.write_number(row, col+ 53, U30BDem)
        worksheet.write_number(row, col+ 54, U40BDem)
        worksheet.write_number(row, col+ 55, U50BDem)
        worksheet.write_number(row, col+ 56, U65BDem)
        worksheet.write_number(row, col+ 57, O65BDem)
        worksheet.write_number(row, col+ 58, U30BRep)
        worksheet.write_number(row, col+ 59, U40BRep)
        worksheet.write_number(row, col+ 60, U50BRep)
        worksheet.write_number(row, col+ 61, U65BRep)
        worksheet.write_number(row, col+ 62, O65BRep)
        worksheet.write_number(row, col+ 63, U30BInd)
        worksheet.write_number(row, col+ 64, U40BInd)
        worksheet.write_number(row, col+ 65, U50BInd)
        worksheet.write_number(row, col+ 66, U65BInd)
        worksheet.write_number(row, col+ 67, O65BInd)
        worksheet.write_number(row, col+ 68, U30HDem)
        worksheet.write_number(row, col+ 69, U40HDem)
        worksheet.write_number(row, col+ 70, U50HDem)
        worksheet.write_number(row, col+ 71, U65HDem)
        worksheet.write_number(row, col+ 72, O65HDem)
        worksheet.write_number(row, col+ 73, U30HRep)
        worksheet.write_number(row, col+ 74, U40HRep)
        worksheet.write_number(row, col+ 75, U50HRep)
        worksheet.write_number(row, col+ 76, U65HRep)
        worksheet.write_number(row, col+ 77, O65HRep)
        worksheet.write_number(row, col+ 78, U30HInd)
        worksheet.write_number(row, col+ 79, U40HInd)
        worksheet.write_number(row, col+ 80, U50HInd)
        worksheet.write_number(row, col+ 81, U65HInd)
        worksheet.write_number(row, col+ 82, O65HInd)
        worksheet.write_number(row, col+ 83, U30WDem)
        worksheet.write_number(row, col+ 84, U40WDem)
        worksheet.write_number(row, col+ 85, U50WDem)
        worksheet.write_number(row, col+ 86, U65WDem)
        worksheet.write_number(row, col+ 87, O65WDem)
        worksheet.write_number(row, col+ 88, U30WRep)
        worksheet.write_number(row, col+ 89, U40WRep)
        worksheet.write_number(row, col+ 90, U50WRep)
        worksheet.write_number(row, col+ 91, U65WRep)
        worksheet.write_number(row, col+ 92, O65WRep)
        worksheet.write_number(row, col+ 93, U30WInd)
        worksheet.write_number(row, col+ 94, U40WInd)
        worksheet.write_number(row, col+ 95, U50WInd)
        worksheet.write_number(row, col+ 96, U65WInd)
        worksheet.write_number(row, col+ 97, O65WInd)
        worksheet.write_number(row, col+ 98, OtherAge)
        row += 1
    workbook.close()


def scrapeVRURL(county): #interchangeable with readFile
    url = 'http://flvoters.com/download/20170131/'+ county + '_H_20170207.txt'#exampleURL--> do not plug in
    html = requests.get(url).text 
    splitLines = html.split('\n')
    listOfLists = []
        for row in splitLines:
            vars = row.split('\t') 
            if len(vars) != 38: 
                rowList.remove(row)
            else:
                listOfLists.append(vars)
    return listOfLists

def scrapeVHURL(county): #interchangeable with readFile
    url = 'http://flvoters.com/download/20170131/'+ county + '_20170207.txt'#exampleURL--> do not plug in
    html = requests.get(url).text 
    splitLines = html.split('\n')
    return splitLines

run()
