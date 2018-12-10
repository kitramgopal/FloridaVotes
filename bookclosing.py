import csv
import xlsxwriter
NumCounties = 67
groupIndex = 0 #put the group index at whatever file you want to read. Download bookclosing files and put them in a group folder

def run():
    f = group[groupIndex]
    a = scrapeCSV(f)
    b = removeBlanks(a) #you dont need to return it, technically
    c = createOtherParty(b)
    d = createRaceParty(c)
    x = delKeys(d)
    y = mergeCounties(x)
    z = sumTotals(y)
    return z

group = ['2018genpartyrace_std.csv', '2016general_partyrace.csv', '2014_cpr.csv', 'gen2012_countypartyrace.csv''2010_countypartyrace.csv']

def scrapeCSV(a):
    with open(a, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        w = list(csv_reader)
        return w

def createOtherParty(dict_list):
    for i in dict_list:
        party = '0'
        namefirst = i['Party Name'].split(' ')
        if namefirst[0] == 'Republican':
            party = 'Rep'
        elif namefirst[0] == '':
            party = 'NULL'
        elif namefirst[0] == 'Democrat': 
            party = 'Dem'
        else: 
            party = 'Other'
        i['party'] = party
    return dict_list

def removeBlanks(dict_list): #blank part for 2018- remove blanks manually for prior years
    for i in dict_list:
        del i['\ufeff']
        if i['County Name']== '':
            dict_list.remove(i)
    return dict_list

def createRaceParty(dict_list):
    for i in dict_list:
        if i['party'] == 'Other':
            i['Oth_AmInd']= i['American Indian or Alaskan Native']
            i['Oth_Asian'] = i['Asian Or Pacific Islander']
            i['Oth_Black'] = i['Black, Not Hispanic']
            i['Oth_His']= i['Hispanic']
            i['Oth_Whi'] = i['White, Not Hispanic']
            i['Oth_Total']= i['Total']
            i['Oth_Unknown']= i['Unknown']
            i['Oth_MR'] = i['Multi-Racial']
            i['Oth_Other']= i['Other']
        elif i['party'] == 'Dem':
            i['Dem_AmInd']= i['American Indian or Alaskan Native']
            i['Dem_Asian'] = i['Asian Or Pacific Islander']
            i['Dem_Black'] = i['Black, Not Hispanic']
            i['Dem_His']= i['Hispanic']
            i['Dem_Whi'] = i['White, Not Hispanic']
            i['Dem_Total']= i['Total']
            i['Dem_Unknown']= i['Unknown']
            i['Dem_MR'] = i['Multi-Racial']
            i['Dem_Other']= i['Other']
        else: 
            i['Rep_AmInd']= i['American Indian or Alaskan Native']
            i['Rep_Asian'] = i['Asian Or Pacific Islander']
            i['Rep_Black'] = i['Black, Not Hispanic']
            i['Rep_His']= i['Hispanic']
            i['Rep_Whi'] = i['White, Not Hispanic']
            i['Rep_Total']= i['Total']
            i['Rep_Unknown']= i['Unknown']
            i['Rep_MR'] = i['Multi-Racial']
            i['Rep_Other']= i['Other']
    return dict_list

def mergeCounties(dictList):
    countyList =[]
    countiesOnce = dictList[0:NumCounties]
    for dictionary in countiesOnce:
        Oth_MR = 0
        Oth_Unknown = 0
        Oth_Other = 0
        Oth_AmInd = 0
        Oth_Asian = 0
        Oth_Whi = 0
        Oth_His = 0
        Oth_Black = 0
        Oth_Total = 0
        for i in dictList[NumCounties:]:
            if dictionary['County Name']== i['County Name']:
                if i['party']== 'Other':
                    Oth_MR += toInt(i['Oth_MR'])
                    Oth_Unknown += toInt(i['Oth_Unknown'])
                    Oth_Other += toInt(i['Oth_Other'])
                    Oth_AmInd += toInt(i['Oth_AmInd'])
                    Oth_Asian += toInt(i['Oth_Asian'])
                    Oth_Whi += toInt(i['Oth_Whi'])
                    Oth_His += toInt(i['Oth_His'])
                    Oth_Black += toInt(i['Oth_Black'])
                    Oth_Total += toInt(i['Oth_Total'])
                dictionary.update(i)
        dictionary['Oth_AmInd'] = Oth_AmInd
        dictionary['Oth_Asian'] = Oth_Asian
        dictionary['Oth_His']= Oth_His
        dictionary['Oth_Black'] = Oth_Black
        dictionary['Oth_Whi']= Oth_Whi
        dictionary['Oth_MR']= Oth_MR
        dictionary['Oth_Unknown'] = Oth_Unknown
        dictionary['Oth_Other'] = Oth_Other
        dictionary['Oth_Total'] = Oth_Total
        countyList.append(dictionary)
    return countyList


def sumTotals(dict_list):
    for i in dict_list:
        i['Tot_Black'] = toInt(i['Rep_Black']) + toInt(i['Dem_Black']) + i['Oth_Black']
        i['Tot_His'] = toInt(i['Rep_His']) + toInt(i['Dem_His']) + i['Oth_His']
        i['Tot_Whi'] = toInt(i['Rep_Whi']) + toInt(i['Dem_Whi']) + i['Oth_Whi']
        i['Tot_AmInd'] = toInt(i['Rep_AmInd']) + toInt(i['Dem_AmInd']) + i['Oth_AmInd']
        i['Tot_Other'] = toInt(i['Rep_Other']) + toInt(i['Dem_Other']) + i['Oth_Other']
        i['Tot_MR'] = toInt(i['Rep_MR']) + toInt(i['Dem_MR']) + toInt(i['Oth_MR'])
        i['Tot_Unknown'] = toInt(i['Rep_Unknown']) + toInt(i['Dem_Unknown']) + toInt(i['Oth_Unknown'])
        i['Tot_Asian'] = toInt(i['Rep_Asian']) + toInt(i['Dem_Asian']) + i['Oth_Asian']
        i['Tot_Total'] = toInt(i['Rep_Total'])+ toInt(i['Dem_Total']) + i['Oth_Total']
    return dict_list

def toInt(stringx):
    if type(stringx) is int:
        return stringx
    else:
        a = stringx.split(',')
        newstring = ''
        for i in a:
            newstring = newstring + i
        x = int(newstring)
        return x

def delKeys(dict_list):
    for i in dict_list:
        del i['Unknown']
        del i['Total']
        del i['Multi-Racial']
        del i['Other']
        del i['White, Not Hispanic']
        del i['Hispanic']
        del i['Black, Not Hispanic']
        del i['Asian Or Pacific Islander']
        del i['American Indian or Alaskan Native']
    return dict_list


def writeSheet(listOfDicts):
    workbook = xlsxwriter.Workbook('2010-1' + 'test.xlsx')
    worksheet =  workbook.add_worksheet()
    worksheet.set_column(1, 1, 15)
    bold = workbook.add_format({'bold': 1})
    worksheet.write('B1', 'County', bold) #totalRegistered
    worksheet.write('C1', 'AmInd Rep', bold)
    worksheet.write('D1', 'AmInd Dem', bold)
    worksheet.write('E1', 'AmInd OP', bold)
    worksheet.write('F1', 'AmInd Total', bold) 
    worksheet.write('G1', 'Asian Rep', bold)
    worksheet.write('H1', 'Asian Dem', bold)
    worksheet.write('I1', 'Asian OP', bold)
    worksheet.write('J1', 'Asian Total', bold) 
    worksheet.write('K1', 'Black Rep', bold)
    worksheet.write('L1', 'Black Dem', bold)
    worksheet.write('M1', 'Black OP', bold)
    worksheet.write('N1', 'Black Total', bold) 
    worksheet.write('O1', 'His Rep', bold)
    worksheet.write('P1', 'His Dem', bold)
    worksheet.write('Q1', 'His OP', bold)
    worksheet.write('R1', 'His Total', bold) 
    worksheet.write('S1', 'White Rep,', bold) #B3
    worksheet.write('T1', 'White Dem', bold) #H4
    worksheet.write('U1', 'White OP,', bold)
    worksheet.write('V1', 'White Total', bold) #B3
    worksheet.write('W1', 'OtherRace Rep', bold)
    worksheet.write('X1', 'OtherRace Dem', bold)
    worksheet.write('Y1', 'OtherRace OP', bold)
    worksheet.write('Z1', 'OtherRace Total', bold)
    worksheet.write('AA1', 'MR Rep', bold)
    worksheet.write('AB1', 'MR Dem', bold)
    worksheet.write('AC1', 'MR OP', bold)
    worksheet.write('AD1', 'MR Total', bold)
    worksheet.write('AE1', 'Unknown Rep', bold)
    worksheet.write('AF1', 'Unknown Dem', bold)
    worksheet.write('AG1', 'Unknown OP', bold)
    worksheet.write('AH1', 'Unknown Total', bold)
    worksheet.write('AI1', 'Rep Total', bold)
    worksheet.write('AJ1', 'Dem Total', bold)
    worksheet.write('AK1', 'Oth Total', bold)
    worksheet.write('AL1', 'Tot_Total', bold)
    row = 1
    col = 1
    for c in listOfDicts:
        worksheet.write_string(row, col, c['County Name'])
        worksheet.write_string(row, col+ 1, c['Rep_AmInd'])
        worksheet.write_string(row, col+ 2, c['Dem_AmInd'])
        worksheet.write_number(row, col+ 3, c['Oth_AmInd'])
        worksheet.write_number(row, col+ 4, c['Tot_AmInd'])
        worksheet.write_string(row, col+ 5, c['Rep_Asian'])
        worksheet.write_string(row, col+ 6, c['Dem_Asian'])
        worksheet.write_number(row, col+ 7, c['Oth_Asian'])
        worksheet.write_number(row, col+ 8, c['Tot_Asian'])
        worksheet.write_string(row, col+ 9, c['Rep_Black'])
        worksheet.write_string(row, col+ 10, c['Dem_Black'])
        worksheet.write_number(row, col+ 11, c['Oth_Black'])
        worksheet.write_number(row, col+ 12, c['Tot_Black'])
        worksheet.write_string(row, col+ 13, c['Rep_His'])
        worksheet.write_string(row, col+ 14, c['Dem_His'])
        worksheet.write_number(row, col+ 15, c['Oth_His'])
        worksheet.write_number(row, col+ 16, c['Tot_His'])
        worksheet.write_string(row, col+ 17, c['Rep_Whi'])
        worksheet.write_string(row, col+ 18, c['Dem_Whi'])
        worksheet.write_number(row, col+ 19, c['Oth_Whi'])
        worksheet.write_number(row, col+ 20, c['Tot_Whi'])
        worksheet.write_string(row, col+ 21, c['Rep_Other'])
        worksheet.write_string(row, col+ 22, c['Dem_Other'])
        worksheet.write_number(row, col+ 23, c['Oth_Other'])
        worksheet.write_number(row, col+ 24, c['Tot_Other'])
        worksheet.write_string(row, col+ 25, c['Rep_MR'])
        worksheet.write_string(row, col+ 26, c['Dem_MR'])
        worksheet.write_number(row, col+ 27, c['Oth_MR'])
        worksheet.write_number(row, col+ 28, c['Tot_MR'])
        worksheet.write_string(row, col+ 29, c['Rep_Unknown'])
        worksheet.write_string(row, col+ 30, c['Dem_Unknown'])
        worksheet.write_number(row, col+ 31, c['Oth_Unknown'])
        worksheet.write_number(row, col+ 32, c['Tot_Unknown'])
        worksheet.write_string(row, col+ 33, c['Rep_Total'])
        worksheet.write_string(row, col+ 34, c['Dem_Total'])
        worksheet.write_number(row, col+ 35, c['Oth_Total'])
        worksheet.write_number(row, col+ 36, c['Tot_Total'])
        row += 1
    workbook.close()
    
run()

