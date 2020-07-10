import re
import xlrd
import time
from collections import defaultdict

def main():
    global doiDict, locDict
    doiDict= {}
    locDict = defaultdict(list) #TODO FINISH BELOW
    wb = xlrd.open_workbook("RAISE_NO_MERGE.xlsx")
    delLST = []


    #loop through the three sheets
    for i in range(1): #TODO SUB IN wb.nsheets-1
        sheet = wb.sheet_by_index(11)
        for col in range(1,3): # cols B & C
            for row in range(4, sheet.nrows):
                cell = sheet.cell_value(row, col)
                del_flag = extractDOI(cell, sheet, row, col)
                #cell is dup, grab for deletion lst
                if del_flag:
                    delLST.append([sheet, row, col])

    print("There are %d citations" %(len(doiDict)))
    print("There are %d duplicates" %(len(delLST)))
    print(delLST)
"""
    for tup in delLST:
        sheet, row, col = tup
        cell = sheet.cell_value(row, col)
        print("----------------------")
        print(cell)
        print("----------------------")


        print(sorted(doiDict.items(), key=lambda x: x[1]))
        print(delLST)
"""


"""
This fcn uses regex to grab the doi number and add it to the global hashtable

@param: cell with citation data
@ret:
- True if doi is already in hashtable
- False if doi is new
"""
def extractDOI(cell, sheet, row, col):
    if cell == "":
        return False
    print(cell)

    chunkACM = re.search('.ca/(.*)', cell) #search for the doi ACM style
    if chunkACM:
        doi = chunkACM.group(1)
    else:
        chunkIEEE = re.search('doi: (.*)URL', cell)
        if chunkIEEE:  #it must be IEEE style
            doi = chunkIEEE.group(1)
        else:
            print("-----------------------------------------")
            print("NO DOI IN IEE OR ACM")
            print(cell)
            print("-------------------------------------------")
            time.sleep(5)
            doi = cell[0:10]

    doi = doi.strip()
    global doiDict, locDict
    doiDict[doi] = doiDict.setdefault(doi, 0) + 1
    locDict[doi].append([sheet, row, col])
    if doiDict[doi] > 1:
        return True # this doi is a duplicate
    else:
        return False

main()
