import re
import xlrd, xlwt, xlutils
import time

from collections import defaultdict
from xlutils.copy import copy

def main():
    global doiDict, locDict
    doiDict= {}
    locDict = defaultdict(list) #TODO FINISH BELOW
    wb = xlrd.open_workbook("RAISE_NO_MERGE.xlsx")
    sb = copy(wb) #sb is the excell sheet to save
    delLST = []


    #loop through the three sheets
    for i in range(wb.nsheets-1): #TODO SUB IN wb.nsheets-1
        sheet = wb.sheet_by_index(i)
        for col in range(1,3): # cols B & C
            for row in range(4, sheet.nrows):
                cell = sheet.cell_value(row, col)
                del_flag = extractDOI(cell, i, row, col)
                #cell is dup, grab for deletion lst
                if del_flag:
                    delLST.append([i, row, col])

    print("There are %d citations" %(len(doiDict)))
    print("There are %d duplicates" %(len(delLST)))
    print(sorted(doiDict.items(), key=lambda x: x[1]))
    # print(delLST)
    print("List of Duplicates:")
    for entry in delLST:
        index, row, col = entry
        printCell(wb, index, row, col)
        deleteCell(sb, index, row, col)

"""
delets a cell
@param: index - the sheet's index in the excell file. Sheet 1 is index = 0
@param: row
@param: col
@param: sb - the write to workbook
"""
def deleteCell(sb, index, row, col):
    sheet = sb.get_sheet(index)
    sheet.write(row, col, "overwritten")
    sb.save('RAISE_MERGED.xlsx')


"""
prints out a cell and the sheet it is from
@param: index - the sheet's index in the excell file. Sheet 1 is index = 0
@param: row
@param: col
@param: wb - read from the workbook
@return: NONE
"""
def printCell(wb, index, row, col):
    sheet = wb.sheet_by_index(index)
    cell = sheet.cell_value(row, col)
    sheetName = wb.sheet_names()
    print("Sheet Title: %s\n" %(sheetName[index]))
    print(cell)
    print("-------------------------------------\n")



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
    #print(cell)

    chunkACM = re.search('.ca/(.*)', cell) #search for the doi ACM style
    if chunkACM:
        doi = chunkACM.group(1)
    else:
        chunkIEEE = re.search('doi: (.*)URL', cell)
        if chunkIEEE:  #it must be IEEE style
            doi = chunkIEEE.group(1)
        else:
            # print("-----------------------------------------")
            # print("NO DOI IN IEE OR ACM")
            # print(cell)
            # print("-------------------------------------------")
            # time.sleep(5)
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
