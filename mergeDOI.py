"""
A fun lil program that reads in an .xlsx file and grabs data from relevant cells
it then uses regex to hunt for DOIs which it uses to find and overwrite
duplicate citations.

Author: Jake Chanenson
Date: July 10, 2020
Written for the RASE Lab
"""
import re
import xlrd, xlwt, xlutils
import time

from collections import defaultdict
from xlutils.copy import copy

def main():
    global doiDict, locDict
    doiDict= {}
    locDict = defaultdict(list)
    wb = xlrd.open_workbook("RAISE_NO_MERGE.xlsx")
    sb = copy(wb) #sb is the excell sheet to save
    delLST = []

    #loop through every sheet except for last because it has stats on it
    for i in range(wb.nsheets-1):
        sheet = wb.sheet_by_index(i)
        for col in range(1,3): # cols B & C
            for row in range(4, sheet.nrows): #start in row 5 b/c header
                cell = sheet.cell_value(row, col)
                del_flag, doi = extractDOI(cell, i, row, col)
                #cell is dup, grab for deletion lst
                if del_flag:
                    delLST.append([i, row, col, doi])

    print("There are %d citations" %(len(doiDict)))
    print("There are %d duplicates" %(len(delLST)))
    print(sorted(doiDict.items(), key=lambda x: x[1]))
    print("List of Duplicates:")
    for entry in delLST:
        index, row, col, doi = entry
        printCell(wb, index, row, col)
        deleteCell(sb, index, row, col, doi)

    #a dash of interactivity to view a doi's duplicates
    while(True):
        cmd = input("\n\nPlease enter doi to see all entries or type 'q' to exit: ")
        if cmd == "q":
            break
        else:
            showDups(wb, cmd)

def showDups(wb, doi):
    """
    Prints out the cells that contain a given doi
    @param: wb - read from the workbook
    @param: doi
    @return: NONE
    """
    try:
        dupLST = locDict[doi]
    except:
        return
    print("There are %d cells with that doi" %(len(dupLST)))
    for entry in dupLST:
        index, row, col = entry
        printCell(wb, index, row, col)
    print("*-"*20)
    return

def deleteCell(sb, index, row, col, doi):
    """
    Overwrites a cell.
    @param: index - the sheet's index in the excell file. Sheet 1 is index = 0
    @param: row
    @param: col
    @param: doi
    @param: sb - the write to workbook
    """
    sheet = sb.get_sheet(index)
    sheet.write(row, col, "overwritten.\nOrginal DOI: %s" %(doi))
    sb.save('RAISE_MERGED.xls')

def printCell(wb, index, row, col):
    """
    prints out a cell and the sheet it is from
    @param: index - the sheet's index in the excell file. Sheet 1 is index = 0
    @param: row
    @param: col
    @param: wb - read from the workbook
    @return: NONE
    """
    sheet = wb.sheet_by_index(index)
    cell = sheet.cell_value(row, col)
    sheetName = wb.sheet_names()
    print("Sheet Title: %s\n" %(sheetName[index]))
    print(cell)
    print("-------------------------------------\n")

def extractDOI(cell, sheet, row, col):
    """
    This fcn uses regex to grab the doi number and add it to the global hashtable
    @param: cell with citation data
    @ret:
    - True if doi is already in hashtable OR False if doi is new
    - doi
    """
    if cell == "":
        return False, None

    chunkACM = re.search('.ca/(.*)', cell) #search for the doi ACM style
    if chunkACM:
        doi = chunkACM.group(1)
    else:
        chunkIEEE = re.search('doi: (.*)URL', cell)
        if chunkIEEE:  #it must be IEEE style
            doi = chunkIEEE.group(1)
        else:
            doi = cell[0:20]

    doi = doi.strip()
    global doiDict, locDict
    doiDict[doi] = doiDict.setdefault(doi, 0) + 1
    locDict[doi].append([sheet, row, col])
    if doiDict[doi] > 1:
        return True, doi # this doi is a duplicate
    else:
        return False, doi

main()
