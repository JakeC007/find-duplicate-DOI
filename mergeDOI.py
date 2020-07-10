import re
import xlrd

def main():
  global doiDict
  doiDict= {}
  wb = xlrd.open_workbook("RAISE_NO_MERGE.xlsx")
  delLST = []


  #loop through the three sheets
  for i in range(1): #TODO SUB IN wb.nsheets-1
    sheet = wb.sheet_by_index(11)
    for col in range(1,3): # cols B & C
        for row in range(4, sheet.nrows):
            cell = sheet.cell_value(row, col)
            del_flag = extractDOI(cell)
            if del_flag: #cell is dup, grab for deletion lst
                delLST.append([sheet, row, col])


  for tup in delLST:
      sheet, row, col = tup
      cell = sheet.cell_value(row, col)
      print("----------------------")
      print(cell)
      print("----------------------")

  print("There are %d number of citations" %(len(doiDict)))
  print(sorted(doiDict.items(), key=lambda x: x[1]))
  print(delLST)


"""
This fcn uses regex to grab the doi number and add it to the global hashtable

@param: cell with citation data
@ret:
    - True if doi is already in hashtable
    - False if doi is new
"""
def extractDOI(cell):
    doi = ""
    if cell == "":
        return False
    chunkACM = re.search('.ca/(.*)', cell) #search for the doi ACM style
    if chunkACM:
      doi = chunkACM.group(1)
    else: #it must be IEEE style
      chunkIEEE = re.search('doi: (.*)URL', cell)
      if chunkIEEE:
        doi = chunkIEEE.group(1)
    doi = doi.strip()
    global doiDict
    doiDict[doi]= doiDict.setdefault(doi, 0) + 1
    if doiDict[doi] > 1:
        print(doi)
        return True # this doi is a duplicate
    else:
        return False

main()
