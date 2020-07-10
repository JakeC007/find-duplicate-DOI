import re 

def main():
  outputlst = []
  doiDict = {}
 
  #read from file
  fInput = open('input.txt', 'r')
  lines = fInput.read().splitlines()
  fInput.close()

  for line in lines:
    if line == "": #skip new char chars
      continue
    chunkACM = re.search('.ca/(.*)', line) #search for the doi ACM style
    if chunkACM:
      doi = chunkACM.group(1)
    else: #it must be IEEE style  
      chunkIEEE = re.search('doi: (.*) URL', line)
      if chunkIEEE:
        doi = chunkIEEE.group(1)
    doi = doi.strip()
    print(doi)
    doiDict[doi]= doiDict.setdefault(doi, 0) + 1
  
  print(sorted(doiDict.items(), key=lambda x: x[1]))
  
main()
