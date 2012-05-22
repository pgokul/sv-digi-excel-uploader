import DLID
import urllib2
import urllib
import urlparse
import string
import os
import errno
import sys
import argparse
import clr
import itertools

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

INPUT_FILE = 'if'
OUTPUT_FOLDER = 'op'
SOFTCOPY = 'fn'
DLI_BARCODE = 'dli'
OTHER_LINKS = 'ol'
PRABHANDAM = 'pn'
AUTHOR = 'an'
PUBLISHER = 'pbn'
MARAN_LINK = 'mdl'

class excelHandler:
    def __init__(self):
        self.__excel__ = Excel.ApplicationClass()
        self.__excel__.visible = True

    def open(self,fileName):
        self.__workBook__ = self.__excel__.Workbooks.open(fileName)
        name = self.__workBook__.name
        self.__workSheets__ = self.__workBook__.Worksheets
        title = name[0:name.rfind('.')]
        return title

    def getSheets(self):
        sheets = []
        for sheet in self.__workSheets__:
            name = sheet.name
            sheets.append(sheet)
        
        return sheets

    def makeSheetActive(self,sheet):
        sheet.Activate()

def parseOptions():
    parser = argparse.ArgumentParser
    parser.add_argument('inputFile','-if',dest='if',required=True,help='The input excel file to be read')
    args = parser.parse_args();
    return args

def makeLower(args):
    lowerArgs = {}
    for k,v in args.iteritems():
        lowerArgs[k] = v.lower()
    return lowerArgs
              

def getHeaderCols(usedRange,args):
    cols = {}
    numRows = usedRange.Rows.count
    numCols = usedRange.Columns.count
    
    for ir in range(1,2):
        for ic in range(1,numCols+1):
            cellVal = usedRange.Cells(ir,ic).Value()
            if(cellVal == None):
                continue

            lowerCellVal = cellVal.lower()
            for k,v in args.iteritems():
                if(v == lowerCellVal):
                    cols[k] = ic

    return cols

def mkdir_p(path):
    try:
        os.makedirs(path)
    except OSError as exc:
        if exc.errno == errno.EEXIST:
            pass
        else: raise

def getValueOrDefaultValue(cellVal,defaultVal):
    value = defaultVal
    temp = cellVal.Value()
    if(temp != None and temp != ""):
        value = temp
    return value

def url2name(url):
    return os.path.basename(urllib.unquote(urlparse.urlsplit(url)[2]))

def downloadURL(url,dirPath):
    localName = url2name(url)
    req = urllib2.Request(url)
    r = urllib2.urlopen(req)
    if r.info().has_key('Content-Disposition'):
        # If the response has Content-Disposition, we take file name from it
        localName = r.info()['Content-Disposition'].split('filename=')[1]
        if localName[0] == '"' or localName[0] == "'":
            localName = localName[1:-1]
    elif r.url != url: 
        # if we were redirected, the real file name we take from the final URL
        localName = url2name(r.url)
    fullpath = os.path.join(dirPath,localName)
    f = open(fullpath, 'wb')
    f.write(r.read())
    f.close()

def handleSheet(excel,sheet,baseDir,args):
    name = sheet.name;
    sheetDir = os.path.join(baseDir,name)
    excel.makeSheetActive(sheet)
    usedRange = sheet.UsedRange
    numRows = usedRange.Rows.count
    curPrabhandam = None

    headerCols = getHeaderCols(usedRange,args)
    # try to get the coloum number that has the value "SoftCopy"
    for ir in range(2,numRows+1):
        curPrabhandam = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[PRABHANDAM]),curPrabhandam)
        if(getValueOrDefaultValue(usedRange.Cells(ir,headerCols[SOFTCOPY]),"N") == "Y"):
            author = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[AUTHOR]),"Unknown")
            publisher = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[PUBLISHER]),"Unknown")
            filepath = os.path.join(sheetDir,curPrabhandam,author,publisher)
            dliCode = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[DLI_BARCODE]),"Unknown")
            mdlLink = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[MARAN_LINK]),"Unknown")
            otherURL = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[OTHER_LINKS]),"Unknown")
            mkdir_p(filepath)
            if(dliCode != "Unknown"):
                filename = string.join([curPrabhandam,author,publisher,dliCode],"-")
                filename = string.join([filename,"pdf"],".")
                fullpath = os.path.join(filepath,filename)
                DLID.download(dliCode,fullpath)
            if(mdlLink != "Unknown"):
                downloadURL(mdlLink,filepath)
            if(otherURL!="Unknown"):
                urls = string.split(otherURL,";")
                map(downloadURL,urls,[filepath] * len(urls))
            #We have a soft copy

def main():
    args = {}
    if(len(sys.argv) != 1):
        args = parseOptions();
    else:
        args[INPUT_FILE]= 'F:\DhivyaPrabhandam.xlsx'
        args[OUTPUT_FOLDER] = 'F:\output'
        args[SOFTCOPY] = 'softcopy'
        args[DLI_BARCODE] = 'DLI Barcode'
        args[MARAN_LINK] =  """Maran's Dog link"""
        args[OTHER_LINKS] = 'Other Links'
        args[PRABHANDAM] = 'Prabhandam'
        args[AUTHOR] = 'Author'
        args[PUBLISHER] = 'Publisher'

    args = makeLower(args)
    basePath = args[OUTPUT_FOLDER]
    excel = excelHandler()

    baseDir = os.path.join(basePath,excel.open(args['if']))
    sheets = excel.getSheets()
    for sheet in sheets:
        handleSheet(excel,sheet,baseDir,args)

    

if(__name__ == '__main__'):
    main()




