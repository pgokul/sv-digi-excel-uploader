import tempfile
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
COMMENTARY_NAME = 'cn'
DRY_RUN = 'dr'
FORCE_DOWNLOAD = 'fd'
args = {}

def enum(*sequential, **named):
    enums = dict(zip(sequential, range(len(sequential))), **named)
    return type('Enum', (), enums)

DOWNLOAD_ACTION = enum('NO_ACTION','DOWNLOAD','CREATE_FILE')

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

    def close(self):
        self.__workBook__.Close(False)

    def quit(self):
        self.__excel__.Quit()

def parseOptions():
    parser = argparse.ArgumentParser()
    parser.add_argument('--inputFile','-if',dest=INPUT_FILE,required=True,help='The input excel file to be read')
    parser.add_argument('--outputDir','-od',dest=OUTPUT_FOLDER,required=True,help='The output directory')
    parser.add_argument('--SoftCopyColName','-sc',dest=SOFTCOPY,required=False,help='The header text for the soft copy available column',default='softcopy')
    parser.add_argument('--DLIBarcodeColName','-dbc',dest=DLI_BARCODE,required=False,help='The header text for the DLI Barcode column',default='DLI Barcode')
    parser.add_argument('--MaranDogLinkColName','-mdc',dest=MARAN_LINK,required=False,help='The header text for the Maran\'s Dog link column',default='Maran\'s Dog link')
    parser.add_argument('--OtherLinksColName','-olc',dest=OTHER_LINKS,required=False,help='The header text for the Other Links column',default='Other Links')
    parser.add_argument('--PrabhandamColName','-pc',dest=PRABHANDAM,required=False,help='The header text for the Prabhandam column',default='Prabhandam')
    parser.add_argument('--AuthorColName','-ac',dest=AUTHOR,required=False,help='The header text for the Author column',default='Author')
    parser.add_argument('--PublisherColName','-pbc',dest=PUBLISHER,required=False,help='The header text for the Publisher column',default='Publisher')
    parser.add_argument('--CommentaryColName','-cnc',dest=COMMENTARY_NAME,required=False,help='The header text for the Commentary Name column',default='CommentaryName')
    parser.add_argument('--DryRun','-dr',dest=DRY_RUN,required=False,help='Dont Download anything just create folders',default=False,action="store_true")
    parser.add_argument('--ForceDownload','-f',dest=FORCE_DOWNLOAD,required=False,help='Force Download even if file exists',default=False,action="store_true")

    pargs = parser.parse_args();
    return pargs

def makeLower():
    lowerArgs = {}
    for k,v in args.iteritems():
        try:
            lowerArgs[k] = v.lower()
        except:
            lowerArgs[k] = v
            pass
    return lowerArgs
              

def getHeaderCols(usedRange):
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

def shouldDownload(fullpath):
    if(args[DRY_RUN] == True):
        if(os.path.exists(fullpath)):
            return DOWNLOAD_ACTION.NO_ACTION
        else:
            return DOWNLOAD_ACTION.CREATE_FILE
    else:
        if(args[FORCE_DOWNLOAD] == False and  os.path.exists(fullpath) and os.path.getsize(fullpath) !=0):
            return DOWNLOAD_ACTION.NO_ACTION
        return DOWNLOAD_ACTION.DOWNLOAD

def doDownloadAction(downloadFunc,downloadFuncArg,fullpath):
    dac = shouldDownload(fullpath)
    if(dac == DOWNLOAD_ACTION.NO_ACTION):
        return
    elif(dac == DOWNLOAD_ACTION.CREATE_FILE):
        f= open(fullpath,'wb')
        f.close()
        return
    elif(dac == DOWNLOAD_ACTION.DOWNLOAD):
        downloadFunc(downloadFuncArg,fullpath)
        return
    else:
        raise " Unknown download action"
    return
 
def actualDownloadUrl(urlHandle,fullpath):
    tempfilehandle = tempfile.TemporaryFile()
    tempfilehandle.write(urlHandle.read())
    tempfilehandle.seek(0)
    f = open(fullpath, 'wb')
    f.write(tempfilehandle.read())
    f.close()
    tempfilehandle.close()
    return

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
    doDownloadAction(actualDownloadUrl,r,fullpath)
    return
    

def downloadDLI(barcode,dirPath,commentaryName,author):
    filename = string.join([author,commentaryName,barcode],"_")
    filename = string.join([filename,"pdf"],".")
    fullpath = os.path.join(dirPath,filename)
    #doDownloadAction(DLID.download,barcode,fullpath)
    return

def handleSheet(excel,sheet,baseDir):
    name = sheet.name
    sheetDir = os.path.join(baseDir,name)
    excel.makeSheetActive(sheet)
    usedRange = sheet.UsedRange
    numRows = usedRange.Rows.count
    curPrabhandam = None
    CommnentaryName = ""

    headerCols = getHeaderCols(usedRange)
    # try to get the coloum number that has the value "SoftCopy"
    for ir in range(2,numRows+1):
        oldPrabhandamName = curPrabhandam
        curPrabhandam = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[PRABHANDAM]),curPrabhandam)
        if(oldPrabhandamName != curPrabhandam):
            CommnentaryName = ""
        CommnentaryName = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[COMMENTARY_NAME]),CommnentaryName)
        if(getValueOrDefaultValue(usedRange.Cells(ir,headerCols[SOFTCOPY]),"N") == "Y"):
            author = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[AUTHOR]),"Unknown")
            publisher = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[PUBLISHER]),"Unknown")
            filepath = os.path.join(sheetDir,curPrabhandam,CommnentaryName,author,publisher)
            dliCode = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[DLI_BARCODE]),"Unknown")
            mdlLink = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[MARAN_LINK]),"Unknown")
            otherURL = getValueOrDefaultValue(usedRange.Cells(ir,headerCols[OTHER_LINKS]),"Unknown")
            mkdir_p(filepath)
            if(dliCode != "Unknown"):
                barcodes = string.split(dliCode,";")
                map(downloadDLI,barcodes,[filepath] * len(barcodes),[CommnentaryName] * len(barcodes),[author] * len(barcodes))
            if(mdlLink != "Unknown"):
                urls = string.split(mdlLink,";")
                map(downloadURL,urls,[filepath] * len(urls))
            if(otherURL!="Unknown"):
                urls = string.split(otherURL,";")
                map(downloadURL,urls,[filepath] * len(urls))
            #We have a soft copy

def main():
    global args
    if(len(sys.argv) != 1):
        argsns = parseOptions()
        args = vars(argsns)
    else:
        args[INPUT_FILE]= 'F:\SV\sv-digi-uploader\DhivyaPrabhandam.xlsx'
        args[OUTPUT_FOLDER] = 'F:\output'
        args[SOFTCOPY] = 'softcopy'
        args[DLI_BARCODE] = 'DLI Barcode'
        args[MARAN_LINK] =  """Maran's Dog link"""
        args[OTHER_LINKS] = 'Other Links'
        args[PRABHANDAM] = 'Prabhandam'
        args[AUTHOR] = 'Author'
        args[PUBLISHER] = 'Publisher'
        args[COMMENTARY_NAME] = 'CommentaryName'
        args[DRY_RUN] = False

    args = makeLower()
    vargs = args
    basePath = args[OUTPUT_FOLDER]
    excel = excelHandler()

    baseDir = os.path.join(basePath,excel.open(args['if']))
    sheets = excel.getSheets()
    for sheet in sheets:
        try:
            handleSheet(excel,sheet,baseDir)
        except:
            pass
    excel.close()
    excel.quit()

if(__name__ == '__main__'):
    main()




