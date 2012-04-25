import os
import sys
import argparse
import clr
clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

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

def handleSheet(excel,sheet,baseDir,args):
    name = sheet.name;
    sheetDir = os.path.join(baseDir,name)
    excel.makeSheetActive(sheet)
    usedRange = sheet.UsedRange
    numRows = usedRange.Rows.count
    numCols = usedRange.Columns.count
    
    headerRow = None
    softCopyCol = None
    linkCells = []
    # try to get the coloum number that has the value "SoftCopy"
    for ir in range(1,numRows+1):
        for ic in range(1,numCols+1):
            cellVal = usedRange.Cells(ir,ic).Value()
            print(cellVal)
            if(cellVal == None):
                continue
            if(args['fn'] == cellVal.lower()):
                headerRow = ir
                softCopyCol = ic
                print(usedRange.Cells(ir,ic).Value())
    os.makedirs(sheetDir)

def main():
    args = {}
    if(len(sys.argv) != 1):
        args = parseOptions();
    else:
        args['if']= 'c:\DhivyaPrabhandam.xlsx'
        args['op'] = 'c:\output'
        args['fn'] = 'softcopy'
        args['dli'] = 'DLI Barcode'
        args['mdl'] =  """Maran's Dog link"""
        args['ol'] = 'Other Links'

    basePath = args['op']
    excel = excelHandler()

    baseDir = os.path.join(basePath,excel.open(args['if']))
    sheets = excel.getSheets()
    for sheet in sheets:
        handleSheet(excel,sheet,baseDir,args)

    

if(__name__ == '__main__'):
    main()




