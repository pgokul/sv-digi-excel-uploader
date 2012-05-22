import subprocess

DLI_PATH = ""
DLI_TITLE = ""
DLID_CMD = "DLIDCmdLine.exe"

def constructArg(arglist,arg,val):
    arglist.append("--"+arg)
    arglist.append(val)
    return

def download(DLIBarcode,pdffile,startpage="10",endpage="40"):
    if(DLIBarcode == ""):
        raise Exception("DLI Barcode is emtpy")
    if(pdffile == ""):
        raise Exception("Output PDF location is empty")
    args = []
    args.append(DLID_CMD)
    if(DLI_PATH != ""):
        constructArg(args,"DLIDPath",DLI_PATH)
    if(DLI_TITLE != ""):
        constructArg(args,"DLIDTitle",DLI_TITLE)
    if(startpage != ""):
        constructArg(args,"start",str(startpage))
    if(endpage != ""):
        constructArg(args,"end",str(endpage))
    constructArg(args,"barcode",DLIBarcode)
    constructArg(args,"output",pdffile)

    subprocess.call(args,shell=False)
    return

if(__name__ == '__main__'):
    download("1990020047793","test.pdf",20,45)