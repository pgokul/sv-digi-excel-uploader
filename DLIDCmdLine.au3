#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=DLIDCmdLine.exe
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <File.au3>
#include <Array.au3>
#include "CmdLineArgParser.au3"

Opt("MustDeclareVars",1)
Opt("PixelCoordMode",0)

Func getTempDir()
   Local $lTempFolder
   $lTempFolder = _TempFile(@TempDir, "DLTemp", "") & "\DLID Books"
   DirCreate($lTempFolder)
   Return $lTempFolder
EndFunc

Func execDLID($path="C:\PR Labs\DLI Downloader 0.23\DLIDownloader0.23.exe",$title="DLI Downloader 0.23")
   Run($path)
   WinWaitActive($title)
   Return 1
EndFunc

Func closeDLID($title="DLI Downloader 0.23",$popupTitle="Select an Option")
   WinClose($title)
   WinWaitActive($popupTitle)
   Send("{SPACE}")
EndFunc

Func changeDLIDBooksRegKey($newval,$path="HKEY_CURRENT_USER\SOFTWARE\DLIDownloader",$key="Download Directory")
   Local $oldval = RegRead($path, $key)
   Local $retVal = RegWrite($path, $key,"REG_SZ",$newval)
   Return $oldval
EndFunc

Func ExactPixelMatch()
   Local $retCode = 2;
   While(1)
	  Sleep(2000)
	  Local $color = Hex(PixelGetColor(64,344),6)
	  if($color = "009900") Then
		 ;MsgBox(4096, "Done:", $color)
		 $retCode = 0
		 ExitLoop
	  ElseIf($color = "FF0000") Then
		 ;MsgBox(4096, "Error:", $color)
		 $retCode = 1
		 ExitLoop
	  EndIf
   WEnd
   Return $retCode
EndFunc

Func BoolPixelSearch($left,$top,$right,$bottom,$color)
   PixelSearch($left,$top,$right,$bottom,$color)
   if(@error = 1) Then
	  Return False
   EndIf
   Return True
EndFunc


Func PixelBlockMatch()
   Local $retCode = 2;
   Local $left = 63
   Local $top = 340
   Local $right = 83
   Local $bottom = 350
   While(1)
	  Sleep(2000)
	  ; Search for green color
	  if(BoolPixelSearch($left,$top,$right,$bottom,0x009900) = True) Then
		 $retCode = 0
		 ExitLoop
	  ElseIf(BoolPixelSearch($left,$top,$right,$bottom,0xFF0000) = True) Then
		 $retCode = 1
		 ExitLoop
	  Elseif(BoolPixelSearch($left,$top,$right,$bottom,0xFF00FF) = True) Then
		 $retCode = 2
	  Else
		 $retCode =2
		 ;ExitLoop
	  EndIf
   WEnd
   Return $retCode
EndFunc

Func DownloadBook($barCode,$startPage=0,$endPage=0,$matchMethod="PixelBlockMatch")
   Send($barCode)
   Send("{TAB}")
   Send("{DEL}")
   if($startPage <> 0) Then
	  Send($startPage)
   EndIf
   if($endPage <> 0) Then
	  Send("-")
	  Send($endPage)
   EndIf
   Send("{TAB}")
   Send("{SPACE}")
   Local $retCode = Call($matchMethod)
   if(@error <> 0xDEAD) Then
	  Return $retCode
   Else
	  Return PixelBlockMatch()
   EndIf
EndFunc

Func PrintHelp($args="")
   ConsoleWrite("usage --barcode DLIBarCode --output OutputFileOrDirectory --help PrintHelp")
   Exit(0)
EndFunc

Func getArgOrDefaultArg($Arg,$DefaultArg)
   Local $argValue = CmdLine($Arg)
   if($argValue = False) Then
	  $argValue = $DefaultArg
   Endif
   Return $argValue
EndFunc


;Local $dliBarCode = 1990020047793
;Local $DestPath = @DesktopDir
Local $vCmdArgs[1]
$vCmdArgs[0] = "barcode"
_ArrayAdd($vCmdArgs, "output")
_ArrayAdd($vCmdArgs, "start")
_ArrayAdd($vCmdArgs, "end")
_ArrayAdd($vCmdArgs, "DLIDPath")
_ArrayAdd($vCmdArgs, "DLIDTitle")

_CmdLine_Parse("--",$vCmdArgs,"PrintHelp")
Local $dliBarCode = getArgOrDefaultArg("barcode",False)
Local $DestPath = getArgOrDefaultArg("output",False)

if($dliBarCode = False or $DestPath = False) Then
   PrintHelp()
EndIf

Local $DLIDPath = getArgOrDefaultArg("DLIDPath","C:\PR Labs\DLI Downloader 0.23\DLIDownloader0.23.exe")
Local $DLIDTitle = getArgOrDefaultArg("DLIDTitle","DLI Downloader 0.23")
Local $startPage = getArgOrDefaultArg("start","")
Local $endPage = getArgOrDefaultArg("end","")

if(1) Then

   Local $tempDir = getTempDir()
   Local $oldRegVal = changeDLIDBooksRegKey($tempDir)
   if(execDLID($DLIDPath,$DLIDTitle)) Then
	  Local $bookDownload = DownloadBook($dliBarCode,$startPage,$endPage)
	  closeDLID()
	  if($bookDownload = 0) Then
		 Local $FileList = _FileListToArray($tempDir,"*.pdf")
		 Local $SrcFilePath = $tempDir & "\" & $FileList[1]
		 FileCopy($SrcFilePath, $DestPath, 9)
	  EndIf
   EndIf
   changeDLIDBooksRegKey($oldRegVal)
EndIf
