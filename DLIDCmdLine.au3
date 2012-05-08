#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <File.au3>
#include <Array.au3>

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

Func DownloadBook($barCode,$startPage=0,$endPage=0)
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



Local $dliBarCode = 1990020047793
Local $DestPath = @DesktopDir

Local $tempDir = getTempDir()
Local $oldRegVal = changeDLIDBooksRegKey($tempDir)
execDLID()
Local $bookDownload = DownloadBook($dliBarCode,25,45)
closeDLID()
changeDLIDBooksRegKey($oldRegVal)

if($bookDownload = 0) Then
   Local $FileList = _FileListToArray($tempDir,"*.pdf")
   Local $SrcFilePath = $tempDir & "\" & $FileList[1]
   FileCopy($SrcFilePath, $DestPath, 9)
EndIf

