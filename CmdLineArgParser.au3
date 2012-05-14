#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

Global $CmdLine_Cache[1][2] = [[0, False]]

; #FUNCTION# ====================================================================================================================
; Name ..........: CmdLine
; Description ...: Returns the value of the command line parameter.
; Syntax ........: CmdLine( [ $sParam ] )
; Parameters ....: $sParam          - [optional] A string specifying the command line parameter to retrieve. Default is a blank
;                                     string, which is the parameter that does not have a named flag first.
; Return values .: Success          - If the parameter has a value, then that is returned. If not then True is.
;                  Failure          - False. Usually means that the parameter wasn't in the command line.
; Author(s) .....: Matt Diesel (Mat)
; Modified ......:
; Remarks .......:
; Related .......: _CmdLine_Parse
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func CmdLine($sParam = "")
    If $CmdLine_Cache[0][0] = 0 And $CmdLine[0] > 0 Then _CmdLine_Parse()

    If $sParam = "" Then
        Return $CmdLine_Cache[0][1]
    Else
        For $i = 1 To UBound($CmdLine_Cache) - 1
            If $CmdLine_Cache[$i][0] = $sParam Then Return $CmdLine_Cache[$i][1]
        Next
        Return False
    EndIf
EndFunc   ;==>CmdLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _CmdLine_Parse
; Description ...:
; Syntax ........: _CmdLine_Parse( [ $sPrefix [, $asAllowed [, $sOnErrFunc ]]] )
; Parameters ....: $sPrefix         - [optional] The prefix for command line arguments. Default is '--'. It is recommended that
;                                     it is one of the following standards (although it could be anything):
;                                   |GNU     - uses '--' to start arguments. <a href="http://www.gnu.org/prep/standards/
;                                              html_node/Command_002dLine-Interfaces.html">Page</a>. '--' on it's own sets the
;                                              unnamed argument but always uses the next parameter, even if prefixed by '--'.
;                                              E.g. '-- --file' will mean CmdLine() = '--file'.
;                                   |MS      - uses '-'. Arguments with values are seperated by a colon: ':'. <a
;                                              href="http://technet.microsoft.com/en-us/library/ee156811.aspx">Page</a>
;                                   |Slashes - Not sure where it's a standard, but using either backslash ( '\' ) or slash
;                                              ( '/' ) is fairly common. AutoIt uses it :) Just make sure the user knows which
;                                              one it is.
;                  $asAllowed       - [optional] A zero based array of possible command line arguments that can be used. When an
;                                     argument is found that does not match any of the values in the array, $sOnErrFunc is called
;                                     with the first parameter being "Unrecognized parameter: PARAM_NAME".
;                  $sOnErrFunc      - [optional] The function to call if an error occurs.
; Return values .: None
; Author(s) .....: Matt Diesel (Mat)
; Modified ......:
; Remarks .......:
; Related .......: This function must be called (rather than used in #OnAutoItStartRegister), as it uses a global array.
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _CmdLine_Parse($sPrefix = "--", $asAllowed = 0, $sOnErrFunc = "")
    If IsString($asAllowed) Then $asAllowed = StringSplit($asAllowed, "|", 3)

    For $i = 1 To $CmdLine[0]
        If $CmdLine[$i] = "--" Then
            If $i <> $CmdLine[0] Then
                $CmdLine_Cache[0][1] = $CmdLine[$i + 1]
                $i += 1
            Else
                $CmdLine_Cache[0][1] = True
            EndIf
        ElseIf StringLeft($CmdLine[$i], StringLen($sPrefix)) = $sPrefix Then
            $CmdLine_Cache[0][0] = UBound($CmdLine_Cache)
            ReDim $CmdLine_Cache[$CmdLine_Cache[0][0] + 1][2]

            If StringInStr($CmdLine[$i], "=") Then
                $CmdLine_Cache[$CmdLine_Cache[0][0]][0] = StringLeft($CmdLine[$i], StringInStr($CmdLine[$i], "=", 2) - 1)
                $CmdLine_Cache[$CmdLine_Cache[0][0]][0] = StringTrimLeft($CmdLine_Cache[$CmdLine_Cache[0][0]][0], StringLen($sPrefix))
                $CmdLine_Cache[$CmdLine_Cache[0][0]][1] = StringTrimLeft($CmdLine[$i], StringInStr($CmdLine[$i], "=", 2))
            ElseIf StringInStr($CmdLine[$i], ":") Then
                $CmdLine_Cache[$CmdLine_Cache[0][0]][0] = StringLeft($CmdLine[$i], StringInStr($CmdLine[$i], ":", 2) - 1)
                $CmdLine_Cache[$CmdLine_Cache[0][0]][0] = StringTrimLeft($CmdLine_Cache[$CmdLine_Cache[0][0]][0], StringLen($sPrefix))
                $CmdLine_Cache[$CmdLine_Cache[0][0]][1] = StringTrimLeft($CmdLine[$i], StringInStr($CmdLine[$i], ":", 2))
            Else
                $CmdLine_Cache[$CmdLine_Cache[0][0]][0] = StringTrimLeft($CmdLine[$i], StringLen($sPrefix))
                If ($i <> $CmdLine[0]) And (StringLeft($CmdLine[$i + 1], StringLen($sPrefix)) <> $sPrefix) Then
                    $CmdLine_Cache[$CmdLine_Cache[0][0]][1] = $CmdLine[$i + 1]
                    $i += 1
                Else
                    $CmdLine_Cache[$CmdLine_Cache[0][0]][1] = True
                EndIf
            EndIf

            If $asAllowed <> 0 Then
                For $n = 0 To UBound($asAllowed) - 1
                    If $CmdLine_Cache[$CmdLine_Cache[0][0]][0] = $asAllowed[$n] Then 
						ContinueLoop 2
					EndIf
                Next

                If $sOnErrFunc <> "" Then 
					Call($sOnErrFunc, "Unrecognized parameter: " & $CmdLine_Cache[$CmdLine_Cache[0][0]][0])
				Endif	
            EndIf
        Else
            If Not $CmdLine_Cache[0][1] Then
                $CmdLine_Cache[0][1] = $CmdLine[$i]
            Else
                If $sOnErrFunc <> "" Then Call($sOnErrFunc, "Unrecognized parameter: " & $CmdLine[$i])
            EndIf
        EndIf
    Next
EndFunc   ;==>_CmdLine_Parse