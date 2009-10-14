Attribute VB_Name = "mHelpers"
Option Explicit

 ' Daisy 2.02 Validator, Daisy 2.02 Regenerator, Bruno
 ' The Daisy Visual Basic Tool Suite
 ' Copyright (C) 2003,2004,2005,2006,2007,2008 Daisy Consortium
 '
 ' This library is free software; you can redistribute it and/or
 ' modify it under the terms of the GNU Lesser General Public
 ' License as published by the Free Software Foundation; either
 ' version 2.1 of the License, or (at your option) any later version.
 '
 ' This library is distributed in the hope that it will be useful,
 ' but WITHOUT ANY WARRANTY; without even the implied warranty of
 ' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 ' Lesser General Public License for more details.
 '
 ' You should have received a copy of the GNU Lesser General Public
 ' License along with this library; if not, write to the Free Software
 ' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

' SMIL-CLOCK constants
Public Const SCV_FullClock = 0
Public Const SCV_PartialClock = 1
Public Const SCV_Npt = 2
Public Const SCV_TimeCount_h = 3
Public Const SCV_TimeCount_m = 4
Public Const SCV_TimeCount_s = 5
Public Const SCV_TimeCount_ms = 6

' Rules for the fncConvertSmilClockVal2S and fncConvertMS2SmilClockVal functions
Private Const sDigit = "[0-9]"
Private Const s2Digits = "(" & sDigit & ", " & sDigit & ")"
Private Const sTimeDigits = "((([0-5])?, [0-9]) | ([0-9]))"
Private Const sTimeCount = "(" & sDigit & ")+"
Private Const sFraction = "(" & sDigit & ")+"
Private Const sSeconds = sTimeDigits
Private Const sMinutes = sTimeDigits
Private Const sHours = "(" & sDigit & ", (" & sDigit & ", (" & sDigit & ")?)?)"
Private Const sTimecount_val = "(" & sTimeCount & ", ('.', (" & sFraction & _
  "))? ('h' | 'min' | 's' | 'ms')?)"
Private Const sPartial_clock_val = "(" & sMinutes & ", ':', " & sSeconds & ", ('.', " & _
  sFraction & ")?)"
Private Const sFull_clock_val = "(" & sHours & ", ':', " & sMinutes & ", ':', " & sSeconds & _
  ", ('.', " & sFraction & ")?)"
Private Const sNPT = "((('n' | 'N'), ('p' | 'P'), ('t' | 'T')), '=', " & sTimeCount & _
  ", ('.', " & sTimeCount & ")?,'s')"

' Allowed time-fluctuation while comparing time-values
Public lTimeSpan As Long

' This function compares two values given in milliseconds regarding the allowed
' time-fluctuation
Public Function fncTimeCompareLog(ilValue1 As Long, ilValue2 As Long, _
  ibolHigher As Boolean) As Boolean
  
  If ibolHigher Then
    If (ilValue1 >= ilValue2 - lTimeSpan) Then fncTimeCompareLog = True
  Else
    If (ilValue1 <= ilValue2 + lTimeSpan) Then fncTimeCompareLog = True
  End If
End Function

' This function is a less advanced version of fncTimeCompareLog
Public Function fncTimeCompare(ilValue1 As Long, ilValue2 As Long) As Boolean
  If (ilValue1 >= ilValue2 - lTimeSpan) And _
     (ilValue1 <= ilValue2 + lTimeSpan) Then fncTimeCompare = True
End Function

' This function compares two smil-clock values regarding the time-fluctuation and
' the number of decimals found in the two values. The function compares down to
' the value with the least number of, or no decimals.
'
Public Function fncSmilClockIsEqual( _
  sSmilClock1 As String, sSmilClock2 As String _
  ) As Boolean
  
  Dim lDecCount1 As Long, lDecCount2 As Long
  Dim lValue1 As Long, lValue2 As Long
  
  sSmilClock1 = Replace(sSmilClock1, ",", ".", , , vbBinaryCompare)
  sSmilClock2 = Replace(sSmilClock2, ",", ".", , , vbBinaryCompare)
  
  lValue1 = fncConvertSmilClockVal2Ms(sSmilClock1)
  lValue2 = fncConvertSmilClockVal2Ms(sSmilClock2)
  
  If (fncRound2((lValue1 - lTimeSpan) / 1000) > fncRound2(lValue2 / 1000)) Or _
     (fncRound2((lValue1 + lTimeSpan) / 1000) < fncRound2(lValue2 / 1000)) Then Exit Function
  
  lDecCount1 = InStr(1, sSmilClock1, ".", vbBinaryCompare)
  lDecCount2 = InStr(2, sSmilClock2, ".", vbBinaryCompare)
  
  If lDecCount1 = 0 Or lDecCount2 = 0 Then fncSmilClockIsEqual = True: _
    Exit Function
  
  Dim lMultiplier As Long, lDecimals1 As Long, lDecimals2 As Long
  lMultiplier = 100: lDecimals1 = 0: lDecimals2 = 0
  
  lDecCount1 = lDecCount1 + 1
  lDecCount2 = lDecCount2 + 1
  Do Until (lDecCount1 = Len(sSmilClock1) Or lDecCount2 = Len(sSmilClock2))
    lDecimals1 = lDecimals1 + _
      (CLng(Mid$(sSmilClock1, lDecCount1, 1) * lMultiplier))
    lDecimals2 = lDecimals2 + _
      (CLng(Mid$(sSmilClock2, lDecCount2, 1) * lMultiplier))
    
    lDecCount1 = lDecCount1 + 1
    lDecCount2 = lDecCount2 + 1
  Loop
  
  If (lDecimals1 >= lDecimals2 - lTimeSpan) And _
     (lDecimals1 <= lDecimals2 + lTimeSpan) Then fncSmilClockIsEqual = True
End Function

' This function substitutes the Visual Basic round function, since it's not
' behaving the way we want (see below). This is a standard rounding routine that
' works like 1.49999999999999999999999999999 = 1 : 1.5 = 2
Public Function fncRound2(ByVal ivInput As Variant)
  Dim isTemp As String, lTemp As Long, lOut As Long
  isTemp = CStr(ivInput)
  lTemp = InStr(1, ivInput, ",", vbBinaryCompare)
  If lTemp < 1 Then fncRound2 = CLng(ivInput): Exit Function
  lOut = Left$(isTemp, lTemp - 1)
  isTemp = "0" & Right$(isTemp, Len(isTemp) - lTemp + 1)
  ivInput = isTemp
  If ivInput >= 0.5 Then lOut = lOut + 1
  fncRound2 = lOut
End Function

' This function converts string decimal time-values to milliseconds
Private Function fncConvertDecimals(isNumber As String) As Long
  Dim lMul As Long, lOut As Long, lCounter As Long
  
  lCounter = InStr(1, isNumber, ",", vbBinaryCompare)
  If lCounter = 0 Then lCounter = InStr(1, isNumber, ".", vbBinaryCompare)
  lCounter = lCounter + 1
  lMul = 100
  
  Do Until (lCounter > Len(isNumber)) Or (lMul = 0)
    lOut = lOut + (CLng(Mid$(isNumber, lCounter, 1)) * lMul)
    lMul = lMul \ 10
    lCounter = lCounter + 1
  Loop
  
  fncConvertDecimals = lOut
End Function

' This function converts a smil-clock or npt (seconds only) value to milliseconds
' using the rules given in the SMIL 1.0 DTD.
Public Function fncConvertSmilClockVal2Ms(ByVal sClockVal As String) As Long
  Dim objRc As New oDTDRuleChecker
  Dim sinTime As Long, lDecimals As Long
  
  objRc.sData = sClockVal
  objRc.lBytePos = 1
  objRc.lDataLength = Len(sClockVal)
  
  If objRc.conformsTo(sFull_clock_val) Then
    sinTime = CLng(fncExtractFromRule(sHours, "", sClockVal)) * 3600 * 1000
    sinTime = sinTime + CLng(fncExtractFromRule(sMinutes, "':'", sClockVal)) * 60 * 1000
    sinTime = sinTime + CLng(fncExtractFromRule(sSeconds, "':'", sClockVal)) * 1000
    sinTime = sinTime + fncConvertDecimals("0," & fncExtractFromRule(sFraction, _
      "'.'", sClockVal))
  ElseIf objRc.conformsTo(sPartial_clock_val) Then
    sinTime = sinTime + CLng(fncExtractFromRule(sMinutes, "':'", sClockVal)) * 60 * 1000
    sinTime = sinTime + CLng(fncExtractFromRule(sSeconds, "':'", sClockVal))
    sinTime = sinTime + fncConvertDecimals("0," & fncExtractFromRule(sFraction, _
      "'.'", sClockVal))
  ElseIf objRc.conformsTo(sTimecount_val) Then
    sinTime = sinTime + CLng(fncExtractFromRule(sTimeCount, "", sClockVal))
    lDecimals = fncConvertDecimals("0," & fncExtractFromRule(sFraction, "'.'", _
      sClockVal))
    
    Select Case LCase$(sClockVal)
      Case "h", "min", "s"
        sinTime = sinTime * 1000 + lDecimals
    End Select
    
    Select Case LCase$(sClockVal)
      Case "h"
        sinTime = sinTime * 3600 * 1000
      Case "min"
        sinTime = sinTime * 60 * 1000
    End Select
  ElseIf objRc.conformsTo(sNPT) Then
    sinTime = sinTime + (fncExtractFromRule(sTimeCount, _
      "((('n' | 'N'), ('p' | 'P'), ('t' | 'T')), '=')", sClockVal) * 1000)
    sinTime = sinTime + fncConvertDecimals("0," & fncExtractFromRule(sFraction, _
      "'.'", sClockVal))
  End If
  
  fncConvertSmilClockVal2Ms = sinTime
End Function

' This function converts milliseconds into a smil-clock or npt (seconds only)
' value using the SMIL 1.0 DTD rules.
Public Function fncConvertMS2SmilClockVal(ByVal sinTime As Long, _
  ByVal lSmilClockVal As Long, Optional ByVal ibolFraction As Variant) As String
  
  Dim sOutput As String
  
  Dim lH As Long, lM As Long, lS As Long, lMS As Long
  Dim lNewTime As Long, lTime As Long, bolFraction
  
  bolFraction = True
  If Not IsMissing(ibolFraction) Then bolFraction = ibolFraction
  
  lTime = sinTime
  lNewTime = lTime Mod 3600000
  lH = ((lTime - lNewTime) \ 3600000)
  lTime = lNewTime
  lNewTime = lTime Mod 60000
  lM = ((lTime - lNewTime) \ 60000)
  lTime = lNewTime
  lNewTime = lTime Mod 1000
  lS = ((lTime - lNewTime) \ 1000)
  lMS = lNewTime
  
  Select Case lSmilClockVal
    Case SCV_FullClock
      If Not (bolFraction) And (CLng(Left$(fnc4d(lMS), 1)) >= 5) Then lS = lS + 1
      If lS >= 60 Then lM = lM + 1: lS = lS - 60
      If lM >= 60 Then lH = lH + 1: lM = lM - 60
      sOutput = fnc2d(lH) & ":" & fnc2d(lM) & ":" & fnc2d(lS)
      If bolFraction Then sOutput = sOutput & "." & fnc4d(lMS)
    
    Case SCV_PartialClock
      If Not (bolFraction) And (CLng(Left$(fnc4d(lMS), 1)) >= 5) Then lS = lS + 1
      If lS >= 60 Then lM = lM + 1: lS = lS - 60
      If lM >= 60 Then lH = lH + 1: lM = lM - 60
      sOutput = fnc2d(lH * 60 + lM) & ":" & fnc2d(lS)
      If bolFraction Then sOutput = sOutput & "." & fnc4d(lMS)

    Case SCV_Npt
      If Not (bolFraction) And (CLng(Left$(fnc4d(lMS), 1)) >= 5) Then lS = lS + 1
      If lS >= 60 Then lM = lM + 1: lS = lS - 60
      If lM >= 60 Then lH = lH + 1: lM = lM - 60
      sOutput = "npt=" & fnc2d(lH * 3600 + lM * 60 + lS) '& "." & lMS
      If bolFraction Then sOutput = sOutput & "." & fnc4d(lMS)

    Case SCV_TimeCount_h
      If bolFraction Then
        sOutput = fnc3d(Round(sinTime / 3600000, 3)) & "h"
      Else
        sOutput = fnc3d(fncRound2(sinTime / 3600000)) & "h"
      End If
    
    Case SCV_TimeCount_m
      If bolFraction Then
        sOutput = fnc3d(Round(sinTime / 60000, 3)) & "m"
      Else
        sOutput = fnc3d(fncRound2(sinTime / 60000)) & "m"
      End If

    Case SCV_TimeCount_s
      If bolFraction Then
        sOutput = fnc3d(Round(sinTime / 1000, 3)) & "s"
      Else
        sOutput = fnc3d(fncRound2(sinTime / 1000)) & "s"
      End If

    Case SCV_TimeCount_ms
      sOutput = sinTime & "ms"
      If Not bolFraction Then sOutput = fncRound2(sOutput)
  End Select
  
  fncConvertMS2SmilClockVal = sOutput
End Function

' This function sets all output to double character
Private Function fnc2d(lInput As Long) As String
  Dim sOutput As String
  If lInput < 10 Then sOutput = "0"
  sOutput = sOutput & lInput
  fnc2d = sOutput
End Function

' This function replaces decimal ',' with '.'
Private Function fnc3d(dInput As Double) As String
  Dim sInput As String
  sInput = CStr(dInput)
  sInput = Replace(sInput, ",", ".", , , vbBinaryCompare)
  fnc3d = sInput
End Function

' This function adds '0' to a string to get the decimal format '000'
Private Function fnc4d(lInput As Long) As String
  Dim sInput As String
  sInput = CStr(lInput)
  Do While Len(sInput) < 3
    sInput = "0" & sInput
  Loop
  fnc4d = sInput
End Function

' This function extracts information from a string using DTD style rules.
' I.E
' a typical smil-clock value is '10:55:23.34', to extract the minutes ('55')
' into another string use:
' sResult = fncExtractFromRule("([0-9]+)", "([0-9]+, ':')", data)
Private Function fncExtractFromRule( _
  sRetrieve As String, sForwardPast As String, sFromString As String) As String
    
  Dim objRc As New oDTDRuleChecker, lStartPos As Long
  
  On Error Resume Next
  
  objRc.lBytePos = 1
  objRc.sData = sFromString
  objRc.lDataLength = Len(sFromString)
  
  objRc.conformsTo sForwardPast
  lStartPos = objRc.lBytePos
  
  objRc.conformsTo sRetrieve
  
  fncExtractFromRule = Mid$(sFromString, lStartPos, objRc.lBytePos - lStartPos)
  
  sFromString = Right$(sFromString, Len(sFromString) - objRc.lBytePos + 1)
End Function

' This function retrieves the filename from a string containing path + file + id
Public Function fncGetFileName(isAbsPath As String) As String
  Dim sTemp As String, sFileName As String
  If Not fncParseURI(isAbsPath, sTemp, sTemp, sFileName, sTemp) Then _
    objEvent.subLog "error in fncGetFileName"
  fncGetFileName = sFileName
End Function

' This function retrieves the path from a string containing path + file + id
Public Function fncGetPathName(isAbsPath As String) As String
  Dim sTemp As String, sDrive As String, sPath As String
  If Not fncParseURI(isAbsPath, sDrive, sPath, sTemp, sTemp) Then _
    objEvent.subLog "error in fncGetPathName"
  fncGetPathName = sDrive & sPath
End Function

' This function removes the id from a string containing path + file + id
Public Function fncStripId(isAbsPath As String) As String
  Dim sD As String, sP As String, sF As String, sTemp As String
  fncParseURI isAbsPath, sD, sP, sF, sTemp
  fncStripId = sD & sP & sF
End Function

' This function adds an absolute path to a relative path and removes the ID
Public Function fncStripIdAddPath( _
    ByVal isRelPath As String, ByVal isBasePath As String _
    ) As String
  
  Dim sD As String, sP As String, sF As String, sTemp As String
  fncParseURI isRelPath, sD, sP, sF, sTemp, isBasePath
  fncStripIdAddPath = sD & sP & sF
End Function

' This function retrieves the id from a string containing path + file + id
Public Function fncGetId( _
    ByVal isPath As String) As String
  Dim sTemp As String, sId As String
    
    fncParseURI isPath, sTemp, sTemp, sTemp, sId
    fncGetId = sId
End Function
Public Function fncFileExists( _
    sCandidate As String _
    ) As Boolean
Dim oFSO As Object: Set oFSO = CreateObject("Scripting.FileSystemObject")

    On Error GoTo ErrHandler
    fncFileExists = False
    If oFSO.FileExists(sCandidate) Then fncFileExists = True
ErrHandler:
    Set oFSO = Nothing
End Function

Public Function fncGetFileAsString(sAbsPath As String) As String
Dim oStream As Object, oFSO As Object
    On Error GoTo ErrHandler
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.opentextfile(sAbsPath)
    fncGetFileAsString = oStream.ReadAll
ErrHandler:
    oStream.Close
End Function

' This function retrives the extension of a file
Public Function fncGetExtension(ByVal sFileName As String) As String
    Dim sD As String, sP As String, sF As String, sId As String, lTemp As Long
    fncParseURI sFileName, sD, sP, sF, sId
    lTemp = InStrRev(sFileName, ".", -1, vbBinaryCompare)
    If lTemp < 1 Or lTemp > Len(sFileName) Then Exit Function
    fncGetExtension = Right$(sFileName, Len(sFileName) - lTemp)
End Function

' This function retrieves the BASE name of a file
Public Function fncGetBaseName(ByVal sFileName As String) As String
    Dim sD As String, sP As String, sF As String, sId As String, lTemp As Long
    fncParseURI sFileName, sD, sP, sF, sId
    lTemp = InStrRev(sFileName, ".", -1, vbBinaryCompare)
    If lTemp < 1 Or lTemp > Len(sFileName) Then Exit Function
    fncGetBaseName = Left$(sFileName, lTemp - 1)
End Function

' This is a multi-purpose path separator, this function takes any relative or
' absolute path and separates it into Drive, Path, Filename, ID. It can also add
' an absolute path to a relative path.
Public Function fncParseURI(ByVal isHref As String, isDrive As String, isPath As String, _
  isFileName As String, isID As String, Optional isDefDrive As Variant) As Boolean
  
  fncParseURI = False
    
  Dim templCounter As Long, sDefPath As String, sDefDrive As String, tempsString As String
  Dim sDefFile As String
  
  If Not IsMissing(isDefDrive) Then
    If Not fncParseURI(CStr(isDefDrive), sDefDrive, sDefPath, sDefFile, tempsString) Then _
      Exit Function
  End If
  
  If (Not Right$(isHref, 1) = "\") And (Not Right$(isHref, 1) = "/") Then
    templCounter = InStrRev(isHref, "#", -1, vbBinaryCompare)
    If templCounter <> 0 Then
      isID = Right$(isHref, Len(isHref) - templCounter)
      isHref = Left$(isHref, templCounter - 1)
    End If
  End If
  
'  Debug.Print "No lcase$ in fncParseUri -> Check if errors"
  
  For templCounter = 1 To Len(isHref)
    If Mid$(isHref, templCounter, 1) = "/" Then Mid$(isHref, templCounter, 1) = "\"
  Next templCounter
  
  If LCase$(Left$(isHref, 7)) = "file:\\" Then _
    isHref = Right$(isHref, Len(isHref) - 7)
  If LCase$(Left$(isHref, 7)) = "http:\\" Or Left$(isHref, 6) = "ftp:\\" Then _
    Exit Function
  If LCase$(Left$(isHref, 7)) = "mailto:" Then Exit Function
     
  templCounter = InStr(1, isHref, ":", vbBinaryCompare)
  If templCounter <> 0 Then
    isDrive = Mid$(isHref, templCounter - 1, 2)
    isHref = Mid$(isHref, templCounter + 1, Len(isHref) - templCounter)
  Else
    isDrive = sDefDrive
  End If
  
  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
  If templCounter <> 0 Then
    isPath = sDefPath & Left$(isHref, templCounter)
    isHref = Right$(isHref, Len(isHref) - templCounter)
  Else
    isPath = sDefPath
  End If
  
  templCounter = InStrRev(isHref, "\", -1, vbBinaryCompare)
  If templCounter <> 0 And templCounter < Len(isHref) Then
    isFileName = Right$(isHref, Len(isHref) - templCounter)
    isHref = Left$(isHref, templCounter - 1)
  ElseIf isHref = "" Then
    isFileName = sDefFile
  Else
    isFileName = isHref
  End If
  
  If Left$(isPath, 1) = "." Then
    If Left$(isPath, 2) = ".." Then templCounter = 2 Else templCounter = 1
    isPath = Right$(isPath, Len(isPath) - templCounter)
    
    If Not IsMissing(isDefDrive) Then
      Dim tempsD As String, tmpsP As String
      fncParseURI CStr(isDefDrive), tempsD, tmpsP, tempsString, tempsString
      If Not tempsD = isDrive Then GoTo SetDefault
      
      If Right$(tmpsP, 1) = "\" And Left$(isPath, 1) = "\" Then _
        isPath = Right$(isPath, Len(isPath) - 1)
      
      If templCounter = 1 Then
        isPath = tmpsP & isPath
      Else
        templCounter = InStrRev(tmpsP, "\", -1, vbBinaryCompare)
        If templCounter > 0 Then isPath = Left$(tmpsP, templCounter) & isPath
      End If
    Else
SetDefault:
    End If
  End If
  
  fncParseURI = True
End Function

' This function is performing a classic "shift-left" bit operation
' (damn vb, doesn't have these functions)
Public Function fncShl(ilValue As Variant, ilShift As Long) As Long
  fncShl = ilValue * (2 ^ ilShift)
End Function

' This function is performing a classic "shift-right" bit operation
' (damn vb, doesn't have these functions)
Public Function fncShr(ilValue As Variant, ilShift As Long) As Long
  fncShr = ilValue \ (2 ^ ilShift)
End Function

' This function converts a string value into an integer. This function is used
' rather than CLNG() or something similar to avoid the use of ON ERROR.
Public Function fncString2Integer(isInput As String) As Long
  Dim lMul As Long, lCounter As Long, lOut As Long, sChar As String
  
  lMul = 1
  If isInput = "" Then lOut = -1: Exit Function
  
  For lCounter = Len(isInput) To 1 Step -1
    sChar = Mid$(isInput, lCounter, 1)
  
    Select Case sChar
      Case "0" To "9"
        lOut = lOut + ((Asc(sChar) - 48) * lMul)
      Case Else
        lOut = -1
        Exit For
    End Select
    
    If Not lCounter = 1 Then lMul = lMul * 10
  Next lCounter
  
  fncString2Integer = lOut
End Function

Public Function fncGetAbsolutePathName(ByRef isPath As String) As String
Dim oFSO As Object
  On Error GoTo ErrH
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  fncGetAbsolutePathName = oFSO.GetAbsolutePathName(isPath)
  Exit Function
ErrH:
 isPath = ""
End Function
