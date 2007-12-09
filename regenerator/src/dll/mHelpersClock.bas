Attribute VB_Name = "mHelpersClock"
' Daisy 2.02 Regenerator DLL
' Copyright (C) 2003 Daisy Consortium
'
'    This file is part of Daisy 2.02 Regenerator.
'
'    Daisy 2.02 Regenerator is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Daisy 2.02 Regenerator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Daisy 2.02 Regenerator; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

Option Explicit

' SMIL-CLOCK constants
Public Const SCV_FullClock = 0
Public Const SCV_PartialClock = 1
Public Const SCV_Npt = 2
Public Const SCV_TimeCount_h = 3
Public Const SCV_TimeCount_m = 4
Public Const SCV_TimeCount_s = 5
Public Const SCV_TimeCount_ms = 6

' Rules for the fncConvertSmilClockVal2S and fncConvertMs2SmilClockVal functions
Private Const sDigit = "[0-9]"
Private Const s2Digits = "(" & sDigit & ", " & sDigit & ")"
Private Const sTimeDigits = "((([0-5])?, [0-9]) | ([0-9]))"
Private Const sTimeCount = "(" & sDigit & ")+"
Private Const sFraction = "(" & sDigit & ")+"
Private Const sSeconds = sTimeDigits
Private Const sMinutes = sTimeDigits
Private Const sHours = "(" & sDigit & ", (" & sDigit & ")?)"
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


' This function converts string decimal time-values to milliseconds
'
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
'
Public Function fncConvertSmilClockVal2Ms(ByVal sClockVal As String) As Long
 Dim objRc As oDTDRuleChecker
 Dim sinTime As Long, lDecimals As Long
 
 Set objRc = New oDTDRuleChecker
 
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
 
 Set objRc = Nothing
End Function

' This function converts milliseconds into a smil-clock or npt (seconds only)
' value using the SMIL 1.0 DTD rules.
'
Public Function fncConvertMS2SmilClockVal(ByVal sinTime As Long, _
  ByVal lSmilClockVal As Long, Optional ByVal ibolFraction As Variant) As String
  
  Dim sOutPut As String
  
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
      sOutPut = fnc2d(lH) & ":" & fnc2d(lM) & ":" & fnc2d(lS)
      If bolFraction Then sOutPut = sOutPut & "." & fnc4d(lMS)
    
    Case SCV_PartialClock
      If Not (bolFraction) And (CLng(Left$(fnc4d(lMS), 1)) >= 5) Then lS = lS + 1
      If lS >= 60 Then lM = lM + 1: lS = lS - 60
      If lM >= 60 Then lH = lH + 1: lM = lM - 60
      sOutPut = fnc2d(lH * 60 + lM) & ":" & fnc2d(lS)
      If bolFraction Then sOutPut = sOutPut & "." & fnc4d(lMS)

    Case SCV_Npt
      If Not (bolFraction) And (CLng(Left$(fnc4d(lMS), 1)) >= 5) Then lS = lS + 1
      If lS >= 60 Then lM = lM + 1: lS = lS - 60
      If lM >= 60 Then lH = lH + 1: lM = lM - 60
      sOutPut = "npt=" & fnc2d(lH * 3600 + lM * 60 + lS) '& "." & lMS
      If bolFraction Then sOutPut = sOutPut & "." & fnc4d(lMS)

    Case SCV_TimeCount_h
      If bolFraction Then
        sOutPut = fnc3d(Round(sinTime / 3600000, 3)) & "h"
      Else
        sOutPut = fnc3d(Round2(sinTime / 3600000)) & "h"
      End If
    
    Case SCV_TimeCount_m
      If bolFraction Then
        sOutPut = fnc3d(Round(sinTime / 60000, 3)) & "m"
      Else
        sOutPut = fnc3d(Round2(sinTime / 60000)) & "m"
      End If

    Case SCV_TimeCount_s
      If bolFraction Then
        sOutPut = fnc3d(Round(sinTime / 1000, 3)) & "s"
      Else
        sOutPut = fnc3d(Round2(sinTime / 1000)) & "s"
      End If

    Case SCV_TimeCount_ms
      sOutPut = sinTime & "ms"
      If Not bolFraction Then sOutPut = Round2(sOutPut)
'      sOutput = sinTime & "ms"
  End Select
  
  fncConvertMS2SmilClockVal = sOutPut
End Function

Public Function Round2(ByVal ivInput As Variant)
  Dim isTemp As String, lTemp As Long, lOut As Long
  isTemp = CStr(ivInput)
  lTemp = InStr(1, ivInput, ",", vbBinaryCompare)
  If lTemp < 1 Then Round2 = CLng(ivInput): Exit Function
  lOut = Left$(isTemp, lTemp - 1)
  isTemp = "0" & Right$(isTemp, Len(isTemp) - lTemp + 1)
  ivInput = isTemp
  If ivInput >= 0.5 Then lOut = lOut + 1
  Round2 = lOut
End Function

' This function sets all output to double character
Private Function fnc2d(lInput As Long) As String
  Dim sOutPut As String
  If lInput < 10 Then sOutPut = "0"
  sOutPut = sOutPut & lInput
  fnc2d = sOutPut
End Function

' This function replaces decimal ',' with '.'
Private Function fnc3d(dInput As Double) As String
  Dim sInput As String
  sInput = CStr(dInput)
  sInput = Replace(sInput, ",", ".", , , vbBinaryCompare)
  fnc3d = sInput
End Function

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
' sResult = fncExtractFromRule("[0-9]+", "([0-9]+, ':')", data)
'
Private Function fncExtractFromRule( _
 sRetrieve As String, sForwardPast As String, sFromString As String) As String
   
 Dim objRc As oDTDRuleChecker, lStartPos As Long
 
 Set objRc = New oDTDRuleChecker
 
 On Error Resume Next
 
 objRc.lBytePos = 1
 objRc.sData = sFromString
 objRc.lDataLength = Len(sFromString)
 
 objRc.conformsTo sForwardPast
 lStartPos = objRc.lBytePos
 
 objRc.conformsTo sRetrieve
 
 fncExtractFromRule = Mid$(sFromString, lStartPos, objRc.lBytePos - lStartPos)
 
 sFromString = Right$(sFromString, Len(sFromString) - objRc.lBytePos + 1)
 
 Set objRc = Nothing
End Function

