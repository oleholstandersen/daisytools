VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oDTDRuleChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

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

'Option Explicit

Public lBytePos As Long, sData As String, lDataLength As Long
Private lTop As Long
Private sErrorChain() As String, lErrorsInChain As Long

Private Const dq = """"
Private Const sq = "'"

Private sLastUnparsed As String, sLastParsed As String

Private sHierarchy As String

Public Function conformsTo(Optional isString As Variant, Optional iclsRules As _
 Variant) As Boolean

 Dim bolResult As Boolean
 
 If IsMissing(isString) And IsMissing(iclsRules) Then
   bolResult = conformsTo2
 ElseIf IsMissing(isString) Then
   bolResult = conformsTo2(, iclsRules)
 ElseIf IsMissing(iclsRules) Then
   bolResult = conformsTo2(isString)
 Else
   bolResult = conformsTo2(isString, iclsRules)
 End If
 
 If (Not lBytePos >= lDataLength) And (bolResult) Then bolResult = False
 conformsTo = bolResult
End Function

Public Function conformsTo2(Optional isString As Variant, Optional iclsRules As _
 Variant) As Boolean
 
 Dim clsRules As cDTDRules, templCounter As Long, lDataCount As Long
 Dim bolResult As Boolean, lOccurance As Long, lBackup As Long
 Dim templValue As Long
 
 On Error GoTo ErrorH

FindAgain:
 
 lBackup = lBytePos
 
 If IsMissing(iclsRules) Then
   If IsMissing(isString) Then Exit Function
   Set clsRules = New cDTDRules
   clsRules.fncParseRule CStr(isString), 1
 Else
   Set clsRules = iclsRules
 End If
 
 sHierarchy = sHierarchy & clsRules.sName
 
 If clsRules.sName = "" Then
   templCounter = 1
   templValue = 0
   Do Until templCounter > clsRules.lRulesCount
     bolResult = conformsTo2(, clsRules.pRule(templCounter))
     If bolResult Then templValue = templValue + 1
     templCounter = templCounter + 1
     
     If clsRules.enChildOccurance = 1 And templValue = 1 Then Exit Do
     If (Not bolResult) And clsRules.enChildOccurance = 3 Then Exit Do
   Loop
   
   Select Case clsRules.enChildOccurance
     Case 1, 0 'OneMustOccur
       If templValue = 1 Then lOccurance = lOccurance + 1: bolResult = True Else _
         bolResult = False
       
     Case 3 'AllMustOccur
       If (templValue = clsRules.lRulesCount) Then lOccurance = lOccurance + 1: _
         bolResult = True Else lOccurance = 0: bolResult = False
         
       'Debug.Print "loccurance should be -1 if if-statement fails ??????"
   End Select
 
 ElseIf Left$(clsRules.sName, 1) = dq Or Left$(clsRules.sName, 1) = sq Then
   If clsRules.enMainOccurance = 6 Then
     lBytePos = lBytePos - (Len(clsRules.sName) - 2): lBackup = lBytePos
     bolResult = hasString(Mid$(clsRules.sName, 2, Len(clsRules.sName) - 2), True)
     If Not bolResult Then
       lBytePos = lBytePos + (Len(clsRules.sName) - 2)
     End If
   Else
     bolResult = hasString(Mid$(clsRules.sName, 2, Len(clsRules.sName) - 2), True)
   End If
   If bolResult Then lOccurance = 1
 ElseIf Left$(clsRules.sName, 1) = "[" Then
   bolResult = hasCharacters(Mid$(clsRules.sName, 2, Len(clsRules.sName) - 2))
   If clsRules.enChildOccurance = 4 And bolResult Then
     templCounter = lBytePos
     lBytePos = lBackup
     If conformsTo(, clsRules.pRule(1)) Then bolResult = False
     lBytePos = templCounter
   End If
   
   If bolResult Then lOccurance = 1
 Else
   bolResult = CallByName(Me, clsRules.sName, VbMethod)
   If clsRules.enChildOccurance = 4 And bolResult Then
     templCounter = lBytePos
     lBytePos = lBackup
     If conformsTo(, clsRules.pRule(1)) Then bolResult = False
     lBytePos = templCounter
   End If
   
   If bolResult Then lOccurance = 1
 End If
 
 Select Case clsRules.enMainOccurance
   Case 1 'OneOrMore
     If lOccurance > 0 Then
       conformsTo2 = True
'lower should maybe only be <
       If lBytePos <= lDataLength And bolResult Then GoTo FindAgain
     Else
       conformsTo2 = False
     End If
   
   Case 3 'ZeroOrOne
     If lOccurance = 0 Or lOccurance = 1 Then conformsTo2 = True Else _
       conformsTo2 = False
     
   Case 4 'ZeroOrMore
     If lOccurance > -1 Then
       conformsTo2 = True
'lower should maybe only be <
       If lBytePos <= lDataLength And bolResult Then GoTo FindAgain
     Else
       conformsTo2 = False
     End If
   
   Case 5 'ExactlyOne
     If lOccurance = 1 Then conformsTo2 = True Else conformsTo2 = False
     
   Case 6
     If lOccurance > 0 Then conformsTo2 = False Else conformsTo2 = True
 End Select
 
 If lTop < lBytePos Then lTop = lBytePos
 If conformsTo2 = False Then lBytePos = lBackup
 
 If conformsTo2 = True Then
   lErrorsInChain = 0
 End If
 
 sHierarchy = Left$(sHierarchy, Len(sHierarchy) - Len(clsRules.sName))
 
 Exit Function
ErrorH:
 conformsTo2 = False
 AddError "conformsTo"
End Function

Public Function hasString(isString As String, ibolCaseSensitive As Boolean) _
 As Boolean
'
'  If isString = "windows-1252" Or Left$(isString, 3) = "utf" Then MsgBox ("TJO")
 
 If Len(isString) > lBytePos + lDataLength Then Exit Function
 If ibolCaseSensitive Then
   If Not Mid$(sData, lBytePos, Len(isString)) = isString Then GoTo ErrorH
 Else
   If Not LCase$(Mid$(sData, lBytePos, Len(isString))) = LCase$(isString) Then _
     GoTo ErrorH:
 End If
 
 lBytePos = lBytePos + Len(isString)
 
 hasString = True
 
 Exit Function
ErrorH:
 AddError "hasString"
End Function

Public Function hasCharacters(ByVal isString As String) As Boolean
 Dim templCounter As Long, bolFound As Boolean
 Dim templCounter2 As Long, templCounter3 As Long
 Dim bolOutside As Boolean, tempsString As String
   
 Dim aLowerRange() As Long, aHigherRange() As Long, lRangeCount As Long
 Dim currChar As Long, lValue1 As Long, lValue2 As Long, bolV2 As Boolean
    
 If (lBytePos > Len(sData)) Then Exit Function

 If Left$(isString, 1) = "^" Then _
   bolOutside = True: isString = Right$(isString, Len(isString) - 1)
     
 templCounter = 1
 
 Do
   currChar = Asc(Mid$(isString, templCounter, 1))
 
   If Mid$(isString, templCounter, 2) = "#x" Then
     templCounter3 = InStr(templCounter, isString, ";", vbBinaryCompare)
     If Not (templCounter3 = 0) Then
       currChar = CLng("&h" & Mid$(isString, templCounter + 2, _
         templCounter3 - templCounter - 2))
       templCounter = templCounter3
     End If
   End If
   'Else
   If Not bolV2 Then lValue1 = currChar Else lValue2 = currChar
   'End If
   
   templCounter = templCounter + 1
   If Mid$(isString, templCounter, 1) = "-" Then
     bolV2 = True
     templCounter = templCounter + 1
   Else
     If Not bolV2 Then lValue2 = lValue1 Else bolV2 = False
     
     lRangeCount = lRangeCount + 1
     ReDim Preserve aLowerRange(lRangeCount)
     ReDim Preserve aHigherRange(lRangeCount)
     
     aLowerRange(lRangeCount) = lValue1
     aHigherRange(lRangeCount) = lValue2
   End If
   
 Loop Until templCounter > Len(isString)
 
 currChar = AscW(Mid$(sData, lBytePos, 1))
 bolFound = False
 For templCounter = 1 To lRangeCount
   If currChar >= aLowerRange(templCounter) And _
     currChar <= aHigherRange(templCounter) Then bolFound = True: Exit For
 Next templCounter

 If bolOutside Then bolFound = Not bolFound
   
 lBytePos = lBytePos + 1
 
 hasCharacters = bolFound
End Function

Public Sub AddError(isError As String)
 lErrorsInChain = lErrorsInChain + 1
 ReDim Preserve sErrorChain(lErrorsInChain)
 sErrorChain(lErrorsInChain) = isError
 
 Debug.Print isError & ":" & sHierarchy
End Sub


