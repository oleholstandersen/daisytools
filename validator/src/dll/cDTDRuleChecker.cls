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

' sData = string data that shall be tested
' lBytePos = character pointer of where to begin testing in sData
' lDataLength = the total length of sData
Public lBytePos As Long, sData As String, lDataLength As Long

' lTop = highest character pointer reached - for debugging
Private lTop As Long

' Error hierarchy
Private sErrorChain() As String, lErrorsInChain As Long

Private Const dq = """"
Private Const sq = "'"

' sHierarchy = the current rules hierarchy
Private sHierarchy As String

' This function runs the conformsTo2 function and checks returns 'true' if ALL
' the data supplied followed the rules given
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

' This function compares the data given in sData, starting at lBytePos with the
' rules given in iclsRules or isString. isString must - if present - contain
' a string with parse rules; "('time:', ([0-9])+)". This information will be
' parsed into a cDTDRules class. iclsRules must - if present - contain a cDTDRules
' class with parsed DTD rules. One of the two arguments must be present.
'
Public Function conformsTo2(Optional isString As Variant, Optional iclsRules As _
  Variant) As Boolean

'                 cDTDRules
  Dim clsRules As Object, templCounter As Long, lDataCount As Long
  Dim bolResult As Boolean, lOccurance As Long, lBackup As Long
  Dim templValue As Long
  
  On Error GoTo ErrorH

FindAgain:
  DoEvents

' save the current character pointer
  lBackup = lBytePos
    
' if iclsRules is missing and isString is present, parse isString into a cDTDRules
' class. If iclsRules is present, set clsRules to point to iclsRules.
  If IsMissing(iclsRules) Then
    If IsMissing(isString) Then Exit Function
    'Set clsRules = New cDTDRules
    Set clsRules = CreateObject("DTDParser.cDTDRules")
    clsRules.fncParseRule CStr(isString), 1
  Else
    Set clsRules = iclsRules
  End If
  
' error info
  sHierarchy = sHierarchy & clsRules.sName

' if the current rules name is null, the structure must contain a set of
' childrules. Example; the rule 'first' has the name 'first' while the rule
' ('first' | 'last') has no name, though it has two childrules; 'first' and 'last'
  If clsRules.sName = "" Then
    templCounter = 1
    templValue = 0

' loop trough all childrules
    Do Until templCounter > clsRules.lRulesCount
      bolResult = conformsTo2(, clsRules.pRule(templCounter))
' if the data followed the rule, increase the number of rules passed
      If bolResult Then templValue = templValue + 1
      templCounter = templCounter + 1
            
' if the data passed the current childrule and the current rule say that one of the
' childrules must pass (the OR '|' statement), exit do.
      If clsRules.enChildOccurance = 1 And templValue = 1 Then Exit Do
' if the data didn't pass the current childrule and the current rule say that all
' childrules must be passed (the AND ',' statement), exit do.
      If (Not bolResult) And clsRules.enChildOccurance = 3 Then Exit Do
    Loop
    
    Select Case clsRules.enChildOccurance
      Case 1, 0 'OneMustOccur
' if the current rule says that one childrule must pass, and the number of
' passed childrules is 1, this rules returnvalue = true and the number of
' succeeded CURRENT rules are increased
        If templValue = 1 Then lOccurance = lOccurance + 1: bolResult = True Else _
          bolResult = False
        
      Case 3 'AllMustOccur
' if the current rule says that all childrules must pass, and the number of
' passed childrules is equal to the number of childrules, this rules
' returnvalue = true and the number of succeeded CURRENT rules are increased
        If (templValue = clsRules.lRulesCount) Then lOccurance = lOccurance + 1: _
          bolResult = True Else lOccurance = 0: bolResult = False
          
'        Debug.Print "loccurance should be -1 if if-statement fails ??????"
    End Select

' if the current rules name starts with a ' or a ", this rule is a string
' expression and will be tested with string functions
  ElseIf Left$(clsRules.sName, 1) = dq Or Left$(clsRules.sName, 1) = sq Then
' if this rule is of negative type - "this must NOT occur" - go down to the
' function which doesn't increse the current character pointer in the data
    If clsRules.enMainOccurance = 6 Then
      lBytePos = lBytePos - (Len(clsRules.sName) - 2): lBackup = lBytePos
      bolResult = hasString(Mid$(clsRules.sName, 2, Len(clsRules.sName) - 2), True)
      If Not bolResult Then
        lBytePos = lBytePos + (Len(clsRules.sName) - 2)
      End If
    Else
      bolResult = hasString(Mid$(clsRules.sName, 2, Len(clsRules.sName) - 2), True)
    End If
' if the current rule passed, set occurance to 1
    If bolResult Then lOccurance = 1
' if the current rules anme starts with a '[' this is a character span and will
' be tested with character span functions
  ElseIf Left$(clsRules.sName, 1) = "[" Then
    bolResult = hasCharacters(Mid$(clsRules.sName, 2, Len(clsRules.sName) - 2))
' if the current rule may only occur one or zero times and has allready passed
' once, test it again. If it is true again, the result is false.
    If clsRules.enChildOccurance = 4 And bolResult Then
      templCounter = lBytePos
      lBytePos = lBackup
      If conformsTo(, clsRules.pRule(1)) Then bolResult = False
      lBytePos = templCounter
    End If
    
    If bolResult Then lOccurance = 1
  Else
' **** Not used in the validator
' if the name doesn't follow standard rules, this must be a function. Made for
' dom parsing / creation.
    bolResult = CallByName(Me, clsRules.sName, VbMethod)
' if the current rule may only occur one or zero times and has allready passed
' once, test it again. If it is true again, the result is false.
    If clsRules.enChildOccurance = 4 And bolResult Then
      templCounter = lBytePos
      lBytePos = lBackup
      If conformsTo(, clsRules.pRule(1)) Then bolResult = False
      lBytePos = templCounter
    End If
    
    If bolResult Then lOccurance = 1
' ****
  End If
  
  Select Case clsRules.enMainOccurance
    Case 1 'OneOrMore
' if the current rule must occur one or several times and has done that the
' function = true. If there's more data to try the rule on and the last data vs
' rule test was successfull, try again.
      If lOccurance > 0 Then
        conformsTo2 = True
        If lBytePos <= lDataLength And bolResult Then GoTo FindAgain
      Else
        conformsTo2 = False
      End If
    
    Case 3 'ZeroOrOne
' if the current rule must occur one or zero times and has done that the function
' = true
      If lOccurance = 0 Or lOccurance = 1 Then conformsTo2 = True Else _
        conformsTo2 = False
      
    Case 4 'ZeroOrMore
' if the current rule must occur zero or several times and has done that the
' function = true. If there's more data to try the rule on and the last data vs
' rule test was successfull, try again.
      If lOccurance > -1 Then
        conformsTo2 = True
'lower should maybe only be <
        If lBytePos <= lDataLength And bolResult Then GoTo FindAgain
      Else
        conformsTo2 = False
      End If
    
    Case 5 'ExactlyOne
' if the current rule must occur one time and has done that the function = true
      If lOccurance = 1 Then conformsTo2 = True Else conformsTo2 = False
      
    Case 6
' if the current rule must not occur and hasn't done that the function = true
      If lOccurance > 0 Then conformsTo2 = False Else conformsTo2 = True
  End Select
  
' debugging information
  If lTop < lBytePos Then lTop = lBytePos

' if the current rule didn't succed, reset the character pointer to what it was
' from the beginning
  If conformsTo2 = False Then lBytePos = lBackup
  
  If conformsTo2 = True Then
    lErrorsInChain = 0
  End If
  
  sHierarchy = Left$(sHierarchy, Len(sHierarchy) - Len(clsRules.sName))
  
  Set clsRules = Nothing
  Exit Function
ErrorH:
  Set clsRules = Nothing

  conformsTo2 = False
  AddError "conformsTo"
End Function

' this function checks wheter the current data pointed at in sData
' is the same as isString
Public Function hasString(isString As String, ibolCaseSensitive As Boolean) _
  As Boolean
  
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

' this function checks wheter the current character pointed at in sData is defined
' within the character rules given in isString
Public Function hasCharacters(ByVal isString As String) As Boolean
  Dim templCounter As Long, bolFound As Boolean
  Dim templCounter2 As Long, templCounter3 As Long
  Dim bolOutside As Boolean, tempsString As String
    
  Dim aLowerRange() As Long, aHigherRange() As Long, lRangeCount As Long
  Dim currChar As Long, lValue1 As Long, lValue2 As Long, bolV2 As Boolean
     
  If (lBytePos > Len(sData)) Then Exit Function

' the ^ is NOT, this means that the current character pointed at in sData must
' NOT be within the given character rules
  If Left$(isString, 1) = "^" Then _
    bolOutside = True: isString = Right$(isString, Len(isString) - 1)
      
  templCounter = 1
  
  Do
' get the current characters ASCII value
    currChar = Asc(Mid$(isString, templCounter, 1))
' if the current character pointed at + 1 character = "#x" then this is a character
' defined in hexadecimals and we'll have to convert it first.
    If Mid$(isString, templCounter, 2) = "#x" Then
      templCounter3 = InStr(templCounter, isString, ";", vbBinaryCompare)
      If Not (templCounter3 = 0) Then
        currChar = CLng("&h" & Mid$(isString, templCounter + 2, _
          templCounter3 - templCounter - 2))
        templCounter = templCounter3
      End If
    End If
        
    If Not bolV2 Then lValue1 = currChar Else lValue2 = currChar
    
    
    templCounter = templCounter + 1
' if the current character pointed at in sData is a - this is a character span
    If Mid$(isString, templCounter, 1) = "-" Then
      bolV2 = True
      templCounter = templCounter + 1
    Else
      If Not bolV2 Then lValue2 = lValue1 Else bolV2 = False
      
      lRangeCount = lRangeCount + 1
' all characters are defined as spans.
      ReDim Preserve aLowerRange(lRangeCount)
      ReDim Preserve aHigherRange(lRangeCount)
      
      aLowerRange(lRangeCount) = lValue1
      aHigherRange(lRangeCount) = lValue2
    End If
    
  Loop Until templCounter > Len(isString)
  
  currChar = AscW(Mid$(sData, lBytePos, 1))
  If currChar < 0 Then currChar = currChar + 65536
  bolFound = False
' compare the current character pointed at in sData with the spans given
  For templCounter = 1 To lRangeCount
    If currChar >= aLowerRange(templCounter) And _
      currChar <= aHigherRange(templCounter) Then bolFound = True: Exit For
  Next templCounter

  If bolOutside Then bolFound = Not bolFound
    
  lBytePos = lBytePos + 1
  
  hasCharacters = bolFound
End Function

' add an error - for debugging
Public Sub AddError(isError As String)
  lErrorsInChain = lErrorsInChain + 1
  ReDim Preserve sErrorChain(lErrorsInChain)
  sErrorChain(lErrorsInChain) = isError
  
'  Debug.Print isError & ":" & sHierarchy
End Sub

