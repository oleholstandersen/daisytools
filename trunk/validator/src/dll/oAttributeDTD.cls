VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oAttributeDTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Daisy 2.02 Validator Engine
' Copyright (C) 2002 Daisy Consortium
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
'
' For information about the Daisy Consortium, visit www.daisy.org or contact
' info@mail.daisy.org. For development issues, contact markus.gylling@tpb.se or
' karl.ekdahl@tpb.se.

Const sErrfDACV As String = "Error in fncDocumentAttributeContentValid: "
Const sErrfIEACV As String = "Error in fncIntraElementAttributesCorrelationValid: "

' This function is checking all attributes found in the given DOM and compares
' the value of them with the ADTD belonging to the given filetype
Public Function fncDocumentAttributesContentValid(iobjReport As Object, _
  iobjDOMStructure As Object, enFileType As enuFileType, _
  isAbsPath As String) As Boolean
  
  fncDocumentAttributesContentValid = False
  
  Dim oNodeList As Object, oNode As Object, oElement As Object
  Dim clsRules As Object, sALEdtd As String
  
  'fncInsertTime "oAttributeDTD.fncDocumentAttributesContentValid"
  
  If sAdtdPath = "" Then _
    objEvent.subLog sErrfDACV & "no adtd specified": Exit Function
  If iobjDOMStructure Is Nothing Then _
    objEvent.subLog sErrfDACV & "dom structure is null": Exit Function
  
  If Not bolRulesAreLoaded Then If Not fncLoadRulesFiles Then Exit Function
  
  ' Set clsRules to point at the correct rule object
  If enFileType = smil Then
    Set clsRules = objRulesSmil
  ElseIf enFileType = ncc Then
    Set clsRules = objRulesNcc
  ElseIf enFileType = mastersmil Then
    Set clsRules = objRulesMasterSmil
  ElseIf enFileType = smilMediaObText Then
    Set clsRules = objRulesContent
  ElseIf enFileType = nccmultivolume Then
    Set clsRules = objRulesMultivolume
  ElseIf enFileType = discinfo Then
    Set clsRules = objRulesDiscinfo
  Else
    Exit Function
  End If
    
  ' Go trough each element found in the file and see if it follows the ADTDs rules
  fncDocumentAttributesContentValid = True
  
  Dim lCurrentItem As Long
  
  If iobjDOMStructure Is Nothing Then _
    objEvent.subLog sErrfDACV & "dom structure is null": Exit Function
  
  ' Select all elements of the document
  Set oNodeList = iobjDOMStructure.documentElement.getElementsByTagName("*")
  fncSetProgress Me, 0, oNodeList.length
  
  For Each oNode In oNodeList
    If (oNode.nodeType = NODE_ELEMENT) Then
      Set oElement = oNode
      ' Test the current element
      If Not fncCheckSingleAttribute(oElement, clsRules, iobjReport, isAbsPath) _
        Then
        
        objEvent.subLog sErrfDACV & " error in fncCheckSingleAttribute while " & _
          "parsing elementtype " & oElement.nodeName
        fncDocumentAttributesContentValid = False
      End If
    End If
    
    lCurrentItem = lCurrentItem + 1
    fncSetProgress Me, lCurrentItem, oNodeList.length
    DoEvents
    
    If bolCancelValidation Then fncDocumentAttributesContentValid = True: Exit For
  Next
  
ALErr:
  Set oNode = Nothing
  Set oElement = Nothing
  Set oNodeList = Nothing
  Set clsRules = Nothing
  If Not fncDocumentAttributesContentValid Then objEvent.subLog sErrfDACV & "Couldn't load ADTD " & isAbsPath & sALEdtd
  
  'fncInsertTime "oAttributeDTD.fncDocumentAttributesContentValid"
End Function

' This function takes all elements that has a defined attribute to attribute
' correlation behiavour. Test is performed trough comparing attribute contents with
' rules depending on the contents of other attributes within the same element.
Public Function fncIntraElementAttributesCorelationValid(iobjReport As oReport, _
    iobjDOMStructure As Object, enFileType As enuFileType, _
    isAbsPath As String) As Boolean
    
  'fncInsertTime "oAttributeDTD.fncIntraElementAttributesCorelationValid"
    
  fncIntraElementAttributesCorelationValid = False

  Dim objNodeList As Object, clsRules As Object
  Dim templCounter As Long, sALEdtd As String
  
  Set clsRules = CreateObject("DTDParser.cDTDData")
  
  If sAdtdPath = "" Then _
    objEvent.subLog sErrfIEACV & "no adtd specified": Exit Function
  If iobjDOMStructure Is Nothing Then _
    objEvent.subLog sErrfIEACV & "dom structure is null": Exit Function
  
  If Not bolRulesAreLoaded Then If Not fncLoadRulesFiles Then Exit Function
  
'Load the correct ADTD file depending of the filetype
  If enFileType = smil Then
    Set clsRules = objRulesSmil
  ElseIf enFileType = ncc Then
    Set clsRules = objRulesNcc
  ElseIf enFileType = mastersmil Then
    Set clsRules = objRulesMasterSmil
  ElseIf enFileType = smilMediaObText Then
    Set clsRules = objRulesContent
  ElseIf enFileType = nccmultivolume Then
    Set clsRules = objRulesMultivolume
  ElseIf enFileType = discinfo Then
    Set clsRules = objRulesDiscinfo
  Else
    Exit Function
  End If

  Dim objAttCorList As Object 'cDTDAttCorList
  
  fncIntraElementAttributesCorelationValid = True
  If clsRules.clsDTDAllAttCorLists.lAttCorlistCount = 0 Then Exit Function
  
  fncSetProgress Me, 0, _
    clsRules.clsDTDAllAttCorLists.lAttCorlistCount

  ' Go trough each ATTCORLIST found in the ADTD and see if the current
  ' file follows the rules
  For templCounter = 1 To clsRules.clsDTDAllAttCorLists.lAttCorlistCount
    Set objAttCorList = clsRules.clsDTDAllAttCorLists.pAttCorlist(templCounter)
    
    ' Select all elements in the DOM that is of the same type as the current
    ' item in the ATTCORLIST.
    If Not objAttCorList Is Nothing Then
      Set objNodeList = iobjDOMStructure.selectNodes("//" & _
        objAttCorList.sElementName)
    
      If Not fncCheckSingleAttCorList(objNodeList, objAttCorList, iobjReport, _
        isAbsPath) Then
        
        objEvent.subLog sErrfDACV & " error in fncCheckSingleAttCorList " & _
          "while parsing elementtype " & objAttCorList.sElementName
        fncIntraElementAttributesCorelationValid = False
      End If
    Else
      objEvent.subLog sErrfIEACV & "objAttCorList " & templCounter & _
        " in file " & sALEdtd & " is null"
      fncIntraElementAttributesCorelationValid = False
    End If
  
    fncSetProgress Me, templCounter, _
      clsRules.clsDTDAllAttCorLists.lAttCorlistCount
    DoEvents
      
    If bolCancelValidation Then _
      fncIntraElementAttributesCorelationValid = True: Exit For
  Next templCounter

'  Exit Function
ALErr:
  Set objAttCorList = Nothing
  Set objNodeList = Nothing
  Set clsRules = Nothing

  If Not fncIntraElementAttributesCorelationValid Then objEvent.subLog sErrfIEACV & "couldn't load ADTD " & isAbsPath & sALEdtd
  
  'fncInsertTime "oAttributeDTD.fncIntraElementAttributesCorelationValid"
End Function

' This function is validating one attribute against the given rules
Private Function fncCheckSingleAttribute(ioElement As Object, _
  iclsRules As Object, iobjReport As oReport, isAbsPath As String) As Boolean
  
  Dim clsRc As New oDTDRuleChecker, templCounter As Long, bolResult As Boolean
  Dim clsAA As Object, clsAL As Object, clsA As Object
  
  'fncInsertTime "oAttributeDTD.fncCheckSingleAttribute"
  
  fncCheckSingleAttribute = False
  
'Look if the current element has got an ATTLIST entry in the ATDT
  Set clsAA = iclsRules.clsDTDAllAttlists
  For templCounter = 1 To clsAA.lAttlistCount
    If clsAA.pAttlist(templCounter) Is Nothing Then GoTo ErrorH
    If (clsAA.pAttlist(templCounter).sElementName = ioElement.nodeName) Then _
      Exit For
  Next templCounter
  If templCounter > clsAA.lAttlistCount Then _
      fncCheckSingleAttribute = True: GoTo ErrorH
  
  Set clsAL = clsAA.pAttlist(templCounter)
  
  Dim oNodeList As Object, oNode As Object
  Set oNodeList = ioElement.selectNodes("@*")
  
'Go trough each attribute in the element and see if it follows the ADTDs rules
  For Each oNode In oNodeList
    For templCounter = 1 To clsAL.lAttributeCount
      If clsAL.pAttributes(templCounter) Is Nothing Then GoTo ErrorH
      If (clsAL.pAttributes(templCounter).sName = oNode.nodeName) Then Exit For
    Next templCounter
    If templCounter > clsAL.lAttributeCount Then GoTo Skip
        
    Set clsA = clsAL.pAttributes(templCounter)

    clsRc.lBytePos = 1
    If Not clsA.bolCaseSensitive Then
      clsRc.sData = LCase$(oNode.nodeValue)
    Else
      clsRc.sData = oNode.nodeValue
    End If
    clsRc.lDataLength = Len(clsRc.sData)

'If the conformsTo function returns true, the attribute follows the rules
    bolResult = clsRc.conformsTo(, clsA.clsTypeRules)
    If Not (clsRc.lBytePos >= Len(clsRc.sData)) Then bolResult = False
    
    If bolResult = False Then
      fncInsFail2Report iobjReport, ioElement, "adtd.attrContentValid", isAbsPath, _
        "wrong content in attribute: " & oNode.nodeName
    Else
      iobjReport.subInsertSucceededTest
    End If
Skip:
  Next
  
  fncCheckSingleAttribute = True
ErrorH:
  'fncInsertTime "oAttributeDTD.fncCheckSingleAttribute"
End Function

' This function is validating one attribute correlation list against the given
' rules
Private Function fncCheckSingleAttCorList( _
    iobjNodeList As Object, iobjAttCorList As Object, _
    iobjReport As oReport, isFileName As String _
    ) As Boolean

  Dim templCounter As Long, bolResult As Boolean, objElement As Object
  Dim objdecidingattribute As Object
  Dim objAttribute As Object, objRc As New oDTDRuleChecker
  Dim bolFound As Boolean, lInstances As Long, sReturnString As String
  
  Dim sTemp As String
  
  'fncInsertTime "oAttributeDTD.fncCheckSingleAttCorList"
  
  fncCheckSingleAttCorList = False
  If iobjAttCorList Is Nothing Then GoTo ErrorH
  
  lInstances = 0
Iterate:
  
  bolResult = False
  bolFound = False
'Go trough each element in the iobjNodeList
  For Each objElement In iobjNodeList
'Get the first attribute, the one that decides whether we shall continue
    
    If iobjAttCorList.pAttCor(1) Is Nothing Then GoTo ErrorH
    
    Set objdecidingattribute = objElement.getAttributeNode( _
      iobjAttCorList.pAttCor(1).sAttName)
        
    If Not objdecidingattribute Is Nothing Then
      objRc.lBytePos = 1
      If iobjAttCorList.pAttCor(1).bolCaseSensitive Then
        objRc.sData = objdecidingattribute.nodeValue
      Else
        objRc.sData = LCase$(objdecidingattribute.nodeValue)
      End If
      objRc.lDataLength = Len(objRc.sData)

'If the attribute conforms to the demands of the deciding attribute we should continue,
'I.E. if the first attribute is 'name' and has the value 'dc:format' then the [x]
'attribute must contain [y]

      If objRc.conformsTo(, _
         iobjAttCorList.pAttCor(1).clsAttValueRules) Then

'OK, we're trough, this increases the instance counter for this element-attribute
'combination.
         lInstances = lInstances + 1
         bolResult = True

'Now the rest of the attributes must follow their given syntax
'- if they have one - or else there will be an error
         For templCounter = 2 To iobjAttCorList.lAttCorCount
           If iobjAttCorList.pAttCor(templCounter) Is Nothing Then GoTo ErrorH
           
           Set objAttribute = objElement.getAttributeNode( _
             iobjAttCorList.pAttCor(templCounter).sAttName)
           
           bolFound = True
'A demanded attribute is not found, this will generate an error
           If objAttribute Is Nothing Then
             fncInsFail2Report iobjReport, objElement, "adtd.missing", isFileName, _
               objdecidingattribute.nodeName & " = " & "'" & _
               objdecidingattribute.nodeValue & "'" & " : " & _
               "attribute missing: " & iobjAttCorList.pAttCor(templCounter).sAttName

             GoTo Done
           Else
             iobjReport.subInsertSucceededTest
           End If
             
           objRc.lBytePos = 1
           If iobjAttCorList.pAttCor(templCounter).bolCaseSensitive Then
             objRc.sData = objAttribute.nodeValue
           Else
             objRc.sData = LCase$(objAttribute.nodeValue)
           End If
           objRc.lDataLength = Len(objRc.sData)
           
'The attribute is found, but if it doesn't have the value decided in the ADTD an error
'will be reported
           If Not objRc.conformsTo(, _
             iobjAttCorList.pAttCor(templCounter).clsAttValueRules) Then
            
            
             sTemp = objdecidingattribute.nodeName & " = " & "'" & _
               objdecidingattribute.nodeValue & "'" & " : " & _
               "attribute " & iobjAttCorList.pAttCor(templCounter).sAttName & _
               " doesn't follow rules "
               
            If bolAdvancedADTD Then sTemp = sTemp & fncWriteHierarchy( _
                 iobjAttCorList.pAttCor(templCounter).clsAttValueRules)
             
             fncInsFail2Report iobjReport, objElement, "adtd.invalidContent", _
               isFileName, sTemp

             bolResult = False
           Else
             iobjReport.subInsertSucceededTest
           End If
         Next templCounter
      End If
Skip:
    End If
  Next
Done:
  
' Check if this element attribute combination is required and the number of
' instances
  If (iobjAttCorList.sInstance = "#" And (Not lInstances = 1)) Or _
     (iobjAttCorList.sInstance = "#+" And lInstances < 1) Or _
     (iobjAttCorList.sInstance = "#?" And lInstances > 1) Then
           
     sReturnString = iobjAttCorList.sElementName & " "
     For templCounter = 1 To iobjAttCorList.lAttCorCount
       If iobjAttCorList.pAttCor(templCounter) Is Nothing Then GoTo ErrorH
       
       If (bolAdvancedADTD) Or (templCounter = 1) Then _
         sReturnString = sReturnString & _
           iobjAttCorList.pAttCor(templCounter).sAttName & " = " & _
           fncWriteHierarchy( _
             iobjAttCorList.pAttCor(templCounter).clsAttValueRules) & " "
     Next templCounter
  
     iobjReport.fncInsertFailedTest "adtd.elmAttrComb", _
       isFileName, -1, -1, sReturnString
   ElseIf ((iobjAttCorList.sInstance = "#?r" Or _
     iobjAttCorList.sInstance = "#*r") And lInstances = 0) Then
           
     sReturnString = iobjAttCorList.sElementName & " "
     For templCounter = 1 To iobjAttCorList.lAttCorCount
       If iobjAttCorList.pAttCor(templCounter) Is Nothing Then GoTo ErrorH
       
       If (bolAdvancedADTD) Or (templCounter = 1) Then _
         sReturnString = sReturnString & _
           iobjAttCorList.pAttCor(templCounter).sAttName & " = " & _
           fncWriteHierarchy( _
             iobjAttCorList.pAttCor(templCounter).clsAttValueRules) & " "
     Next templCounter

     iobjReport.fncInsertFailedTest "adtd.elmAttrCombRec", _
       isFileName, -1, -1, sReturnString
   Else
     iobjReport.subInsertSucceededTest
   End If

  fncCheckSingleAttCorList = True
ErrorH:
  'fncInsertTime "oAttributeDTD.fncCheckSingleAttCorList"
End Function

' This function writes out the hierarchy of the rules that a certain string must
' follow I.E.
' "(('d'|'D'), ('a'|'A'), ('i'|'I'), ('s'|'S'), ('y'|'Y'), ' 2.02')" for content
' attribute on meta element with name attribute set to 'dc:format'
Public Function fncWriteHierarchy(iclsRules As Object, _
  Optional isFocusOn As Variant) As String
  Dim clsRules As Object, templCounter As Long, tempsString As String
  Dim sFocusOn As String
  
  Set clsRules = iclsRules
  If clsRules Is Nothing Then Exit Function
  
  If Not IsMissing(isFocusOn) Then sFocusOn = LCase$(isFocusOn)
  
    If clsRules.sName = "" Or clsRules.sName = "Rules" Then
    tempsString = tempsString & "("
    For templCounter = 1 To clsRules.lRulesCount
      If clsRules.pRule(templCounter) Is Nothing Then Exit Function
    
      tempsString = tempsString & fncWriteHierarchy( _
        clsRules.pRule(templCounter), sFocusOn)
      
      If templCounter = clsRules.lRulesCount Then Exit For
      
      Select Case clsRules.enChildOccurance
        Case 1
          tempsString = tempsString & "|"
        Case 3
          tempsString = tempsString & ","
      End Select
      
    Next templCounter
    tempsString = tempsString & ")"
  Else
    If LCase$(clsRules.sName) = sFocusOn Then
      tempsString = tempsString & " <<" & clsRules.sName & ">> "
    Else
      tempsString = tempsString & " " & clsRules.sName & " "
    End If
  End If
  
  Select Case clsRules.enMainOccurance
    Case 1 'DTD_OneOrMore
      tempsString = tempsString & "+"
    Case 3 'DTD_ZeroOrOne
      tempsString = tempsString & "?"
    Case 4 'DTD_ZeroOrMore
      tempsString = tempsString & "*"
  End Select
    
  fncWriteHierarchy = tempsString
End Function
