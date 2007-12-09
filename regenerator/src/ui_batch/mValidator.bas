Attribute VB_Name = "mValidator"
' Daisy 2.02 Regenerator Batch UI
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

' This constant is the head of a XHTML file
Private Const sXhtmlStart As String = _
  "<?xml version='1.0' encoding='utf-8'?>" & vbCrLf & _
  "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' " & _
  "'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>" & vbCrLf & _
  "<html>" & vbCrLf & _
  "  <head>" & vbCrLf & _
  "    <style>" & vbCrLf & _
  "      html { background: rgb(210,210,240); margin-right: 4em; margin-left: 4em; font-family: arial; }" & vbCrLf & _
  "      h1 { font-size: 18; }" & vbCrLf & _
  "      div { margin-left: +2em; font-size: 16; }" & vbCrLf & _
  "      div.suggestion { margin-left: +3em; font-size: 12; }" & vbCrLf & _
  "      div.warning { margin-left: 1em; font-size: 20; }" & vbCrLf & _
  "    </style>" & vbCrLf & _
  "  </head>" & vbCrLf & _
  "  <body>" & vbCrLf
  
' This constant is the bottom of a XHTML file
Private Const sXhtmlEnd As String = "  </body>" & vbCrLf & "</html>"

Private sReportPath As String

' Misc path variables
Public sVTMPath As String, sExtPath As String, sTempPath As String

' Misc settings variables
Public bolUseValidator As Boolean
Public bolIncludeNCErrors As Boolean
Public bolIncludeWarnings As Boolean
Public lTimeFluctuation As Long
Public bolIncludeADVADTD As Boolean
Public bolCreateStandalone As Boolean
Public bolValidatorLightMode As Boolean
Public sStandalonePath As String
Public bolSyncWValidator As Boolean

Public Function fncValidate(ByVal isNccPath As String) As Boolean
Dim objValidateDTB As Object
Dim lCounter As Long
Dim objReportItem As Object
  
  On Error Resume Next
  
  'objUI.fncAddLog "<validating>", True
  
  Set objValidateDTB = CreateObject("ValidatorEngine.oValidateDTB")
  If objValidateDTB Is Nothing Then
    objUI.fncAddLog "<error from='ui' in='fncValidate'>Validator dll not registered, aborting.</error>", True
    Exit Function
  End If
  
' Parse the location of the report BEFORE validation, because if the user has supplied
' the same directory for the report as the regenerated book, it will be moved with
' the book in the final stage
  Dim sFileName As String, objFile As Object
  
'mg 20030325 below didnt really work
'  If aJobItems(lCurrentJob).bolSaveSame Then
'    Dim sTemp As String
'    fncParsePathWithConstants sStandalonePath, sTemp, lCurrentJob
'    sReportPath = fncGetPathName(aJobItems(lCurrentJob).sPath) + fncGetUriFileName(sTemp)
'  Else
'    fncParsePathWithConstants sStandalonePath, sReportPath, lCurrentJob
'  End If
'instead:
  fncParsePathWithConstants sStandalonePath, sReportPath, lCurrentJob
  
  sReportPath = fncStripIdAddPath(sReportPath, aJobItems(lCurrentJob).sPath)
  sReportPath = fncGetPathName(sReportPath)
  
' Default the jobs validation status to true
  aJobItems(lCurrentJob).bolValResult = True

' Create a new oValidateDTB instance
'  Set objValidateDTB = CreateObject("ValidatorEngine.oValidateDTB")

' All paths MUST end with a slash '\'
  If Not Right$(isNccPath, 1) = "\" Then isNccPath = isNccPath & "\"
  
' Validate job
  fncValidate = objValidateDTB.fncValidate(isNccPath)
'  Set aJobItems(lCurrentJob).objReport = objValidateDTB.objReport
  
' Go trough all errors
  objUI.fncAddLog "<report>", True
  For lCounter = 0 To objValidateDTB.objReport.lFailedTestCount - 1
    objValidateDTB.objReport.fncRetrieveFailedTestItem lCounter, objReportItem
    
    With aJobItems(lCurrentJob)
' Ok, an error was found, this means that the book has failed
      .bolValResult = False

' The following lines are determining the WORST error type and error class that has
' occured within the job. Lowest priority is type="warning" class="" and highest
' priority is type="error" class="critical"
      
      If objReportItem.sFailType = "warning" And .sErrorType = "" Then _
        .sErrorType = "warning": .sErrorClass = ""
    
      If objReportItem.sFailType = "error" Then
        .sErrorType = "error"
        
        If objReportItem.sFailClass = "non-critical" And .sErrorClass = "" Then _
          .sErrorClass = "non-critical"
          
        If objReportItem.sFailClass = "critical" Then .sErrorClass = "critical"
      End If
    
      objUI.fncAddLog _
        "<item type='" & objReportItem.sFailType & _
        "' class='" & objReportItem.sFailClass & "'>" & _
        objReportItem.sFailType & " @ " _
        & objReportItem.sAbsPath & " [" & _
        objReportItem.lLine & ":" & objReportItem.lColumn & "]" & " " & _
        objReportItem.sShortDesc & ", " & objReportItem.sComment & "</item>", True
    End With
  Next lCounter ' depending on the UI settings
  
  If objValidateDTB.objReport.lFailedTestCount = 0 Then
    objUI.fncAddLog "<congrats>No errors or warnings reported by validator</congrats>", True
  End If
  
  objUI.fncAddLog "</report>", True

' If the user has chosen to create a standalone validator report, do it
  If bolCreateStandalone Then fncSaveReport objValidateDTB.objReport
  
  Set objValidateDTB.objReport = Nothing
  Set objValidateDTB = Nothing
  
  'objUI.fncAddLog "</validating>", True
  
  If Not objValidatorUserControl.fncdeinitializevalidator Then Exit Function
End Function

Public Function fncSaveReport(iobjReport As Object) As Boolean
Dim lCounter As Long
Dim objReportItem As Object
Dim oFile As Object
Dim lIteration As Long
Dim sXhtmlOutput As String
Dim sTemp As String
Dim msgbResult As VbMsgBoxResult
Dim sFileName As String
Dim objFile As Object

  ' Function for saving a validation report XHTML file that is readable by
  ' Daisy 2.02 Validator ((c) TPB 2002)
  
  sXhtmlOutput = sXhtmlStart
  
  sTemp = "dtb"
  
  sXhtmlOutput = sXhtmlOutput & "    <h1 class='0' id='" & aJobItems(lCurrentJob).sPath & "'>Validator report for " & sTemp & _
    " at " & aJobItems(lCurrentJob).sPath & "</h1>" & vbCrLf & "    <br />" & vbCrLf

' Create the report file
  fncCreateDirectoryChain sReportPath
  
  Dim sFileName2 As String
  If aJobItems(lCurrentJob).sID <> "" Then
    sFileName2 = "validator_report_" & Trim$(aJobItems(lCurrentJob).sID) & ".html"
  Else
    sFileName2 = "validator_report.html"
  End If
  
  Set objFile = oFSO.createtextfile(sReportPath & sFileName2)
    
  If iobjReport.lFailedTestCount = 0 Then
    sXhtmlOutput = sXhtmlOutput & "<div>No errors reported by validator</div>" & vbCrLf
  End If
  
  If bolValidatorLightMode Then
    sXhtmlOutput = sXhtmlOutput & "<div>Light Mode was used.</div>" & vbCrLf
  End If
    
  lIteration = 0
Again:

  For lCounter = 0 To iobjReport.lFailedTestCount - 1
    iobjReport.fncRetrieveFailedTestItem lCounter, objReportItem
    
    With objReportItem
      If (LCase$(.sFailType) = "error" And LCase$(.sFailClass) = "critical" And lIteration = 0) Or _
         (LCase$(.sFailType) = "error" And LCase$(.sFailClass) = "non-critical" And lIteration = 1) Or (LCase$(.sFailType) = "warning" And lIteration = 2) Then
        
        sXhtmlOutput = sXhtmlOutput & "    <div class='" & .sTestId & "'>" & vbCrLf
        
        sXhtmlOutput = sXhtmlOutput & "      <div class='failType'>" & _
          .sFailType
        If lIteration < 2 Then _
          sXhtmlOutput = sXhtmlOutput & "[<span class='failClass'>" & _
            .sFailClass & "</span>]"
        sXhtmlOutput = sXhtmlOutput & "</div>" & vbCrLf
        
        sXhtmlOutput = sXhtmlOutput & "      <div class='shortDesc'>" & _
          .sShortDesc & "</div>" & vbCrLf
        
        If Not .sLongDesc = "" Then _
          sXhtmlOutput = sXhtmlOutput & "      <div class='longDesc'>" & .sLongDesc & "</div>" & vbCrLf
        
        If Not .sAbsPath = "" Then _
          sXhtmlOutput = sXhtmlOutput & "      <div>" & vbCrLf & _
            "        <span class='absPath'>" & .sAbsPath & "</span>" & vbCrLf & _
            "        <span>[</span>" & vbCrLf & _
            "        <span class='line'>" & .lLine & "</span>" & vbCrLf & _
            "        <span>:</span>" & vbCrLf & _
            "        <span class='column'>" & .lColumn & "</span>" & vbCrLf & _
            "        <span>]</span>" & vbCrLf & _
            "      </div>" & vbCrLf
                
        If Not (.sComment = "" And .sLink = "") Then
          sXhtmlOutput = sXhtmlOutput & "      <div>" & vbCrLf
          
          If Not .sComment = "" Then _
            sXhtmlOutput = sXhtmlOutput & "        <span class='comment'>" & _
            .sComment & "</span>" & vbCrLf
          If Not .sLink = "" Then _
            sXhtmlOutput = sXhtmlOutput & "        <span class='link'><a href='" & _
            .sLink & "'>" & .sLink & "</a></span>" & vbCrLf
        
          sXhtmlOutput = sXhtmlOutput & "      </div>" & vbCrLf
        End If
    
        sXhtmlOutput = sXhtmlOutput & "    </div>" & vbCrLf
        sXhtmlOutput = sXhtmlOutput & "    <br />" & vbCrLf
      End If
    End With
    
    
' If our textbuffer exeeds 64000 characters, write it. This technique is used because
' it's too slow writing all the time and VB:s string handling is too worthless to
' handle the size of the reports we're creating (the program get's slooooooowed down).
    If Len(sXhtmlOutput) > 64000 Then
      objFile.write (sXhtmlOutput)
      sXhtmlOutput = ""
    End If
    DoEvents
  Next lCounter
    
  If lIteration < 2 Then lIteration = lIteration + 1: GoTo Again
  
  sXhtmlOutput = sXhtmlOutput & sXhtmlEnd
  
  objFile.write (sXhtmlOutput)
  objFile.Close
ErrorH:
End Function

Public Function fncInitValidator() As Boolean
  fncInitValidator = False
  objUI.bolBusy = True

  objUI.fncAddLog "Initializing validator engine...", False
  
  On Error Resume Next
  
' Get the regenerator control object
  Set objValidatorUserControl = CreateObject("ValidatorEngine.oUserControl")
  If objValidatorUserControl Is Nothing Then
    objUI.fncAddLog "<message from='ui' in='fncInitValidator'>Validator dll not registered, validation will not be avaliable</message>", True
    bolValidatorExists = False
    GoTo ErrorH
  Else
    Set objValidatorUserControl.xobjevent.objowner = objUI
    bolValidatorExists = True
  End If
  
' Set paths and settings in the validator engine

  objValidatorUserControl.fncSetAdtdPath sExtPath
  objValidatorUserControl.fncSetDTDPath sExtPath
  objValidatorUserControl.fncSetTempPath sTempPath
  objValidatorUserControl.fncSetVtmPath sVTMPath
  objValidatorUserControl.fncSetTimeSpan lTimeFluctuation
  objValidatorUserControl.fncSetAdvancedADTD bolIncludeADVADTD
  objValidatorUserControl.fncSetLightMode bolValidatorLightMode
  If objValidatorUserControl.fncInitializeValidator Then fncInitValidator = True
ErrorH:
  objUI.fncAddLog "Done.", False
  objUI.bolBusy = False
End Function

Public Function fncDeinitValidator() As Boolean
  If Not bolValidatorExists Then Exit Function
  If Not objValidatorUserControl.fncdeinitializevalidator Then Exit Function
  'fncAddMemLog "Validatorengine deinitialized"
  fncDeinitValidator = True
End Function

