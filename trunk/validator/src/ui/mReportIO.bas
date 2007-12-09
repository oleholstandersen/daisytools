Attribute VB_Name = "mReportIO"
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
  "      div.lightmode { margin-left: 1em; border: 1px solid red; }" & vbCrLf & _
  "    </style>" & vbCrLf & _
  "  </head>" & vbCrLf & _
  "  <body>" & vbCrLf
  
Private Const sXhtmlEnd As String = "  </body>" & vbCrLf & "</html>"

Public Function fncSaveReport(typCandidateInfo As tCandidateInfo) As Boolean
  Dim objDom As New MSXML2.DOMDocument40, lCounter As Long
  
  Dim objReportItem As oReportItem, oFSO As Object, oFile As Object
  Dim lIteration As Long, sXhtmlOutput As String, sTemp As String
  Dim msgbResult As VbMsgBoxResult
    
  Set oFSO = CreateObject("scripting.FileSystemObject")
  Set oFile = oFSO.Createtextfile(propTempPath & "tempreport.html")

  sXhtmlOutput = sXhtmlStart
  
  Select Case typCandidateInfo.lCandidateType
    Case TYPE_SINGLEDTB
      sTemp = "dtb"
    Case TYPE_MULTIVOLUME
      sTemp = "TYPE_MULTIVOLUME dtb"
    Case TYPE_SINGLE_NCC
      sTemp = "single ncc"
    Case TYPE_SINGLE_SMIL
      sTemp = "single smil"
    Case TYPE_SINGLE_MSMIL
      sTemp = "single mastersmil"
    Case TYPE_SINGLE_CONTENTDOC
      sTemp = "single content document"
    Case TYPE_SINGLE_DISCINFO
      sTemp = "single discinfo"
  End Select
  
  sXhtmlOutput = sXhtmlOutput & "    <h1 class='" & typCandidateInfo.lCandidateType & "' id='" & typCandidateInfo.sAbsPath & "'>Validator report for " & sTemp & _
    " at " & typCandidateInfo.sAbsPath & "</h1>" & vbCrLf & "    <br />" & vbCrLf
  
  If frmMain.mnuLightMode.Checked Then
    sXhtmlOutput = sXhtmlOutput & "    <div class='lightmode'>This validation was done using <em>light mode</em>. Full conformance checking has <strong>not</strong> been performed.</div>" & vbCrLf
  End If
  
  frmMain.MousePointer = 11
  sCurrentState = "building report "
  lIntProgMax = (typCandidateInfo.objReport.lFailedTestCount - 1) * 3
  
  lIteration = 0
Again:
  For lCounter = 0 To typCandidateInfo.objReport.lFailedTestCount - 1
    typCandidateInfo.objReport.fncRetrieveFailedTestItem lCounter, objReportItem
    
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
    
    If Len(sXhtmlOutput) > 64000 Then
      oFile.write (sXhtmlOutput)
      sXhtmlOutput = ""
    End If
    
    lIntProgress = lCounter: DoEvents
  Next lCounter
  
  If lIteration < 2 Then lIteration = lIteration + 1: GoTo Again
  
  sXhtmlOutput = sXhtmlOutput & sXhtmlEnd
  
  frmMain.MousePointer = 1
  sCurrentState = "idle"
  
  oFile.write (sXhtmlOutput)
  oFile.Close
  
  Set oFile = oFSO.opentextfile(propTempPath & "tempreport.html")
  sXhtmlOutput = oFile.readall
  oFile.Close
  
  oFSO.deletefile (propTempPath & "tempreport.html")
  
  fncSaveFile sXhtmlOutput, report
ErrorH:
End Function

Public Function fncLoadReport() As Boolean
  frmMain.CommonDialog1.Filter = "Report file (*.html)|*.html"
  SetSingleSelectFlags
  On Error GoTo ErrorH
  
  frmMain.CommonDialog1.ShowOpen
  
  Dim objDom As New MSXML2.DOMDocument40, objNode As IXMLDOMNode
  Dim objNodeList As IXMLDOMNodeList, objNode2 As IXMLDOMNode
  Dim sTestId As String, sAbsPath As String, lLine As Long, lColumn As Long
  Dim sComment As String, bolErrInReport As Boolean
  
  objDom.async = False
  objDom.validateOnParse = False
  objDom.preserveWhiteSpace = False
  objDom.resolveExternals = False
  objDom.setProperty "SelectionLanguage", "XPath"
  
  If Not objDom.Load(frmMain.CommonDialog1.FileName) Then
    MsgBox "File couldn't be loaded!", vbOKOnly, "File error"
    Exit Function
  End If
  
  bolErrInReport = True
  
  ReDim Preserve aCandidateQueue(lCandidatesAdded)
  
  Set objNode = objDom.selectSingleNode("//h1/@class")
  If objNode Is Nothing Then GoTo ErrorH
  aCandidateQueue(lCandidatesAdded).lCandidateType = CLng(objNode.nodeValue)
  
  Set objNode = objDom.selectSingleNode("//h1/@id")
  If objNode Is Nothing Then GoTo ErrorH
  aCandidateQueue(lCandidatesAdded).sAbsPath = objNode.nodeValue
  
  Set aCandidateQueue(lCandidatesAdded).objReport = New oReport
  
  Set objNodeList = objDom.selectNodes("//body/child::div")
  For Each objNode In objNodeList
    Set objNode2 = objNode.selectSingleNode("@class")
    If Not objNode2 Is Nothing Then
      sTestId = objNode2.nodeValue
      
      Set objNode2 = objNode.selectSingleNode("div/span[@class='absPath']")
      If Not objNode2 Is Nothing Then sAbsPath = objNode2.Text
      
      Set objNode2 = objNode.selectSingleNode("div/span[@class='line']")
      If Not objNode2 Is Nothing Then lLine = CLng(objNode2.Text)
      
      Set objNode2 = objNode.selectSingleNode("div/span[@class='column']")
      If Not objNode2 Is Nothing Then lColumn = CLng(objNode2.Text)
      
      Set objNode2 = objNode.selectSingleNode("div/span[@class='comment']")
      If Not objNode2 Is Nothing Then sComment = objNode2.Text
      
      aCandidateQueue(lCandidatesAdded).objReport.fncInsertFailedTest sTestId, _
        sAbsPath, lLine, lColumn, sComment
    End If
  Next objNode
  lCandidatesAdded = lCandidatesAdded + 1
  
  bolErrInReport = False
ErrorH:
  If bolErrInReport Then MsgBox "Error in reportfile!", vbOKOnly, "Error"
  If Not lCandidatesAdded = 0 Then fncAddCandidateToTree aCandidateQueue(lCandidatesAdded - 1)
End Function
