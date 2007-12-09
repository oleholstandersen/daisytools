Attribute VB_Name = "IteratorMdlReport"
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

Public Function fncSaveReport(typCandidateInfo As tCandidateInfo, bolIsValid As Boolean) As Boolean
  Dim objDom As New MSXML2.DOMDocument40, lCounter As Long
    objDom.async = False
    objDom.validateOnParse = False
    objDom.resolveExternals = False
    objDom.preserveWhiteSpace = False
    objDom.setProperty "SelectionLanguage", "XPath"
    objDom.setProperty "SelectionNamespaces", "xmlns:xht='http://www.w3.org/1999/xhtml'"
    objDom.setProperty "NewParser", True
  
  Dim sDtbId As String
  Dim objReportItem As oReportItem, oFile As Object
  Dim lIteration As Long, sXhtmlOutput As String, sTemp As String
  Dim msgbResult As VbMsgBoxResult
      
  sDtbId = fncGetDtbId(typCandidateInfo.sAbsPath, objDom)
      
  Set oFile = oFSO.Createtextfile(propTempPath & "tempreport.html")
  'if there were no errors, output an empty file:
  If typCandidateInfo.objReport.lFailedTestCount = 0 Then GoTo emptyfile
  
  sXhtmlOutput = sXhtmlStart
  sTemp = "dtb"
  sXhtmlOutput = sXhtmlOutput & "    <h1 class='" & typCandidateInfo.lCandidateType & "' id='" & typCandidateInfo.sAbsPath & "'>Validator report for " & sTemp & _
    " at " & typCandidateInfo.sAbsPath & "</h1>" & vbCrLf & "    <br />" & vbCrLf
  
  If bolLightMode Then
    sXhtmlOutput = sXhtmlOutput & "    <div class='lightmode'>This validation was done using <em>light mode</em>. Full conformance checking has <strong>not</strong> been performed.</div>" & vbCrLf
  End If
    
  lIteration = 0
Again:
  For lCounter = 0 To typCandidateInfo.objReport.lFailedTestCount - 1
    typCandidateInfo.objReport.fncRetrieveFailedTestItem lCounter, objReportItem
    
    With objReportItem
      If (LCase$(.sFailType) = "error" And LCase$(.sFailClass) = "critical" And lIteration = 0) Or _
         (LCase$(.sFailType) = "error" And LCase$(.sFailClass) = "non-critical" And lIteration = 1) Then
         '(LCase$(.sFailType) = "error" And LCase$(.sFailClass) = "non-critical" And lIteration = 1) Or (LCase$(.sFailType) = "warning" And lIteration = 2) Then
        
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
    
    'lIntProgress = lCounter: DoEvents
  Next lCounter
  
  If lIteration < 2 Then lIteration = lIteration + 1: GoTo Again
  
  sXhtmlOutput = sXhtmlOutput & sXhtmlEnd
  oFile.write (sXhtmlOutput)
  
emptyfile:
  oFile.Close
  Dim sDestinationFolder As String
  
  If bolIsValid Then
   sDestinationFolder = sReportDir & "pass\"
  Else
   sDestinationFolder = sReportDir & "fail\"
  End If
  
  If Not fncFolderExists(sDestinationFolder) Then
    fncCreateDirectoryChain (sDestinationFolder)
  End If
  
  oFSO.copyfile propTempPath & "tempreport.html", sDestinationFolder & sDtbId & ".html", True
  oFSO.deletefile (propTempPath & "tempreport.html")
  
  fncSaveReport = True
  
ErrorH:
Set objDom = Nothing
End Function

Private Function fncGetDtbId(sNccPath As String, objDom As MSXML2.DOMDocument40) As String
Dim oNode As IXMLDOMNode
Dim sIdString As String, sFile As String

 If objDom.Load(sNccPath) Then
   Set oNode = objDom.selectSingleNode("//xht:meta[@name='dc:identifier']/@content") '[/@name='dc:identifier']/@content"
   If Not oNode Is Nothing Then
     sIdString = fncTruncToValidUriChars(oNode.Text) & "_"
   End If
 End If
 Set oNode = Nothing
 
 'if there was no id string or length was zero
 
 If Len(sIdString) < 1 Then
   sIdString = "unkownId_"
 End If
 
 sIdString = sIdString & Format(Date, "yyyy-mm-dd") & Time
 sIdString = Replace(sIdString, ":", "")
 sIdString = Replace(sIdString, "-", "")
 
 fncGetDtbId = sIdString

End Function
