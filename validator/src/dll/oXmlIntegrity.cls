VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oXmlIntegrity"
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

' This function verifies that the given file is a wellformed XML file
Public Function fncIsWellformedXML( _
    ByRef iobjReport As oReport, _
    ByVal isAbsPath As String, _
    ByRef oDOMStructure As Object _
    ) As Boolean
    
    'fncInsertTime "oXMLIntegrity.fncIsWellformedXML"

    fncIsWellformedXML = False

    'Dim tmpDoc As New object
    Dim tmpDoc As Object
    
    On Error Resume Next
    
    Set tmpDoc = CreateObject("Msxml2.DOMDocument.4.0")
    
    'this fnc is also used for generic loading
    
    On Error GoTo ErrH

    Dim oFSO As Object, oFile As Object, sData As String
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFSO.opentextfile(isAbsPath)
    
    If Not oFile.AtEndOfStream Then 'if the file is not empty
        sData = oFile.ReadAll
        oFile.Close
    Else
        iobjReport.fncInsertFailedTest "xi.isWellformed", isAbsPath, -1, -1, "file is empty"
        oFile.Close
        GoTo ErrH
    End If
    
    Dim lValue1 As Long, lValue2 As Long, lIteration As Long, sChar As String
    lValue1 = InStr(1, sData, "xmlns", vbTextCompare)
    If lValue1 > 0 Then
      lValue2 = lValue1 + 5
      Do
        Select Case Mid$(sData, lValue2, 1)
          Case Chr(34), "'"
            If lIteration = 1 Then
              lIteration = 2: sChar = Mid$(sData, lValue2, 1)
            ElseIf (lIteration = 2) And (sChar = Mid$(sData, lValue2, 1)) Then
              sData = Replace(sData, Mid$(sData, lValue1, lValue2 - lValue1 + 1), "")
              Exit Do
            End If
          Case "="
            If lIteration = 0 Then
                lIteration = 1
            Else
                objEvent.subLog "error in oXmlIntegrity.fncIsWellformedXML on" & isAbsPath
                Exit Function
            End If

          Case ">", "<"
            objEvent.subLog "error in oXmlIntegrity.fncIsWellformedXML on" & isAbsPath
            Exit Function
        End Select

        lValue2 = lValue2 + 1
      Loop
    End If
    
    Set oFile = oFSO.CreateTextFile(sTempPath & "temporary.dat")
    oFile.Write sData
    oFile.Close
    
    Set oFile = Nothing
    Set oFSO = Nothing
    
    'use a temporary doc since these properties must be set for wellformedness checking
    tmpDoc.async = False
    tmpDoc.validateOnParse = False
    tmpDoc.preserveWhiteSpace = True
    tmpDoc.resolveExternals = False
    tmpDoc.setProperty "SelectionLanguage", "XPath"
    tmpDoc.setProperty "NewParser", True

    On Error GoTo ErrH2
    
    If Not tmpDoc.Load(sTempPath & "temporary.dat") Then
        iobjReport.fncInsertFailedTest "xi.isWellformed", _
          isAbsPath, tmpDoc.parseError.Line, tmpDoc.parseError.linepos, _
          tmpDoc.parseError.reason
    Else
        iobjReport.subInsertSucceededTest
        Set oDOMStructure = tmpDoc
    End If

    Dim objNode As Object
    Set objNode = tmpDoc.selectSingleNode("//body")

ErrH2:
'    Kill sTempPath & "temporary.dat"
ErrH:
    Set tmpDoc = Nothing
    If Err.Number = 62 Then iobjReport.fncInsertFailedTest "xi.isWellformed", isAbsPath
    Err.Clear
    
    fncIsWellformedXML = True
    
    'fncInsertTime "oXMLIntegrity.fncIsWellformedXML"
End Function

' This function verifies that the given file has the supplied doctype
Public Function fncHasAndIsDocumenttype(ByRef iobjReport As oReport, _
    ByVal oDOMStructure As Object, _
    sDemandDocumenttype As enuDocuType, _
    isAbsPath As String) As Boolean

    Dim k As Long
    Dim oDcNode As Object, oRootElemNode As Object
    Dim sExpectedDocType As String, objNode As Object
    
    'fncInsertTime "oXMLIntegrity.fncHasAndIsDocumenttype"
    'checks:
    'a) that doctype exists and
    'b) that doctype assumed from filename is actually the xml document inside the file
            
            fncHasAndIsDocumenttype = False
            'On Error GoTo ErrH
            
            'find the doctype node
            If Not oDOMStructure Is Nothing Then
                For k = 0 To oDOMStructure.childNodes.length - 1
                    If oDOMStructure.childNodes.Item(k).nodeType = NODE_DOCUMENT_TYPE Then
                       Set oDcNode = oDOMStructure.childNodes.Item(k)
                       Exit For
                    End If
                Next
                
                'if doctypedecl node was not found
                If oDcNode Is Nothing Then
                    iobjReport.fncInsertFailedTest "xi.hasDoctype", isAbsPath, -1, -1
                    'Exit Function
                Else 'node was found, so check that root element name equals expectation
                    iobjReport.subInsertSucceededTest
                    
                    Select Case sDemandDocumenttype
                      Case smil10
                        sExpectedDocType = "smil"
                      Case xhtml10
                        sExpectedDocType = "html"
                    End Select
                    
                    Set oRootElemNode = oDOMStructure.documentElement
                    
                    'ke20030602
                    Dim sPID As String
                    For Each objNode In oDcNode.Attributes
                      If objNode.nodeName = "PUBLIC" Then
                        sPID = oDcNode.Attributes.Item(0).nodeValue
                        Exit For
                      End If
                    Next objNode
                    
                    Dim lpIDValue As Long
                    
                    lpIDValue = 0
                    If sExpectedDocType = "html" And (sPID = "-//W3C//DTD XHTML 1.0 Strict//EN" Or sPID = "-//W3C//DTD XHTML 1.0 Transitional//EN") Then lpIDValue = 1
                    If sExpectedDocType = "smil" And sPID = "-//W3C//DTD SMIL 1.0//EN" Then lpIDValue = 1
                    
                    'If sExpectedDocType = "html" And (sPID = "-//W3C//DTD XHTML 1.1//EN") Then lpIDValue =
                    
                    If Not objNode Is Nothing Then
                      If lpIDValue = 1 Then
                        iobjReport.subInsertSucceededTest
                      Else
                        fncInsFail2Report iobjReport, objNode, _
                          "xi.doctypeValid", isAbsPath, sPID
                      End If
                    End If
                                        
                    If (oDcNode.baseName = sExpectedDocType) And _
                      ((Not oRootElemNode Is Nothing) And _
                      (oRootElemNode.nodeName = sExpectedDocType)) And (lpIDValue = 1) Then
                      
                        iobjReport.subInsertSucceededTest
                    Else
                        iobjReport.fncInsertFailedTest "xi.isDoctype", isAbsPath, -1, -1
                    End If
                    
                End If
            End If 'if not is nothing
            fncHasAndIsDocumenttype = True
ErrH:
        If Not fncHasAndIsDocumenttype Then objEvent.subLog ("failure in fncHasAndIsDocumenttype on " & fncGetFileName(isAbsPath))
        Set oDcNode = Nothing
        Set oRootElemNode = Nothing
    
    'fncInsertTime "oXMLIntegrity.fncHasAndIsDocumenttype"
End Function
