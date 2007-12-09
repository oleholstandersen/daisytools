VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAbstractDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public lAbstractType As Long
Public sFileName As String
Public lInputContentDocSource As String 'ref to inputarray content doc if content
Public oDom As New MSXML2.DOMDocument40

Private Sub Class_Initialize()
 'debug.Print "cAbstractDocument.initialize"
 oDom.async = False
 oDom.validateOnParse = False
 oDom.resolveExternals = False
 oDom.preserveWhiteSpace = False
 oDom.setProperty "SelectionLanguage", "XPath"
 oDom.setProperty "SelectionNamespaces", oBruno.oDriver.sDefaultNameSpace
 oDom.setProperty "NewParser", True
End Sub

Public Function fncSetFileName(ByRef sCandidateFilename As String) As Boolean
  fncSetFileName = False
  sFileName = fncTruncToValidUriChars(sCandidateFilename)
  fncSetFileName = True
End Function

Private Function fncTruncToValidUriChars(ByVal sString As String) As String

    Dim i As Long
    Dim sErr As String

    'do some pretty from windows-1252:
    ' space to underscore
    sString = Replace$(sString, Chr(32), Chr(95), , , vbBinaryCompare)
    ' german double s to ss
    sString = Replace$(sString, Chr(223), Chr(115) & Chr(115), , , vbTextCompare)
    '224-227 to a
    sString = Replace$(sString, Chr(224), Chr(97), , , vbTextCompare)
    sString = Replace$(sString, Chr(225), Chr(97), , , vbTextCompare)
    sString = Replace$(sString, Chr(226), Chr(97), , , vbTextCompare)
    sString = Replace$(sString, Chr(227), Chr(97), , , vbTextCompare)
    ' ä/Ä to ae
    sString = Replace$(sString, Chr(228), Chr(97) & Chr(101), , , vbTextCompare)
    ' å/Å to aa
    sString = Replace$(sString, Chr(229), Chr(97) & Chr(97), , , vbTextCompare)
    '230: danish ae to ae
    sString = Replace$(sString, Chr(230), Chr(97) & Chr(101), , , vbTextCompare)
    '231: ccedil to c
    sString = Replace$(sString, Chr(231), Chr(99), , , vbTextCompare)
    '232-235 to e
    sString = Replace$(sString, Chr(232), Chr(101), , , vbTextCompare)
    sString = Replace$(sString, Chr(233), Chr(101), , , vbTextCompare)
    sString = Replace$(sString, Chr(234), Chr(101), , , vbTextCompare)
    sString = Replace$(sString, Chr(235), Chr(101), , , vbTextCompare)
    '236-239 to i
    sString = Replace$(sString, Chr(236), Chr(105), , , vbTextCompare)
    sString = Replace$(sString, Chr(237), Chr(105), , , vbTextCompare)
    sString = Replace$(sString, Chr(238), Chr(105), , , vbTextCompare)
    sString = Replace$(sString, Chr(239), Chr(105), , , vbTextCompare)
    '240 (icelandic d) to d
    sString = Replace$(sString, Chr(240), Chr(100), , , vbTextCompare)
    '241(ana) to n
    sString = Replace$(sString, Chr(241), Chr(110), , , vbTextCompare)
    '242-245 to o
    sString = Replace$(sString, Chr(242), Chr(111), , , vbTextCompare)
    sString = Replace$(sString, Chr(243), Chr(111), , , vbTextCompare)
    sString = Replace$(sString, Chr(244), Chr(111), , , vbTextCompare)
    sString = Replace$(sString, Chr(245), Chr(111), , , vbTextCompare)
    ' ö/Ö to oe
    sString = Replace$(sString, Chr(246), Chr(111) & Chr(101), , , vbTextCompare)
    ' danish ö to oe
    sString = Replace$(sString, Chr(248), Chr(111) & Chr(101), , , vbTextCompare)
    '249-252 to u
    sString = Replace$(sString, Chr(249), Chr(117), , , vbTextCompare)
    sString = Replace$(sString, Chr(250), Chr(117), , , vbTextCompare)
    sString = Replace$(sString, Chr(251), Chr(117), , , vbTextCompare)
    sString = Replace$(sString, Chr(252), Chr(117), , , vbTextCompare)
    '253 and 255 to y
    sString = Replace$(sString, Chr(253), Chr(121), , , vbTextCompare)
    sString = Replace$(sString, Chr(255), Chr(121), , , vbTextCompare)
        
    'now loop and do the rest ugly (trunc char)
    sErr = sString

    For i = 1 To Len(sString)
        If Not fncIsValidUriChar(Asc(Mid$(sString, i, 1))) Then
            sErr = Replace$(sErr, Mid$(sString, i, 1), "", , , vbTextCompare)
        End If
    Next i

    If sErr = "" Then sErr = "dtb_"
     
    sString = sErr
                
    fncTruncToValidUriChars = sString

End Function

Private Function fncIsValidUriChar(sChar As String) As Boolean

    fncIsValidUriChar = False
    
    If (CInt(sChar) > 44 And CInt(sChar) < 58) Or _
       (CInt(sChar) > 63 And CInt(sChar) < 91) Or _
       (CInt(sChar) > 94 And CInt(sChar) < 123) _
       Then
        'char is ok
    Else
        Exit Function
    End If
    
    fncIsValidUriChar = True

End Function

Private Sub Class_Terminate()
  Set oDom = Nothing
End Sub
