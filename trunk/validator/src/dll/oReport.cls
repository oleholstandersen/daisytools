VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

' Number of tests failed in object instance
Public lFailedTestCount As Long
' Number of tests not failed in object instance
Public lSucceededTestCount As Long

' Array containing failed tests (uses oReportItem class as type)
Private aFailedTestItems() As oReportItem

' Used to prevent event raising while adding tests from other report objects
Private bolNoRaise As Boolean

' This function increases the "lSucceededTestCount" variable
Public Sub subInsertSucceededTest()
  lSucceededTestCount = lSucceededTestCount + 1
  If Not bolNoRaise Then objEvent.subSucceededTest
End Sub

' This function inserts a new test into the "aFiledTestItems" array
'
Public Function fncInsertFailedTest(isTestId As String, _
        Optional isAbsPath As Variant, _
        Optional ilLine As Variant, _
        Optional ilColumn As Variant, _
        Optional ivComment As Variant) As Boolean
  
  fncInsertFailedTest = False
  
  On Error GoTo ErrH
  
  ReDim Preserve aFailedTestItems(lFailedTestCount) 'zero-based!
  
  ' since reportitem is a class and not a type, set is needed to allocate memory
  Set aFailedTestItems(lFailedTestCount) = New oReportItem
  
  With aFailedTestItems(lFailedTestCount)
    .sTestId = isTestId
    Debug.Print isTestId
    If Not IsMissing(isAbsPath) Then
      .sAbsPath = isAbsPath
      Debug.Print isAbsPath
    End If
    If Not IsMissing(ilLine) Then
      .lLine = ilLine
      Debug.Print ilLine
    End If
    If Not IsMissing(ilColumn) Then
      .lColumn = ilColumn
      Debug.Print ilColumn
    End If
    If Not IsMissing(ivComment) Then
      .sComment = ivComment
      Debug.Print ivComment
    End If
  End With

  lFailedTestCount = lFailedTestCount + 1
  
  fncInsertFailedTest = True
  If Not bolNoRaise Then objEvent.subFailedTest
ErrH:
End Function
  
Public Function fncRetrieveFailedTestItem(lItem As Long, ByRef tItem As oReportItem) As Boolean
  fncRetrieveFailedTestItem = False
  On Error GoTo ErrH
  
  If (lItem < 0 Or lItem > lFailedTestCount - 1) Then Exit Function
  Set tItem = aFailedTestItems(lItem)
  
  fncRetrieveFailedTestItem = True
ErrH:
End Function

' This function will add entries from another report object to this report object,
' it is also able to add context to the added entries. See the function below
' for help on context adding.
'
Friend Function fncMergeReportsWithContext( _
    iobjFromReport As oReport, Optional isContext As Variant _
    ) As Boolean
  
  bolNoRaise = True
  
  Dim lCounter As Long, objReportItem As oReportItem
  
  For lCounter = 0 To iobjFromReport.lFailedTestCount - 1
    iobjFromReport.fncRetrieveFailedTestItem lCounter, objReportItem
    
    With objReportItem
      If Not IsMissing(isContext) Then .sTestId = isContext & "." & .sTestId
      fncInsertFailedTest .sTestId, .sAbsPath, .lLine, .lColumn, _
        .sComment
    End With
  Next lCounter
  
  bolNoRaise = False
  
  lSucceededTestCount = lSucceededTestCount + iobjFromReport.lSucceededTestCount
End Function

' This function will add a context, a prefix to all report entries. I.E if an
' error is named 'xi.wellFormed' and you add context 'ncc' the new name is
' 'ncc.xi.wellformed'
'
Friend Function fncAddContext(isContext As String) As Boolean
  Dim lCounter As Long
  
  For lCounter = 0 To lFailedTestCount - 1
    With aFailedTestItems(lCounter)
      .sTestId = isContext & "." & .sTestId
    End With
  Next lCounter
End Function

Private Sub Class_Initialize()
  bolNoRaise = False
End Sub

Private Sub Class_Terminate()
  Dim lCounter
  For lCounter = 0 To lFailedTestCount - 1
    Set aFailedTestItems(lCounter) = Nothing
  Next lCounter
  
  ReDim aFailedTestItems(0)
End Sub
