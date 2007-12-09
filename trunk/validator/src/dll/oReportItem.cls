VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oReportItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

' Items in oReport failed tests array uses these

Private sPrivAbsPath As String, _
        lPrivLine As Long, _
        lPrivColumn As Long, _
        sPrivComment As String, _
        sPrivTestId As String

        'sAbsPath as string  = full path inc filename
        'lLine As Long = line in doc
        'lColumn As Long, = column in doc
        'sTestID As String = test id (used for retrieval in validator.xml)
        'sComment As String = used for comments such as suggestions for change

Private objVtmItem As oVtmItem

' This property sets the report items ID
Friend Property Let sTestId(isTestId As String)
  sPrivTestId = isTestId
  
  ' If the VTM file isn't parsed, parse it
  If Not bolVtmIsLoaded Then fncParseVtm
  ' Find the VTM with the same context and set "objVtmItem" to point at it
  Set objVtmItem = fncGetVtmItem(isTestId)
End Property

' This property gets the report items ID
Public Property Get sTestId() As String
  sTestId = sPrivTestId
End Property

' This property sets the report items AbsPath property
Friend Property Let sAbsPath(isAbsPath As String)
  sPrivAbsPath = isAbsPath
End Property

' This property gets the report items AbsPath property
Public Property Get sAbsPath() As String
  sAbsPath = sPrivAbsPath
End Property

' This property sets the report items Line property
Friend Property Let lLine(ilLine As String)
  lPrivLine = ilLine
End Property

' This property gets the report items Line property
Public Property Get lLine() As String
  lLine = lPrivLine
End Property

' This property sets the report items Column property
Friend Property Let lColumn(ilColumn As String)
  lPrivColumn = ilColumn
End Property

' This property gets the report items Column property
Public Property Get lColumn() As String
  lColumn = lPrivColumn
End Property

' This property sets the report items Comment property
Friend Property Let sComment(isComment As String)
  sPrivComment = isComment
End Property

' This property gets the report items Comment property
Public Property Get sComment() As String
  sComment = sPrivComment
End Property

' This property gets the report items FailClass property from the VTM item that
' "objVtmItem" is pointing at
'
Public Property Get sFailClass() As String
  If objVtmItem Is Nothing Then Exit Property
  sFailClass = objVtmItem.sFailClass
End Property

' This property gets the report items FailType property from the VTM item that
' "objVtmItem" is pointing at
'
Public Property Get sFailType() As String
  If objVtmItem Is Nothing Then Exit Property
  sFailType = objVtmItem.sFailType
End Property

' This property gets the report items Link property from the VTM item that
' "objVtmItem" is pointing at
'
Public Property Get sLink() As String
  If objVtmItem Is Nothing Then Exit Property
  sLink = objVtmItem.sLink
End Property

' This property gets the report items LongDesc property from the VTM item that
' "objVtmItem" is pointing at
'
Public Property Get sLongDesc() As String
  If objVtmItem Is Nothing Then Exit Property
  sLongDesc = objVtmItem.sLongDesc
End Property

' This property gets the report items ShortDesc property from the VTM item that
' "objVtmItem" is pointing at
'
Public Property Get sShortDesc() As String
  If objVtmItem Is Nothing Then Exit Property
  sShortDesc = objVtmItem.sShortDesc
End Property

' This property gets the report items Name property from the VTM item that
' "objVtmItem" is pointing at
'
Public Property Get sName() As String
  If objVtmItem Is Nothing Then Exit Property
  sName = objVtmItem.sName
End Property

Private Sub Class_Terminate()
  Set objVtmItem = Nothing
End Sub
