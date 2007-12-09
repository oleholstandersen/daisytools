VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oEvent"
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

Public Event evErrorLog(sError As String)
Public Event evFailedTest()
Public Event evSucceededTest()
Public Event evProgressChanged()

Public objOwner As Object

' This event is fired each time an INTERNAL ERROR occurs
Friend Sub subLog(isError As String)
    RaiseEvent evErrorLog(isError)
    subRaiseVBFriendlyEvent "ValidatorEngine.Log", isError, ""
    Debug.Print isError
End Sub

' This event is fired each time a test fails
Friend Sub subFailedTest()
    RaiseEvent evFailedTest
    subRaiseVBFriendlyEvent "ValidatorEngine.FailedTest", "", ""
End Sub

' This event is fired each time a test passes
Friend Sub subSucceededTest()
    RaiseEvent evSucceededTest
    subRaiseVBFriendlyEvent "ValidatorEngine.SucceededTest", "", ""
End Sub

' This event is fired each time the current progress changes
Friend Sub subProgressChanged()
  RaiseEvent evProgressChanged
  subRaiseVBFriendlyEvent "ValidatorEngine.ProgressChanged", "", ""
End Sub

' This function is necessary if using late bound in Visual Basic. To use this
' function, set the 'objOwner' object to point to a object that contains a
' function with the name 'subVBFriendlyEvent' and the input parameters
' 'isEvent as string, vParam1 as variant, vParam2 as variant'.
' This function will receive all event notifications sent by the validator engine.
Private Sub subRaiseVBFriendlyEvent(isEvent As String, vParam1 As Variant, _
  vParam2 As Variant)
  
  If objOwner Is Nothing Then Exit Sub
  objOwner.subVBFriendlyEvent isEvent, vParam1, vParam2
End Sub

