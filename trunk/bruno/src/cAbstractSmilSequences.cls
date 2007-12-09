VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAbstractSmilSequences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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

Public initialized As Boolean
Private aSequences() As cAbstractSmilSequence
Public lSequencesCount As Long

Public Function initialize(ByRef oContentDocRoot As IXMLDOMNode, oAbstractSmil As cAbstractDocument)
  Me.AddSequence oContentDocRoot, oAbstractSmil.oDom.documentElement
  'oabstractsmil root elem is now first item in asequences array
  initialized = True
End Function

Public Function Sequence(lArrayItem) As cAbstractSmilSequence
   Set Sequence = aSequences(lArrayItem)
End Function

Public Function AddSequence( _
    ByRef oContentDocNode As IXMLDOMNode, _
    ByRef oSeqElement As IXMLDOMNode) _
    As Boolean
  ReDim Preserve aSequences(lSequencesCount)
  Set aSequences(lSequencesCount) = New cAbstractSmilSequence
  aSequences(lSequencesCount).initialize oContentDocNode, oSeqElement
  lSequencesCount = lSequencesCount + 1
  
End Function

Public Function GetPointForTimeContainer( _
    oContentDocNode As IXMLDOMNode _
    ) As IXMLDOMNode
'gets a contentdocnode as input
'returns the cAbstractSmilSequence.oAbstractSmilPoint
'that includes oContentDocNode in cAbstractSmilSequence.oContentDocNodesWithinSequence
'iterate backwards through aSequences() array
'therefore always ends at rootSequence
'Note: ideally, aSequences should not be a flat array but a branched tree structure
Dim i As Long

  'add for optim:
  If lSequencesCount = 1 Then
    Set GetPointForTimeContainer = aSequences(0).oAbstractSmilPoint
    Exit Function
  End If

  For i = (lSequencesCount - 1) To 0 Step -1
    If oBruno.oCmn.oDomCmn.fncIsInNodeList(oContentDocNode, aSequences(i).oContentDocNodesWithinSequence) Then
      Set GetPointForTimeContainer = aSequences(i).oAbstractSmilPoint
      Exit Function
    End If
  Next
  
End Function

Private Sub Class_Initialize()
  'Debug.Print cAbstractSmilSequences.initialize
End Sub

Private Sub Class_Terminate()
  'Debug.Print cAbstractSmilSequences.terminate
  lSequencesCount = 0
  ReDim Preserve aSequences(lSequencesCount)
End Sub

