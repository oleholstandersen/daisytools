VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lSpineItemCount As Long
Private oSpineDom As New MSXML2.DOMDocument40
Private oSpineItemIdGetter As cIdGetter

Private Sub Class_Initialize()
  Set oSpineItemIdGetter = New cIdGetter
  oSpineDom.loadXML "<spine></spine>"
  lSpineItemCount = 0
End Sub

Public Function fncAppendSpineItem(sSmilFileName As String) As IXMLDOMNode
Dim oNewNode As IXMLDOMNode
Dim oNewAttr As IXMLDOMAttribute
  Set oNewNode = oSpineDom.createNode(NODE_ELEMENT, "item", "")
  Set oNewAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oNewNode, "name", sSmilFileName)
  Set oNewAttr = oBruno.oCmn.oDomCmn.fncAppendAttribute(oNewNode, "id", "s" & oSpineItemIdGetter.fncGetId)
  Set oNewNode = oSpineDom.documentElement.appendChild(oNewNode)
  lSpineItemCount = lSpineItemCount + 1
  Set fncAppendSpineItem = oNewNode
End Function

Public Function fncGetSpineItem(lSpineItem As Long) As IXMLDOMElement
Dim oNode As IXMLDOMNode
Dim k As Long

  For Each oNode In oSpineDom.documentElement.childNodes
    If k = lSpineItem Then
      Set fncGetSpineItem = oNode
      Exit For
    End If
    k = k + 1
  Next

End Function



