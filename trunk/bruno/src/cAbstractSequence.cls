VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAbstractSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'is used when creating oAbstractSmil
'identifies which nodes in contentdoc belong to this sequence
'there is alwayws at least one oAbstractSequnce
'which correlates to the root of abstractsmil
'and therefore identifies all contentdoc nodes

Public oContentDocNodes As IXMLDOMNodeList
Public oAbstractSmilPoint As IXMLDOMNode

Private Sub Class_Initialize(ByRef oContentDocNode As IXMLDOMNode)
  Set oContentDocNodes = oContentDocNode.selectNodes(".//*")
End Sub
