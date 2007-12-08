VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAbstractSmilSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'is used when creating oAbstractSmil
'identifies which nodes in contentdoc belong to this sequence
'there is always at least one oAbstractSequence
'which correlates to the root of abstractsmil
'and therefore identifies all contentdoc nodes

Public oContentDocNodesWithinSequence As IXMLDOMNodeList
Public oAbstractSmilPoint As IXMLDOMNode
Public oContentDocNode As IXMLDOMNode

Public Function initialize(ByRef oInContentDocNode As IXMLDOMNode, ByRef oAbstractSmilSeqNode)
  Set Me.oContentDocNodesWithinSequence = oInContentDocNode.selectNodes(".//*")
  Set Me.oAbstractSmilPoint = oAbstractSmilSeqNode
  'added for update below
  Set Me.oContentDocNode = oInContentDocNode
End Function

Public Function update() As Boolean
  Set Me.oContentDocNodesWithinSequence = oContentDocNode.selectNodes(".//*")
End Function

Private Sub Class_Initialize()
  'debug.Print cAbstractSmilSequence.initialize
End Sub

Private Sub Class_Terminate()
  'debug.Print cAbstractSmilSequence.terminate
  Set oContentDocNodesWithinSequence = Nothing
  Set oAbstractSmilPoint = Nothing
End Sub


