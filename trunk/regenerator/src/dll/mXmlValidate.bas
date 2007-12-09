Attribute VB_Name = "mXmlValidate"
Option Explicit

Public Function fncDtdValidateString(sDomData As String, sLocalName As String)
Dim i As Long
Dim oDom As New MSXML2.DOMDocument40
    oDom.async = False
    oDom.validateOnParse = True
    oDom.resolveExternals = True
    oDom.preserveWhiteSpace = True
    
    If fncParseString(sDomData, oDom) Then objOwner.addlog sLocalName & " valid."
    
    Set oDom = Nothing
    
End Function



