Attribute VB_Name = "mXsl"
Option Explicit

Public Function fncRunXslTransform( _
    ByRef oDocToTransForm As MSXML2.DOMDocument40, _
    ByVal sXsltPath As String _
) As Boolean

  'fncRunXslTransform = True: Exit Function

Dim oXsl As New MSXML2.DOMDocument40: oXsl.async = False
'
'sXsltPath = sAppPath & "xml_pretty_printer\xsl\xml_pp_clean.xsl"


  'If bDebugMode Then objowner.addlog "fncRunXslTransform in"
  fncRunXslTransform = False
  If fncParseFile(sXsltPath, oXsl) Then
    ' Parse results into byref DOM Document.
    Dim oDomOutput As New MSXML2.DOMDocument40
        oDomOutput.async = False
        oDomOutput.validateOnParse = False
        oDomOutput.resolveExternals = False
        oDomOutput.preserveWhiteSpace = True
        
    oDocToTransForm.transformNodeToObject oXsl, oDomOutput
    Set oDocToTransForm = oDomOutput
  Else
    objOwner.addlog fncGetFileName(sXsltPath) & " parse error"
    GoTo ErrHandler
  End If
  
  fncRunXslTransform = True
  'If bDebugMode Then objowner.addlog "fncRunXslTransform out"
ErrHandler:
  If Not fncRunXslTransform Then objOwner.addlog "fncRunXslTransform ErrH"
End Function
