Attribute VB_Name = "mXmlLoad"
Option Explicit

Public Function fncParseFile( _
    ByVal isAbsPath As String, _
    ByRef ioDom As MSXML2.DOMDocument40 _
    ) As Boolean
 
    fncParseFile = False
    
    If Not ioDom.Load(isAbsPath) Then
        addLog "Parse error in " & fncGetFileName(isAbsPath) & ": " & ioDom.parseError.reason & vbCrLf & _
           "filepos: " & ioDom.parseError.filepos & vbCrLf & _
           "line: " & ioDom.parseError.Line & vbCrLf & _
           "linepos: " & ioDom.parseError.linepos & vbCrLf & _
           "srctext: " & ioDom.parseError.srcText & vbCrLf & _
           "Url: " & ioDom.parseError.url & vbCrLf
    Else
        fncParseFile = True
    End If
    
End Function

Public Function fncParseString( _
    ByVal isContent As String, _
    ByRef ioDom As MSXML2.DOMDocument40 _
    ) As Boolean
    
    fncParseString = False

    If Not ioDom.loadXML(isContent) Then
        addLog "Parse error: " & ioDom.parseError.reason & vbCrLf & _
           "filepos: " & ioDom.parseError.filepos & vbCrLf & _
           "line: " & ioDom.parseError.Line & vbCrLf & _
           "linepos: " & ioDom.parseError.linepos & vbCrLf & _
           "srctext: " & ioDom.parseError.srcText & vbCrLf & _
           "Url: " & ioDom.parseError.url '& vbCrLf & vbCrLf & _
           'isContent
    Else
        fncParseString = True
    End If
End Function
