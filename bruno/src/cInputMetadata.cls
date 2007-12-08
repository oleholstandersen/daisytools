VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInputMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public sDcIdentifier As String
Public sDcTitle As String
Public sXhtmlTitle As String
Public sDcCreator As String
Public sDcPublisher As String
Public sDcLanguage As String
Public sDcDate As String
Public sDcSubject As String

'20050309 added these for niels ibos
Public sDcSource
Public sNccMultimediaType
Public sNccNarrator
Public sNccSourceDate
Public sNccSourceEdition
Public sNccSourcePublisher

Public sDtbookDocAuthor As String
Public sDtbookDocTitle As String
Public sDtbookUid As String

Public sXmlLang As String

Public Function SetCommonMetas(oDom As MSXML2.DOMDocument40) As Boolean
Dim oMetaElement As IXMLDOMNode
Dim oMetaElements As IXMLDOMNodeList
'outfileset (d202|z3986.2005|raw) #REQUIRED

  If Not oDom.documentElement.Attributes Is Nothing Then
    Set oMetaElement = oDom.documentElement.Attributes.getNamedItem("xml:lang")
    If Not oMetaElement Is Nothing Then
      sXmlLang = oMetaElement.Text
    End If
  End If
  
  If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='dc:identifier']/@content")
      If Not oMetaElement Is Nothing Then
        sDcIdentifier = oMetaElement.Text
      Else
        sDcIdentifier = "unknown"
      End If
        
      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='dc:creator']/@content")
      If Not oMetaElement Is Nothing Then
        sDcCreator = oMetaElement.Text
      End If
    
      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='dc:publisher']/@content")
      If Not oMetaElement Is Nothing Then
        sDcPublisher = oMetaElement.Text
      End If
    
      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='dc:language']/@content")
      If Not oMetaElement Is Nothing Then
        sDcLanguage = oMetaElement.Text
      End If
    
      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='dc:date']/@content")
      If Not oMetaElement Is Nothing Then
        sDcDate = oMetaElement.Text
      End If
    
      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='dc:subject']/@content")
      If Not oMetaElement Is Nothing Then
        sDcSubject = oMetaElement.Text
      End If
      
      '20050309 added these for niels ibos
      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='dc:source']/@content")
      If Not oMetaElement Is Nothing Then
        sDcSource = oMetaElement.Text
      End If

      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='ncc:multimediaType']/@content")
      If Not oMetaElement Is Nothing Then
        sNccMultimediaType = oMetaElement.Text
      End If

      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='ncc:narrator']/@content")
      If Not oMetaElement Is Nothing Then
        sNccNarrator = oMetaElement.Text
      End If

      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='ncc:sourceDate']/@content")
      If Not oMetaElement Is Nothing Then
        sNccSourceDate = oMetaElement.Text
      End If

      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='ncc:sourceEdition']/@content")
      If Not oMetaElement Is Nothing Then
        sNccSourceEdition = oMetaElement.Text
      End If

      Set oMetaElement = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='ncc:sourcePublisher']/@content")
      If Not oMetaElement Is Nothing Then
        sNccSourcePublisher = oMetaElement.Text
      End If
      
      ' end niels adds
      
      
      'try to extract a title using xhtml:title and dc:title
      Dim oDcTitle As IXMLDOMNode, oXhtmlTitle As IXMLDOMNode
      Dim sTempDcTitle As String, sTempXhtmlTitle As String
      
      Set oDcTitle = oDom.selectSingleNode("//xhtml:head/xhtml:meta[@name='dc:title']/@content")
      Set oXhtmlTitle = oDom.selectSingleNode("//xhtml:head/xhtml:title")
      
      sTempDcTitle = ""
      sTempXhtmlTitle = ""
      
      If Not oDcTitle Is Nothing Then sTempDcTitle = Trim$(oDcTitle.Text)
      If Not oXhtmlTitle Is Nothing Then sTempXhtmlTitle = Trim$(oXhtmlTitle.Text)
                        
      If (Len(sTempDcTitle) > 0) And (Len(sTempXhtmlTitle) > 0) Then
        'Debug.Print "both titles have values"
        sXhtmlTitle = sTempXhtmlTitle
        sDcTitle = sTempDcTitle
        
      ElseIf (Len(sTempDcTitle) < 1) And (Len(sTempXhtmlTitle) > 0) Then
        'Debug.Print "only xhtml:title was given"
        sXhtmlTitle = sTempXhtmlTitle
        sDcTitle = sTempXhtmlTitle
        
      ElseIf (Len(sTempDcTitle) > 0) And (Len(sTempXhtmlTitle) < 1) Then
        'Debug.Print "only dc:title was given"
        sDcTitle = sTempDcTitle
        sXhtmlTitle = sTempDcTitle
      Else
        'Debug.Print "both are empty"
        sDcTitle = "unknown title"
        sXhtmlTitle = "unknown title"
      End If
      'Stop
  ElseIf oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
      Set oMetaElement = oDom.selectSingleNode("//dtbook:head/dtbook:meta[@name='dtb:uid']/@content")
      If Not oMetaElement Is Nothing Then
        sDtbookUid = oMetaElement.Text
      Else
        sDtbookUid = "unknown"
      End If
      
      Set oMetaElement = oDom.selectSingleNode("//dtbook:frontmatter/dtbook:doctitle")
      If Not oMetaElement Is Nothing Then
        sDtbookDocTitle = oMetaElement.Text
      Else
        'go for dublin core as fallback
        Set oMetaElement = oDom.selectSingleNode("//dtbook:meta[@name='dc:Title']/@content")
        If Not oMetaElement Is Nothing Then
            sDtbookDocTitle = oMetaElement.Text
        Else
            sDtbookDocTitle = "title unspecified"
        End If
      End If
      
      Set oMetaElements = oDom.selectNodes("//dtbook:frontmatter/dtbook:docauthor")
      If (Not oMetaElements Is Nothing) Then
        If (oMetaElements.length > 0) Then
          If oMetaElements.length = 1 Then
            sDtbookDocAuthor = oMetaElements.Item(0).Text
          Else
            For Each oMetaElement In oMetaElements
              sDtbookDocAuthor = sDtbookDocAuthor & oMetaElement.Text & ", "
            Next
            sDtbookDocAuthor = Mid(sDtbookDocAuthor, 1, Len(sDtbookDocAuthor) - 2)
          End If
        End If
      End If
      If sDtbookDocAuthor = "" Then
        'go for dublin core as fallback
        Set oMetaElement = oDom.selectSingleNode("//dtbook:meta[@name='dc:Creator']/@content")
        If Not oMetaElement Is Nothing Then
            sDtbookDocAuthor = oMetaElement.Text
        Else
           sDtbookDocAuthor = "author unspecified"
        End If
      End If
      
      
      
      Set oMetaElement = oDom.selectSingleNode("//dtbook:head/dtbook:meta[@name='dc:Title']/@content")
      If Not oMetaElement Is Nothing Then
        sDcTitle = oMetaElement.Text
      Else
        sDcTitle = "title unknown"
      End If
      
      Set oMetaElement = oDom.selectSingleNode("//dtbook:head/dtbook:meta[@name='dc:Creator']/@content")
      If Not oMetaElement Is Nothing Then
        sDcCreator = oMetaElement.Text
      Else
        sDcCreator = "unknown"
      End If
              
      Set oMetaElement = oDom.selectSingleNode("//dtbook:head/dtbook:meta[@name='dc:Date']/@content")
      If Not oMetaElement Is Nothing Then
        sDcDate = oMetaElement.Text
      Else
        sDcDate = "0000-00-00"
      End If
              
      Set oMetaElement = oDom.selectSingleNode("//dtbook:head/dtbook:meta[@name='dc:Publisher']/@content")
      If Not oMetaElement Is Nothing Then
        sDcPublisher = oMetaElement.Text
      Else
        sDcPublisher = "unknown"
      End If
              
      Set oMetaElement = oDom.selectSingleNode("//dtbook:head/dtbook:meta[@name='dc:Language']/@content")
      If Not oMetaElement Is Nothing Then
        sDcLanguage = oMetaElement.Text
      Else
        sDcLanguage = "xx"
      End If
              
              
  Else
    'raw, no metas
  End If
  
End Function

