VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInputAuxFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private aInputAuxFiles() As cInputAuxFile
Private lInputAuxFileCount As Long

Private sBrunoCssFullPath As String 'css added by bruno if no css in contentdoc

Private Sub Class_Initialize()
  'debug.print"cInputAuxFiles.initialize"
  lInputAuxFileCount = 0
End Sub

Public Function InputAuxFileCount() As Long
  InputAuxFileCount = lInputAuxFileCount
End Function

Public Function InputAuxFile(lAuxFile As Long) As cInputAuxFile
  Set InputAuxFile = aInputAuxFiles(lAuxFile)
End Function

Public Function fncCheckInputDocumentForAuxFiles( _
    ByVal ilOwnerInputDocument As Long, _
    ByVal sOwnerInputDocumentPath As String, _
    ByRef oInputDom As MSXML2.DOMDocument40 _
    )
Dim oAuxNodes As IXMLDOMNodeList
Dim oAuxNode As IXMLDOMNode
Dim sAuxXPath As String, sCurrSrc As String
Dim bContentDocHasCss As Boolean

  If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
    sAuxXPath = "//xhtml:head/xhtml:link[@rel='stylesheet' and @type='text/css']/@href | //xhtml:img/@src"
  ElseIf oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
    sAuxXPath = "//processing-instruction('xml-stylesheet') | //dtbook:img/@src "
  Else
    Stop
  End If
  Set oAuxNodes = oInputDom.selectNodes(sAuxXPath)

  If Not (oAuxNodes.length = 0) Or (oAuxNodes Is Nothing) Then
    For Each oAuxNode In oAuxNodes
      If oAuxNode.nodeType = NODE_PROCESSING_INSTRUCTION Then
        sCurrSrc = oBruno.oCmn.oUriParser.fncGetAttrValueFromString(oAuxNode.Text, "href")
      Else
        sCurrSrc = oAuxNode.nodeValue
      End If
      
      'sCurrSrc = oBruno.oCmn.oUriParser.fncStripIdAddPath(sCurrSrc, oInputDom.url)
      sCurrSrc = oBruno.oCmn.oUriParser.fncGetPathName(sOwnerInputDocumentPath) & sCurrSrc
      '20050420 add fso.path to normalize string
      sCurrSrc = oBruno.oCmn.oFsoCmn.fncGetFSOPathFromString(sCurrSrc)
      
      If oBruno.oCmn.oFsoCmn.fncFileExists(LCase$(sCurrSrc)) Then
        If Not fncIsObjectInAuxFilesArray(sCurrSrc) Then
          fncAddAuxFile ilOwnerInputDocument, sCurrSrc
        End If
        If LCase$(oBruno.oCmn.oFsoCmn.fncGetExtensionFromString(sCurrSrc)) = "css" Then
         bContentDocHasCss = True
         fncCheckCssForUrlRefs ilOwnerInputDocument, sCurrSrc
        End If
      Else
        ' oNode.nodeValue not found on filesystem
        DoEvents
        frmMain.fncAddMessage oAuxNode.nodeValue & " not found on filesystem"
      End If
    Next
  Else
    'no aux files found
  End If 'Not (oNodes.length = 0) Or (oNodes Is Nothing)


  'now check if brunos built in css should be added to this doc
  'it is added if:
  'no css exists in content doc, and
  'a css exists in css subfolder
  
  'If (Not bContentDocHasCss) Then no linked css was found above
  
  If (Not bContentDocHasCss) Then
    'check that 202 not has inline style
    If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
        If Not oInputDom.selectSingleNode("//xhtml:head/xhtml:style[@type='text/css']") Is Nothing Then
          bContentDocHasCss = True
        End If
    End If
    'no previous css file existed
    If (Not bContentDocHasCss) Then
      'check of there is a css in css folder
      '(sBrunoCssFullPath is set when this class initializes )
      If sBrunoCssFullPath <> "" Then
        If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
          oBruno.oCmn.oDomCmn.fncAppendElement oInputDom.selectSingleNode("//xhtml:head"), "link", , "rel", "stylesheet", "type", "text/css", "href", oBruno.oCmn.oFsoCmn.fncGetFileName(sBrunoCssFullPath)
        Else
          Dim oPI As MSXML2.IXMLDOMProcessingInstruction
          Set oPI = oInputDom.createProcessingInstruction("xml-stylesheet", "type='text/css' href='" & oBruno.oCmn.oFsoCmn.fncGetFileName(sBrunoCssFullPath) & "'")
          oInputDom.insertBefore oPI, oInputDom.documentElement
        End If
        If Not fncIsObjectInAuxFilesArray(sBrunoCssFullPath) Then
          fncAddAuxFile ilOwnerInputDocument, sBrunoCssFullPath
        End If
      End If
    End If '(Not bContentDocHasCss) Then
  End If '(Not bContentDocHasCss) Then
End Function

Private Function fncAddAuxFile( _
 ilOwnerInputDocument As Long, _
 isFullPath As String _
 )
 'mg20041029: check if auxfile already added
 If Not fncIsObjectInAuxFilesArray(isFullPath) Then
   ReDim Preserve aInputAuxFiles(lInputAuxFileCount)
   Set aInputAuxFiles(lInputAuxFileCount) = New cInputAuxFile
   With aInputAuxFiles(lInputAuxFileCount)
    .lOwnerInputDocument = ilOwnerInputDocument
    .lType = TYPE_ACTUAL_AUXILLIARY
    .sFileName = oBruno.oCmn.oFsoCmn.fncGetFileName(isFullPath)
    .sDirectory = oBruno.oCmn.oUriParser.fncGetPathName(isFullPath)
    .sFullPath = isFullPath
    Debug.Print "added auxfile " & .sFullPath
   End With
   lInputAuxFileCount = lInputAuxFileCount + 1
   
 End If
 
End Function

Public Function fncInstantiate()
Dim oFiles As Object, oFile As Object, sSubPath As String
  'check of there is a css in css folder
  If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_D202 Then
    sSubPath = "d202/"
  Else
    sSubPath = "z2004/"
  End If
  Set oFiles = oBruno.oCmn.oFsoCmn.fncGetFolderFiles(oBruno.oPaths.CssPath & sSubPath)
  
  For Each oFile In oFiles
    If oBruno.oCmn.oFsoCmn.fncGetExtensionFromFileObject(oFile) = "css" Then
      sBrunoCssFullPath = oFile.Path
    End If
  Next

End Function

Private Function fncIsObjectInAuxFilesArray( _
  sFullPath As String _
  ) As Boolean
Dim i As Long
    fncIsObjectInAuxFilesArray = False
      
    'mg20041029: check for filename alone, not fullpath
    For i = 0 To lInputAuxFileCount - 1
        If LCase$(oBruno.oCmn.oUriParser.fncGetUriFileName(aInputAuxFiles(i).sFullPath)) = LCase$(oBruno.oCmn.oUriParser.fncGetUriFileName(sFullPath)) Then
            fncIsObjectInAuxFilesArray = True
            Exit For
        End If
    Next i
End Function

Private Function fncCheckCssForUrlRefs(ByRef ilOwnerInputDocument As Long, ByRef sCurrSrc As String) As Boolean
Dim sCssUrlRefs() As String
Dim sCssUrlRefFullPath As String

  'check css for url() entries
  If (fncGetCssUrlValues(sCurrSrc, sCssUrlRefs)) Then
    Dim i As Long
    For i = 1 To UBound(sCssUrlRefs)
       sCssUrlRefFullPath = oBruno.oCmn.oUriParser.fncGetPathName(sCurrSrc) & sCssUrlRefs(i)
      'check integrity and add
      If oBruno.oCmn.oFsoCmn.fncFileExists(LCase$(sCssUrlRefFullPath)) Then
        If Not fncIsObjectInAuxFilesArray(sCssUrlRefFullPath) Then
          fncAddAuxFile ilOwnerInputDocument, sCssUrlRefFullPath
        End If
      Else
        frmMain.fncAddMessage "File referenced by CSS not found on filesystem: " & sCssUrlRefFullPath
      End If '.fncFileExists
    Next i
  Else
    'no url values found, or error
  End If

End Function

Private Function fncGetCssUrlValues(ByRef sCurrSrc As String, ByRef sCssUrlRefs() As String) As Boolean
Dim sCssFile As String
Dim bUrlFound As Boolean
Dim lCurrentUrlPos As Long
Dim lUrlOpenParen As Long
Dim lUrlCloseParen As Long
Dim sUrlString As String
Dim sCssUrlRefsCount As Long
'populate the byref input array with any found url() values
'return false on error or if array length = 0
  fncGetCssUrlValues = False
  On Error GoTo errhandler
  
  sCssFile = oBruno.oCmn.oFsoCmn.fncGetFileAsString(sCurrSrc)
  If sCssFile <> "" Then
    lCurrentUrlPos = 1
    Do
      bUrlFound = False
      lCurrentUrlPos = InStr(lCurrentUrlPos, sCssFile, "url", vbTextCompare)
      If lCurrentUrlPos > 0 Then
        bUrlFound = True
        lUrlOpenParen = InStr(lCurrentUrlPos, sCssFile, "(", vbTextCompare)
        lUrlCloseParen = InStr(lCurrentUrlPos, sCssFile, ")", vbTextCompare)
        'get the string within the paren
        sUrlString = Mid(sCssFile, lUrlOpenParen + 1, lUrlCloseParen - (lUrlOpenParen + 1))
        'Stop
        'remove single or double quotes
        sUrlString = Trim$(sUrlString)
        If (Mid(sUrlString, 1, 1) = Chr(39)) Or (Mid(sUrlString, 1, 1) = Chr(34)) Then
          'remove the leading char
          sUrlString = Mid(sUrlString, 2, Len(sUrlString))
        End If
        
        If (Mid(sUrlString, Len(sUrlString), 1) = Chr(39)) Or (Mid(sUrlString, Len(sUrlString), 1) = Chr(34)) Then
          'remove the trailing char
          sUrlString = Mid(sUrlString, 1, Len(sUrlString) - 1)
        End If
        
        'add it to array
        sCssUrlRefsCount = sCssUrlRefsCount + 1
        ReDim Preserve sCssUrlRefs(sCssUrlRefsCount)
        sCssUrlRefs(sCssUrlRefsCount) = sUrlString
        'indicate there is content on the array
        fncGetCssUrlValues = True
        'move cursor forward
        lCurrentUrlPos = lUrlCloseParen
      End If
    Loop Until bUrlFound = False
  
  End If 'sCssFile <> ""
  
  Exit Function
errhandler:
  fncGetCssUrlValues = False
End Function


