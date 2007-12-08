VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutputSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
  'debug.Print "cOutputSaver.Initialize"
End Sub

Private Sub Class_Terminate()
  'debug.Print "cOutputSaver.Terminate"
End Sub

Public Function fncRenderOutputDocuments() As Boolean
Dim i As Long
  
  On Error GoTo errh
  fncRenderOutputDocuments = False
    
  'create the output path
  If Not oBruno.oCmn.oFsoCmn.fncCreateDirectoryChain(oBruno.oPaths.OutputPath) Then Exit Function
    
  For i = 0 To oBruno.oOutputDocuments.OutputDocumentCount - 1
    DoEvents
    With oBruno.oOutputDocuments.Document(i)
    If (.lType <> TYPE_ACTUAL_LPP) Then
        If (.lType <> TYPE_ACTUAL_MDF) Then
          If (.lType <> TYPE_ACTUAL_AUXILLIARY) Then
            If Not .oDom Is Nothing Then
              'we have xml documents
              'remove redundant namespace attributes
              If .lType <> TYPE_ACTUAL_202SMIL And .lType <> TYPE_ACTUAL_202MSMIL Then
                'unnecessary to do it for 202smil; no namespace at all
                If Not oBruno.oCmn.oDomCmn.fncRedundantNsRemove(.oDom, True) Then GoTo errh
              End If
              'pretty print
              If Not oBruno.oCmn.oDomCmn.fncXmlIndentDom(.oDom.documentElement, vbTab) Then GoTo errh
              'set output encoding using first input document encoding
              If Not oBruno.oCmn.oDomCmn.fncSetEncoding(.oDom, oBruno.oInputDocuments.InputDocument(0).sEncoding) Then GoTo errh
              'save
              .oDom.save (oBruno.oPaths.OutputPath & .sFileName)
              
              If .lType = TYPE_ACTUAL_202CONTENT Or .lType = TYPE_ACTUAL_Z39CONTENT Then
                'check if this doc had auxfiles
                'note - this renders only auxfiles that have a contentdocument owner
                fncRenderAuxFiles i
              End If
            Else
              'this is an uncatered for nonxml type
              Stop
            End If
          Else
            'we have an unkown type
            Stop
          End If
        Else
            'we have an mdf file
            If Not oBruno.oCmn.oFsoCmn.fncCopyFile(oBruno.oPaths.ShellPath & .sFileName, oBruno.oPaths.OutputPath & .sFileName, True) Then
              'couldnt copy mdf
              frmMain.fncAddMessage "could not copy dummy MDF"
            End If
        End If
    Else
      'we have an lpp file
      
      'If Not oBruno.oCmn.oFsoCmn.fncMoveFile(oBruno.oPaths.ShellPath & "temp.lpp", oBruno.oPaths.OutputPath & oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension(oBruno.oInputDocuments.InputDocument(0).sFileName) & ".lpp", True) Then
      If Not oBruno.oCmn.oFsoCmn.fncMoveFile( _
        oBruno.oPaths.ShellPath & "temp.lpp", _
        oBruno.oPaths.OutputPath & _
        oBruno.oCmn.oFsoCmn.fncGetFileNameLessExtension( _
        oBruno.oAbstractDocuments.AbstractDocument(0).sFileName) _
        & ".lpp", True) Then
          frmMain.fncAddMessage "could not copy lpp"
      End If
    End If
    End With
  Next
  
  'finally, copy possible loose aux files (not .hasBeenCopied) into destination dir
  Dim r As Long

  For r = 0 To oBruno.oInputDocuments.oInputAuxFiles.InputAuxFileCount - 1
    With oBruno.oInputDocuments.oInputAuxFiles.InputAuxFile(r)
      If Not .bHasBeenCopiedToOutput Then
        If oBruno.oCmn.oFsoCmn.fncCopyFile(.sFullPath, oBruno.oPaths.OutputPath & .sFileName, True) Then
            'frmMain.fncAddMessage "Auxilliary file " & .sFullPath & " copied to " & oBruno.oPaths.OutputPath & " which is probably wrong destination"
        Else
            'could not copy aux
        End If
      End If
    End With
  Next
  
  fncRenderOutputDocuments = True
errh:

End Function

Private Function fncRenderAuxFiles(lOutputContentDoc As Long) As Boolean
Dim r As Long
Dim lActiveInputDocument As Long
Dim sInputDocDir As String

  'renders auxfiles beloning to inparam contentdoc
  fncRenderAuxFiles = True
  
  lActiveInputDocument = oBruno.oOutputDocuments.Document(lOutputContentDoc).lInputDocumentSource
  
  For r = 0 To oBruno.oInputDocuments.oInputAuxFiles.InputAuxFileCount - 1
    With oBruno.oInputDocuments.oInputAuxFiles.InputAuxFile(r)
      If .lOwnerInputDocument = lActiveInputDocument Then
        'this auxfile belongs to inparam contentdoc
        'check path of contentdoc vs auxfile
        sInputDocDir = oBruno.oInputDocuments.InputDocument(lActiveInputDocument).sDirectory
        If .sDirectory = sInputDocDir Then
          'if same path, copy to outputdir
          If oBruno.oCmn.oFsoCmn.fncCopyFile(.sFullPath, oBruno.oPaths.OutputPath & .sFileName, True) Then
            .bHasBeenCopiedToOutput = True
          End If
        Else '.sDirectory <> sInputDocDir
          'if subdir
          If oBruno.oCmn.oFsoCmn.fncIsFileDescendantOfFolder(.sFullPath, sInputDocDir) Then
            Dim sOutputPathSubDir As String
            Dim sSubDir As String
            'find the subdir path to create
            sSubDir = oBruno.oCmn.oFsoCmn.fncGetParentFolderName(.sFullPath)
            sOutputPathSubDir = oBruno.oPaths.OutputPath & sSubDir & "\"
            'create it
            If Not oBruno.oCmn.oFsoCmn.fncFolderExists(sOutputPathSubDir) Then
              oBruno.oCmn.oFsoCmn.fncCreateDirectoryChain (sOutputPathSubDir)
            End If
            'copy auxfile to created subdir
            If oBruno.oCmn.oFsoCmn.fncCopyFile(.sFullPath, sOutputPathSubDir & .sFileName, True) Then
              .bHasBeenCopiedToOutput = True
            End If
          Else
            'external dir (src="../../etc), or other error in fncIsFileDescendantOfFolder
          End If 'fncIsFileDescendantOfFolder
        End If '.sDirectory = sInputDocDir Then
      Else
        'this auxfile does not belong to inparam contentdoc
      End If
    End With
  Next r
        
  fncRenderAuxFiles = False
End Function
