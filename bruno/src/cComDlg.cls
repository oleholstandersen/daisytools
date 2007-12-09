VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
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
 
DefStr S
DefLng N
DefBool B
DefVar V
 
' OFN constants.
Const OFN_ALLOWMULTISELECT   As Long = &H200
Const OFN_CREATEPROMPT       As Long = &H2000
Const OFN_EXPLORER           As Long = &H80000
Const OFN_EXTENSIONDIFFERENT As Long = &H400
Const OFN_FILEMUSTEXIST      As Long = &H1000
Const OFN_HIDEREADONLY       As Long = &H4
Const OFN_LONGNAMES          As Long = &H200000
Const OFN_NOCHANGEDIR        As Long = &H8
Const OFN_NODEREFERENCELINKS As Long = &H100000
Const OFN_OVERWRITEPROMPT    As Long = &H2
Const OFN_PATHMUSTEXIST      As Long = &H800
Const OFN_READONLY           As Long = &H1
 
' The maximum length of a single file path.
Const MAX_PATH As Long = 260
' This MAX_BUFFER value allows you to select approx.
' 500 files with an average length of 25 characters.
' Change this value as needed.
Const MAX_BUFFER As Long = 50 * MAX_PATH
' String constants:
Const sBackSlash As String = "\"
Const sPipe As String = "|"
 
' API functions to use the Windows common dialog boxes.
Private Declare Function GetOpenFileName _
  Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" _
  (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName _
  Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" _
  (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetActiveWindow _
  Lib "user32" () As Long
 
' Type declaration, used by GetOpenFileName and
' GetSaveFileName.
Private Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String  ' Can also be a Long.
End Type
 
' Private variables.
Private OFN As OPENFILENAME
Private colFileTitles As New Collection
Private colFileNames As New Collection
Private sFullName
Private sFileTitle
Private sPath
Private sExtension
 
' Public enumeration variable.
Public Enum XFlags
  PathMustExist = OFN_PATHMUSTEXIST
  FileMustExist = OFN_FILEMUSTEXIST
  PromptToCreateFile = OFN_CREATEPROMPT
End Enum
 
Property Let AllowMultiSelect(bFlag)
  SetFlag OFN_ALLOWMULTISELECT, bFlag
End Property
 
Property Let DialogTitle(sCaption)
  OFN.lpstrTitle = sCaption
End Property
 
Property Let filter(vFilter)
  If IsArray(vFilter) Then _
    vFilter = Join(vFilter, vbNullChar)
  OFN.lpstrFilter = _
    Replace(vFilter, sPipe, vbNullChar) & vbNullChar
End Property
 
Property Get filter()
  With OFN
    If .nFilterIndex Then
      Dim sTemp()
      sTemp = Split(.lpstrFilter, vbNullChar)
      filter = sTemp(.nFilterIndex * 2 - 2) & sPipe & _
        sTemp(.nFilterIndex * 2 - 1)
    End If
  End With
End Property
 
Property Let FilterIndex(nIndex)
  OFN.nFilterIndex = nIndex
End Property
 
Property Get FilterIndex() As Long
  FilterIndex = OFN.nFilterIndex
End Property
 
Property Let RestoreCurDir(bFlag)
  SetFlag OFN_NOCHANGEDIR, bFlag
End Property
 
Property Let ExistFlags(nFlags As XFlags)
  OFN.Flags = OFN.Flags Or nFlags
End Property
 
Property Let CheckBoxVisible(bFlag)
  SetFlag OFN_HIDEREADONLY, Not bFlag
End Property
 
Property Let CheckBoxSelected(bFlag)
  SetFlag OFN_READONLY, bFlag
End Property
 
Property Get CheckBoxSelected() As Boolean
  CheckBoxSelected = OFN.Flags And OFN_READONLY
End Property
 
Property Let FileName(sFileName)
  If Len(sFileName) <= MAX_PATH Then _
    OFN.lpstrFile = sFileName
End Property
 
Property Get FileName() As String
  FileName = sFullName
End Property
 
Property Get FileNames() As Collection
  Set FileNames = colFileNames
End Property
 
Property Get FileTitle() As String
  FileTitle = sFileTitle
End Property
 
Property Get FileTitles() As Collection
  Set FileTitles = colFileTitles
End Property
 
Property Let Directory(sInitDir)
  OFN.lpstrInitialDir = sInitDir
End Property
 
Property Get Directory() As String
  Directory = sPath
End Property
 
Property Let Extension(sDefExt)
  OFN.lpstrDefExt = LCase$(Left$( _
    Replace(sDefExt, ".", vbNullString), 3))
End Property
 
Property Get Extension() As String
  Extension = sExtension
End Property
 
Function ShowOpen() As Boolean
  ShowOpen = Show(True)
End Function
 
Function ShowSave() As Boolean
  ' Set or clear appropriate flags for Save As dialog.
  SetFlag OFN_ALLOWMULTISELECT, False
  SetFlag OFN_PATHMUSTEXIST, True
  SetFlag OFN_OVERWRITEPROMPT, True
  ShowSave = Show(False)
End Function
 
Private Function Show(bOpen)
  With OFN
    .lStructSize = Len(OFN)
    ' Could be zero if no owner is required.
    .hwndOwner = GetActiveWindow
    ' If the RO checkbox must be checked, we should also
    ' display it.
    If .Flags And OFN_READONLY Then _
      SetFlag OFN_HIDEREADONLY, False
    ' Create large buffer if multiple file selection
    ' is allowed.
    .nMaxFile = IIf(.Flags And OFN_ALLOWMULTISELECT, _
      MAX_BUFFER + 1, MAX_PATH + 1)
    .nMaxFileTitle = MAX_PATH + 1
    ' Initialize the buffers.
    .lpstrFile = .lpstrFile & String$( _
      .nMaxFile - 1 - Len(.lpstrFile), 0)
    .lpstrFileTitle = String$(.nMaxFileTitle - 1, 0)
 
    ' Display the appropriate dialog.
    If bOpen Then
      Show = GetOpenFileName(OFN)
    Else
      Show = GetSaveFileName(OFN)
    End If
 
    If Show Then
      ' Remove trailing null characters.
      Dim nDoubleNullPos
      nDoubleNullPos = InStr(.lpstrFile & vbNullChar, _
                              String$(2, 0))
      .lpstrFile = Trim$(.lpstrFile)
      If nDoubleNullPos Then
        ' Get the file name including the path name.
        sFullName = Left$(.lpstrFile, nDoubleNullPos - 1)
        ' Get the file name without the path name.
        sFileTitle = Left$(.lpstrFileTitle, _
          InStr(.lpstrFileTitle, vbNullChar) - 1)
        ' Get the path name.
        sPath = Left$(sFullName, .nFileOffset - 1)
        ' Get the extension.
        If .nFileExtension Then
          sExtension = Mid$(sFullName, .nFileExtension + 1)
        End If
        ' If sFileTitle is a string,
        ' we have a single selection.
        If Len(sFileTitle) Then
          ' Add to the collections.
          colFileTitles.Add _
            Mid$(sFullName, .nFileOffset + 1)
          colFileNames.Add sFullName
        Else  ' Tear multiple selection apart.
          Dim sTemp(), nCount
          sTemp = Split(sFullName, vbNullChar)
          ' If array contains no elements,
          ' UBound returns -1.
          If UBound(sTemp) > LBound(sTemp) Then
            ' We have more than one array element!
            ' Remove backslash if sPath is the root folder.
            If Len(sPath) = 3 Then _
              sPath = Left$(sPath, 2)
            ' Loop through the array, and create the
            ' collections; skip the first element
            ' (containing the path name), so start the
            ' counter at 1, not at 0.
            For nCount = 1 To UBound(sTemp)
              colFileTitles.Add sTemp(nCount)
              ' If the string already contains a backslash,
              ' the user must have selected a shortcut
              ' file, so we don't add the path.
              'colFileNames.Add If(InStr(sTemp(nCount), sBackSlash), sTemp(nCount), sPath & sBackSlash & sTemp(nCount))
              colFileNames.Add sPath & sBackSlash & sTemp(nCount)
            Next
            ' Clear this variable.
            sFullName = vbNullString
          End If
        End If
        ' Add backslash if sPath is the root folder.
        If Len(sPath) = 2 Then _
          sPath = sPath & sBackSlash
      End If
    End If
  End With
End Function
 
Private Sub SetFlag(nValue, bTrue)
  ' Wrapper routine to set or clear bit flags.
  With OFN
    If bTrue Then
      .Flags = .Flags Or nValue
    Else
      .Flags = .Flags And Not nValue
    End If
  End With
End Sub
 
Private Sub Class_Initialize()
  ' This routine runs when the object is created.
  OFN.Flags = OFN.Flags Or OFN_EXPLORER Or _
              OFN_LONGNAMES Or OFN_HIDEREADONLY
End Sub


