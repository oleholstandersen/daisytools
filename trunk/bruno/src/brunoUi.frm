VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3345
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10320
   ForeColor       =   &H00FFFFFF&
   Icon            =   "brunoUi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmUiUpdate 
      Interval        =   1000
      Left            =   9720
      Top             =   -120
   End
   Begin VB.Frame Frame2 
      Caption         =   "infobar"
      Height          =   1215
      Left            =   5400
      TabIndex        =   16
      Top             =   1920
      Width           =   4695
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00000000&
         ForeColor       =   &H80000009&
         Height          =   795
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "infobar"
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   5400
      TabIndex        =   13
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton cmdSaveOutput 
         Caption         =   "create fileset"
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "create and save output files"
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtOutputPath 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Text            =   "D:\output\"
         ToolTipText     =   "Path to save output files to"
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdOutputPath 
         Caption         =   "browse"
         Height          =   315
         Left            =   3600
         TabIndex        =   9
         ToolTipText     =   "browse and select/create output path"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "output path"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "remove"
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "delete document from list"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdDocDown 
      Caption         =   "down"
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      ToolTipText     =   "move document down in list"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdDocUp 
      Caption         =   "up"
      Height          =   315
      Left            =   3240
      TabIndex        =   6
      ToolTipText     =   "move document up in list"
      Top             =   2640
      Width           =   855
   End
   Begin MSComctlLib.ListView oInputDocList 
      Height          =   1815
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Input Document List"
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame frameInput 
      Caption         =   "input"
      Height          =   3255
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdRevalidate 
         Caption         =   "revalidate"
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cmbDriverList 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "select driver to apply to input document"
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmdAddDoc 
         Caption         =   "add"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "add input document"
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "driver"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame frameGenerate 
      Caption         =   "output"
      Height          =   3255
      Left            =   5280
      TabIndex        =   1
      Top             =   0
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog oCommonDialog 
      Left            =   4800
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAddSourceDoc 
         Caption         =   "Add Source Document(s)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSetOutPath 
         Caption         =   "&Set output path"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuSaveOutput 
         Caption         =   "&Save Output Files"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuInputEdit 
      Caption         =   "Input Doc &Edit"
      Begin VB.Menu mnuRemoveSelected 
         Caption         =   "Remove Selected Doc"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuRevalidateAll 
         Caption         =   "Revalidate All Docs"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuMoveDocUp 
         Caption         =   "Move Selected Doc Up"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuMoveDocDown 
         Caption         =   "Move Selected Doc Down"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuFocus 
      Caption         =   "Focus"
      Begin VB.Menu mnuDriverSelect_setFocus 
         Caption         =   "Driver selector"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuInputDocList_setFocus 
         Caption         =   "Added documents"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuOutputPath_setFocus 
         Caption         =   "Output Path"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuInfoBar_setFocus 
         Caption         =   "Infobar"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuUserDocumenation 
         Caption         =   "User Documentation"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAdvDoc 
         Caption         =   "Driver Documentation"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Dim lCaptionTimerCounter As Long

Private Sub cmbDriverList_Click()
 If frmMain.Visible Then
  txtStatus.Text = _
   oDriverList.Driver(cmbDriverList.ListIndex).sDesc
 End If
End Sub

Private Sub mnuAbout_Click()
 frmAbout.Show
End Sub

Private Sub mnuRevalidateAll_Click()
  cmdRevalidate_Click
End Sub

Private Sub cmdRevalidate_Click()
  'revalidate all added documents
  If oBruno.oInputDocuments.InputDocumentCount > 0 Then
    Me.txtStatus = ""
    oBruno.oInputDocuments.fncReValidateInputDocuments
    If oBruno.oInputDocuments.fncCheckDocuments Then
      fncAddMessage ("no problems in doc(s). congrats.")
    End If
  End If
  fncPopulateInputList
End Sub

Private Sub Form_Load()
  oInputDocList.ColumnHeaders.Add , , "Name", ((oInputDocList.Width / 100) * 50)
  oInputDocList.ColumnHeaders.Add , , "Type", ((oInputDocList.Width / 100) * 24)
  oInputDocList.ColumnHeaders.Add , , "Status", ((oInputDocList.Width / 100) * 24)
  subPopulateDriverList
  
  'get prev output path from registry
  Dim sValue As String
  oRegistryControl.fncLoadRegistryData "OutputPath", sValue, &H80000001, , "d:/bruno_output/"
  Me.txtOutputPath.Text = sValue
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

 oRegistryControl.fncSaveRegistryData "OutputPath", Trim$(Me.txtOutputPath.Text), &H80000001
 oRegistryControl.fncSaveRegistryData "DriverName", Me.cmbDriverList.Text, &H80000001
 
 If Not oBruno.fncTerminateChildren Then
   Stop
 End If
 Set oBruno = Nothing
End Sub

Private Sub mnuAddSourceDoc_Click()
Dim oCdlg As cComDlg
Dim i As Long
Dim k As Long
Dim bdupeadded As Boolean

  If (oBruno.oInputDocuments.InputDocumentCount > 0) Then
   If oBruno.oDriver.lOutFileSet = OUTPUT_TYPE_Z39 Then
     fncAddMessage "You can only add one dtbook document per session. Muliple docs is allowed only under d202."
     Exit Sub
   End If
  End If
  fncAddMessage ""
  Set oCdlg = New cComDlg
  With oCdlg
    .DialogTitle = "Select Input Source Documents"
    Dim f(3) As String
    f(0) = "Xml Documents (*.xml)|*.xml"
    f(1) = "Xhtml Documents (*.html)|*.html"
    f(2) = "All Files (*.*)|*.*"
    .filter = f
    .FilterIndex = 3
    .CheckBoxSelected = True
    .AllowMultiSelect = True
    .ExistFlags = FileMustExist + PathMustExist
    If .ShowOpen Then
      If .FileNames.Count > 0 Then
        For i = 1 To .FileNames.Count
          bdupeadded = False
          For k = 0 To oBruno.oInputDocuments.InputDocumentCount - 1
            If LCase$(oBruno.oInputDocuments.InputDocument(k).sFullPath) = LCase$(.FileNames(i)) Then
              bdupeadded = True
              fncAddMessage "cant add same file more than once"
            End If
          Next
          If Not bdupeadded Then
            If (oBruno.oDriver Is Nothing) Then
              'Driver was not loaded
              If Not oBruno.fncCreateDriver(oDriverList.Driver(cmbDriverList.ListIndex).sFullPath) Then
                'fncAddMessage "Driver " & oDriverList.Driver(cmbDriverList.Index).sName & " could not be loaded"
                'fncAddMessage "Driver could not be loaded"
                Exit Sub
              End If
            End If
            If (oBruno.oDriver.sFullPath <> oDriverList.Driver(cmbDriverList.ListIndex).sFullPath) Then
              'user selected another Driver in the ui than the one loaded
              'clear all added docs
              oBruno.oInputDocuments.fncResetArrays
              'create new Driver
              oBruno.fncCreateDriver (oDriverList.Driver(cmbDriverList.ListIndex).sFullPath)
            End If
            'add input doc
            If Not oBruno.fncAddInputDocument(.FileNames(i)) Then
              Stop
            End If
          End If
        Next
        fncPopulateInputList
      End If
    End If
  End With
  Set oCdlg = Nothing
End Sub

Private Sub cmdAddDoc_Click()
  mnuAddSourceDoc_Click
End Sub

Private Sub mnuExit_Click()
 fncResetAll
 Unload Me
End Sub

Private Function fncResetAll() As Boolean
  fncResetAll = False
  Set oBruno.oInputDocuments = Nothing
  fncResetAll = True
End Function

Public Function fncPopulateInputList() As Boolean
Dim i As Long
Dim oListItem As ListItem

 oInputDocList.ListItems.Clear
 
 With oBruno.oInputDocuments
  For i = 0 To .InputDocumentCount - 1
    Set oListItem = oInputDocList.ListItems.Add(, , .InputDocument(i).sFileName)
    oListItem.ListSubItems.Add , , oBruno.oInputDocuments.InputDocument(i).sDocTypeNiceName
    If .InputDocument(i).bWellformed Then
      If .InputDocument(i).bValid Then
          oListItem.ListSubItems.Add , , "valid"
      Else
        oListItem.ListSubItems.Add , , "invalid"
      End If
    Else
      oListItem.ListSubItems.Add , , "malformed"
    End If
  Next
 End With
 
 oInputDocList.SetFocus
 If (oInputDocList.ListItems.Count > 0) Then
   oInputDocList.ListItems(oInputDocList.ListItems.Count).Selected = True
 End If
 
End Function

Private Sub mnuMoveDocDown_Click()
Dim lReturnIndex As Long
  If oInputDocList.ListItems.Count > 0 Then
    lReturnIndex = oBruno.oInputDocuments.fncMoveDocument(oInputDocList.SelectedItem.Index - 1, DIRECTION_DOWN)
  End If
  fncPopulateInputList
  If oInputDocList.ListItems.Count > 0 Then oInputDocList.ListItems.Item(lReturnIndex + 1).Selected = True
End Sub

Private Sub cmdDocDown_Click()
  mnuMoveDocDown_Click
End Sub

Private Sub mnuMoveDocUp_Click()
Dim lReturnIndex As Long
  If oInputDocList.ListItems.Count > 0 Then
    lReturnIndex = oBruno.oInputDocuments.fncMoveDocument(oInputDocList.SelectedItem.Index - 1, DIRECTION_UP)
  End If
  fncPopulateInputList
  If oInputDocList.ListItems.Count > 0 Then oInputDocList.ListItems.Item(lReturnIndex + 1).Selected = True
End Sub

Private Sub cmdDocUp_Click()
  mnuMoveDocUp_Click
End Sub

Private Sub mnuRemoveSelected_Click()
  If oInputDocList.ListItems.Count > 0 Then
    oBruno.oInputDocuments.fncRemoveDocument (oInputDocList.SelectedItem.Index - 1)
  End If
  fncPopulateInputList
End Sub

Private Sub cmdDelete_Click()
  mnuRemoveSelected_Click
End Sub

Private Sub cmdSaveOutput_Click()
  mnuSaveOutput_Click
End Sub


Private Sub mnuSaveOutput_Click()
    
    Debug.Print "start:" & Now
    
    If oBruno.oInputDocuments.InputDocumentCount > 0 Then
      If oBruno.oInputDocuments.fncCheckDocuments Then
        oBruno.fncSetStatus (STATUS_WORKING)
        'create and set the metadata object
        'only one object (first input doc) regardless of number of input docs
        'note: this was earlier done in oInputDocuments.fncadddocument
        'but since multiple adds might be reordered later
        'has to be done here since now, document order is known
        If oBruno.oInputDocuments.oInputMetadata Is Nothing Then
          Set oBruno.oInputDocuments.oInputMetadata = New cInputMetadata
          oBruno.oInputDocuments.oInputMetadata.SetCommonMetas oBruno.oInputDocuments.InputDocument(0).oDom
        End If

        fncAddMessage "working..."
        DoEvents
        'we have one or more input documents, and a working Driver

         If oBruno.fncCreateAbstractDocuments Then
           'we now have the abstract representation of the fileset
           If oBruno.fncCreateOutputDocuments Then
             'we now have a representation of a "real" fileset
             If Trim$(frmMain.txtOutputPath.Text) = "" Then
               frmMain.txtOutputPath.Text = oBruno.oPaths.AppPath & "out\"
             Else
               If InStr(1, Trim$(frmMain.txtOutputPath.Text), " ") > 0 Then
                 'output path contains spaces
                 frmMain.txtOutputPath.Text = Replace(frmMain.txtOutputPath.Text, " ", "_")
                 frmMain.fncAddMessage "warning: replaced spaces in output path with underscores"
               End If
             End If

             If oBruno.fncRenderOutputDocuments(frmMain.txtOutputPath.Text) Then
               'all ok
               fncAddMessage "all done"
               'mnuInfoBar_setFocus_Click
               Beep
               DoEvents
               Debug.Print "end:" & Now
               'reset all
               If Not oBruno.fncTerminateChildren Then Stop
               Set oBruno = Nothing
               Set oBruno = New cBruno
               fncPopulateInputList
               oBruno.fncSetStatus (STATUS_IDLE)
             Else
               'render failed
                fncAddMessage "RenderOutputDocuments failed"
                oBruno.fncSetStatus (STATUS_ABORTED)
             End If
            Else
             'creation of output documents failed
             fncAddMessage "CreateOutputDocuments failed"
             oBruno.fncSetStatus (STATUS_ABORTED)
           End If
         Else
           'CreateAbstractDocuments failed
           fncAddMessage "createAbstractDocuments failed"
           oBruno.fncSetStatus (STATUS_ABORTED)
         End If
      Else
        fncAddMessage "invalid documents added"
      End If
    Else
      fncAddMessage "no documents added"
    End If
    txtStatus.SetFocus
End Sub

Private Sub cmdOutputPath_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose Save Directory", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)

  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub

  txtOutputPath.Text = objFolder.Self.Path
  If Not Right$(Trim$(txtOutputPath), 1) = "\" Then txtOutputPath.Text = txtOutputPath.Text & "\"

End Sub

Private Sub mnuSetOutPath_Click()
 cmdOutputPath_Click
End Sub

Private Sub mnuUserDocumenation_Click()
  ShellExecute frmMain.hwnd, "Open", App.Path & _
    "\externals\manual\bruno_user_manual.html", "", "", vbNormalFocus
End Sub

Private Sub mnuAdvDoc_Click()
  ShellExecute frmMain.hwnd, "Open", App.Path & _
    "\externals\manual\bruno_Driver.html", "", "", vbNormalFocus

End Sub

Private Sub tmUiUpdate_Timer()
 
 Select Case oBruno.fncGetStatus
   Case STATUS_IDLE
     Me.Caption = "idle - " & oBruno.sAppVersion
   Case STATUS_WORKING
     Select Case lCaptionTimerCounter
       Case 0
         Me.Caption = "working     - " & oBruno.sAppVersion
       Case 1
         Me.Caption = "working.    - " & oBruno.sAppVersion
       Case 2
         Me.Caption = "working..   - " & oBruno.sAppVersion
       Case 3
         Me.Caption = "working...  - " & oBruno.sAppVersion
       Case 4
         Me.Caption = "working..   - " & oBruno.sAppVersion
       Case 5
         Me.Caption = "working.    - " & oBruno.sAppVersion
     End Select
     Refresh
     lCaptionTimerCounter = lCaptionTimerCounter + 1
     If lCaptionTimerCounter = 6 Then lCaptionTimerCounter = 0
   Case STATUS_ABORTED
     Me.Caption = "idle - " & oBruno.sAppVersion
   Case STATUS_DONE
     Me.Caption = "idle - " & oBruno.sAppVersion
   Case STATUS_UNKNOWN
     Me.Caption = "unknown state - " & oBruno.sAppVersion
 End Select
 DoEvents
End Sub

Private Sub txtOutputPath_LostFocus()
  If Not Len(Trim$(txtOutputPath.Text)) = 0 Then
    txtOutputPath.Text = Replace$(txtOutputPath.Text, "/", "\")
    If Not Right$(Trim$(txtOutputPath), 1) = "\" Then
      txtOutputPath.Text = txtOutputPath.Text & "\"
    End If
  End If
End Sub

Private Sub subPopulateDriverList()
Dim i As Long
    cmbDriverList.Clear
    If oDriverList Is Nothing Then Exit Sub
    If oDriverList.DriverCount > 0 Then
      For i = 0 To oDriverList.DriverCount - 1
        cmbDriverList.AddItem (oDriverList.Driver(i).sName)
      Next
      'try to find a match in registry for driver name
      Dim sValue As String
      oRegistryControl.fncLoadRegistryData "DriverName", sValue, &H80000001, , ""
      For i = 0 To oDriverList.DriverCount - 1
       If oDriverList.Driver(i).sName = sValue Then
         cmbDriverList.ListIndex = i
         Exit Sub
       End If
      Next i
      'if no match found in registry, just set first driver as focus
      cmbDriverList.ListIndex = 0
    End If

End Sub

Public Function fncAddMessage(sMessage As String)
  On Error GoTo errh
  txtStatus.Text = sMessage & vbCrLf & txtStatus.Text
  txtStatus.SelStart = 0
errh:
End Function

Private Sub mnuDriverSelect_setFocus_Click()
  cmbDriverList.SetFocus
End Sub

Private Sub mnuInfoBar_setFocus_Click()
 txtStatus.SetFocus
 txtStatus.SelStart = 0
 'txtStatus.SelLength = Len(txtStatus.Text)
End Sub

Private Sub mnuInputDocList_setFocus_Click()
 oInputDocList.SetFocus
End Sub

Private Sub mnuOutputPath_setFocus_Click()
  txtOutputPath.SetFocus
  txtOutputPath.SelStart = 0
  txtOutputPath.SelLength = Len(txtOutputPath.Text)
End Sub
 
Private Sub oInputDocList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Dim Item As MSComctlLib.ListItem
    Set Item = oInputDocList.SelectedItem
    If Item.Index <= oBruno.oInputDocuments.InputDocumentCount Then
      If oBruno.oInputDocuments.InputDocument(Item.Index - 1).bValid Then
        fncAddMessage _
        "encoding: " & oBruno.oInputDocuments.InputDocument(Item.Index - 1).sEncoding _
        & vbCrLf & "namespace: " & oBruno.oInputDocuments.InputDocument(Item.Index - 1).oDom.documentElement.namespaceURI _
        & vbCrLf & "path: " & oBruno.oInputDocuments.InputDocument(Item.Index - 1).sFullPath
      Else
        fncAddMessage _
          "invalid document: " & vbCrLf & _
          "path: " & oBruno.oInputDocuments.InputDocument(Item.Index - 1).sFullPath
      End If
    End If
  End If 'KeyCode = vbKeyReturn Then
End Sub
