VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form IteratorFrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validator Iterator"
   ClientHeight    =   5760
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10530
   Icon            =   "IteratorFrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   2295
      Left            =   4920
      TabIndex        =   12
      Top             =   360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4048
      _Version        =   393217
      TextRTF         =   $"IteratorFrmMain.frx":0442
   End
   Begin VB.TextBox txTimeFluctuation 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Text            =   "1000"
      Top             =   3240
      Width           =   735
   End
   Begin VB.Frame frameSettings 
      Caption         =   "settings"
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4455
      Begin VB.CheckBox chkDisableAudioTests 
         Caption         =   "disable audio tests"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox chkLightMode 
         Caption         =   "light mode"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Time Fluctuation (ms)"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdLocateCandidates 
      Caption         =   "locate"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdRunValidation 
      Caption         =   "run validation"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstCandidates 
      Height          =   2655
      Left            =   4920
      TabIndex        =   13
      Top             =   3000
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtOutputReportPath 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtMotherFolder 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "dtb list"
      Height          =   255
      Left            =   4920
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "status log"
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "output report path"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "mother folder"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSetMotherFolder 
         Caption         =   "Set Mother Folder..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSetReportOutputPath 
         Caption         =   "Set Report Output Path..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDividor1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuCandidates 
      Caption         =   "Candidates"
      Begin VB.Menu mnuLocateCandidates 
         Caption         =   "Locate Candidates"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuDividor2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunValidation 
         Caption         =   "Run Validation"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuDivisor3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove All"
      End
      Begin VB.Menu mnuRemoveSelected 
         Caption         =   "Remove Selected"
      End
      Begin VB.Menu mnuRemoveUnselected 
         Caption         =   "Remove Unselected"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "Validator Iterator Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "IteratorFrmMain"
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

Const BIF_RETURNONLYFSDIRS   As Long = &H1
Const BIF_DONTGOBELOWDOMAIN  As Long = &H2
Const BIF_VALIDATE           As Long = &H20
Const BIF_EDITBOX            As Long = &H10
Const BIF_NEWDIALOGSTYLE     As Long = &H40
Const BIF_USENEWUI As Long = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)


Private Sub chkLightMode_Click()
 If chkLightMode.Value = vbChecked Then
   bolLightMode = True
 Else
   bolLightMode = False
 End If
End Sub

Private Sub chkDisableAudioTests_Click()
 If chkDisableAudioTests.Value = vbChecked Then
   bolDisableAudioTests = True
 Else
   bolDisableAudioTests = False
 End If
End Sub

'Private Sub chkMoveDtbAfter_Click()
'  If chkMoveDtbAfter.Value = vbChecked Then
'    bolMoveDtbAfter = True
'  Else
'    bolMoveDtbAfter = False
'  End If
'  fncUpdateUi
'End Sub

Private Sub cmdCancel_Click()
  bolCancel = True
End Sub

Private Sub cmdLocateCandidates_Click()
  mnuLocateCandidates_Click
End Sub

Private Sub cmdRunValidation_Click()
  mnuRunValidation_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
  fncSaveToRegistry
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
  IteratorFrmAbout.Show vbModal
End Sub

Private Sub mnuHelpManual_Click()
  ShellExecute IteratorFrmMain.hwnd, "Open", App.Path & _
    "\Manual\validatorIterator_user_manual.html", "", "", vbNormalFocus
End Sub

Private Sub mnuRunValidation_Click()
  fncRunQueue
End Sub

Private Sub mnuSetMotherFolder_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose Mother folder of one or several DTBs", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)
  
  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub
  sMotherDir = objFolder.Self.Path: If Not Left$(sMotherDir, 1) = "\" Then sMotherDir = sMotherDir & "\"
  fncUpdateUi
  mnuLocateCandidates_Click
End Sub

Private Sub mnuLocateCandidates_Click()
  If Not oFSO.folderexists(sMotherDir) Then
    MsgBox "Mother directory '" & sMotherDir & "' does not exist"
    Exit Sub
  End If
  
  'empty the array
  sNccPathArrayItems = 0
  ReDim Preserve sNccPathArray(sNccPathArrayItems)
  fncUpdateUi
  
  If Not fncFindFiles(sMotherDir, "ncc.html") Then
    MsgBox ("locate candidates failed")
    Exit Sub
  Else
    fncUpdateCandidateList
  End If
End Sub

'Private Sub mnuSetMoveDestination_Click()
'Dim objShell As New Shell, objFolder As Folder, vRoot As Variant
'Dim oFile As Object, oFolder As Object, oFolders As Object
'  On Error GoTo ErrHandler
'  Set oFSO = CreateObject("Scripting.FileSystemObject")
'  Set objFolder = objShell.BrowseForFolder(hWnd, "Choose DTB Move Destination", _
'    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
'    BIF_VALIDATE)
'
'  If objFolder Is Nothing Then Exit Sub
'  If Not objFolder.Self.IsFileSystem Then Exit Sub
'  sDtbMoveDestination = objFolder.Self.Path: If Not Left$(sDtbMoveDestination, 1) = "\" Then sDtbMoveDestination = sDtbMoveDestination & "\"
'  If Not oFSO.folderexists(sDtbMoveDestination) Then oFSO.createFolder (sDtbMoveDestination)
'  If Not oFSO.folderexists(sDtbMoveDestination & "pass") Then oFSO.createFolder (sDtbMoveDestination & "pass")
'  If Not oFSO.folderexists(sDtbMoveDestination & "fail") Then oFSO.createFolder (sDtbMoveDestination & "fail")
'
'  fncUpdateUi
'ErrHandler:
'
'
'End Sub

Private Sub mnuSetReportOutputPath_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant
Dim oFile As Object, oFolder As Object
  On Error GoTo Errhandler
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose Report Output Directory", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)
  
  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub
  sReportDir = objFolder.Self.Path: If Not Left$(sReportDir, 1) = "\" Then sReportDir = sReportDir & "\"

  If Not oFSO.folderexists(sReportDir) Then oFSO.createFolder (sReportDir)
  fncUpdateUi
Errhandler:
End Sub

Private Sub mnuRemoveSelected_Click()
Dim mItem As ListItem
Dim i As Long
    If sNccPathArrayItems > 0 Then
        With lstCandidates
            For i = 1 To .ListItems.Count
                Set mItem = .ListItems.Item(i)
                If mItem.Checked Then
                    fncRemoveCandidateFromArray (i - 1)
                End If
            Next i
        End With
    End If
    fncUpdateCandidateList
End Sub

Private Sub mnuRemoveUnselected_Click()
Dim mItem As ListItem
Dim i As Long
    If sNccPathArrayItems > 0 Then
        With lstCandidates
            For i = 1 To .ListItems.Count
                Set mItem = .ListItems.Item(i)
                If Not mItem.Checked Then
                    fncRemoveCandidateFromArray (i - 1)
                End If
            Next i
        End With
    End If
    fncUpdateCandidateList
End Sub

Private Sub mnuRemoveAll_Click()
 With lstCandidates
   .ListItems.Clear
 End With
 sNccPathArrayItems = 0
 ReDim Preserve sNccPathArray(sNccPathArrayItems)
 fncUpdateCandidateList
End Sub

Public Function fncUpdateCandidateList()
Dim colItem As ColumnHeader
Dim mItem As ListItem
Dim i As Long, aChecked() As CheckBoxConstants, lCheckCount As Long
    
    With lstCandidates
       .View = lvwReport
       .ColumnHeaders.Clear
       .ListItems.Clear
       .LabelEdit = lvwManual
       
        Set colItem = .ColumnHeaders.Add()
            colItem.Text = "Candidate"
            colItem.Width = .Width * 0.8
        Set colItem = .ColumnHeaders.Add()
            colItem.Text = "Result"
            colItem.Width = .Width * 0.2
      For i = 0 To (sNccPathArrayItems - 1)
          Set mItem = .ListItems.Add()
          mItem.Text = sNccPathArray(i)
          mItem.Checked = True
      Next i
    End With
End Function

Public Function fncUpdateUi()

  Me.txtMotherFolder = sMotherDir
  Me.txtOutputReportPath = sReportDir
  
  If bolLightMode Then
    Me.chkLightMode.Value = vbChecked
  Else
    Me.chkLightMode.Value = vbUnchecked
  End If
  
  If bolDisableAudioTests Then
    Me.chkDisableAudioTests.Value = vbChecked
  Else
    Me.chkDisableAudioTests.Value = vbUnchecked
  End If
    
  Me.txTimeFluctuation = lTimeFluct
  
'  If bolMoveDtbAfter Then
'    Me.chkMoveDtbAfter.Value = vbChecked
'    Me.txtDtbDestination.Text = sDtbMoveDestination
'  Else
'    Me.chkMoveDtbAfter.Value = vbUnchecked
'    Me.txtDtbDestination.Text = ""
'  End If
    
'  Me.txtDtbDestination.Enabled = Not bolValidating
  Me.txTimeFluctuation.Enabled = Not bolValidating
  Me.txtMotherFolder.Enabled = Not bolValidating
  Me.txtOutputReportPath.Enabled = Not bolValidating
  Me.cmdLocateCandidates.Enabled = Not bolValidating
  Me.cmdRunValidation.Enabled = Not bolValidating
  Me.chkLightMode.Enabled = Not bolValidating
  Me.chkDisableAudioTests.Enabled = Not bolValidating
'  Me.chkMoveDtbAfter.Enabled = Not bolValidating
End Function

Private Sub txTimeFluctuation_lostfocus()
Dim lTest As Long
  If Not IsNumeric(txTimeFluctuation.Text) Then
    lTest = "1000"
  Else
   lTest = CLng(txTimeFluctuation.Text)
   If lTest < 0 Then lTest = 0
   If lTest > 1500 Then lTest = 1500
  End If
  
  txTimeFluctuation.Text = lTest
  lTimeFluct = lTest
End Sub

