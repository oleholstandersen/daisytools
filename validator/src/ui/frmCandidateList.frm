VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCandidateList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Candidate List"
   ClientHeight    =   3915
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6180
   Icon            =   "frmCandidateList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6180
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRunAll 
      Caption         =   "Run &All"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdRunChecked 
      Caption         =   "Run &Checked"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstCandidates 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuAddCandidates 
      Caption         =   "&Add"
      Begin VB.Menu mnuAddSingleDTB 
         Caption         =   "Add &Single DTB"
      End
      Begin VB.Menu mnuAddMultiVolume 
         Caption         =   "Add &Multivolume DTB"
      End
      Begin VB.Menu mnuAddSingleFiles 
         Caption         =   "Add Single &File(s)..."
         Begin VB.Menu mnuAddSingleNcc 
            Caption         =   "ncc..."
         End
         Begin VB.Menu mnuAddSingleSmil 
            Caption         =   "smil..."
         End
         Begin VB.Menu mnuAddSingleMasterSmil 
            Caption         =   "master smil..."
         End
         Begin VB.Menu mnuAddSingleContentDoc 
            Caption         =   "content doc..."
         End
         Begin VB.Menu mnuAddSingleDiscInfo 
            Caption         =   "discinfo..."
         End
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run"
      Begin VB.Menu mnuRunChecked 
         Caption         =   "Run Checked"
      End
      Begin VB.Menu mnuRunAll 
         Caption         =   "Run All"
      End
   End
   Begin VB.Menu mnuRemove 
      Caption         =   "&Remove"
      Begin VB.Menu mnuRemoveSelected 
         Caption         =   "Remove &Checked"
      End
      Begin VB.Menu mnuRemoveUnselected 
         Caption         =   "Remove Unchecked"
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "Remove &All"
      End
   End
End
Attribute VB_Name = "frmCandidateList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Daisy 2.02 Validator Engine
' Copyright (C) 2002 Daisy Consortium
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
'
' For information about the Daisy Consortium, visit www.daisy.org or contact
' info@mail.daisy.org. For development issues, contact markus.gylling@tpb.se or
' karl.ekdahl@tpb.se.

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdRunAll_Click()
    If lCandidatesAdded > 0 Then RunQueue (runall)
End Sub

Private Sub cmdRunChecked_Click()
    If lCandidatesAdded > 0 Then RunQueue (runchecked)
End Sub

Private Sub lstCandidates_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Item.Checked = True Then aCandidateQueue(Item.Index - 1).bolChecked = True Else _
    aCandidateQueue(Item.Index - 1).bolChecked = False
End Sub

Private Sub mnuAddMultiVolume_Click()
    AddMultiVolume
    'Form_Activate
End Sub

'Private Sub mnuAddMotherDir_Click()
'Dim objShell As New Shell, objFolder As Folder, vRoot As Variant
'Dim oFSO As Object, oFile As Object, oFolder As Object
'Dim sMotherDir As String
'
'  Set oFSO = CreateObject("Scripting.FileSystemObject")
'  Set objFolder = objShell.BrowseForFolder(hwnd, "Choose Mother folder of one or several DTBs", _
'    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
'    BIF_VALIDATE)
'
'  If objFolder Is Nothing Then Exit Sub
'  If Not objFolder.Self.IsFileSystem Then Exit Sub
'  sMotherDir = objFolder.Self.Path: If Not Left$(sMotherDir, 1) = "\" Then sMotherDir = sMotherDir & "\"
'
'  Set oFolder = oFSO.getfolder(sTemp)
'
'  If oFolder.fileExists(ncc.html) Then
'  'xxx
'  End If
'
'  objTextSaveLogPath.Text = objFolder.Self.Path
'  If Not Left$(objTextSaveLogPath.Text, 1) = "\" Then objTextSaveLogPath.Text = _
'    objTextSaveLogPath.Text & "\"
'End Sub

Private Sub mnuAddSingleContentDoc_Click()
    AddSingleContentDoc
    'Form_Activate
End Sub

Private Sub mnuAddSingleDiscInfo_Click()
    AddSingleDiscInfo
    'Form_Activate
End Sub

Private Sub mnuAddSingleDtb_Click()
    AddSingleDtb
    'Form_Activate
End Sub

Private Sub mnuAddSingleMasterSmil_Click()
    AddSingleMasterSmil
    'Form_Activate
End Sub

Private Sub mnuAddSingleNcc_Click()
    AddSingleNcc
    'Form_Activate
End Sub

Private Sub mnuAddSingleSmil_Click()
    AddSingleSmil
    'Form_Activate
End Sub

Private Sub mnuRemoveAll_Click()
Dim i As Long
    If lCandidatesAdded > 0 Then
        'For i = lCandidatesAdded To 1 Step -1
        '    fncRemoveCandidateFromArray (i - 1)
        'Next i
        Do Until lCandidatesAdded = 0
          fncRemoveCandidateFromArray (0)
        Loop
        'Form_Activate
    End If
End Sub

Private Sub mnuRemoveSelected_Click()
Dim mItem As ListItem
Dim i As Long
    If lCandidatesAdded > 0 Then
        With lstCandidates
            For i = lCandidatesAdded To 1 Step -1
                Set mItem = .ListItems.Item(i)
                If mItem.Checked Then
                    fncRemoveCandidateFromArray (i - 1)
                End If
            Next i
        End With
'        Form_Activate
    End If
    
End Sub

Private Sub mnuRemoveUnselected_Click()
Dim mItem As ListItem
Dim i As Long
    If lCandidatesAdded > 0 Then
        With lstCandidates
            For i = lCandidatesAdded To 1 Step -1
                Set mItem = .ListItems.Item(i)
                If Not mItem.Checked Then
                    fncRemoveCandidateFromArray (i - 1)
                End If
            Next i
        End With
        'Form_Activate
    End If
End Sub

Private Sub mnuRunAll_Click()
    If lCandidatesAdded > 0 Then RunQueue (runall)
End Sub

Private Sub mnuRunChecked_Click()
    If lCandidatesAdded > 0 Then RunQueue (runchecked)
End Sub


