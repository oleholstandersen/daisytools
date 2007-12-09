VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Regenerator AudioNameReverter"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Multi DTB"
      Enabled         =   0   'False
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   5775
      Begin VB.CommandButton cmdRunMulti 
         Caption         =   "run multi rename"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdBrowseMother 
         Caption         =   "browse"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txMotherDirPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Text            =   "all in this mother directory"
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Single DTB"
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton cmdRun 
         Caption         =   "run single  rename"
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdBrowseAud 
         Caption         =   "browse"
         Height          =   300
         Left            =   4200
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowseNfo 
         Caption         =   "browse"
         Height          =   300
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txAudioDirPath 
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Text            =   "path to audio file directory"
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txRenameNfoPath 
         Height          =   300
         Left            =   240
         TabIndex        =   0
         Text            =   "path to audioRenameNfo.xml"
         Top             =   360
         Width           =   3735
      End
   End
   Begin MSComDlg.CommonDialog objCommonDialog 
      Left            =   5640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfLog 
      Height          =   3975
      Left            =   6120
      TabIndex        =   6
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7011
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":0442
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

Private Sub cmdBrowseAud_Click()

Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hWnd, "Point to the folder where the audio files to be un-renamed reside", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)

  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub

  txAudioDirPath.Text = objFolder.Self.Path
  If Not Left$(txAudioDirPath.Text, 1) = "\" Then txAudioDirPath.Text = _
    txAudioDirPath.Text & "\"

End Sub

Private Sub cmdBrowseMother_Click()
Dim objShell As New Shell, objFolder As Folder, vRoot As Variant

  Set objFolder = objShell.BrowseForFolder(hWnd, "Point to the folder where DTB folders for rename exists", _
    BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
    BIF_VALIDATE)

  If objFolder Is Nothing Then Exit Sub
  If Not objFolder.Self.IsFileSystem Then Exit Sub

  txMotherDirPath.Text = objFolder.Self.Path
  If Not Left$(txMotherDirPath.Text, 1) = "\" Then txMotherDirPath.Text = _
    txMotherDirPath.Text & "\"

End Sub

Private Sub cmdBrowseNfo_Click()
Dim sFileName As String
  If Not fncOpenFile("Xml file (*.xml)|*.xml", "*.xml", True, sFileName) Then Exit Sub
  txRenameNfoPath.Text = sFileName
End Sub

Private Function fncOpenFile( _
  ByVal sMask As String, _
  ByVal sFileName As String, _
  ByVal bolMustExist As Boolean, _
  ByRef sOutPut As String _
  ) As Boolean
  
  On Error GoTo ErrorH
  
  With objCommonDialog
    .CancelError = True
    .Filter = sMask
    .FilterIndex = 1
    .FileName = sFileName
    If bolMustExist Then .Flags = cdlOFNFileMustExist
    .Flags = cdlOFNNoChangeDir
    .ShowOpen
    sOutPut = .FileName
  End With
   
  fncOpenFile = True
ErrorH:
End Function


Private Sub cmdRun_Click()
  
  If Not fncFolderExists(txAudioDirPath.Text) Then
    addLog "audio folder" & txAudioDirPath.Text & "not found on filesystem"
    Exit Sub
  End If
  
  If Not fncFileExists(txRenameNfoPath.Text) Then
    addLog "rename nfo" & txRenameNfoPath.Text & "not found on filesystem"
    Exit Sub
  End If
  
  If Right(txAudioDirPath.Text, 1) <> "\" Then txAudioDirPath.Text = txAudioDirPath.Text & "\"
  If Not fncRunRename(txAudioDirPath.Text, txRenameNfoPath.Text) Then
    Exit Sub
  Else
    addLog "rename done"
  End If
  
End Sub


