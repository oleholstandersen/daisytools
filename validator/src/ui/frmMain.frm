VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10350
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfEdit 
      Height          =   6615
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   11668
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   1e6
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin VB.CommandButton cmSave 
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":04C4
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Save document"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmFindReplace 
      Height          =   375
      Left            =   5880
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":0806
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Search/Replace"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin MSComctlLib.ImageList ilMain 
      Left            =   8760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1830
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ED4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2226
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2578
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4496
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":528C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmBack 
      Height          =   375
      Left            =   6480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":55DE
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Go back to previous document"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmFollowLink 
      Height          =   375
      Left            =   7080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":5860
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Follow link"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.TextBox txErrView 
      BackColor       =   &H8000000F&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   7680
      Width           =   10095
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":5AE2
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Abort validation process"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdRunSelected 
      Height          =   375
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":5E24
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Run selected candidate"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmUiUpdate 
      Interval        =   1000
      Left            =   9840
      Top             =   0
   End
   Begin VB.CommandButton cmdRunAll 
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":6166
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Run all added candidates"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin MSComctlLib.TreeView treeErrorView 
      Height          =   6615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11668
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8430
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "idle"
            TextSave        =   "idle"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1746
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1746
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1746
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1746
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1746
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1746
            MinWidth        =   1058
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1746
            MinWidth        =   1058
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRtf 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Document Editor"
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblTree 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Validation Report View"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveReport 
         Caption         =   "&Save report..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuLoadReport 
         Caption         =   "&Load report..."
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuFileLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuCandidates 
      Caption         =   "&Candidates"
      Begin VB.Menu mnuAddSingleDtb 
         Caption         =   "Add single &DTB..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuAddMultiVolume 
         Caption         =   "Add &multivolume DTB..."
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuAddSingleFiles 
         Caption         =   "Add &Single File(s)..."
         Begin VB.Menu mnuAddSingleNcc 
            Caption         =   "&ncc..."
         End
         Begin VB.Menu mnuAddSingleSmil 
            Caption         =   "&smil..."
         End
         Begin VB.Menu mnuAddSingleMasterSmil 
            Caption         =   "&mastersmil..."
         End
         Begin VB.Menu mnuAddSingleContentDoc 
            Caption         =   "&content doc..."
         End
         Begin VB.Menu mnuAddSingleDiscInfo 
            Caption         =   "&discinfo..."
         End
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddSingleDtbLight 
         Caption         =   "Add single DTB light test ..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCandidateLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLightMode 
         Caption         =   "Light mode"
      End
      Begin VB.Menu mnuDisableAudio 
         Caption         =   "Disable all audio tests"
      End
      Begin VB.Menu mnuCandidateLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCandidateList 
         Caption         =   "View &Candidate List..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuClearAllCandidates 
         Caption         =   "Clear all added"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRunSelected 
         Caption         =   "Run &selected"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuRunAll 
         Caption         =   "&Run all"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Abort validation"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuEditor 
      Caption         =   "Editor"
      Begin VB.Menu mnuEditorSave 
         Caption         =   "&Save document..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEditorFind 
         Caption         =   "S&earch/Replace"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFollowLink 
         Caption         =   "&Follow link"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuGoBack 
         Caption         =   "Go &back"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuOnlineOfflineSettings 
         Caption         =   "Online/Offline Settings..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSettingsDocuPaths 
         Caption         =   "Document Paths..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuGeneralSettings 
         Caption         =   "&General settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuOpenHelp 
         Caption         =   "Open User Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuOpenDevManual 
         Caption         =   "Open Developer Manual"
      End
      Begin VB.Menu mnuViewErrorLog 
         Caption         =   "Error &log..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopupRTF 
      Caption         =   "popupRT"
      Visible         =   0   'False
      Begin VB.Menu mnuPopFollowLink 
         Caption         =   "&Follow link"
      End
      Begin VB.Menu mnuPopGoBack 
         Caption         =   "Go &back"
      End
   End
   Begin VB.Menu mnuPopupTree 
      Caption         =   "popupTree"
      Visible         =   0   'False
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnuGotoSpec 
         Caption         =   "&Goto specification"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINELENGTH = &HC1

Public WithEvents objDllEvent As oEvent
Attribute objDllEvent.VB_VarHelpID = -1
Dim lCaptionTimerCounter As Long, lLastStart, lLastEnd

Public Sub subSetInterfaceStatus(ibolStatus As Boolean)
  mnuFile.Enabled = ibolStatus
  mnuCandidates.Enabled = ibolStatus
  mnuRunAll.Enabled = ibolStatus
  mnuRunSelected.Enabled = ibolStatus
  mnuEditor.Enabled = ibolStatus
  mnuSettings.Enabled = ibolStatus
  cmdRunAll.Enabled = ibolStatus
  cmdRunSelected.Enabled = ibolStatus
  cmFollowLink.Enabled = ibolStatus
  cmBack.Enabled = ibolStatus
  cmSave.Enabled = ibolStatus
  cmFindReplace.Enabled = ibolStatus
End Sub

Private Sub cmBack_Click()
  fncShowFile "", False, False
End Sub

Private Sub cmdCancel_Click()
  fncSetCancelFlag
End Sub

Private Sub cmFindReplace_Click()
  mnuEditorFind_Click
End Sub

Private Sub cmFollowLink_Click()
  subFollowLink
End Sub

Private Sub cmSave_Click()
  mnuEditorSave_Click
End Sub

Public Sub Form_Load()
    Set objDllEvent = xobjEvent
    tmUiUpdate_Timer
    
    Dim lTemp As Long
    sBaseKey = "Software\DaisyWare\Validator"
    fncLoadRegistryData "Apperance\Windowstate", _
      lTemp, HKEY_CURRENT_USER, , vbMaximized: frmMain.WindowState = lTemp
    
    frmMain.Show
    Set treeErrorView.ImageList = ilMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Not bolQuit Then
'        bolQuit = True
'        Cancel = 1
'    End If
  fncSaveToRegistry
  Unload frmAbout
  Unload frmCandidateList
  Unload frmDocuPaths
  Unload frmErrorLog
  Unload dlgSearchReplace
  Unload frmGeneralSettings
End Sub

Private Sub Form_Resize()
    'treeErrorView.Height = (frmMain.Height - 1440) - 480
    'rtfEdit.Height = (frmMain.Height - 1440) - 480
    If (Not Visible) Or (WindowState = vbMinimized) Then Exit Sub
    
    On Error Resume Next
    
    treeErrorView.Height = (frmMain.Height - 2760) - 120
    rtfEdit.Height = (frmMain.Height - 2760) - 120
    
    rtfEdit.Width = frmMain.Width - treeErrorView.Width - 500
    lblTree.Width = treeErrorView.Width
    lblRtf.Width = rtfEdit.Width
    Dim i As Long
    For i = 1 To 7
        StatusBar.Panels(i).Width = frmMain.Width * 0.143
    Next i
    
    txErrView.Top = treeErrorView.Height + treeErrorView.Top + 105
    txErrView.Width = Width - 375
End Sub

Private Sub cmdRunAll_Click()
    If lCandidatesAdded > 0 Then RunQueue (runall)
End Sub
 
Private Sub cmdRunSelected_Click()
    If lCandidatesAdded > 0 Then RunQueue (runselected)
End Sub


Private Sub mnuCancel_Click()
  fncSetCancelFlag
End Sub

Private Sub mnuClearAllCandidates_Click()
Dim i As Long
    For i = 0 To lCandidatesAdded - 1
        fncRemoveCandidateFromArray (0)
    Next i
    lCandidatesAdded = 0
    fncUpdateCandidateList
    
    fncClearFileHistory
End Sub

Private Sub mnuEditorFind_Click()
    dlgSearchReplace.Show
End Sub

Private Sub mnuEditorSave_Click()
    fncSaveFile rtfEdit.Text, filesetfile, aFileHistory(lCurrentHistory).sAbsPath
End Sub

Private Sub mnuExit_Click()
'    bolQuit = True
  Unload Me
End Sub

Private Sub mnuFollowLink_Click()
  subFollowLink
End Sub

Private Sub mnuGeneralSettings_Click()
  frmGeneralSettings.Show vbModal
End Sub

Private Sub mnuGotoSpec_Click()
  Dim objReportItem As oReportItem
  
  Set objReportItem = fncGetSelectedCanRepItem
  If objReportItem Is Nothing Then Exit Sub
  
  On Error Resume Next
  Shell "explorer " & objReportItem.sLink, vbMaximizedFocus
End Sub

Private Sub mnuLightMode_Click()
  bolLightMode = Not bolLightMode
  mnuLightMode.Checked = bolLightMode
  If bolLightMode Then
    StatusBar.Panels.Item(5).Text = "Light Mode ON"
  Else
    StatusBar.Panels.Item(5).Text = ""
  End If
  fncSetLightMode bolLightMode
End Sub

Private Sub mnuDisableAudio_Click()
  bolDisableAudioTests = Not bolDisableAudioTests
  mnuDisableAudio.Checked = bolDisableAudioTests
  If bolDisableAudioTests Then
    StatusBar.Panels.Item(6).Text = "Audio Tests OFF"
  Else
    StatusBar.Panels.Item(6).Text = ""
  End If
  fncSetDisableAudioTests bolDisableAudioTests

End Sub


Private Sub mnuOpenHelp_Click()
  ShellExecute frmMain.hwnd, "Open", App.Path & _
    "\Manual\validator_user_manual.html", "", "", vbNormalFocus
End Sub

Private Sub mnuOpenDevManual_Click()
  ShellExecute frmMain.hwnd, "Open", App.Path & _
    "\Manual\validator_developer_manual.html", "", "", vbNormalFocus
End Sub


Private Sub mnuPopFollowLink_Click()
  subFollowLink
End Sub

Private Sub mnuGoBack_Click()
  fncShowFile "", False, False
End Sub

Private Sub mnuPopGoBack_Click()
  fncShowFile "", False, False
End Sub

Private Sub mnuLoadReport_Click()
  fncLoadReport
End Sub

Private Sub mnuProperties_Click()
  Dim objReportItem As oReportItem, sOutput As String, lCounter As Long
  
  lCounter = fncGetSelectedCandidate
  If lCounter = -1 Then Exit Sub
  
  Select Case aCandidateQueue(lCounter).lCandidateType
    Case TYPE_SINGLEDTB
      sOutput = "Single dtb "
    Case TYPE_MULTIVOLUME
      sOutput = "TYPE_MULTIVOLUME dtb "
    Case TYPE_SINGLE_NCC
      sOutput = "Single ncc document "
    Case TYPE_SINGLE_SMIL
      sOutput = "Single smil document "
    Case TYPE_SINGLE_MSMIL
      sOutput = "Single master smil document "
    Case TYPE_SINGLE_CONTENTDOC
      sOutput = "Single content document "
    Case TYPE_SINGLE_DISCINFO
      sOutput = "Single discinfo document "
  End Select
  
  sOutput = sOutput & "@ " & aCandidateQueue(lCounter).sAbsPath
  
  Set objReportItem = fncGetSelectedCanRepItem
   
  If Not objReportItem Is Nothing Then
    sOutput = sOutput & vbCrLf & vbCrLf
    sOutput = sOutput & objReportItem.sFailType
    If LCase$(objReportItem.sFailType) = "error" Then _
      sOutput = sOutput & " (" & objReportItem.sFailClass & ")"
    If Not objReportItem.sAbsPath = "" Then
      sOutput = sOutput & " " & objReportItem.sAbsPath
      sOutput = sOutput & " [" & objReportItem.lLine & ":" & _
        objReportItem.lColumn & "]"
    End If
    sOutput = sOutput & vbCrLf
    sOutput = sOutput & "Test id: " & objReportItem.sTestId & vbCrLf
    sOutput = sOutput & "Name: " & objReportItem.sName & vbCrLf
    sOutput = sOutput & "Short description: " & objReportItem.sShortDesc & vbCrLf
    sOutput = sOutput & "Long description: " & objReportItem.sLongDesc & vbCrLf
    sOutput = sOutput & "Comment: " & objReportItem.sComment & vbCrLf
    sOutput = sOutput & "Link: " & objReportItem.sLink
  End If
  
  'MsgBox sOutput
  frmRepItemProperties.txRepItemInfo = sOutput
  frmRepItemProperties.Show vbModal
End Sub

    
Private Sub mnuAddSingleContentDoc_Click()
    AddSingleContentDoc
End Sub

Private Sub mnuAddSingleDiscInfo_Click()
    AddSingleDiscInfo
End Sub
    
Private Sub mnuAddSingleMasterSmil_Click()
    AddSingleMasterSmil
End Sub
        
Private Sub mnuAddSingleSmil_Click()
  AddSingleSmil
End Sub

Private Sub mnuAddSingleNcc_Click()
AddSingleNcc
End Sub

Private Sub mnuAddMultiVolume_Click()
    AddMultiVolume
End Sub

Private Sub mnuAddSingleDtb_Click()
    AddSingleDtb
End Sub

'Private Sub mnuAddTYPE_SINGLEDTBLight_Click()
'    AddSingleDTBLight
'End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

'Private Sub mnuOnlineOfflineSettings_Click()
'    frmOnOffLine.Show
'End Sub

Private Sub mnuRemoveAll_Click()
Dim i As Long
    If lCandidatesAdded > 0 Then
        For i = lCandidatesAdded To 1 Step -1
            fncRemoveCandidateFromArray (i)
        Next i
    End If
End Sub

Private Sub mnuRemoveLast_Click()
    If lCandidatesAdded > 0 Then
        fncRemoveCandidateFromArray (lCandidatesAdded)
    End If
End Sub

Private Sub mnuRunAll_Click()
    If lCandidatesAdded > 0 Then RunQueue (runall)
End Sub

Private Sub mnuRunSelected_Click()
    If lCandidatesAdded > 0 Then RunQueue (runselected)
End Sub

Private Sub mnuSaveReport_Click()
    Dim lCounter As Long
    lCounter = fncGetSelectedCandidate
    If (lCounter = -1) Then Exit Sub
    fncSaveReport aCandidateQueue(lCounter)
End Sub

Private Sub mnuSettingsDocuPaths_Click()
    frmDocuPaths.Show vbModal
End Sub

Private Sub mnuViewCandidateList_Click()
    frmCandidateList.Show vbModal
End Sub

Private Sub mnuViewErrorLog_Click()
    frmErrorLog.Show vbModal
End Sub

'Events in dll oErrorLog:
'Public Event evErrorLog(sError As String)
'Public Event evFailedTest()
'Public Event evSucceededTest()

Public Sub objDllEvent_evErrorLog(sError As String)
    lErrorLogs = lErrorLogs + 1
    StatusBar.Panels.Item(4).Text = "log items: " & lErrorLogs
    frmErrorLog.txtErrorCount.Text = lErrorLogs
    frmErrorLog.rtfErrorLog.Text = frmErrorLog.rtfErrorLog.Text & sError & vbCrLf
End Sub

Private Sub objDllEvent_evFailedTest()
    lTestsFailed = lTestsFailed + 1
End Sub

Private Sub objDllEvent_evSucceededTest()
    lTestsSucceeded = lTestsSucceeded + 1
End Sub



Public Sub tmUiUpdate_Timer()
    If sCurrentState = "validating" Then
        StatusBar.Panels.Item(2).Text = "tests passed: " & lTestsSucceeded
        StatusBar.Panels.Item(3).Text = "tests failed: " & lTestsFailed
        
        Select Case lCaptionTimerCounter
            Case 0
                frmMain.Caption = "2.02 dtb validator -- " & sCurrentState & " |"
                frmMain.StatusBar.Panels.Item(1).Text = "status: " & sCurrentState & " [" & lCurrentCandidate & " of " & lCandidatesAdded & "] " & fncGetProgress & "% " & " |"
            Case 1
                frmMain.Caption = "2.02 dtb validator -- " & sCurrentState & " /"
                frmMain.StatusBar.Panels.Item(1).Text = "status: " & sCurrentState & " [" & lCurrentCandidate & " of " & lCandidatesAdded & "] " & fncGetProgress & "% " & " /"
            Case 2
                frmMain.Caption = "2.02 dtb validator -- " & sCurrentState & " -"
                frmMain.StatusBar.Panels.Item(1).Text = "status: " & sCurrentState & " [" & lCurrentCandidate & " of " & lCandidatesAdded & "] " & fncGetProgress & "% " & " -"
            Case 3
                frmMain.Caption = "2.02 dtb validator -- " & sCurrentState & " \"
                frmMain.StatusBar.Panels.Item(1).Text = "status: " & sCurrentState & " [" & lCurrentCandidate & " of " & lCandidatesAdded & "] " & fncGetProgress & "% " & " \"
        End Select
        Refresh
        lCaptionTimerCounter = lCaptionTimerCounter + 1
        If lCaptionTimerCounter = 4 Then lCaptionTimerCounter = 0
    ElseIf sCurrentState = "idle" Then
        frmMain.Caption = "2.02 dtb validator -- " & sCurrentState
        frmMain.StatusBar.Panels.Item(1).Text = "status: " & sCurrentState
        
        Dim lSC As Long, lTP As Long, lTF As Long
        lSC = fncGetSelectedCandidate
        If (lSC > -1) Then
          If (Not aCandidateQueue(lSC).objReport Is Nothing) Then
            lTP = aCandidateQueue(lSC).objReport.lSucceededTestCount
            lTF = aCandidateQueue(lSC).objReport.lFailedTestCount
          End If
        End If
        StatusBar.Panels.Item(2).Text = "tests passed: " & lTP
        StatusBar.Panels.Item(3).Text = "tests failed: " & lTF
    Else
      Caption = "2.02 dtb validator -- " & sCurrentState
      StatusBar.Panels.Item(1).Text = "status: " & sCurrentState & _
        "(" & Int((100 / lIntProgMax) * lIntProgress) & "%)"
      StatusBar.Panels.Item(2).Text = ""
      StatusBar.Panels.Item(3).Text = ""
    End If
    
    If lCurrentHistory < 2 Then _
      frmMain.cmBack.Enabled = False Else frmMain.cmBack.Enabled = True
    
    If lCurrentHistory < 1 Then
      frmMain.cmSave.Enabled = False: frmMain.cmFindReplace.Enabled = False
      frmMain.cmFollowLink.Enabled = False
    Else
      frmMain.cmSave.Enabled = True: frmMain.cmFindReplace.Enabled = True
      frmMain.cmFollowLink.Enabled = True
    End If
    
    DoEvents
End Sub

Private Sub treeErrorView_Click()
  Dim lCounter As Long, lCounter2 As Long
  Dim objReportItem As oReportItem
  
  Set objReportItem = fncGetSelectedCanRepItem
  tmUiUpdate_Timer
  If objReportItem Is Nothing Then txErrView.Text = "": Exit Sub
  
  txErrView.Text = objReportItem.sFailType
  If LCase$(objReportItem.sFailType) = "error" Then _
    txErrView.Text = txErrView.Text & " (" & objReportItem.sFailClass & ")"
  
  If Not objReportItem.sAbsPath = "" Then _
    txErrView.Text = txErrView.Text & " in " & objReportItem.sAbsPath & " [" & _
      objReportItem.lLine & ":" & objReportItem.lColumn & "]"

  txErrView.Text = txErrView.Text & vbCrLf & ": " & objReportItem.sLongDesc
  If Not objReportItem.sComment = "" Then _
    txErrView.Text = txErrView & ", " & objReportItem.sComment
  If Not objReportItem.sLink = "" Then _
    txErrView.Text = txErrView.Text & vbCrLf & "For further information see " & _
    objReportItem.sLink
End Sub

Private Sub treeErrorView_DblClick()
  Dim lCounter As Long, lCounter2 As Long
  
  Dim objReportItem As oReportItem
  
  Set objReportItem = fncGetSelectedCanRepItem
  If objReportItem Is Nothing Then txErrView.Text = "": Exit Sub

  fncShowFile objReportItem.sAbsPath, True, True

  If (objReportItem.lLine < 1) Or (objReportItem.lColumn < 1) Then
    rtfEdit.SelStart = 0
    Exit Sub
  End If
  
  'rtfEdit.SelStart + 1
'  Private Const EM_GETLINECOUNT = &HBA
'Private Const EM_LINEINDEX = &HBB
'Private Const EM_LINELENGTH = &HC1

  Dim lLastLine As Long, lLastChar As Long
  lLastLine = SendMessage(rtfEdit.hwnd, EM_GETLINECOUNT, 0, 0)
  lLastChar = SendMessage(rtfEdit.hwnd, EM_LINEINDEX, lLastLine - 1, 0)
  
  lCounter = 1
  lCounter2 = 1
  Do
    'If lCounter2 = 0 Then Exit Sub
    'lCounter = lCounter + 1
    'lCounter2 = InStr(lCounter2 + 1, rtfEdit.Text, vbCr, vbBinaryCompare)
    lCounter = rtfEdit.GetLineFromChar(lCounter2) + 1
    lCounter2 = lCounter2 + 1
    If lCounter2 > lLastChar Then GoTo Skip
    DoEvents
  Loop Until (lCounter >= CLng(objReportItem.lLine))
  
  rtfEdit.SelStart = lCounter2 + CLng(objReportItem.lColumn) - 2
Skip:
  rtfEdit.SetFocus
End Sub

Private Function fncMayBeUnicode(sStringToSearch) As Boolean
Dim sFirstDocLine As String
Dim lEndOfFirstLine As Long
Dim bolTest As Boolean

 bolTest = False
 ' get the first line only
 lEndOfFirstLine = InStr(1, sStringToSearch, "?>")
 If lEndOfFirstLine > 0 Then
   sFirstDocLine = Mid(sStringToSearch, 1, lEndOfFirstLine)
   If (InStr(1, sFirstDocLine, "utf-8", vbTextCompare) > 0) Or _
     (InStr(1, sFirstDocLine, "utf8", vbTextCompare) > 0) Or _
     (InStr(1, sFirstDocLine, "utf-16", vbTextCompare) > 0) Or _
     (InStr(1, sFirstDocLine, "utf16", vbTextCompare) > 0) Or _
     (InStr(1, sFirstDocLine, "encoding", vbTextCompare) < 1) _
   Then
     bolTest = True
   End If
 End If 'lEndOfFirstLine > 0 Then
 
 fncMayBeUnicode = bolTest

End Function

Private Sub rtfEdit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    PopupMenu mnuPopupRTF
  End If
End Sub

Private Sub treeErrorView_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then treeErrorView_DblClick
End Sub

Private Sub treeErrorView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
    Dim objReportItem As Object
    
    Set objReportItem = fncGetSelectedCanRepItem
    If objReportItem Is Nothing Then Exit Sub
    
    If objReportItem.sLink = "" Then
      mnuGotoSpec.Enabled = False
    Else
      mnuGotoSpec.Enabled = True
    End If
  
    PopupMenu mnuPopupTree
  End If
End Sub

Private Sub treeErrorView_NodeClick(ByVal Node As MSComctlLib.Node)
  treeErrorView_Click
End Sub

Private Sub rtfEdit_LostFocus()
        With StatusBar.Panels
            .Item(6).Text = ""
            .Item(7).Text = ""
            .Item(8).Text = ""
        End With
        rtfEdit.Locked = False
End Sub

Private Sub rtfEdit_GotFocus()
    Dim lCharIndex As Long, lLine As Long
    lLine = rtfEdit.GetLineFromChar(rtfEdit.SelStart)
    lCharIndex = (rtfEdit.SelStart + 1) - _
      SendMessage(rtfEdit.hwnd, EM_LINEINDEX, lLine, 0)
    
    With StatusBar.Panels
        .Item(6).Text = "Line " & (rtfEdit.GetLineFromChar(rtfEdit.SelStart) + 1)
        .Item(7).Text = "Column " & lCharIndex
    End With
    
    If fncMayBeUnicode(rtfEdit.Text) Then
      rtfEdit.Locked = True
      StatusBar.Panels.Item(8).Text = "READ ONLY"
    Else
      rtfEdit.Locked = False
      StatusBar.Panels.Item(8).Text = ""
    End If

    
End Sub

Private Sub rtfEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lCharIndex As Long, lLine As Long
    lLine = rtfEdit.GetLineFromChar(rtfEdit.SelStart)
    lCharIndex = (rtfEdit.SelStart + 1) - _
      SendMessage(rtfEdit.hwnd, EM_LINEINDEX, lLine, 0)
    
    With StatusBar.Panels
        .Item(6).Text = "Line " & (rtfEdit.GetLineFromChar(rtfEdit.SelStart) + 1)
        .Item(7).Text = "Column " & lCharIndex
'        .Item(7).Text = "Column " & (fncGetColumn(rtfEdit.Text, rtfEdit.SelStart))
    End With

'    With StatusBar.Panels
'        .Item(6).Text = "Line " & (rtfEdit.GetLineFromChar(rtfEdit.SelStart) + 1)
'        .Item(7).Text = "Column " & (fncGetColumn(rtfEdit.Text, rtfEdit.SelStart) + 1)
'    End With
End Sub

'Private Function fncHighlightLine()
'  Dim lStart As Long, lEnd As Long, lPreservePos As Long
'
'  fncRestoreHighlight
'
'  lStart = InStrRev(rtfEdit.Text, vbCrLf, rtfEdit.SelStart, vbBinaryCompare)
'  lEnd = InStr(rtfEdit.SelStart, rtfEdit.Text, vbCrLf, vbBinaryCompare)
'
'  lPreservePos = rtfEdit.SelStart
'
'  lLastStart = lStart
'  lLastEnd = lEnd
'
'  rtfEdit.SelStart = lStart + 2
'  rtfEdit.SelLength = lEnd - lStart - 2
'  rtfEdit.SelColor = RGB(255, 0, 0)
'
'  rtfEdit.SelStart = lPreservePos
'End Function
'
'Public Function fncRestoreHighlight()
'  If lLastStart = 0 And lLastEnd = 0 Then Exit Function
'
'  Dim lPreservePos As Long
'
'  lPreservePos = rtfEdit.SelStart
'
'  rtfEdit.SelStart = lLastStart + 2
'  rtfEdit.SelLength = lLastEnd - lLastStart - 2
'  rtfEdit.SelColor = RGB(0, 0, 0)
'
'  rtfEdit.SelStart = lPreservePos
'End Function

Public Function fncGetColumn(ByVal strTemp As String, ByVal lngCursorPos As Long) As Long
Dim i As Long, lngStepCounter As Long

    lngStepCounter = 0
    
    If lngCursorPos < 2 Then 'on first rows first two chars
        lngStepCounter = lngCursorPos
    ElseIf InStrRev(strTemp, vbCrLf, lngCursorPos) = 0 Then 'on first row where there is no leading VbCrLf
        lngStepCounter = lngCursorPos
    Else 'all other rows
        For i = lngCursorPos To 0 Step -1
            If Mid$(strTemp, i, 2) = vbCrLf Then
                Exit For
            Else
                lngStepCounter = lngStepCounter + 1
            End If
        Next i
        lngStepCounter = lngStepCounter - 1
    End If

    fncGetColumn = lngStepCounter
End Function

Private Sub treeErrorView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  'data.GetData
  '
  'Stop
  
  Dim oFSO As Object, lValue As Long, lValue2 As Long, sName As String
  Dim sTemp As String
  Set oFSO = CreateObject("scripting.FileSystemObject")
  
  For lValue = 1 To Data.Files.Count
    sName = Data.Files.Item(lValue)
    
    If oFSO.folderexists(sName) Then
      If Not Right$(sName, 1) = "\" Then sName = sName & "\"
      fncAddCandidateToArray TYPE_SINGLEDTB, sName, True
    ElseIf oFSO.fileExists(sName) Then
      lValue2 = InStrRev(sName, "\", , vbBinaryCompare)
      If lValue2 > 0 Then
        sTemp = LCase$(Right$(sName, Len(sName) - lValue2))
        
        If sTemp = "ncc.html" Then
          fncAddCandidateToArray TYPE_SINGLE_NCC, Data.Files.Item(lValue), True
        ElseIf sTemp = "master.smil" Then
          fncAddCandidateToArray TYPE_SINGLE_MSMIL, Data.Files.Item(lValue), True
        ElseIf sTemp = "discinfo.html" Then
          fncAddCandidateToArray TYPE_SINGLE_DISCINFO, Data.Files.Item(lValue), True
        ElseIf Right$(sTemp, 4) = "smil" Then
          fncAddCandidateToArray TYPE_SINGLE_SMIL, Data.Files.Item(lValue), True
        ElseIf Right$(sTemp, 4) = "html" Or Right$(sTemp, 3) = "htm" Then
          fncAddCandidateToArray TYPE_SINGLE_CONTENTDOC, Data.Files.Item(lValue), True
        End If
          
      End If
    End If
    
    DoEvents
  Next lValue
  
  fncUpdateCandidateList
End Sub

Private Sub txErrView_GotFocus()
    With txErrView
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
