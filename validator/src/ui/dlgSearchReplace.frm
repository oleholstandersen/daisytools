VERSION 5.00
Begin VB.Form dlgSearchReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Replace"
   ClientHeight    =   2580
   ClientLeft      =   6360
   ClientTop       =   7335
   ClientWidth     =   6780
   Icon            =   "dlgSearchReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGoToBeginning 
      Caption         =   "reset cursor"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Set the cursor at beginning of file"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "undo"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox chkbxMatchCase 
      Caption         =   "match case"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtBoxreplaceStatusBar 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "replace dialog status bar"
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox txtBoxReplaceWith 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "replace with"
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtBoxSearchFor 
         Height          =   375
         Left            =   1320
         TabIndex        =   0
         ToolTipText     =   "search for"
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "replace all"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "replace"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "search"
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "close"
         Default         =   -1  'True
         Height          =   375
         Left            =   5400
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "search for"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "replace with"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
   End
End
Attribute VB_Name = "dlgSearchReplace"
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

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1

Dim sUndoBuffer As String

Private Sub Form_Activate()
    If Len(frmMain.rtfEdit.SelText) > 2 And Len(frmMain.rtfEdit.SelText) < 10 Then txtBoxSearchFor.Text = frmMain.rtfEdit.SelText
    If Me.Visible Then txtBoxSearchFor.SetFocus
    sUndoBuffer = frmMain.rtfEdit.Text
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdUndo_Click()
    frmMain.rtfEdit.Text = sUndoBuffer
    With txtBoxreplaceStatusBar
        .Text = "Last edit undone."
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cmdSearch_Click()
    
    With frmMain.rtfEdit 'if text is highlighted, set cursor at end of highlight
        If .SelLength > 0 Then
            .SelStart = .SelStart + .SelLength
            .SelLength = 0
        End If
    End With
    
    With txtBoxreplaceStatusBar
        If RunSearch(txtBoxSearchFor.Text) Then
            .Text = "'" & txtBoxSearchFor.Text & "' found at line " & frmMain.rtfEdit.GetLineFromChar(frmMain.rtfEdit.SelStart) + 1
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        Else
            .Text = "'" & txtBoxSearchFor.Text & "' not found."
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

Private Function RunSearch(isStringToFind As String) As Boolean
    RunSearch = False
        
    With frmMain.rtfEdit
        If chkbxMatchCase Then
            If .Find(isStringToFind, .SelStart + .SelLength, , 4) > -1 Then RunSearch = True
        Else
            If .Find(isStringToFind, .SelStart + .SelLength) > -1 Then RunSearch = True
        End If
    End With
End Function

Private Sub cmdReplace_Click()
    With frmMain.rtfEdit
        If .SelText = txtBoxSearchFor.Text Then ' if rtf seltext is focused on text to search and replace
            sUndoBuffer = frmMain.rtfEdit.Text
            .SelText = txtBoxReplaceWith.Text
            With txtBoxreplaceStatusBar
                .Text = "'" & txtBoxSearchFor.Text & "' replaced at line " & frmMain.rtfEdit.GetLineFromChar(frmMain.rtfEdit.SelStart) + 1
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        Else                                    ' if its not focused on text to search and replace
            If RunSearch(txtBoxSearchFor.Text) Then
                cmdReplace_Click
            Else
                With txtBoxreplaceStatusBar
                    .Text = "'" & txtBoxSearchFor.Text & "' not found."
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
            End If
        End If
    End With
End Sub

Private Sub cmdGoToBeginning_Click()
    frmMain.rtfEdit.SelStart = 0
    frmMain.rtfEdit.SelLength = 0
    With txtBoxreplaceStatusBar
          .Text = "Cursor set at beginning of document."
          .SetFocus
          .SelStart = 0
          .SelLength = Len(.Text)
    End With
End Sub

Private Sub cmdReplaceAll_Click()
Dim lCount As Long, bDontCount As Boolean

        bDontCount = False
        sUndoBuffer = frmMain.rtfEdit.Text
        If chkbxMatchCase Then
            If InStr(1, frmMain.rtfEdit.Text, txtBoxSearchFor, vbBinaryCompare) = 0 Then
                lCount = 0
            Else
                If LCase$(txtBoxSearchFor.Text) = LCase$(txtBoxReplaceWith.Text) Then
                  frmMain.rtfEdit.Text = Replace(frmMain.rtfEdit.Text, txtBoxSearchFor.Text, txtBoxReplaceWith.Text, , -1, vbBinaryCompare)
                  bDontCount = True
                Else
                  Do
                    frmMain.rtfEdit.Text = Replace(frmMain.rtfEdit.Text, txtBoxSearchFor.Text, txtBoxReplaceWith.Text, , 1, vbTextCompare)
                    lCount = lCount + 1
                  Loop Until InStr(1, frmMain.rtfEdit.Text, txtBoxSearchFor, vbBinaryCompare) = 0
                End If
            End If
        Else
            If InStr(1, frmMain.rtfEdit.Text, txtBoxSearchFor, vbTextCompare) = 0 Then
                lCount = 0
            Else
                If LCase$(txtBoxSearchFor.Text) = LCase$(txtBoxReplaceWith.Text) Then
                  frmMain.rtfEdit.Text = Replace(frmMain.rtfEdit.Text, txtBoxSearchFor.Text, txtBoxReplaceWith.Text, , -1, vbTextCompare)
                  bDontCount = True
                Else
                  'have to do the above since if doing the below means eternal loop
                  Do
                    frmMain.rtfEdit.Text = Replace(frmMain.rtfEdit.Text, txtBoxSearchFor.Text, txtBoxReplaceWith.Text, , 1, vbTextCompare)
                    lCount = lCount + 1
                  Loop Until InStr(1, frmMain.rtfEdit.Text, txtBoxSearchFor, vbTextCompare) = 0
                End If
            End If
        End If
        
    With txtBoxreplaceStatusBar
         If bDontCount Then
          .Text = "'" & txtBoxSearchFor.Text & "' replace done."
         Else
          .Text = "'" & txtBoxSearchFor.Text & "' replaced " & lCount & " times."
         End If
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Form_Load()
  SetWindowPos hwnd, HWND_TOPMOST, Left / Screen.TwipsPerPixelX, _
    Top / Screen.TwipsPerPixelY, Width / Screen.TwipsPerPixelX, _
    Height / Screen.TwipsPerPixelY, 0
End Sub
