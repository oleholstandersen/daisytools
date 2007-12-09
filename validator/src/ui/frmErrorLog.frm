VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmErrorLog 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Candidate Session Error Log"
   ClientHeight    =   4995
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6555
   ForeColor       =   &H0000FF00&
   Icon            =   "frmErrorLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtErrorCount 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   4935
   End
   Begin RichTextLib.RichTextBox rtfErrorLog 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6800
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmErrorLog.frx":0442
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Errors Logged:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveErrorLog 
         Caption         =   "&Save Error Log"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmErrorLog"
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

Private Sub mnuClose_Click()
    Me.Hide
End Sub

Private Sub mnuSaveErrorLog_Click()
    fncSaveFile rtfErrorLog.Text, eventlog
End Sub
