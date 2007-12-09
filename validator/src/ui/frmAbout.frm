VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Daisy 2.02 validator"
   ClientHeight    =   2445
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4215
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1687.583
   ScaleMode       =   0  'User
   ScaleWidth      =   3958.103
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label lblExeVersion 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblDllVersion 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0442
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      Caption         =   "Daisy 2.02 Validator"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
 On Error GoTo Errhandler
 lblTitle.Caption = "Daisy 2.02 DTB Validator"
 lblDllVersion.Caption = "Engine Version: " & fncGetDllVersion
 lblExeVersion.Caption = "Interface Version: " & App.Major & "." & App.Minor & "." & App.Revision
Errhandler:
End Sub

