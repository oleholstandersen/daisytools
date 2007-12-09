VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Daisy 2.02 Regenerator"
   ClientHeight    =   3345
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4215
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2308.779
   ScaleMode       =   0  'User
   ScaleWidth      =   3958.103
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2880
      TabIndex        =   0
      Top             =   2880
      Width           =   1260
   End
   Begin VB.Label lblUiVersion 
      Caption         =   "Interface version:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblDllVersion 
      Caption         =   "Engine version:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout.frx":0442
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "This program is licensed under the GPL license. Please refer to the supplied textfile or www.gnu.org for more information."
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      Caption         =   "Daisy 2.02 Regenerator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
' Daisy 2.02 Regenerator Batch UI
' Copyright (C) 2003 Daisy Consortium
'
'    This file is part of Daisy 2.02 Regenerator.
'
'    Daisy 2.02 Regenerator is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Daisy 2.02 Regenerator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Daisy 2.02 Regenerator; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


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
 On Error GoTo errhandler
 lblUiVersion.Caption = "Interface version: " & App.Major & "." & App.Minor & "." & App.Revision
 lblDllVersion.Caption = "Engine version: " & objRegeneratorUserControl.fncGetDllVersion
errhandler:
End Sub
