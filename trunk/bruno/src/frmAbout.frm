VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Bruno"
   ClientHeight    =   2355
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4215
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1625.463
   ScaleMode       =   0  'User
   ScaleWidth      =   3958.103
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2760
      TabIndex        =   0
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "This program is licensed under the LGPL license. Please refer to the supplied textfile or www.gnu.org for more information."
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      Caption         =   "Bruno [Daisy Fileset Generator]"
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
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
 On Error GoTo errhandler
 Label1.Visible = False
 lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
errhandler:
End Sub
