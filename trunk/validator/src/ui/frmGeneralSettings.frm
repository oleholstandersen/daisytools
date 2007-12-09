VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGeneralSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Settings"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmGeneralSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame objFrameATF 
      Caption         =   "Allowed time fluctuation"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txTimeFluct 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   1080
         Width           =   615
      End
      Begin MSComctlLib.Slider slTimeFluct 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   1500
         TickFrequency   =   25
      End
      Begin VB.Label lbl0ms 
         Caption         =   "0 ms"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lbl100ms 
         Caption         =   "1500 ms"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Current value"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.CheckBox objChkAdvancedADTDInfo 
      Caption         =   "Show advanced ADTD information"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3855
   End
   Begin VB.CommandButton cmClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmGeneralSettings"
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

Private Sub cmClose_Click()
  If objChkAdvancedADTDInfo.Value = vbChecked Then bolADTDAdvanced = True Else _
    bolADTDAdvanced = False
  
  fncSetAdvancedADTD bolADTDAdvanced
  
  Unload Me
End Sub

Private Sub Form_Load()
  slTimeFluct.Value = lTimeFluct
  If bolADTDAdvanced Then objChkAdvancedADTDInfo.Value = vbChecked Else _
    objChkAdvancedADTDInfo.Value = vbUnchecked
End Sub

Private Sub slTimeFluct_Change()
  lTimeFluct = slTimeFluct.Value
  fncSetTimeSpan lTimeFluct
  txTimeFluct.Text = lTimeFluct
End Sub
