VERSION 5.00
Begin VB.Form frmDocuPaths 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Paths"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "frmDocuPaths.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txDefRepPath 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox txTempPath 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   6015
   End
   Begin VB.TextBox txDtdPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label laDefRepPath 
      Caption         =   "Default report save path"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   6015
   End
   Begin VB.Label laTempPath 
      Caption         =   "Temporary path"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label laDtdPath 
      Caption         =   "DTD / ADTD Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmDocuPaths"
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

Private Sub cmCancel_Click()
  Unload Me
End Sub

Private Sub cmOk_Click()
  sDefaultReportPath = txDefRepPath.Text
  
  fncSetDTDPath txDtdPath.Text
  fncSetAdtdPath txDtdPath.Text
  fncSetTempPath txTempPath.Text
  fncSaveToRegistry

  Unload Me
End Sub

Private Sub Form_Load()
  txDefRepPath.Text = sDefaultReportPath
  txDtdPath.Text = propDTDPath
  txTempPath.Text = propTempPath
End Sub
