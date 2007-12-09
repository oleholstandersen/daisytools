VERSION 5.00
Begin VB.Form frmOnOffLine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Online/Offline Settings"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmHttpTest 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   2640
   End
   Begin VB.Frame frmOnOffLineSettings 
      Caption         =   "Proxy settings"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      Begin VB.TextBox txProxySettings 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "ip:port"
         Top             =   720
         Width           =   4215
      End
      Begin VB.CheckBox chUseProxy 
         Caption         =   "Use proxy"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame frDtdAccess 
      Caption         =   "DTD access"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton raUseOnlineDtds 
         Caption         =   "Use online dtd:s via http"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4215
      End
      Begin VB.OptionButton raUseLocalDtds 
         Caption         =   "Use local dtd:s."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmOnOffLine"
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

Dim lTestState As Long
Dim bolUODTDs As Boolean, bolUP As Boolean, sP As String

Private Sub chUseProxy_Click()
  subDtdDataChanged
End Sub

Private Sub cmClose_Click()
  sP = txProxySettings.Text
  
  fncSaveRegistryData "Elcel\UseOnlineDtds", raUseOnlineDtds.Value
  
  If chUseProxy.Value = vbChecked Then
    fncSaveRegistryData "Elcel\UseProxy", True
  Else
    fncSaveRegistryData "Elcel\UseProxy", False
  End If
  
  fncSaveRegistryData "Elcel\ProxySettings", txProxySettings.Text
  
  fncSetElcelOnlineOptions bolUODTDs, bolUP, sP
  
  If raUseOnlineDtds.Value = True Then
    Dim vRes As VbMsgBoxResult, bolResult As Boolean

    Enabled = False
    tmHttpTest.Enabled = True
    bolResult = fncTestElcelOnline
    tmHttpTest.Enabled = False
    Caption = "Elcel xmlvalid options"
    Enabled = True
    
    If Not bolResult Then
      vRes = MsgBox("Http access malfunctioning! " & vbCrLf & _
        "Do you want to continue anyway?", vbYesNo)
      If vRes = vbNo Then Exit Sub
    End If
  End If
  
  Unload Me
End Sub

Private Sub Form_Load()
  Dim bolTemp As Boolean, tempsString As String
  
  sBaseKey = "software\validator"
  
  fncLoadRegistryData "Elcel\UseOnlineDtds", bolUODTDs, , , propUseOnlineDtds
  
  raUseOnlineDtds.Value = bolTemp
  raUseLocalDtds.Value = Not bolTemp
  
  fncLoadRegistryData "Elcel\UseProxy", bolUP, , , propUseProxy
  fncLoadRegistryData "Elcel\ProxySettings", sP, , , propProxy
    
  txProxySettings.Text = tempsString
    
  If bolTemp Then chUseProxy.Value = vbChecked Else _
    chUseProxy.Value = vbUnchecked
  
  fncSetElcelOnlineOptions bolUODTDs, bolUP, Trim(sP)
  subDtdDataChanged
End Sub

Private Sub raUseLocalDtds_Click()
  subDtdDataChanged
End Sub

Private Sub raUseOnlineDtds_Click()
  subDtdDataChanged
End Sub

Private Sub subDtdDataChanged()
  If raUseOnlineDtds.Value Then
    chUseProxy.Enabled = True
    txProxySettings.Enabled = True
  
    If chUseProxy.Value = vbChecked Then
      bolUP = True
      txProxySettings.Enabled = True
    Else
      bolUP = False
      txProxySettings.Enabled = False
    End If
    bolUODTDs = True
  Else
    chUseProxy.Enabled = False
    txProxySettings.Enabled = False
    bolUODTDs = False
  End If
  
  fncSetElcelOnlineOptions bolUODTDs, bolUP, sP
End Sub

Private Sub tmHttpTest_Timer()
  Select Case lTestState
    Case 0
      Caption = "Testing HTTP settings |"
    Case 1
      Caption = "Testing HTTP settings /"
    Case 2
      Caption = "Testing HTTP settings -"
    Case 3
      Caption = "Testing HTTP settings \"
  End Select
  
  Refresh
  
  lTestState = lTestState + 1
  If lTestState = 4 Then lTestState = 0
End Sub
