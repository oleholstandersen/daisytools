VERSION 5.00
Begin VB.Form frmApplySettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apply settings"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmApplySettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton objCmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton objCmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame objFrameApplySettings 
      Caption         =   "Apply settings"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CheckBox objChkMoveBookLocation 
         Caption         =   "Move book location"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CheckBox objChkMoveBookOption 
         Caption         =   "Move book option"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox objchkSaveRegeneratedOption 
         Caption         =   "Save regenerated book option"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox objChkPrefix 
         Caption         =   "Prefix"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox objChkUseNumeric 
         Caption         =   "Use numeric portion of id"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.CheckBox objChkSequentialRename 
         Caption         =   "Sequential rename"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox objChkMetadataLocation 
         Caption         =   "Metadata location"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox objChkMetadataImportOption 
         Caption         =   "Metadata import option"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox objChkCharacterSet 
         Caption         =   "Character set"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox objChkDtbType 
         Caption         =   "Dtb type"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmApplySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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



Option Explicit

Private Sub objCmdCancel_Click()
  Unload Me
End Sub

Private Sub objCmdOk_Click()
Dim i As Long
  'For lCurrentJob = 1 To lJobCount
  For i = 1 To lJobCount
    'With aJobItems(lCurrentJob)
    With aJobItems(i)
      If (objChkDtbType.Value = vbChecked) Then .lDtbType = lDtbType
      If (objChkCharacterSet.Value = vbChecked) Then
        .lInputCharset = lCharset
        .lCharsetOther = lIANACharset
      End If
      If (objChkMetadataImportOption.Value = vbChecked) Then _
        .bolPreserveMeta = bolPreserveMeta
      If (objChkMetadataLocation.Value = vbChecked) Then .sMetaImport = sMetaFile
      If (objChkSequentialRename.Value = vbChecked) Then .bolSeqRename = bolSeqRename
      If (objChkUseNumeric.Value = vbChecked) Then .bolUseNumeric = bolUseNumeric
      If (objChkPrefix.Value = vbChecked) Then .sPrefix = sPrefix
      If (objchkSaveRegeneratedOption.Value = vbChecked) Then .bolSaveSame = bolSameFolder
      If (objChkMoveBookOption.Value = vbChecked) Then .bolMoveBook = bolMoveBook
      If (objChkMoveBookLocation.Value = vbChecked) Then .sNewFolder = sSavePath
    End With
  'Next lCurrentJob
  Next i

  Unload Me
End Sub

Private Sub Form_Load()
  Dim bolTemp As Boolean
  
  sBaseKey = "Software\DaisyWare\Regenerator\Apply Settings"
  
  fncLoadRegistryData "Dtb type", bolTemp, HKEY_CURRENT_USER, , True
  objChkDtbType.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Character set", bolTemp, HKEY_CURRENT_USER, , True
  objChkCharacterSet.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Metadata import option", bolTemp, HKEY_CURRENT_USER, , True
  objChkMetadataImportOption.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Metadata location", bolTemp, HKEY_CURRENT_USER, , True
  objChkMetadataLocation.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Sequential rename", bolTemp, HKEY_CURRENT_USER, , True
  objChkSequentialRename.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Use numeric portion of id", bolTemp, HKEY_CURRENT_USER, , True
  objChkUseNumeric.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Prefix", bolTemp, HKEY_CURRENT_USER, , True
  objChkPrefix.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Save regenerated book option", bolTemp, HKEY_CURRENT_USER, , True
  objchkSaveRegeneratedOption.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Move book option", bolTemp, HKEY_CURRENT_USER, , True
  objChkMoveBookOption.Value = fncBol2Check(bolTemp)
  fncLoadRegistryData "Move book location", bolTemp, HKEY_CURRENT_USER, , True
  objChkMoveBookLocation.Value = fncBol2Check(bolTemp)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim bolTemp As Boolean
  
  sBaseKey = "Software\DaisyWare\Regenerator\Apply Settings"
  
  fncSaveRegistryData "Dtb type", fncCheck2Bol(objChkDtbType), HKEY_CURRENT_USER
  fncSaveRegistryData "Character set", fncCheck2Bol(objChkCharacterSet), HKEY_CURRENT_USER
  fncSaveRegistryData "Metadata import option", _
    fncCheck2Bol(objChkMetadataImportOption), HKEY_CURRENT_USER
  fncSaveRegistryData "Metadata location", fncCheck2Bol(objChkMetadataLocation), HKEY_CURRENT_USER
  fncSaveRegistryData "Sequential rename", fncCheck2Bol(objChkSequentialRename), HKEY_CURRENT_USER
  fncSaveRegistryData "Use numeric portion of id", fncCheck2Bol(objChkUseNumeric), HKEY_CURRENT_USER
  fncSaveRegistryData "Prefix", fncCheck2Bol(objChkPrefix), HKEY_CURRENT_USER
  fncSaveRegistryData "Save regenerated book option", _
    fncCheck2Bol(objchkSaveRegeneratedOption), HKEY_CURRENT_USER
  fncSaveRegistryData "Move book option", fncCheck2Bol(objChkMoveBookOption), HKEY_CURRENT_USER
  fncSaveRegistryData "Move book location", fncCheck2Bol(objChkMoveBookLocation), HKEY_CURRENT_USER
End Sub
