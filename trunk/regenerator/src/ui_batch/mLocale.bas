Attribute VB_Name = "mLocale"
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

Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const CHINESEBIG5_CHARSET = 136
Public Const CHINESESIMPLIFIED_CHARSET = 134

Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Public Sub subApplyUserLcid()
    'below should include all UI forms
    fncApplyCharsetToForm frmMain
    fncApplyCharsetToForm frmApplySettings
    fncApplyCharsetToForm frmAbout
End Sub
    
' This function is changing the charset and font on all controls in the given form
' to match the current LCID
Public Function fncApplyCharsetToForm(iForm As Form) As Boolean
  Dim colControls As Collection, templCounter As Long
  Dim sFontName As String, lCharset As Long, lSize As Long
      
  On Error GoTo ErrorSetProperFont
  Select Case GetUserDefaultLCID
  Case &H404 ' Traditional Chinese
    lCharset = CHINESEBIG5_CHARSET
    sFontName = ChrW(&H65B0) + ChrW(&H7D30) + ChrW(&H660E) + ChrW(&H9AD4)
    lSize = 9
    
  Case &H411 ' Japan
    lCharset = SHIFTJIS_CHARSET
    sFontName = ChrW(&HFF2D) + ChrW(&HFF33) + ChrW(&H20) + ChrW(&HFF30) + ChrW(&H30B4) + ChrW(&H30B7) + ChrW(&H30C3) + ChrW(&H30AF)
    lSize = 9
    
  Case &H412 'Korea UserLCID
    lCharset = HANGEUL_CHARSET
    sFontName = ChrW(&HAD74) + ChrW(&HB9BC)
    lSize = 9
    
  Case &H804 ' Simplified Chinese
    lCharset = CHINESESIMPLIFIED_CHARSET
    sFontName = ChrW(&H5B8B) + ChrW(&H4F53)
    lSize = 9
    
  Case Else   ' The other countries
    fncApplyCharsetToForm = True
    sFontName = "Arial"
    Exit Function
  End Select

  On Error Resume Next
  For templCounter = 0 To iForm.Count - 1
    With iForm.Controls.Item(templCounter).Font
      .Charset = lCharset
      .Name = sFontName
      .Size = lSize
    End With
NextItem:
  Next templCounter
  
  fncApplyCharsetToForm = True
ErrorSetProperFont:
End Function
