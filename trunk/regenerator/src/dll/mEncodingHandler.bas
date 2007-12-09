Attribute VB_Name = "mEncodingHandler"
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

Public Function fncSetEncodingName( _
    ByVal lInputEncoding As Long, _
    ByRef sOutCharsetName As String, _
    ByVal sWantsThisEncoding As String, _
    ByRef objOwner As oRegenerator) As Boolean

  'set the output charset encoding name string
  fncSetEncodingName = True
  Select Case lInputEncoding
    Case CHARSET_WESTERN
      sOutCharsetName = "utf-8"
    Case CHARSET_UTF8
     'mg20030921: do this to be able to take books away from utf-8
     If (LCase$(Trim$(sWantsThisEncoding)) = "windows-1252") Then
      sOutCharsetName = "windows-1252"
     Else
       sOutCharsetName = "utf-8"
     End If
    Case CHARSET_SHIFTJIS
      sOutCharsetName = "Shift_JIS" 'preferred mime name
    Case CHARSET_BIG5
      sOutCharsetName = "Big5" 'preferred mime name
    Case CHARSET_SPECIAL
      sOutCharsetName = Trim$(sWantsThisEncoding)
      If sOutCharsetName = "" Then objOwner.addlog "<error in='fncSetEncodingName'>zerolength special characterset name</error>"
    Case Else
      objOwner.addlog "<error in='fncSetEncodingName'>unrecognized charset type encountered in fncRunTransform selectcase lInputEncoding</error>"
  End Select

End Function

Public Function fncCheckUtf8Encoding( _
    ByRef lInputEncoding As Long, _
    ByVal sNccPath As String, _
    ByRef objOwner As oRegenerator _
    ) As Boolean
Dim sNcc As String
Dim lActualEncoding As Long

  On Error GoTo ErrHandler
  fncCheckUtf8Encoding = False
  
  sNcc = fncReadFile(sNccPath)

  If sNcc <> "" Then
    'trunc to prolog and head only
    sNcc = Mid(sNcc, 1, InStr(1, sNcc, "<body", vbTextCompare))
   'look for indications that this might be utf-8
    If sNcc <> "" Then
      If InStr(1, sNcc, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>", vbTextCompare) > 1 _
      Or InStr(1, sNcc, "<?xml version=" & Chr(39) & "1.0" & Chr(39) & "?>", vbTextCompare) > 1 _
      Or InStr(1, sNcc, "utf-8", vbTextCompare) > 1 Then
        lActualEncoding = CHARSET_UTF8
        If lInputEncoding <> lActualEncoding Then
          objOwner.addlog "<warning>warning: input ncc indicates this is utf-8 encoded. Input charset changed to utf-8 for all files. Check the integrity of the output ncc/content/smil files manually in your browser.</warning>"
          lInputEncoding = lActualEncoding
        End If
      End If
    Else
      'file is weird (but still pre tidy), do nothing, dont change lInputEncoding
    End If 'ncc <> ""
    Else
    objOwner.addlog "<error in='fncCheckUtf8Encoding'>sNcc is empty in fncCheckUtf8Encoding</error>"
  End If 'sNcc <> ""
  
  fncCheckUtf8Encoding = True

ErrHandler:
 If Not fncCheckUtf8Encoding Then objOwner.addlog "<errH in='fncCheckUtf8Encoding' file='ncc'>fncCheckUtf8Encoding ErrH</errH>"
End Function


