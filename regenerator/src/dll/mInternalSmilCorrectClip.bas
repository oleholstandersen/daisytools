Attribute VB_Name = "mInternalSmilCorrectClip"
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

Public Function fncCorrectFirstPhraseInPar( _
    ByRef oSmilDom As MSXML2.DOMDocument40, _
    ByRef objOwner As oRegenerator, _
    ByVal lCurrentArrayItem As Long, _
    ByRef lClipSpan As Long, _
    ByRef lClipLessThan As Long, _
    ByRef lFirstClipLessThan As Long, _
    ByRef lNextClipLessThan As Long _
    ) As Boolean
  
  On Error GoTo ErrHandler
    
  fncCorrectFirstPhraseInPar = False

  Dim objPar As IXMLDOMNode, objParList As IXMLDOMNodeList
  Dim objAudio As IXMLDOMNode, objAudioList As IXMLDOMNodeList
  Dim lClipTime As Long, lNextClip As Long, bolChange As Boolean
  Dim lFirstEnd As Long, lSecondStart As Long
  Dim objNode As IXMLDOMNode, sFirstSrc As String, sNextSrc As String
  
  Set objParList = oSmilDom.selectNodes("/smil/body/seq/par")
  For Each objPar In objParList
    Set objAudioList = objPar.selectNodes(".//audio")
    If objAudioList.length >= 2 Then
      bolChange = False
      Set objAudio = objAudioList.Item(0)
      
      Set objNode = objAudio.Attributes.getNamedItem("src")
      sFirstSrc = LCase$(objNode.nodeValue)
      
      Set objNode = objAudio.Attributes.getNamedItem("clip-end")
      If objNode Is Nothing Then GoTo Skip
      lClipTime = fncConvertSmilClockVal2Ms(objNode.nodeValue)
      lFirstEnd = lClipTime
      
      Set objNode = objAudio.Attributes.getNamedItem("clip-begin")
      If objNode Is Nothing Then GoTo Skip
      lClipTime = lClipTime - fncConvertSmilClockVal2Ms(objNode.nodeValue)
      
      If lClipTime < lClipLessThan Then bolChange = True
    
      Set objAudio = objAudioList.Item(1)
        
      Set objNode = objAudio.Attributes.getNamedItem("src")
      sNextSrc = LCase$(objNode.nodeValue)
        
      Set objNode = objAudio.Attributes.getNamedItem("clip-end")
      If objNode Is Nothing Then GoTo Skip
      lNextClip = fncConvertSmilClockVal2Ms(objNode.nodeValue)
      
      Set objNode = objAudio.Attributes.getNamedItem("clip-begin")
      If objNode Is Nothing Then GoTo Skip
      lSecondStart = fncConvertSmilClockVal2Ms(objNode.nodeValue)
      lNextClip = lNextClip - lSecondStart
      
      If Not bolChange Then
        If lClipTime < lFirstClipLessThan And _
          lNextClip < lNextClipLessThan Then bolChange = True
      End If
      
      If (Not sFirstSrc = sNextSrc) _
        Or (lFirstEnd < lSecondStart - lClipSpan) Then
          GoTo Skip
      End If
      
      If bolChange Then
        Dim sBeginVal As String
        sBeginVal = objAudioList.Item(0).Attributes.getNamedItem("clip-begin").nodeValue
        Set objNode = objAudioList.Item(1).Attributes.getNamedItem("clip-begin")
        objNode.nodeValue = sBeginVal
        Set objAudio = objAudioList.Item(0)
        objAudio.parentNode.removeChild objAudio
'        Set objNode = oSmilDom.createAttribute("clip-begin")
'        objNode.nodeValue = objAudioList.Item(0).Attributes.getNamedItem("clip-begin").nodeValue
'        objAudio.Attributes.removeNamedItem ("clip-begin")
'        objAudio.Attributes.setNamedItem objNode
'        Set objAudio = objAudioList.Item(0)
'        objAudio.parentNode.removeChild objAudio
        objOwner.addlog "<message>merged first audio phrase in " & objOwner.objFileSetHandler.aOutFileSet(lCurrentArrayItem).sFileName & "</message>"
      End If
Skip:
    End If 'objAudioList.length >= 2
  Next objPar
  
  fncCorrectFirstPhraseInPar = True
ErrHandler:
  If Not fncCorrectFirstPhraseInPar Then objOwner.addlog "<errH in='fncCorrectFirstPhraseInPar' arrayItem='" & lCurrentArrayItem & "'>fncCorrectFirstPhraseInPar ErrH</errH>"
End Function

'----------------------------------------------------------------------
' Name:         fncCorrectLastClipTimeInSmil
' Copyright:    © 2002 DAISY consortium / Talboks- och punktskriftsbiblioteket
' Author:       DAISY consortium / Talboks- och punktskriftsbiblioteket
' Created:      2002-07-12 @ 14:39:54 (Vers: 0.4.0049)
'
' Description:  For each smil dom object sent in, calls
'               Getmp3Length for the mp3 file referenced in the
'               last node that contains a clip time value. Replaces
'               the clip-end value if the current value is outside
'               the playtime of the mp3 file.
'
'               Wastes with CPU presently as values for the same file
'               will be calculated multiple times
'----------------------------------------------------------------------

'the above string may be used by fncFixMultiTextInPar
'if it finds dupe text in par
'it inserts this string as audio src in new par
'should be of same format as other audio in smilfile
'currently assuming mono
'currently assuming always mp3
'currently assuming all mp3files of dtb use same format

Public Function fncCorrectLastClipTimeInSmil( _
    xmlDoc As MSXML2.DOMDocument40, _
    i As Long, _
    objOwner As oRegenerator _
    ) As Boolean

Dim k As Integer
Dim oNode As IXMLDOMNode
Dim oNodes As IXMLDOMNodeList
Dim oNodeValue As IXMLDOMNode
Dim oNodeValueClipBegin As IXMLDOMNode
Dim mp3filename As String
Dim mp3playtime As Double
Dim orgmp3playtime As String
Dim mp3starttime As String

  On Error GoTo ErrHandler
  mp3filename = ""
  
  'Find the last audio element in the smile file
  'and then get the file name of the mp3 file
  Set oNodes = xmlDoc.selectNodes("//audio[last()]")
  Set oNode = oNodes.Item(oNodes.length - 1)
            
  If Not oNode Is Nothing Then
    Set oNodeValue = oNode.Attributes.getNamedItem("src")
    If Not oNodeValue Is Nothing Then
      mp3filename = oNodeValue.Text
      '**changed If fncFileExists(Left$(aInFileSet(i).sAbsPath, Len(aInFileSet(i).sAbsPath) - Len(aOutFileSet(i).sFileName)) & mp3filename) And Right$(Trim$(mp3filename), 4) = ".mp3" Then
      If (fncFileExists(objOwner.sDtbFolderPath & oNodeValue.Text, objOwner)) And (Right$(Trim$(LCase$(mp3filename)), 4) = ".mp3") Then
        'Get the current clip-end/-begin values
        Set oNodeValue = oNode.Attributes.getNamedItem("clip-end")
        Set oNodeValueClipBegin = oNode.Attributes.getNamedItem("clip-begin")
        'Get the calculated playtime of the referenced mp3 file
        '**changed ** mp3playtime = Getmp3Length(Left$(aInFileSet(i).sAbsPath, Len(aInFileSet(i).sAbsPath) - Len(aOutFileSet(i).sFileName)) & mp3filename)
        mp3playtime = Getmp3Length(objOwner.sDtbFolderPath & mp3filename, objOwner)
        
        'Strip the characters from the clip-end attribute value
        orgmp3playtime = Right$(Trim$(oNodeValue.Text), Len(oNodeValue.Text) - 4)
        orgmp3playtime = Trim$(Left$(Trim$(orgmp3playtime), Len(orgmp3playtime) - 1))
        mp3starttime = Right$(Trim$(oNodeValueClipBegin.Text), Len(oNodeValueClipBegin.Text) - 4)
        mp3starttime = Trim$(Left$(Trim$(mp3starttime), Len(mp3starttime) - 1))
        If mp3playtime <> 0 And IsNumeric(orgmp3playtime) And IsNumeric(mp3starttime) Then
          'We only change the value if the current value points outside the referenced mp3 file
          'AND the difference is less than 1.5 seconds AND of course the new value can't really
          'be before the start-time...
          If Val(orgmp3playtime) > mp3playtime Then
            'If Val(orgmp3playtime) - mp3playtime < 1.5 Then
            If Val(mp3starttime) <= mp3playtime Then
              objOwner.addlog ("<message>clip-end value outside mp3 file, changing from " & orgmp3playtime & " to " & mp3playtime & "</message>")
              oNodeValue.Text = "npt=" & Trim$(Str(mp3playtime)) & "s"
            Else
              objOwner.addlog ("<error in='fncCorrectLastClipTimeInSmil'>Error replacing clip-end value for " & mp3filename & " new value would be before start-time</error>")
            End If 'Val(mp3starttime) <= mp3playtime
            'Else
              'objowner.addlog ("Error replacing clip-end value for " & mp3filename & " clip-time difference larger than 1.5 seconds" & vbCrLf & "Old value: " & orgmp3playtime & vbCrLf & "Recommended value: " & mp3playtime)
            'End If 'Val(orgmp3playtime) - mp3playtime < 1.5 Then
            End If 'mp3playtime <> 0 And IsNumeric(orgmp3playtime) And
        Else
           objOwner.addlog ("<error in='fncCorrectLastClipTimeInSmil'>Error calculating or obtaining clip-end value for " & mp3filename & "</error>")
        End If 'fncFileExists(Left$(aInFileSet
      Else
        'objOwner.addlog ("Could not find mp3 file: " & mp3filename)
      End If 'Not oNodeValue Is Nothing
    Else
      objOwner.addlog ("<error in='fncCorrectLastClipTimeInSmil'>Could not find src attribute in last audio element</error>")
    End If
  Else
      objOwner.addlog ("<error in='fncCorrectLastClipTimeInSmil'>No audio element in smil file</error>")
  End If 'Not oNode Is Nothing

  Set oNode = Nothing
  fncCorrectLastClipTimeInSmil = True

ErrHandler:
    If Err.Number <> 0 Then fncCorrectLastClipTimeInSmil = False
    If Not fncCorrectLastClipTimeInSmil Then objOwner.addlog "<errH in='fncCorrectLastClipTimeInSmil'>fncCorrectLastClipTimeInSmil ErrH</errH>"
End Function

'----------------------------------------------------------------------
' Name:         Getmp3Length
' Copyright:    © 2002 DAISY consortium / Talboks- och punktskriftsbiblioteket
' Author:       DAISY consortium / Talboks- och punktskriftsbiblioteket
' Created:      2002-07-12 @ 14:39:54 (Vers: 0.4.0049)
'
' Description:  Returns the playtime in seconds to three decimal points
'               for the given mp3 file.
'               The playtime is calculated as filesize/bitrate
'----------------------------------------------------------------------
Public Function Getmp3Length(mp3name As String, objOwner As oRegenerator) As Double

    Dim fso As Object
    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(mp3name) Then
        Getmp3Length = CleanAndConvertDecimalNumber(fso.GetFile(mp3name).Size * 8 / GetBitRate(mp3name, objOwner))
    Else
        Getmp3Length = 0
    End If

ErrHandler:
    If Err.Number <> 0 Then
        Getmp3Length = 0
    End If
End Function

'----------------------------------------------------------------------
' Name:         CleanAndConvertDecimalNumber
' Copyright:    © 2002 DAISY consortium / Talboks- och punktskriftsbiblioteket
' Author:       DAISY consortium / Talboks- och punktskriftsbiblioteket
' Created:      2002-07-12 @ 14:39:54 (Vers: 0.4.0049)
'
' Description:  Takes a double and returns a number
'               with three decimals
'----------------------------------------------------------------------
Private Function CleanAndConvertDecimalNumber(nbrtoclean As Double) As Double
Dim valuearr() As String
Dim integerpart As String
Dim decimalpart As String

    On Error GoTo ErrHandler
    If nbrtoclean <> 0 Then
        If InStr(1, Trim$(Str(nbrtoclean)), ".", vbTextCompare) <> 0 Then
            'The full stop intead of the comma in the split is because of the str function
            valuearr = Split(Str(nbrtoclean), ".", -1, vbTextCompare)
            integerpart = valuearr(0)
            decimalpart = valuearr(1)
            If Len(decimalpart) > 3 Then
                decimalpart = Left$(decimalpart, 3)
            Else
                While Len(decimalpart) < 3
                    decimalpart = decimalpart & "0"
                Wend
            End If
            'Sub second clip
            If integerpart = "" Or integerpart = " " Then integerpart = 0
            CleanAndConvertDecimalNumber = Val(integerpart & "." & decimalpart)
        Else
            CleanAndConvertDecimalNumber = nbrtoclean
        End If
    Else
        CleanAndConvertDecimalNumber = 0
    End If

ErrHandler:

    If Err.Number <> 0 Then
        CleanAndConvertDecimalNumber = ""
    End If

End Function

'----------------------------------------------------------------------
' Name:         GetBitRate
' Copyright:    © 2002 DAISY consortium / Talboks- och punktskriftsbiblioteket
' Author:       DAISY consortium / Talboks- och punktskriftsbiblioteket
' Created:      2002-07-12 @ 14:39:54 (Vers: 0.4.0049)
'
' Description:  Calculates the bitrate for the given mp3file.
'               Reads the first 4 bytes in the mp3 file and uses the
'               bit values to determine the bitrate by combining
'               the values for Bitrate, Mpeg version and layer
'----------------------------------------------------------------------
Public Function GetBitRate(mp3name As String, objOwner As oRegenerator) As Long
Dim fnum As Long
Dim getmpegbyte As Byte
Dim getbitAndSamplerate As Byte
Dim mpegversion As String
Dim bitrate As String
Dim layer As String
Dim samplerate As String
    
    On Error GoTo ErrHandler
    fnum = FreeFile
    Open mp3name For Binary As #fnum
    
    'skip the first byte
    Get #fnum, , getmpegbyte
    Get #fnum, , getmpegbyte
    Get #fnum, , getbitAndSamplerate
    Close #fnum
    mpegversion = GetMpegVersion(Mid$(decimaltobinary(getmpegbyte), 4, 2))
    'Not neccessary to get bit rate
    'samplerate = GetSampleRate(Mid$(decimaltobinary(getbitAndSamplerate), 5, 2), mpegversion)
    layer = GetLayerValue(Mid$(decimaltobinary(getmpegbyte), 6, 2))
    bitrate = GetBitValue(Left$(decimaltobinary(getbitAndSamplerate), 4), layer, mpegversion)
    GetBitRate = CLng(bitrate) * 1000
       
    'use these values to determine sEmptyMp3FileName
    'objowner.addlog mpegversion & " " & layer & " " & GetBitRate
       Select Case GetBitRate
         Case 16000
           objOwner.sEmptyMp3Filename = "rgn_empty_16.mp3"
         Case 24000
           objOwner.sEmptyMp3Filename = "rgn_empty_24.mp3"
         Case 32000
           objOwner.sEmptyMp3Filename = "rgn_empty_32.mp3"
         Case 48000
           objOwner.sEmptyMp3Filename = "rgn_empty_48.mp3"
         Case 56000
           objOwner.sEmptyMp3Filename = "rgn_empty_56.mp3"
         Case 64000
           objOwner.sEmptyMp3Filename = "rgn_empty_64.mp3"
         Case 96000
           objOwner.sEmptyMp3Filename = "rgn_empty_96.mp3"
         Case 128000
           objOwner.sEmptyMp3Filename = "rgn_empty_128.mp3"
         Case Else
           'objowner.addlog "found audio encoding [" & mpegversion & " " & layer & " " & GetBitRate & "]not existing in audiofile insert set; will use mp3 64 kbps if inserting par"
           objOwner.sEmptyMp3Filename = "rgn_empty_64.mp3"
       End Select
     
ErrHandler:

    If Err.Number <> 0 Then
        GetBitRate = 0
    End If

End Function

Private Function GetBitValue(bitratevalue As String, LayerValue As String, mpegversion As String) As String

    On Error GoTo ErrHandler

    Select Case mpegversion
        Case "1"
            Select Case LayerValue
                Case "1"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "32"
                        Case "0010"
                            GetBitValue = "64"
                        Case "0011"
                            GetBitValue = "96"
                        Case "0100"
                            GetBitValue = "128"
                        Case "0101"
                            GetBitValue = "160"
                        Case "0110"
                            GetBitValue = "192"
                        Case "0111"
                            GetBitValue = "224"
                        Case "1000"
                            GetBitValue = "256"
                        Case "1001"
                            GetBitValue = "288"
                        Case "1010"
                            GetBitValue = "320"
                        Case "1011"
                            GetBitValue = "352"
                        Case "1100"
                            GetBitValue = "384"
                        Case "1101"
                            GetBitValue = "416"
                        Case "1110"
                            GetBitValue = "448"
                        Case Else
                            GetBitValue = ""
                    End Select
                Case "2"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "32"
                        Case "0010"
                            GetBitValue = "48"
                        Case "0011"
                            GetBitValue = "56"
                        Case "0100"
                            GetBitValue = "64"
                        Case "0101"
                            GetBitValue = "80"
                        Case "0110"
                            GetBitValue = "96"
                        Case "0111"
                            GetBitValue = "112"
                        Case "1000"
                            GetBitValue = "128"
                        Case "1001"
                            GetBitValue = "160"
                        Case "1010"
                            GetBitValue = "192"
                        Case "1011"
                            GetBitValue = "224"
                        Case "1100"
                            GetBitValue = "256"
                        Case "1101"
                            GetBitValue = "320"
                        Case "1110"
                            GetBitValue = "384"
                        Case Else
                            GetBitValue = ""
                    End Select
                Case "3"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "32"
                        Case "0010"
                            GetBitValue = "40"
                        Case "0011"
                            GetBitValue = "48"
                        Case "0100"
                            GetBitValue = "56"
                        Case "0101"
                            GetBitValue = "64"
                        Case "0110"
                            GetBitValue = "80"
                        Case "0111"
                            GetBitValue = "96"
                        Case "1000"
                            GetBitValue = "112"
                        Case "1001"
                            GetBitValue = "128"
                        Case "1010"
                            GetBitValue = "160"
                        Case "1011"
                            GetBitValue = "192"
                        Case "1100"
                            GetBitValue = "224"
                        Case "1101"
                            GetBitValue = "256"
                        Case "1110"
                            GetBitValue = "320"
                        Case Else
                            GetBitValue = ""
                    End Select
                Case Else
                    GetBitValue = ""
            End Select
        Case "2"
            Select Case LayerValue
                Case "1"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "32"
                        Case "0010"
                            GetBitValue = "48"
                        Case "0011"
                            GetBitValue = "56"
                        Case "0100"
                            GetBitValue = "64"
                        Case "0101"
                            GetBitValue = "80"
                        Case "0110"
                            GetBitValue = "96"
                        Case "0111"
                            GetBitValue = "112"
                        Case "1000"
                            GetBitValue = "128"
                        Case "1001"
                            GetBitValue = "144"
                        Case "1010"
                            GetBitValue = "160"
                        Case "1011"
                            GetBitValue = "176"
                        Case "1100"
                            GetBitValue = "192"
                        Case "1101"
                            GetBitValue = "224"
                        Case "1110"
                            GetBitValue = "256"
                        Case Else
                            GetBitValue = ""
                    End Select
                Case "2"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "8"
                        Case "0010"
                            GetBitValue = "16"
                        Case "0011"
                            GetBitValue = "24"
                        Case "0100"
                            GetBitValue = "32"
                        Case "0101"
                            GetBitValue = "40"
                        Case "0110"
                            GetBitValue = "48"
                        Case "0111"
                            GetBitValue = "56"
                        Case "1000"
                            GetBitValue = "64"
                        Case "1001"
                            GetBitValue = "80"
                        Case "1010"
                            GetBitValue = "96"
                        Case "1011"
                            GetBitValue = "112"
                        Case "1100"
                            GetBitValue = "128"
                        Case "1101"
                            GetBitValue = "144"
                        Case "1110"
                            GetBitValue = "160"
                        Case Else
                            GetBitValue = ""
                    End Select
                Case "3"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "8"
                        Case "0010"
                            GetBitValue = "16"
                        Case "0011"
                            GetBitValue = "24"
                        Case "0100"
                            GetBitValue = "32"
                        Case "0101"
                            GetBitValue = "40"
                        Case "0110"
                            GetBitValue = "48"
                        Case "0111"
                            GetBitValue = "56"
                        Case "1000"
                            GetBitValue = "64"
                        Case "1001"
                            GetBitValue = "80"
                        Case "1010"
                            GetBitValue = "96"
                        Case "1011"
                            GetBitValue = "112"
                        Case "1100"
                            GetBitValue = "128"
                        Case "1101"
                            GetBitValue = "144"
                        Case "1110"
                            GetBitValue = "160"
                        Case Else
                            GetBitValue = ""
                    End Select
                Case Else
                    GetBitValue = ""
            End Select
        Case "2.5"
            Select Case LayerValue
                Case "1"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "32"
                        Case "0010"
                            GetBitValue = "48"
                        Case "0011"
                            GetBitValue = "56"
                        Case "0100"
                            GetBitValue = "64"
                        Case "0101"
                            GetBitValue = "80"
                        Case "0110"
                            GetBitValue = "96"
                        Case "0111"
                            GetBitValue = "112"
                        Case "1000"
                            GetBitValue = "128"
                        Case "1001"
                            GetBitValue = "144"
                        Case "1010"
                            GetBitValue = "160"
                        Case "1011"
                            GetBitValue = "176"
                        Case "1100"
                            GetBitValue = "192"
                        Case "1101"
                            GetBitValue = "224"
                        Case "1110"
                            GetBitValue = "256"
                        Case Else
                            GetBitValue = ""
                    End Select
                Case "2"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "8"
                        Case "0010"
                            GetBitValue = "16"
                        Case "0011"
                            GetBitValue = "24"
                        Case "0100"
                            GetBitValue = "32"
                        Case "0101"
                            GetBitValue = "40"
                        Case "0110"
                            GetBitValue = "48"
                        Case "0111"
                            GetBitValue = "56"
                        Case "1000"
                            GetBitValue = "64"
                        Case "1001"
                            GetBitValue = "80"
                        Case "1010"
                            GetBitValue = "96"
                        Case "1011"
                            GetBitValue = "112"
                        Case "1100"
                            GetBitValue = "128"
                        Case "1101"
                            GetBitValue = "144"
                        Case "1110"
                            GetBitValue = "160"
                        Case Else
                            GetBitValue = ""
                    End Select
                Case "3"
                    Select Case bitratevalue
                        Case "0000"
                            GetBitValue = "free"
                        Case "0001"
                            GetBitValue = "8"
                        Case "0010"
                            GetBitValue = "16"
                        Case "0011"
                            GetBitValue = "24"
                        Case "0100"
                            GetBitValue = "32"
                        Case "0101"
                            GetBitValue = "40"
                        Case "0110"
                            GetBitValue = "48"
                        Case "0111"
                            GetBitValue = "56"
                        Case "1000"
                            GetBitValue = "64"
                        Case "1001"
                            GetBitValue = "80"
                        Case "1010"
                            GetBitValue = "96"
                        Case "1011"
                            GetBitValue = "112"
                        Case "1100"
                            GetBitValue = "128"
                        Case "1101"
                            GetBitValue = "144"
                        Case "1110"
                            GetBitValue = "160"
                        Case Else
                            GetBitValue = ""
                    End Select
            End Select
    End Select

ErrHandler:
    If Err.Number <> 0 Then
        GetBitValue = ""
    End If
End Function

Private Function GetSampleRate(samplevalue As String, mpegversion As String) As String

    On Error GoTo ErrHandler
    Select Case samplevalue
        Case "00"
            If mpegversion = "1" Then GetSampleRate = "44100" Else GetSampleRate = "22050"
        Case "01"
            If mpegversion = "1" Then GetSampleRate = "48000" Else GetSampleRate = "24000"
        Case "10"
            If mpegversion = "1" Then GetSampleRate = "32000" Else GetSampleRate = "16000"
        Case Else
            GetSampleRate = ""
    End Select
ErrHandler:
    If Err.Number <> 0 Then
        GetSampleRate = ""
    End If
End Function

Private Function GetMpegVersion(versionvalue As String) As String

    On Error GoTo ErrHandler

    Select Case versionvalue
        Case "00"
            GetMpegVersion = "2.5"
        Case "01"
            GetMpegVersion = "reserved"
        Case "10"
            GetMpegVersion = "2"
        Case "11"
            GetMpegVersion = "1"
        Case Else
            GetMpegVersion = ""
    End Select

ErrHandler:
    If Err.Number <> 0 Then
        GetMpegVersion = ""
    End If
End Function

Private Function GetLayerValue(LayerValue As String) As String

    On Error GoTo ErrHandler
    Select Case LayerValue
        Case "00"
            GetLayerValue = "reserved"
        Case "01"
            GetLayerValue = "3"
        Case "10"
            GetLayerValue = "2"
        Case "11"
            GetLayerValue = "1"
        Case Else
            GetLayerValue = ""
    End Select

ErrHandler:
    If Err.Number <> 0 Then
        GetLayerValue = ""
    End If
End Function

Private Function decimaltobinary(decvalue As Byte) As String

    Dim TV As Long
    Dim RM As Byte
    Dim BS As String
    Dim i As Long
    On Error GoTo ErrHandler
    BS = ""
    TV = decvalue
    Do
        RM = TV Mod 2
        BS = Right(Str(RM), 1) & BS
        TV = TV \ 2
    Loop Until TV = 0
    'Pad
    While Len(BS) < 8
        BS = "0" & BS
    Wend

    decimaltobinary = BS

ErrHandler:
    If Err.Number <> 0 Then
        decimaltobinary = ""
    End If
End Function
