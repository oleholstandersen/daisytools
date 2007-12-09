VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SAXContentHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
 
' This class is made to locate the line/column of an element found in DOM.
' There are two methods of doing this, the first one is to get an ID of the element
' or of an nearby element and use that as the criteria. The other method is to
' create an ancestorchain. The ancestor chain has the partially the same syntax as
' XPath;
'
' html/body/h1/a[@href="http://www.daisy.org" and alt="The Daisy consortium"]
'
' This will look for an <a> element nested within a 'html/body/h1' context. The
' <a> element must have the attributes 'href' and 'alt' and their defined content

' Implement a few sax interfaces
Implements IVBSAXContentHandler
Implements IVBSAXLocator

' oLocator = locator object (IVBSAXLocator)
Private oLocator As IVBSAXLocator

' bolFound = true if a matching pair of attributes / elements have been found
' bolDone = true if the right element has been found
Private bolFound As Boolean, bolDone As Boolean

' sId = ID of the element to find
' sText = Name of the element or attribute to find
' lLine = Found elements line
' lColumn = Found elements column
Public sId As String, sText As String, lLine As Long, lColumn As Long

' a type defining an attribute
Private Type typAttributes
  sName As String
  sValue As String
  bolFound As Boolean
End Type

' sAncestors() = the ancestor chain, starting at 1
' lAncestors = number of items in sAncestors()
' lCurrentAncestor = number of the current ancestor
Private sAncestors() As String, lAncestors As Long, lCurrentAncestor As Long

' sAttributes() = the attributes that are to be found within the element
' lAttributes() = number of items in sAttributes()
' lLevel = the current elements hierarchy level
Private sAttributes() As typAttributes, lAttributes As Long, lLevel As Long

' nodeType of node to be found
Public eNodeType As Long

' this method sets and parses the ancestorchain that should be found
Public Property Let sAncestorChain(isAncestorChain As String)
  Dim lCounter As Long, lStart As Long, lEnd As Long
  Dim sTemp As String

  lCounter = 1
  Do
    lCounter = 1
    sTemp = Mid$(isAncestorChain, lCounter, 1)
' search for the next / or [
    Do Until sTemp = "/" Or sTemp = "["
      lCounter = lCounter + 1
      If lCounter >= Len(isAncestorChain) Then Exit Do
      sTemp = Mid$(isAncestorChain, lCounter, 1)
    Loop
    
' insert the ancestor in the ancestorchain
    lAncestors = lAncestors + 1
    ReDim Preserve sAncestors(lAncestors)
    sAncestors(lAncestors) = Left$(isAncestorChain, lCounter - 1)
    
    If (sTemp = "[") Or (lCounter > Len(isAncestorChain)) Then Exit Do
    
' remove the inserted ancestor from the ancestorchain
    isAncestorChain = Right$(isAncestorChain, Len(isAncestorChain) - lCounter)
  Loop
  
  ReDim Preserve sAncestors(lAncestors)
  
' look for [
  lCounter = InStr(1, isAncestorChain, "[", vbBinaryCompare)
' if found, we have attribute declarations
  If lCounter > 0 Then
' look for an attribute declaration (allways starts with an @)
    lCounter = InStr(1, isAncestorChain, "@", vbBinaryCompare)
    Do Until lCounter = 0
      lStart = lCounter + 1
' search for the = sign (the end of the attribute name)
      lEnd = InStr(1, isAncestorChain, "=", vbBinaryCompare)
      
' insert the attribute in the attribute list
      lAttributes = lAttributes + 1
      ReDim Preserve sAttributes(lAttributes)
      sAttributes(lAttributes).sName = Mid$(isAncestorChain, lStart, lEnd - lStart)
      
' search for the end of the attribute value
      lStart = lEnd + 2
      lEnd = InStr(lStart, isAncestorChain, Mid$(isAncestorChain, lStart - 1, 1), vbBinaryCompare) - 1
      If lEnd < 1 Then lEnd = InStr(1, isAncestorChain, "]", vbBinaryCompare) - 1
      lEnd = lEnd + 1
           
' insert the attribute value in the attribute list
      sAttributes(lAttributes).sValue = Mid$(isAncestorChain, lStart, lEnd - lStart)

' if we haven't reached the end of the string, keep on searching. Otherwise, exit.
      If Len(isAncestorChain) - lEnd - 4 > 0 Then
        isAncestorChain = Right$(isAncestorChain, Len(isAncestorChain) - lEnd - 5)
      Else
        Exit Do
      End If
      
      lCounter = InStr(1, isAncestorChain, "@", vbBinaryCompare)
    Loop
  End If
End Property

Private Sub Class_Initialize()
  bolFound = False
  lLine = -1
  lColumn = -1
End Sub

Private Property Set IVBSAXContentHandler_documentLocator(ByVal RHS As _
   MSXML2.IVBSAXLocator)
'  MSXML2.IVBSAXLocator)
    
    Set oLocator = RHS
End Property

' This function is called every time a new element is starting, this is the
' function that we're using to locate the element that we're looking for.
'Private Sub IVBSAXContentHandler_startElement( _
'  strNamespaceURI As String, strLocalName As String, strQName As String, _
'  ByVal oAttributes As MSXML2.IVBSAXAttributes)
Private Sub IVBSAXContentHandler_startElement( _
  strNamespaceURI As String, strLocalName As String, strQName As String, _
  ByVal oAttributes As MSXML2.IVBSAXAttributes)
    
  Dim lCounter As Integer, sElement As String, bolAC As Boolean, lCounter2 As Long
  Dim lL As Long, lC As Long, bolTooMany As Boolean
  
  If bolDone Then Exit Sub
  
' if the next ancestor in the chain has the same name and level as the current
' element, increase the ancestor count.
  If Not lCurrentAncestor = lAncestors Then
    If (lLevel = lCurrentAncestor) And _
      (sAncestors(lCurrentAncestor + 1) = strLocalName) Then
    
      lCurrentAncestor = lCurrentAncestor + 1
      If lCurrentAncestor = lAncestors Then bolAC = True
    End If
  ElseIf lAncestors = 0 Then

  ElseIf sAncestors(lCurrentAncestor) = strLocalName Then
' if this element is the same as the current ancestor, set AC to true
    bolAC = True
  End If
  
  lLevel = lLevel + 1
  
  lL = oLocator.lineNumber
  lC = oLocator.columnNumber
  bolTooMany = False

' Go trough this elements attributes
  For lCounter = 0 To (oAttributes.length - 1)
' if this element is the element we're looking for, check the attributes.
    If (bolAC) And (Not bolTooMany) Then
' go trough all the attributes defined in the attribute and see if any of the
' defintions has the same values as the current attribute
      For lCounter2 = 1 To lAttributes
        If sAttributes(lCounter2).sName = oAttributes.getLocalName(lCounter) And _
          sAttributes(lCounter2).sValue = oAttributes.getValue(lCounter) Then _
          sAttributes(lCounter2).bolFound = True
      Next lCounter2
    End If
  
' if an ID has bin given and it is found on this attribute, then this is the
' element we're looking for
    If bolFound = False Then
      If (Not sId = "") And (oAttributes.getLocalName(lCounter) = "id") Then
        If oAttributes.getValue(lCounter) = sId Then bolFound = True
      End If
    End If
  Next

' if the number of attributes in this element isn't the same as the number of
' attributes in the element we're looking for, bolTooMany = true
  If oAttributes.length <> lAttributes Then bolTooMany = True

' If we have the same nuber of attributes, check so that all attributes .bolFound
' variable = true
  If Not bolTooMany Then
    bolFound = True
    For lCounter = 1 To lAttributes
      If sAttributes(lCounter).bolFound = False Then bolFound = False
    Next lCounter
  End If
  
' If bolFound is still = true then we're at it. If it's the same node type and has
' the same name as given in sText; this is our element.
  If bolFound Then
    Select Case eNodeType
      Case NODE_ATTRIBUTE
        For lCounter = 0 To (oAttributes.length - 1)
          If (oAttributes.getLocalName(lCounter) = sText) And (oAttributes.getValue(lCounter) = sId) Then _
            bolDone = True: lLine = oLocator.lineNumber: lColumn = oLocator.columnNumber
        Next
      
      Case NODE_ELEMENT
        If strLocalName = sText Then _
          bolDone = True: lLine = lL: lColumn = lC
    End Select
  End If
End Sub

' This function receives a 'end element' event
Private Sub IVBSAXContentHandler_endElement(strNamespaceURI As String, _
                                            strLocalName As String, _
                                            strQName As String)
                                            
  If bolDone Then Exit Sub

  If Not lCurrentAncestor = lAncestors Then
    If (lLevel = lCurrentAncestor) And _
      (sAncestors(lCurrentAncestor) = strLocalName) Then
    
      lCurrentAncestor = lCurrentAncestor - 1
    End If
  End If
  
  lLevel = lLevel - 1
End Sub

' The following functions / subroutines / properties are required to be implemented
' by the IVBSAXContentHandler interface

Private Sub IVBSAXContentHandler_characters(text As String)
End Sub

Private Sub IVBSAXContentHandler_endDocument()

End Sub

Private Sub IVBSAXContentHandler_endPrefixMapping(strPrefix As String)

End Sub

Private Sub IVBSAXContentHandler_ignorableWhitespace(strChars As String)

End Sub

Private Sub IVBSAXContentHandler_processingInstruction(target As String, _
                                                       data As String)

End Sub

Private Sub IVBSAXContentHandler_skippedEntity(strName As String)

End Sub

Private Sub IVBSAXContentHandler_startDocument()

End Sub

Private Sub IVBSAXContentHandler_startPrefixMapping(strPrefix As String, _
                                                    strURI As String)

End Sub

Private Property Get IVBSAXLocator_columnNumber() As Long

End Property

Private Property Get IVBSAXLocator_lineNumber() As Long

End Property

Private Property Get IVBSAXLocator_publicId() As String

End Property

Private Property Get IVBSAXLocator_systemId() As String

End Property


