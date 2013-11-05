'-------------------------------------------------------------------------------------------
' Name:         clsSaxContentHandlerCountNodes
' Purpose:      Counts the number "SupplementalRecord" elements in the document
' Author:       Diego F.Pereira-Perdomo
' Date:         Oct-28-2013
' References:   Requires the Microsoft XML, v6.0 library
'-------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
    
Dim lngC As Long

Implements IVBSAXContentHandler
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_startDocument()
    Debug.Print Now()
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_startPrefixMapping(strPrefix As String, _
                                                       strURI As String)
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_startElement(strNamespaceURI As String, _
                                                 strLocalName As String, _
                                                     strQName As String, _
                                            ByVal oAttributes As MSXML2.IVBSAXAttributes)
    Select Case strLocalName
        Case "SupplementalRecord"
            lngC = lngC + 1
        Case Else
    End Select

End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_characters(strChars As String)
'
End Sub
'-------------------------------------------------------------------------------------------
Private Property Set IVBSAXContentHandler_documentLocator(ByVal RHS As MSXML2.IVBSAXLocator)
'
End Property
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_ignorableWhitespace(strChars As String)
'
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_processingInstruction(strTarget As String, _
                                                         strData As String)
'
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_skippedEntity(strName As String)
'
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_endElement(strNamespaceURI As String, _
                                               strLocalName As String, _
                                                   strQName As String)
'
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_endPrefixMapping(strPrefix As String)
'
End Sub
'-------------------------------------------------------------------------------------------
Private Sub IVBSAXContentHandler_endDocument()
    Debug.Print Now()
    Debug.Print "There are " & lngC & " SupplementalRecord elements"
End Sub
