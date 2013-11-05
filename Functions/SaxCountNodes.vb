'-------------------------------------------------------------------------------------------
' Name:         SaxCountNodes
' Purpose:      Calls the class clsSaxContentHandlerCountNodes.
' Description:  strPath = The complete Path.
'               strFile = The name of the file.
' Author:       Diego F.Pereira-Perdomo
' Date:         Oct-28-2013
' References:   Requires the classes clsSaxContentHandlerCountNodes and clsSaxErrorHandler
'-------------------------------------------------------------------------------------------

Sub SaxCountNodes(strPath As String, _
                  strFile As String)
On Error GoTo SaxCountNodes_Error

    Dim strPF           As String
    Dim reader          As SAXXMLReader60
    Dim contentHandler  As clsSaxContentHandlerCountNodes
    Dim errorHandler    As clsSAXErrorHandler
        
    strPath = strPath & "\"
    strPF = strPath & strFile
    
    Set reader = New SAXXMLReader60
    Set contentHandler = New clsSaxContentHandlerCountNodes
    Set errorHandler = New clsSAXErrorHandler
    
    Set reader.contentHandler = contentHandler
    Set reader.errorHandler = errorHandler
    
    reader.parseURL (strPF)

SaxCountNodes_Error:

    Select Case Err.Number
        Case 0
        Case Else
            Debug.Print Err.Number & ": " & Err.Description
    End Select
    Exit Sub
    
End Sub

