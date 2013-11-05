'-------------------------------------------------------------------------------------------
' Name:         FirstLines
' Purpose:      Reads the first lines of a file.
' Description:  strPath = The complete Path.
'               strFile = The name of the file.
'               lngNodes = The number of lines to be read
' Author:       Diego F.Pereira-Perdomo
' Date:         Oct-27-2013
' References:   Requires the Microsoft Scripting Runtime library
'-------------------------------------------------------------------------------------------
Public Function FirstLines(strPath As String, _
                           strFile As String, _
                          lngNodes As Long) As String
                      
    Dim oFSO    As Scripting.FileSystemObject
    Dim oTSt    As Scripting.TextStream
    Dim strPF   As String
    Dim strSL   As String
    Dim i       As Long
       
    strPF = strPath & "\" & strFile
    
    Set oFSO = New Scripting.FileSystemObject
    Set oTSt = oFSO.OpenTextFile(strPF, ForReading)
    
    With oTSt
        For i = 1 To lngNodes
            strSL = strSL & .ReadLine & vbCrLf
        Next i
        .Close
    End With
    
    Set oFSO = Nothing
    
    FirstLines = strSL

End Function
