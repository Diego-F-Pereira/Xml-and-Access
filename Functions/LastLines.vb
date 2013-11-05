'-------------------------------------------------------------------------------------------
' Name:         LastLines
' Purpose:      Reads the last lines of a file.
' Description:  strPath = The complete Path.
'               strFile = The name of the file.
'               lngNodes = The number of lines to be read
' Author:       Diego F.Pereira-Perdomo
' Date:         Oct-27-2013
' References:   Requires the Microsoft Scripting Runtime library
'-------------------------------------------------------------------------------------------

Public Function LastLines(strPath As String, _
                          strFile As String, _
                         lngNodes As Long) As String
                      
    Dim oFSO    As Scripting.FileSystemObject
    Dim oTSt    As Scripting.TextStream
    Dim strPF   As String
    Dim strSL   As String
    Dim i       As Long
    Dim j       As Long
       
    strPF = strPath & "\" & strFile
    
    Set oFSO = New Scripting.FileSystemObject
    Set oTSt = oFSO.OpenTextFile(strPF, ForReading)
    
    
    With oTSt
        Do While Not .AtEndOfStream
            .SkipLine
        Loop
        i = .Line
        .Close
    End With
        j = i - lngNodes + 1
        
    Set oTSt = Nothing
    Set oTSt = oFSO.OpenTextFile(strPF, ForReading)
    
    With oTSt
        Do While Not .AtEndOfStream
            Select Case .Line
                Case j To i
                    strSL = strSL & .ReadLine & vbCrLf
                Case Else
                    .SkipLine
            End Select
        Loop
        .Close
    End With
    
    Set oTSt = Nothing
    Set oFSO = Nothing
    
    LastLines = strSL

End Function