'-------------------------------------------------------------------------------------------
' Name:         CountNodes
' Purpose:      Counts the number of times a specific node appears in a XML file.
' Description:  strPath = The complete Path.
'               strFile = The name of the file.
'               strNode = The name of the node.
' Author:       Diego F.Pereira-Perdomo
' Date:         Oct-27-2013
' References:   Requires the Microsoft Scripting Runtime library
'-------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Dim lngC As Long

Public Function CountNodes(strPath As String, _
                           strFile As String, _
                           strNode As String) As Long
                      
    Dim oFSO    As Scripting.FileSystemObject
    Dim oTSt    As Scripting.TextStream
    Dim strPF   As String
    Dim strRL   As String

    strPF = strPath & "\" & strFile
    
    Set oFSO = New Scripting.FileSystemObject
    Set oTSt = oFSO.OpenTextFile(strPF, ForReading)
    
    With oTSt
        Do While Not .AtEndOfStream
            Counting Trim$(.ReadLine), strNode
        Loop
        .Close
    End With
        
    Set oFSO = Nothing
    CountNodes = lngC
    lngC = 0
End Function
'-------------------------------------------------------------------------------------------
Sub Counting(strRL As String, strNode As String)
    If strRL Like "*<" & strNode & "[> ]*" Then
            lngC = lngC + 1
    End If
End Sub