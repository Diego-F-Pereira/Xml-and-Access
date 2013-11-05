'-------------------------------------------------------------------------------------------
' Name:         SplitXml
' Purpose:      Splits a XML file into smaller files.
' Description:  strPath = The complete Path.
'               strInputFile = The name of the original XML file.
'               strOutputName = The name of the resulting XML files.
'               strNode = Parent node used for splitting the file.
'               lngNodes = Number of nodes to be included in each output file.
' Author:       Diego F.Pereira-Perdomo
' Date:         Oct-27-2013
' References:   Requires the Microsoft Scripting Runtime library
'-------------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
    Const xmlHeading = "<?xml version='1.0' encoding='utf-8'?>"
    Const xmlFirst = "<root>"
    Const xmlLast = "</root>"
    
    Dim booW    As Boolean  ' Excludes nodes between the end node and the next start node
    Dim booN    As Boolean  ' A new file has to be created
    Dim lngC    As Long     ' Counts start nodes
    Dim lngF    As Long     ' File Number
    Dim strP    As String   ' Path
    Dim strO    As String   ' Output
    Dim fFile   As Long     ' FreeFile Number
'-------------------------------------------------------------------------------------------    

Sub SplitXml(strPath As String, _
        strInputFile As String, _
       strOutputName As String, _
             strNode As String, _
            lngNodes As Long)
    
    Dim oFSO    As Scripting.FileSystemObject
    Dim oTSt    As Scripting.TextStream
    Dim strPF   As String
    
    strP = strPath
    strPF = strP & "\" & strInputFile
    strO = strOutputName

    Set oFSO = New Scripting.FileSystemObject
    Set oTSt = oFSO.OpenTextFile(strPF, ForReading)

    With oTSt

        Do While Not .AtEndOfStream
            txtStream Trim$(.ReadLine), strNode, lngNodes
        Loop

        If .AtEndOfStream Then
            EndXml
        End If
        .Close
    End With
    
'   Cleaning
    lngF = 0
    lngC = 0
    strP = ""
    strPF = ""
    strO = ""
    
    Set oTSt = Nothing
    Set oFSO = Nothing
    
End Sub
'-------------------------------------------------------------------------------------------

Sub txtStream(strRL As String, _
            strNode As String, _
           lngNodes As Long)
           
    booN = False

    Select Case True
        Case strRL Like "*<" & strNode & "[> ]*"
            booW = True
            IsNewFile lngNodes
            ValidTxt strRL
            CountUpNode
        Case strRL Like "*</" & strNode & "[> ]*"
            ValidTxt strRL
            booW = False
        Case Else
            ValidTxt strRL
    End Select
        
End Sub
'-------------------------------------------------------------------------------------------

Sub ValidTxt(strRL As String)

    If booW Then
        PrintXml strRL
    End If
    
End Sub
'-------------------------------------------------------------------------------------------

Sub PrintXml(strRL As String)
    
    Select Case booN
        Case True
        
            Select Case lngF
                Case 0
                    NewXml strRL
                Case Else
                    EndXml
                    NewXml strRL
            End Select
            
            CountUpFile
            
        Case False
            ContentXml strRL
    End Select
    
End Sub
'-------------------------------------------------------------------------------------------

Sub NewXml(strRL As String)

    Dim strPF   As String
    Dim strFile As String
    
    fFile = FreeFile

    strFile = strO & lngF & ".xml"
    strPF = strP & "\" & strFile
    
    Open strPF For Output As #fFile
        Print #fFile, xmlHeading
        Print #fFile, xmlFirst
        Print #fFile, strRL
        
    Debug.Print lngF, Now()
        
End Sub
'-------------------------------------------------------------------------------------------

Sub EndXml()

    Print #fFile, xmlLast
    Close #fFile

End Sub
'-------------------------------------------------------------------------------------------

Sub ContentXml(strRL As String)
    Print #fFile, strRL
End Sub
'-------------------------------------------------------------------------------------------

Function CountUpNode() As Long
    lngC = lngC + 1
End Function
'-------------------------------------------------------------------------------------------

Sub IsNewFile(lngNodes As Long)

    If lngC Mod lngNodes = 0 Then
        booN = True
    End If
    
End Sub
'-------------------------------------------------------------------------------------------

Sub CountUpFile()
    lngF = lngF + 1
End Sub
