'-------------------------------------------------------------------------------------------
' Name:         RandomNumbers
' Purpose:      Generates pseudo-random numbers and stores them in an array.
' Description:  lngA = Amount of pseudorandom numbers to be generated.
'               lngL = Lower number.
'               lngU = Upper number
'               lngSeed = Optional. Initial value to be used by Rnd
' Author:       Diego F.Pereira-Perdomo
' Date:         Oct-27-2013
' References:   Requires the Microsoft Scripting Runtime library
'-------------------------------------------------------------------------------------------
Function RandomNumbers(lngA As Long, _
                       lngL As Long, _
                       lngU As Long, _
           Optional lngSeed As Long) As Variant
           
    Dim i       As Long
    Dim j       As Long
    Dim aRnd
    
    j = lngA - 1
    ReDim aRnd(j)
    
    If IsMissing(lngSeed) Then
        For i = 0 To j
            aRnd(i) = Int((lngU - lngL + 1) * Rnd + lngL)
        Next i
    Else
        For i = 0 To j
            aRnd(i) = Int((lngU - lngL + 1) * Rnd(lngSeed) + lngL)
        Next i
    End If
    
    aRnd = BubbleSort(aRnd)
    
    RandomNumbers = aRnd

End Function
'-------------------------------------------------------------------------------------------

' Original function from http://support.microsoft.com/kb/133135
Public Function BubbleSort(ByVal tempArray As Variant) As Variant
Dim Temp        As Variant
Dim i           As Integer
Dim NoExchanges As Integer

    ' Loop until no more "exchanges" are made.
    Do
        NoExchanges = True
        
        ' Loop through each element in the array.
        For i = 0 To UBound(tempArray) - 1
        
            ' Substitution when element is greater than the element following int
            If tempArray(i) > tempArray(i + 1) Then
                NoExchanges = False
                Temp = tempArray(i)
                tempArray(i) = tempArray(i + 1)
                tempArray(i + 1) = Temp
            End If
        
        Next i
    
    Loop While Not (NoExchanges)
    
    BubbleSort = tempArray

End Function