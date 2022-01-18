Sub CalculatePrimeNumbers()

    Worksheets("Sheet1").Rows(6).Cells.Borders.LineStyle = xlNone
    Worksheets("Sheet1").Rows(6).Value = ""
    Worksheets("Sheet1").Range("B6").Borders.LineStyle = xlContinuous
    Worksheets("Sheet1").Range("B6").Value = "Result"

    If (IsEmpty(Worksheets("Sheet1").Range("C3"))) Then
        Exit Sub
    End If
    
    Dim valueToBeChecked As String
    valueToBeChecked = Worksheets("Sheet1").Range("C3").Value
    
    Dim factorialDecomposition As Scripting.Dictionary
    Set factorialDecomposition = New Scripting.Dictionary
    
    Dim temporaryPower
    Dim denominator
    denominator = 2
    
    While valueToBeChecked > 1
        While valueToBeChecked Mod denominator = 0
            valueToBeChecked = valueToBeChecked / denominator
            
            temporaryPower = 1
            If factorialDecomposition.Exists(denominator) Then
                temporaryPower = factorialDecomposition.Item(denominator)
                temporaryPower = temporaryPower + 1
                factorialDecomposition.Remove (denominator)
            End If
            
            factorialDecomposition.Add denominator, temporaryPower
        Wend
    denominator = denominator + 1
    Wend
    
    
    Dim ColumnPrime, ColumnValue
    
    Dim index
    index = 0
    
    
    For Each varKey In factorialDecomposition.Keys()
        Dim CellCoords
        CellCoords = Chr(Asc("C") + index) & "6"
        Worksheets("Sheet1").Range(CellCoords) = varKey & " ^ " & factorialDecomposition(varKey)
        Worksheets("Sheet1").Range(CellCoords).Borders.LineStyle = xlContinuous
        index = index + 1
    Next

End Sub
