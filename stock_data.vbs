Sub Stock_Data():
    ' Declare Current as a worksheet object variable.
    Dim Current As Worksheet
     ' Loop through all of the worksheets in the active workbook.
    For Each Current In Worksheets
        'make a header for each sheet/year
        Current.Range("I1").Value = "Ticker"
        Current.Range("J1").Value = "Yearly Change"
        Current.Range("K1").Value = "Percent Change"
        Current.Range("L1").Value = "Total Stock Volume"
        'Finds the last non-blank cell in a single row or column
        Dim NumRows As Long
        'Find the last non-blank cell in column A(1)
        NumRows = Current.Range("A2", Current.Range("A2").End(xlDown)).Rows.Count
        Dim i As Long
        Dim k As Long
        'j is the row number of new ticker in the result table, initialize j as 2
        j = 2
        'k is the start of a new ticker in the data table, initialize k as 2
        k = 2
        For i = 2 To NumRows
        If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
            '1. assign ticker
            Current.Cells(j, 9).Value = Current.Cells(i, 1).Value
            '2. calculate yearly change for each ticker: <close> - <open>
            Current.Cells(j, 10).Value = Current.Cells(i, 6).Value - Current.Cells(k, 3).Value
            '3. calculate percent change for each ticker: (<close> - <opem>)/<open>
            'be careful if <open> = 0!
            If Current.Cells(k, 3) = 0 Then
                Current.Cells(j, 11).Value = Current.Cells(i, 6).Value
            Else
                Current.Cells(j, 11).Value = (Current.Cells(i, 6).Value - Current.Cells(k, 3).Value) / Current.Cells(k, 3).Value
            End If
            Current.Cells(j, 11).NumberFormat = "0.00%"
            '4. calculate total stock volume
            Dim TotalStockVolume As Double
            TotalStockVolume = 0
            m = k
            'for each ticker, the range is k to i
            For tickerRange = k To i
               TotalStockVolume = Current.Cells(m, 7).Value + TotalStockVolume
               m = m + 1
            Next tickerRange
           Current.Cells(j, 12).Value = TotalStockVolume
           'conditional formating: red for negative, green for positive
           If Current.Cells(j, 10).Value < 0 Then
            Current.Cells(j, 10).Interior.ColorIndex = 3
           ElseIf Current.Cells(j, 10).Value > 0 Then
            Current.Cells(j, 10).Interior.ColorIndex = 4
           End If
            'Update j And k
            j = j + 1
            k = i + 1
         End If
    Next i
     
    'Bonus question
     'make a header for each sheet/year
    Current.Range("O2").Value = "Greatest % Increase"
    Current.Range("O3").Value = "Greatest % Decrease"
    Current.Range("O4").Value = "Greatest Total Volume"
    Current.Range("P1").Value = "Ticker"
    Current.Range("Q1").Value = "Value"
    
    Dim ResultRows As Long
    'Find the last non-blank cell in column A(1)
    ResultRows = Current.Range("I2", Current.Range("I2").End(xlDown)).Rows.Count
    '1. find the greatest % increase
    'initiate a temp value of percent change and ticker value
    temp = Current.Cells(2, 11).Value
    ticker = Current.Cells(2, 9).Value
    For i = 2 To ResultRows
        If Current.Cells(i + 1, 11).Value > temp Then
            temp = Current.Cells(i + 1, 11).Value
            ticker = Current.Cells(i + 1, 9).Value
        End If

    Next i
    Current.Range("Q2").Value = temp
    Current.Range("Q2").NumberFormat = "0.00%"
    Current.Range("P2").Value = ticker
    
    '2. find the greatest % decrease
     'initiate a temp value of percent change and ticker value
    temp = Current.Cells(2, 11).Value
    ticker = Current.Cells(2, 9).Value
    For i = 2 To ResultRows
        If Current.Cells(i + 1, 11).Value < temp Then
            temp = Current.Cells(i + 1, 11).Value
            ticker = Current.Cells(i + 1, 9).Value
        End If

    Next i
    Current.Range("Q3").Value = temp
    Current.Range("Q3").NumberFormat = "0.00%"
    Current.Range("P3").Value = ticker
    
    '3. find the greatest total volume
     'initiate a temp value of total stock volume and ticker value
    temp = Current.Cells(2, 12).Value
    ticker = Current.Cells(2, 9).Value
    For i = 2 To ResultRows
        If Current.Cells(i + 1, 12).Value > temp Then
            temp = Current.Cells(i + 1, 12).Value
            ticker = Current.Cells(i + 1, 9).Value
        End If
    Next i
    Current.Range("Q4").Value = temp
    Current.Range("P4").Value = ticker

    Next Current
 
End Sub
