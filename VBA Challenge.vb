Sub test1():
    
    
    Dim Ticker As String
    Dim aRow As Integer
    aRow = 2
    
    Cells(1, 9).Value = "Ticker"
    ' Get the count of the entire rows
    
    Dim cRow As Double
    cRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Get the Loop for the Ticker names
    
    For i = 2 To cRow
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Ticker = Cells(i, 1).Value
            Cells(aRow, 9).Value = Ticker
            
            aRow = aRow + 1
            
            End If
            
        Next i
        
        
        
        ' Yearly Change
    Dim YChange As Double
    Dim YinitialValue As Double
    Dim YendValue As Double
    
    Cells(1, 10) = "YChange"
    
    aRow = 2
    
    Dim cRowsC As Double
    cRowsC = Cells(Rows.Count, 3).End(xlUp).Row
        
        ' Loop for Year Change
        
        
    For i = 2 To cRowsC
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        YinitialValue = Cells(i, 3).Value
        
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        YendValue = Cells(i, 6).Value
        
        YChange = YinitialValue - YendValue
        Cells(aRow, 10).Value = YChange
        aRow = aRow + 1
        
        End If
        
        Next i
        
       
        
        ' Percentage Change
        
    Dim PChange As Double
    Dim YinitialValuePercentage As Double
    Dim YendValuePercentage As Double
    
    Cells(1, 11) = "PChange"
    
    aRow = 2
    
    Dim cRowsPercentage As Double
    cRowsPercentage = Cells(Rows.Count, 3).End(xlUp).Row
        
        ' Loop for Percental Change
          
    For i = 2 To cRowsPercentage
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
             YinitialValuePercentage = Cells(i, 3).Value
        
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            YendValuePercentage = Cells(i, 6).Value
         
            If YinitialValuePercentage = 0 Then
            PChange = 0
        
            Else
            PChange = YendValuePercentage / YinitialValuePercentage - 1
        
            End If
        
        Cells(aRow, 11).Value = PChange
        aRow = aRow + 1
        
        End If
        
        Next i
        
        
        ' Total Stock Volume
    
    Cells(1, 12) = "Total Stock Volume"
    
    Dim TotalStockVolume As Double
        TotalStockVolume = 0
        
    Dim InitialStockVolume As Double
        InitialStockVolume = 0
        
    Dim cRowVolume As Double
    cRowVolume = Cells(Rows.Count, 1).End(xlUp).Row

        aRow = 2

    For i = 2 To cRowVolume
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        TotalStockVolume = InitialStockVolume + Cells(i, 7).Value
        Cells(aRow, 12).Value = TotalStockVolume
        InitialStockVolume = Cells(aRow, 12).Value
        
    ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Cells(aRow, 12).Value = TotalStockVolume + Cells(i, 7).Value
        aRow = aRow + 1
        InitialStockVolume = 0
        TotalStockVolume = 0
    End If
    
    Next i
    
        ' To change colors for YChange
        
    Dim cRowYChange As Double
        cRowYChange = Cells(Rows.Count, 10).End(xlUp).Row
        
        For i = 2 To cRowYChange
        
            If Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
            
            ElseIf Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
            
        End If
        
        Next i
End Sub
