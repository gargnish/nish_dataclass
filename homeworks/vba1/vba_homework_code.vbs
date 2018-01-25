Sub one()
    Dim x As String
    
    
    For i = 1 To Sheets.Count
        Sheets(i).Activate
        With ActiveSheet
            x = mycalculation()
        End With
        Next i
        
        
        
        
    End Sub
    
    Function mycalculation() As String
        
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly change"
        Range("K1").Value = "Percentage change"
        Range("L1").Value = "Total Stock Volume"
        
'scount source count
        scount = Cells(Rows.Count, 1).End(xlUp).Row
'sorting
        Range("A2", "G" & scount).Sort [A2], xlAscending
        
'loop and fill at target
        
        Dim v1 As Variant
        Dim x As Range
        Dim m As String
        m = ""
        Set x = Range("A2:A" & scount)
        Dim i As Long
        i = 1
        
        Dim k As Long
        k = 2
        
'loop thou source table
        For Each v1 In x
            
' unique ids
            If m <> v1.Value Then
                i = i + 1
                m = v1.Value
                Range("I" & i).Value = m
            End If
            
'sum of stock volume
            Range("L" & i).Value = Range("L" & i).Value + Range("G" & k).Value
'min of date
            If Range("M" & i).Value > Range("B" & k).Value Or Range("M" & i).Value = 0 Then
'min value
                Range("M" & i).Value = Range("B" & k).Value
'open at min
                Range("O" & i).Value = Range("C" & k).Value
            End If
'max of date
            If Range("N" & i).Value < Range("B" & k).Value Or Range("N" & i).Value = 0 Then
'max value
                Range("N" & i).Value = Range("B" & k).Value
'close at max
                Range("P" & i).Value = Range("F" & k).Value
            End If
            
            k = k + 1
            Next v1
            
            Dim tcount As Long
            
'tcount target count
            tcount = i
            
            For k1 = 2 To tcount
'increase or decreases
                Range("J" & k1).Value = Range("P" & k1).Value - Range("O" & k1).Value
                If Range("O" & k1).Value <> 0 Then
                    Range("K" & k1).Value = (Range("P" & k1).Value - Range("O" & k1).Value) / Range("O" & k1).Value
                Else
                    Range("K" & k1).Value = 0
                End If
                If Range("J" & k1).Value > 0 Then
                    Range("J" & k1).Interior.Color = RGB(0, 255, 0)
                Else
                    Range("J" & k1).Interior.Color = RGB(255, 0, 0)
                End If
                
                Range("K" & k1).NumberFormat = "0.00%"
                Range("J" & k1).NumberFormat = "0.########"
                
' truncate columns
                Range("M" & k1).Clear
                Range("N" & k1).Clear
                Range("O" & k1).Clear
                Range("P" & k1).Clear
                
'Greatest % increase
                If Range("R" & 2).Value < Range("K" & k1).Value Or Range("R" & 2).Value = 0 Then
'max value
                    Range("R" & 2).Value = Range("K" & k1).Value
'ticker
                    Range("Q" & 2).Value = Range("I" & k1).Value
                End If
'Greatest % Decrease
                If Range("R" & 3).Value > Range("K" & k1).Value Or Range("R" & 3).Value = 0 Then
'min value
                    Range("R" & 3).Value = Range("K" & k1).Value
'ticker
                    Range("Q" & 3).Value = Range("I" & k1).Value
                End If
'Greatest Total Volume
                If Range("R" & 4).Value < Range("L" & k1).Value Or Range("R" & 4).Value = 0 Then
'max value
                    Range("R" & 4).Value = Range("L" & k1).Value
'ticker
                    Range("Q" & 4).Value = Range("I" & k1).Value
                End If
                Next k1
                
                
                Range("R" & 2).NumberFormat = "0.00%"
                Range("R" & 3).NumberFormat = "0.00%"
                
                Range("Q1").Value = "Ticker"
                Range("R1").Value = "Value"
                
                
                
                
                Range("P2").Value = "Greatest % Increase"
                Range("P3").Value = "Greatest % Decrease"
                Range("P4").Value = "Greatest Total Volume"
                
                Columns("O").EntireColumn.Delete
                
                
            End Function
            
            
            
            
            
            



