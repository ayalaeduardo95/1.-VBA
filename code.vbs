Sub Stocks()

'Variable List
'Worksheet = WS
'Ticker = T
'Yearly Change = YC
'Percent Change = PC
'Total Stock Volume = TSV
'Summary Table Row = STR
'Last Row = LR
'Last Row Format = LRF
'Closing Price = CP
'Opening Price = OP
'Greatest Increase = GI
'Greatest Decrease = GD
'Greates Total Volume = GTV

    Dim WS As Worksheet
    Dim T As String
    Dim YC As Double
    Dim PC As Double
    Dim TSV As Double
    Dim STR As Integer
    Dim LR As Long
    Dim LRF As Long
    Dim CP As Double
    Dim OP As Double
    Dim GI As Double
    Dim GD As Double
    Dim GTV As Double
    
    'First Activate For Loop for all Worksheets
    
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
    'Then For Loop for to bring value of T, YC, PC & TSV
    
    STR = 2
    
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    
    OP = Cells(2, 3).Value
    
        For i = 2 To LR
    
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                LP = Cells(i, 6).Value
                T = Cells(i, 1).Value
                YC = LP - OP
                
                    If OP = 0 Or LP = 0 Then
                    PC = 0
                    Else
                    PC = (LP / OP) - 1
                    End If
                    
                TSV = TSV + Cells(i, 7).Value
                
                Cells(STR, 9).Value = T
                Cells(STR, 10).Value = YC
                Cells(STR, 11).Value = PC
                Cells(STR, 12).Value = TSV
                
                STR = STR + 1
                
                TSV = 0
                
                OP = Cells(i + 1, 3).Value

            Else
                
               TSV = TSV + Cells(i, 7).Value
                
            End If
            
        Next i
            
            'For Loop to get Values and Ticker for GI, GD & GTV
            
            LRF = Cells(Rows.Count, 10).End(xlUp).Row
            
            GI = 2
            GD = 2
            GTV = 2
            
                For j = 2 To LRF
                
                'Format for YC
            
                    If Cells(j, 10).Value <= 0 Then
                    
                    Cells(j, 10).Interior.ColorIndex = 3
                    
                    Else
                    
                    Cells(j, 10).Interior.ColorIndex = 4
                    
                    End If
                    
                    'Greatest Increase
                    
                    If Cells(j, 11).Value > Cells(GI, 11).Value Then
                    
                    GI = j
                    
                    End If
                    
                    'Greatest Decrease
                    
                    If Cells(j, 11).Value < Cells(GD, 11).Value Then
                    
                    GD = j
                    
                    End If
                    
                    'GTV
                    
                    If Cells(j, 12).Value > Cells(GTV, 12).Value Then
                    
                    GTV = j
                    
                    End If
                    
                Next j
                
                'Assigning Values to GI, GD & GTV
                
                Cells(2, 16).Value = Cells(GI, 11).Value
                Cells(2, 15).Value = Cells(GI, 9).Value
                Cells(2, 16).NumberFormat = "0.00%"
                
                Cells(3, 16).Value = Cells(GD, 11).Value
                Cells(3, 15).Value = Cells(GD, 9).Value
                Cells(3, 16).NumberFormat = "0.00%"
            
                Cells(4, 16).Value = Cells(GTV, 12).Value
                Cells(4, 15).Value = Cells(GTV, 9).Value
        
                'Format for PC
                
                Range("K:K").NumberFormat = "0.00%"
                
    
    
    
    Next
        
End Sub
