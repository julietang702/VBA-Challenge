Sub MultipleYearStockData():
Const YEARLY_CHANGE_COL As Integer = 10
    Const CLOSE_COL As Integer = 6
    Dim WS As Worksheet
    For Each WS In Worksheets
    
     WS.Activate
        'Create column headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'Functionality/Summary headers
        Range(" O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
       ' Dim variables
       
        Dim InputRow As Long
        Dim OutputRow As Long
        Dim LastAValue As Long
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As Double
        Dim TotalStockVolume As Double
        Dim GreatVolume As Double
        Dim GreatDecrease As Double
        Dim GreatIncrease As Double
        OpenPrice = Cells(2, 3).Value
        OutputRow = 2
        'Row starts at 2
        X = 2
        'Find the last value in Row A


        LastAValue = Cells(Rows.Count, "A").End(xlUp).Row
        For InputRow = 2 To LastAValue
            If Cells(InputRow + 1, 1).Value <> Cells(InputRow, 1).Value Then
            Cells(OutputRow, 9).Value = Cells(InputRow, 1).Value
                
                'Inputs
            ClosePrice = Cells(InputRow, 6).Value
                
                'Calculate Yearly Change
            YearlyChange = ClosePrice - OpenPrice
                
                'Outputs
            Cells(OutputRow, 10).Value = YearlyChange
                'CONDITIONAL FORMATTING
                'RED IS NEGATIVE, GREEN IS POSITIVE

                    'Conditional formating
                If Cells(OutputRow, 10).Value < 0 Then
                    
                        'Set cell background color to red
                    Cells(OutputRow, 10).Interior.ColorIndex = 3
                    
                Else
                    
                        'Set cell background color to green
                    Cells(OutputRow, 10).Interior.ColorIndex = 4
                    
                End If
            OutputRow = OutputRow + 1
            End If
            'Calculate PercentChange
            If Cells(X, 3).Value <> 0 Then
                PercentChange = ((Cells(InputRow, 6).Value - Cells(X, 3).Value) / Cells(X, 3).Value)
                    
                Cells(OutputRow, 11).Value = Format(PercentChange, "Percent")
                    
            Else
                    
                Cells(OutputRow, 11).Value = Format(0, "Percent")
                    
            End If
            TotalStockVolume = TotalStockVolume + Cells(InputRow, 7).Value
                'Total Volume
            Cells(OutputRow, 12).Value = WorksheetFunction.Sum(Range(Cells(X, 7), Cells(InputRow, 7)))
                
                'Increase OutputRow
            OutputRow = OutputRow + 1
                
                'Set new start row of the ticker block
            X = InputRow + 1
  
            Next InputRow
            
        'Find Last I Value
        LastIValue = Cells(Rows.Count, "I").End(xlUp).Row
        '
        
        'Cells for Summary and find Loop
        GreatVolume = Cells(2, 12).Value
        GreatIncrease = Cells(2, 11).Value
        GreatDecrease = Cells(2, 11).Value
        
            'Indent
            For InputRow = 2 To LastIValue
            
                'Greatest Total Volume
                If Cells(InputRow, 12).Value > GreatVolume Then
                GreatVolume = Cells(InputRow, 12).Value
                Cells(4, 16).Value = Cells(InputRow, 9).Value
                Cells(4, 17).Value = Format(GreatVolume, "Scientific")
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                
            
                If Cells(InputRow, 11).Value > GreatIncrease Then
                GreatIncrease = Cells(InputRow, 11).Value
                Cells(2, 16).Value = Cells(InputRow, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                Cells(2, 17).Value = Format(GreatIncrease, "Percent")
                End If
                
                'Greatest decrease
                If Cells(InputRow, 11).Value < GreatDecrease Then
                GreatDecrease = Cells(InputRow, 11).Value
                Cells(3, 16).Value = Cells(InputRow, 9).Value
                Cells(2, 17).Value = Format(GreatDecrease, "Percent")
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
        
        
        Next InputRow
    Next WS

End Sub
    
