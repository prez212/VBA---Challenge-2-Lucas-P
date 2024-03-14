
Sub MultiYearStock()
    
    Dim w As Long

    Application.ScreenUpdating = False

    For w = 1 To worksheets.Count
    Sheets(w).Select


    Application.ScreenUpdating = True
        
        ' Declare variables outside of the loop
        Dim LastRow As Long
        Dim Summary_Table_Row As Integer
        Dim Stock_Name As String
        Dim Stock_Volume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentageChange As Double
        Dim IncTicker As String
        Dim DecTicker As String
        Dim VolTicker As String
        Dim GreatestVolume As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        
        
        ' Set column headers
        Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        Range("P1:Q1").Value = Array("Ticker", "Value")
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Dim rng As Range
             Dim maxvalue As Variant
                Set rng = Range("K:K")
                maxvalue = Application.WorksheetFunction.Max(rng)

        Set rng = Range("K:K")
        maxvalue = Application.WorksheetFunction.Max(rng)
        Range("Q2").Value = maxvalue
        Range("Q2").NumberFormat = "0.00%"
        
        Set rng = Range("K:K")
        minvalue = Application.WorksheetFunction.Min(rng)
        Range("Q3").Value = minvalue
        Range("Q3").NumberFormat = "0.00%"
        
        Set rng = Range("L:L")
        maxvalue = Application.WorksheetFunction.Max(rng)
        Range("Q4").Value = maxvalue
        
        ' Initialize IncTicker
        GreatestIncrease = 0
        IncTicker = " "
        
        'Initialize DecTicker
        GreatestDecrease = 0
        DecTicker = " "
        
        'Initialize VolTicker
        GreatestVolume = 0
        VolTicker = " "
        
        ' Find the last row
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize the Summary Table Row
        Summary_Table_Row = 2
        
        ' Loop through rows from 2 to the last row
        For i = 2 To LastRow
        
        'Check if new ticker
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        'Set Ticker
        Ticker = Cells(i, 1).Value
        
    End If
            
            ' Check if the next row has a different Ticker name
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                
        ' Assign values to variables
            Stock_Name = Cells(i, 1).Value
            OpenPrice = Cells(i, 3).Value
        ' Reset Stock Volume
            Stock_Volume = 0
            
    End If
    
            'Add Stock_Volume
            Stock_Volume = Stock_Volume + Cells(i, 7).Value

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Calculate YearlyChange and PercentageChange
            ClosePrice = Cells(i, 6).Value
            YearlyChange = ClosePrice - OpenPrice
                
                If OpenPrice <> 0 Then
                    PercentageChange = YearlyChange / OpenPrice
                Else
                    PercentageChange = 0
                End If
                
         'Check IncTicker
                If PercentageChange > GreatestIncrease Then
                GreatestIncrease = PercentageChange
                IncTicker = Cells(i, 1).Value
                End If
        'Check DecTicker
                If PercentageChange < GreatestDecrease Then
                GreatestDecrease = PercentageChange
                DecTicker = Cells(i, 1).Value
                End If
        'Check VolTicker
                If Stock_Volume > GreatestVolume Then
                GreatestVolume = Stock_Volume
                VolTicker = Cells(i, 1).Value
                End If
                
    
        ' Print the Summary Table
                Range("I" & Summary_Table_Row).Value = Stock_Name
                Range("L" & Summary_Table_Row).Value = Stock_Volume
                Range("J" & Summary_Table_Row).Value = YearlyChange
            If YearlyChange > 0 Then
                Range("J" & Summary_Table_Row).Interior.Color = vbGreen
            Else
                Range("J" & Summary_Table_Row).Interior.Color = vbRed
            End If
                Range("K" & Summary_Table_Row).Value = PercentageChange
                Range("P2").Value = IncTicker
                Range("P3").Value = DecTicker
                Range("P4").Value = VolTicker
                
        ' Increment the Summary Table Row for the next iteration
        Summary_Table_Row = Summary_Table_Row + 1
                
        
Else
        
End If
        
Next i

Next w


End Sub
