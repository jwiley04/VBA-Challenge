Sub StockYearData()


'Loop Through WorkBook................................................................

Dim MainWrksht As Worksheet

Dim WB As Workbook
    Set WB = ActiveWorkbook
    
For Each MainWrksht In Worksheets

'Name Variables .......................................................................

Dim Ticker As String
    Ticker = ""

Dim TotalStockVolume As Double
    TotalStockVolume = 0
    
Dim YearOpen As Double
    YearOpen = 0

Dim YearClose As Double
    YearClose = 0
    
Dim YearlyChange As Double
    YearlyChange = 0
    
Dim PercentChange As Double
    PercentChange = 0
    
Dim MaxTicker As String
    MaxTicker = ""
    
Dim MaxPercent As Double
    MaxPercent = 0
    
Dim MinTicker As String
    MinTicker = ""
    
Dim MinPercent As Double
    MinPercent = 0
    
Dim MaxVolumeTicker As String
    MaxVolumeTicker = ""
    
Dim MaxVolume As Double
    MaxVolume = 0
    
'**********************************************************************************

'Set Summary Table and Last Row....................................................
    
Dim SummaryRow As Long
    SummaryRow = 2
    
Dim OpenPriceRow As Long
    OpenPriceRow = 2
    
        'Dim LastRow As Double
            LastRow = MainWrksht.Cells(Rows.Count, "A").End(xlUp).Row
                YearOpen = MainWrksht.Cells(OpenPriceRow, 3).Value
                
        MainWrksht.Cells(1, 9).Value = "Tickers"
        MainWrksht.Cells(1, 10).Value = "Yearly Change"
        MainWrksht.Cells(1, 11).Value = "Percent Change"
        MainWrksht.Cells(1, 12).Value = "Total Stock Volume"

'Begin Code......................................................................

For i = 2 To LastRow

    If MainWrksht.Cells(i + 1, 1).Value <> MainWrksht.Cells(i, 1).Value Then
        
        Ticker = MainWrksht.Cells(i, 1).Value
        
            YearClose = MainWrksht.Cells(i, 6).Value
            YearlyChange = YearClose - YearOpen
            
                'Condition for Percent Change Calculation
                
                    If YearOpen <> 0 Then
                    
                        PercentChange = (YearlyChange / YearOpen) * 100
                    
                    End If

                
                'Stock Total Volume
                    StockTotalVolume = StockTotalVolume + MainWrksht.Cells(i, 7).Value
                
                'Summary Tables for Column "I" and "J"
                
                    MainWrksht.Range("I" & SummaryRow).Value = Ticker
                    
                    MainWrksht.Range("J" & SummaryRow).Value = YearlyChange
                    
                        'ColorIndex
                            
                            If YearlyChange > 0 Then
                                MainWrksht.Range("J" & SummaryRow).Interior.ColorIndex = 4
                                
                            ElseIf YearlyChange <= 0 Then
                                MainWrksht.Range("J" & SummaryRow).Interior.ColorIndex = 3
                                
                            End If
                            
                            
                'Summary Tables for Column "K" and "L"
                
                    MainWrksht.Range("K" & SummaryRow).Value = (CStr(PercentChange) & "%")
                    
                    MainWrksht.Range("L" & SummaryRow).Value = StockTotalVolume
                    
                        'SummaryRow
                            
                            SummaryRow = SummaryRow + 1
                            
                                StockTotalVolume = 0
                                
                                OpenPriceRow = i + 1
                            
                'Bonus Section...................
                
                If PercentChange > MaxPercent Then
                    MaxPercent = PercentChange
                    MaxTicker = Ticker
                    
                ElseIf PercentChange < MinPercent Then
                    MinPercent = PercentChange
                    MinTicker = Ticker
                    
                End If
                
                If StockTotalVolume > MaxVolume Then
                    MaxVolume = StockTotalVolume
                    MaxVolumeTicker = Ticker
                    
                End If
                            

                    'Reset Values
                        PercentChange = 0
                        StockTotalVolume = 0
                                
                                                
    Else
        
        StockTotalVolume = StockTotalVolume + Cells(i, 7).Value
        
 End If
 
Next i

            MainWrksht.Range("Q2").Value = (CStr(MaxPercent) & "%")
            MainWrksht.Range("Q3").Value = (CStr(MinPercent) & "%")
            MainWrksht.Range("P2").Value = MaxTicker
            MainWrksht.Range("P3").Value = MinTicker
            MainWrksht.Range("Q4").Value = MaxVolume
            MainWrksht.Range("O2").Value = "Greatest % Increase"
            MainWrksht.Range("O3").Value = "Greatest % Decrease"
            MainWrksht.Range("O4").Value = "Greatest Total Volume"
            

Next MainWrksht

End Sub


