VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockSummary()

Dim i As Integer
Dim j As Long
Dim k As Integer
Dim WS_count As Integer
Dim LastRow As Long
Dim LastRowTable As Long
Dim NewTickerFound As Integer
Dim OldStockStartRow As Long
Dim NewStockStartRow As Long
Dim OpeningValueOldStock As Double
Dim ClosingValueOldStock As Double
Dim YearlyStockValueChange As Double
Dim TotalStockVolume As Variant
Dim MaxVol As Variant
Dim MaxInc As Double
Dim MaxDcr As Double

    'Find the number of sheets
    WS_count = ActiveWorkbook.Worksheets.Count
    
    
    For i = 1 To WS_count
    'For i = 1 To 1
 
        ' Cycle through all the sheets
       
        NewStockStartRow = 2
        LastRow = 0
    
        ' Find the number of rows of data in the active sheet
        LastRow = Worksheets(i).Cells(ActiveWorkbook.Worksheets(i).Rows.Count, "A").End(xlUp).Row
        
        ' Put headings for the additional columns for the summary table on each sheet
        Worksheets(i).Cells(1, 9) = "Ticker"
        Worksheets(i).Cells(1, 10) = "Yearly Change"
        Worksheets(i).Cells(1, 11) = "Percent Change"
        Worksheets(i).Cells(1, 12) = "Total Stock Volume"
               

        NewTickerFound = 0
        NewStockStartRow = 2
    
        For j = 2 To LastRow
                
            'Cycle through all the rows on the sheet
        
            'Find the number of rows on the summary table
            LastRowTable = Worksheets(i).Cells(Rows.Count, 9).End(xlUp).Row
                        
           
            
            'Cycle through all the rows and compare the ticker code in the data row with
            'the ticker code in the next row to find out whether it is a new ticker or one that is already present
                
            If (StrComp(Worksheets(i).Cells(j, 1).Value, Worksheets(i).Cells(j + 1, 1).Value, vbBinaryCompare) <> 0) Then
                
                NewTickerFound = 1
                                    
            Else
            
                 NewTickerFound = 0
                    
            End If
                
           
            
            If NewTickerFound = 1 Then
            
                'LastRowTable = Worksheets(i).Cells(Rows.Count, 9).End(xlUp).Row
                          
                'New ticker found so write the name of the ticker into the next row of the summary table
                Worksheets(i).Cells(LastRowTable + 1, 9) = Worksheets(i).Cells(j, 1)
                               
                OldStockStartRow = NewStockStartRow
                NewStockStartRow = j + 1
                TotalStockVolume = 0
                 
                
                
                    'Opening and closing values of the last ticker
                    OpeningValueOldStock = Worksheets(i).Cells(OldStockStartRow, 3)
                    ClosingValueOldStock = Worksheets(i).Cells(NewStockStartRow - 1, 6)
                    
                    'Find the change over the year
                    YearlyStockValueChange = ClosingValueOldStock - OpeningValueOldStock
                    Worksheets(i).Cells(LastRowTable + 1, 10) = YearlyStockValueChange
                    
                    'Find the percentage change and reformat the column to show %
                    Worksheets(i).Cells(LastRowTable + 1, 11).Value = Format(Round((YearlyStockValueChange / OpeningValueOldStock), 4), "Percent")
                                        
                    If YearlyStockValueChange >= 0 Then
                        'Stock has increased in value or stayed the same - colour green
                        Worksheets(i).Cells(LastRowTable + 1, 10).Interior.ColorIndex = 4
                        Worksheets(i).Cells(LastRowTable + 1, 11).Interior.ColorIndex = 4
                    Else
                        'Stock has decreased in value - colour red
                        Worksheets(i).Cells(LastRowTable + 1, 10).Interior.ColorIndex = 3
                        Worksheets(i).Cells(LastRowTable + 1, 11).Interior.ColorIndex = 3
                    End If
                                        
                    'Sum all the daily values of this stock an put it in the summary table
                    TotalStockVolume = WorksheetFunction.Sum(Range(Worksheets(i).Cells(OldStockStartRow, 7), Worksheets(i).Cells(NewStockStartRow - 1, 7)))
                    Worksheets(i).Cells(LastRowTable + 1, 12) = TotalStockVolume
                               
            End If
            
        Next j
        
        
        MaxInc = 0
        MaxDcr = 0
        MaxVol = 0
        
        'MsgBox ("stop 1")
        
        'Headings for the next summary table
        Worksheets(i).Cells(1, 16).Value = "Ticker"
        Worksheets(i).Cells(1, 17).Value = "Value"
        
        
       'Identify the stock with the greatest % increase over the year
        Worksheets(i).Cells(2, 15).Value = "Greatest % Increase"
        
        MaxInc = WorksheetFunction.Max(Range(Worksheets(i).Cells(2, 11), Worksheets(i).Cells(LastRowTable, 11)))
        Worksheets(i).Cells(2, 17).Value = Format(MaxInc, "Percent")
        
        'Find the index of the stock with the greatest increase and use this to find the sock name
        Ind = WorksheetFunction.Match(Worksheets(i).Cells(2, 17).Value, Range(Worksheets(i).Cells(1, 11), Worksheets(i).Cells(LastRowTable, 11)), 0)
        Worksheets(i).Cells(2, 16).Value = Worksheets(i).Cells(Ind, 9)
        
        
        'Identify the stock with the greatest % decrease over the year
        Worksheets(i).Cells(3, 15).Value = "Greatest % Decrease"
        
        MaxDcr = WorksheetFunction.Min(Range(Worksheets(i).Cells(2, 11), Worksheets(i).Cells(LastRowTable, 11)))
        Worksheets(i).Cells(3, 17).Value = Format(MaxDcr, "Percent")
        'Find the index of the stock with the greatest decrease and use this to find the sock name
        Ind = WorksheetFunction.Match(Worksheets(i).Cells(3, 17).Value, Range(Worksheets(i).Cells(1, 11), Worksheets(i).Cells(LastRowTable, 11)), 0)
        Worksheets(i).Cells(3, 16).Value = Worksheets(i).Cells(Ind, 9)
        
        
        'Identify the stock with the greatest total volume over the year
        Worksheets(i).Cells(4, 15).Value = "Greatest Total Volume"
        
        MaxVol = WorksheetFunction.Max(Range(Worksheets(i).Cells(2, 12), Worksheets(i).Cells(LastRowTable, 12)))
        Worksheets(i).Cells(4, 17).Value = MaxVol
        'Find the index of the stock with the greatest total volume and use this to find the sock name
        Ind = WorksheetFunction.Match(Worksheets(i).Cells(4, 17).Value, Range(Worksheets(i).Cells(1, 12), Worksheets(i).Cells(LastRowTable, 12)), 0)
        Worksheets(i).Cells(4, 16).Value = Worksheets(i).Cells(Ind, 9)
        
        Worksheets(i).Columns("J:L").AutoFit
        Worksheets(i).Columns("O").AutoFit
        Worksheets(i).Columns("Q").AutoFit
        
        'MsgBox ("stop 2")
    
    Next i

End Sub

