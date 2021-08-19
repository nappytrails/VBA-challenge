Sub StockAnalysis()

Dim worksheetCount As Integer
Dim i As Integer

' Set worksheetCount equal to the number of worksheets in the active workbook.
         
worksheetCount = ActiveWorkbook.Worksheets.Count

' Begin the loop.
For i = 1 To worksheetCount
         
    ActiveWorkbook.Worksheets(i).Activate


' Establish summary table header values

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "% Change"
Cells(1, 12).Value = "Total Stock Volume"

' Make the font bold for the top row

Range("1:1").Font.Bold = True

' Freeze top row

With ActiveWindow
    If .FreezePanes Then .FreezePanes = False
    .SplitColumn = 0
    .SplitRow = 1
    .FreezePanes = True
End With


' Declare variables

Dim firstOpeningPrice As Double
Dim lastClosingPrice As Double
Dim summaryTableTicker As String
Dim summaryTableYearChange As Double
Dim summaryTablePercentChange As Double
Dim summaryTableTotalVolume As LongLong
Dim summaryTableRow As Long
Dim lastRow As Long

' Establish values for variables

firstOpeningPrice = Cells(2, 3).Value
summaryTableTicker = Cells(2, 1).Value
summaryTableTotalVolume = 0
summaryTableRow = 2
lastRow = Cells(Rows.Count, 1).End(xlUp).Row


For Row = 2 To lastRow

    If Cells(Row, 1).Value = summaryTableTicker Then
        summaryTableTotalVolume = summaryTableTotalVolume + Cells(Row, 7).Value
                 
        lastClosingPrice = Cells(Row, 6).Value
    
    End If
    
    
    If Cells(Row, 1).Value <> summaryTableTicker Then
        
        ' Enter ticker symbol in summary table
        
        Cells(summaryTableRow, 9).Value = summaryTableTicker
         
        
        ' Calculate yearly change from opening price at the beginning of a given year
        ' to the closing price at the end of that year
        
        summaryTableYearChange = lastClosingPrice - firstOpeningPrice
        
        ' Enter yearly change in summary table
                
        Cells(summaryTableRow, 10).Value = summaryTableYearChange
        
            
            ' Insert contitional formatting for Cells(summaryTableRow, 10):
            ' Color cell green if summaryTableYearChange is >0
            ' Color cell red if summaryTableYearChange is < 0
       
            If summaryTableYearChange > 0 Then
                    
                Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0)
            
            ElseIf summaryTableYearChange < 0 Then
        
                Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0)
            
            End If
        
        
        ' Calculate percentage change
        
        If firstOpeningPrice <> 0 Then
        
            summaryTablePercentChange = Round(((summaryTableYearChange / firstOpeningPrice) * 100), 2)
        
            ' Enter percentage change in summary table
        
            Cells(summaryTableRow, 11).Value = summaryTablePercentChange & "%"
                    
        Else
       
            Cells(summaryTableRow, 11).Value = "0%"
        
        End If
        
        ' Enter total volume of the stock in summary table
        
        Cells(summaryTableRow, 12).Value = summaryTableTotalVolume
        
                
        ' Reset variables for next stock
        
        firstOpeningPrice = Cells(Row, 3).Value
        summaryTableTicker = Cells(Row, 1).Value
        summaryTableTotalVolume = Cells(Row, 7)
        summaryTableRow = summaryTableRow + 1
        

    End If

Next Row

' Autofit columns I through L

Range("I:L").Columns.AutoFit


' Enter column headers for analysis table

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

' Enter row names for analysis table

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

' Make row names bold

Range("O2:O4").Font.Bold = True

' Declare variables for analysis table

Dim lastRowSummaryTable As Integer
Dim percentChangeColumn As Range
Dim totalStockVolumeColumn As Range
Dim maxPercentChange As Variant
Dim maxPercentChangeLocation As Range
Dim minPercentChange As Variant
Dim minPercentChangeLocation As Range
Dim maxTotalStockVolume As LongLong
Dim maxTotalStockVolumeLocation As Range

lastRowSummaryTable = Cells(Rows.Count, 9).End(xlUp).Row
Set percentChangeColumn = Range("K2:K" & lastRowSummaryTable)
Set totalStockVolumeColumn = Range("L2:L" & lastRowSummaryTable)


' Find value and ticker symbol of gretest percent increase

    ' Find and enter greatest percent increase
    
    maxPercentChange = WorksheetFunction.Max(percentChangeColumn) * 100 & "%"
        
    Range("Q2") = maxPercentChange
    
    ' Find and enter ticker symbol
    
    Set maxPercentChangeLocation = percentChangeColumn.Find(maxPercentChange)
    
    Range("P2") = maxPercentChangeLocation.Offset(0, -2).Value
    
    
' Find value and ticker symbol of gretest percent decrease

    ' Find and enter greatest percent decrease
    
    minPercentChange = WorksheetFunction.Min(percentChangeColumn) * 100 & "%"
            
    Range("Q3") = minPercentChange
    
    ' Find and enter ticker symbol
    
    Set minPercentChangeLocation = percentChangeColumn.Find(minPercentChange)
    
    Range("P3") = minPercentChangeLocation.Offset(0, -2).Value
    

' Find value and address of maxTotalStockVolume

    ' Find and enter max total stock volume
    
    maxTotalStockVolume = WorksheetFunction.Max(totalStockVolumeColumn)

    Range("Q4") = maxTotalStockVolume
    
    ' Find and enter ticker symbol
    
    Set maxTotalStockVolumeLocation = totalStockVolumeColumn.Find(maxTotalStockVolume)

    Range("P4") = maxTotalStockVolumeLocation.Offset(0, -3).Value


' Autofit columns O through Q

Range("O:Q").Columns.AutoFit

Next i

End Sub

