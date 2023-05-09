Attribute VB_Name = "StockDataAnalysis"
Sub StockDataAnalysis()
'Define worksheet and create for/loop to loop through each worksheet
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Select
    
    'Define variables
    Dim Ticker As String
    Dim Volume As Double
    Dim BegPrice As Double
    Dim EndPrice As Double
    Dim Change As Double
    Dim Percent As Double
    Dim RowCounter As Integer
    Dim LastRow As Double
    Dim LastRowCounter As Integer
    
    'Define last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Set starting values
    Volume = 0
    RowCounter = 2
    BegPrice = Cells(2, 3).Value
    
    'Sort data to make sure Ticker and Date are in descending order
    ActiveSheet.Sort.SortFields.Clear
    Range("A1:G" & LastRow).Sort key1:=Range("A1"), key2:=Range("B1"), Header:=xlYes, order1:=xlAscending, order2:=xlAscending
    
    'If Ticker changes
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set ticker value
            Ticker = Cells(i, 1).Value
        
            'Copy ticker value to table
            Cells(RowCounter, 9).Value = Ticker
        
            'Set closing price
            EndPrice = Cells(i, 6).Value
        
            'Calculate yearly price change
            Change = EndPrice - BegPrice
        
            'Calculate percentage change
            Percent = Change / BegPrice
        
            'Copy yearly price change and percentage changes to table
            Cells(RowCounter, 10).Value = Change
            Cells(RowCounter, 11).Value = Percent
        
            'record new beginning price
            BegPrice = Cells(i + 1, 3).Value
        
            'Record Stock Volume
            Volume = Volume + Cells(i, 7).Value
        
            'Copy Stock Volume to Table
            Cells(RowCounter, 12).Value = Volume
        
            'Add one to RowCounter
            RowCounter = RowCounter + 1
        
            'Reset Volume total
            Volume = 0
    
            'If Ticker hasn't changed
            Else
        
            'Add daily volume to Volume total
            Volume = Volume + Cells(i, 7).Value
        
            End If
    
    Next i
    
    'Define lastrowcounter
    LastRowCounter = Cells(Rows.Count, 9).End(xlUp).Row
        
    'Conditional formatting for gains and losses
    For j = 2 To LastRowCounter
        
        'check to make sure that yearly price and % moves are the same direction (ensure no data sorting issues)
        If Cells(j, 10).Value > 0 And Cells(j, 11).Value < 0 Then
        MsgBox ("Yearly Price Change and Yearly % Change are the opposite direction.  Error in Data!")
            
        'change the fill of cells with yearly price and % gains to green
        ElseIf Cells(j, 10).Value > 0 And Cells(j, 11).Value > 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
            Cells(j, 11).Interior.ColorIndex = 4
        'change formatting of cells to number for yearly price change and % for yearly % change
            Cells(j, 10).Value = FormatNumber(Cells(j, 10), 2)
            Cells(j, 10).Value = FormatPercent(Cells(j, 10), 2)
            
        'change the fill of cells with yearly price and % losses to red
        Else
            Cells(j, 10).Interior.ColorIndex = 3
            Cells(j, 11).Interior.ColorIndex = 3
        'change formatting of cells to number for yearly price change and % for yearly % change
            Cells(j, 10).Value = FormatNumber(Cells(j, 10), 2)
            Cells(j, 10).Value = FormatPercent(Cells(j, 10), 2)
        
        End If
        
    Next j
        
    'Locate and copy largest movers to 2nd table
    
    'define largest movers
    Dim LargestInc As Double
    Dim LargestDec As Double
    Dim LargestVol As Double
        
    'set largest movers values
    LargestInc = WorksheetFunction.Max(Range("K2:K" & LastRowCounter))
    LargestDec = WorksheetFunction.Min(Range("K2:K" & LastRowCounter))
    LargestVol = WorksheetFunction.Max(Range("L2:L" & LastRowCounter))
    
    'Copy largest movers to 2nd table and format cells
    Range("Q2").Value = LargestInc
    Range("Q2").Value = FormatPercent(Range("Q2"), 2)
    Range("Q3").Value = LargestDec
    Range("Q3").Value = FormatPercent(Range("Q3"), 2)
    Range("Q4").Value = LargestVol
    Range("Q4").Value = FormatNumber(Range("Q4"))
    
    'loop through ticker summary data and copy tickers to 2nd table
    For k = 2 To LastRowCounter
            
        If Cells(k, 11).Value = LargestInc Then
            Range("P2").Value = Cells(k, 9).Value
            
        ElseIf Cells(k, 11).Value = LargestDec Then
            Range("P3").Value = Cells(k, 9).Value
            
        ElseIf Cells(k, 12).Value = LargestVol Then
            Range("P4").Value = Cells(k, 9).Value
        
        End If
    
    Next k

Next ws

End Sub


