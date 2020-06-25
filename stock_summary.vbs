Sub stock_summary():

Dim ws_count As Integer
Dim ws As Integer

' Set WS_Count equal to the number of worksheets in the active workbook.
ws_count = ActiveWorkbook.Worksheets.Count

' Loop through worksheets
For ws = 1 To ws_count
    'Change worksheet upon new loop
    worksheet_name = ActiveWorkbook.Worksheets(ws).Name
    Worksheets(worksheet_name).Activate

    ' Set colnames for summary table
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change(USD)"
    Range("K1") = "Yearly Change(%)"
    Range("L1") = "Total Stock Volume"
    ' Remove unnecessary gridlines
    Range("I1:L1").Interior.ColorIndex = 2
    
    ' Set index names for standout table
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    Range("N2") = "Top Leader"
    Range("N3") = "Top Laggard"
    Range("N4") = "Most Liquid"
    
    ' Remove unnecessary gridlines
    Range("N1:P1").Interior.ColorIndex = 2
    Range("N2:O4").Interior.ColorIndex = 2
    
    ' Color Leader and Laggard Table Cells
    Range("P2").Interior.ColorIndex = 4
    Range("P3").Interior.ColorIndex = 3
    
    ' Add col index borders
    With Range("I1:L1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    
    ' Setup ticker var
    Dim ticker As String
    ' Setup price vars
    Dim initial_open As Double
    Dim final_close As Double
    ' Setup volume var
    Dim total_vol As Long
    total_vol = 0
    
    ' Setup and determine final iterator
    Dim last_iter As Long
    last_iter = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    ' Setup table row indexer
    Dim table_row_index As Integer
    table_row_index = 2
    
    ' Set first stock's first open price
    initial_open = Cells(2, 3).Value
    
    ' Setup greatest leader
    Dim top_leader As Single
    Dim top_leader_value As Double
    
    top_leader_value = 0
    
    ' Setup greatest laggard
    Dim top_laggard As Single
    Dim top_laggard_value As Double
    
    top_laggard_value = 0
    
    ' Setup highest liquidity
    Dim most_liquid As Single
    Dim most_liquid_value As Long
    
    most_liquid_value = 0
    
    ' Setup temporary holding variables for price and volume
    Dim price_change As Double
    Dim liquidity As Long
    
    ' Loop through stock data for table index Yearly Change(USD), Yearly Change(%), and Total Volume
    For I = 2 To last_iter:
        total_vol = 0
        ' Conditional check for ticker change
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            
        ' Take in ticker
        ticker = ActiveSheet.Cells(I, 1).Value
            
        ' Take in final price and volume
        total_vol = total_vol + Cells(I, 7).Value
        final_close = Cells(I, 6).Value
            
            If initial_open = 0 Then
            ' take precaution for stock that did not trade
                
                Range("I" & table_row_index).Value = ticker
                
                ' Eliminate uneeded gridlines
                Range("I" & table_row_index).Interior.ColorIndex = 2
                Range("J" & table_row_index).Interior.ColorIndex = 2
                Range("K" & table_row_index).Interior.ColorIndex = 2
                Range("L" & table_row_index).Interior.ColorIndex = 2
                
                ' borders
                With Range("I" & table_row_index).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                End With
                With Range("I" & table_row_index).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                
                Range("J" & table_row_index).Value = "no activity"
                Range("K" & table_row_index).Value = "no activity"
                Range("L" & table_row_index).Value = "no activity"
                
                'borders
                With Range("L" & table_row_index).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                End With
                With Range("L" & table_row_index).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                
                ' Set new table row index
                table_row_index = table_row_index + 1
                
                ' Reset initial price and volume
                initial_open = Cells(I + 1, 3).Value
                total_vol = 0
                
            Else
                ' Print Values
                Range("I" & table_row_index).Value = ticker
                
                ' Eliminate uneeded gridlines
                Range("I" & table_row_index).Interior.ColorIndex = 2
                Range("L" & table_row_index).Interior.ColorIndex = 2
                
                ' borders
                With Range("I" & table_row_index).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                End With
                With Range("I" & table_row_index).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                
                Range("J" & table_row_index).Value = final_close - initial_open
                Range("K" & table_row_index).Value = final_close / initial_open - 1
                Range("L" & table_row_index).Value = total_vol
                
                price_change = Range("K" & table_row_index).Value
                liquidity = Range("L" & table_row_index).Value
                
                ' test for standout table
                If price_change > top_leader_value Then
                    Range("O2").Value = ticker
                    Range("P2").Value = price_change
                    top_leader_value = Range("P2").Value
                
                ElseIf price_change < top_laggard_value Then
                    Range("O3").Value = ticker
                    Range("P3").Value = price_change
                    top_laggard_value = Range("P3").Value
                
                ElseIf liquidity > most_liquid_value Then
                    Range("O4").Value = ticker
                    Range("P4").Value = liquidity
                    most_liquid_value = Range("P4").Value
                    
                End If
                
                'borders
                With Range("L" & table_row_index).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                End With
                With Range("L" & table_row_index).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                
                ' Format Price Cells
                If final_close > initial_open Then
                
                    Range("J" & table_row_index).Interior.ColorIndex = 4
                    Range("K" & table_row_index).Interior.ColorIndex = 4
                    
                Else
                
                    Range("J" & table_row_index).Interior.ColorIndex = 3
                    Range("K" & table_row_index).Interior.ColorIndex = 3
                    
                End If
            
                ' Set new table row index
                table_row_index = table_row_index + 1
            
                ' Reset initial price and volume
                initial_open = Cells(I + 1, 3).Value
                total_vol = 0
            
            End If
        
        Else
            
            total_vol = total_vol + Cells(I, 7).Value
        
        End If
        
    Next I
    ' Format Columns
    Columns("I:L").AutoFit
    Range("K:K").NumberFormat = "0.000%"
    Range("L:L").NumberFormat = "#,##0"
    Range("J:K").HorizontalAlignment = Excel.Constants.xlCenter
    ' Add Bottom Border
    With Range("I" & table_row_index).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Range("J" & table_row_index).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Range("K" & table_row_index).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Range("L" & table_row_index).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    ' Format Columns
    Columns("N:P").AutoFit
    Range("P2:P3").NumberFormat = "0.000%"
    Range("P4").NumberFormat = "#,##0"
    Range("O1:P4").HorizontalAlignment = Excel.Constants.xlCenter
    ' Add borders for standout table
    With Range("N2:N4").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Range("P2:P4").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Range("P2:P4").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Range("N4:P4").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Range("N2:P2").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With Range("N2:N4").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
Next ws


End Sub