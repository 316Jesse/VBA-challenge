Attribute VB_Name = "Module1"
Sub stocks()
    Dim ws As Worksheet
    Dim table_row As Double
    Dim select_row As Double
    Dim last_row As Double
    Dim year_opening As Single
    Dim year_closing As Single
    Dim totalstock As Double

    
    For Each ws In Sheets
        Worksheets(ws.Name).Activate
        table_row = 2
        select_row = 2
        last_row = WorksheetFunction.CountA(ActiveSheet.Columns(1))
        totalstock = 0
        
            ws.Range("I1").EntireColumn.Insert
            ws.Cells(1, 9).Value = "Ticker"
            ws.Range("J1").EntireColumn.Insert
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Range("K1").EntireColumn.Insert
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Range("L1").EntireColumn.Insert
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Range("N1").EntireColumn.Insert
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        For i = 2 To last_row
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker = Cells(i, 1).Value
            totalstock = totalstock + Cells(i, 7).Value
            ws.Range("L" & table_row).Value = totalstock
            ws.Range("I" & table_row).Value = ticker
            table_row = table_row + 1
            totalstock = 0
        Else
        totalstock = totalstock + Cells(i, 7).Value
            End If
         Next i
    
        table_row = 2
        For i = 2 To last_row
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                year_closing = Cells(i, 6).Value
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                year_opening = Cells(i, 3).Value
            End If
            If year_opening > 0 And year_closing > 0 Then
                increase = year_closing - year_opening
                percent_increase = increase / year_opening
                Cells(table_row, 10).Value = increase
                Cells(table_row, 11).Value = FormatPercent(percent_increase)
                year_closing = 0
                year_opening = 0
                table_row = table_row + 1
            End If
        Next i
    
        max_per = WorksheetFunction.Max(ActiveSheet.Columns("k"))
        min_per = WorksheetFunction.Min(ActiveSheet.Columns("k"))
        max_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        
        Range("Q2").Value = FormatPercent(max_per)
        Range("Q3").Value = FormatPercent(min_per)
        Range("Q4").Value = max_vol
    
        For i = 2 To last_row
            If max_per = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
            ElseIf min_per = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
            ElseIf max_vol = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            End If
        Next i
        
        For i = 2 To last_row
            If IsEmpty(Cells(i, 10).Value) Then Exit For
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    Next ws
                
End Sub



