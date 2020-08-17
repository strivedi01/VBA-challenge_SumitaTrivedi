VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub test()
    Dim ws As Worksheet
    For Each ws In Sheets
        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim row_index As Long
        row_index = 2
        Dim iterator As Long
        iterator = 2
        Dim ticker_placer As Long
        ticker_placer = 2
        Dim opening_price As Double
        Dim closing_price As Double
        Dim percent_change As Double
        Dim total_stock_volume As Double
        Do While ws.Cells(iterator, 1).Value <> ""
            
            Do While ws.Cells(row_index, 1).Value = ws.Cells(iterator, 1)
                If iterator = row_index Then
                    opening_price = ws.Cells(iterator, 3).Value
                End If
                total_stock_volume = total_stock_volume + ws.Cells(iterator, 7).Value
                iterator = iterator + 1
            Loop
            row_index = iterator
            
            
            closing_price = ws.Cells(row_index - 1, 6).Value
            ws.Cells(ticker_placer, 10).Value = closing_price - opening_price
            If ws.Cells(ticker_placer, 10).Value >= 0 Then
                ws.Cells(ticker_placer, 10).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(ticker_placer, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            If opening_price = 0 Then
                opening_price = 0.0000000001
                ws.Cells(ticker_placer, 11).Value = "0%"
            Else
                percent_change = ((closing_price - opening_price) / opening_price) * 100
                ws.Cells(ticker_placer, 11).Value = CStr(percent_change) & "%"
                ws.Cells(ticker_placer, 13).Value = percent_change
            End If
            
            
            ws.Cells(ticker_placer, 9).Value = ws.Cells(row_index - 1, 1).Value
            
            ws.Cells(ticker_placer, 12).Value = total_stock_volume
            total_stock_volume = 0
            ticker_placer = ticker_placer + 1
        Loop
    Next ws
End Sub




Sub bonus()

    Dim ws As Worksheet
    
    For Each ws In Sheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Dim MyMax As Double
        Dim MyMin As Double
        Dim MyVolume As Double
        
        
        MyMax = Application.WorksheetFunction.Max(ws.Range("M:M"))
        MyMin = Application.WorksheetFunction.Min(ws.Range("M:M"))
        MyVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
        
        
        ws.Range("Q2") = MyMax
        ws.Range("Q3") = MyMin
        ws.Range("Q4") = MyVolume
        
        
        Dim index As Integer
        index = 2
        Dim row_index_max As Integer
        Dim row_index_min As Integer
        Dim row_index_volume As Integer
        
        
        Do While ws.Cells(index, 13).Value <> ""
            If ws.Cells(index, 13).Value = MyMax Then
                row_index_max = index
          
            End If
            If ws.Cells(index, 13).Value = MyMin Then
                row_index_min = index
                
            End If
            If ws.Cells(index, 12).Value = MyVolume Then
                row_index_volume = index
            End If
            index = index + 1
        Loop
        
        ws.Range("P2") = ws.Cells(row_index_max, 9).Value
        ws.Range("P3") = ws.Cells(row_index_min, 9).Value
        ws.Range("P4") = ws.Cells(row_index_volume, 9).Value
        
        ws.Columns(13).ClearContents
    Next ws


   
End Sub

