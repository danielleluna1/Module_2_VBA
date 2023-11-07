VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alphabetical_testing()
'Create a script that loops through all the stocks for one year and outputs the following information:


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Dim row_summary As Integer
Dim last_row As Long
Dim i As Variant
Dim vol As Variant
Dim close_p As Double
Dim open_p As Double

    Dim currStock As String
    Dim total_vol As Double
    Dim nextStock As String
    Dim row As Long
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim ws As Worksheet
            
            Dim lastSummaryRow As Long
            Dim Gr_Vol As Double
            Dim Gr_Vol_Stock As String
            Dim Gr_In As Double
            Dim Gr_In_Stock As String
            Dim Gr_Decr As Double
            Dim Gr_Decr_Stock As String
        

last_row = Cells(Rows.Count, 1).End(xlUp).row

row_summary = 2


For i = 2 To last_row
    vol = vol + Cells(i, 7).Value
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    close_p = Cells(i, 6).Value
    Cells(row_summary, 10).Value = close_p - open_p
    
    
    If open_p = 0 Then
    Cells(row_summary, 11).Value = 0
    
    Else
    Cells(row_summary, 11).Value = Cells(row_summary, 10).Value / open_p
    
        End If

   
    open_p = Cells(i + 1, 3).Value
    Cells(row_summary, 12).Value = vol
    vol = 0
    Cells(row_summary, 9).Value = Cells(i, 1).Value
    row_summary = row_summary + 1
    
    End If
    
        Next i
        
        
' Loop
        For row = 2 To last_row
            currStock = ws.Cells(row, 1).Value
            nextStock = ws.Cells(row + 1, 1).Value
            
' Add total
            totalVolume = totalVolume + ws.Cells(row, 7).Value
        
            If nextStock <> currStock Then
' Close price
                closePrice = ws.Cells(row, 6).Value
                yearlyChange = close_p - open_p
                
' Error handling
            If openPrice = 0 Then
                    percentChange = 0
            Else
                    percentChange = yearlyChange / open_p
            
            End If
            
               
        ws.Cells(summaryRow, 9).Value = currStock
        ws.Cells(summaryRow, 10).Value = yearlyChange
        ws.Cells(summaryRow, 11).Value = percentChange
        ws.Cells(summaryRow, 12).Value = total_vol
        
        
                
                
            If yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 3
            Else
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 4
            End If
                
' reset
                    totalVolume = 0
                    summaryRow = summaryRow + 1
                
' open price
                open_p = ws.Cells(row + 1, 3).Value
            End If
        
        Next row
        
       

        lastSummaryRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
        Gr_Vol = ws.Cells(2, 12).Value
        Gr_Vol_Stock = ws.Cells(2, 9).Value
        
            Gr_In = ws.Cells(2, 11).Value
            Gr_Vol_Stock = ws.Cells(2, 9).Value
        
                Gr_Decr = ws.Cells(2, 11).Value
                Gr_Decr_Stock = ws.Cells(2, 9).Value
        
        For row = 2 To lastSummaryRow
        
            
            If ws.Cells(row, 12).Value > Gr_Vol Then
                Gr_Vol = ws.Cells(row, 12).Value
                Gr_Vol_Stock = ws.Cells(row, 9).Value
            End If
            
            If ws.Cells(row, 11).Value > Gr_In Then
                Gr_In = ws.Cells(row, 11).Value
                Gr_In_Stock = ws.Cells(row, 9).Value
            End If
            
        
            If ws.Cells(row, 11).Value < Gr_Decr Then
                Gr_Decr = ws.Cells(row, 11).Value
                Gr_Decr_Stock = ws.Cells(row, 9).Value
            End If
        
        Next row
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = Gr_In_Stock
        ws.Cells(2, 17).Value = Gr_In
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(3, 16).Value = Gr_Decr_Stock
            ws.Cells(3, 17).Value = Gr_Decr
            ws.Cells(3, 17).NumberFormat = "0.00%"
        
                ws.Cells(4, 15).Value = "Greatest Total Volume"
                ws.Cells(4, 16).Value = Gr_Vol_Stock
                ws.Cells(4, 17).Value = Gr_Vol
        
                    ws.Range("J2:J" & summaryRow).NumberFormat = "0.00"
                    ws.Range("K2:K" & summaryRow).NumberFormat = "0.00%"
        
    
End Sub

