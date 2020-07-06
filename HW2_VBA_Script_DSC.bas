Attribute VB_Name = "Module4"
Sub MasterSolution_VBA_Homework_DSC()

Dim ws As Worksheet
' iterate through all worksheets
   For Each ws In ActiveWorkbook.Worksheets
   ws.Activate
   
   ' GENERAL SECTION
   ' Variables
       Dim last_row As Long
       Dim ticker As String
       Dim ticker_index As Integer
       Dim header_index As Integer
       Dim open_price As Double
       Dim close_price As Double
       Dim yearly_change As Double
       Dim percent_change As Double
       Dim total_volume As Double
       Dim GPI_value As Double
       Dim GPI_ticker As String
       Dim GPD_value As Double
       Dim GPD_ticker As String
       Dim GTV_value As Double
       Dim GTV_ticker As String
       
       last_row = Cells(Rows.Count, 1).End(xlUp).Row
       ticker_index = 0
       header_index = 2
       total_volume = 0
       
    ' determine open price on first day and close price on last day in year
    
   ' summarize data by ticker
   ' ***PLEASE READ: The Homework says to compute the values as of beginning of year and end of year, meaning first day (20160101) and last day (20161230)
    For I = 2 To last_row
        ' designate first day of year
            If Right(Str(Cells(I, 2).Value), 4) = "0101" Then
                open_price = Cells(I, 3).Value
                ticker = Cells(I, 1).Value
            End If
        ' designate last day of year
            If Right(Str(Cells(I, 2).Value), 4) = "1231" Then
                close_price = Cells(I, 6).Value
                ticker = Cells(I, 1).Value
            ElseIf Right(Str(Cells(I, 2).Value), 4) = "1230" Then
                close_price = Cells(I, 6).Value
                ticker = Cells(I, 1).Value
            End If
           ' compute total volume per ticker
           total_volume = total_volume + Cells(I, 7)
           ' compute yearly change per ticker
           If ticker <> Cells(I + 1, 1).Value Then
               yearly_change = close_price - open_price
           ' computer percentage change
           If open_price <> 0 Then
               percent_change = yearly_change / open_price
           Else
               percent_change = 100
       End If
           
       ' input tickers, yearly change, percentage change, and total stock volume
           Range("I" & header_index).Value = ticker
           Range("J" & header_index).Value = yearly_change
           Range("K" & header_index).Value = percent_change
           Range("L" & header_index).Value = total_volume
           
       ' formatting
           
           ' negative values are red
           If yearly_change < 0 Then
               Range("J" & header_index).Interior.ColorIndex = 3
           ' zero values and positive values are green
           Else
               Range("J" & header_index).Interior.ColorIndex = 4
           End If
                      
           ' format percentage change as percent
           Range("K:K").NumberFormat = "0.00%"
           Cells(2, 17).NumberFormat = "0.00%"
           Cells(3, 17).NumberFormat = "0.00%"
           
           ' add headers
           Range("I1").Value = "Ticker"
           Range("J1").Value = "Yearly Change"
           Range("K1").Value = "Percentage Change"
           Range("L1").Value = "Total Stock Volume"
           
           ' move header_index to the next row
           header_index = header_index + 1
           
           ' reset the open_price and ticker
           open_price = Cells(I + 1, 3).Value
           ticker = Cells(I + 1, 1).Value
           total_volume = 0
           
       End If
       Next I
   
   ' CHALLENGES SECTION
   ' stats table, far right
   
       ' set table headers
      Cells(2, 15).Value = "Greatest % Increase"
      Cells(3, 15).Value = "Greatest % Decrease"
      Cells(4, 15).Value = "Greatest Total Volume"
      Cells(1, 16).Value = "Ticker"
      Cells(1, 17).Value = "Value"
         
       ' input ticker and value for Greatest % Increase (GPI)
           GPI_value = WorksheetFunction.Max(Columns("K"))
           Cells(2, 17).Value = GPI_value
           For I = 2 To last_row
            For j = 9 To 12
                If Cells(I, j) = GPI_value Then
                    GPI_ticker = Cells(I, 9)
                End If
            Next j
        Next I
            Cells(2, 16).Value = GPI_ticker

       ' input ticker and value for Greatest % Decrease (GPD)
           GPD_value = WorksheetFunction.Min(Columns("K"))
           Cells(3, 17).Value = GPD_value
           For I = 2 To last_row
            For j = 9 To 12
                If Cells(I, j) = GPD_value Then
                    GPD_ticker = Cells(I, 9)
                End If
            Next j
        Next I
           Cells(3, 16).Value = GPD_ticker
           
        ' input ticker and value for Greatest Total Volume (GTV)
           GTV_value = WorksheetFunction.Max(Columns("L"))
           Cells(4, 17).Value = GTV_value
           For I = 2 To last_row
            For j = 9 To 12
                If Cells(I, j) = GTV_value Then
                    GTV_ticker = Cells(I, 9)
                End If
            Next j
        Next I
           Cells(4, 16).Value = GTV_ticker
           
       ' auto fit data to cells
           Columns("I:Q").AutoFit
       
Next ws
End Sub
