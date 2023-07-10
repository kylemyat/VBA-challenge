Attribute VB_Name = "Module1"
Sub ticker()

Dim WS As Worksheet
Application.ScreenUpdating = False
 For Each WS In Worksheets
 WS.Select
 

Dim ticker As String
Dim open_price As Double
open_price = Cells(2, 3).Value

Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim g_increase As Double
g_increase = 0

Dim gi_ticker As String

Dim g_decrease As Double
g_decrease = 0

Dim gd_ticker As String

Dim g_volume As LongLong
g_volume = 0

Dim gv_ticker As String


Dim volume As LongLong
volume = 0

Dim table As Integer
table = 2

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"



For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
 If Cells(i + 1, 1) <> Cells(i, 1) Then
 
  'Ticker Symbol Output
  ticker = Cells(i, 1).Value
  Range("I" & table).Value = ticker
 
  'Yearly Change Output
  close_price = Cells(i, 6).Value
  yearly_change = close_price - open_price
  Range("J" & table).Value = yearly_change
    If yearly_change > 0 Then
    Range("J" & table).Interior.ColorIndex = 4
    Else
    Range("J" & table).Interior.ColorIndex = 3
    End If
    

  'Percent Change Output
  percent_change = yearly_change / open_price
  Range("K" & table).NumberFormat = "0.00%"
  Range("K" & table).Value = percent_change
  
  'Total Stock Volume Output
  volume = volume + Cells(i, 7).Value
  Range("L" & table).Value = volume
  
  
  'Greatest Change
    If g_increase < percent_change Then
    g_increase = percent_change
    Range("Q2").NumberFormat = "0.00%"
    Range("Q2").Value = g_increase
    gi_ticker = Cells(i, 1)
    Range("P2").Value = gi_ticker
    End If
    
    If g_decrease > percent_change Then
    g_decrease = percent_change
    Range("Q3").NumberFormat = "0.00%"
    Range("Q3").Value = g_decrease
    gd_ticker = Cells(i, 1)
    Range("P3").Value = gd_ticker
    End If
    
    If g_volume < volume Then
    g_volume = volume
    Range("Q4").Value = g_volume
    gv_ticker = Cells(i, 1)
    Range("P4").Value = gv_ticker
    End If
    
  volume = 0
 
  'Output Row
 table = table + 1
 
 open_price = Cells(i + 1, 3).Value
 
 Else
  'Sum Volume
 volume = volume + Cells(i, 7).Value
 
 End If
 
 Next i
 
 Next
 


End Sub


