Attribute VB_Name = "Module1"
Sub stock()

For Each ws In Sheets
Dim i As Long
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim total_volume As Double
total_volume = 0
Dim percent_change As Double
percent_change = 0

'challenge
Dim gre_in_name As String
gre_in_name = ""
Dim gre_de_name As String
gre_de_name = ""
Dim gre_ine As Double
gre_in = 0
Dim gre_de As Double
gre_de = 0
Dim gre_vol_name As String
gre_vol_name = ""
Dim gre_vol As Double
gre_vol = 0


d = 2
stockcount = 0
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("K1").Value = "Ticker"
ws.Range("l1").Value = "Yearly_change"
ws.Range("M1").Value = "Percent_change"
ws.Range("N1").Value = "Total_stock_volume"
ws.Range("Q2").Value = "Greatest_%_increase"
ws.Range("Q3").Value = "Greatest_%_decrease"
ws.Range("Q4").Value = "Greatest_total_volume"
ws.Range("R1").Value = "Ticker"
ws.Range("S1").Value = "Value"
ws.Range("S2:S3").NumberFormat = "0.00%"

For i = 2 To lastrow
If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
stockcount = stockcount + 1
total_volume = total_volume + ws.Cells(i, 7)

Else
ws.Cells(d, 11).Value = ws.Cells(i, 1).Value
close_price = ws.Cells(i, 6).Value
open_price = ws.Cells(i - stockcount, 3).Value
yearly_change = close_price - open_price
ws.Cells(d, 12).Value = yearly_change
total_volume = total_volume + ws.Cells(i, 7)
ws.Cells(d, 14).Value = total_volume
    If open_price <> 0 Then
    percent_change = (yearly_change / open_price)
    ws.Cells(d, 13).Value = percent_change
      
    Else
    ws.Cells(d, 13).Value = "Null"
    End If
    
        If ws.Cells(d, 13).Value > 0 Then
        ws.Cells(d, 13).Interior.ColorIndex = 4
        Else
        ws.Cells(d, 13).Interior.ColorIndex = 3
        End If
        
      If (percent_change > gre_in) Then
            gre_in = percent_change
            gre_in_name = ws.Cells(d, 11).Value
      ElseIf (percent_change < gre_de) Then
            gre_de = percent_change
            gre_de_name = ws.Cells(d, 11).Value
      End If
        
        If (total_volume > gre_vol) Then
            gre_vol = total_volume
            gre_vol_name = ws.Cells(d, 11).Value
        
        End If
        
      ws.Range("R2").Value = gre_in_name
      ws.Range("R3").Value = gre_de_name
      ws.Range("R4").Value = gre_vol_name
      ws.Range("S2").Value = gre_in
      ws.Range("S3").Value = gre_de
      ws.Range("S4").Value = gre_vol
      
       
d = d + 1
stockcount = 0
total_volume = 0

              
End If
Next i
            
Next ws


End Sub

