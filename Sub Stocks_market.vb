Sub Stocks_market()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

Dim Ticker As String
Dim Yearlyopen As Double
Dim Yearlyclose As Double
Dim Yearlypricechange As Double
Dim yearlypercentchange As Double
Dim Totalstockvol As Variant
Dim Rowcount As Long
Dim Greatestpercentincrease As Double
Dim Greatestpercentdecrease As Double
Dim Greatesttotalvolume As Integer
Dim summary_table As Integer

Dim lastrow As Long



        summary_table = 2

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearlypricechange"
    ws.Range("K1").Value = "yearlypercentchange"
    ws.Range("L1").Value = "Totalstockvol"

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Cells(3, 14).Value = "Greatest%increase"
    ws.Cells(4, 14).Value = "Greatest%decrease"
    ws.Cells(5, 14).Value = "Greatesttotalvolume"
    ws.Range("o1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
   
                         
    
For i = 2 To lastrow

    Totalstockvol = Totalstockvol + ws.Cells(i, 7).Value

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
       

     
  Ticker = ws.Cells(i, 1).Value
  
    ws.Range("I" & summary_table).Value = Ticker
    ws.Range("J" & summary_table).Value = Yearlypricechange
    ws.Range("L" & summary_table).Value = Totalstockvol
    Totalstockvol = 0

  summary_table = summary_table + 1
  
  End If
  
  
    Yearlyopen = ws.Cells(i, 3).Value
    
   Yearlyclose = ws.Cells(i, 6).Value
   
   Yearlypricechange = Yearlyclose - Yearlyopen
   
   If (Yearlyopen = 0 And Yearlyclose = 0) Then
   yearlypercentchange = 0
   
    ElseIf Yearlyopen = 0 And Yearlyclose <> 0 Then
    yearlypercentchange = -1
    
    Else: yearlypercentchange = (Yearlypricechange / Yearlyopen)
    
    ws.Range("K" & summary_table).Value = yearlypercentchange
    ws.Range("K" & summary_table).NumberFormat = "0.00%"
    
    If ws.Range("j" & summary_table).Value >= 0 Then
    ws.Range("j" & summary_table).Interior.ColorIndex = 4
    
    Else
    ws.Range("j" & summary_table).Interior.ColorIndex = 3
    End If
    
    
    
     
     End If
     
     
     
     Next i
     
    
     
     
     
     Next ws
     
     
    End Sub

