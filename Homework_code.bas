Attribute VB_Name = "Module2"
Sub MultiYearStock()
Dim LastRow As Long
Dim FirstTickerRow As Long
Dim NextTickerRow As Long
Dim Ticker As Long
Dim TickerClose As Double
Dim TickerOpen As Double
Dim Volume As Double
Dim TickerName As String
Dim NextTickerName  As String
Dim Total As Integer

Total = ThisWorkbook.Sheets.Count

For j = 1 To Total

  Sheets(j).Activate
  
  Ticker = 1
  FirstTickerRow = 2

  Range("I1").Value = "Ticker"

  Range("J1").Value = "Yearly Change"

  Range("K1").Value = "Percentage Change"

  Range("L1").Value = "Total Stock Volume"

  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  

  Volume = Cells(2, 7).Value

  For i = 2 To LastRow

    TickerName = Cells(i, 1).Value

    NextTickerName = Cells(i + 1, 1).Value

    If TickerName = NextTickerName Then
      
      Volume = Volume + Cells(i + 1, 7).Value

    Else

       Ticker = Ticker + 1

       NextTickerRow = i + 1

       Cells(Ticker, 9).Value = TickerName

       Cells(Ticker, 12).Value = Volume

       TickerClose = Cells(NextTickerRow - 1, 6).Value

       TickerOpen = Cells(FirstTickerRow, 3).Value
      
       While TickerOpen = 0

             FirstTickerRow = FirstTickerRow + 1

             TickerOpen = Cells(FirstTickerRow, 3).Value

       Wend

       Cells(Ticker, 10).Value = TickerClose - TickerOpen
  
       Cells(Ticker, 11).Value = Cells(Ticker, 14).Value / TickerOpen
  
       Cells(Ticker, 11).NumberFormat = "0.00%"

       If Cells(Ticker, 10).Value > 0 Then

            Cells(Ticker, 10).Interior.ColorIndex = 4

        Else

            Cells(Ticker, 10).Interior.ColorIndex = 3

        End If
    
      FirstTickerRow = NextTickerRow

      Volume = Cells(FirstTickerRow, 7).Value
      
    End If

 Next i

 Range("I1:L1").Columns.AutoFit

Next j

End Sub

Sub Sort(vArray As Variant, arrLbound As Double, arrUbound As Double)

'Small to larg

Dim pivotValue As Variant
Dim trade    As Variant
Dim Low   As Double
Dim High    As Double

Low = arrLbound
High = arrUbound

pivotValue = vArray((arrLbound + arrUbound) \ 2)

While (Low <= High) 'divide

   While (vArray(Low) < pivotValue And Low < arrUbound)

      Low = Low + 1

   Wend
  
   While (pivotValue < vArray(High) And High > arrLbound)

      High = High - 1

   Wend

   If (Low <= High) Then

      trade = vArray(Low)

      vArray(Low) = vArray(High)

      vArray(High) = trade

      Low = Low + 1

      High = High - 1

   End If

Wend

  If (arrLbound < High) Then Sort vArray, arrLbound, High

  If (Low < arrUbound) Then Sort vArray, Low, arrUbound

End Sub

