Sub vba_challenge()

Dim WS_Count As Integer
Dim Z As Integer
Dim Stock_Name As String
Dim lastrow, lastrow2 As Double
Dim Open_abs As Double
Dim High_abs As Double
Dim Low_abs As Double
Dim Close_abs As Double
Dim Volume_abs As Double
Dim Open_Change As Double
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Volume_Change As Integer
Dim Open_first As Double
Dim Close_last As Double
Dim Date_abs As String
Dim Inc_stock, Vol_stock As String
Dim Inc_value, Vol_value As Double
Dim Ticker_label, Yearly_label, Percent_label, Volume_label, Value_label, Great_inc, Great_dec, Great_vol As String


' Populate all summary table column and row labels
Ticker_label = "Ticker"
Yearly_label = "Yearly Change"
Percent_label = "Percent Change"
Volume_label = "Total Stock Volume"
Value_label = "Value"
Great_inc = "Greatest % Increase"
Great_dec = "Greatest % Decrease"
Great_vol = "Greatest Total Volume"


Range("I1").Value = Ticker_label
Range("J1").Value = Yearly_label
Range("K1").Value = Percent_label
Range("L1").Value = Volume_label
Range("Q1").Value = Ticker_label
Range("R1").Value = Value_label
Range("P2").Value = Great_inc
Range("P3").Value = Great_dec
Range("P4").Value = Great_vol

  ' Keep track of the location for each stock symbol in the summary table
  Dim Table_Row As Integer
  Table_Row = 2

  ' Loop through all daily stock results
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  For I = 2 To lastrow

    ' Check if we are still within the same stock symbol; if not, sum up findings
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
 
      ' Establish last closing price date, since this is the last line of a particular symbol's section
      Close_last = Cells(I, 6).Value


      ' Set the Stock_Name
      Stock_Name = Cells(I, 1).Value
      
            ' Add to Open_abs
            Open_abs = Open_abs + Cells(I, 3).Value
            
               ' Add to High_abs
            High_abs = High_abs + Cells(I, 4).Value
            
                 ' Add to Low_abs
            Low_abs = Low_abs + Cells(I, 5).Value
            
                ' Add to Close_abs
            Close_abs = Close_abs + Cells(I, 6).Value
            
             ' Add to Volume_abs
            Volume_abs = Volume_abs + Cells(I, 7).Value
            
         ' Add to the Yearly Change
      Yearly_Change = Close_last - Open_first
      

        WS_Count = ActiveWorkbook.Worksheets.Count

                  For Z = 1 To WS_Count
     
   ' Print the Stock_Name to the Summary Table
      Range("I" & Table_Row).Value = Stock_Name

      ' Print the Yearly_Change to the Summary Table
      Range("J" & Table_Row).Value = Yearly_Change
      
      ' Create conditional formatting, so that yearly change has a red background for negative amounts and a green one for positive
      If Yearly_Change < 0 Then
      Range("J" & Table_Row).Interior.ColorIndex = 3
      Else: Range("J" & Table_Row).Interior.ColorIndex = 4
      End If
      
   
        ' Print Volume_abs to the Summary Table
      Range("L" & Table_Row).Value = Volume_abs
      
         ' Create the Percentage_Change
      Percentage_Change = Yearly_Change / Open_first
      
         ' Print the Percentage_Change to the Summary Table
      Range("K" & Table_Row).Value = FormatPercent(Percentage_Change)

Next Z

      ' Add one to the summary table row
      Table_Row = Table_Row + 1
      

      
      'Reset Values, since you'll now be starting a new symbol's section
Open_abs = 0
High_abs = 0
Low_abs = 0
Close_abs = 0
Volume_abs = 0
Open_Change = 0
Yearly_Change = 0
Percentage_Change = 0
Volume_Change = 0

    ' BUT If the cell immediately following a row has the same stock symbol as the one before it:
    Else

         ' Add to Open_abs
            Open_abs = Open_abs + Cells(I, 3).Value
            
               ' Add to High_abs
            High_abs = High_abs + Cells(I, 4).Value
            
                 ' Add to Low_abs
            Low_abs = Low_abs + Cells(I, 5).Value
            
                ' Add to Close_abs
            Close_abs = Close_abs + Cells(I, 6).Value
            
             ' Add to Volume_abs
            Volume_abs = Volume_abs + Cells(I, 7).Value
            
         ' Add to the Yearly Change
      Yearly_Change = Close_abs - Open_abs
      
         ' If the date  is the first market day of the year (Jan 02), record it as the first opening price
         If Right(Cells(I, 2), 4) = "0102" Then
      Open_first = Cells(I, 3).Value
      
      End If

      

    End If


  Next I
  

  
    ' Now that the original summary tanble is done, loop through results to create another summary table of superlatives
  lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
  For j = 2 To lastrow2


    ' Calculate and assign the highest percent change and its stock symbol
    If Cells(j, 11).Value = Application.WorksheetFunction.Max(Range("K:K")) Then
    Inc_value = Cells(j, 11).Value
    Inc_stock = Cells(j, 9).Value
    Range("R2").Value = FormatPercent(Inc_value)
    Range("Q2").Value = Inc_stock
    
    
    ' Calculate and assign the lowest percent change and its stock symbol
    ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Min(Range("K:K")) Then
    Inc_value = Cells(j, 11).Value
    Inc_stock = Cells(j, 9).Value
    Range("R3").Value = FormatPercent(Inc_value)
    Range("Q3").Value = Inc_stock
    
          ' Calculate and assign the greatest total colume and its stock symbol
      ElseIf Cells(j, 12).Value = Application.WorksheetFunction.Max(Range("L:L")) Then
    Vol_value = Cells(j, 12).Value
    Vol_stock = Cells(j, 9).Value
    Range("R4").Value = Vol_value
    Range("Q4").Value = Vol_stock
    

     
     End If
     
     Next j
  

End Sub
