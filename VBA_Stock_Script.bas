Attribute VB_Name = "Module1"
Sub alphabetical_testing()
  Const FIRST_DATA_ROW As Integer = 2
  Const CLOSING_COL As Integer = 6
  Const OPEN_COL As Integer = 3
  Const VOLUME_COL As Integer = 7
  Const TICKER_COL As Integer = 1
  'Set an initial variable for holding the brand name
  Dim stock As String
  Dim openprice As Double
  Dim closingprice As Double
  Dim yearly_change As Double
  Dim percentage_change As Double
  'Set the total volume of the stock starting at 0
  Dim volume As Double
  
  'adding headers
  Range("i1").Value = "ticker"
  Range("j1").Value = "yearly change"
  Range("k1").Value = "percentage change"
  Range("l1").Value = "stock volume"

  'Keep track of the location for each stock in the summary table
  Dim output_row As Integer
  output_row = FIRST_DATA_ROW

  lastRow = Cells(Rows.Count, TICKER_COL).End(xlUp).Row
  'Loop through all stocks volume
  For input_row = FIRST_DATA_ROW To lastRow
      'Process first row of ticker,
    If Cells(input_row - 1, TICKER_COL).Value <> Cells(input_row, TICKER_COL).Value Then
       stock = Cells(input_row, TICKER_COL).Value
       openprice = Cells(input_row, OPEN_COL).Value
       volume = 0
    End If
    
    'every row add the volume
    volume = volume + Cells(input_row, VOLUME_COL).Value
    
    'Process last row of ticker,
    If Cells(input_row + 1, TICKER_COL).Value <> Cells(input_row, TICKER_COL).Value Then
      closingprice = Cells(input_row, CLOSING_COL).Value
      yearly_change = closingprice - openprice
      percentage_change = (Range("F" & output_row).Value - Range("C" & output_row).Value) / Range("C" & output_row).Value * 100
     'Print the output
      Range("I" & output_row).Value = stock
      Range("L" & output_row).Value = volume
      Range("J" & output_row).Value = yearly_change
      Range("J" & output_row).Style = "Currency"
      Range("K" & output_row).Value = percentage_change
       Range("K" & output_row).NumberFormat = "0.00%"
      If yearly_change > 0 Then
        Range("J" & output_row).Interior.ColorIndex = 4    'Green
        Range("k" & output_row).Interior.ColorIndex = 4    'Green
     ElseIf yearly_change < 0 Then
         Range("J" & output_row).Interior.ColorIndex = 3    'Red
         Range("k" & output_row).Interior.ColorIndex = 3    'Red
     End If

       'prepare for the next row of output
      output_row = output_row + 1
    End If
  Next input_row
  

End Sub

