Sub multiple_year_stock_data()

  ' Set an initial variable for holding the Ticker
  Dim ticker As String

  ' Set an initial variable for holding the volume
  Dim volume As Double
  volume = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stocks
  For i = 2 To 760192

    ' Check if we are still within the same stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker
      ticker = Cells(i, 1).Value

      ' Add to the Volume
      volume = volume + Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker

      ' Print the volume to the Summary Table
      Range("J" & Summary_Table_Row).Value = volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the ticker
      ticker = 0

    ' If the cell immediately following a row is the same stock...
    Else

      ' Add to the volume
      volume = volume + Cells(i, 7).Value

    End If

  Next i

End Sub

