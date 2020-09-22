Sub Ticker()

  ' Set an initial variable for holding the Ticker name
  Dim Ticker_Name As String

  ' Set an initial variables for holding the total per Column
  Dim Ticker As Double
  Ticker = 0
  Dim StartValue As Double
  StartValue = 0
  Dim EndValue As Double
  EndValue = 0
  Dim YearChange As Double
  YearChange = 0
  Dim PercentageChange As Double
  PercentageChange = 0
  Dim ws As Worksheet
  
  For Each ws In Worksheets
  
  
  ' Keep track of the location for each Ticker brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Dim Start As Long
  Start = 2
  
  'Counts the number of rows
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
  ' Loop through each row
  For i = 2 To lastrow

    ' Check if we are still within the same Ticker brand, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
      
           
       ' Add to the StartValue Total
      StartValue_Total = ws.Cells(Start, 3)
      
       ' Add to the EndValue Total
      EndValue_Total = ws.Cells(i, 6)
      
      ' Calaculate the YearChange Total
      YearChange = EndValue_Total - StartValue_Total
            
      ' Calaculate the PercentageChange Total
      'PercentageChange = YearChange / StartValue_Total
           
      If YearChange <> 0 And StartValue_Total <> 0 Then
         PercentageChange = YearChange / StartValue_Total * 100
         Else
           PercentageChange = 0
         End If
                  PercentageChange = Round(PercentageChange, 2)

      ' Print the Credit Card Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      ' Print the YearChange Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = YearChange
      
      ' Print the PercentageChange Amount to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = PercentageChange

      ' Print the Ticker Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
      
      'Formate positive changes green and negative changes red
      
      If YearChange > 0 Then
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      
      Else
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      
      End If

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Ticker_Total = 0
      Start = i + 1

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Brand Total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
   End If

  Next i
  
  Next ws

End Sub





