Sub Wall_Street()
For Each ws In Worksheets
    
  Dim WorksheetName As String
'Get the WorksheetName
        WorksheetName = ws.Name
        
  ' Set an initial variable for holding the Ticker
  Dim Ticker As String
  
  ' Set Variable to keep track of the location for each Ticker in the Ticker Column
  Dim Summary_ticker As Integer

  ' Set Variable for Last Row of the whole worksheet
  Dim lastrow As Long
  
  ' Set Variable for last row of Calculated columns
  Dim lastrow1 As Double
    
  ' Set Variables for Stock Volume, Yearly Change and Stock Volume
    Dim YChange As Double
    Dim PercentChange As Double
    Dim StockVol As Double
  
  ' Set Variable for Greates Percentage Increased
  Dim MaxPerIn As Double
  
    ' Set Variable for Greates Percentage Decreased
  Dim MinPerIn As Double
  
  'Set Variable for Greatest Total Volum
  Dim TotalVol As Double
        
  'Set Variable for Start Row for each ticker blocks
  Dim STB As Long
  
  ' Adding new Columns
  ws.Cells(1, 8).Value = "Ticker"
  ws.Cells(1, 9).Value = "Yearly Change"
  ws.Cells(1, 10).Value = "Percent Change"
  ws.Cells(1, 11).Value = "Total Stock Vol."
  ws.Cells(1, 14).Value = "Ticker"
  ws.Cells(1, 15).Value = "Value"
  
  ' Set Start ticker block and Summary Ticker to first row
  Summary_ticker = 2
  STB = 2
  
  'Defining last row
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all Rows
    For i = 2 To lastrow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker = ws.Cells(i, 1).Value

      ' Print the Ticker in the Ticker Column
      ws.Range("H" & Summary_ticker).Value = Ticker
      
      ' Calculate Yearly Change and print it in the I column
      YChange = ws.Cells(i, 6).Value - ws.Cells(STB, 3)
      ws.Range("I" & Summary_ticker).Value = YChange
      ' Highlight positive change in green and negative change in red for Yearly Change
      If YChange < 0 Then ws.Range("I" & Summary_ticker).Interior.ColorIndex = 3
      If YChange >= 0 Then ws.Range("I" & Summary_ticker).Interior.ColorIndex = 4
      
      
      ' Calculate Percent Change and print it in the J Column and Format as Percentage
      PercentChange = YChange / ws.Cells(STB, 3)
      ws.Range("J" & Summary_ticker).Value = Format(PercentChange, "Percent")
      
      ' Calculate Stock Vol and print it in K Column
      StockVol = WorksheetFunction.Sum(ws.Range(ws.Cells(STB, 7), ws.Cells(i, 7)))
            ws.Range("K" & Summary_ticker) = StockVol

      ' Add one to the summary Ticker row
      Summary_ticker = Summary_ticker + 1
      STB = i + 1


    ' If the cell immediately following a row is the same ticker...
    'Else
        

    End If
    
  Next i
    ws.Range("H:H").Font.Bold = True
    ws.Range("I:I").Font.Bold = True
    ws.Range("J:J").Font.Bold = True
    ws.Range("K:K").Font.Bold = True
    ws.Range("M:M").Font.Bold = True
    ws.Range("N:N").Font.Bold = True
    ws.Range("O:O").Font.Bold = True

'Defining last row of the calculated columns
  lastrow1 = ws.Cells(Rows.Count, 9).End(xlUp).Row
  
' Calculating Max Percentage Increase
MaxPerIn = WorksheetFunction.Max(ws.Range("J:J"))
ws.Cells(2, 13).Value = "Greates % Increased"
ws.Cells(2, 13).Font.Bold = True
ws.Cells(2, 15).Value = Format(MaxPerIn, "Percent")

'Loop through Percentage Change to get the ticker Name for Greatest Percentage Decrease
For i = 2 To lastrow1
    If ws.Cells(i, 10).Value = MaxPerIn Then
        ws.Cells(2, 14).Value = ws.Cells(i, 8).Value
    End If
Next i
    
' Calculating Greatest Percentage Decreased
MinPerIn = WorksheetFunction.Min(ws.Range("J:J"))
ws.Cells(3, 13).Value = "Greates % Decreased"
ws.Cells(3, 13).Font.Bold = True
ws.Cells(3, 15).Value = Format(MinPerIn, "Percent")

'Loop through Percentage Change to get the ticker Name for Greatest Percentage Decrease
For i = 2 To lastrow1
    If ws.Cells(i, 10).Value = MinPerIn Then
        ws.Cells(3, 14).Value = ws.Cells(i, 8).Value
    End If
Next i

' Calculating Greatest Total Volum
TotalVol = WorksheetFunction.Max(ws.Range("K:K"))
ws.Cells(4, 13).Value = "Greates Volume"
ws.Cells(4, 13).Font.Bold = True
ws.Cells(4, 15).Value = TotalVol

'Loop through total stock Volum to get the ticker Name
For i = 2 To lastrow1
        If ws.Cells(i, 11).Value = TotalVol Then
        ws.Cells(4, 14).Value = ws.Cells(i, 8).Value
        End If
Next i

'Auto fit all columns
Worksheets(WorksheetName).Columns("A:Z").AutoFit
Next ws

End Sub



