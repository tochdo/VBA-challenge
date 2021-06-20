Attribute VB_Name = "Module1"
Sub VBA_challenge()

  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total volume per ticker name
  Dim Total_Volume As Double
  Total_Volume = 0
  
  'ws
  Dim ws As Worksheet
  
  'define lastrow
  Dim lastrow As Double
  

  'Loop through all worksheet
  For Each ws In Worksheets
  
           'lastrow formula
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    ' Create Summary Table Headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
  
    ' Keep track of the location for each Ticker Name in the summary table
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
  
    'Format cell size
        ws.Columns("I:L").EntireColumn.AutoFit
  
        ' Loop through all ticker names
         For i = 2 To lastrow

        ' Check if we are still within the same Ticker Name, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
       ' Set the Ticker Name
        Ticker_Name = ws.Cells(i, 1).Value
        
        ' Add to the Volume Total
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
         
        ' Print the Ticker Name in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

        ' Print the Total Volume Amount to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = Volume_Total

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset the Volume Total
        Volume_Total = 0
        
       ' If the cell immediately following a row is the same Ticker Name..
                Else
                
        'Add to the Volume Total
        Volume_Total = Volume_Total + ws.Cells(i, 3).Value
     
        End If

        Next i

    Next ws

End Sub

'Set initial variable for first iteration
'Dim firsties as double

'Set initial variable for first opening price
'Dim Open_At As Double

'Set initial variable for closing price
'Dim Close_At as double

'set initial variable for Percent Change
'Dim PC as double

'Set initial varaible for Yearly Change
'Dim YC

'Track the firstime it goes through the loop
'firsties = firsties + 1

'Test if it is the first time through
'If firsties = 1 Then

'Snatch opening price to save for later
'Open_At = ws.cells(i,3)

'Else

'Snatch the closing price to save for later
'Close_At =ws.cells(i,6)

'Formula to get % and yearly change with stored variables
'If Open_At <> 0 Then
    'PC = (Close_At - Open_At) / Open_At)
    'YC = (Open_At - Close_At)
