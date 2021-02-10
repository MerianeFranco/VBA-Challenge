Sub Summary_Stock() 'Sub  code available at from Extendoffice.com

Dim xsh As Worksheet
Application.ScreenUpdating = False
For Each xsh In Worksheets
    xsh.Select
    Call stock
    Call test
Next
Application.ScreenUpdating = True

End Sub


Sub stock()

'* Create a script that will loop through all the stocks for one year and output the following information.

'  * The ticker symbol.

'  * Yearly change from opening price athe beginning of a given year to the closing price at the end of that year.

'  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

'  * The total stock volume of the stock.

'* You should also have conditional formatting that will highlight positive change in green and negative change in red.

'* The result should look as follows.

'![moderate_solution](Images/moderate_solution.png)



' Need 4 buckets for calculation
'(Ticker symbol[literal string],
'Yearly change [double],    math: latest close minus first open
'PChange [double],         math:  Yearly change [double]/first open
'Volume [integer])          math:  sum of the ticker colume during an year

'Excercise example
'Find the last ticker
'Credit card excercise


    'Dim ws As Worksheet
    Dim lastrow As Double
    lastrow = 0
    'Dim i As Integer

'Build the titles


  
  
  'For Each ws In Worksheets
  
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Percent % Decrease"
Cells(4, 15).Value = "Greatest TOtal Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
  ' Find the last row of each worksheet

    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
   


  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String


  ' Set an initial variable for holding the yearly Change
  Dim YChange As Double
  

  ' Set an initial variable for holding the % Change
  Dim PChange As Double
   
  ' Set an initial variable for holding the total volume
  Dim YVolume_Total As Double
  YVolume_Total = 0

  ' Keep track of the location for each ticket in the summary table
  Dim Ticket_Table_Row As Integer
  Ticket_Table_Row = 2
    

  
  
  ' Loop through all tickets
  For i = 2 To lastrow '*************************************************(need to change the hard code)
 
 
    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Brand name
      Ticker_Name = Cells(i, 1).Value
      
      'Set The first value of stock at opening
 
    FirstOpen = Cells(i + 1, 3).Value
      
      ' Add to the Brand Total
     YVolume_Total = YVolume_Total + Cells(i, 7).Value

      ' Print the Ticket in the Summary Table
      Range("I" & Ticket_Table_Row).Value = Ticker_Name

      ' Print the Brand Amount to the Summary Table
      Range("L" & Ticket_Table_Row).Value = YVolume_Total

      ' Add one to the summary table row
      Ticket_Table_Row = Ticket_Table_Row + 1
      
      ' Reset the Brand Total
      YVolume_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else
    
    
    
      ' Add to the Volume Total
      YVolume_Total = YVolume_Total + Cells(i, 7).Value
      
     'Yearly change [double],    math: latest close minus first open
     
     Dim LatestClose As Double
    LatestClose = Cells(i + 1, 6).Value
     
     
     If Cells(i, 1).Value = Cells(i + 1, 1).Value Then

    'fix the firstopen value for the first ticker
     If FirstOpen = 0 Then
     FirstOpen = Cells(2, 3).Value
     End If
     
     
      YChange = LatestClose - FirstOpen
        ' Print YChange in column J
     Range("J" & Ticket_Table_Row).Value = YChange
 
     
      
      'PChange [double],         math:  Yearly change [double]/first open
           PChange = Round(YChange / FirstOpen, 2)
        ' Print YChange in column k
     Range("k" & Ticket_Table_Row).Value = PChange
      
      
      End If
    End If
    
 
Next i


'conditional formatting that will highlight positive change in green and negative change in red ***% Change [double]


  
'Cells(2, 3).Interior.ColorIndex = 4

End Sub

Sub test()


lastrow2 = Cells(Rows.Count, "I").End(xlUp).Row


 

'Greatest PIncrease
Dim GPIncrease As Double
GPIncrease = Application.WorksheetFunction.Max(Range("k:k"))
Cells(2, 17).Value = GPIncrease

For i = 1 To lastrow2
If Cells(i, 11).Value = GPIncrease Then
Cells(2, 16).Value = Cells(i, 9).Value
End If
Next i





'Greatest PDecrease
Dim GPDecrease As Double
GPDecrease = Application.WorksheetFunction.Min(Range("k:k"))
Cells(3, 17).Value = GPDecrease

For i = 1 To lastrow2
If Cells(i, 11).Value = GPDecrease Then
Cells(3, 16).Value = Cells(i, 9).Value
End If
Next i


'Greatest Total Volume
Dim GTotalVolume As Double
GTotalVolume = Application.WorksheetFunction.Max(Range("L:L"))
Cells(4, 17).Value = GTotalVolume


For i = 1 To lastrow2
If Cells(i, 12).Value = GTotalVolume Then
Cells(4, 16).Value = Cells(i, 9).Value
End If
Next i



'formating colors
For i = 2 To lastrow2
   If (Cells(i, 10).Value >= 0) Then
    Cells(i, 10).Interior.ColorIndex = 4
     'MsgBox "fica verde"
     Else
     Cells(i, 10).Interior.ColorIndex = 3
      'MsgBox "fica vermelho"
End If

Next i
End Sub

