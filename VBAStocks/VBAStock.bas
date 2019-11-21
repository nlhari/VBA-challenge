Attribute VB_Name = "Module1"
Sub VBAStock()
  ' Set an initial variable for holding the ticker name
  Dim Ticker As String
  Dim lastrow As Long
  Dim tickerStartRow As Long, tickerEndRow As Long
  Dim yearOpeningValue As Double
  Dim yearClosingValue As Double
  Dim yearlyChange As Double
  Dim percentChange As Double
  Dim Summary_Table_Row As Long
  Dim Ticker_Total As Double

  Dim WS_Count As Integer

  ' Set WS_Count equal to the number of worksheets in the active
  ' workbook.
  WS_Count = ActiveWorkbook.Worksheets.Count

  ' Loop to navigate one sheet at a time. Begin with first sheet.
  For J = 1 To WS_Count
    Sheets(J).Activate
    ActiveSheet.Select
  
    'reset number of rows variable for each sheet.
    lastrow = 0
    
    'count the number of rows in active sheet.
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Set an initial variable for holding the total volume per ticker
    Ticker_Total = 0

    ' Keep track of the location for each ticker in the summary table
    Summary_Table_Row = 2

    ' Loop through all ticker data to summarize volume
    ' Summarizing volume gives us unique Ticker list in column I
    For I = 2 To lastrow

      ' Check if we are still within the same ticker brand, if it is not...
      If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        ' Set the ticker name
        Ticker = Cells(I, 1).Value

        ' Add to the ticker Total
        Ticker_Total = Ticker_Total + Cells(I, 7).Value

        ' Print the ticker Brand in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker

        ' Print the ticker total volume to the Summary Table
        Range("L" & Summary_Table_Row).Value = Ticker_Total

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the ticker Total
        Ticker_Total = 0

        ' If the cell immediately following a row is the same ticker...
     Else

        ' Add to the ticker Total
        Ticker_Total = Ticker_Total + Cells(I, 7).Value

     End If

    Next I
  
    'Set header for 1st output
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
  
    'Adjust the width of the output columns
    Range("J:L").Columns.AutoFit
  
    ' Now that we have the unique Ticker list in the sheet,
    ' run another loop where for each unique ticker,
    ' find the first row and the last row.
    ' First Row will have the Year Opening value of the ticker and
    ' Last Row will have the Year Closing value.
    ' Once this is read, the yearly change and percentage change is simple math.
    
    For I = 2 To (Summary_Table_Row - 1)
    
      'Initialize all values to 0 for each ticker
      tickerStartRow = 0
      tickerEndRow = 0
      yearOpeningValue = 0
      yearClosingValue = 0
      yearlyChange = 0
      percentChange = 0
    
      'Find first and last rows of Ticker - read from Open value of first row and Close value of last row
      tickerStartRow = Range("A:A").Find(What:=Cells(I, 9).Value, After:=Range("A1")).Row
      tickerEndRow = Range("A:A").Find(What:=Cells(I, 9).Value, After:=Range("A1"), lookat:=xlWhole, searchdirection:=xlPrevious).Row
    
      'Set Opening and Closing Values
      yearOpeningValue = Cells(tickerStartRow, 3).Value
      yearClosingValue = Cells(tickerEndRow, 6).Value
    
      'Difference between Close value from last row and Open value from first row will give Yearly Change
      yearlyChange = yearClosingValue - yearOpeningValue
    
      'Write the results to columns J and K
      Cells(I, 10).Value = yearlyChange
      If yearOpeningValue <> 0 Then
        ' <>0 is to ensure that if opening value is 0, then it doesn't become a divident throwing an error.
        ' The challenge assignment requires to search the Tickers for Min and Max percentage change from the unique list
        ' turns out the search for Min and Max rows requires exact match of value.
        ' the Round function trims the digits to 4 decimals so the search is match.
        
        Cells(I, 11).Value = Round(yearlyChange / yearOpeningValue, 4)
        
      Else
      
        Cells(I, 11).Value = 0
      
      End If
    
      'Conditional formatting of Yearly Change - Green for positive and Red for negative
      If yearlyChange > 0 Then
       Cells(I, 10).Interior.ColorIndex = 4
      Else
       Cells(I, 10).Interior.ColorIndex = 3
      End If
      'MsgBox ("Start Row " & tickerStartRow & ", End Row " & tickerEndRow)
    
    Next I
  
    'The challenge assignment
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim IncreaseTicker As Range
    Dim DecreaseTicker As Range
    Dim VolumeTicker As Range
    
    ' get the Min and Max values
    greatestIncrease = WorksheetFunction.Max(Range("K2", "K" & (Summary_Table_Row - 1)))
    greatestDecrease = WorksheetFunction.Min(Range("K2", "K" & (Summary_Table_Row - 1)))
    greatestVolume = WorksheetFunction.Max(Range("L2", "L" & (Summary_Table_Row - 1)))
    
    ' find the rows with Min and Max values
    
    Set IncreaseTicker = Range("K2", "K" & (Summary_Table_Row - 1)).Find(What:=greatestIncrease)
    Set DecreaseTicker = Range("K2", "K" & (Summary_Table_Row - 1)).Find(What:=greatestDecrease)
    Set VolumeTicker = Range("L2:L290").Find(What:=greatestVolume)
         
    ' get the ticker name for the Min and Max changes
    
    Range("P2").Value = IncreaseTicker.Offset(, -2).Value
    Range("Q2").Value = greatestIncrease
    Range("P3").Value = DecreaseTicker.Offset(, -2).Value
    Range("Q3").Value = greatestDecrease
    
    'Set header for challenge output
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
   
    
    Range("P4").Value = VolumeTicker.Offset(, -3).Value
    Range("Q4").Value = greatestVolume
    
    'Format Yearly Change to $ and Percent Change to % with two decimals
    Range("J:J").NumberFormat = "$#,##0.00"
    Range("K:K").NumberFormat = "0.00%"
    Range("O:O").Columns.AutoFit
    Range("Q2:Q3").NumberFormat = "0.00%"
    

  Next J

End Sub

Sub WorksheetLoop()

    Dim WS_Count As Integer
    Dim I As Integer

    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop.
    For I = 1 To WS_Count
       Sheets(I).Activate
       ' Insert your code here.
       ' The following line shows how to reference a sheet within
       ' the loop by displaying the worksheet name in a dialog box.
       MsgBox ActiveWorkbook.Worksheets(I).Name

    Next I

End Sub

Sub Challenge()

    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim IncreaseTicker As Range
    Dim DecreaseTicker As Range
    Dim VolumeTicker As Range
    
    Summary_Table_Row = 172
    Range("K:K").NumberFormat = "General"
    
    greatestIncrease = WorksheetFunction.Max(Range("K2", "K" & (Summary_Table_Row - 1)))
    greatestDecrease = WorksheetFunction.Min(Range("K2", "K" & (Summary_Table_Row - 1)))
    greatestVolume = WorksheetFunction.Max(Range("L2", "L" & (Summary_Table_Row - 1)))
    
    MsgBox ("Increase " & greatestIncrease & ", Decrease " & greatestDecrease)
    
    
    Set IncreaseTicker = Range("K2", "K" & (Summary_Table_Row - 1)).Find(What:=greatestIncrease)
    Set DecreaseTicker = Range("K2", "K" & (Summary_Table_Row - 1)).Find(What:=greatestDecrease)
    Set VolumeTicker = Range("L2:L290").Find(What:=greatestVolume)
         
    'MsgBox (Range("K2:K290").Find(What:=greatestDecrease))
         
    'Set header for 1st output
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
   
    'Adjust the width of the output columns
    Range("O:O").Columns.AutoFit
    
    
    Range("P2").Value = IncreaseTicker.Offset(, -2).Value
    Range("Q2").Value = greatestIncrease
    Range("P3").Value = DecreaseTicker.Offset(, -2).Value
    Range("Q3").Value = greatestDecrease
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    
    Range("P4").Value = VolumeTicker.Offset(, -3).Value
    Range("Q4").Value = greatestVolume
    Range("K:K").NumberFormat = "0.00%"


End Sub
