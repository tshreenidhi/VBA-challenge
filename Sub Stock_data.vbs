Sub Stock_data()
 
For Each ws In Worksheets

 ' Set an initial variable for holding the Ticker Data
  Dim Stock_Name As String
  Dim Initial_Value As Double
  Dim Final_Value As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As String
  Dim TotalVolume As Double
       
  'Populate the result table headers
  
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage change"
    ws.Range("L1").Value = "Total Stock Volume"

  ' Keep track of the location for each stock value in the summary table
   Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2
   'Variable to keep track of the initial Value
   Init_index = 2
   
   ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
      
      TotalVolume = TotalVolume + ws.Cells(i, 7).Value
          

    ' Check if we are still within the same Stock, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
     ' Read the stock name, initial and final values of the stock
      Stock_Name = ws.Cells(i, 1).Value
      Initial_Value = ws.Cells(Init_index, 3).Value
      Final_Value = ws.Cells(i, 6).Value
      Init_index = i + 1

      ' Calculate the yearly change
      Yearly_Change = Final_Value - Initial_Value
      
      'Calculate the percentage change
      If Initial_Value <> 0 Then
           Percent_Change = FormatPercent((Yearly_Change / Initial_Value), 2, vbUseDefault, vbUseDefault, vbUseDefault)
          
        Else
           Percent_Change = 0
      End If
          

      ' Print the Stock name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Stock_Name

      ' Print the yearly change to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
     ' Print the Percentage Change to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
    ' Print the Total Stock Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = TotalVolume

      
      
      'Format the yearly change cells based on the value
      If Yearly_Change > 0 Then
         ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(0, 255, 0)
       ElseIf Yearly_Change < 0 Then
          ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(255, 0, 0)
       Else
          ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(0, 0, 255)
       End If
       
    ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Values
      Final_Value = 0
      TotalVolume = 0
    
    End If

  Next i

 'Find the max, min decrease and greatest Volume
 Max_Increase = WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row))
 Max_Decrease = WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row))
 Max_Volume = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row))
 Max_Increase_percent = FormatPercent(Max_Increase, 2, vbUseDefault, vbUseDefault, vbUseDefault)
 Max_Decrease_percent = FormatPercent(Max_Decrease, 2, vbUseDefault, vbUseDefault, vbUseDefault)
 


 
 'Write the max, min decrease and greatest Volume values to the result table
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("R2").Value = Max_Increase_percent
    ws.Range("R3").Value = Max_Decrease_percent
    ws.Range("R4").Value = Max_Volume
    
    'Max_IncreaseTicker_index = Application.Match(Max_Increase, Range("K2:K" & Summary_Table_Row), 0)
    'Max_IncreaseTicker = ws.Range("I" & (Max_IncreaseTicker_index + 1)).Value
    
    'Max_DecreaseTicker_index = Application.Match(Max_Decrease, Range("K2:K" & Summary_Table_Row), 0)
   ' Max_DecreaseTicker = ws.Range("I" & (Max_DecreaseTicker_index + 1)).Value
    
   ' Max_VolumeTicker_index = Application.Match(Max_Volume, Range("L2:L" & Summary_Table_Row), 0)
   ' Max_VolumeTicker = ws.Range("I" & (Max_VolumeTicker_index + 1)).Value
    
    Max_IncreaseTicker = ws.Cells(Application.WorksheetFunction.Match(Max_Increase, ws.Range("K2:K" & Summary_Table_Row), 0) + 1, 9).Value
    Max_DecreaseTicker = ws.Cells(Application.WorksheetFunction.Match(Max_Decrease, ws.Range("K2:K" & Summary_Table_Row), 0) + 1, 9).Value
    Max_VolumeTicker = ws.Cells(Application.WorksheetFunction.Match(Max_Volume, ws.Range("L2:L" & Summary_Table_Row), 0) + 1, 9).Value

    ws.Range("Q2").Value = Max_IncreaseTicker
    ws.Range("Q3").Value = Max_DecreaseTicker
    ws.Range("Q4").Value = Max_VolumeTicker
    
 
 Next ws
 
End Sub



