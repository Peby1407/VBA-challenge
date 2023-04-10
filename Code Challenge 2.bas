Attribute VB_Name = "Module1"
Sub Challenge2()
'Enabeling work for each worksheet
For Each ws In Worksheets
'Defining the quantity of each Ticker
Dim cantidadticker As Integer
cantidadticker = WorksheetFunction.CountIf(ws.Range("A1:A700"), "AAB") - 1

Dim i As Long
'create new column headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

Dim ticker As String
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Summary_Table_Row2 As Integer
Summary_Table_Row2 = 2
Dim Percent_Change As Double
Percent_Change = 0
'Define the lastrow
Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ' Set the ticker
        ticker = ws.Cells(i, 1).Value
    
         ' Add to the Volume Total
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
         ' Print the ticker
        ws.Range("I" & Summary_Table_Row).Value = ticker
    
         ' Print the Volume total
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
         ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
    
         ' Reset the Total Volume
        Total_Stock_Volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Volume total
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    End If
 
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Calculate the difference between the las close and first open value
        Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(i - cantidadticker, 3).Value
        'Enter the difference to the new column
        ws.Range("J" & Summary_Table_Row2).Value = Yearly_Change
    
        
    
        'Calculate the differente in %
        Percent_Change = Yearly_Change / ws.Cells(i - cantidadticker, 3).Value
        'Enter the percent calculated in the new column
        ws.Range("K" & Summary_Table_Row2).Value = Percent_Change

        Summary_Table_Row2 = Summary_Table_Row
        'reset the Yearly Change
        Yearly_Change = 0
    
    Else
    End If
    
    

    
    'Set the conditional formatting
Next i
For i = 2 To 3001
    If ws.Cells(i, 10) < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
    Else
    
        ws.Cells(i, 10).Interior.ColorIndex = 4
    
    End If
    If ws.Cells(i, 11) < 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 3
        
    Else
    
        ws.Cells(i, 11).Interior.ColorIndex = 4
    
    End If

Next i
'Give % format
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"
'Look for maximum
ws.Range("Q2") = Application.WorksheetFunction.Max(ws.Range("K2:K800000"))
'Look for Minimum
ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K2:K800000"))
'Look for biggest volume
ws.Range("Q4") = Application.WorksheetFunction.Max(ws.Range("L2:L800000"))
'Autofit Cell
ws.Columns("I:Q").AutoFit
'Look for the ticker related to the minimum and maximums
ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I2:K800000"), WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K800000"), 0), 1)
ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I2:K800000"), WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K800000"), 0), 1)
ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("I2:L800000"), WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L800000"), 0), 1)
                 

Next ws
End Sub

