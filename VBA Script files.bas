Attribute VB_Name = "Module1"
Sub Stock_Analysis()
'Declare variables

Dim ws As Worksheet
Dim LasRow As Long
Dim Ticker As String
Dim TotalVolume As Double
Dim PrintCount As Long
Dim YearOpen As Double
Dim YearClose As Double
Dim YearCount As Long
Dim RowCount As Long
Dim Maximum_Value As Double
Dim Most_Volume As Double
Dim Increase_index As Long
Dim Decrease_index As Long
Dim Volume_index As Long

'Loop through worksheets

For Each ws In ThisWorkbook.Worksheets

'Titles for I1:L1

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

'Variables for every sheet

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
PrintCount = 1
YearCount = 2

'Loop through the rows
For i = 2 To LastRow
    YearOpen = ws.Cells(YearCount, 3).Value
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
'Ticker Column
Ticker = ws.Cells(i, 1).Value

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        PrintCount = PrintCount + 1
        ws.Cells(PrintCount, 9).Value = Ticker
        YearClose = ws.Cells(i, 6).Value
        ws.Cells(PrintCount, 10).Value = (YearClose - YearOpen)
        
    'Add Column color
    
    If ws.Cells(PrintCount, 10).Value > 0 Then
    ws.Cells(PrintCount, 10).Interior.ColorIndex = 4
    
    ElseIf ws.Cells(PrintCount, 10).Value < 0 Then
    ws.Cells(PrintCount, 10).Interior.ColorIndex = 3
    
    Else
    '0 will be neglected
    ws.Cells(PrintCount, 10).Interior.ColorIndex = 2
    End If
    
YearCount = i + 1
    
    If YearOpen <> 0 Then
    ws.Cells(PrintCount, 11).Value = (((YearClose - YearOpen) / YearOpen))
    ws.Cells(PrintCount, 11).Value = FormatPercent(ws.Cells(PrintCount, 11), 2)
    
    Else
    End If
    ws.Cells(PrintCount, 12).Value = TotalVolume
    TotalVolume = 0
    
    Else
    
    End If
    
    Next i
    
    'Get Percentages
    ws.Range("O2").Value = "Greatest Increase Perentage"
    ws.Range("O3").Value = "Greatest Decrese Percentage"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    RowCount = ws.Cells(Rows.Count, "K").End(x1Up).Row
    
    'Find Maximum
    
    Maximum_Value = WorksheetFunction.Max(ws.Range("K2:K" & RowCount))
    Increase_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    
    'Find Minimum
    Minimum_Value = WorksheetFunction.Min(ws.Range("K2:K" & RowCount))
    Decrease_index = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    
    'Find Volume
    Most_Volume = WorksheetFunction.Max(ws.Range("L2:l" & RowCount))
    Volume_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:l" & RowCount)), ws.Range("L2:l" & RowCount), 0)
    
    'Print Results
    
    ws.Range("P2") = ws.Cells(Increase_index + 1, 9)
    ws.Range("Q2") = Maximum_Value
    ws.Range("Q2") = FormatPercent(ws.Range("Q2"), 2)
    
    ws.Range("P3") = ws.Cells(Decrease_index + 1, 9)
    ws.Range("Q3") = Minimum_Value
    ws.Range("Q3") = FormatPercent(ws.Range("Q3"), 2)
    
    ws.Range("P4") = ws.Cells(Volume_index + 1, 9)
    ws.Range("Q4") = Most_Volume
    ws , Range("Q4") = FormatPercent(ws.Range("Q4"), 2)
    Next ws
    
    
End Sub
