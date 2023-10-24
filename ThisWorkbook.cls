VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBA_challenge()
    Dim WS As Worksheet
    Dim LastRow As Long
    Dim TickerSymbol As String
    Dim StartPrice As Double
    Dim EndingPrice As Double
    Dim Change As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestTotalVolumeTicker As String
    
    ' Loop through all worksheets in the workbook
    For Each WS In ThisWorkbook.Worksheets
        ' Initialize variables for each worksheet
        TotalVolume = 0
        SummaryRow = 2 ' Start the summary rows from row 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0
        GreatestIncreaseTicker = ""
        GreatestDecreaseTicker = ""
        GreatestTotalVolumeTicker = ""
        
        ' Activate the current worksheet
        WS.Activate
        
        ' Add headers for the required columns if not already present
        If WS.Range("I1").Value <> "Ticker" Then
            WS.Range("I1").Value = "Ticker"
            WS.Range("J1").Value = "Yearly Change"
            WS.Range("K1").Value = "Percent Change"
            WS.Range("L1").Value = "Total Stock Volume"
        End If
        
        ' Find the last row in column A
        LastRow = WS.Cells(WS.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through the data and extract ticker symbols starting from cell A2
        For i = 2 To LastRow
            ' Copy the ticker symbol from column A to column I
            WS.Cells(i, 9).Value = WS.Cells(i, 1).Value
            
            ' Calculate the yearly change (opening prices - closing prices)
            StartPrice = WS.Cells(i, 3).Value
            EndingPrice = WS.Cells(i, 6).Value
            Change = EndingPrice - StartPrice
            
            ' Calculate and display yearly change in column J
            WS.Cells(i, 10).Value = Change
            
            ' Calculate and display percentage change in column K
            If StartPrice <> 0 Then
                WS.Cells(i, 11).Value = (Change / StartPrice)
                ' Format as percentage
                WS.Cells(i, 11).NumberFormat = "0.00%"
            Else
                WS.Cells(i, 11).Value = 0
            End If
            
            ' Apply conditional formatting for positive and negative changes
            If Change > 0 Then
                ' Green color for positive change
                WS.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf Change < 0 Then
                ' Red color for negative change
                WS.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
            ' Calculate and display total stock volume in column L
            TotalVolume = TotalVolume + WS.Cells(i, 7).Value ' Assuming volume is in column G
            
            ' Check if the next row has a different ticker symbol or if it's the last row
            If WS.Cells(i + 1, 1).Value <> TickerSymbol Or i = LastRow Then
                ' Write the total stock volume to the summary column
                WS.Cells(SummaryRow, 12).Value = TotalVolume
                
                ' Reset total volume for the next ticker symbol
                TotalVolume = 0
                
                ' Move to the next summary row
                SummaryRow = SummaryRow + 1
            End If
            
            ' Check for the greatest % increase, % decrease, and total volume
            If WS.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = WS.Cells(i, 11).Value
                GreatestIncreaseTicker = WS.Cells(i, 9).Value
            End If
            
            If WS.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = WS.Cells(i, 11).Value
                GreatestDecreaseTicker = WS.Cells(i, 9).Value
            End If
            
            If TotalVolume > GreatestTotalVolume Then
                GreatestTotalVolume = TotalVolume
                GreatestTotalVolumeTicker = WS.Cells(i, 9).Value
            End If
        Next i
        
' Output the greatest % increase, % decrease, and total volume for this worksheet
        WS.Cells(SummaryRow, 15).Value = "Greatest % Increase"
        WS.Cells(SummaryRow + 1, 15).Value = "Greatest % Decrease"
        WS.Cells(SummaryRow + 2, 15).Value = "Greatest Total Volume"
        WS.Cells(SummaryRow, 16).Value = "Ticker"
        WS.Cells(SummaryRow + 1, 16).Value = "Value"
        WS.Cells(SummaryRow, 17).Value = GreatestIncreaseTicker
        WS.Cells(SummaryRow + 1, 17).Value = GreatestDecreaseTicker
        WS.Cells(SummaryRow + 2, 17).Value = GreatestTotalVolumeTicker

        ' Format the output
        WS.Range("O2:Q4").HorizontalAlignment = xlCenter
        WS.Range("O2:Q4").Font.Bold = True

    Next WS
End Sub
