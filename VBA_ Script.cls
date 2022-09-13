VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarketData()
    'Set an intial Variables
    Dim Ticker_Symbol As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim TotalTickerVolume As Double
    TotalTickerVolume = 0
    Dim SummaryTableRow As Long
    SummaryTableRow = 2
    Dim StartingPrice As Double
    StartingPrice = Cells(2, 3).Value
    Dim EndingPrice As Double
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            EndingPrice = Cells(i, 6).Value
            Yearly_Change = EndingPrice - StartingPrice
            Ticker_Symbol = Cells(i, 1).Value
            Percent_Change = Yearly_Change / StartingPrice * 100
            TotalTickerVolume = TotalTickerVolume + Cells(i, 7).Value
            Range("I" & SummaryTableRow).Value = Ticker_Symbol
            Range("J" & SummaryTableRow).Value = Yearly_Change
            Range("K" & SummaryTableRow).Value = Percent_Change
            Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            Range("L" & SummaryTableRow).Value = TotalTickerVolume
            TotalTickerVolume = 0
            SummaryTableRow = SummaryTableRow + 1
            StartingPrice = Cells(i + 1, 3).Value
            
        
        Else:
         TotalTickerVolume = TotalTickerVolume + Cells(i, 7).Value
         End If
         
        If Range("J" & SummaryTableRow).Value >= 0 Then
                   Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        Else
                   Range("J" & SummaryTableRow).Interior.ColorIndex = 3
        
        End If
    Next i
End Sub
