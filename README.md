# VBA_Challenges
VBA stock data analysis
Sub MarketAnalysis():

' Declare your variables
Dim i, SummChart As Long
SummChart = 2
Dim Ticker As String
Dim LastRow As Long
Dim TotalStockVolume As Long
Dim PercentChange As Double
Dim YearlyChange As Double
Dim PreviousAmount As Long
PreviousAmount = 2
Dim ws As Worksheets

'Define my columns in SummChart
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change %"
Cells(1, 12).Value = "Total Stock Volume"

'Find the last Row

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow

'SummaryChart Results
'Conditionals/Storing BOY data here

If Cells(i, 1).Value <> Cells(i - 1, 1) Then
    Ticker = Cells(i, 1).Value
    YearlyChange = Cells(i, 3).Value
    PercentChange = Cells(i, 3).Value
    Cells(SummChart, 12).Value = Cells(i, 7).Value

ElseIf Cells(i, 1).Value <> Cells(i + 1, 1) Then
       YearlyChange = Cells(i, 6).Value - YearlyChange
       PercentChange = YearlyChange / PercentChange * 100
       
           
'Paste data SummChart
Cells(SummChart, 9).Value = Ticker
Cells(SummChart, 10).Value = Round(YearlyChange, 2)
Cells(SummChart, 11).Value = Round(PercentChange, 2)
Cells(SummChart, 12).Value = Cells(i, 7).Value + Cells(SummChart, 12).Value

SummChart = SummChart + 1

'Middle TSV data included
ElseIf Cells(i, 1).Value = Cells(i - 1, 1) And Cells(i, 1).Value = Cells(i + 1, 1) Then
    Cells(SummChart, 12).Value = Cells(i, 7).Value + Cells(SummChart, 12).Value
    
    
    
    

' Nested looping for Formatting

'If Cells(SummChart, 10).Value < 0 Then
    'Cells(SummChart, 10).Interior.ColorIndex = 3

'ElseIf Cells(SummChart, 10).Value >= 0 Then
    'Cells(SummChart, 10).Interior.ColorIndex = 4
# causingerror



End If


'Define

' Create a script that will loop through all the stocks for one year
' Searches for when the value of the next cell is different than that of the current cell
    
    'If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
    'Cells(i, 9).Value = Ticker
    
        
'End If


' Last Row

Next i

End Sub

