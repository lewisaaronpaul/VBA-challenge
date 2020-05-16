Attribute VB_Name = "StockData_Analysis_Challenges"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name: Aaron Paul Lewis                                                 '
'Rice University: Data Analytics and Visualization Boot Camp            '
'Assignment #2: VBA Homework - The VBA of Wall Street plus Challenges   '
'Date: May 16, 2020                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub StockDataAnalysis_Challenges()

' Declare Variables
Dim WkShtCnt As Integer
Dim TotalData As Range
Dim SortingCol As Range
Dim RowCnt As Long
Dim ColCnt As Long
Dim CurrentTicker As String
Dim NextTicker As String
Dim TotStockVol As Double
Dim Tickersymb As String
Dim PrtRow As Integer
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChg As Double
Dim PercentChg As Double
Dim Counter As Long
Dim Start As Double
Dim MaxPercentChg As Double
Dim MinPercentChg As Double
Dim MaxTicker As String
Dim MinTicker As String
Dim StartStockVol As Double
Dim MaxStockVol As Double
Dim MaxTotVolTicker As String

'Initialize the Worksheet count.
WkShtCnt = 1

'Loop across each worksheet.
Do While WkShtCnt <= Worksheets.Count           '# of Worksheets

'Select the Sheet and then select cell A1.
    Worksheets(WkShtCnt).Select
    Range("A1").Select
    
'Ensure that the table is sorted relative to the first column (ticker)!
'This is important for looping later.
    Set TotalData = Selection.CurrentRegion  'Select the table of data
    RowCnt = TotalData.Rows.Count            '# of rows
    ColCnt = TotalData.Columns.Count         '# of columns
    
'This is the range of the sorting column,'ticker'.
    Set SortingCol = Worksheets(WkShtCnt).Range("A2:A" & RowCnt)

'Sorting:- ascending order
    TotalData.Sort key1:=SortingCol, order1:=xlAscending, Header:=xlYes
    
'Print header for the summary.
    PrtRow = 1
    Cells(PrtRow, ColCnt + 2).Value = "Ticker"
    Cells(PrtRow, ColCnt + 2).Font.Bold = True
    Columns(ColCnt + 2).AutoFit
    Cells(PrtRow, ColCnt + 3).Value = "Yearly Change"
    Cells(PrtRow, ColCnt + 3).Font.Bold = True
    Columns(ColCnt + 3).AutoFit
    Cells(PrtRow, ColCnt + 4).Value = "Percent Change"
    Cells(PrtRow, ColCnt + 4).Font.Bold = True
    Columns(ColCnt + 4).AutoFit
    Cells(PrtRow, ColCnt + 5).Value = "Total Stock Volume"
    Cells(PrtRow, ColCnt + 5).Font.Bold = True
    Columns(ColCnt + 5).AutoFit
    
    TotStockVol = 0
    Column = 1
    Counter = 0
    MaxPercentChg = 0
    MinPercentChg = 0
    MaxStockVol = 0
    
'Loop over each row in the table of data.
    For i = 2 To RowCnt
        CurrentTicker = TotalData.Cells(i, 1).Value
        NextTicker = TotalData.Cells(i + 1, 1).Value
    
        If NextTicker = CurrentTicker Then
            Counter = Counter + 1
            
            TotStockVol = TotStockVol + TotalData.Cells(i, 7).Value
        
        ElseIf NextTicker <> CurrentTicker Then
            PrtRow = PrtRow + 1
            Worksheets(WkShtCnt).Select
            
            Counter = Counter + 1
            OpenPrice = TotalData.Cells(i - Counter + 1, 3).Value
            ClosePrice = TotalData.Cells(i, 6).Value
            YearlyChg = ClosePrice - OpenPrice
                        
            Tickersymb = CurrentTicker
            TotStockVol = TotStockVol + TotalData.Cells(i, 7).Value
            
'Print summary!
            Cells(PrtRow, ColCnt + 2).Value = Tickersymb
            Cells(PrtRow, ColCnt + 3).Value = YearlyChg
                If YearlyChg > 0 Then
                    Cells(PrtRow, ColCnt + 3).Interior.Color = vbGreen
                ElseIf YearlyChg < 0 Then
                    Cells(PrtRow, ColCnt + 3).Interior.Color = vbRed
                End If
                
                If OpenPrice <> 0 Then
                    PercentChg = YearlyChg / OpenPrice
                    Cells(PrtRow, ColCnt + 4).Value = FormatPercent(PercentChg)
                Else
                Cells(PrtRow, ColCnt + 4).Value = "Open Price Zero"
                End If
            Cells(PrtRow, ColCnt + 5).Value = TotStockVol
            
'Find greatest percentage increase, the greatest percentage decrease and greatest total volume.
                Start = PercentChg
                                                
                If Start > MaxPercentChg Then
                    MaxPercentChg = Start
                    MaxTicker = TotalData.Cells(i, 1).Value
                
                ElseIf Start < MinPercentChg Then
                    MinPercentChg = Start
                    MinTicker = TotalData.Cells(i, 1).Value
                End If
                
                StartStockVol = TotStockVol
                
                If StartStockVol > MaxStockVol Then
                    MaxStockVol = StartStockVol
                    MaxTotVolTicker = TotalData.Cells(i, 1).Value
                End If
                    
'Reset TotStockVol before starting the next stock.
            Counter = 0
            TotStockVol = 0
            
        End If
    
    Next i
'Autofit the contents of the newly created columns.
    Columns(ColCnt + 2).AutoFit
    Columns(ColCnt + 3).AutoFit
    Columns(ColCnt + 4).AutoFit
    Columns(ColCnt + 5).AutoFit
    
'Print the greatest summary for the challenges.
    Cells(1, ColCnt + 9).Value = "Ticker"
    Cells(1, ColCnt + 9).Font.Bold = True
    Cells(1, ColCnt + 10).Value = "Value"
    Cells(1, ColCnt + 10).Font.Bold = True
    Cells(2, ColCnt + 8).Value = "Greatest % Increase"
    Cells(2, ColCnt + 8).Font.Bold = True
    Cells(2, ColCnt + 9).Value = MaxTicker
    Cells(2, ColCnt + 10).Value = FormatPercent(MaxPercentChg)
    Cells(3, ColCnt + 8).Value = "Greatest % Decrease"
    Cells(3, ColCnt + 8).Font.Bold = True
    Cells(3, ColCnt + 9).Value = MinTicker
    Cells(3, ColCnt + 10).Value = FormatPercent(MinPercentChg)
    Cells(4, ColCnt + 8).Value = "Greatest Total Volume"
    Cells(4, ColCnt + 8).Font.Bold = True
    Cells(4, ColCnt + 9).Value = MaxTotVolTicker
    Cells(4, ColCnt + 10).NumberFormat = "0.0000E+00"
    Cells(4, ColCnt + 10).Value = MaxStockVol
    
    Columns(ColCnt + 8).AutoFit
    Columns(ColCnt + 9).AutoFit
    Columns(ColCnt + 10).AutoFit
    
'Reset WkShtCnt for the Do While loop, which changes Worksheet!
    WkShtCnt = WkShtCnt + 1
    
Loop

End Sub

