Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()

Dim LastRow As Long
Dim TotalVolume As LongLong
Dim SummaryRowNumber As Integer
Dim StockName As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim PriceDiff As Double

Dim LRange As Range
Dim MinValue As Double
Dim MaxValue As Double
Dim MaxVolume As LongLong
Dim MaxVolumeRow As Long

Dim MinMaxCell As Range
Dim MinMaxRowNumber As Integer
Dim StrVal As String

Dim i As Long
Dim BeginRow As Long

    For Each ws In Worksheets
            MaxVolume = 0
            MaxVolumeRow = 0
            TotalVolume = 0
            BeginRow = 2
            'Keep track of the row summary info goes on
            SummaryRowNumber = 2
            LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
            
            'MsgBox (ws.Name)
            'Set the header
            ws.Cells(1, 10).Value = "Ticker Symbol"
            ws.Cells(1, 11).Value = "Yearly Change"
            ws.Cells(1, 12).Value = "% Change"
            ws.Cells(1, 13).Value = "Total Stock Volume"
            ws.Cells(1, 14).Value = "Open Price (BOY)"
            ws.Cells(1, 15).Value = "Close Price (EOY)"
            ws.Cells(1, 18).Value = "Ticker"
            ws.Cells(1, 19).Value = "Value"
            
            'iterate through every stock symbol
            
            For i = 2 To LastRow
                     If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                     'Dont' forget to add the
                            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                            If TotalVolume > MaxVolume Then
                                MaxVolume = TotalVolume
                                MaxVolumeRow = SummaryRowNumber
                            End If
                            
                            ws.Cells(SummaryRowNumber, 10).Value = ws.Cells(i, 1)
                            
                            'Keep track of Opening price
                            OpenPrice = ws.Cells(BeginRow, 3).Value
                            
                            'Close Price for the end of the year
                            ClosePrice = ws.Cells(i, 6).Value
                            
                            ws.Cells(SummaryRowNumber, 11).Value = ClosePrice - OpenPrice
                            
                            If OpenPrice > 0# Then
                                PriceDiff = ((ClosePrice - OpenPrice) / OpenPrice)
                                ws.Cells(SummaryRowNumber, 12).Value = Round(PriceDiff, 4)
                            Else
                                PriceDiff = 0#
                                ws.Cells(SummaryRowNumber, 12).Value = 0
                            End If
                            
                            ws.Cells(SummaryRowNumber, 13).Value = TotalVolume
                            ws.Cells(SummaryRowNumber, 14).Value = OpenPrice
                            ws.Cells(SummaryRowNumber, 15).Value = ClosePrice
                            
                            If PriceDiff > 0# Then
                                ws.Cells(SummaryRowNumber, 12).Interior.ColorIndex = 4
                            ElseIf PriceDiff < 0# Then
                                ws.Cells(SummaryRowNumber, 12).Interior.ColorIndex = 3
                            Else
                                ws.Cells(SummaryRowNumber, 12).Interior.ColorIndex = 8
                            End If
                            
                            SummaryRowNumber = SummaryRowNumber + 1
                            
                            'Update the begin row to point to the beginning of next ticker
                            BeginRow = i + 1
                            'Reset Total volume for the next ticker
                            TotalVolume = 0
                     Else
                            'Add the volume for the year
                            TotalVolume = TotalVolume + Cells(i, 7).Value
                            If TotalVolume > MaxVolume Then
                                MaxVolume = TotalVolume
                                MaxVolumeRow = SummaryRowNumber
                            End If
                    End If
            
            Next i
            'Find the min Max values
            
                    LastRow = ws.Cells(Rows.Count, "L").End(xlUp).Row
                    Set LRange = ws.Range("L1:L" & LastRow)
                    MinValue = WorksheetFunction.Min(LRange)
                    ws.Cells(2, 17).Value = " Greatest % Decrease"
                    Set MinMaxCell = LRange.Find(what:=MinValue, LookIn:=xlValues)
                    MinMaxRowNumber = MinMaxCell.Row
                    ws.Cells(2, 18).Value = ws.Cells(MinMaxRowNumber, 10)
                    ws.Cells(2, 19).Value = MinValue
                    ws.Cells(2, 19).NumberFormat = "0.00%"
                    
                    ws.Cells(4, 17).Value = " Greatest % Increase"
                    MaxValue = WorksheetFunction.Max(LRange)
                    Set MinMaxCell = LRange.Find(what:=MaxValue, LookIn:=xlValues)
                    MinMaxRowNumber = MinMaxCell.Row
                    ws.Cells(4, 18).Value = ws.Cells(MinMaxRowNumber, 10)
                    ws.Cells(4, 19).Value = MaxValue
                    ws.Cells(4, 19).NumberFormat = "0.00%"
                    
                    ws.Cells(6, 17).Value = " Greatest % Volume"
                    ws.Cells(6, 18).Value = ws.Cells(MaxVolumeRow, 10)
                    ws.Cells(6, 19).Value = MaxVolume
                    
                    'Does not workwhile comparing largenumbers that have scientific notation
                    'Set LRange = ws.Range("M1:M" & LastRow)
                    'MaxVolume = WorksheetFunction.Max(LRange)
                    'Set MinMaxCell = LRange.Find(what:=MaxVolume, LookIn:=xlValues)
                    'MinMaxRowNumber = MinMaxCell.Row
                    
                    'MsgBox (ws.Cells(MinMaxRowNumber, 1))
                    
                    ws.Range("L1:L" & LastRow).NumberFormat = "0.00%"
    Next ws
End Sub
Sub Main()

For Each ws In Worksheets
    'MsgBox (ws.Name)
Next ws

End Sub
