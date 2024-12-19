VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub QuarterlyStockAnalysis()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

    For Each ws In ThisWorkbook.Worksheets

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 11).Value = "Quarterly Change ($)"
        ws.Cells(1, 12).Value = "Percent Change (%)"

        Dim i As Long

        For i = 2 To lastRow

            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            totalVolume = ws.Cells(i, 7).Value

            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If

            ws.Cells(i, 9).Value = ticker
            ws.Cells(i, 10).Value = totalVolume
            ws.Cells(i, 11).Value = quarterlyChange
            ws.Cells(i, 12).Value = percentChange

            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If

            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If

            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If

        Next i

        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatestIncreaseTicker
        ws.Cells(2, 16).Value = greatestIncrease

        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = greatestDecreaseTicker
        ws.Cells(3, 16).Value = greatestDecrease

        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatestVolumeTicker
        ws.Cells(4, 16).Value = greatestVolume

        ' Apply conditional formatting for Quarterly Change
        With ws.Range("K2:K" & lastRow)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0)  ' Green for positive change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0)  ' Red for negative change
        End With

        ' Apply conditional formatting for Percent Change
        With ws.Range("L2:L" & lastRow)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0)  ' Green for positive change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0)  ' Red for negative change
        End With

    Next ws

End Sub

