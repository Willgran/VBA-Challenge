Attribute VB_Name = "Module1"
Sub Stonks()
'Variables

Dim ws As Worksheet
Dim ticker As String
Dim Volume As Long
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim totalStockVolume As Long
Dim Summary_Table_Row As Integer
Volume = 0

'Print for each worksheet
For Each ws In ThisWorkbook.Worksheets
'Titles printed
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
Next ws

'Overrides error
On Error Resume Next


'Run For Loop
Summary_Table_Row = 2

For i = 2 To 797712
    If year_open = 0 Then
        year_open = Cells(i, 3).Value
    End If

If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    year_close = Cells(i, 6).Value
    yearly_change = year_close - year_open

    ticker = Cells(i, 1).Value
    Volume = Volume + Cells(i, 7)
    Range("I" & Summary_Table_Row).Value = ticker
    Range("J" & Summary_Table_Row).Value = yearly_change
    Range("K" & Summary_Table_Row).Value = percent_change
    Range("L" & Summary_Table_Row).Value = Volume
    Summary_Table_Row = Summary_Table_Row + 1
    Volume = 0
Else: Volume = Volume + Cells(i, 7).Value


    End If
Next i

'format columns colors
Dim color_range As Range
Dim green_Red As Long
Dim CCount As Long
Dim color_cell As Range

Set color_range = ws.Range("J2", Range("J2").End(xlDown))

CCount = color_range.Cells.Count
For green_Red = 1 To Count
    Set color_cell = color_range(green_Red)

    Select Case color_cell
    Case Is >= 0
    With color_cell
    .Interior.Color = vbGreen
    End With

    Case Is < 0
    With color_cell
    .Interior.Color = vbRed
    End With

    End Select
Next green_Red



End Sub

