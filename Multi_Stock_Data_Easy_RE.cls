VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks():
'Step 1 - Define the workbook and worksheets
Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

'Step 2 - Determine the last row
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Step 3 - Create the summary table
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Total Stock Volume"

'Step 4 - Define your variables
    Dim Ticker As String
    Dim Volume As Double
        Volume = 0
    Dim Row As Double
        Row = 2
    Dim Column As Integer
        Column = 1
    Dim i As Long

'Step 5 - Start loop
    For i = 2 To LastRow
        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            'String Ticker into the loop
            Ticker = Cells(i, Column).Value
            Cells(Row, Column + 8).Value = Ticker

            'Make sure the Volume resets or it will keep changing the values on you (i.e. Overflow Error)
            Volume = Volume + Cells(i, Column + 6).Value
            Cells(Row, Column + 9).Value = Volume
            Volume = 0

            'Make sure the table starts AFTER the header row (this can go anywhere, when i figured it out I just put it here, WITHIN THE LOOP)
            Row = Row + 1
        Else
            Volume = Volume + Cells(i, Column + 6).Value
        End If
    Next i
Next WS
End Sub

