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
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'Step 4 - Define your variables
    Dim Ticker As String
    Dim Volume As Double
        Volume = 0
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Row As Double
        Row = 2
    Dim i As Long

'Step 5 - Define Year_Open value (prevents Overflow Error)
    Year_Open = Cells(2, 3).Value

'Step 6 - Start loop
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'String Ticker into the loop
            Ticker = Cells(i, 1).Value
            Cells(Row, 9).Value = Ticker

            'Reset Year_Open
            Year_Open = Cells(2, 3).Value
            Year_Opem = Cells(i + 1, 3)

            'Set Year_Close Value
            Year_Close = Cells(i, 6).Value
            
            'Set Yearly_Change values (using defined Year_Open and Year_Close values)
            Yearly_Change = Year_Close - Year_Open
            Cells(Row, 10).Value = Yearly_Change
            
            'Calculate Percent_Change (use conditionals or the entire columns sets to 0.00%)
                If (Year_Open = 0 And Year_Close = 0) Then
                    Percent_Change = 0
                ElseIf (Year_Open = 0 And Year_Close <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Year_Open
                    Cells(Row, 11).Value = Percent_Change
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
            
            'Make sure the Volume resets or it will keep changing the values on you (i.e. Overflow Error)
            Volume = Volume + Cells(i, 7).Value
            Cells(Row, 12).Value = Volume
            Volume = 0

            'Make sure the table starts AFTER the header row (this can go anywhere, when i figured it out I just put it WITHIN THE LOOP)
            Row = Row + 1
        Else
            Volume = Volume + Cells(i, 7).Value
        End If
    Next i

'Step 7 - Create new lastrow function (for column 9) to loop conditional formatting
    YCLastRow = WS.Cells(Rows.Count, 10).End(xlUp).Row

'Step 8 - use If/ Else function for conditional formatting
    For j = 2 To YCLastRow
        If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
            Cells(j, 10).Interior.ColorIndex = 4
        ElseIf Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j

Next WS
End Sub

