Attribute VB_Name = "Module2"
Sub Stock_Volume_medium()

'Create store variables
Dim stockvolume As Double
Dim laststoredrow As Integer
Dim annualcounter As Long
Dim percentchange As Variant
Dim annualchange As Double
Dim initialvalue As Double

' Find last row of data and store into lastrow variable
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through worksheets

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
'Set counter values to initial position
    stockvolume = 0
    laststoredrow = 2
    annualcounter = 0
    percentchange = 0
    initialvalue = 0

'Set titles for columns

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Annual Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'Create loop from first row to last row

    For i = 2 To lastrow
    'Check to see if this row is a different ticker than next row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'If it is different print stockinfo into new cells
            initialvalue = Cells(i - annualcounter, 3).Value
            annualchange = Cells(i, 6).Value - initialvalue
            If initialvalue <> 0 Then
                percentchange = annualchange / initialvalue * 100
                Else: percentchange = "Null"
                End If
            Cells(laststoredrow, 9).Value = Cells(i, 1).Value
            Cells(laststoredrow, 10).Value = annualchange
            Cells(laststoredrow, 11).Value = percentchange
            Cells(laststoredrow, 12).Value = stockvolume
            'Change percentage background color
            If percentchange > 0 Then
                Cells(laststoredrow, 11).Interior.ColorIndex = 4
                Else: Cells(laststoredrow, 11).Interior.ColorIndex = 3
                End If
            'Advance the aggregate row counter variable
            laststoredrow = laststoredrow + 1
            'Reset counter variables
            stockvolume = 0
            annualcounter = 0
            initialvalue = 0
            annualchange = 0
        'If it is not different then add that row's volume to stockvolume variable and annual counter
            Else:
                stockvolume = stockvolume + Cells(i, 7).Value
                annualcounter = annualcounter + 1
        End If
    Next i
Next

starting_ws.Activate

End Sub


