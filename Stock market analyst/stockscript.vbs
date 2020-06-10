Sub test12test()
'DEFINITIONS'
Dim ticker As String
Dim sum_table_row As Integer
    sum_table_row = 2
Dim openvalue As Double
Dim closevalue As Double
Dim total As Double
    total = 0
Dim occhange As Double
Dim ocpercent As Double
Dim last As Long
last = Cells(Rows.Count, 1).End(xlUp).Row
    
'START'
For i = 2 To last

    If openvalue = 0 Then
    openvalue = Cells(i, 3).Value
    End If

    'CONDITION'
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'LOCATIONS'
    ticker = Cells(i, 1).Value
    closevalue = Cells(i, 6).Value
    total = Cells(i, 7).Value

    'MATH'
    occhange = closevalue - openvalue
    ocpercent = (closevalue - openvalue) / openvalue
    
    On Error Resume Next
    
    
    'PRINT'
    Range("I" & sum_table_row).Value = ticker
    Range("J" & sum_table_row).Value = occhange
    Range("K" & sum_table_row).Value = ocpercent
   
    '+1 SUMMERY ROW'
    sum_table_row = sum_table_row + 1
    
    'RESET'
    total = 0
    openvalue = 0

Else
    
    total = total + Cells(i, 7).Value
    
    Range("L" & sum_table_row).Value = total
    
End If

'DONEZO'
Next i


End Sub
