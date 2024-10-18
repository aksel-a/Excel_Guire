Attribute VB_Name = "InsertRowsBasedOnColumnValue"
Sub InsertRowsBasedOnColumnValue()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim numRows As Long
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A (adjust if necessary)
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Set the range to the column where the values are (e.g., Column A)
    ' Change "A" to the desired column
    Set rng = ws.Range("I1:I" & lastRow)
    
    ' Loop through each cell in the range (go backwards to avoid conflict when inserting rows)
    For i = lastRow To 1 Step -1
        Set cell = rng.Cells(i)
        
        ' Get the number of rows to insert
        If IsNumeric(cell.Value) And cell.Value > 0 Then
            numRows = cell.Value
            
            ' Insert the rows below the current cell
            If numRows > 0 Then
                cell.Offset(1, 0).Resize(numRows).EntireRow.Insert Shift:=xlDown
            End If
        End If
    Next i
    
    ' Refresh last row reference after insertion
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
End Sub


