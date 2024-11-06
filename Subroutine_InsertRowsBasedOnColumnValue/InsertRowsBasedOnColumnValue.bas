Attribute VB_Name = "InsertRowsBasedOnColumnValue"
' Author: Aksel Alvarez
' Github: https://github.com/aksel-a/Excel_Guire
' Date: 2024-10-18
' Description: This VBA macro inserts rows based on the numeric values in one column (you will need to adapt it to suit your needs) of the active worksheet.
' Version: v1.0

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
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Set the range to the column where the values are (e.g., Column I)
    ' Change "I" to the desired column
    Set rng = ws.Range("A1:A" & lastRow)
    
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
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
End Sub