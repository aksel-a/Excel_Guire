Attribute VB_Name = "CopyRowsIfNotInWsB"
Function CopyRowsIfNotInWsB(wsAName As String, wsBName As String, wsCName As String, colA As String, colB As String, startRowA As Long, startRowB As Long)
    Dim wsA As Worksheet, wsB As Worksheet, wsC As Worksheet
    Dim lastRowA As Long, lastRowB As Long, lastRowC As Long
    Dim i As Long, found As Range
    
    ' Set worksheet references
    Set wsA = ThisWorkbook.Sheets(wsAName)
    Set wsB = ThisWorkbook.Sheets(wsBName)
    Set wsC = ThisWorkbook.Sheets(wsCName)
    
    ' Get the last row for WS_A and WS_B
    lastRowA = wsA.Cells(wsA.Rows.Count, colA).End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, colB).End(xlUp).Row
    lastRowC = 1 ' Assuming you want to start copying in Row 1 of WS_C
    
    ' Loop through WS_A rows starting from startRowA
    For i = startRowA To lastRowA
        Set found = wsB.Range(colB & startRowB & ":" & colB & lastRowB).Find(wsA.Cells(i, colA).Value, LookIn:=xlValues, LookAt:=xlWhole)
        If found Is Nothing Then
            wsA.Rows(i).Copy wsC.Rows(lastRowC)
            lastRowC = lastRowC + 1
        End If
    Next i
End Function

Sub ExecuteCopyRowsIfNotInWSB()
    ' Define worksheet names, column references, and starting rows
    Dim wsAName As String, wsBName As String, wsCName As String
    Dim colA As String, colB As String
    Dim startRowA As Long, startRowB As Long
    
    wsAName = "01-QS-Rooms-SOLL_IST_Werte"
    wsBName = "_FRP413"
    wsCName = "QS_NotInFRP"
    colA = "F"
    colB = "F"
    startRowA = 3 ' Data starts at row 3 in WS_A
    startRowB = 2 ' Data starts at row 2 in WS_B
    
    ' Call the CopyRowsIfNotInWSB function with dynamic parameters
    Call CopyRowsIfNotInWsB(wsAName, wsBName, wsCName, colA, colB, startRowA, startRowB)
End Sub
