VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)

Static xRow

If xRow <> "" Then
    With Rows(xRow).Interior
        .Color = xlNone
    End With

End If

Active_Row = Selection.Row
xRow = Active_Row

With Rows(Active_Row).Interior
    .Color = RGB(127, 223, 200) ' Replace with your desired RGB values
    .Pattern = xlSolid

End With

End Sub
