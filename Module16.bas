Attribute VB_Name = "Module16"
Sub RefreshAllPivotTables()
'Updateby20140724
Dim xWs As Worksheet
Dim xTable As PivotTable
For Each xWs In Application.ActiveWorkbook.Worksheets
    For Each xTable In xWs.PivotTables
        xTable.RefreshTable
    Next
Next
End Sub
