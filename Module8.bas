Attribute VB_Name = "Module8"
Sub Yearend()
Worksheets("Filtered").Activate
'LastRow = ActiveSheet.UsedRange.Rows.Count
Range("Y2").Select
'Range("Y2") = Application.WorksheetFunction.Text(Range("W2"), "yyyy")
Range("Y2").Formula = "=Year(W2)"
'Range("Y2").AutoFill Destination:=Range("Y2:Y" & LastRow)
Selection.AutoFill Destination:=Range("Y2:Y" & Range("W" & Rows.Count).End(xlUp).Row)
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlDown)).NumberFormat = "General"
End Sub



