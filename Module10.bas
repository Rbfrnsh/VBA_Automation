Attribute VB_Name = "Module10"
Sub ChangeValue()
Worksheets("Filtered").Activate
LastRow = ActiveSheet.UsedRange.Rows.Count
Range("X2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Selection.PasteSpecial xlPasteValues
'Range("Y2").AutoFill Destination:=Range("Y2:Y" & LastRow)

End Sub
