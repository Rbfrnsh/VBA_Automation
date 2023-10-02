Attribute VB_Name = "Module5"
Sub MFBAPS()
Worksheets("Filtered").Activate
LastRow = ActiveSheet.UsedRange.Rows.Count
Range("O2").Formula = "=M2+N2"
Range("O2").AutoFill Destination:=Range("O2:O" & LastRow)
End Sub




