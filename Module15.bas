Attribute VB_Name = "Module15"
Sub Move_sheet()
Sheets("Filtered").Copy Before:=Workbooks("Dashboard.xlsm").Sheets(1)
End Sub
