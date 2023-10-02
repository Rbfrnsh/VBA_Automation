Attribute VB_Name = "Module4"
Sub Insert_Multiple_Columns()
'insert multiple columns as columns B, C and D
Worksheets("Filtered").Range("O:O").EntireColumn.Insert
Worksheets("Filtered").Range("X:X").EntireColumn.Insert
Worksheets("Filtered").Range("Y:Y").EntireColumn.Insert
End Sub
Sub column_name()
Dim name1: name1 = Split("Revenue")
Dim name2: name2 = Split("Tenant_Rank1")
Dim name3: name3 = Split("year_end")

Worksheets("Filtered").Range("O:O").Resize(1, UBound(name1) + 1) = name1
Worksheets("Filtered").Range("X:X").Resize(1, UBound(name2) + 1) = name2
Worksheets("Filtered").Range("Y:Y").Resize(1, UBound(name3) + 1) = name3
End Sub
