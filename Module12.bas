Attribute VB_Name = "Module12"
Sub Rename_Column_tenant()
Dim name1: name1 = Split("Tenant_Rank")

Worksheets("Filtered").Range("Y:Y").Resize(1, UBound(name1) + 1) = name1
End Sub
