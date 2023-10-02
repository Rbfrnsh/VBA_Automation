Attribute VB_Name = "Module6"
Sub Sort_Table_Based_on_Multiple_Columns()

Worksheets("Filtered").Range("DynamicTable1").Sort Key1:=Range("DynamicTable1[SiteID]"), _
Order1:=xlAscending, _
Header:=xlYes, _
Key2:=Range("DynamicTable1[SiteName]"), _
Order2:=xlAscending, _
Key3:=Range("DynamicTable1[Rental_start]"), _
Order3:=xlAscending

End Sub

