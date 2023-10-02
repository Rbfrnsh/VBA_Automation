Attribute VB_Name = "Module9"
Sub Tenant_Rank()
Attribute Tenant_Rank.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Tenant_Rank Macro
'

'
With Worksheets("Filtered")
Range("X2").Select
'Range(Selection, Selection.End(xlDown)).NumberFormat = "General"
ActiveCell.FormulaR1C1 = _
    "=SUMPRODUCT(([SiteID]=[@SiteID])*([@[Rental_start]]>[Rental_start]), ([SiteName]=[@SiteName])*([@[Rental_start]]>[Rental_start]))+1"
Range("X3").Select
Range("X2", Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlDown)).NumberFormat = "General"
End With
    
End Sub
