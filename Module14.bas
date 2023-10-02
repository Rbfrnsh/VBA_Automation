Attribute VB_Name = "Module14"
Option Explicit


Sub multi_filter2()

Sheets("database").Range("A1").AutoFilter Field:=10, Criteria1:=Array( _
"TSEL", "XL", "HCPT", "SMART", "SMART8", "ISAT"), Operator:=xlFilterValues
Sheets("database").Range("A1").AutoFilter Field:=63, Criteria1:=Array( _
"ONAIR"), Operator:=xlFilterValues
Sheets("database").Range("A1").AutoFilter Field:=29, Criteria1:=Array( _
"COLO 3G", "COLO", "COLO IBS", "COLO MAKRO", "COLO MCP NON FO", "Colocation MMP (Fiberization)", "HANDHOLE FIBERIZE", "IBS/DAS", "INTERSITE FO", "MMP (Fiberization)", "NEW MCP FO", "NEW MCP NON FO", "Sewa MPP", "SITE ACCESS", "SWAP ANTENNA", "TOWER", "UP GRADE 3G"), Operator:=xlFilterValues
'ActiveSheet.Range("A1").AutoFilter Field:=64, Criteria1:=Array( _
"ONAIR"), Operator:=xlFilterValues
End Sub




