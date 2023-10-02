Attribute VB_Name = "Module1"
Option Explicit


Sub multi_filter()

Sheets("database").Range("A1").AutoFilter Field:=10, Criteria1:=Array( _
"TSEL", "XL", "HCPT", "SMART", "SMART8", "ISAT"), Operator:=xlFilterValues
Sheets("database").Range("A1").AutoFilter Field:=63, Criteria1:=Array( _
"ONAIR"), Operator:=xlFilterValues
'ActiveSheet.Range("A1").AutoFilter Field:=64, Criteria1:=Array( _
"ONAIR"), Operator:=xlFilterValues
End Sub


