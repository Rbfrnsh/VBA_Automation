Attribute VB_Name = "Module2"
Sub CopyVisibleCells()
Dim wsSource    As Worksheet
Dim wsDest      As Worksheet
Dim LR          As Long

Application.ScreenUpdating = False
Set wsSource = ThisWorkbook.Worksheets("database")    'Source Sheet you want to copy data from
Set wsDest = ThisWorkbook.Worksheets("Filtered")      'Destination sheet where you want to paste the data

wsDest.Range("A1").CurrentRegion.Offset(1).Columns("A:CB").ClearContents 'Clearing the existing data from the destination sheet

LR = wsSource.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

wsSource.Range("A1:CB" & LR).SpecialCells(xlCellTypeVisible).Copy
wsDest.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

Application.CutCopyMode = 0
Application.ScreenUpdating = True
End Sub

Sub Create_Dynamic_Table1()
Dim tbOb As ListObject
Dim TblRng As Range
With Sheets("Filtered")
lLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
lLastColumn = .Cells(1, .Columns.Count).End(xlToLeft).column
Set TblRng = .Range("A1", .Cells(lLastRow, lLastColumn))
Set tbOb = .ListObjects.Add(xlSrcRange, TblRng, , xlYes)
tbOb.name = "DynamicTable1"
tbOb.TableStyle = "TableStyleMedium2"
End With
End Sub
