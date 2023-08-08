Attribute VB_Name = "Main"
Option Explicit


Public Function Get_Config()

Dim wbThis As Workbook, wbOut As Workbook
Dim ws As Worksheet
Dim rng As Range
Dim row As Single, col As Single
Dim i As Long
Dim Arr() As Variant ' declare an unallocated array.



Set wbThis = ThisWorkbook
Set ws = Worksheets("CF")
Set rng = ws.Range("_Outlook").Offset(1, 0)
col = rng.End(xlToRight).Column
row = rng.End(xlDown).row
Set rng = Range(rng, Cells(row, col))
rng.Select
Arr = rng ' Arr is now an allocated array

Get_Config = Arr

End Function


Public Sub Search_Mail(Arr)


End Sub
