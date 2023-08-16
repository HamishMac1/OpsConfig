Attribute VB_Name = "Formatting"


Private Sub Print_Header_Format()
'
' Print_Header_Format Macro

       With ActiveSheet.PageSetup
        .LeftHeader = "&Z &F"
        .CenterHeader = "&A"
        .RightHeader = "&T &D"
        .CenterFooter = "Page &P of &N"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
    End With
    

    
End Sub
Sub Grid_Toggle()
Dim v As Variant
v = ActiveWindow.DisplayGridlines
If v = True Then
    v = False
Else
    v = True
End If
ActiveWindow.DisplayGridlines = v
End Sub


Private Sub Red_Brackets_Per_Cent()
' Red_Brackets_Per_Cent Macro
    Selection.NumberFormat = "#,##0.00%;[Red](#,##0.00%);""-"""
End Sub

Public Sub Format_Red_Brackets()
Attribute Format_Red_Brackets.VB_ProcData.VB_Invoke_Func = "F\n14"

Selection.NumberFormat = "#,##0;[Red](#,##0);" & """-"""
'Selection.FormatConditions.Delete
End Sub

Sub Paste_Values()
Attribute Paste_Values.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' Paste_Values Macro
On Error GoTo Crash
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Crash:
    Exit Sub
End Sub

Sub Paste_Fomula()
Attribute Paste_Fomula.VB_ProcData.VB_Invoke_Func = "L\n14"

On Error GoTo Crash
Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
Crash:
    Exit Sub
End Sub
Sub Paste_Transpose()
Attribute Paste_Transpose.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' tran Macro
On Error GoTo Crash

Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=True
Crash:

End Sub
Sub Paste_multiplyer()
Attribute Paste_multiplyer.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' multi_ply Macro
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, _
        SkipBlanks:=False, Transpose:=False
End Sub

Private Sub List_All_Sheets()

Dim Wks As Worksheet

activeworkbook.Sheets.Add
Set Wks = ActiveSheet


For N = 1 To Worksheets.Count
        
        ActiveCell.Value = Worksheets(N).Name
        ActiveCell.Offset(1, 0).Select
        
Next N

End Sub


Public Sub ReformatNumbers()
Attribute ReformatNumbers.VB_ProcData.VB_Invoke_Func = "F\n14"
ReformatNumbersII
End Sub


Public Sub ReformatNumbers2(Optional rng As Range)

'activates cell, checks if null or less than 10e-7 and sets to zero
If rng Is Nothing Then Set rng = Selection

rng.Value = rng.Value

End Sub

Public Sub ReformatNumbers3()
Dim Arr() As Variant
Dim rng As Range

On Error GoTo Crash
Set rng = Selection
Arr = rng
'Stop

Dim Destination As Range
Set Destination = rng
Destination.Resize(UBound(Arr, 1), UBound(Arr, 2)).Value = Arr
''You can transpose the array when writing to the worksheet:

'Set Destination = Range("K1")
'Destination.Resize(UBound(Arr, 2), UBound(Arr, 1)).Value = Application.Transpose(Arr)
Exit Sub

Crash:
'Stop
Call ReformatNumbers2
End Sub
Public Sub RenameSheet()
Attribute RenameSheet.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' Macro1 Macro
' Macro recorded 06/10/2003 by J.P. Morgan
'

'
    Dim SHEETNAME As String
        SHEETNAME = InputBox("Input New Sheet Name", , ActiveSheet.Name)
    If SHEETNAME = "" Then
        SHEETNAME = ActiveSheet.Name
    Else
    End If
    
        ActiveSheet.Name = SHEETNAME
End Sub


Private Sub CalculateSheet()

'Refresh Worksheet calcs
Selection.Calculate

End Sub

Public Sub Last_Cell_CleanUp_Call()
'If ws is not specified then runs for all sheets in activeworkbook
Call Last_Cell_CleanUp

End Sub

Public Sub Last_Cell_CleanUp(Optional ws As Worksheet)
'If ws is not specified then runs for all sheets in activeworkbook
'ReReferences LastCell to actual last cell in sheet
Dim i As Single
Dim str As String
Dim w As Worksheet
Dim rng As Range

'Dim N As Single
str = Application.Calculation
Set rng = ActiveCell
Application.Calculation = xlCalculationManual
If Not ws Is Nothing Then
    ws.Activate
    Call Filter_Showall(ws)
    Call DeleteUnusedFormats
    
Else
    
    For Each w In Worksheets
        w.Activate
        Call Filter_Showall(w)
        If w.ProtectContents = False Then
            Call DeleteUnusedFormats
        End If
    Next w
End If

On Error Resume Next
rng.Activate
On Error GoTo 0

Application.Calculation = str
End Sub

Public Sub DeleteUnusedFormats()
     Dim lLastRow As Long, lLastColumn As Long
     Dim lRealLastRow As Long, lRealLastColumn As Long
     With Range("A1").SpecialCells(xlCellTypeLastCell)
         lLastRow = .Row
         lLastColumn = .Column
     End With
     On Error GoTo Crash1:
     lRealLastRow = Cells.Find("*", Range("A1"), xlFormulas, , xlByRows, xlPrevious).Row
     On Error GoTo Crash2:
     lRealLastColumn = Cells.Find("*", Range("A1"), xlFormulas, , _
               xlByColumns, xlPrevious).Column
     On Error GoTo 0
     If lRealLastRow < lLastRow Then
         Range(Cells(lRealLastRow + 1, 1), Cells(lLastRow, 1)).EntireRow.Select
         'Stop
         Range(Cells(lRealLastRow + 1, 1), Cells(lLastRow, 1)).EntireRow.Delete
         'Stop
     End If
     If lRealLastColumn < lLastColumn Then
         Range(Cells(1, lRealLastColumn + 1), _
              Cells(1, lLastColumn)).EntireColumn.Delete
     End If
     ActiveSheet.UsedRange 'Resets LastCell
     On Error GoTo Crash0
     ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell) = Cells(lRealLastRow, lRealLastColumn)
     ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Select
'     Stop
Skip:
     
     Exit Sub
Crash1:
    lRealLastRow = 1
    Resume Next
Crash2:
    lRealLastColumn = 1
    Resume Next
Crash0:
    GoTo Skip
End Sub


Private Sub Cycle_sheets()
Worksheets(1).Activate
For Each Worksheet In Worksheets
    Call Print_Header_Format
    ActiveSheet.Next.Activate
    On Error GoTo 0
Next Worksheet
End Sub
Private Sub PageSetup()
'
' PageSetup Macro
'

'
Dim i As Integer
Dim myArray As Variant
Dim SHEETNAME As String

    myArray = Array(Sheets(Array("Legend", "Bench %", "Spot %", "Tracking HG", "Inter MAP Portfolio HG" _
        , "Conservative HG", "Cons", "Moderate HG", "Mod", "Plus HG", "Plus", "Balanced HG", _
        "Bal", "Cautious HG", "Cau", "GH", "MF", "Comment")))
    For i = 0 To 18
    myArray(i).Select
        With ActiveSheet.PageSetup
            .LeftHeader = "&Z &F"
            .CenterHeader = "&A"
            .RightHeader = "&T &D"
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.511811023622047)
            .TopMargin = Application.InchesToPoints(0.433070866141732)
            .BottomMargin = Application.InchesToPoints(0.275590551181102)
            .HeaderMargin = Application.InchesToPoints(0.275590551181102)
            .FooterMargin = Application.InchesToPoints(0.15748031496063)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = True
            .CenterVertically = False
            .Orientation = xlLandscape
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = xlPrintErrorsDash
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = False
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
    Next i
End Sub

Private Sub Filter_Unique()

    Dim Rw As Long
    
    Set rng = Selection
    Rw = Cells.Find("*", Range("A1"), xlFormulas, , xlByRows, xlPrevious).Row + 2
    rng.Select
    rng.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range( _
        "A" & Rw), Unique:=True
    Range(Range("A" & Rw), Range("A" & Rw).End(xlToRight).End(xlDown)).Select
End Sub

Private Sub Paste_Transpose_Values()

    Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
End Sub

Private Sub Reverse_Date()
Dim dd, mm, yyyy As Integer
Dim Lfirst, Llast, lCol As Long
Dim str As String
Dim rng As Range
Dim Dat As Date

Application.DisplayAlerts = False
str = Application.Calculation
Application.ScreenUpdating = False

Set rng = Selection
lCol = rng.Column
Lfirst = rng(1).Row
Llast = rng.Rows.Count
If Cells(Llast, lCol).Value = "" Then
    Llast = Cells(Llast, lCol).End(xlUp).Row
End If
rng.EntireColumn.Insert (xlShiftToRight)
Set rng = Range(Cells(Lfirst, lCol), Cells(Llast, lCol))
rng.Select
rng.FormulaR1C1 = "=DATE(LEFT(RC[1],4),MID(RC[1],5,2),RIGHT(RC[1],2))"
rng.Value = rng.Value

Application.DisplayAlerts = True
Application.Calculation = str
Application.ScreenUpdating = True
End Sub


Private Sub Evaluate_Cell_Formula()
Dim c As Range

For Each c In Selection
    c.Formula = c.Formula
Next c
End Sub

Public Sub T()
Attribute T.VB_ProcData.VB_Invoke_Func = "R\n14"

If Application.ReferenceStyle = xlR1C1 Then
    Application.ReferenceStyle = xlA1
Else
    Application.ReferenceStyle = xlR1C1
End If
End Sub


Sub AutoFilter_Off(ws)


If Not ws.AutoFilter Is Nothing Then
    ws.AutoFilter.ShowAllData
End If
End Sub

Private Sub Unhide_All()
Dim ws As Worksheet

For Each ws In Sheets
    ws.Visible = xlSheetVisible
Next ws

End Sub
Public Sub UnMergeFill()

Dim cell As Range, joinedCells As Range
'set rngUsed
For Each cell In ActiveSheet.UsedRange
    If cell.MergeCells Then
        Set joinedCells = cell.MergeArea
        cell.MergeCells = False
        joinedCells.Value = cell.Value
    End If
Next

End Sub

Public Sub Go_Sheet1()
Attribute Go_Sheet1.VB_ProcData.VB_Invoke_Func = "H\n14"
activeworkbook.Sheets(1).Activate
End Sub

Public Sub Go_Sheet_End()
Attribute Go_Sheet_End.VB_ProcData.VB_Invoke_Func = "E\n14"
activeworkbook.Sheets(activeworkbook.Sheets.Count).Activate
End Sub



Public Sub Goto_Sheet()
Attribute Goto_Sheet.VB_Description = "Goto Sheet"
Attribute Goto_Sheet.VB_ProcData.VB_Invoke_Func = "G\n14"

'Dim ws As Worksheet
'Dim str As String
'
'str = ActiveCell.Value
On Error GoTo Crash
Worksheets(ActiveCell.Value).Activate

Exit Sub
Crash:
'    Stop
End Sub
Public Sub Goto_Home_Sheet()

'Dim ws As Worksheet
'Dim str As String
'
'str = ActiveCell.Value
On Error GoTo Crash
Worksheets(Range("_Home").Value).Activate

Exit Sub
Crash:
'    Stop
End Sub







