Attribute VB_Name = "Functions"

Sub Sync_FTP()
Dim RetVal
'RetVal = Shell("C:\Users\hmacnamara\Documents\Visual Studio 2013\Projects\WinSCP_UpdaterII\WinSCP_UpdaterII\obj\Debug\WinSCP_UpdaterII.exe", 1)
RetVal = Shell("U:\Operations\Operations\Admin\Projects\Code\Visual Studio\WinSCP_L1\WinSCP_L1\bin\release\WinSCP_L1.exe", 1)
'or run winscp.com from network location using VBA
End Sub

Public Function EndCellRow(Optional rng)
Dim eRow As Long

If rng Is Nothing Then
    EndCellRow = ActiveCell.End(xlDown).Row
Else
    EndCellRow = rng.End(xlDown).Row
End If
End Function


Public Sub UnHide_Sheets()
Attribute UnHide_Sheets.VB_ProcData.VB_Invoke_Func = "U\n14"
Dim wsThis, ws As Worksheet
Dim wb As Workbook

Set wsThis = ActiveSheet
Set wb = activeworkbook
For Each ws In wb.Worksheets
    ws.Visible = xlSheetVisible
Next ws
wsThis.Activate
End Sub

Public Sub Key_Dates()
Dim dtPriorMonthEnd, dtPriorYearEnd, dtIncept, dtThisMonthEnd As Date
Dim wbPnL As Workbook

Set wbPnL = activeworkbook

End Sub



Public Function Name_Check(strFind As String) As Boolean
'checks if activeworkbook name contains a string argument

Dim strWB As String
'Dim res As Variant
strFind = LCase(strFind)
strWB = LCase(activeworkbook.Name)
Res = InStr(1, strWB, strFind)
If IsNumeric(Res) And Res > 0 Then
    Name_Check = True
Else
    Name_Check = False
End If

End Function

Public Function Cxl_Filter(Optional ws As Worksheet)

If ws.AutoFilterMode = True Then
    ws.AutoFilter.ShowAllData
End If


End Function
Public Function Cxl_All_Filtered_Sheets(Optional wb As Workbook)
Dim ws1 As Worksheet
'Dim wb As Workbook
Dim ws As Worksheet

If wb Is Nothing Then Set wb = activeworkbook
Set ws1 = ActiveSheet
For Each ws In wb.Worksheets
    Cxl_Filter ws
Next ws
ws1.Activate
End Function

Sub Filter_Showall(Optional wsO As Worksheet)
Dim ws As Worksheet

If Not wsO Is Nothing Then
    Set ws = wsO
End If
Set ws = ActiveSheet
On Error GoTo Crash
If ws.AutoFilter.FilterMode = True Then
    ws.AutoFilter.ShowAllData
End If

Exit1:
Exit Sub

Crash:
Select Case True
Case Err.Number = 91
    On Error GoTo 0
    Resume Exit1
End Select
End Sub
Public Function wb_Check(str As String)

str = "Cpty OTC Val"
If Name_Check(LCase(str)) = False Then
    msg = MsgBox("Select Correct " & str & " Workbook!", vbOKOnly)
    Exit Function
End If
End Function


Public Function Refresh_Sheet_Pivot_Tables(Optional ws As Worksheet)
Dim i As Integer

If ws Is Nothing Then
    Set ws = ActiveSheet
End If

On Error GoTo Crash
With ws
    .Activate
    For Each pt In .PivotTables
        pt.PivotCache.Refresh
        i = i + 1
        
    Next pt
End With
Refresh_Sheet_Pivot_Tables = i & " Pivot Table(s) Refreshed!"
Crash:
End Function

Public Function Refresh_WBook_Pivot_Tables()
Dim wb As Workbook
Dim ws As Worksheet

Set wb = activeworkbook
With wb
    For Each ws In wb.Sheets
        'ws.Activate
        Call Refresh_Sheet_Pivot_Tables(ws)
    Next ws
End With

End Function


Public Sub Open_Code()
Dim wb As Workbook
Dim Isopen As Boolean
Dim msg As String

For Each wb In Workbooks
    If wb.Name = "Operations VBA.xlsm" Then Isopen = True
Next wb
If Isopen = False Then
    Workbooks.Open ("U:\Operations\Operations\Admin\Projects\Code\Operations VBA.xlsm")
End If
msg = MsgBox("Hi")
End Sub


Public Function Unique_Array(rng As Range)
'convert range to single dimensioned array

Dim idOld() As Variant
Dim idUnique(), str As String
Dim i, N As Single

'rng
'array whole
idOld = rng
N = UBound(idOld, 1)

'pass to dictionary
Set d = CreateObject("scripting.dictionary")

'...if NOT exists
For i = 1 To N
    str = idOld(i, 1)
    If Not d.Exists(str) And str <> "" Then
        'j = j + 1
        d.Add str, 1
    End If
'Loop
Next i

'pass back to array
Set s = CreateObject("scripting.dictionary")
s = d.Keys
For i = 1 To d.Count
    ReDim Preserve idUnique(1 To i)
    idUnique(i) = s(i - 1)
Next i

Unique_Array = idUnique
End Function





Public Function DirMax(strDir As String, Optional str1 As String)
'Find most recent File with a date in it's name
'

Dim strFile As String, str As String
Dim dtRec As Date, dtMax As Date
Dim iYear As Integer, iMonth As Integer, iDay As Integer, iRow As Integer

If str1 = "" Then str1 = "*"
'If str2 = "" Then str1 = "*"

strFile = Dir(strDir & "*" & str1 & "*")

'find max date of pos file
Do While strFile <> ""
    
    str = Mid(strFile, InStr(1, strFile, "201", 1), 8) ', "yyyymmdd")
    iYear = Left(str, 4)
    iMonth = Mid(str, 5, 2)
    iDay = Right(str, 2)
    dtRec = DateSerial(iYear, iMonth, iDay)
    If dtRec > dtMax Then
        dtMax = dtRec
        DirMax = strFile
    End If
    strFile = Dir()
Loop

End Function

Public Function DirMax2(strDir As String, strMain As String, Optional str1 As String, Optional str2 As String)
'Find most recently modified File with up to 3 strings in its name. Requires: directory path and one search string. Insert wildcards as required
'strDir=Directory Path
'strMain= main search string
'str1=string prefix
'str2=string suffix

Dim dtMod As Date, dtMax As Date
Dim iYear As Integer, iMonth As Integer, iDay As Integer, iRow As Integer
'
'If str1 = "" Then str1 = "*"
'If str2 = "" Then str2 = "*"

If Right(strDir, 1) <> "\" Then strDir = strDir & "\"

strFile = Dir(strDir & str1 & strMain & str2)

'find max date of pos file
Do While strFile <> ""
    
'    str = Mid(strFile, InStr(1, strFile, "201", 1), 8) ', "yyyymmdd")
'    iYear = Left(str, 4)
'    iMonth = Mid(str, 5, 2)
'    iDay = Right(str, 2)
'    dtRec = DateSerial(iYear, iMonth, iDay)

    dtMod = FileDateTime(strDir & strFile)
    
    If dtMod > dtMax Then
        dtMax = dtMod
        DirMax2 = strFile
    End If
    
    strFile = Dir()
Loop
'If strFile = "" Then
'
'End If
End Function


Public Function DirMax3(strDir As String, str1 As String, Optional str2 As String, Optional str3 As String)
'Find most recently modified File with up to 3 strings in its name. Requires: directory path and one search string. Insert wildcards as required
'strDir=Directory Path
'str1= 1st search string (required)
'str2=2nd search string (optional)
'str3=3rd search string (optional)

Dim dtMod As Date, dtMax As Date
Dim iYear As Integer, iMonth As Integer, iDay As Integer, iRow As Integer
'
'If str1 = "" Then str1 = "*"
'If str2 = "" Then str2 = "*"

If Right(strDir, 1) <> "\" Then strDir = strDir & "\"

strFile = Dir(strDir & str1 & "*" & str2)

'find max date of pos file
Do While strFile <> ""
    
'    str = Mid(strFile, InStr(1, strFile, "201", 1), 8) ', "yyyymmdd")
'    iYear = Left(str, 4)
'    iMonth = Mid(str, 5, 2)
'    iDay = Right(str, 2)
'    dtRec = DateSerial(iYear, iMonth, iDay)

    dtMod = FileDateTime(strDir & strFile)
    
    If dtMod > dtMax Then
        dtMax = dtMod
        DirMax3 = strFile
    End If
    
    strFile = Dir()
Loop
'If strFile = "" Then
'
'End If
End Function
Public Function DirMax4(strDir As String, strMain As String, Optional str1 As String, Optional str2 As String)
'Find most recently modified File with up to 3 strings in its name. Requires: directory path and one search string. Insert wildcards as required
'this is an attempt on improved version that correctly interprets null strings as wildcards
'strDir=Required Directory Path
'strMain=Required  main search string
'str1=Optional string prefix
'str2=Optional string suffix

Dim dtMod As Date, dtMax As Date
Dim iYear As Integer, iMonth As Integer, iDay As Integer, iRow As Integer
'
If str1 = "" Then str1 = "*" Else str1 = str1 & "*"
If str2 = "" Then str2 = "*" Else str2 = "*" & str2

If Right(strDir, 1) <> "\" Then strDir = strDir & "\"

strFile = Dir(strDir & str1 & strMain & str2)

'find max date of pos file
Do While strFile <> ""
    
'    str = Mid(strFile, InStr(1, strFile, "201", 1), 8) ', "yyyymmdd")
'    iYear = Left(str, 4)
'    iMonth = Mid(str, 5, 2)
'    iDay = Right(str, 2)
'    dtRec = DateSerial(iYear, iMonth, iDay)

    dtMod = FileDateTime(strDir & strFile)
    
    If dtMod > dtMax Then
        dtMax = dtMod
        DirMax4 = strFile
    End If
    
    strFile = Dir()
Loop
'If strFile = "" Then
'
'End If
End Function
Public Function Find_Last_Real_Row(Optional ws As Worksheet)
'finds row of last populated cell in a sheet

If ws Is Nothing Then
    Set ws = ActiveSheet
End If
On Error GoTo Crash
Find_Last_Real_Row = ws.Cells.Find("*", Range("A1"), xlFormulas, , xlByRows, xlPrevious).Row
On Error GoTo 0

Exit Function
Crash:
If Err.Number = 91 Then
    Find_Last_Real_Row = 1
    Resume Next
Else
    Stop
End If

End Function


Public Function Fcn_Filter_Unique(Optional rng As Range) As Range

Dim Rw As Long

On Error GoTo Crash
If rng Is Nothing Then
    Set rng = Selection
End If
Rw = Cells.Find("*", Range("A1"), xlFormulas, , xlByRows, xlPrevious).Row + 2
rng.Select
rng.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range( _
    "A" & Rw), Unique:=True
Rw = Rw + 1
If Range("B" & Rw).Value = "" Then
    Range(Range("A" & Rw), Range("A" & Rw).End(xlDown)).Select
Else
    Range(Range("A" & Rw), Range("A" & Rw).End(xlToRight).End(xlDown)).Select
End If
Set Fcn_Filter_Unique = Selection
Exit Function
    
Crash:
    Set Fcn_Filter_Unique = Nothing
End Function


Public Sub test()
Dim v() As String
Dim i As Single
Dim bool As Boolean

bool = Select_Data_Range(Range("A2"), 18)
End Sub


Public Function Select_Data_Range(rngStart As Range, Optional lColEnd As Long, Optional lRowEnd As Long, Optional ws As Worksheet)
'  Input Starting range, worksheet (optional), End column (optional) and end row (optional) to select the range bounded within
'Endrow is the last cell with data in in the lcolend specifies, else end row row of last real cell


If ws Is Nothing Then
    Set ws = rngStart.Worksheet
End If

On Error GoTo Crash

If lColEnd = 0 Then
    Set rng = ws.Cells.Find("*", Range("A1"), xlFormulas, , xlByColumns, xlPrevious)
    lColEnd = rng.Column
End If

If lRowEnd = 0 Then
    Set rng = ws.Cells(1, lColEnd)
    Set rng = rng.EntireColumn.Find("*", rng, xlFormulas, , xlByColumns, xlPrevious)
    lRowEnd = rng.Row
End If

Set rng = Range(rngStart, ws.Cells(lRowEnd, lColEnd))
rng.Select
Select_Data_Range = True
Exit Function
Crash:

Clear_Data_Range = False
On Error GoTo 0
End Function

Public Sub Clean_QueryTable_Connectionss()  '27/1/2016 Worked! didn't tidy up sorry!

Dim wb As Workbook
Dim ws As Worksheet
Dim qt As QueryTable
Dim rng As Range
Dim nm As Name
Dim nqt As Integer, iqt As Integer

Set wb = activeworkbook
For Each ws In wb.Worksheets
    Debug.Print ws.Name
    nqt = ws.QueryTables.Count
    iqt = 0
'    Do Until iqt = nqt
'        Set qt = ws.QueryTables(iqt + 1)
'
'        Debug.Print "   " & qt.Name & "   " & qt.Connection
''        For Each cn In qt.Connection
''            Debug.Print
''        '    'Debug.Print nm.RefersTo
''        '    If nm.Name Like "*Expected_CA*" Or nm.Name Like "*Current_Deposits*" Or nm.Name Like "*CASH_BAL_BOOK*" Or nm.Name Like "*BalanceSummary*" _
''        '    Or nm.Name Like "*Near_Cash_Asset*" Or nm.Name Like "*Unsettled_Trades*" Or nm.Name Like "*Unsettled_Trades*" _
''        '    Or nm.Name Like "*L1_Off_All_Transfers_20_days_hist*" Or nm.Name Like "*L1_Off_Unrealised_Listed*" Or nm.Name Like "*L1_Off_All_Transfers_Forecast*" _
''        '    Or nm.RefersTo Like "=*'C:\Op*" Or Format(nm.RefersTo, Text) Like "*#REF*" Then
''        '        'Stop
''        '        Debug.Print nm.Name & "   " & nm.RefersTo
''        '        nm.Delete
''        '    End If
''        Next cn
'        qt.Delete
'        iqt = iqt + 1
'    Loop
For Each qt In ws.QueryTables
    
    Debug.Print "   " & qt.Name & "   " & qt.Connection
    qt.Delete
Next qt

Next ws
End Sub


Public Sub Names_Manager_Listing()
'List all Named Ranges in Workbook. Fields: Name    Refers To   Scope   Comment
Dim rng As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim str As String
Dim i As Integer
Dim N As Name

Set wb = activeworkbook
Set ws = wb.Sheets("1")
Set rng = ws.Range("A4")

For Each N In wb.Names
    rng.Value = """" & N.Name & """"
    rng.Offset(0, 1).Value = """" & N.RefersTo & """"
'    Rng.Offset(0, 2).Value = n.Scope
'    Rng.Offset(0, 3).Value = n.Comment
    Set rng = rng.Offset(1, 0)
Next N
End Sub

Public Sub Names_Manager_Listing_Import()
'List all Named Ranges in Workbook. Fields: Name    Refers To   Scope   Comment
Dim rng As Range
Dim wb As Workbook
Dim ws As Worksheet
Dim str As String
Dim i As Integer, x As Integer
Dim N As Name

Set wb = activeworkbook
Set ws = wb.Sheets("1")
Set rng = ws.Range("A4")
x = rng.End(xlDown).Row
x = Range(rng, "a" & x).Rows.Count - 1

'For Each n In wb.Names
For i = 0 To x
    On Error GoTo Crash
    strName = rng.Offset(i, 0).Value
    strName = Mid(strName, 2, Len(strName) - 2)
    strRefersTo = rng.Offset(i, 1).Value
    strRefersTo = Mid(strRefersTo, 2, Len(strRefersTo) - 2)
    wb.Names.Add strName, strRefersTo
    On Error GoTo 0
'    Rng.Value = """" & n.Name & """"
'    Rng.Offset(0, 1).Value = """" & n.RefersTo & """"
'    Rng.Offset(0, 2).Value = n.Scope
'    Rng.Offset(0, 3).Value = n.Comment
    Set rng = rng.Offset(1, 0)
Next i
Exit Sub
Crash:
    Stop
    Resume Next
End Sub


Public Function Unzip_Files(strTarget, strDest, Optional Clear_Folder_at_Start As Boolean) As Boolean
'Copies Target Files to destination folder. Clear_Folder_at_Start As Boolean toggles delete existing folders option
Dim oApp As Object
Dim strTargetZip, vba_Folder_Items, FolderItem


On Error GoTo Crash
If Right(strDest, 1) <> "\" Then strDest = strDest & "\"

If Clear_Folder_at_Start = True Then Kill strDest & "*.*"

Set oApp = CreateObject("Shell.Application") 'Windows Shell
Set vba_Folder_Items = oApp.Namespace(strTarget).Items

oApp.Namespace(strDest).CopyHere vba_Folder_Items
strTargetZip = oApp.Namespace(strDest).Title
'Pause


Unzip_Files = True
Exit Function
Crash:
Unzip_Files = False

End Function
Public Function Unzip_Files_iRecs(strTarget, strDst, Optional Clear_Folder_at_Start As Boolean) As Boolean
'Copies Target Files to destination folder. Clear_Folder_at_Start As Boolean toggles delete existing folders option
Dim oApp As Object
Dim strTargetZip, vba_Folder_Items, FolderItem


'On Error GoTo Crash
If Right(strDst, 1) <> "\" Then strDst = strDst & "\"

If Clear_Folder_at_Start = True Then Kill strDst & "*.*"

Set oApp = CreateObject("Shell.Application") 'Windows Shell
Set vba_Folder_Items = oApp.Namespace(strTarget).Items
strTargetZip = oApp.Namespace(strDest).Title 'test for folder object FAIL!!!
oApp.Namespace(strDst).CopyHere vba_Folder_Items
'Pause
'oApp.Namespace(strDst) = Nothing

Unzip_Files_iRecs = True
Exit Function
Crash:
Unzip_Files_iRecs = False

End Function

Public Sub Fetch_Links()

Dim aLinks As Variant, aLinkInfo As Variant
Dim app As Application
Dim wb As Workbook

Set app = activeworkbook.Application
Set wb = activeworkbook
aLinks = wb.LinkSources(xlExcelLinks)

If Not IsEmpty(aLinks) Then
Sheets.Add
For i = 1 To UBound(aLinks)
    Cells(i, 1).Value = aLinks(i)
    Cells(i, 2).Value = wb.LinkInfo(aLinks(i), xlEditionDate, , 1)
Next i
End If
End Sub


Sub ShowAllLinksInfo()
'Original Author:        JLLatham
'Purpose:       Identify which cells in which worksheets are using Linked Data
'Requirements:  requires a worksheet to be added to the workbook and named LinksList
'Modified From: http://answers.microsoft.com/en-us/office/forum/office_2007-excel/workbook-links-cannot-be-updated/b8242469-ec57-e011-8dfc-68b599b31bf5?page=1&tm=1301177444768
    Dim aLinks           As Variant
    Dim i                As Integer
    Dim wb               As Workbook
    Dim ws               As Worksheet
    Dim anyWS            As Worksheet
    Dim anyCell          As Range
    Dim reportWS         As Worksheet
    Dim nextReportRow    As Long
    Dim shtName          As String
    Dim bWsExists        As Boolean
 
    shtName = "LinksList"
    Set wb = activeworkbook
    'Create the result sheet if one does not already exist
    For Each ws In Application.Worksheets
        If ws.Name = shtName Then bWsExists = True
    Next ws
    If bWsExists = False Then
        Application.DisplayAlerts = False
        Set ws = activeworkbook.Worksheets.Add(Type:=xlWorksheet)
        ws.Name = shtName
        ws.Select
        ws.Move After:=activeworkbook.Worksheets(activeworkbook.Worksheets.Count)
        Application.DisplayAlerts = True
    End If
 
    'Now start looking of linked data cells
    Set reportWS = wb.Worksheets(shtName)
    reportWS.Cells.Clear
    reportWS.Range("A1") = "Worksheet"
    reportWS.Range("B1") = "Cell"
    reportWS.Range("C1") = "Formula"
 
    aLinks = activeworkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(aLinks) Then
        'there are links somewhere in the workbook
        For Each anyWS In wb.Worksheets
            If anyWS.Name <> reportWS.Name Then
                For Each anyCell In anyWS.UsedRange
                    If anyCell.HasFormula Then
                        If InStr(anyCell.Formula, "[") > 0 Then
                            nextReportRow = reportWS.Range("A" & Rows.Count).End(xlUp).Row + 1
                            reportWS.Range("A" & nextReportRow) = anyWS.Name
                            reportWS.Range("B" & nextReportRow) = anyCell.Address
                            reportWS.Range("C" & nextReportRow) = "'" & anyCell.Formula
                        End If
                    End If
                Next    ' end anyCell loop
            End If
        Next    ' end anyWS loop
    Else
        MsgBox "No links to Excel worksheets detected."
    End If
    'housekeeping
    Set reportWS = Nothing
    Set ws = Nothing
End Sub

Public Sub connection_cleanup() ' Careful. this deletes all connections (except those that are open)
Dim conns As Connections
'Dim c As Connection
Dim wb As Workbook

Set wb = activeworkbook
Set conns = wb.Connections

For Each c In conns
    Debug.Print c
    c.Delete
Next c
Debug.Print conns.Count
End Sub

Public Function ImportText(wb As Workbook, ws As Worksheet, rng As Range, strFilePath As String, StrFileName As String, ArrColDataTypes As Variant, _
                            bFieldNames As Boolean, StartRow As Single, strDelim As String, strdeci As String, strThousands As String) As String
'Generic Text File importer
Dim qt As QueryTable
On Error GoTo Crash
Set qt = ws.QueryTables.Add(Connection:= _
    "TEXT;" & strFilePath & StrFileName, Destination:=rng)
With qt
    .Name = ws.Name
    .FieldNames = bFieldNames
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .RefreshStyle = xlOverwriteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .TextFilePromptOnRefresh = False
    .TextFilePlatform = 850
    .TextFileStartRow = StartRow
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileOtherDelimiter = "|"
    .TextFileColumnDataTypes = ArrColDataTypes
        .TextFileDecimalSeparator = strdeci
        .TextFileThousandsSeparator = strThousands
    .TextFileTrailingMinusNumbers = False
    .Refresh BackgroundQuery:=False
End With
qt.Delete
ImportText = "True"
Exit Function
Crash:
    ImportText = strFilePath & StrFileName
    
End Function



Public Sub SaveAs_Defined(Optional iFormat, Optional strDir As String, Optional strInitial As String)
'Saves copy of the activeworkbook via a dialogue box which requests dir and filename. inputs are fileformat id, dir and filename

'Save as with dialogue
Dim strFile As String

If iFormat = 0 Then
    iFormat = activeworkbook.FileFormat
End If

ChDir strDir

Do
    'StrFile = Application.GetSaveAsFilename(strDir & strInitial, iFormat, , "Save to '" & strDir & "'?")  ' For Dev to include filefilter eg
    strFile = Application.GetSaveAsFilename(strDir & strInitial, , , "Save to '" & strDir & "'?")
Loop Until strFile <> "False"
strFile = Left(strFile, InStr(strFile, ".") - 1)
activeworkbook.SaveAs FileName:=strFile, FileFormat:=iFormat
'ActiveWorkbook.SaveAs (StrFile)

End Sub

Public Sub test_proc()

Dim strInitial As String, strDir As String, strDt As String
Dim strformat As String
Dim wb As Workbook
Dim iFormat  As Integer

Set wb = activeworkbook
iFormat = wb.FileFormat
'strformat=
strDt = Format(Now(), "yyyymmdd")
strInitial = "Enfusion Cash Activity ITD " & strDt
strDir = "c:\temp\"

Call SaveAs_Input(iFormat, strDir, strInitial)
End Sub

Public Function Find_Last_Real_Cell() As Range
'Returns a cell at the intersection of the last real row & column

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
Set Find_Last_Real_Cell = Cells(lRealLastRow, lRealLastColumn)

ActiveSheet.UsedRange 'Resets LastCell
     
Exit Function

Crash1:
    lRealLastRow = 1
    Resume Next
Crash2:
    lRealLastColumn = 1
    Resume Next
End Function

Public Sub List_WSs()
Dim ws As Worksheet
For Each ws In Worksheets
    Debug.Print ws.Name
Next ws
End Sub



Public Function TransposeArray(InputArr As Variant, OutputArr As Variant) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TransposeArray
' This transposes a two-dimensional array. It returns True if successful or
' False if an error occurs. InputArr must be two-dimensions. OutputArr must be
' a dynamic array. It will be Erased and resized, so any existing content will
' be destroyed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim RowNdx As Long
Dim ColNdx As Long
Dim LB1 As Long
Dim LB2 As Long
Dim UB1 As Long
Dim UB2 As Long

'''''''''''''''''''''''''''''''''''
' Ensure InputArr and OutputArr
' are arrays.
'''''''''''''''''''''''''''''''''''
If (IsArray(InputArr) = False) Or (IsArray(OutputArr) = False) Then
    TransposeArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''
' Ensure OutputArr is a dynamic
' array.
'''''''''''''''''''''''''''''''''''
If IsArrayDynamic(Arr:=OutputArr) = False Then
    TransposeArray = False
    Exit Function
End If

''''''''''''''''''''''''''''''''''
' Ensure InputArr is two-dimensions,
' no more, no lesss.
''''''''''''''''''''''''''''''''''
If NumberOfArrayDimensions(Arr:=InputArr) <> 2 Then
    TransposeArray = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''
' Get the Lower and Upper bounds of
' InputArr.
'''''''''''''''''''''''''''''''''''''''
LB1 = LBound(InputArr, 1)
LB2 = LBound(InputArr, 2)
UB1 = UBound(InputArr, 1)
UB2 = UBound(InputArr, 2)

'''''''''''''''''''''''''''''''''''''''''
' Erase and ReDim OutputArr
'''''''''''''''''''''''''''''''''''''''''
Erase OutputArr
ReDim OutputArr(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)

For RowNdx = LBound(InputArr, 2) To UBound(InputArr, 2)
    For ColNdx = LBound(InputArr, 1) To UBound(InputArr, 1)
        OutputArr(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
    Next ColNdx
Next RowNdx

TransposeArray = True

End Function

Public Sub Tint_Alternate_Rows()
With Cells
    .FormatConditions.Delete
    .FormatConditions.Add Type:=xlExpression, Formula1:="=ROW()/2-ROUNDUP(ROW()/2,0)<>0"
    .FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
End With
With Cells.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
End With
Cells.FormatConditions(1).StopIfTrue = False
ActiveWindow.DisplayGridlines = False
End Sub

Public Function SHEETNAME(rng As Range) As String
'Retunrs sheet name
'Dim rng As Range

'i=application.WorksheetFunction.s
'Set rng = ActiveCell

SHEETNAME = rng.Worksheet.Name

End Function


Function Import_Single_xls_Report(arrConfig(), i As Integer, wb As Workbook, Optional rng As Range)

'-------------------NB THIS CODE HAS DIVERGED FROM THE RISK VERSION!!!------------------
'WANT generic function that can be rolled out to any xls to be imported with appropriate flexibility

Dim wbCopy As Workbook
Dim ws As Worksheet
Dim rngStart As Range, rngCopy As Range
Dim strPath As String, strMain As String, strPrefix As String, strSuffix As String, strFile As String, arrForm(1 To 5) As String
Dim iMax As Integer, iRowStart As Integer, iRow As Integer, iColStart As Integer, iCol As Integer, N As Integer, iHdr As Integer
Dim dtMod As Date
Dim bool As Boolean

'Assign Variables
Set ws = wb.Sheets(arrConfig(i, 1))
If arrConfig(i, 2) = Empty Or arrConfig(i, 2) = "n/a" Then
    Set rngStart = ws.Range("A1")
Else
    Set rngStart = ws.Range(arrConfig(i, 2))
End If

strPath = arrConfig(i, 3)
If Right(strPath, 1) <> "\" Then
    strPath = strPath & "\"
End If
strMain = arrConfig(i, 4)
strPrefix = arrConfig(i, 5)
strSuffix = arrConfig(i, 6)
iMax = arrConfig(i, 9)
iHdr = arrConfig(i, 10)

'Find File
strFile = DirMax2(strPath, strMain, strPrefix, strSuffix)
dtMod = FileDateTime(strPath & strFile)

'copy or import data to next empty line
'Select Case bool
If DateValue(dtMod) = DateValue(arrConfig(i, 7)) And strSuffix <> "" Then ' If xls then it's a paste else import
    ws.Activate
    'show all if filter set
    Call Filter_Showall(ws)

    'clear old data
    iRow = Find_Last_Real_Cell.Row + 1
    
    'Check width of Import will not ovewrite permanaent formulae
    If iMax > 0 Then
        iCol = iMax
    Else '=================EXPLANATION REQUIRED===========
        If iColStart <> 0 Then '?? iColStart NOT DEFINED YET!
            iCol = iColStart - 2 'WHY DO WE ADJUST -2?
        Else
            iCol = Find_Last_Real_Cell.Column + 1
        End If
    End If
    Set rng = ws.Range(rngStart, Cells(iRow, iCol))
    rng.ClearContents

    'Copy & Paste New Data
    Workbooks.Open FileName:=strPath & strFile, ReadOnly:=True, notify:=False, Format:=2, Local:=True
'    Stop
    Set wbCopy = activeworkbook
    Set rngCopy = Range("a1", Cells(Range("a1").End(xlDown).Row, Range("a1").SpecialCells(xlCellTypeLastCell).Column)) 'ignores total
    rngCopy.Copy
    ws.Activate
    rngStart.Select
    ActiveSheet.Paste
'    iRow = ws.Range("A1").End(xlDown).Row
    Debug.Print wbCopy.Name
    wbCopy.Close False
    
End If
'Stop
Import_Single_xls_Report = strPath & strFile
End Function
Function Import_Enfusion_csv(wb As Workbook, arrConfig, i As Integer) As String
'import IR risk all csv

'text import spec
Dim ws As Worksheet
Dim rng As Range, rngClear As Range
Dim strDir As String, strFile As String, strSearchMain As String, strSuffix As String
Dim iHdr As Integer

Set ws = wb.Sheets(arrConfig(i, 1))
ws.Activate
Call Filter_Showall(ws)
If arrConfig(i, 2) = "n/a" Or arrConfig(i, 2) = "" Then
    Set rng = Range("A1")
Else
    Set rng = Range(arrConfig(i, 2))
End If

strDir = arrConfig(i, 3)
strSearchMain = arrConfig(i, 4)
strSuffix = arrConfig(i, 6)
iHdr = arrConfig(i, 10)
If iHdr = 0 Then
    iHdr = 1
End If
'Range(rng, Find_Last_Real_Cell).ClearContents 'Clear old data
Set rngClear = Cells(iHdr, rng.Column)
Set rngClear = Range(rngClear, Range(rngClear.End(xlDown), rngClear.End(xlToRight))) 'Clear old data
'Set rngClear = Range(rngClear, Cells(rngClear.End(xlDown).Row, rngClear.End(xlDown).Column)) 'Clear old data
'Stop
rngClear.Select
rngClear.ClearContents
rng.Activate

Import_Enfusion_csv = Search_and_import_csv_NEW(strDir, strSearchMain, rng, iHdr, , strSuffix, True) 'Return File name to function
'search_and_import_csv_NEW(strdir , strMain , rng As Range, iHdr As Integer, Optional str1 As String, Optional str2 As String, Optional Hdr As Boolean
End Function


Sub Mail_Selection_Range_Outlook_Body_COLLATERAL()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'https://www.rondebruin.nl/win/s1/outlook/bmail2.htm
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strDir As String, strFile As String
    Dim strYear As String, strMonth As String
    Dim dtPBD As Date
    
    'Print Range
    Range("_Print").Worksheet.Activate
    Set rng = Range("_Print")
        
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = rng.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'Define attachment
'    strFile
    dtPBD = Sheets("Summary").Range("_PBD")
    strMonth = Format(Month(dtPBD), "00")
    strYear = Year(dtPBD)
    strDir = "\\192.168.25.12\lts\Operations\Operations\" & strYear & "\" & strYear & strMonth & "\Collateral\"
    strFile = DirMax3(strDir, "Collateral CoB", "xlsm")
    
    'email Body
    'To List
    'CC List
    'BCC List
    'Subject
    
    
    
'    On Error Resume Next
    On Error GoTo Crash
    With OutMail
        .To = "ahamid@letterone.com; hmacnamara@letterone.com; jferber@letterone.com; crayner-cook@letterone.com; jlai@letterone.com; moprea@letterone.com; riskmanagement@letterone.com; dtalbot@letterone.com; edale@letterone.com; mhumphreys@letterone.com"
        .CC = "operations@letterone.com"
        .BCC = ""
        .Subject = "OTC Collateral Moves CoB " & dtPBD
        .HTMLBody = "MS Futures: LTS SA " & "<br>"
        .HTMLBody = .HTMLBody & "<br>"
        .HTMLBody = .HTMLBody & "n/a" & "<br>"
        .HTMLBody = .HTMLBody & "<br>"
        .HTMLBody = .HTMLBody & "Enfusion Collateral " & "<br>"
        .HTMLBody = .HTMLBody & RangetoHTML(rng)

        .attachments.Add (strDir & "\" & strFile)
        .display
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Exit Sub
Crash:
    Stop
    Resume
End Sub

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Function Mail_get_Config() As String
'Read Config from mail workbook
Dim wsConfig As Worksheet, wsPrint As Worksheet
Dim rngConfig As Range, rngPrint As Range
Dim strEmail(1 To 6) As Variant ', strBody As String, strFrom As String, strTo As String, strCC As String, strSubject As String


Set wsConfig = Range("_Config_email").Worksheet
wsConfig.Activate
Set rngConfig = Range("_Config_email")
'ReDim strEmail(1) as Range
Set strEmail(1) = Range(rngConfig.Offset(1, 1).Value)
strEmail(2) = rngConfig.Offset(2, 1).Value
strEmail(3) = rngConfig.Offset(3, 1)
strEmail(4) = rngConfig.Offset(4, 1)
strEmail(5) = rngConfig.Offset(5, 1)

'Set wsPrint = rngPrint.Worksheet
'wsPrint.Activate
'rngPrint.Select

End Function

Sub Mail_Selection_Range_Outlook_Body() 'strBody As String, strSubject As String, strTo As String, Optional rng As Range, Optional strCC As String, Optional strBCC As String
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'https://www.rondebruin.nl/win/s1/outlook/bmail2.htm
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strDir As String, strFile As String, strEmail() As Variant
    Dim strYear As String, strMonth As String
    Dim dtPBD As Date
    
    'Print Range
    'email Body
    'To List
    'CC List
    'BCC List
    'Subject
    Range("_Print").Worksheet.Activate
    ReDim strEmail(LBound(Mail_get_Config()) To UBound(Mail_get_Config()))
    strEmail = Mail_get_Config()
    
    Set rng = Range("_Print")
        
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = rng.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'Define attachment
'    strFile
    dtPBD = Sheets("Summary").Range("_PBD")
    strMonth = Format(Month(dtPBD), "00")
    strYear = Year(dtPBD)
    strDir = "\\192.168.25.12\lts\Operations\Operations\" & strYear & "\" & strYear & strMonth & "\Collateral\"
    strFile = DirMax3(strDir, "Collateral CoB", "xlsm")
    
    'email Body
    'To List
    'CC List
    'BCC List
    'Subject
    
    
    
'    On Error Resume Next
    On Error GoTo Crash
    With OutMail
        .To = "ahamid@letterone.com; hmacnamara@letterone.com; jferber@letterone.com; crayner-cook@letterone.com; jlai@letterone.com; moprea@letterone.com; riskmanagement@letterone.com; dtalbot@letterone.com; edale@letterone.com; mhumphreys@letterone.com"
        .CC = "operations@letterone.com"
        .BCC = ""
        .Subject = "OTC Collateral Moves CoB " & dtPBD
        .HTMLBody = "MS Futures: LTS SA " & "<br>"
        .HTMLBody = .HTMLBody & "<br>"
        .HTMLBody = .HTMLBody & "n/a" & "<br>"
        .HTMLBody = .HTMLBody & "<br>"
        .HTMLBody = .HTMLBody & "Enfusion Collateral " & "<br>"
        .HTMLBody = .HTMLBody & RangetoHTML(rng)

        .attachments.Add (strDir & "\" & strFile)
        .display
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Exit Sub
Crash:
    Stop
    Resume
End Sub

