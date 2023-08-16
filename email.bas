Attribute VB_Name = "email"
Option Explicit

Sub email_List_Addresses()
Dim OutApp As Outlook.Application
Dim outNs As Outlook.Namespace ' The Namespace Object (Session) has a collection of accounts.
Dim DictAddress As Object, DictName As Object, DictDelete As Object
Dim outFldr As Outlook.MAPIFolder
Dim msgs As Outlook.Items
Dim msg As Outlook.MailItem
Dim rng As Range
Dim i As Integer, iFolder As Integer, iMsg As Integer
Dim strFldrNm As String, msgAddress As String
Dim oldStatusBar As Boolean

oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True

Set OutApp = CreateObject("Outlook.Application")
Set outNs = OutApp.GetNamespace("MAPI")
Set DictAddress = CreateObject("Scripting.Dictionary")
Set DictName = CreateObject("Scripting.Dictionary")
Set DictDelete = Create_Delete_Email_Dictionary() 'Get existing index of emails for deletion

'Clear Old Data
ThisWorkbook.Sheets(1).Activate
Range("A1", "F" & Range("A1").SpecialCells(xlCellTypeLastCell).Row).ClearContents

For i = 1 To outNs.Folders.Count
    ReDim outAc(1 To outNs.Folders.Count)
    Set outAc(i) = outNs.Folders(i)
    Debug.Print outAc(i).Name
    'loop thru folders
    For iFolder = 1 To outAc(i).Folders.Count
        Set outFldr = outAc(i).Folders(iFolder)
        
        strFldrNm = outFldr.Name
        Debug.Print "    " & strFldrNm & ": Count=" & outFldr.Items.Count
        If strFldrNm = "Inbox" And outAc(i).Name = "hamishmacnamara@yahoo.co.uk" Then
            'outFldr.Sort (Items)
            Set msgs = outFldr.Items
            'Call msgOps1
            GoTo Skip1 ' easier than parsing all local variables to a new sub!
        End If
    Next iFolder
Next i

Skip1:
'recycle i
i = 1
For iMsg = 1 To msgs.Count
    Application.StatusBar = "Msg Count: " & iMsg & " of " & outFldr.Items.Count & " Msgs in " & outFldr.Name
    msgs.Sort ("[SenderEmailAddress]")
    Set msg = msgs(iMsg)
    msgAddress = msg.SenderEmailAddress
    If Not msgAddress = "" And DictDelete(msgAddress) = "" Then 'is a valid mail address AND is not in existing deletion list
        DictAddress(msgAddress) = msg.Sender.Name
        DictName(msg.Sender.Name) = True
        'msgAddress(iMsg) = msg.Sender.Address
        'msgName(iMsg) = msg.Sender.Name
        'msgdetail(iMsg, 2) = msg.Subject
        'ReDim Preserve msDetail(1 To iMsg)
        'Call Email_Item_Ops(Item, msDetail)
        i = i + 1
    End If
    'Exit For
Next iMsg


Set rng = Range("e2")
rng.Resize(i, 1).Value = WorksheetFunction.Transpose(DictAddress.Keys)
Set rng = Range("f2")
rng.Resize(i, 1).Value = WorksheetFunction.Transpose(DictAddress.Items)
Set rng = Range("e2").Resize(i, 2)
rng.Sort key1:=Range("e2"), order1:=xlAscending, Header:=xlNo






End Sub
'
Public Sub msgOps1()
Dim DictDelete As Object
Dim OutApp As Outlook.Application
Dim outNs As Outlook.Namespace ' The Namespace Object (Session) has a collection of accounts.
Dim outFldr As Outlook.MAPIFolder
Dim msgs As Outlook.Items
Dim msg As Outlook.MailItem
Dim strFldrNm As String, msgAddress As String
Dim i As Integer, iFolder As Integer, iMsg As Integer

'Get existing index of emails for deletion
Set DictDelete = Create_Delete_Email_Dictionary()
'k = DictDelete.Keys

Set OutApp = CreateObject("Outlook.Application")
Set outNs = OutApp.GetNamespace("MAPI")

For i = 1 To outNs.Folders.Count
    ReDim outAc(1 To outNs.Folders.Count)
    Set outAc(i) = outNs.Folders(i)
    Debug.Print outAc(i).Name
    'loop thru folders
    For iFolder = 1 To outAc(i).Folders.Count
        Set outFldr = outAc(i).Folders(iFolder)
        
        strFldrNm = outFldr.Name
        Debug.Print "    " & strFldrNm & ": Count=" & outFldr.Items.Count
        If strFldrNm = "Inbox" And outAc(i).Name = "hamishmacnamara@yahoo.co.uk" Then
            'outFldr.Sort (Items)
            Set msgs = outFldr.Items
            'Call msgOps1
            GoTo Skip1 ' easier than parsing all local variables to a new sub!
        End If
    Next iFolder
Next i

Skip1:
'recycle i
i = msgs.Count ' total loops
iMsg = 1
On Error GoTo Crash1
Do While iMsg < i
    Application.StatusBar = "Msg Count: " & iMsg & " of " & outFldr.Items.Count & " Msgs in " & outFldr.Name
    msgs.Sort ("[SenderEmailAddress]")
    Set msg = msgs(iMsg)
    msgAddress = msg.SenderEmailAddress
    If Not msgAddress = "" Then
        If DictDelete.Exists(msg.SenderEmailAddress) And DictDelete.Item(msg.SenderEmailAddress) Then
            Debug.Print "True" & iMsg & ": " & msg.SenderEmailAddress & ", " & iMsg + 1 & msgs(iMsg + 1).SenderEmailAddress
            On Error GoTo Crash1
            msg.Delete
            On Error GoTo Crash2
            iMsg = iMsg - 1
            'i = i - 1 'reduce loops
        Else
            Debug.Print "False" & iMsg & ": " & msg.SenderEmailAddress
        End If
        
        
    End If
    iMsg = iMsg + 1
    'Exit For
Loop
Exit Sub

Crash1:
    i = i - 1
Resume Next

Crash2:
    Debug.Print Err.Description
    Stop
Resume
End Sub

Function Unique2(DRange As Variant) As Variant ' takes 2d range and convertss to unique list of variants
 
Dim Dict As Object
Dim i As Long, j As Long, NumRows As Long, NumCols As Long
 
If k <= 0 Then k = 1

'Convert range to array and count rows and columns
If TypeName(DRange) = "Range" Then DRange = DRange.Value2
NumRows = UBound(DRange)
'NumCols = UBound(DRange, 2)
NumCols = 1
'put unique data elements in a dictionay
Set Dict = CreateObject("Scripting.Dictionary")
For i = 1 To NumCols
    For j = 1 To NumRows
        If NumCols = 1 Then
            Dict(DRange(j)) = 1
        Else
            Dict(DRange(j, i)) = 1
        End If
    Next j
Next i
 
'Dict.Keys() is a Variant array of the unique values in DRange
 'which can be written directly to the spreadsheet
 'but transpose to a column array first
 
Unique2 = WorksheetFunction.Transpose(Dict.Keys)
 
End Function

Public Function Unique(myArray As Variant) As Variant
Dim d As Object
Dim i As Long

'takes a 1d array, converts to unique values and returns as an array

If k <= 0 Then k = 1

Set d = CreateObject("Scripting.Dictionary")
'Set d = New Scripting.Dictionary


For i = LBound(myArray) To UBound(myArray)
    d(myArray(i)) = 1
Next i

'Dim v As Variant
'For Each v In d.Keys()
'    'd.Keys() is a Variant array of the unique values in myArray.
'    'v will iterate through each of them.
'Next

Unique = WorksheetFunction.Transpose(d.Keys)

End Function

Public Function Create_Delete_Email_Dictionary() As Object
'Dim arr As Variant
Dim DictDelete As Object
Dim rng As Range, rngItem As Range
Dim Arr(), k
Dim iRow As Integer, i As Integer
Dim bool As Boolean

'get array
'create dictionary from array
'loop thru messages
'if address in delete list then delete

Set DictDelete = CreateObject("Scripting.Dictionary")
iRow = Range("I1").End(xlDown).Row
Set rng = Range("H2", "I" & iRow)
Set rngItem = Range("I2", "I" & iRow)
Arr = rng
For i = LBound(Arr, 1) To UBound(Arr, 1)
    If Arr(i, 1) = LCase("X") Then
        'Bool = True
        DictDelete(Arr(i, 2)) = True
    Else
        DictDelete(Arr(i, 2)) = False
    End If
'Debug.Print DictDelete(arr(i, 2))


Next i

Set Create_Delete_Email_Dictionary = DictDelete


End Function

