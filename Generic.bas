Attribute VB_Name = "Generic"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Save_Vals_Attachments_Single_email1()
'Run look back for single email address
Dim email_address As String

'email_address=msgbox("Paste email address to run attachgment lookback for",
email_address = "@pamplonafunds.com"
Call Save_Vals_Attachments_Single_email2(email_address)
End Sub

Public Sub Save_Vals_Attachments_Single_email2(olAddress As String)
'Parse mail item
Dim itm As MailItem, olMail As MailItem, arrMail() As MailItem
Dim Fldr() As Folder, flInbox As Variant
Dim myNameSpace As Namespace
Dim myItems As Items
Dim N As Single, iFldr As Single, iItem As Single
Dim tmNow As Date

Set myNameSpace = Application.GetNamespace("MAPI")
'Set flInbox = myNameSpace.Items
Arr = GetXLconfig

''folder map
'Dim Fldr2 As Folder, Fldr3 As Folder
'For Each Fldr2 In myNameSpace.Folders
'    Debug.Print Fldr2.Name
'        For Each Fldr3 In Fldr2.Folders
'            Debug.Print "       " & Fldr3.Name
'        Next Fldr3
'Next Fldr2
'find an item to test
ReDim Fldr(1 To 1) 'Is an array in case we want to search more than one folder in future
iFldr = 1
'Set Fldr(iFldr) = myNameSpace.Folders("hmacnamara@letterone.com").Folders("Inbox")
Set Fldr(iFldr) = myNameSpace.Folders("operations@letterone.com").Folders("Done")
'Does mail meet criteria?
'Fldr(iFldr).Items.Sort "[SENTon]", False


'minitest
Set myItems = Fldr(iFldr).Items
myItems.Sort "[ReceivedTime]", True
tmNow = Now() - 2
For iItem = 1 To myItems.Count
    Debug.Print myItems(iItem).ReceivedTime
    If myItems(iItem).ReceivedTime < tmNow Then Exit For
Next iItem
    
For iItem = iItem To 1 Step -1
'    If n = 500 Then Exit For
    Set olMail = myItems(iItem)
'    Debug.Print arrMail(n).ReceivedTime & " " & n
    'loop thru arr matching against sender and subject
    'Loop Config array (arr) for match with parsed mail item
    N = 2 'first line is headers!
    For N = 2 To UBound(Arr, 1)
        iEndField = UBound(Arr, 2)
        If InStr(1, olMail.SenderEmailAddress, olAddress) > 1 And InStr(olMail.Subject, Arr(N, 3)) <> 0 And Arr(N, iEndField) <> True Then
'            arr(n, iEndField) = True
'            Debug.Print olMail.Sender.Name & "  " & olMail.Subject
            'invoke some save attachment code
            Call saveAttachVals2(olMail, Arr, N)
        End If
        Debug.Print myItems(iItem).ReceivedTime, olMail.Sender.Name & "  " & olMail.Subject
    Next N
Next iItem
End Sub

Public Sub Save_Vals_Attachments()
'Parse mail item
Dim itm As MailItem, olMail As MailItem, arrMail() As MailItem
Dim Fldr() As Folder, flInbox As Variant
Dim myNameSpace As Namespace
Dim myItems As Items
Dim N As Single, iFldr As Single, iItem As Single
Dim tmNow As Date

Set myNameSpace = Application.GetNamespace("MAPI")
'Set flInbox = myNameSpace.Items
Arr = GetXLconfig

''folder map
'Dim Fldr2 As Folder, Fldr3 As Folder
'For Each Fldr2 In myNameSpace.Folders
'    Debug.Print Fldr2.Name
'        For Each Fldr3 In Fldr2.Folders
'            Debug.Print "       " & Fldr3.Name
'        Next Fldr3
'Next Fldr2
'find an item to test
ReDim Fldr(1 To 1) 'Is an array in case we want to search more than one folder in future
iFldr = 1
'Set Fldr(iFldr) = myNameSpace.Folders("hmacnamara@letterone.com").Folders("Inbox")
Set Fldr(iFldr) = myNameSpace.Folders("operations@letterone.com").Folders("Valuations")
'Does mail meet criteria?
'Fldr(iFldr).Items.Sort "[SENTon]", False


'minitest
Set myItems = Fldr(iFldr).Items
myItems.Sort "[ReceivedTime]", True
tmNow = Now() - 2
For iItem = 1 To myItems.Count
    Debug.Print myItems(iItem).ReceivedTime
    If myItems(iItem).ReceivedTime < tmNow Then Exit For
Next iItem
    
For iItem = iItem To 1 Step -1
'    If n = 500 Then Exit For
    Set olMail = myItems(iItem)
'    Debug.Print arrMail(n).ReceivedTime & " " & n
    'loop thru arr matching against sender and subject
    'Loop Config array (arr) for match with parsed mail item
    N = 2 'first line is headers!
    For N = 2 To UBound(Arr, 1)
        iEndField = UBound(Arr, 2)
        If olMail.SenderEmailAddress = Arr(N, 2) And InStr(olMail.Subject, Arr(N, 3)) <> 0 And Arr(N, iEndField) <> True Then
'            arr(n, iEndField) = True
            Debug.Print olMail.Sender.Name & "  " & olMail.Subject
            'invoke some save attachment code
            Call saveAttachVals2(olMail, Arr, N)
        End If
    Next N
Next iItem


''''Fldr(iFldr).Items.Sort "[ReceivedTime]", True
'''''1st arrMail is single dim
''''ReDim arrMail(1 To Fldr(iFldr).Items.Count) As MailItem
''''n = 1
'''''Set Limit opn number of days to look back
''''tmNow = Now() - 7
''''For iItem = Fldr(iFldr).Items.Count To 1 Step -1
''''    Set arrMail(n) = Fldr(iFldr).Items(iItem)
''''    If arrMail(n).ReceivedTime < tmNow Then Exit For
'''''    If n = 500 Then Exit For
''''    n = n + 1
''''Next iItem
'''''Cut size of aray
''''ReDim Preserve arrMail(1 To n)
''''For iItem = 1 To UBound(arrMail)
''''    Set olMail = arrMail(iItem)
'''''    Debug.Print arrMail(n).ReceivedTime & " " & n
''''    'loop thru arr matching against sender and subject
''''    'Loop Config array (arr) for match with parsed mail item
''''    n = 2 'first line is headers!
''''    For n = 2 To UBound(arr, 1)
''''        iEndField = UBound(arr, 2)
''''        If olMail.SenderEmailAddress = arr(n, 2) And InStr(olMail.Subject, arr(n, 3)) <> 0 And arr(n, iEndField) <> True Then
''''            arr(n, iEndField) = True
''''            Debug.Print olMail.Sender.Name & "  " & olMail.Subject
''''            'invoke some save attachment code
''''            Call saveAttachVals2(olMail, arr, n)
''''        End If
''''    Next n
''''Next iItem

End Sub


Public Sub saveAttachVals_Test()
'Parse a group of mail items to backflush those that may have been excluded if mailbox rule breaks or stops
'######## ADD ", Optional arr" into saveAttachVals in order to speed up##########
'Parse mail item
Dim itm As MailItem, olMail As MailItem, arrMail() As MailItem
Dim Fldr(1 To 3) As Folder, flInbox As Variant
Dim myNameSpace As Namespace
Dim myItems As Items
Dim N As Single, iFldr As Single, iItem As Single, iDim As Single
Dim tmLookBack As Date

On Error GoTo Crash:
Arr = GetXLconfig

Set myNameSpace = Application.GetNamespace("MAPI")

Set Fldr(1) = myNameSpace.Folders("operations@letterone.com").Folders("Done")
Set Fldr(2) = myNameSpace.Folders("operations@letterone.com").Folders("Valuations")
Set Fldr(3) = myNameSpace.Folders("operations@letterone.com").Folders("Valuations").Folders("Margin Calls")

'ReDim Fldr(1 To UBound(strFdr)) 'Is an array in case we want to search more than one folder in future
For iFldr = 1 To UBound(Fldr)
'    Set myNameSpace = Application.GetNamespace("MAPI")
    'Set flInbox = myNameSpace.Items
    
    
    'find an item to test
    
'    Set Fldr(iFldr) = strFdr(iFldr)
    
    'minitest
    Set myItems = Fldr(iFldr).Items
    myItems.Sort "[ReceivedTime]", True
    tmLookBack = Now() - 1 '04/10/18 Re run through inbox to populate folders
    ReDim Preserve arrMail(1 To 1)
    iDim = 1
    For iItem = 1 To myItems.Count 'This counts number of emails >lookback time
        If myItems(iItem).ReceivedTime < tmLookBack Then
    '        If myItems(iItem).ReceivedTime < tmLookBack Then
    '            Debug.Print myItems(iItem).ReceivedTime
    '            Set arrMail(iDim) = myItems(iItem)
    '            ReDim Preserve arrMail(1 To UBound(arrMail) + 1)
    '            iDim = iDim + 1
    '        End If
            Exit For
        End If
        
    '    Debug.Print iItem
    '    If iItem >= 1000 Then
    '        Stop
    '        Exit For
    '    End If
    Next iItem
    
    'FYI 6/3 i created arrmail array above to cope with count down but it's not needed due to sort
    For iItem = iItem To 1 Step -1
    '    If n = 500 Then Exit For
        Set olMail = myItems(iItem)
        ' Debug.Print arrMail(n).ReceivedTime & " " & n
        'loop thru arr matching against sender and subject
        'Loop Config array (arr) for match with parsed mail item
        
        N = 2 'first line is headers!
        For N = 2 To UBound(Arr, 1)
            iEndField = UBound(Arr, 2)
            If olMail.SenderEmailAddress = Arr(N, 2) And InStr(olMail.Subject, Arr(N, 3)) <> 0 And Arr(N, iEndField) <> True Then  'INSERT TEST CRITERIA HERE IF CONFIG NOT WORKING
    '            arr(n, iEndField) = True
                Debug.Print olMail.Sender.Name & "  " & olMail.Subject
                'invoke some save attachment code
                Call saveAttachVals(olMail, Arr)
            End If
        Next N
     Next iItem
Next iFldr
'Stop
Exit Sub
Crash:
    Debug.Print Err.Number, ": " & Err.Description
    'Stop
    Resume
End Sub

Public Sub saveAttachVals(olMail As Outlook.MailItem) ', Optional arr)
'Check in config and save attachments per config
Dim myNameSpace As Namespace
Dim Fldr As Folder
Dim N As Single, c As Single

On Error GoTo Crash
If DateValue(Now) - DateValue(olMail.ReceivedTime) > 5 Then
    On Error Resume Next
    sendExceedError (olMail)
'    Debug.Print olMail.Subject; olMail.SenderEmailAddress
    On Error GoTo 0
    GoTo Reached_time_Limit
End If
On Error GoTo Crash
Set myNameSpace = Application.GetNamespace("MAPI")
On Error GoTo CrashArr
If Arr = Empty Then 'this is to accomodate the Test Sub. If run from Test then arr will be populated. Add ", Optional arr" back into arguments to speed up
    Arr = GetXLconfig
End If
ResumeArr:
On Error GoTo Crash
'loop thru arr matching against sender and subject
'Set Fldr = myNameSpace.Folders("Operations").Folders("Valuations")
'Set Fldr = myNameSpace.Folders("hmacnamara@letterone.com").Folders("Inbox")
'Loop Config array (arr) for match with parsed mail item

For N = 2 To UBound(Arr, 1)
    If olMail.SenderEmailAddress = Arr(N, 2) And InStr(olMail.Subject, Arr(N, 3)) <> 0 And olMail.attachments.Count > 0 Then
        Call sendNotice(olMail, N, Arr)
        On Error GoTo Crash
        Debug.Print olMail.ReceivedTime & "  " & olMail.SenderEmailAddress & "  " & olMail.Subject
        Call saveAttachVals2(olMail, Arr, N)
        
    'Workaround for UBS
    'ElseIf InStr(olMail.Subject, arr(n, 3)) <> 0 And InStr(olMail.Subject, "UBS Trade Summary for LETTERONE TREASURY SERVICES LLP S A (USD)") Then
     ElseIf InStr(olMail.Subject, Arr(N, 3)) <> 0 And InStr(olMail.Subject, "UBS Trade Summary for LETTERONE TREASURY SERVICES LLP S A") Then
        Call sendNotice(olMail, N, Arr)
        On Error GoTo Crash
        Debug.Print olMail.ReceivedTime & "  " & olMail.SenderEmailAddress & "  " & olMail.Subject
        Call saveAttachVals2(olMail, Arr, N)
    End If
    
Next N
Exit Sub

CrashArr:
    If Arr(1, 1) <> "" Then GoTo ResumeArr
Crash:
    Debug.Print Err.Number, Err.Description
    Call sendError(olMail, N, Arr, Err.Number, Err.Description)
    Resume
    Exit Sub
Reached_time_Limit:
    Exit Sub
End Sub

Private Sub saveAttachVals2(olMail As MailItem, Arr As Variant, iConfig As Single)

'saves attachments according to parsed mail item parameters and config array 'arr'
Dim objAtt As Outlook.Attachment
Dim dateFormat
If Not Arr(iConfig, 6) = "" Then
        dateFormat = Format(olMail.ReceivedTime, Arr(iConfig, 6))
    Else
        dateFormat = ""
End If

Dim saveFolder As String, strSender As String, strAttach As String
Dim i As Single

'Unzip file logic definitions
    Dim localZipFile As Variant, destFolder As Variant, renameFolder As Variant  'Both must be Variant with late binding of Shell object
    Dim Sh As Object
    Dim File As String, Folder As String, FileShort As String, Path As String, LDate As String, LDateFormat As String, Lrand As Integer, CurrencyString As String

    renameFolder = "S:\Do Not Delete - Unzip Rename Empty Folder\"
    LDate = Date
    LDateFormat = Format(LDate, "DDMMYY")
    Lrand = Int((200 - 100 + 1) * Rnd + 100)
'Unzip file logic definitions end

strSender = Arr(iConfig, 1)

On Error GoTo CrashHere
saveFolder = Arr(iConfig, 4)
For Each objAtt In olMail.attachments
    Select Case Arr(iConfig, 7)
    ' Include L1 Sender Lable Y/N in FileName Saved
    Case "Yes"
        objAtt.SaveAsFile saveFolder & "\" & Arr(iConfig, 1) & dateFormat & Arr(iConfig, 5) & objAtt.DisplayName
        Debug.Print saveFolder & "\" & objAtt.DisplayName
        
        'Unzip files logic
        localZipFile = saveFolder & "\" & Arr(iConfig, 5) & objAtt.DisplayName
        destFolder = saveFolder
        
        If InStr(localZipFile, "EUR") Then
        CurrencyString = "EUR"
        ElseIf InStr(localZipFile, "USD") Then
        CurrencyString = "USD"
        Else
        End If
        
        If Right(localZipFile, 3) = "zip" Then
        Set Sh = CreateObject("Shell.Application")
        With Sh
                .Namespace(renameFolder).CopyHere .Namespace(localZipFile).Items
        End With
        Set Sh = Nothing
        File = Dir(renameFolder & "\*")
        If Right(File, 3) = "lsx" Then
        Path = Right(File, 4)
        FileShort = Left(File, Len(File) - 5)
        Else
        Path = Right(File, 3)
        FileShort = Left(File, Len(File) - 4)
        End If
            Do While File <> ""
                File = Dir(renameFolder & "\*")
                If Right(File, 3) = "lsx" Then
                Path = Right(File, 4)
                FileShort = Left(File, Len(File) - 5)
                Else
                Path = Right(File, 3)
                FileShort = Left(File, Len(File) - 4)
                End If
                Name renameFolder & File As destFolder & "\" & FileShort & "_" & CurrencyString & Lrand & "_" & LDateFormat & "." & Path
                File = Dir
            Loop
        Else
        End If
        'Unzip files logic end
        
    Case Else
        objAtt.SaveAsFile saveFolder & "\" & dateFormat & Arr(iConfig, 5) & objAtt.DisplayName
        Debug.Print saveFolder & "\" & objAtt.DisplayName
        
        'Unzip files logic
            localZipFile = saveFolder & "\" & Arr(iConfig, 5) & objAtt.DisplayName
            destFolder = saveFolder
            
            If InStr(localZipFile, "EUR") Then
            CurrencyString = "EUR"
            ElseIf InStr(localZipFile, "USD") Then
            CurrencyString = "USD"
            Else
            End If
            
            If Right(localZipFile, 3) = "zip" Then
            Set Sh = CreateObject("Shell.Application")
            With Sh
                .Namespace(renameFolder).CopyHere .Namespace(localZipFile).Items
            End With
            Set Sh = Nothing
            File = Dir(renameFolder & "\*")
            If Right(File, 3) = "lsx" Then
            Path = Right(File, 4)
            FileShort = Left(File, Len(File) - 5)
            Else
            Path = Right(File, 3)
            FileShort = Left(File, Len(File) - 4)
            End If
                Do While File <> ""
                    File = Dir(renameFolder & "\*")
                    If Right(File, 3) = "lsx" Then
                    Path = Right(File, 4)
                    FileShort = Left(File, Len(File) - 5)
                    Else
                    Path = Right(File, 3)
                    FileShort = Left(File, Len(File) - 4)
                    End If
                    Name renameFolder & File As destFolder & "\" & FileShort & "_" & CurrencyString & Lrand & "_" & LDateFormat & "." & Path
                    File = Dir
                Loop
            Else
            End If
            'Unzip files logic end
            
    End Select
    Set objAtt = Nothing
    i = i + 1
Next objAtt
On Error GoTo 0
Exit Sub

CrashHere:
    Debug.Print Err.Number, Err.Description
    'MsgBox Err.Description
    'Stop
    Resume Next

End Sub

Public Function GetXLconfig()

Dim objExcelApp As Object
Dim wb As Workbook

Set objExcelApp = CreateObject("Excel.Application")

Set wb = objExcelApp.Workbooks.Open(FileName:="U:\Operations\Operations\Admin\Projects\Code\Operations VBA.xlsm", UpdateLinks:=False, ReadOnly:=True, notify:=False)
Dim Arr()
Dim ws As Worksheet
Dim rng As Range, rng2 As Range
Dim iRow1 As Single, iRow2 As Single

Set ws = wb.Sheets(1)
Set rng = ws.Range("_Vals_Mail_Config") ' dynamic named range
iRow1 = ws.Range("_Vals_Config_Cell").Row
iRow2 = ws.Range("_Vals_Config_Cell").End(xlDown).Row

'define the config range
Set rng2 = ws.Range(iRow1 & ":" & iRow2)
Set rng = objExcelApp.Application.Intersect(rng, rng2)

'commit range to 2d array
Arr = rng
'ws.Cells(1, 1).Value = "Hello"
'ws.Cells(1, 2).Value = "World"

'Close the workbook
wb.Close SaveChanges:=False
Set wb = Nothing
GetXLconfig = Arr
End Function

Sub sendNotice(inMail As MailItem, iConfig As Single, Arr)
    Dim olApp As Outlook.Application
    Dim olSubj As String, strBody As String
    Dim att As Attachment
    Dim N As Single
    
    On Error GoTo Crash2
    olSubj = inMail.SenderEmailAddress
'    olSubj = Right(olSubj, InStr(1, olSubj, "@"))
    Set olApp = Outlook.Application
    Set objMail = olApp.CreateItem(olMailItem)

    objMail.BodyFormat = olFormatPlain
    objMail.Subject = "Outlook SaVE Rule Triggered: " & olSubj
    strBody = olSubj + Chr(10) & inMail.SenderEmailAddress + Chr(10) & "no. Attachments: " & inMail.attachments.Count & Chr(10) & "Dest Path:- " & Arr(iConfig, 4)
    objMail.Body = strBody
    If inMail.attachments.Count > 0 Then
        N = 1
        For Each att In inMail.attachments
            objMail.Body = objMail.Body + Chr(10) & N & att.DisplayName
            N = N + 1
        Next att
    End If
    objMail.To = "hmacnamara@letterone.com"
    objMail.Send
Exit Sub
Crash2:
    Stop
    Resume
End Sub

Sub sendError(inMail As MailItem, N As Single, Arr, sErr As Single, strErr As String)
    Dim olApp As Outlook.Application
    Dim olSubj As String
    Dim olBody As String
On Error GoTo CrashSendError
olSubj = inMail.SenderEmailAddress
'    olSubj = Right(olSubj, InStr(1, olSubj, "@"))
Set olApp = Outlook.Application
Set objMail = olApp.CreateItem(olMailItem)

objMail.BodyFormat = olFormatPlain
objMail.Subject = "Outlook SaVE Rule Error: " & olSubj & " @  " & Now()
objBody = olSubj
objBody = objBody + Chr(10)
objBody = objBody & inMail.Sender
objBody = objBody + Chr(10) + inMail.SenderEmailAddress
objBody = objBody + Chr(10) & "no. Attachments:- "
objBody = objBody & inMail.attachments.Count
objBody = objBody & Chr(10) & "Dest Path:- " & Arr(N, 4)
objBody = objBody + Chr(10) & "Error Number/Description:- "
objBody = objBody & sErr & "/" & strErr
objMail.Body = objBody
objMail.To = "hmacnamara@letterone.com"
objMail.Send
    
Exit Sub
CrashSendError:
    Stop
    Resume Next
End Sub

Sub sendExceedError(inMail As MailItem)
    Dim olApp As Outlook.Application
    Dim olSubj As String
    
    olSubj = inMail.SenderEmailAddress
'    olSubj = Right(olSubj, InStr(1, olSubj, "@"))
    Set olApp = Outlook.Application
    Set objMail = olApp.CreateItem(olMailItem)

    objMail.BodyFormat = olFormatPlain
    objMail.Subject = "Outlook Rule has Exceeded T-5: " & olSubj & " @  " & Now()
    objMail.Body = olSubj + Chr(10) & inMail.Sender + Chr(10) + inMail.SenderEmailAddress + Chr(10) & "no. Attachments:- " & inMail.attachments.Count
    objMail.To = "hmacnamara@letterone.com"
    objMail.Send
End Sub

'UAT ENVIRONMENT//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Sub Save_Vals_Attachments_Single_email1TEST()
'Run look back for single email address
Dim email_address As String

'email_address=msgbox("Paste email address to run attachgment lookback for",
email_address = "@pamplonafunds.com"
Call Save_Vals_Attachments_Single_email2TEST(email_address)
End Sub

Public Sub Save_Vals_Attachments_Single_email2TEST(olAddress As String)
'Parse mail item
Dim itm As MailItem, olMail As MailItem, arrMail() As MailItem
Dim Fldr() As Folder, flInbox As Variant
Dim myNameSpace As Namespace
Dim myItems As Items
Dim N As Single, iFldr As Single, iItem As Single
Dim tmNow As Date

Set myNameSpace = Application.GetNamespace("MAPI")
'Set flInbox = myNameSpace.Items
Arr = GetXLconfigTEST

''folder map
'Dim Fldr2 As Folder, Fldr3 As Folder
'For Each Fldr2 In myNameSpace.Folders
'    Debug.Print Fldr2.Name
'        For Each Fldr3 In Fldr2.Folders
'            Debug.Print "       " & Fldr3.Name
'        Next Fldr3
'Next Fldr2
'find an item to test
ReDim Fldr(1 To 1) 'Is an array in case we want to search more than one folder in future
iFldr = 1
'Set Fldr(iFldr) = myNameSpace.Folders("mhumphreys@letterone.com").Folders("Inbox")
Set Fldr(iFldr) = myNameSpace.Folders("mhumphreys@letterone.com").Folders("Done")
'Does mail meet criteria?
'Fldr(iFldr).Items.Sort "[SENTon]", False


'minitest
Set myItems = Fldr(iFldr).Items
myItems.Sort "[ReceivedTime]", True
tmNow = Now() - 2
For iItem = 1 To myItems.Count
    Debug.Print myItems(iItem).ReceivedTime
    If myItems(iItem).ReceivedTime < tmNow Then Exit For
Next iItem
    
For iItem = iItem To 1 Step -1
'    If n = 500 Then Exit For
    Set olMail = myItems(iItem)
'    Debug.Print arrMail(n).ReceivedTime & " " & n
    'loop thru arr matching against sender and subject
    'Loop Config array (arr) for match with parsed mail item
    N = 2 'first line is headers!
    For N = 2 To UBound(Arr, 1)
        iEndField = UBound(Arr, 2)
        If InStr(1, olMail.SenderEmailAddress, olAddress) > 1 And InStr(olMail.Subject, Arr(N, 3)) <> 0 And Arr(N, iEndField) <> True Then
'            arr(n, iEndField) = True
'            Debug.Print olMail.Sender.Name & "  " & olMail.Subject
            'invoke some save attachment code
            Call saveAttachVals2TEST(olMail, Arr, N)
        End If
        Debug.Print myItems(iItem).ReceivedTime, olMail.Sender.Name & "  " & olMail.Subject
    Next N
Next iItem
End Sub

Public Sub Save_Vals_AttachmentsTEST()
'Parse mail item
Dim itm As MailItem, olMail As MailItem, arrMail() As MailItem
Dim Fldr() As Folder, flInbox As Variant
Dim myNameSpace As Namespace
Dim myItems As Items
Dim N As Single, iFldr As Single, iItem As Single
Dim tmNow As Date

Set myNameSpace = Application.GetNamespace("MAPI")
'Set flInbox = myNameSpace.Items
Arr = GetXLconfigTEST

''folder map
'Dim Fldr2 As Folder, Fldr3 As Folder
'For Each Fldr2 In myNameSpace.Folders
'    Debug.Print Fldr2.Name
'        For Each Fldr3 In Fldr2.Folders
'            Debug.Print "       " & Fldr3.Name
'        Next Fldr3
'Next Fldr2
'find an item to test
ReDim Fldr(1 To 1) 'Is an array in case we want to search more than one folder in future
iFldr = 1
'Set Fldr(iFldr) = myNameSpace.Folders("mhumphreys@letterone.com").Folders("Inbox")
Set Fldr(iFldr) = myNameSpace.Folders("mhumphreys@letterone.com").Folders("Valuations")
'Does mail meet criteria?
'Fldr(iFldr).Items.Sort "[SENTon]", False


'minitest
Set myItems = Fldr(iFldr).Items
myItems.Sort "[ReceivedTime]", True
tmNow = Now() - 2
For iItem = 1 To myItems.Count
    Debug.Print myItems(iItem).ReceivedTime
    If myItems(iItem).ReceivedTime < tmNow Then Exit For
Next iItem
    
For iItem = iItem To 1 Step -1
'    If n = 500 Then Exit For
    Set olMail = myItems(iItem)
'    Debug.Print arrMail(n).ReceivedTime & " " & n
    'loop thru arr matching against sender and subject
    'Loop Config array (arr) for match with parsed mail item
    N = 2 'first line is headers!
    For N = 2 To UBound(Arr, 1)
        iEndField = UBound(Arr, 2)
        If olMail.SenderEmailAddress = Arr(N, 2) And InStr(olMail.Subject, Arr(N, 3)) <> 0 And Arr(N, iEndField) <> True Then
'            arr(n, iEndField) = True
            Debug.Print olMail.Sender.Name & "  " & olMail.Subject
            'invoke some save attachment code
            Call saveAttachVals2TEST(olMail, Arr, N)
        End If
    Next N
Next iItem


''''Fldr(iFldr).Items.Sort "[ReceivedTime]", True
'''''1st arrMail is single dim
''''ReDim arrMail(1 To Fldr(iFldr).Items.Count) As MailItem
''''n = 1
'''''Set Limit opn number of days to look back
''''tmNow = Now() - 7
''''For iItem = Fldr(iFldr).Items.Count To 1 Step -1
''''    Set arrMail(n) = Fldr(iFldr).Items(iItem)
''''    If arrMail(n).ReceivedTime < tmNow Then Exit For
'''''    If n = 500 Then Exit For
''''    n = n + 1
''''Next iItem
'''''Cut size of aray
''''ReDim Preserve arrMail(1 To n)
''''For iItem = 1 To UBound(arrMail)
''''    Set olMail = arrMail(iItem)
'''''    Debug.Print arrMail(n).ReceivedTime & " " & n
''''    'loop thru arr matching against sender and subject
''''    'Loop Config array (arr) for match with parsed mail item
''''    n = 2 'first line is headers!
''''    For n = 2 To UBound(arr, 1)
''''        iEndField = UBound(arr, 2)
''''        If olMail.SenderEmailAddress = arr(n, 2) And InStr(olMail.Subject, arr(n, 3)) <> 0 And arr(n, iEndField) <> True Then
''''            arr(n, iEndField) = True
''''            Debug.Print olMail.Sender.Name & "  " & olMail.Subject
''''            'invoke some save attachment code
''''            Call saveAttachVals2TEST(olMail, arr, n)
''''        End If
''''    Next n
''''Next iItem

End Sub


Public Sub saveAttachVals_TestTEST()
'Parse a group of mail items to backflush those that may have been excluded if mailbox rule breaks or stops
'######## ADD ", Optional arr" into saveAttachValsTEST in order to speed up##########
'Parse mail item
Dim itm As MailItem, olMail As MailItem, arrMail() As MailItem
Dim Fldr(1 To 3) As Folder, flInbox As Variant
Dim myNameSpace As Namespace
Dim myItems As Items
Dim N As Single, iFldr As Single, iItem As Single, iDim As Single
Dim tmLookBack As Date

On Error GoTo CrashTEST:
Arr = GetXLconfigTEST

Set myNameSpace = Application.GetNamespace("MAPI")

Set Fldr(1) = myNameSpace.Folders("mhumphreys@letterone.com").Folders("Done")
Set Fldr(2) = myNameSpace.Folders("mhumphreys@letterone.com").Folders("Valuations")
Set Fldr(3) = myNameSpace.Folders("mhumphreys@letterone.com").Folders("Valuations").Folders("Margin Calls")

'ReDim Fldr(1 To UBound(strFdr)) 'Is an array in case we want to search more than one folder in future
For iFldr = 1 To UBound(Fldr)
'    Set myNameSpace = Application.GetNamespace("MAPI")
    'Set flInbox = myNameSpace.Items
    
    
    'find an item to test
    
'    Set Fldr(iFldr) = strFdr(iFldr)
    
    'minitest
    Set myItems = Fldr(iFldr).Items
    myItems.Sort "[ReceivedTime]", True
    tmLookBack = Now() - 0.5 '#3/5/2018 4:13:00 PM# =====================================SET MAX LOOKBACK
    ReDim Preserve arrMail(1 To 1)
    iDim = 1
    For iItem = 1 To myItems.Count 'This counts number of emails >lookback time
        If myItems(iItem).ReceivedTime < tmLookBack Then
    '        If myItems(iItem).ReceivedTime < tmLookBack Then
    '            Debug.Print myItems(iItem).ReceivedTime
    '            Set arrMail(iDim) = myItems(iItem)
    '            ReDim Preserve arrMail(1 To UBound(arrMail) + 1)
    '            iDim = iDim + 1
    '        End If
            Exit For
        End If
        
    '    Debug.Print iItem
    '    If iItem >= 1000 Then
    '        Stop
    '        Exit For
    '    End If
    Next iItem
    
    'FYI 6/3 i created arrmail array above to cope with count down but it's not needed due to sort
    For iItem = iItem To 1 Step -1
    '    If n = 500 Then Exit For
        Set olMail = myItems(iItem)
        ' Debug.Print arrMail(n).ReceivedTime & " " & n
        'loop thru arr matching against sender and subject
        'Loop Config array (arr) for match with parsed mail item
        
        N = 2 'first line is headers!
        For N = 2 To UBound(Arr, 1)
            iEndField = UBound(Arr, 2)
            If olMail.SenderEmailAddress = Arr(N, 2) And InStr(olMail.Subject, Arr(N, 3)) <> 0 And Arr(N, iEndField) <> True Then  'INSERT TEST CRITERIA HERE IF CONFIG NOT WORKING
    '            arr(n, iEndField) = True
                Debug.Print olMail.Sender.Name & "  " & olMail.Subject
                'invoke some save attachment code
                Call saveAttachValsTEST(olMail, Arr)
            End If
        Next N
     Next iItem
Next iFldr
Stop
Exit Sub
CrashTEST:
    Debug.Print Err.Number, ": " & Err.Description
    Stop
    Resume
End Sub

Public Sub saveAttachValsTEST(olMail As Outlook.MailItem) ', Optional arr)
'Check in config and save attachments per config
Dim myNameSpace As Namespace
Dim Fldr As Folder
Dim N As Single, c As Single

On Error GoTo CrashTEST
If DateValue(Now) - DateValue(olMail.ReceivedTime) > 5 Then
    On Error Resume Next
    sendExceedErrorTEST (olMail)
'    Debug.Print olMail.Subject; olMail.SenderEmailAddress
    On Error GoTo 0
    GoTo Reached_time_Limit
End If

On Error GoTo CrashTEST
Set myNameSpace = Application.GetNamespace("MAPI")
On Error GoTo CrashTESTArr
If Arr = Empty Then 'this is to accomodate the Test Sub. If run from Test then arr will be populated. Add ", Optional arr" back into arguments to speed up
    Arr = GetXLconfigTEST
End If
ResumeArr:
On Error GoTo CrashTEST
'loop thru arr matching against sender and subject
'Set Fldr = myNameSpace.Folders("Operations").Folders("Valuations")
'Set Fldr = myNameSpace.Folders("mhumphreys@letterone.com").Folders("Inbox")
'Loop Config array (arr) for match with parsed mail item

For N = 2 To UBound(Arr, 1)

    If olMail.SenderEmailAddress = Arr(N, 2) And InStr(olMail.Subject, Arr(N, 3)) <> 0 And olMail.attachments.Count > 0 Then
        Call sendNoticeTEST(olMail, N, Arr)
        On Error GoTo CrashTEST
        Debug.Print olMail.ReceivedTime & "  " & olMail.SenderEmailAddress & "  " & olMail.Subject
        Call saveAttachVals2TEST(olMail, Arr, N)
    
    'Workaround for UBS
    ElseIf InStr(olMail.Subject, Arr(N, 3)) <> 0 And InStr(olMail.Subject, "UBS Trade Summary for LETTERONE TREASURY SERVICES LLP S A (USD)") Then
        Call sendNoticeTEST(olMail, N, Arr)
        On Error GoTo CrashTEST
        Debug.Print olMail.ReceivedTime & "  " & olMail.SenderEmailAddress & "  " & olMail.Subject
        Call saveAttachVals2TEST(olMail, Arr, N)
    End If
    
Next N
Exit Sub

CrashTESTArr:
    If Arr(1, 1) <> "" Then GoTo ResumeArr
CrashTEST:
    Debug.Print Err.Number, Err.Description
    Call sendErrorTEST(olMail, N, Arr, Err.Number, Err.Description)
    Resume
    Exit Sub
Reached_time_Limit:
    Exit Sub
End Sub

Private Sub saveAttachVals2TEST(olMail As MailItem, Arr As Variant, iConfig As Single)

'Unzip file logic definitions
    Dim localZipFile As Variant, destFolder As Variant, renameFolder As Variant  'Both must be Variant with late binding of Shell object
    Dim Sh As Object
    Dim File As String, Folder As String, FileShort As String, Path As String, LDate As String, LDateFormat As String, Lrand As Integer, CurrencyString As String

    renameFolder = "U:\Operations\Operations\Admin\Projects\Code\UAT\SFTP\Do Not Delete - Unzip Rename Empty Folder\"
    LDate = Date
    LDateFormat = Format(LDate, "DDMMYY")
    Lrand = Int((200 - 100 + 1) * Rnd + 100)
'Unzip file logic definitions end
    
'saves attachments according to parsed mail item parameters and config array 'arr'
Dim objAtt As Outlook.Attachment
Dim dateFormat
If Not Arr(iConfig, 6) = "" Then
        dateFormat = Format(olMail.ReceivedTime, Arr(iConfig, 6))
    Else
        dateFormat = ""
End If

Dim saveFolder As String, strSender As String, strAttach As String
Dim i As Single

strSender = Arr(iConfig, 1)

On Error GoTo CrashTESTHere
saveFolder = Arr(iConfig, 4)
For Each objAtt In olMail.attachments
    Select Case Arr(iConfig, 7)
    ' Include L1 Sender Lable Y/N in FileName Saved
    Case "Yes"
        objAtt.SaveAsFile saveFolder & "\" & Arr(iConfig, 1) & dateFormat & Arr(iConfig, 5) & objAtt.DisplayName
        Debug.Print saveFolder & "\" & objAtt.DisplayName
        
        'Unzip files logic
        localZipFile = saveFolder & "\" & Arr(iConfig, 5) & objAtt.DisplayName
        destFolder = saveFolder
        
        If InStr(localZipFile, "EUR") Then
        CurrencyString = "EUR"
        ElseIf InStr(localZipFile, "USD") Then
        CurrencyString = "USD"
        Else
        End If
    
        If Right(localZipFile, 3) = "zip" Then
        Set Sh = CreateObject("Shell.Application")
        With Sh
                'Do Until .NameSpace(renameFolder).Items.Count = .NameSpace(localZipFile).Items.Count
                .Namespace(renameFolder).CopyHere .Namespace(localZipFile).Items
                'Sleep 500
                'Loop
        End With
        Set Sh = Nothing
        File = Dir(renameFolder & "\*")
        If Right(File, 3) = "lsx" Then
        Path = Right(File, 4)
        FileShort = Left(File, Len(File) - 5)
        Else
        Path = Right(File, 3)
        FileShort = Left(File, Len(File) - 4)
        End If
              
            Do While File <> ""
                File = Dir(renameFolder & "\*")
                If Right(File, 3) = "lsx" Then
                Path = Right(File, 4)
                FileShort = Left(File, Len(File) - 5)
                Else
                Path = Right(File, 3)
                FileShort = Left(File, Len(File) - 4)
                End If
                Name renameFolder & File As destFolder & "\" & FileShort & "_" & CurrencyString & "_" & Lrand & "_" & LDateFormat & "." & Path
                File = Dir
            Loop
        Else
        End If
        'Unzip files logic end
        
    Case Else
        objAtt.SaveAsFile saveFolder & "\" & dateFormat & Arr(iConfig, 5) & objAtt.DisplayName
        Debug.Print saveFolder & "\" & objAtt.DisplayName
        
         'Unzip files logic
        localZipFile = saveFolder & "\" & Arr(iConfig, 5) & objAtt.DisplayName
        destFolder = saveFolder
        
        If InStr(localZipFile, "EUR") Then
        CurrencyString = "EUR"
        ElseIf InStr(localZipFile, "USD") Then
        CurrencyString = "USD"
        Else
        End If
    
        If Right(localZipFile, 3) = "zip" Then
        Set Sh = CreateObject("Shell.Application")
        With Sh
                'Do Until .NameSpace(renameFolder).Items.Count = .NameSpace(localZipFile).Items.Count
                .Namespace(renameFolder).CopyHere .Namespace(localZipFile).Items
                'Sleep 500
                'Loop
        End With
        Set Sh = Nothing
        File = Dir(renameFolder & "\*")
        If Right(File, 3) = "lsx" Then
        Path = Right(File, 4)
        FileShort = Left(File, Len(File) - 5)
        Else
        Path = Right(File, 3)
        FileShort = Left(File, Len(File) - 4)
        End If
              
            Do While File <> ""
                File = Dir(renameFolder & "\*")
                If Right(File, 3) = "lsx" Then
                Path = Right(File, 4)
                FileShort = Left(File, Len(File) - 5)
                Else
                Path = Right(File, 3)
                FileShort = Left(File, Len(File) - 4)
                End If
                Name renameFolder & File As destFolder & "\" & FileShort & "_" & CurrencyString & "_" & Lrand & "_" & LDateFormat & "." & Path
                File = Dir
            Loop
        Else
        End If
        'Unzip files logic end
        
    End Select
    Set objAtt = Nothing
    i = i + 1
Next objAtt
On Error GoTo 0
Exit Sub

CrashTESTHere:
    Debug.Print Err.Number, Err.Description
    'Stop
    Resume Next
End Sub

Public Function GetXLconfigTEST()

Dim objExcelApp As Object
Dim wb As Workbook

Set objExcelApp = CreateObject("Excel.Application")

Set wb = objExcelApp.Workbooks.Open(FileName:="U:\Operations\Operations\Admin\Projects\Code\UAT\Operations VBA Test.xlsm", UpdateLinks:=False, ReadOnly:=True, notify:=False)
Dim Arr()
Dim ws As Worksheet
Dim rng As Range, rng2 As Range
Dim iRow1 As Single, iRow2 As Single

Set ws = wb.Sheets(1)
Set rng = ws.Range("_Vals_Mail_Config") ' dynamic named range
iRow1 = ws.Range("_Vals_Config_Cell").Row
iRow2 = ws.Range("_Vals_Config_Cell").End(xlDown).Row

'define the config range
Set rng2 = ws.Range(iRow1 & ":" & iRow2)
Set rng = objExcelApp.Application.Intersect(rng, rng2)

'commit range to 2d array
Arr = rng
'ws.Cells(1, 1).Value = "Hello"
'ws.Cells(1, 2).Value = "World"

'Close the workbook
wb.Close SaveChanges:=False
Set wb = Nothing
GetXLconfigTEST = Arr
End Function

Sub sendNoticeTEST(inMail As MailItem, iConfig As Single, Arr)
    Dim olApp As Outlook.Application
    Dim olSubj As String, strBody As String
    Dim att As Attachment
    Dim N As Single
    
    On Error GoTo CrashTEST2
    olSubj = inMail.SenderEmailAddress
'    olSubj = Right(olSubj, InStr(1, olSubj, "@"))
    Set olApp = Outlook.Application
    Set objMail = olApp.CreateItem(olMailItem)

    objMail.BodyFormat = olFormatPlain
    objMail.Subject = "Outlook SaVE Rule Triggered: " & olSubj
    strBody = olSubj + Chr(10) & inMail.SenderEmailAddress + Chr(10) & "no. Attachments: " & inMail.attachments.Count & Chr(10) & "Dest Path:- " & Arr(iConfig, 4)
    objMail.Body = strBody
    If inMail.attachments.Count > 0 Then
        N = 1
        For Each att In inMail.attachments
            objMail.Body = objMail.Body + Chr(10) & N & att.DisplayName
            N = N + 1
        Next att
    End If
    'objMail.To = "mhumphreys@letterone.com"
    'objMail.Send
Exit Sub
CrashTEST2:
    Stop
    Resume
End Sub

Sub sendErrorTEST(inMail As MailItem, N As Single, Arr, sErr As Single, strErr As String)
    Dim olApp As Outlook.Application
    Dim olSubj As String
    Dim olBody As String
On Error GoTo CrashTESTsendErrorTEST
olSubj = inMail.SenderEmailAddress
'    olSubj = Right(olSubj, InStr(1, olSubj, "@"))
Set olApp = Outlook.Application
Set objMail = olApp.CreateItem(olMailItem)

objMail.BodyFormat = olFormatPlain
objMail.Subject = "Outlook SaVE Rule Error: " & olSubj & " @  " & Now()
objBody = olSubj + Chr(10) & inMail.Sender + Chr(10) + inMail.SenderEmailAddress + Chr(10) & "no. Attachments:- " & inMail.attachments.Count & Chr(10) & "Dest Path:- " & Arr(N, 4) _
    + Chr(10) & "Error Number/Description:- " & sErr & "/" & strErr
objMail.Body = objBody
'objMail.To = "mhumphreys@letterone.com"
'objMail.Send
    
CrashTESTsendErrorTEST:
    Stop
    Resume Next
End Sub

Sub sendExceedErrorTEST(inMail As MailItem)
    Dim olApp As Outlook.Application
    Dim olSubj As String
    
    olSubj = inMail.SenderEmailAddress
'    olSubj = Right(olSubj, InStr(1, olSubj, "@"))
    Set olApp = Outlook.Application
    Set objMail = olApp.CreateItem(olMailItem)

    objMail.BodyFormat = olFormatPlain
    objMail.Subject = "Outlook Rule has Exceeded T-5: " & olSubj & " @  " & Now()
    objMail.Body = olSubj + Chr(10) & inMail.Sender + Chr(10) + inMail.SenderEmailAddress + Chr(10) & "no. Attachments:- " & inMail.attachments.Count
    'objMail.To = "mhumphreys@letterone.com"
    'objMail.Send
End Sub

Private m_Folder As Outlook.MAPIFolder
Private m_Find As String
Private m_Wildcard As Boolean

Private Const SpeedUp As Boolean = True
Private Const StopAtFirstMatch As Boolean = True
