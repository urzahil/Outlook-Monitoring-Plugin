Attribute VB_Name = "TeamCaseCounter"
Option Explicit

' Team Case Counter for classic Outlook on Windows.
'
' Counts how many cases (email threads) each team member handled in a shared
' mailbox:
'   - several replies by the same person on one thread count as 1 case;
'   - if several people reply on the same thread, each of them gets 1 case.
'
' Run CountTeamCases. The first run asks for your settings and stores them in
' the Windows registry for your user only. Run ResetTeamCaseSettings to change
' them.

Private Const APP_NAME As String = "TeamCaseCounter"
Private Const SECTION As String = "Settings"

' MAPI properties: the person who sent the message, and the mailbox it was
' sent from (they differ for "Send on behalf of" replies).
Private Const PR_SENDER_SMTP As String = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
Private Const PR_SENDER_NAME As String = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
Private Const PR_FROM_SMTP As String = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"
Private Const PR_FROM_NAME As String = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
Private Const PR_FROM_ENTRYID As String = "http://schemas.microsoft.com/mapi/proptag/0x00410102"
Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const PR_CLIENT_SUBMIT_TIME As String = "http://schemas.microsoft.com/mapi/proptag/0x00390040"

Private teamSuffix As String
Private ignored As Object       ' Scripting.Dictionary: addresses to ignore
Private people As Object        ' address -> Dictionary(name, messages, threads)
Private allCases As Object      ' conversation ID -> True
Private unattributed As Long
Private scanned As Long
Private rangeStart As Date
Private rangeEnd As Date
Private rangeEndShown As Date        ' last day inside the period, for display
Private rangeLabel As String
Private onlyNewThreads As Boolean    ' ignore threads started before rangeStart
Private skipOwnThreads As Boolean    ' ignore threads started by the team
Private threadVerdict As Object      ' conversation ID -> 0 ok, 1 too old, 2 ours
Private skippedOld As Long
Private skippedOwn As Long
Private skippedErrors As Long        ' items that couldn't be read (encrypted, protected, corrupt)
Private duplicates As Long
Private seenIds As Object            ' internet message ID -> True
Private excludedFolders As Object    ' EntryID -> True: folders that are never counted
Private itemsDone As Long

Public Sub CountTeamCases()
    Dim store As Outlook.store
    Set store = LoadSettings()
    If store Is Nothing Then Exit Sub
    If Not AskPeriod() Then Exit Sub
    AskThreadOptions

    Set people = CreateObject("Scripting.Dictionary")
    Set allCases = CreateObject("Scripting.Dictionary")
    Set threadVerdict = CreateObject("Scripting.Dictionary")
    unattributed = 0
    scanned = 0
    skippedOld = 0
    skippedOwn = 0
    skippedErrors = 0
    duplicates = 0
    itemsDone = 0
    Set seenIds = CreateObject("Scripting.Dictionary")
    Set excludedFolders = CreateObject("Scripting.Dictionary")
    FindExcludedFolders store

    ScanFolder store.GetRootFolder
    ShowReport store.DisplayName
End Sub

Public Sub ResetTeamCaseSettings()
    On Error Resume Next
    DeleteSetting APP_NAME
    On Error GoTo 0
    MsgBox "Settings cleared. Run CountTeamCases to set them up again.", vbInformation, "Team Case Counter"
End Sub

' ---------------------------------------------------------------- Settings

Private Function LoadSettings() As Outlook.store
    Dim storeName As String, mailbox As String, domain As String, extra As String
    Dim s As Outlook.store, part As Variant

    storeName = GetSetting(APP_NAME, SECTION, "MailboxStore", "")
    If storeName = "" Then
        storeName = PickStore()
        If storeName = "" Then Exit Function
        SaveSetting APP_NAME, SECTION, "MailboxStore", storeName
    End If

    mailbox = GetSetting(APP_NAME, SECTION, "MailboxAddress", "")
    If mailbox = "" Then
        mailbox = Trim$(InputBox("Email address of the shared mailbox" & vbCrLf & _
            "(for example support@yourcompany.com):", "Team Case Counter"))
        If mailbox = "" Then Exit Function
        SaveSetting APP_NAME, SECTION, "MailboxAddress", mailbox
    End If

    domain = GetSetting(APP_NAME, SECTION, "TeamDomain", "")
    If domain = "" Then
        domain = Trim$(InputBox("Your team's email domain" & vbCrLf & _
            "(for example yourcompany.com):", "Team Case Counter"))
        If domain = "" Then Exit Function
        SaveSetting APP_NAME, SECTION, "TeamDomain", domain
    End If

    extra = GetSetting(APP_NAME, SECTION, "IgnoreAddresses", "?")
    If extra = "?" Then
        extra = Trim$(InputBox("Team addresses that aren't people and should be ignored," & vbCrLf & _
            "separated by commas (for example noreply@yourcompany.com)." & vbCrLf & vbCrLf & _
            "Leave empty if there are none.", "Team Case Counter"))
        SaveSetting APP_NAME, SECTION, "IgnoreAddresses", extra
    End If

    teamSuffix = "@" & LCase$(Replace(domain, "@", ""))
    Set ignored = CreateObject("Scripting.Dictionary")
    ignored(LCase$(mailbox)) = True
    For Each part In Split(extra, ",")
        If Trim$(part) <> "" Then ignored(LCase$(Trim$(part))) = True
    Next

    For Each s In Application.Session.Stores
        If s.DisplayName = storeName Then
            Set LoadSettings = s
            Exit Function
        End If
    Next
    MsgBox "Can't find the mailbox """ & storeName & """ in Outlook." & vbCrLf & _
        "Run ResetTeamCaseSettings and pick it again.", vbExclamation, "Team Case Counter"
End Function

Private Function PickStore() As String
    Dim s As Outlook.store, names() As String, count As Long, i As Long
    Dim prompt As String, answer As String

    ReDim names(1 To Application.Session.Stores.count)
    For Each s In Application.Session.Stores
        count = count + 1
        names(count) = s.DisplayName
        prompt = prompt & count & ".  " & s.DisplayName & vbCrLf
    Next

    answer = Trim$(InputBox("Which mailbox is the shared support mailbox?" & vbCrLf & vbCrLf & _
        prompt & vbCrLf & "Type its number:", "Team Case Counter"))
    If answer = "" Then Exit Function
    If Not IsNumeric(answer) Then GoTo Invalid
    i = CLng(answer)
    If i < 1 Or i > count Then GoTo Invalid
    PickStore = names(i)
    Exit Function

Invalid:
    MsgBox "Please type one of the numbers shown.", vbExclamation, "Team Case Counter"
End Function

Private Function AskPeriod() As Boolean
    Dim answer As String, y As Integer, m As Integer

    answer = Trim$(InputBox("Which period?" & vbCrLf & vbCrLf & _
        "7  = last 7 days" & vbCrLf & _
        "30 = last 30 days" & vbCrLf & _
        "or a calendar month as year-month, for example " & Format(Date, "yyyy-mm"), _
        "Team Case Counter", "7"))

    Select Case answer
        Case ""
            Exit Function
        Case "7", "30"
            rangeEnd = Now
            rangeStart = rangeEnd - CLng(answer)
            rangeEndShown = rangeEnd
            rangeLabel = "Last " & answer & " days"
        Case Else
            If Not answer Like "####-##" Then GoTo Invalid
            y = CInt(Left$(answer, 4))
            m = CInt(Mid$(answer, 6, 2))
            If m < 1 Or m > 12 Then GoTo Invalid
            rangeStart = DateSerial(y, m, 1)
            rangeEnd = DateSerial(y, m + 1, 1)
            If rangeStart > Now Then GoTo Invalid
            rangeEndShown = rangeEnd - 1 / 86400   ' the last second of the month
            If rangeEnd > Now Then
                rangeEnd = Now
                rangeEndShown = Now
            End If
            rangeLabel = Format(rangeStart, "mmmm yyyy")
    End Select
    AskPeriod = True
    Exit Function

Invalid:
    MsgBox "Type 7, 30, or a past month like " & Format(Date, "yyyy-mm") & ".", vbExclamation, "Team Case Counter"
End Function

Private Sub AskThreadOptions()
    onlyNewThreads = (MsgBox("Only count threads that were started inside this period?" & vbCrLf & vbCrLf & _
        "Yes = replies to threads started before " & Format(rangeStart, "d mmm yyyy") & " are not counted." & vbCrLf & _
        "No  = every reply in the period is counted.", vbYesNo + vbQuestion, "Team Case Counter") = vbYes)
    skipOwnThreads = (MsgBox("Ignore threads started by your own team?" & vbCrLf & vbCrLf & _
        "Yes = only threads started by someone outside " & teamSuffix & " are counted." & vbCrLf & _
        "No  = threads started by the team are counted too.", vbYesNo + vbQuestion, "Team Case Counter") = vbYes)
End Sub

' ---------------------------------------------------------------- Counting

' Date as text for a DASL filter: fixed layout, UTC, no locale-dependent
' separators or AM/PM (which made Restrict fail silently on non-US systems).
Private Function DaslDate(folder As Outlook.folder, d As Date) As String
    Dim u As Date
    u = folder.PropertyAccessor.LocalTimeToUTC(d)
    DaslDate = Year(u) & "-" & Two(Month(u)) & "-" & Two(Day(u)) & " " & _
        Two(Hour(u)) & ":" & Two(Minute(u)) & ":" & Two(Second(u))
End Function

Private Function Two(ByVal n As Long) As String
    Two = Right$("0" & n, 2)
End Function

' Remembers the folders whose mail must not be counted: deleted mistakes, spoofed
' mail in Junk, and unsent or failed items.
Private Sub FindExcludedFolders(store As Outlook.store)
    Dim kind As Variant, f As Outlook.folder
    ' Deleted Items, Outbox, Drafts, Junk, Conflicts, Local Failures, Server Failures, Sync Issues
    For Each kind In Array(3, 4, 16, 23, 19, 21, 22, 20)
        Set f = Nothing
        On Error Resume Next
        Set f = store.GetDefaultFolder(kind)
        On Error GoTo 0
        If Not f Is Nothing Then excludedFolders(f.EntryID) = True
    Next
End Sub

Private Function IsExcludedFolder(folder As Outlook.folder) As Boolean
    If excludedFolders.Exists(folder.EntryID) Then IsExcludedFolder = True: Exit Function
    IsExcludedFolder = (folder.Name = "Sync Issues")
End Function

Private Sub ScanFolder(folder As Outlook.folder)
    Dim items As Outlook.items, item As Object, subFolder As Outlook.folder
    Dim i As Long, query As String

    If IsExcludedFolder(folder) Then Exit Sub
    On Error GoTo FolderFailed

    If folder.DefaultItemType = olMailItem Then
        ' Some system folders can't be filtered; skip those.
        On Error Resume Next
        query = "@SQL=""" & PR_CLIENT_SUBMIT_TIME & """ >= '" & DaslDate(folder, rangeStart) & "' AND " & _
            """" & PR_CLIENT_SUBMIT_TIME & """ < '" & DaslDate(folder, rangeEnd) & "'"
        If Err.Number = 0 Then Set items = folder.items.Restrict(query)
        Err.Clear
        On Error GoTo 0
        If Not items Is Nothing Then
            For i = 1 To items.count
                ' Outlook allows only about 250 open items per session: release each one.
                Set item = Nothing
                On Error Resume Next
                Set item = items.item(i)
                On Error GoTo 0
                If Not item Is Nothing Then
                    If TypeOf item Is Outlook.MailItem Then CountMessage item
                End If
                Set item = Nothing
                itemsDone = itemsDone + 1
                If itemsDone Mod 100 = 0 Then DoEvents
            Next
        End If
        Set items = Nothing
    End If

    ' Skip subfolders Outlook can't open rather than stopping the whole count.
    For Each subFolder In folder.Folders
        ScanFolder subFolder
        Set subFolder = Nothing
    Next
    Exit Sub

FolderFailed:
    skippedErrors = skippedErrors + 1
    Err.Clear
End Sub

' One unreadable item (S/MIME-encrypted, IRM-protected, corrupt) must not stop the scan.
Private Sub CountMessage(ByVal mail As Outlook.MailItem)
    On Error GoTo Skip
    CountMessageCore mail
    Exit Sub
Skip:
    skippedErrors = skippedErrors + 1
    Err.Clear
End Sub

Private Sub CountMessageCore(ByVal mail As Outlook.MailItem)
    Dim senderAddr As String, fromAddr As String
    Dim who As String, whoName As String, conv As String, msgId As String
    Dim person As Object, threads As Object, thread As Variant

    If Not mail.Sent Then Exit Sub   ' drafts
    conv = mail.ConversationID
    If conv = "" Then Exit Sub

    ' The same message can be filed in several folders (rules, CC to the shared mailbox).
    msgId = GetProp(mail, PR_INTERNET_MESSAGE_ID)
    If msgId <> "" Then
        If seenIds.Exists(msgId) Then
            duplicates = duplicates + 1
            Exit Sub
        End If
        seenIds(msgId) = True
    End If
    scanned = scanned + 1

    If ThreadVerdictFor(mail, conv) <> 0 Then Exit Sub

    senderAddr = SenderAddress(mail)
    fromAddr = FromAddress(mail)

    ' The sender is checked first so "Send on behalf of" replies (sent from the
    ' shared mailbox by a person) are credited to the person.
    If IsTeamPerson(senderAddr) Then
        who = senderAddr
        whoName = GetProp(mail, PR_SENDER_NAME)
    ElseIf IsTeamPerson(fromAddr) Then
        who = fromAddr
        whoName = GetProp(mail, PR_FROM_NAME)
    Else
        ' Sent "as" the shared mailbox: nobody can be credited.
        If IsTeam(senderAddr) Or IsTeam(fromAddr) Then unattributed = unattributed + 1
        Exit Sub
    End If
    If whoName = "" Then whoName = who

    If Not people.Exists(who) Then
        Set person = CreateObject("Scripting.Dictionary")
        person("name") = whoName
        person("messages") = 0
        Set person("threads") = CreateObject("Scripting.Dictionary")
        people.Add who, person
    End If
    Set person = people(who)
    person("messages") = person("messages") + 1

    Set threads = person("threads")
    If threads.Exists(conv) Then
        thread = threads(conv)
        thread(1) = thread(1) + 1
        threads(conv) = thread
    Else
        threads.Add conv, Array(SubjectOf(mail), 1)
    End If
    allCases(conv) = True
End Sub

' Decides once per conversation whether it is counted: 0 = yes, 1 = started
' before the period, 2 = started by the team. The first message of the thread
' is found with Outlook's conversation API, so it works across folders and even
' when the start is older than the period. If it can't be found the thread is
' counted rather than silently dropped.
Private Function ThreadVerdictFor(ByVal mail As Outlook.MailItem, conv As String) As Long
    Dim verdict As Long, root As Object, best As Object
    Dim roots As Outlook.SimpleItems, i As Long, started As Date, bestDate As Date

    If Not (onlyNewThreads Or skipOwnThreads) Then Exit Function
    If threadVerdict.Exists(conv) Then
        ThreadVerdictFor = threadVerdict(conv)
        Exit Function
    End If

    On Error Resume Next
    Set roots = mail.GetConversation.GetRootItems
    If Not roots Is Nothing Then
        For i = 1 To roots.count
            Set root = roots.Item(i)
            If TypeOf root Is Outlook.MailItem Then
                started = root.SentOn
                If best Is Nothing Or started < bestDate Then
                    Set best = root
                    bestDate = started
                End If
            End If
            Set root = Nothing   ' don't keep hundreds of items open
        Next
    End If
    Set roots = Nothing
    On Error GoTo 0

    If Not best Is Nothing Then
        If onlyNewThreads And bestDate < rangeStart Then
            verdict = 1
        ElseIf skipOwnThreads Then
            If IsTeam(SenderAddress(best)) Or IsTeam(FromAddress(best)) Then verdict = 2
        End If
    End If

    If verdict = 1 Then skippedOld = skippedOld + 1
    If verdict = 2 Then skippedOwn = skippedOwn + 1
    Set best = Nothing
    threadVerdict(conv) = verdict
    ThreadVerdictFor = verdict
End Function

Private Function SenderAddress(ByVal mail As Outlook.MailItem) As String
    Dim addr As String
    addr = GetProp(mail, PR_SENDER_SMTP)
    If addr = "" Then
        On Error Resume Next
        If mail.SenderEmailType = "EX" Then
            addr = mail.Sender.GetExchangeUser().PrimarySmtpAddress
        Else
            addr = mail.SenderEmailAddress
        End If
        On Error GoTo 0
    End If
    SenderAddress = LCase$(Trim$(addr))
End Function

' The mailbox a message was sent from. Internal Exchange senders have no SMTP
' property, so fall back to resolving the "sent representing" address entry.
Private Function FromAddress(ByVal mail As Outlook.MailItem) As String
    Dim addr As String, entry As Outlook.AddressEntry, id As String
    addr = GetProp(mail, PR_FROM_SMTP)
    If addr = "" Then
        On Error Resume Next
        id = mail.PropertyAccessor.BinaryToString(mail.PropertyAccessor.GetProperty(PR_FROM_ENTRYID))
        If id <> "" Then
            Set entry = Application.Session.GetAddressEntryFromID(id)
            If Not entry Is Nothing Then
                If entry.AddressEntryUserType = olExchangeUserAddressEntry Or entry.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
                    addr = entry.GetExchangeUser().PrimarySmtpAddress
                Else
                    addr = entry.Address
                End If
            End If
        End If
        On Error GoTo 0
    End If
    FromAddress = LCase$(Trim$(addr))
End Function

Private Function GetProp(ByVal mail As Outlook.MailItem, tag As String) As String
    On Error Resume Next
    GetProp = mail.PropertyAccessor.GetProperty(tag)
End Function

Private Function SubjectOf(ByVal mail As Outlook.MailItem) As String
    SubjectOf = mail.ConversationTopic
    If SubjectOf = "" Then SubjectOf = mail.Subject
    If SubjectOf = "" Then SubjectOf = "(no subject)"
End Function

Private Function IsTeam(addr As String) As Boolean
    IsTeam = Len(addr) > Len(teamSuffix) And Right$(addr, Len(teamSuffix)) = teamSuffix
End Function

Private Function IsTeamPerson(addr As String) As Boolean
    IsTeamPerson = IsTeam(addr) And Not ignored.Exists(addr)
End Function

' ---------------------------------------------------------------- Report

Private Sub ShowReport(mailboxName As String)
    Dim order As Variant, i As Long, j As Long, tmp As Variant
    Dim person As Object, threads As Object, conv As Variant, thread As Variant
    Dim html As String, csv As String, folderPath As String, stamp As String, csvName As String

    ' Sort people by number of cases, highest first.
    order = people.keys
    For i = 0 To people.count - 2
        For j = i + 1 To people.count - 1
            If people(order(j))("threads").count > people(order(i))("threads").count Then
                tmp = order(i): order(i) = order(j): order(j) = tmp
            End If
        Next
    Next

    ' A new file name every run: Excel locks an open CSV, so overwriting it would fail.
    stamp = Format(Now, "yyyymmdd") & "-" & Format(Now, "hhnnss")
    csvName = "TeamCaseCounter-" & stamp & ".csv"

    html = "<!doctype html><html><head><meta charset=""utf-8""><title>Team cases - " & Esc(rangeLabel) & "</title>" & _
        "<style>body{font:14px/1.4 'Segoe UI',sans-serif;color:#242424;margin:24px;max-width:760px}" & _
        "table{border-collapse:collapse;width:100%}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid #ddd;vertical-align:top}" & _
        ".num{text-align:right}.muted{color:#616161;font-size:12px}.warn{background:#fff4ce;padding:8px;border-radius:4px}" & _
        "ul{margin:4px 0 0;padding-left:18px}</style></head><body>" & _
        "<h1>Team cases: " & Esc(rangeLabel) & "</h1>" & _
        "<p>" & Esc(mailboxName) & " &middot; " & Format(rangeStart, "d mmm yyyy") & " to " & Format(rangeEndShown, "d mmm yyyy") & _
        " &middot; " & allCases.count & " distinct cases handled &middot; " & scanned & " messages scanned</p>"

    If onlyNewThreads Or skipOwnThreads Then
        html = html & "<p class=""muted"">Filters:" & _
            IIf(onlyNewThreads, " only threads started in this period (" & skippedOld & " older " & IIf(skippedOld = 1, "thread", "threads") & " skipped).", "") & _
            IIf(skipOwnThreads, " threads started by the team ignored (" & skippedOwn & " skipped).", "") & "</p>"
    End If

    If duplicates > 0 Or skippedErrors > 0 Then
        html = html & "<p class=""muted"">" & _
            IIf(duplicates > 0, duplicates & " duplicate " & IIf(duplicates = 1, "copy", "copies") & " ignored. ", "") & _
            IIf(skippedErrors > 0, skippedErrors & " " & IIf(skippedErrors = 1, "message", "messages") & " couldn't be read (encrypted or protected) and " & IIf(skippedErrors = 1, "was", "were") & " skipped.", "") & "</p>"
    End If

    If unattributed > 0 Then
        html = html & "<p class=""warn"">" & unattributed & " team " & IIf(unattributed = 1, "reply was", "replies were") & _
            " sent ""as"" the shared mailbox, so the sender can't be identified and " & _
            IIf(unattributed = 1, "it is", "they are") & " not counted. " & _
            "Use ""Send on behalf"" or reply from personal addresses to track them.</p>"
    End If

    html = html & "<table><tr><th>Person</th><th class=""num"">Cases</th><th class=""num"">Messages</th></tr>"
    csv = """Name"",""Email"",""Cases"",""Messages sent""" & vbCrLf

    If people.count = 0 Then
        html = html & "<tr><td colspan=""3"">No replies from the team in this period.</td></tr>"
    End If
    For i = 0 To people.count - 1
        Set person = people(order(i))
        Set threads = person("threads")
        html = html & "<tr><td><details><summary>" & Esc(person("name")) & "</summary>" & _
            "<div class=""muted"">" & Esc(order(i)) & "</div><ul class=""muted"">"
        For Each conv In threads.keys
            thread = threads(conv)
            html = html & "<li>" & Esc(thread(0))
            If thread(1) > 1 Then html = html & " (" & thread(1) & " replies)"
            html = html & "</li>"
        Next
        html = html & "</ul></details></td><td class=""num"">" & threads.count & "</td>" & _
            "<td class=""num"">" & person("messages") & "</td></tr>"
        csv = csv & CsvField(person("name")) & "," & CsvField(order(i)) & "," & threads.count & "," & person("messages") & vbCrLf
    Next
    html = html & "</table><p><a href=""" & csvName & """ download=""" & csvName & """>Open as a spreadsheet (CSV)</a></p></body></html>"

    folderPath = Environ$("TEMP") & "\"
    WriteUtf8 folderPath & "TeamCaseCounter-" & stamp & ".html", html
    WriteUtf8 folderPath & csvName, csv
    CreateObject("Shell.Application").ShellExecute folderPath & "TeamCaseCounter-" & stamp & ".html"
End Sub

Private Function Esc(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    Esc = Replace(text, """", "&quot;")
End Function

' Quotes a CSV text cell. A leading = + - @ (or tab / carriage return) would be
' run as a formula by Excel, and display names come from email headers.
Private Function CsvField(ByVal text As String) As String
    Dim first As String
    first = Left$(text, 1)
    If first = "=" Or first = "+" Or first = "-" Or first = "@" Or first = vbTab Or first = vbCr Or first = vbLf Then text = "'" & text
    CsvField = """" & Replace(text, """", """""") & """"
End Function

Private Sub WriteUtf8(path As String, text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2          ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile path, 2   ' overwrite
    stream.Close
End Sub
