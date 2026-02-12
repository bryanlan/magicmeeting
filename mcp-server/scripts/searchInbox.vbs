' searchInbox.vbs - Searches inbox for emails matching criteria
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Constants
Const olFolderInbox = 6

' Main function
Sub Main()
    ' Get command line arguments
    Dim subjectContains, fromAddresses, toAddresses, receivedAfter, receivedBefore
    Dim bodyContains, folderName
    Dim limitStr, includeBodyStr
    Dim limit, includeBody

    subjectContains = GetArgument("subjectContains")
    fromAddresses = GetArgument("fromAddresses")
    toAddresses = GetArgument("toAddresses")
    receivedAfter = GetArgument("receivedAfter")
    receivedBefore = GetArgument("receivedBefore")
    bodyContains = GetArgument("bodyContains")
    folderName = GetArgument("folder")
    limitStr = GetArgument("limit")
    includeBodyStr = GetArgument("includeBody")

    ' Parse limit (default 50)
    If limitStr = "" Then
        limit = 50
    Else
        If Not IsNumeric(limitStr) Then
            OutputError "Limit must be a number"
            WScript.Quit 1
        End If
        limit = CInt(limitStr)
        If limit < 1 Then limit = 1
        If limit > 200 Then limit = 200
    End If

    ' Parse includeBody (default false, but true if bodyContains is specified)
    If bodyContains <> "" Then
        includeBody = True
    Else
        includeBody = (LCase(includeBodyStr) = "true")
    End If

    ' Search inbox
    Dim result
    result = SearchInbox(subjectContains, fromAddresses, toAddresses, receivedAfter, receivedBefore, bodyContains, folderName, limit, includeBody)

    ' Output result
    OutputSuccess result
End Sub

' Searches inbox for matching emails
Function SearchInbox(subjectContains, fromAddresses, toAddresses, receivedAfter, receivedBefore, bodyContains, folderName, limit, includeBody)
    On Error Resume Next

    Dim outlookApp, namespace, inbox, items, item, folder
    Dim json, count, filter, filterParts
    Dim fromList, toList, fromEmail, i
    Dim afterDate, beforeDate
    Dim bodySearchTerms, bodyMatch, term

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()

    ' Get namespace
    Set namespace = outlookApp.GetNamespace("MAPI")

    ' Get folder (default to Inbox)
    If folderName = "" Or LCase(folderName) = "inbox" Then
        Set folder = namespace.GetDefaultFolder(olFolderInbox)
    ElseIf LCase(folderName) = "sent" Then
        Set folder = namespace.GetDefaultFolder(5) ' olFolderSentMail
    ElseIf LCase(folderName) = "drafts" Then
        Set folder = namespace.GetDefaultFolder(16) ' olFolderDrafts
    Else
        ' Try to find folder by name in Inbox subfolders
        Set folder = namespace.GetDefaultFolder(olFolderInbox)
        On Error Resume Next
        Set folder = folder.Folders(folderName)
        If Err.Number <> 0 Then
            Set folder = namespace.GetDefaultFolder(olFolderInbox)
        End If
        On Error Goto 0
    End If

    If Err.Number <> 0 Then
        OutputError "Failed to access folder: " & Err.Description
        WScript.Quit 1
    End If

    ' Parse body search terms (space-separated for multi-word search)
    If bodyContains <> "" Then
        bodySearchTerms = Split(LCase(bodyContains), " ")
    Else
        bodySearchTerms = Array()
    End If

    ' Get items sorted by received time (newest first)
    Set items = folder.Items
    items.Sort "[ReceivedTime]", True

    ' Build filter if criteria provided
    filterParts = Array()

    ' Subject filter
    If subjectContains <> "" Then
        ReDim Preserve filterParts(UBound(filterParts) + 1)
        filterParts(UBound(filterParts)) = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%" & Replace(subjectContains, "'", "''") & "%'"
    End If

    ' Date filters
    If receivedAfter <> "" Then
        afterDate = ParseDate(receivedAfter)
        ReDim Preserve filterParts(UBound(filterParts) + 1)
        filterParts(UBound(filterParts)) = "[ReceivedTime] >= '" & FormatDate(afterDate) & " 12:00 AM'"
    End If

    If receivedBefore <> "" Then
        beforeDate = ParseDate(receivedBefore)
        ReDim Preserve filterParts(UBound(filterParts) + 1)
        filterParts(UBound(filterParts)) = "[ReceivedTime] <= '" & FormatDate(beforeDate) & " 11:59 PM'"
    End If

    ' Apply filter if any criteria
    If UBound(filterParts) >= 0 Then
        filter = Join(filterParts, " AND ")
        Set items = items.Restrict(filter)
    End If

    ' Parse fromAddresses (semicolon-separated)
    If fromAddresses <> "" Then
        fromList = Split(LCase(fromAddresses), ";")
    Else
        fromList = Array()
    End If

    ' Parse toAddresses (semicolon-separated) - for searching recipients in sent items
    If toAddresses <> "" Then
        toList = Split(LCase(toAddresses), ";")
    Else
        toList = Array()
    End If

    ' Build JSON response
    json = "{"
    json = json & """query"":{"
    If subjectContains <> "" Then json = json & """subjectContains"":""" & EscapeJSON(subjectContains) & ""","
    If fromAddresses <> "" Then json = json & """fromAddresses"":""" & EscapeJSON(fromAddresses) & ""","
    If toAddresses <> "" Then json = json & """toAddresses"":""" & EscapeJSON(toAddresses) & ""","
    If bodyContains <> "" Then json = json & """bodyContains"":""" & EscapeJSON(bodyContains) & ""","
    If folderName <> "" Then json = json & """folder"":""" & EscapeJSON(folderName) & ""","
    If receivedAfter <> "" Then json = json & """receivedAfter"":""" & EscapeJSON(receivedAfter) & ""","
    If receivedBefore <> "" Then json = json & """receivedBefore"":""" & EscapeJSON(receivedBefore) & ""","
    json = json & """limit"":" & limit & ","
    json = json & """includeBody"":" & LCase(CStr(includeBody))
    json = json & "},"
    json = json & """emails"":["

    count = 0

    For Each item In items
        ' Check if we've hit the limit
        If count >= limit Then Exit For

        ' Skip non-mail items
        If TypeName(item) <> "MailItem" Then
            ' Skip
        Else
            ' Check sender filter if specified
            Dim senderMatch
            senderMatch = True

            If UBound(fromList) >= 0 Then
                senderMatch = False
                For i = 0 To UBound(fromList)
                    If Trim(fromList(i)) <> "" Then
                        If InStr(1, LCase(item.SenderEmailAddress), Trim(fromList(i)), vbTextCompare) > 0 Then
                            senderMatch = True
                            Exit For
                        End If
                        If InStr(1, LCase(item.SenderName), Trim(fromList(i)), vbTextCompare) > 0 Then
                            senderMatch = True
                            Exit For
                        End If
                    End If
                Next
            End If

            ' Check recipient filter if specified (for sent items)
            Dim recipientMatch, recip, recipEmail, recipName
            recipientMatch = True

            If senderMatch And UBound(toList) >= 0 Then
                recipientMatch = False
                For Each recip In item.Recipients
                    recipEmail = LCase(recip.Address)
                    recipName = LCase(recip.Name)
                    For i = 0 To UBound(toList)
                        If Trim(toList(i)) <> "" Then
                            If InStr(1, recipEmail, Trim(toList(i)), vbTextCompare) > 0 Then
                                recipientMatch = True
                                Exit For
                            End If
                            If InStr(1, recipName, Trim(toList(i)), vbTextCompare) > 0 Then
                                recipientMatch = True
                                Exit For
                            End If
                        End If
                    Next
                    If recipientMatch Then Exit For
                Next
            End If

            ' Check body filter if specified
            bodyMatch = True
            If senderMatch And recipientMatch And UBound(bodySearchTerms) >= 0 Then
                Dim bodyLower
                bodyLower = LCase(item.Body)

                ' All search terms must be present (AND logic)
                For Each term In bodySearchTerms
                    If Trim(term) <> "" Then
                        If InStr(1, bodyLower, Trim(term), vbTextCompare) = 0 Then
                            bodyMatch = False
                            Exit For
                        End If
                    End If
                Next
            End If

            If senderMatch And recipientMatch And bodyMatch Then
                If count > 0 Then json = json & ","

                ' Get SMTP email address (handle Exchange DN format)
                Dim senderEmail
                senderEmail = item.SenderEmailAddress

                ' If it's an Exchange DN (starts with /O=), try to get SMTP address
                If Left(senderEmail, 3) = "/O=" Then
                    On Error Resume Next
                    If Not item.Sender Is Nothing Then
                        Dim pa
                        Set pa = item.Sender.PropertyAccessor
                        senderEmail = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                        If Err.Number <> 0 Or senderEmail = "" Then
                            senderEmail = item.SenderEmailAddress
                        End If
                    End If
                    On Error Goto 0
                End If

                json = json & "{"
                json = json & """id"":""" & EscapeJSON(item.EntryID) & ""","
                json = json & """subject"":""" & EscapeJSON(item.Subject) & ""","
                json = json & """from"":""" & EscapeJSON(item.SenderName) & ""","
                json = json & """fromEmail"":""" & EscapeJSON(senderEmail) & ""","
                json = json & """received"":""" & FormatDateTime(item.ReceivedTime) & ""","

                ' Add recipients (useful for sent items)
                Dim recipJson, r, rEmail
                recipJson = "["
                Dim recipCount
                recipCount = 0
                For Each r In item.Recipients
                    If recipCount > 0 Then recipJson = recipJson & ","
                    rEmail = r.Address
                    ' Try to get SMTP address for Exchange recipients
                    If Left(rEmail, 3) = "/O=" Then
                        On Error Resume Next
                        rEmail = r.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                        If Err.Number <> 0 Or rEmail = "" Then rEmail = r.Address
                        On Error Goto 0
                    End If
                    recipJson = recipJson & "{""name"":""" & EscapeJSON(r.Name) & """,""email"":""" & EscapeJSON(rEmail) & """}"
                    recipCount = recipCount + 1
                    If recipCount >= 5 Then Exit For ' Limit to first 5 recipients
                Next
                recipJson = recipJson & "]"
                json = json & """to"":" & recipJson

                If includeBody Then
                    json = json & ","
                    json = json & """body"":""" & EscapeJSON(item.Body) & """"
                End If

                json = json & "}"

                count = count + 1
            End If
        End If
    Next

    json = json & "],"
    json = json & """count"":" & count
    json = json & "}"

    If Err.Number <> 0 Then
        OutputError "Failed to search inbox: " & Err.Description
        WScript.Quit 1
    End If

    SearchInbox = json

    ' Clean up
    Set items = Nothing
    Set folder = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
