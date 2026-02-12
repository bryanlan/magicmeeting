' listEvents.vbs - Lists calendar events within a specified date range
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line arguments
    Dim startDateStr, endDateStr, calendarName, limitStr, offsetStr, compactStr
    Dim subjectContains, attendeeEmail, locationContains
    Dim startDate, endDate, limit, offset, compact

    ' Get and validate arguments
    startDateStr = GetArgument("startDate")
    endDateStr = GetArgument("endDate")
    calendarName = GetArgument("calendar")
    limitStr = GetArgument("limit")
    offsetStr = GetArgument("offset")
    compactStr = GetArgument("compact")
    subjectContains = GetArgument("subjectContains")
    attendeeEmail = GetArgument("attendeeEmail")
    locationContains = GetArgument("locationContains")

    ' Require start date
    RequireArgument "startDate"

    ' Parse dates
    startDate = ParseDate(startDateStr)

    ' If end date is not provided, use start date (single day)
    If endDateStr = "" Then
        endDate = startDate
    Else
        endDate = ParseDate(endDateStr)
    End If

    ' Ensure end date is not before start date
    If endDate < startDate Then
        OutputError "End date cannot be before start date"
        WScript.Quit 1
    End If

    ' Parse limit (default 50, max 200)
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

    ' Parse offset (default 0)
    If offsetStr = "" Then
        offset = 0
    Else
        If Not IsNumeric(offsetStr) Then
            OutputError "Offset must be a number"
            WScript.Quit 1
        End If
        offset = CInt(offsetStr)
        If offset < 0 Then offset = 0
    End If

    ' Parse compact mode (default false)
    compact = (LCase(compactStr) = "true")

    ' Get calendar events with pagination
    Dim result
    result = GetCalendarEventsPaginated(startDate, endDate, calendarName, limit, offset, compact, subjectContains, attendeeEmail, locationContains)

    ' Output events as JSON
    OutputSuccess result
End Sub

' Gets calendar events within the specified date range with pagination
Function GetCalendarEventsPaginated(startDate, endDate, calendarName, limit, offset, compact, subjectContains, attendeeEmail, locationContains)
    On Error Resume Next

    ' Create Outlook objects
    Dim outlookApp, calendar, calItems, restrictedItems, filter, appt
    Dim json, count, totalCount
    Dim subjectMatch, attendeeMatch, locationMatch, i, recipient

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()

    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If

    ' Get calendar items with recurrence support
    ' IMPORTANT: IncludeRecurrences must be set before Sort
    Set calItems = calendar.Items
    calItems.IncludeRecurrences = True
    calItems.Sort "[Start]"

    ' Create filter for date range and apply it
    Dim startStr, endStr
    startStr = Month(startDate) & "/" & Day(startDate) & "/" & Year(startDate) & " 12:00 AM"
    endStr = Month(endDate) & "/" & Day(endDate) & "/" & Year(endDate) & " 11:59 PM"
    filter = "[Start] >= '" & startStr & "' AND [End] <= '" & DateAdd("d", 1, endDate) & "'"
    Set restrictedItems = calItems.Restrict(filter)

    ' Build JSON with pagination info
    json = "{"
    json = json & """startDate"":""" & FormatDate(startDate) & ""","
    json = json & """endDate"":""" & FormatDate(endDate) & ""","
    json = json & """limit"":" & limit & ","
    json = json & """offset"":" & offset & ","
    json = json & """events"":["

    count = 0
    totalCount = 0

    ' Iterate through restricted items
    For Each appt In restrictedItems
        ' Double-check the appointment falls within our date range
        If appt.Start >= startDate And appt.Start < DateAdd("d", 1, endDate) Then
            ' Check subject filter if specified
            subjectMatch = True
            If subjectContains <> "" Then
                subjectMatch = (InStr(1, LCase(appt.Subject), LCase(subjectContains), vbTextCompare) > 0)
            End If

            ' Check attendee filter if specified
            attendeeMatch = True
            If attendeeEmail <> "" And appt.Recipients.Count > 0 Then
                attendeeMatch = False
                For i = 1 To appt.Recipients.Count
                    Set recipient = appt.Recipients.Item(i)
                    If InStr(1, LCase(recipient.Address), LCase(attendeeEmail), vbTextCompare) > 0 Then
                        attendeeMatch = True
                        Exit For
                    End If
                Next
            End If

            ' Check location filter if specified
            locationMatch = True
            If locationContains <> "" Then
                locationMatch = (InStr(1, LCase(appt.Location), LCase(locationContains), vbTextCompare) > 0)
            End If

            ' Only process if matches all filters
            If subjectMatch And attendeeMatch And locationMatch Then
                totalCount = totalCount + 1

                ' Skip items before offset
                If totalCount > offset Then
                    ' Only include up to limit items
                    If count < limit Then
                        If count > 0 Then json = json & ","

                        If compact Then
                            ' Compact format: just subject, start, end, busyStatus
                            json = json & "{"
                            json = json & """subject"":""" & EscapeJSON(appt.Subject) & ""","
                            json = json & """start"":""" & FormatDateTime(appt.Start) & ""","
                            json = json & """end"":""" & FormatDateTime(appt.End) & ""","
                            json = json & """busyStatus"":""" & GetBusyStatusText(appt.BusyStatus) & """"
                            json = json & "}"
                        Else
                            ' Full format using existing function
                            json = json & AppointmentToJSON(appt)
                        End If

                        count = count + 1
                    End If
                End If
            End If
        End If

        ' Safety limit to prevent infinite loops with recurring events
        If totalCount > 5000 Then Exit For
    Next

    json = json & "],"
    json = json & """returned"":" & count & ","
    json = json & """totalInRange"":" & totalCount & ","
    json = json & """hasMore"":" & LCase(CStr(totalCount > offset + limit))
    json = json & "}"

    If Err.Number <> 0 Then
        OutputError "Failed to get calendar events: " & Err.Description
        WScript.Quit 1
    End If

    GetCalendarEventsPaginated = json

    ' Clean up
    Set restrictedItems = Nothing
    Set calItems = Nothing
    Set calendar = Nothing
    Set outlookApp = Nothing
End Function

' Helper function to get busy status text
Function GetBusyStatusText(busyStatus)
    Select Case busyStatus
        Case olBusy
            GetBusyStatusText = "Busy"
        Case olTentative
            GetBusyStatusText = "Tentative"
        Case olFree
            GetBusyStatusText = "Free"
        Case olOutOfOffice
            GetBusyStatusText = "Out of Office"
        Case Else
            GetBusyStatusText = "Unknown"
    End Select
End Function

' Run the main function
Main
