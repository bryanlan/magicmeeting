' createEvent.vbs - Creates a new calendar event
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Recurrence type constants
Const olRecursDaily = 0
Const olRecursWeekly = 1
Const olRecursMonthly = 2
Const olRecursYearly = 5

' Day of week mask constants (for weekly recurrence)
Const olSunday = 1
Const olMonday = 2
Const olTuesday = 4
Const olWednesday = 8
Const olThursday = 16
Const olFriday = 32
Const olSaturday = 64

' Main function
Sub Main()
    ' Get command line arguments
    Dim subject, startDateStr, startTimeStr, endDateStr, endTimeStr, location, body, isMeeting, attendeesStr, calendarName
    Dim startDateTime, endDateTime, attendees
    Dim roomEmail, isTeamsMeeting
    Dim recurrenceType, recurrenceInterval, recurrenceDays, recurrenceEndDateStr, recurrenceOccurrences

    ' Get and validate arguments
    subject = GetArgument("subject")
    startDateStr = GetArgument("startDate")
    startTimeStr = GetArgument("startTime")
    endDateStr = GetArgument("endDate")
    endTimeStr = GetArgument("endTime")
    location = GetArgument("location")
    isMeeting = LCase(GetArgument("isMeeting")) = "true"

    ' Recurrence arguments
    recurrenceType = LCase(GetArgument("recurrenceType"))
    recurrenceInterval = GetArgument("recurrenceInterval")
    recurrenceDays = LCase(GetArgument("recurrenceDays"))
    recurrenceEndDateStr = GetArgument("recurrenceEndDate")
    recurrenceOccurrences = GetArgument("recurrenceOccurrences")

    ' Read body from temp file if path provided, otherwise from command line
    Dim bodyFilePath
    bodyFilePath = GetArgument("bodyFile")
    If bodyFilePath <> "" Then
        Dim bodyFso, bodyStream
        Set bodyFso = CreateObject("Scripting.FileSystemObject")
        Set bodyStream = bodyFso.OpenTextFile(bodyFilePath, 1)
        body = bodyStream.ReadAll
        bodyStream.Close
        Set bodyStream = Nothing
        Set bodyFso = Nothing
    Else
        body = GetArgument("body")
    End If
    attendeesStr = GetArgument("attendees")
    calendarName = GetArgument("calendar")
    roomEmail = GetArgument("room")
    isTeamsMeeting = LCase(GetArgument("teamsMeeting")) = "true"
    
    ' Require subject and start date/time
    RequireArgument "subject"
    RequireArgument "startDate"
    RequireArgument "startTime"
    
    ' Parse start date/time
    startDateTime = ParseDateTime(startDateStr, startTimeStr)
    
    ' Parse end date/time (if not provided, default to 30 minutes after start)
    If endDateStr = "" Then endDateStr = startDateStr
    If endTimeStr = "" Then
        endDateTime = DateAdd("n", 30, startDateTime)
    Else
        endDateTime = ParseDateTime(endDateStr, endTimeStr)
    End If
    
    ' Ensure end time is not before start time
    If endDateTime <= startDateTime Then
        OutputError "End time cannot be before or equal to start time"
        WScript.Quit 1
    End If
    
    ' Parse attendees (if provided and it's a meeting)
    If isMeeting And attendeesStr <> "" Then
        attendees = Split(attendeesStr, ";")
    Else
        attendees = Array()
    End If
    
    ' Create the event
    Dim eventId, roomName
    eventId = CreateCalendarEvent(subject, startDateTime, endDateTime, location, body, isMeeting, attendees, calendarName, roomEmail, isTeamsMeeting, roomName, recurrenceType, recurrenceInterval, recurrenceDays, recurrenceEndDateStr, recurrenceOccurrences)

    ' Output success with the event ID and room info
    Dim json
    json = "{"
    json = json & """eventId"":""" & eventId & """"
    If roomName <> "" Then
        json = json & ",""room"":""" & EscapeJSON(roomName) & """"
        json = json & ",""roomEmail"":""" & EscapeJSON(roomEmail) & """"
    End If
    json = json & ",""teamsMeeting"":" & LCase(CStr(isTeamsMeeting))
    json = json & "}"
    OutputSuccess json
End Sub

' Parses a date and time string into a DateTime object
Function ParseDateTime(dateStr, timeStr)
    Dim dateObj, timeObj, dateTimeStr
    
    ' Parse date
    dateObj = ParseDate(dateStr)
    
    ' Combine date and time
    dateTimeStr = FormatDate(dateObj) & " " & timeStr
    
    ' Parse combined date/time
    If Not IsDate(dateTimeStr) Then
        OutputError "Invalid time format: " & timeStr
        WScript.Quit 1
    End If
    
    ParseDateTime = CDate(dateTimeStr)
End Function

' Creates a new calendar event with the specified properties
Function CreateCalendarEvent(subject, startDateTime, endDateTime, location, body, isMeeting, attendees, calendarName, roomEmail, isTeamsMeeting, ByRef roomName, recurrenceType, recurrenceInterval, recurrenceDays, recurrenceEndDateStr, recurrenceOccurrences)
    On Error Resume Next

    ' Create Outlook objects
    Dim outlookApp, namespace, calendar, appointment, i, recipient
    Dim recPattern, dayMask

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    Set namespace = outlookApp.GetNamespace("MAPI")

    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If

    ' Create new appointment item
    Set appointment = calendar.Items.Add(olAppointmentItem)

    ' Set appointment properties
    appointment.Subject = subject
    appointment.Start = startDateTime
    appointment.End = endDateTime
    appointment.Body = body

    ' Set up recurrence if specified
    If recurrenceType <> "" And recurrenceType <> "none" Then
        Set recPattern = appointment.GetRecurrencePattern()

        ' Set recurrence type
        Select Case recurrenceType
            Case "daily"
                recPattern.RecurrenceType = olRecursDaily
            Case "weekly"
                recPattern.RecurrenceType = olRecursWeekly
            Case "monthly"
                recPattern.RecurrenceType = olRecursMonthly
            Case "yearly"
                recPattern.RecurrenceType = olRecursYearly
            Case Else
                OutputError "Invalid recurrence type: " & recurrenceType
                WScript.Quit 1
        End Select

        ' Set interval (default to 1)
        If recurrenceInterval <> "" And IsNumeric(recurrenceInterval) Then
            recPattern.Interval = CInt(recurrenceInterval)
        Else
            recPattern.Interval = 1
        End If

        ' Set days of week for weekly recurrence
        If recurrenceType = "weekly" And recurrenceDays <> "" Then
            dayMask = 0
            If InStr(recurrenceDays, "sunday") > 0 Then dayMask = dayMask + olSunday
            If InStr(recurrenceDays, "monday") > 0 Then dayMask = dayMask + olMonday
            If InStr(recurrenceDays, "tuesday") > 0 Then dayMask = dayMask + olTuesday
            If InStr(recurrenceDays, "wednesday") > 0 Then dayMask = dayMask + olWednesday
            If InStr(recurrenceDays, "thursday") > 0 Then dayMask = dayMask + olThursday
            If InStr(recurrenceDays, "friday") > 0 Then dayMask = dayMask + olFriday
            If InStr(recurrenceDays, "saturday") > 0 Then dayMask = dayMask + olSaturday
            If dayMask > 0 Then
                recPattern.DayOfWeekMask = dayMask
            End If
        End If

        ' Set end condition
        If recurrenceEndDateStr <> "" Then
            recPattern.PatternEndDate = ParseDate(recurrenceEndDateStr)
        ElseIf recurrenceOccurrences <> "" And IsNumeric(recurrenceOccurrences) Then
            recPattern.Occurrences = CInt(recurrenceOccurrences)
        Else
            recPattern.NoEndDate = True
        End If

        Set recPattern = Nothing
    End If

    ' Set location - only use location parameter if no room provided
    ' (room resource attendees automatically populate location)
    If roomEmail <> "" Then
        ' Resolve room to get its display name for output only
        Set recipient = namespace.CreateRecipient(roomEmail)
        recipient.Resolve
        If recipient.Resolved Then
            roomName = recipient.Name
        Else
            roomName = roomEmail
        End If
        ' Don't set Location - let Outlook populate it from the room resource
    ElseIf location <> "" Then
        appointment.Location = location
    End If

    ' If it's a meeting, add attendees
    If isMeeting Then
        appointment.MeetingStatus = olMeeting

        ' Add attendees
        For i = LBound(attendees) To UBound(attendees)
            If Trim(attendees(i)) <> "" Then
                Set recipient = appointment.Recipients.Add(Trim(attendees(i)))
                recipient.Type = 1 ' Required attendee (olRequired)
            End If
        Next

        ' Add room as resource attendee if provided
        If roomEmail <> "" Then
            Set recipient = appointment.Recipients.Add(roomEmail)
            recipient.Type = 3 ' Resource attendee (olResource)
        End If

        ' Enable Teams meeting if requested - use UI automation
        If isTeamsMeeting Then
            ' Display the appointment to access the ribbon (don't save first)
            appointment.Display

            ' Give Outlook time to fully render the window
            WScript.Sleep 1000

            ' Use SendKeys to invoke Teams Meeting button via ribbon keyboard shortcuts
            Dim wshShell
            Set wshShell = CreateObject("WScript.Shell")

            ' Activate the Outlook window (title will be "subject - Meeting")
            wshShell.AppActivate subject & " - Meeting"
            WScript.Sleep 300

            ' Try ribbon shortcut: Alt to activate ribbon, then navigate to Teams Meeting
            wshShell.SendKeys "%"  ' Alt to activate ribbon
            WScript.Sleep 200
            wshShell.SendKeys "H"  ' Home tab
            WScript.Sleep 200
            wshShell.SendKeys "TM" ' Teams Meeting
            WScript.Sleep 2500     ' Wait for Teams to generate the link

            ' Use Alt+S to Send (avoids Ctrl+Enter dialog)
            wshShell.SendKeys "%S"
            WScript.Sleep 500

            ' Press Enter to dismiss any confirmation dialogs
            wshShell.SendKeys "{ENTER}"
            WScript.Sleep 300
            wshShell.SendKeys "{ENTER}"
            WScript.Sleep 300

            Set wshShell = Nothing
        Else
            ' Save first to get EntryID, then send
            appointment.Save
            If Err.Number <> 0 Then
                OutputError "Failed to save meeting: " & Err.Description
                WScript.Quit 1
            End If

            ' Verify we have an EntryID before sending
            If appointment.EntryID = "" Then
                OutputError "Meeting was not created (empty EntryID after save)"
                WScript.Quit 1
            End If

            ' Now send the meeting request
            appointment.Send
        End If
    Else
        ' Save the appointment
        appointment.Save
    End If

    If Err.Number <> 0 Then
        OutputError "Failed to create calendar event: " & Err.Description
        WScript.Quit 1
    End If

    ' Return the EntryID as the event ID
    CreateCalendarEvent = appointment.EntryID

    ' Final check - ensure we have a valid ID
    If CreateCalendarEvent = "" Then
        OutputError "Event creation failed (no EntryID returned)"
        WScript.Quit 1
    End If

    ' Clean up
    Set appointment = Nothing
    Set calendar = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
