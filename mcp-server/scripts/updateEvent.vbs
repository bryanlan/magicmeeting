' updateEvent.vbs - Updates an existing calendar event
Option Explicit

' Outlook recurrence state constants (not in utils.vbs)
Const olApptNotRecurring = 0
Const olApptMaster = 1
Const olApptOccurrence = 2
Const olApptException = 3

' Include utility functions (defines olMeeting and other constants)
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line arguments
    Dim eventId, subject, startDateStr, startTimeStr, endDateStr, endTimeStr, location, body, calendarName
    Dim originalStartStr, sendUpdate
    Dim startDateTime, endDateTime, originalStart

    ' Get and validate arguments
    eventId = GetArgument("eventId")
    subject = GetArgument("subject")
    startDateStr = GetArgument("startDate")
    startTimeStr = GetArgument("startTime")
    endDateStr = GetArgument("endDate")
    endTimeStr = GetArgument("endTime")
    location = GetArgument("location")
    calendarName = GetArgument("calendar")
    originalStartStr = GetArgument("originalStart")  ' For recurring: original occurrence date/time
    sendUpdate = GetArgument("sendUpdate")           ' Whether to send meeting update to attendees
    Dim updateSeriesStr
    updateSeriesStr = GetArgument("updateSeries")    ' For recurring: update entire series

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

    ' Require event ID
    RequireArgument "eventId"

    ' Parse original start date/time for recurring meetings
    ' Format: "MM/DD/YYYY HH:MM AM/PM" (must match occurrence's original start exactly)
    If originalStartStr <> "" Then
        If Not IsDate(originalStartStr) Then
            OutputError "Invalid originalStart format. Use: MM/DD/YYYY HH:MM AM/PM"
            WScript.Quit 1
        End If
        originalStart = CDate(originalStartStr)
    End If

    ' Parse date/time if provided
    If startDateStr <> "" And startTimeStr <> "" Then
        startDateTime = ParseDateTime(startDateStr, startTimeStr)
    ElseIf startDateStr <> "" And startTimeStr = "" Then
        ' Handle date-only update - will need existing time from event
        startDateTime = "DATE_ONLY:" & startDateStr
    ElseIf startDateStr = "" And startTimeStr <> "" Then
        ' Handle time-only update - will need existing date from event
        startDateTime = "TIME_ONLY:" & startTimeStr
    End If

    If endDateStr <> "" And endTimeStr <> "" Then
        endDateTime = ParseDateTime(endDateStr, endTimeStr)
    ElseIf endDateStr <> "" And endTimeStr = "" Then
        ' Handle date-only update - will need existing time from event
        endDateTime = "DATE_ONLY:" & endDateStr
    ElseIf endDateStr = "" And endTimeStr <> "" Then
        ' Handle time-only update - will need existing date from event
        endDateTime = "TIME_ONLY:" & endTimeStr
    End If

    ' Ensure end time is not before start time if both are provided
    ' Skip validation for date-only or time-only updates since we need the existing appointment to calculate final times
    If Not IsEmpty(startDateTime) And Not IsEmpty(endDateTime) Then
        Dim startPrefix, endPrefix
        startPrefix = Left(CStr(startDateTime), 10)
        endPrefix = Left(CStr(endDateTime), 10)
        If startPrefix <> "DATE_ONLY:" And startPrefix <> "TIME_ONLY:" And endPrefix <> "DATE_ONLY:" And endPrefix <> "TIME_ONLY:" Then
            If endDateTime <= startDateTime Then
                OutputError "End time cannot be before or equal to start time"
                WScript.Quit 1
            End If
        End If
    End If

    ' Update the event
    Dim result, updateSeries
    updateSeries = (LCase(updateSeriesStr) = "true")
    result = UpdateCalendarEvent(eventId, subject, startDateTime, endDateTime, location, body, calendarName, originalStart, (LCase(sendUpdate) = "true"), updateSeries)

    ' Output success
    OutputSuccess "{""success"":" & LCase(CStr(result)) & "}"
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

' Parses a time-only string (for use with existing appointment date)
Function ParseTimeOnly(timeStr)
    Dim dt
    dt = "1/1/2000 " & timeStr
    If Not IsDate(dt) Then
        OutputError "Invalid time format: " & timeStr
        WScript.Quit 1
    End If
    ParseTimeOnly = TimeValue(CDate(dt))
End Function

' Updates an existing calendar event with the specified properties
' For recurring meetings, pass originalStart to identify which occurrence to modify
' Or set updateSeries=True to update the entire series (all occurrences)
Function UpdateCalendarEvent(eventId, subject, startDateTime, endDateTime, location, body, calendarName, originalStart, sendUpdate, updateSeries)
    On Error Resume Next

    ' Create Outlook objects
    Dim outlookApp, namespace, calendar, appointment, targetAppt
    Dim recPattern, recState

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()

    ' Get MAPI namespace
    Set namespace = outlookApp.GetNamespace("MAPI")

    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If

    ' Try to get the appointment by EntryID
    Set appointment = namespace.GetItemFromID(eventId)

    ' Check if appointment was found
    If appointment Is Nothing Then
        OutputError "Event not found with ID: " & eventId
        UpdateCalendarEvent = False
        Exit Function
    End If

    ' Determine the target appointment (handle recurring vs non-recurring)
    recState = appointment.RecurrenceState

    If recState = olApptMaster Then
        ' This is a recurring series master
        If updateSeries Then
            ' Update the entire series - use the master appointment directly
            Set targetAppt = appointment
        ElseIf Not IsEmpty(originalStart) Then
            ' Get the recurrence pattern and fetch the specific occurrence
            Set recPattern = appointment.GetRecurrencePattern()
            Err.Clear
            Set targetAppt = recPattern.GetOccurrence(originalStart)

            If Err.Number <> 0 Or targetAppt Is Nothing Then
                OutputError "Could not find occurrence at " & originalStart & ". Make sure originalStart matches the occurrence's ORIGINAL start date+time exactly. Error: " & Err.Description
                UpdateCalendarEvent = False
                Exit Function
            End If
            Err.Clear
        Else
            OutputError "This is a recurring meeting series. Either set updateSeries=true to update ALL occurrences, or provide originalStart parameter to modify a specific instance (format: MM/DD/YYYY HH:MM AM/PM)."
            UpdateCalendarEvent = False
            Exit Function
        End If
    ElseIf recState = olApptOccurrence Or recState = olApptException Then
        ' Already have a specific occurrence or exception - use it directly
        Set targetAppt = appointment
    Else
        ' Non-recurring appointment - use directly
        Set targetAppt = appointment
    End If

    ' Capture original values and duration
    Dim origStart, origEnd, origSubject, origLocation, origBody
    Dim durMins, newStart, newEnd, startChanged, endExplicit, updated

    origStart = targetAppt.Start
    origEnd = targetAppt.End
    origSubject = targetAppt.Subject
    origLocation = targetAppt.Location
    origBody = targetAppt.Body
    durMins = DateDiff("n", origStart, origEnd)

    newStart = origStart
    newEnd = origEnd
    startChanged = False
    endExplicit = False
    updated = False

    ' Handle start date/time updates
    If Not IsEmpty(startDateTime) Then
        If Left(CStr(startDateTime), 10) = "DATE_ONLY:" Then
            ' Date-only update: combine new date with existing time
            Dim newStartDate, existingStartTime
            newStartDate = ParseDate(Mid(CStr(startDateTime), 11))
            existingStartTime = TimeValue(origStart)
            newStart = newStartDate + existingStartTime
        ElseIf Left(CStr(startDateTime), 10) = "TIME_ONLY:" Then
            ' Time-only update: combine existing date with new time
            newStart = DateValue(origStart) + ParseTimeOnly(Mid(CStr(startDateTime), 11))
        Else
            ' Full date/time update
            newStart = startDateTime
        End If
        startChanged = True
    End If

    ' Handle end date/time updates
    If Not IsEmpty(endDateTime) Then
        If Left(CStr(endDateTime), 10) = "DATE_ONLY:" Then
            ' Date-only update: combine new date with existing time
            Dim newEndDate, existingEndTime
            newEndDate = ParseDate(Mid(CStr(endDateTime), 11))
            existingEndTime = TimeValue(origEnd)
            newEnd = newEndDate + existingEndTime
        ElseIf Left(CStr(endDateTime), 10) = "TIME_ONLY:" Then
            ' Time-only update: combine existing date with new time
            newEnd = DateValue(origEnd) + ParseTimeOnly(Mid(CStr(endDateTime), 11))
        Else
            ' Full date/time update
            newEnd = endDateTime
        End If
        endExplicit = True
    End If

    ' If start moved but no end provided, preserve duration
    If startChanged And Not endExplicit Then
        newEnd = DateAdd("n", durMins, newStart)
    End If

    ' Validate final times
    If newEnd <= newStart Then
        OutputError "End time cannot be before or equal to start time"
        UpdateCalendarEvent = False
        Exit Function
    End If

    ' Apply changes only if different
    If newStart <> origStart Then
        targetAppt.Start = newStart
        updated = True
    End If

    If newEnd <> origEnd Then
        targetAppt.End = newEnd
        updated = True
    End If

    If subject <> "" And subject <> origSubject Then
        targetAppt.Subject = subject
        updated = True
    End If

    If location <> "" And location <> origLocation Then
        targetAppt.Location = location
        updated = True
    End If

    If body <> "" And body <> origBody Then
        targetAppt.Body = body
        updated = True
    End If

    ' Check if anything was actually updated
    If Not updated Then
        OutputError "No fields were updated (check your parameters)"
        UpdateCalendarEvent = False
        Exit Function
    End If

    ' Save the changes (this creates an exception for recurring items)
    Err.Clear
    targetAppt.Save

    If Err.Number <> 0 Then
        ' Check for common recurring meeting errors
        If InStr(1, Err.Description, "skip", vbTextCompare) > 0 Or _
           InStr(1, Err.Description, "earlier", vbTextCompare) > 0 Or _
           InStr(1, Err.Description, "later", vbTextCompare) > 0 Then
            OutputError "Cannot move this occurrence to the new time because it would skip over other instances in the recurring series. Consider creating a new standalone meeting instead."
        Else
            OutputError "Failed to update calendar event: " & Err.Description
        End If
        UpdateCalendarEvent = False
        Exit Function
    End If

    ' For meetings with attendees, send update if requested
    If sendUpdate And targetAppt.MeetingStatus = olMeeting Then
        Err.Clear
        targetAppt.Send
        If Err.Number <> 0 Then
            ' Save succeeded but send failed - warn but don't fail
            OutputError "Warning: Meeting saved but failed to send update to attendees: " & Err.Description
        End If
    End If

    UpdateCalendarEvent = True

    ' Clean up (important for recurring items)
    Set targetAppt = Nothing
    Set recPattern = Nothing
    Set appointment = Nothing
    Set calendar = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
