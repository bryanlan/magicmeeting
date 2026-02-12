' removeAttendee.vbs - Remove an attendee from an existing meeting
Option Explicit

' Outlook recurrence state constants
Const olApptNotRecurring = 0
Const olApptMaster = 1
Const olApptOccurrence = 2
Const olApptException = 3

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    Dim eventId, attendee, sendUpdate, originalStartStr, updateSeriesStr
    Dim originalStart

    eventId = GetArgument("eventId")
    attendee = GetArgument("attendee")
    sendUpdate = GetArgument("sendUpdate")
    originalStartStr = GetArgument("originalStart")  ' For recurring: original occurrence date/time
    updateSeriesStr = GetArgument("updateSeries")    ' For recurring: update entire series

    RequireArgument "eventId"
    RequireArgument "attendee"

    ' Default to sending update
    If sendUpdate = "" Then sendUpdate = "true"

    ' Parse original start date/time for recurring meetings
    If originalStartStr <> "" Then
        If Not IsDate(originalStartStr) Then
            OutputError "Invalid originalStart format. Use: MM/DD/YYYY HH:MM AM/PM"
            WScript.Quit 1
        End If
        originalStart = CDate(originalStartStr)
    End If

    Dim result, updateSeries
    updateSeries = (LCase(updateSeriesStr) = "true")
    result = RemoveAttendeeFromMeeting(eventId, attendee, (LCase(sendUpdate) = "true"), originalStart, updateSeries)

    OutputSuccess result
End Sub

' Remove an attendee from an existing meeting
Function RemoveAttendeeFromMeeting(eventId, attendeeToRemove, sendUpdate, originalStart, updateSeries)
    On Error Resume Next

    Dim outlookApp, ns, appt, targetAppt, recip, recPattern
    Dim json, found, removedName, removedEmail
    Dim i, recipName, recipAddress, recState

    found = False

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    Set ns = outlookApp.GetNamespace("MAPI")

    ' Get the appointment by EntryID
    Set appt = ns.GetItemFromID(eventId)

    If appt Is Nothing Then
        RemoveAttendeeFromMeeting = "{""success"":false,""error"":""Event not found with ID: " & EscapeJSON(eventId) & """}"
        Exit Function
    End If

    If Err.Number <> 0 Then
        RemoveAttendeeFromMeeting = "{""success"":false,""error"":""Failed to get event: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    ' Verify it's an appointment
    If TypeName(appt) <> "AppointmentItem" Then
        RemoveAttendeeFromMeeting = "{""success"":false,""error"":""Item is not an appointment. Type: " & TypeName(appt) & """}"
        Exit Function
    End If

    ' Determine the target appointment (handle recurring vs non-recurring)
    recState = appt.RecurrenceState

    If recState = olApptMaster Then
        ' This is a recurring series master
        If updateSeries Then
            ' Update the entire series - use the master appointment directly
            Set targetAppt = appt
        ElseIf Not IsEmpty(originalStart) Then
            ' Get the recurrence pattern and fetch the specific occurrence
            Set recPattern = appt.GetRecurrencePattern()
            Err.Clear
            Set targetAppt = recPattern.GetOccurrence(originalStart)

            If Err.Number <> 0 Or targetAppt Is Nothing Then
                RemoveAttendeeFromMeeting = "{""success"":false,""error"":""Could not find occurrence at " & originalStart & ". Make sure originalStart matches the occurrence's ORIGINAL start date+time exactly. Error: " & EscapeJSON(Err.Description) & """}"
                Exit Function
            End If
            Err.Clear
        Else
            RemoveAttendeeFromMeeting = "{""success"":false,""error"":""This is a recurring meeting series. Either set updateSeries=true to update ALL occurrences, or provide originalStart parameter to modify a specific instance (format: MM/DD/YYYY HH:MM AM/PM).""}"
            Exit Function
        End If
    ElseIf recState = olApptOccurrence Or recState = olApptException Then
        ' Already have a specific occurrence or exception - use it directly
        Set targetAppt = appt
    Else
        ' Non-recurring appointment - use directly
        Set targetAppt = appt
    End If

    ' Normalize the search string for matching
    Dim searchStr
    searchStr = LCase(attendeeToRemove)

    ' Find and remove the attendee (iterate backwards to avoid index issues)
    For i = targetAppt.Recipients.Count To 1 Step -1
        Set recip = targetAppt.Recipients.Item(i)
        recipName = LCase(recip.Name)
        recipAddress = LCase(recip.Address)

        ' Match by email, name, or partial match
        If InStr(recipAddress, searchStr) > 0 Or _
           InStr(recipName, searchStr) > 0 Or _
           recipAddress = searchStr Or _
           recipName = searchStr Then
            ' Found the attendee to remove
            removedName = recip.Name
            removedEmail = recip.Address

            ' Remove from recipients
            targetAppt.Recipients.Remove i
            found = True
            Exit For
        End If
    Next

    If Not found Then
        RemoveAttendeeFromMeeting = "{""success"":false,""error"":""Attendee not found: " & EscapeJSON(attendeeToRemove) & """}"
        Exit Function
    End If

    ' Save the appointment
    Err.Clear
    targetAppt.Save

    If Err.Number <> 0 Then
        RemoveAttendeeFromMeeting = "{""success"":false,""error"":""Failed to save: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    ' Send update to all attendees if requested
    If sendUpdate And targetAppt.MeetingStatus = olMeeting Then
        Err.Clear
        targetAppt.Send

        If Err.Number <> 0 Then
            ' Save succeeded but send failed
            RemoveAttendeeFromMeeting = "{""success"":true,""attendeeRemoved"":""" & EscapeJSON(removedName) & """,""email"":""" & EscapeJSON(removedEmail) & """,""updateSent"":false,""warning"":""Attendee removed but update send failed: " & EscapeJSON(Err.Description) & """}"
            Exit Function
        End If
    End If

    ' Build success response
    json = "{"
    json = json & """success"":true,"
    json = json & """attendeeRemoved"":""" & EscapeJSON(removedName) & ""","
    json = json & """email"":""" & EscapeJSON(removedEmail) & ""","
    json = json & """updateSent"":" & LCase(CStr(sendUpdate))
    json = json & "}"

    RemoveAttendeeFromMeeting = json

    ' Cleanup
    Set recip = Nothing
    Set targetAppt = Nothing
    Set recPattern = Nothing
    Set appt = Nothing
    Set ns = Nothing
    Set outlookApp = Nothing
End Function

' Run main
Main
