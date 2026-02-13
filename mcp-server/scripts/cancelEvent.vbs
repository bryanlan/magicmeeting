' cancelEvent.vbs - Cancels a meeting with an optional custom message
' For recurring meetings: use occurrenceStart for one instance, or cancelSeries for entire series
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Additional Outlook constants not in utils.vbs
Const olMeetingCanceled = 5

' Main function
Sub Main()
    ' Get command line arguments
    Dim eventId, calendarName, comment, occurrenceStartStr, cancelSeries

    ' Get and validate arguments
    eventId = GetArgument("eventId")
    calendarName = GetArgument("calendar")
    occurrenceStartStr = GetArgument("occurrenceStart")
    cancelSeries = LCase(GetArgument("cancelSeries")) = "true"

    ' Read comment from temp file if path provided, otherwise from command line
    Dim bodyFilePath
    bodyFilePath = GetArgument("bodyFile")
    If bodyFilePath <> "" Then
        Dim bodyFso2, bodyStream
        Set bodyFso2 = CreateObject("Scripting.FileSystemObject")
        Set bodyStream = bodyFso2.OpenTextFile(bodyFilePath, 1)
        comment = bodyStream.ReadAll
        bodyStream.Close
        Set bodyStream = Nothing
        Set bodyFso2 = Nothing
    Else
        comment = GetArgument("body")
    End If

    ' Require event ID
    RequireArgument "eventId"

    ' Cancel the meeting
    Dim result
    result = CancelMeeting(eventId, calendarName, comment, occurrenceStartStr, cancelSeries)

    ' Output success
    OutputSuccess "{""success"":" & LCase(CStr(result)) & "}"
End Sub

' Cancels a meeting with an optional custom cancellation message
' For recurring meetings: use occurrenceStart for one instance, or cancelSeries for entire series
' Only works if the user is the organizer
Function CancelMeeting(eventId, calendarName, comment, occurrenceStartStr, cancelSeries)
    On Error Resume Next

    ' Create Outlook objects
    Dim outlookApp, namespace, calendar, master, targetItem
    Dim recPattern, occurrenceStart

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()

    ' Get MAPI namespace
    Set namespace = outlookApp.GetNamespace("MAPI")

    ' Get calendar folder (needed for StoreID)
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If

    ' Get the appointment by EntryID (this is the series master for recurring meetings)
    Set master = namespace.GetItemFromID(eventId, calendar.StoreID)

    ' Check if appointment was found
    If master Is Nothing Then
        OutputError "Event not found with ID: " & eventId
        CancelMeeting = False
        Exit Function
    End If

    ' Check if this is a meeting (MeetingStatus = olMeeting means you're the organizer)
    If master.MeetingStatus <> olMeeting Then
        OutputError "This event is not a meeting or you are not the organizer"
        CancelMeeting = False
        Exit Function
    End If

    ' Handle recurring vs single meetings differently
    If master.IsRecurring Then
        If cancelSeries Then
            ' Cancel entire series - use the master directly
            Set targetItem = master
        ElseIf occurrenceStartStr = "" Then
            ' Neither cancelSeries nor occurrenceStart provided - error
            OutputError "For recurring meetings: provide occurrenceStart to cancel one instance, or cancelSeries=true to cancel the entire series"
            CancelMeeting = False
            Exit Function
        Else
            ' Cancel specific occurrence
            ' Parse the occurrence start datetime
            If Not IsDate(occurrenceStartStr) Then
                OutputError "Invalid occurrenceStart format. Expected: MM/DD/YYYY HH:MM AM/PM"
                CancelMeeting = False
                Exit Function
            End If
            occurrenceStart = CDate(occurrenceStartStr)

            ' Get the recurrence pattern and then the specific occurrence
            Set recPattern = master.GetRecurrencePattern()

            Err.Clear
            Set targetItem = recPattern.GetOccurrence(occurrenceStart)

            If Err.Number <> 0 Then
                OutputError "Failed to get occurrence at " & occurrenceStartStr & ": " & Err.Description
                CancelMeeting = False
                Exit Function
            End If

            If targetItem Is Nothing Then
                OutputError "No occurrence found at " & occurrenceStartStr
                CancelMeeting = False
                Exit Function
            End If

            Set recPattern = Nothing
        End If
    Else
        ' Single (non-recurring) meeting - cancel directly
        Set targetItem = master
    End If

    ' Now cancel the target item (either the single meeting or the specific occurrence)
    Err.Clear

    ' If comment is provided, prepend to body
    If comment <> "" Then
        targetItem.Body = comment & vbCrLf & vbCrLf & "---" & vbCrLf & vbCrLf & targetItem.Body
    End If

    ' Set the meeting status to canceled and send the cancellation
    targetItem.MeetingStatus = olMeetingCanceled
    targetItem.Send

    If Err.Number <> 0 Then
        OutputError "Failed to cancel meeting: " & Err.Description
        CancelMeeting = False
    Else
        CancelMeeting = True
    End If

    ' Clean up
    Set targetItem = Nothing
    Set master = Nothing
    Set calendar = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
