' findFreeSlots.vbs - Finds available time slots in the calendar
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Constants for working hours
Const DEFAULT_WORK_DAY_START = 9 ' 9 AM
Const DEFAULT_WORK_DAY_END = 17 ' 5 PM
Const DEFAULT_SLOT_DURATION = 30 ' 30 minutes

' Main function
Sub Main()
    ' Get command line arguments
    Dim startDateStr, endDateStr, durationStr, workDayStartStr, workDayEndStr, calendarName
    Dim startDate, endDate, duration, workDayStart, workDayEnd

    ' Get and validate arguments
    startDateStr = GetArgument("startDate")
    endDateStr = GetArgument("endDate")
    durationStr = GetArgument("duration")
    workDayStartStr = GetArgument("workDayStart")
    workDayEndStr = GetArgument("workDayEnd")
    calendarName = GetArgument("calendar")

    ' Require start date
    RequireArgument "startDate"

    ' Parse dates
    startDate = ParseDate(startDateStr)

    ' If end date is not provided, use 7 days from start date
    If endDateStr = "" Then
        endDate = DateAdd("d", 7, startDate)
    Else
        endDate = ParseDate(endDateStr)
    End If

    ' Ensure end date is not before start date
    If endDate < startDate Then
        OutputError "End date cannot be before start date"
        WScript.Quit 1
    End If

    ' Parse duration (in minutes)
    If durationStr = "" Then
        duration = DEFAULT_SLOT_DURATION
    Else
        If Not IsNumeric(durationStr) Then
            OutputError "Duration must be a number (minutes)"
            WScript.Quit 1
        End If
        duration = CInt(durationStr)
    End If

    ' Parse work day start/end hours
    If workDayStartStr = "" Then
        workDayStart = DEFAULT_WORK_DAY_START
    Else
        If Not IsNumeric(workDayStartStr) Then
            OutputError "Work day start hour must be a number (0-23)"
            WScript.Quit 1
        End If
        workDayStart = CInt(workDayStartStr)
        If workDayStart < 0 Or workDayStart > 23 Then
            OutputError "Work day start hour must be between 0 and 23"
            WScript.Quit 1
        End If
    End If

    If workDayEndStr = "" Then
        workDayEnd = DEFAULT_WORK_DAY_END
    Else
        If Not IsNumeric(workDayEndStr) Then
            OutputError "Work day end hour must be a number (0-23)"
            WScript.Quit 1
        End If
        workDayEnd = CInt(workDayEndStr)
        If workDayEnd < 0 Or workDayEnd > 23 Then
            OutputError "Work day end hour must be between 0 and 23"
            WScript.Quit 1
        End If
    End If

    ' Ensure work day end is after work day start
    If workDayEnd <= workDayStart Then
        OutputError "Work day end hour must be after work day start hour"
        WScript.Quit 1
    End If

    ' Find free slots
    Dim freeSlots
    freeSlots = FindFreeTimeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendarName)

    ' Output free slots as JSON
    OutputSuccess freeSlots
End Sub

' Finds free time slots in the calendar
Function FindFreeTimeSlots(startDate, endDate, duration, workDayStart, workDayEnd, calendarName)
    On Error Resume Next

    ' Create Outlook objects
    Dim outlookApp, namespace, calendar, calItems, restrictedItems
    Dim currentDate, currentSlotStart, currentSlotEnd, i, appt, isSlotFree
    Dim busyStarts(), busyEnds(), busyCount

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    Set namespace = outlookApp.GetNamespace("MAPI")

    If Err.Number <> 0 Then
        OutputError "Failed to get MAPI namespace: " & Err.Description
        WScript.Quit 1
    End If

    ' Get calendar folder
    If calendarName = "" Then
        Set calendar = GetDefaultCalendar(outlookApp)
    Else
        Set calendar = GetCalendarByName(outlookApp, calendarName)
    End If

    ' Get calendar items - IMPORTANT: IncludeRecurrences must be set BEFORE Sort
    ' and the filter must use a specific format for recurring items
    Set calItems = calendar.Items
    calItems.IncludeRecurrences = True
    calItems.Sort "[Start]"

    ' For recurring items, we need to use a different filter format
    ' The date must be in a format Outlook understands for recurrence expansion
    Dim filterStr, startStr, endStr
    startStr = Month(startDate) & "/" & Day(startDate) & "/" & Year(startDate) & " 12:00 AM"
    endStr = Month(endDate) & "/" & Day(endDate) & "/" & Year(endDate) & " 11:59 PM"
    filterStr = "[Start] >= '" & startStr & "' AND [End] <= '" & endStr & "'"
    Set restrictedItems = calItems.Restrict(filterStr)

    ' Build arrays of busy time slots
    busyCount = 0
    ReDim busyStarts(100)
    ReDim busyEnds(100)

    ' Iterate through items and filter by actual date (not series date)
    For Each appt In restrictedItems
        ' Double-check the appointment falls within our date range
        ' (recurring items can be tricky with Restrict)
        If appt.Start >= startDate And appt.Start < DateAdd("d", 1, endDate) Then
            ' Only consider events marked as Busy or Out of Office (not Tentative or Free)
            If appt.BusyStatus = olBusy Or appt.BusyStatus = olOutOfOffice Then
                If busyCount > UBound(busyStarts) Then
                    ReDim Preserve busyStarts(busyCount + 100)
                    ReDim Preserve busyEnds(busyCount + 100)
                End If
                busyStarts(busyCount) = appt.Start
                busyEnds(busyCount) = appt.End
                busyCount = busyCount + 1
            End If
        End If
    Next

    ' Build JSON array of free slots
    Dim json, slotCount
    json = "["
    slotCount = 0

    ' Loop through each day in the date range
    currentDate = startDate
    Do While currentDate <= endDate
        ' Skip weekends (Saturday = 7, Sunday = 1)
        If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
            ' Loop through each potential slot in the work day
            currentSlotStart = DateAdd("h", workDayStart, currentDate)

            Do While DateAdd("n", duration, currentSlotStart) <= DateAdd("h", workDayEnd, currentDate)
                currentSlotEnd = DateAdd("n", duration, currentSlotStart)

                ' Check if the slot is free
                isSlotFree = True

                For i = 0 To busyCount - 1
                    ' If the slot overlaps with a busy slot, it's not free
                    If (currentSlotStart < busyEnds(i)) And (currentSlotEnd > busyStarts(i)) Then
                        isSlotFree = False
                        Exit For
                    End If
                Next

                ' If the slot is free, add it to the JSON
                If isSlotFree Then
                    If slotCount > 0 Then json = json & ","
                    json = json & "{""start"":""" & FormatDateTime(currentSlotStart) & """,""end"":""" & FormatDateTime(currentSlotEnd) & """}"
                    slotCount = slotCount + 1
                End If

                ' Move to the next slot
                currentSlotStart = DateAdd("n", 30, currentSlotStart) ' 30-minute increments
            Loop
        End If

        ' Move to the next day
        currentDate = DateAdd("d", 1, currentDate)
    Loop

    json = json & "]"

    FindFreeTimeSlots = json

    ' Clean up
    Set restrictedItems = Nothing
    Set calItems = Nothing
    Set calendar = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
