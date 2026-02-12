' getFreeBusy.vbs - Gets free/busy information for multiple attendees
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' FreeBusy status constants
Const FB_FREE = 0
Const FB_TENTATIVE = 1
Const FB_BUSY = 2
Const FB_OOF = 3

' Main function
Sub Main()
    ' Get command line arguments
    Dim startDateStr, endDateStr, attendeesStr, durationStr
    Dim startDate, endDate, duration

    ' Get and validate arguments
    startDateStr = GetArgument("startDate")
    endDateStr = GetArgument("endDate")
    attendeesStr = GetArgument("attendees")
    durationStr = GetArgument("duration")

    ' Require start date and attendees
    RequireArgument "startDate"
    RequireArgument "attendees"

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

    ' Parse duration (in minutes), default to 30
    If durationStr = "" Then
        duration = 30
    Else
        If Not IsNumeric(durationStr) Then
            OutputError "Duration must be a number (minutes)"
            WScript.Quit 1
        End If
        duration = CInt(durationStr)
    End If

    ' Get free/busy information
    Dim result
    result = GetAttendeesFreeBusy(startDate, endDate, attendeesStr, duration)

    ' Output result as JSON
    OutputSuccess result
End Sub

' Gets free/busy information for multiple attendees
Function GetAttendeesFreeBusy(startDate, endDate, attendeesStr, duration)
    On Error Resume Next

    ' Create Outlook objects
    Dim outlookApp, namespace, recipient, addressEntry
    Dim attendees, i, freeBusyStr, freeBusyData
    Dim json, attendeeJson, attendeeResults
    Dim totalDays, slotsPerDay, totalSlots
    Dim commonFreeSlots, slotIndex, dayOffset, slotInDay

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    Set namespace = outlookApp.GetNamespace("MAPI")

    If Err.Number <> 0 Then
        OutputError "Failed to get MAPI namespace: " & Err.Description
        WScript.Quit 1
    End If

    ' Split attendees by semicolon
    attendees = Split(attendeesStr, ";")

    ' Calculate time range parameters
    ' GetFreeBusy uses minutes per character (we'll use 30 min intervals)
    Dim minPerChar
    minPerChar = 30

    ' Calculate number of days and slots
    totalDays = DateDiff("d", startDate, endDate) + 1
    slotsPerDay = 24 * 60 / minPerChar ' 48 slots per day at 30 min intervals
    totalSlots = totalDays * slotsPerDay

    ' Initialize common free slots array (all free initially)
    ReDim commonFreeSlots(totalSlots - 1)
    For i = 0 To totalSlots - 1
        commonFreeSlots(i) = FB_FREE
    Next

    ' Build JSON for attendee results
    attendeeResults = "["

    ' Process each attendee
    For i = 0 To UBound(attendees)
        Dim attendeeEmail, freeBusyString, status
        attendeeEmail = Trim(attendees(i))

        If attendeeEmail <> "" Then
            ' Create and resolve recipient
            Set recipient = namespace.CreateRecipient(attendeeEmail)
            recipient.Resolve

            If Err.Number <> 0 Then
                Err.Clear
                ' Add attendee with error status
                If i > 0 Then attendeeResults = attendeeResults & ","
                attendeeResults = attendeeResults & "{"
                attendeeResults = attendeeResults & """email"":""" & EscapeJSON(attendeeEmail) & ""","
                attendeeResults = attendeeResults & """resolved"":false,"
                attendeeResults = attendeeResults & """error"":""Could not resolve recipient"""
                attendeeResults = attendeeResults & "}"
            ElseIf Not recipient.Resolved Then
                ' Recipient not resolved
                If i > 0 Then attendeeResults = attendeeResults & ","
                attendeeResults = attendeeResults & "{"
                attendeeResults = attendeeResults & """email"":""" & EscapeJSON(attendeeEmail) & ""","
                attendeeResults = attendeeResults & """resolved"":false,"
                attendeeResults = attendeeResults & """error"":""Recipient not found in address book"""
                attendeeResults = attendeeResults & "}"
            Else
                ' Get the address entry
                Set addressEntry = recipient.AddressEntry

                ' Get free/busy string
                ' Parameters: Start date, minutes per character, complete format (True = include start/end)
                freeBusyString = addressEntry.GetFreeBusy(startDate, minPerChar, True)

                If Err.Number <> 0 Then
                    Err.Clear
                    If i > 0 Then attendeeResults = attendeeResults & ","
                    attendeeResults = attendeeResults & "{"
                    attendeeResults = attendeeResults & """email"":""" & EscapeJSON(attendeeEmail) & ""","
                    attendeeResults = attendeeResults & """name"":""" & EscapeJSON(addressEntry.Name) & ""","
                    attendeeResults = attendeeResults & """resolved"":true,"
                    attendeeResults = attendeeResults & """error"":""Could not retrieve free/busy information"""
                    attendeeResults = attendeeResults & "}"
                Else
                    ' Update common free slots based on this attendee's schedule
                    Dim j, charStatus
                    For j = 1 To Len(freeBusyString)
                        If j <= totalSlots Then
                            charStatus = CInt(Mid(freeBusyString, j, 1))
                            ' If attendee is busy/tentative/OOF, mark slot as not available
                            If charStatus > commonFreeSlots(j - 1) Then
                                commonFreeSlots(j - 1) = charStatus
                            End If
                        End If
                    Next

                    ' Build detailed schedule for this attendee
                    Dim scheduleBlocks, blockStart, blockEnd, currentStatus, prevStatus
                    Dim blockDate, blockHour, blockMinute
                    scheduleBlocks = "["

                    ' Parse the free/busy string into time blocks
                    Dim blockCount
                    blockCount = 0
                    prevStatus = -1

                    For j = 1 To Len(freeBusyString)
                        If j <= totalSlots Then
                            currentStatus = CInt(Mid(freeBusyString, j, 1))

                            ' Calculate date/time for this slot
                            dayOffset = Int((j - 1) / slotsPerDay)
                            slotInDay = (j - 1) Mod slotsPerDay
                            blockDate = DateAdd("d", dayOffset, startDate)

                            ' If status changed or first slot, start new block
                            If currentStatus <> prevStatus Then
                                ' Close previous block if exists
                                If prevStatus >= 0 And blockCount > 0 Then
                                    blockEnd = DateAdd("n", slotInDay * minPerChar, blockDate)
                                    scheduleBlocks = scheduleBlocks & """end"":""" & FormatDateTime(blockEnd) & ""","
                                    scheduleBlocks = scheduleBlocks & """status"":""" & GetFreeBusyStatusText(prevStatus) & """}"
                                End If

                                ' Start new block
                                blockStart = DateAdd("n", slotInDay * minPerChar, blockDate)
                                If blockCount > 0 Then scheduleBlocks = scheduleBlocks & ","
                                scheduleBlocks = scheduleBlocks & "{"
                                scheduleBlocks = scheduleBlocks & """start"":""" & FormatDateTime(blockStart) & ""","

                                blockCount = blockCount + 1
                                prevStatus = currentStatus
                            End If
                        End If
                    Next

                    ' Close final block
                    If prevStatus >= 0 Then
                        dayOffset = Int((Len(freeBusyString)) / slotsPerDay)
                        slotInDay = (Len(freeBusyString)) Mod slotsPerDay
                        blockDate = DateAdd("d", dayOffset, startDate)
                        blockEnd = DateAdd("n", slotInDay * minPerChar, blockDate)
                        scheduleBlocks = scheduleBlocks & """end"":""" & FormatDateTime(blockEnd) & ""","
                        scheduleBlocks = scheduleBlocks & """status"":""" & GetFreeBusyStatusText(prevStatus) & """}"
                    End If

                    scheduleBlocks = scheduleBlocks & "]"

                    ' Add attendee result with schedule
                    If i > 0 Then attendeeResults = attendeeResults & ","
                    attendeeResults = attendeeResults & "{"
                    attendeeResults = attendeeResults & """email"":""" & EscapeJSON(attendeeEmail) & ""","
                    attendeeResults = attendeeResults & """name"":""" & EscapeJSON(addressEntry.Name) & ""","
                    attendeeResults = attendeeResults & """resolved"":true,"
                    attendeeResults = attendeeResults & """freeBusyRetrieved"":true,"
                    attendeeResults = attendeeResults & """schedule"":" & scheduleBlocks
                    attendeeResults = attendeeResults & "}"
                End If
            End If
        End If
    Next

    attendeeResults = attendeeResults & "]"

    ' Find common free slots (only during work hours 9 AM - 5 PM, weekdays)
    Dim commonSlots, slotStart, slotEnd, currentDate, currentSlotIndex
    Dim hour, isWorkHour, isWeekday
    Dim slotsNeeded

    ' Calculate how many 30-min slots we need for the requested duration
    slotsNeeded = duration / minPerChar
    If slotsNeeded < 1 Then slotsNeeded = 1

    commonSlots = "["
    Dim slotCount
    slotCount = 0

    ' Iterate through each slot
    For currentSlotIndex = 0 To totalSlots - 1
        ' Calculate the date/time for this slot
        dayOffset = Int(currentSlotIndex / slotsPerDay)
        slotInDay = currentSlotIndex Mod slotsPerDay
        currentDate = DateAdd("d", dayOffset, startDate)
        hour = Int(slotInDay * minPerChar / 60)

        ' Check if it's a weekday (2-6 = Mon-Fri)
        isWeekday = (Weekday(currentDate) >= 2 And Weekday(currentDate) <= 6)

        ' Check if it's work hours (9 AM - 5 PM)
        isWorkHour = (hour >= 9 And hour < 17)

        ' Only consider work hours on weekdays
        If isWeekday And isWorkHour Then
            ' Check if this slot and enough following slots are all free
            Dim allFree, k
            allFree = True

            For k = 0 To slotsNeeded - 1
                If currentSlotIndex + k >= totalSlots Then
                    allFree = False
                    Exit For
                End If
                If commonFreeSlots(currentSlotIndex + k) <> FB_FREE Then
                    allFree = False
                    Exit For
                End If
            Next

            If allFree Then
                ' Calculate start and end times
                slotStart = DateAdd("n", slotInDay * minPerChar, currentDate)
                slotEnd = DateAdd("n", duration, slotStart)

                ' Only add if end time is still within work hours
                If Hour(slotEnd) <= 17 Or (Hour(slotEnd) = 17 And Minute(slotEnd) = 0) Then
                    If slotCount > 0 Then commonSlots = commonSlots & ","
                    commonSlots = commonSlots & "{"
                    commonSlots = commonSlots & """start"":""" & FormatDateTime(slotStart) & ""","
                    commonSlots = commonSlots & """end"":""" & FormatDateTime(slotEnd) & """"
                    commonSlots = commonSlots & "}"
                    slotCount = slotCount + 1
                End If
            End If
        End If
    Next

    commonSlots = commonSlots & "]"

    ' Build final JSON result
    json = "{"
    json = json & """startDate"":""" & FormatDate(startDate) & ""","
    json = json & """endDate"":""" & FormatDate(endDate) & ""","
    json = json & """duration"":" & duration & ","
    json = json & """attendees"":" & attendeeResults & ","
    json = json & """commonFreeSlots"":" & commonSlots & ","
    json = json & """totalSlotsFound"":" & slotCount
    json = json & "}"

    GetAttendeesFreeBusy = json

    ' Clean up
    Set addressEntry = Nothing
    Set recipient = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Helper function to convert free/busy status code to text
Function GetFreeBusyStatusText(statusCode)
    Select Case statusCode
        Case FB_FREE
            GetFreeBusyStatusText = "Free"
        Case FB_TENTATIVE
            GetFreeBusyStatusText = "Tentative"
        Case FB_BUSY
            GetFreeBusyStatusText = "Busy"
        Case FB_OOF
            GetFreeBusyStatusText = "Out of Office"
        Case Else
            GetFreeBusyStatusText = "Unknown"
    End Select
End Function

' Run the main function
Main
