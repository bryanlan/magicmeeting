' findAvailableRooms.vbs - Finds available conference rooms
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

' Max candidates per floor to check free/busy
Const MAX_CANDIDATES_PER_FLOOR = 10

' Main function
Sub Main()
    ' Get command line arguments
    Dim buildingStr, floorStr, startDateStr, startTimeStr, endDateStr, endTimeStr, capacityStr
    Dim building, floor, startDate, startTime, endDate, endTime, capacity
    Dim startDateTime, endDateTime

    ' Get and validate arguments
    buildingStr = GetArgument("building")
    floorStr = GetArgument("floor")
    startDateStr = GetArgument("startDate")
    startTimeStr = GetArgument("startTime")
    endDateStr = GetArgument("endDate")
    endTimeStr = GetArgument("endTime")
    capacityStr = GetArgument("capacity")

    ' Require building, startDate, startTime, endTime, capacity
    RequireArgument "building"
    RequireArgument "startDate"
    RequireArgument "startTime"
    RequireArgument "endTime"
    RequireArgument "capacity"

    building = Trim(buildingStr)

    ' Parse floor (optional)
    If floorStr <> "" Then
        If Not IsNumeric(floorStr) Then
            OutputError "Floor must be a number"
            WScript.Quit 1
        End If
        floor = CInt(floorStr)
    Else
        floor = -1 ' -1 means no floor filter
    End If

    ' Parse dates
    startDate = ParseDate(startDateStr)

    ' If end date not provided, use start date
    If endDateStr = "" Then
        endDate = startDate
        endDateStr = startDateStr
    Else
        endDate = ParseDate(endDateStr)
    End If

    ' Parse times and combine with dates
    startDateTime = CDate(startDateStr & " " & startTimeStr)
    endDateTime = CDate(endDateStr & " " & endTimeStr)

    If endDateTime <= startDateTime Then
        OutputError "End time must be after start time"
        WScript.Quit 1
    End If

    ' Parse capacity
    If Not IsNumeric(capacityStr) Then
        OutputError "Capacity must be a number"
        WScript.Quit 1
    End If
    capacity = CInt(capacityStr)

    ' Find available rooms
    Dim result
    result = FindAvailableRooms(building, floor, startDateTime, endDateTime, capacity)

    ' Output result as JSON
    OutputSuccess result
End Sub

' Finds available conference rooms matching criteria
Function FindAvailableRooms(building, floor, startDateTime, endDateTime, requiredCapacity)
    On Error Resume Next

    Dim outlookApp, namespace
    Dim i, j
    Dim candidates, candidateCount
    Dim json, roomName, roomEmail, roomCapacity, roomFloor, roomNumber
    Dim buildingUpper, emailPrefix

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    Set namespace = outlookApp.GetNamespace("MAPI")

    If Err.Number <> 0 Then
        OutputError "Failed to get MAPI namespace: " & Err.Description
        WScript.Quit 1
    End If

    ' Determine email prefix based on building
    ' STUDIO E, STUDIO D -> cfh prefix
    ' Building numbers -> cf{building} prefix (e.g., cf50 for Building 50)
    buildingUpper = UCase(building)

    If InStr(buildingUpper, "STUDIO") > 0 Then
        emailPrefix = "cfh"
    Else
        ' Assume numeric building - use cf{building}
        emailPrefix = "cf" & Replace(building, " ", "")
    End If

    ' Initialize candidates array
    ReDim candidates(100)
    candidateCount = 0

    ' Track count per floor
    Dim floorCountArray(9)
    For i = 0 To 9
        floorCountArray(i) = 0
    Next

    ' Determine which floors to scan
    Dim floorsToScan
    If floor >= 0 Then
        floorsToScan = Array(floor)
    Else
        ' Scan floors 1-5 (most common)
        floorsToScan = Array(1, 2, 3, 4, 5)
    End If

    ' Scan room numbers for each floor
    ' Room numbers are typically: {floor}{3 digits}, e.g., 2235 = floor 2, room 235
    Dim floorIdx, floorNum, roomNumStart, roomNumEnd, testRoomNum
    Dim recipient, testEmail

    For floorIdx = 0 To UBound(floorsToScan)
        floorNum = floorsToScan(floorIdx)

        ' Check if we've hit the cap for this floor
        If floorCountArray(floorNum) >= MAX_CANDIDATES_PER_FLOOR Then
            ' Skip this floor, we have enough candidates
        Else
            ' Scan room numbers for this floor
            ' Try common room number ranges: {floor}000-{floor}999
            roomNumStart = floorNum * 1000
            roomNumEnd = floorNum * 1000 + 999

            ' For efficiency, try specific sub-ranges that are more likely to have rooms
            Dim subRanges
            subRanges = Array(0, 100, 200, 300, 400, 500, 600, 700, 800, 900)

            Dim rangeIdx, rangeStart
            For rangeIdx = 0 To UBound(subRanges)
                If floorCountArray(floorNum) >= MAX_CANDIDATES_PER_FLOOR Then Exit For

                rangeStart = roomNumStart + subRanges(rangeIdx)

                ' Try 100 numbers in this range, but stop if we find enough
                For testRoomNum = rangeStart To rangeStart + 99
                    If floorCountArray(floorNum) >= MAX_CANDIDATES_PER_FLOOR Then Exit For

                    testEmail = emailPrefix & testRoomNum & "@microsoft.com"

                    Set recipient = namespace.CreateRecipient(testEmail)
                    recipient.Resolve

                    If recipient.Resolved Then
                        ' Check if this is a real Exchange room resource (not just an SMTP address)
                        ' Real rooms have Address starting with "/o=ExchangeLabs"
                        If InStr(1, recipient.Address, "/o=ExchangeLabs", vbTextCompare) > 0 Then
                            ' Found a real room! Parse its info
                            roomName = recipient.Name
                            roomEmail = testEmail
                            roomNumber = CStr(testRoomNum)
                            roomFloor = floorNum

                            ' Parse capacity from name
                            roomCapacity = 0
                            Dim capStart, capEnd, capStr
                            capStart = InStrRev(roomName, "(")
                            capEnd = InStrRev(roomName, ")")
                            If capStart > 0 And capEnd > capStart Then
                                capStr = Mid(roomName, capStart + 1, capEnd - capStart - 1)
                                If IsNumeric(capStr) Then
                                    roomCapacity = CInt(capStr)
                                End If
                            End If

                            ' Check if this room matches our building
                            If InStr(1, UCase(roomName), buildingUpper, vbTextCompare) > 0 Then
                            ' Add to candidates
                            If candidateCount > UBound(candidates) Then
                                ReDim Preserve candidates(candidateCount + 100)
                            End If

                                candidates(candidateCount) = roomName & "|" & roomEmail & "|" & roomCapacity & "|" & roomFloor & "|" & roomNumber
                                candidateCount = candidateCount + 1
                                floorCountArray(floorNum) = floorCountArray(floorNum) + 1
                            End If
                        End If
                    End If

                    Err.Clear
                Next
            Next
        End If
    Next

    ' Now check free/busy for each candidate using Recipient.FreeBusy
    Dim availableRooms, availableCount
    ReDim availableRooms(candidateCount)
    availableCount = 0

    Dim minPerChar, freeBusyString
    Dim isFree, slotStatus
    minPerChar = 30 ' 30-minute intervals

    ' Calculate number of slots needed for the time range
    Dim totalMinutes, slotsNeeded
    totalMinutes = DateDiff("n", startDateTime, endDateTime)
    slotsNeeded = Int(totalMinutes / minPerChar)
    If slotsNeeded < 1 Then slotsNeeded = 1

    For i = 0 To candidateCount - 1
        ' Parse candidate info
        Dim parts
        parts = Split(candidates(i), "|")
        roomName = parts(0)
        roomEmail = parts(1)
        roomCapacity = CInt(parts(2))
        roomFloor = CInt(parts(3))
        roomNumber = parts(4)

        ' Get free/busy for this room using Recipient.FreeBusy (not AddressEntry.GetFreeBusy)
        Set recipient = namespace.CreateRecipient(roomEmail)
        recipient.Resolve

        If recipient.Resolved Then
            ' Use Recipient.FreeBusy instead of AddressEntry.GetFreeBusy
            freeBusyString = recipient.FreeBusy(DateValue(startDateTime), minPerChar, True)

            If Err.Number = 0 And Len(freeBusyString) > 0 Then
                ' Calculate which slot in the free/busy string corresponds to our start time
                Dim startOfDay, minutesFromMidnight, startSlot
                startOfDay = DateValue(startDateTime)
                minutesFromMidnight = DateDiff("n", startOfDay, startDateTime)
                startSlot = Int(minutesFromMidnight / minPerChar) + 1 ' 1-based

                ' Check if all slots in our range are free
                isFree = True
                For j = 0 To slotsNeeded - 1
                    If startSlot + j <= Len(freeBusyString) Then
                        slotStatus = CInt(Mid(freeBusyString, startSlot + j, 1))
                        If slotStatus <> FB_FREE Then
                            isFree = False
                            Exit For
                        End If
                    End If
                Next

                If isFree Then
                    ' Room is available - add to results
                    availableRooms(availableCount) = candidates(i)
                    availableCount = availableCount + 1
                End If
            End If
        End If

        Err.Clear
    Next

    ' Build JSON response
    json = "{"
    json = json & """building"":""" & EscapeJSON(building) & ""","
    json = json & """requestedFloor"":" & floor & ","
    json = json & """requestedCapacity"":" & requiredCapacity & ","
    json = json & """timeRange"":{"
    json = json & """start"":""" & FormatDateTime(startDateTime) & ""","
    json = json & """end"":""" & FormatDateTime(endDateTime) & """"
    json = json & "},"
    json = json & """candidatesChecked"":" & candidateCount & ","
    json = json & """availableRooms"":["

    ' Sort and output available rooms
    ' Priority: 1) meets capacity + floor match, 2) meets capacity + other floor, 3) under capacity
    Dim tier1, tier2, tier3
    tier1 = ""
    tier2 = ""
    tier3 = ""

    For i = 0 To availableCount - 1
        parts = Split(availableRooms(i), "|")
        roomName = parts(0)
        roomEmail = parts(1)
        roomCapacity = CInt(parts(2))
        roomFloor = CInt(parts(3))
        roomNumber = parts(4)

        Dim roomJson
        roomJson = "{"
        roomJson = roomJson & """name"":""" & EscapeJSON(roomName) & ""","
        roomJson = roomJson & """email"":""" & EscapeJSON(roomEmail) & ""","
        roomJson = roomJson & """capacity"":" & roomCapacity & ","
        roomJson = roomJson & """floor"":" & roomFloor & ","
        roomJson = roomJson & """roomNumber"":""" & EscapeJSON(roomNumber) & ""","
        roomJson = roomJson & """meetsCapacity"":" & LCase(CStr(roomCapacity >= requiredCapacity)) & ","
        roomJson = roomJson & """floorMatch"":" & LCase(CStr(floor < 0 Or roomFloor = floor))
        roomJson = roomJson & "}"

        If roomCapacity >= requiredCapacity Then
            If floor < 0 Or roomFloor = floor Then
                ' Tier 1: meets capacity AND floor match (or no floor specified)
                If tier1 <> "" Then tier1 = tier1 & ","
                tier1 = tier1 & roomJson
            Else
                ' Tier 2: meets capacity but different floor
                If tier2 <> "" Then tier2 = tier2 & ","
                tier2 = tier2 & roomJson
            End If
        Else
            ' Tier 3: under capacity
            If tier3 <> "" Then tier3 = tier3 & ","
            tier3 = tier3 & roomJson
        End If
    Next

    ' Combine tiers
    Dim allRooms
    allRooms = ""
    If tier1 <> "" Then allRooms = tier1
    If tier2 <> "" Then
        If allRooms <> "" Then allRooms = allRooms & ","
        allRooms = allRooms & tier2
    End If
    If tier3 <> "" Then
        If allRooms <> "" Then allRooms = allRooms & ","
        allRooms = allRooms & tier3
    End If

    json = json & allRooms
    json = json & "],"
    json = json & """totalAvailable"":" & availableCount
    json = json & "}"

    FindAvailableRooms = json

    ' Clean up
    Set addressEntry = Nothing
    Set recipient = Nothing
    Set addressEntries = Nothing
    Set addressList = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
