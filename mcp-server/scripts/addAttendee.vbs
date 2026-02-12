' addAttendee.vbs - Add an attendee to an existing meeting and send update
Option Explicit

' Recipient type constants
Const olRequired = 1
Const olOptional = 2
Const olResource = 3

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    Dim eventId, attendee, attendeeType, sendUpdate

    eventId = GetArgument("eventId")
    attendee = GetArgument("attendee")
    attendeeType = GetArgument("type")
    sendUpdate = GetArgument("sendUpdate")

    RequireArgument "eventId"
    RequireArgument "attendee"

    ' Default to required attendee
    If attendeeType = "" Then attendeeType = "required"
    ' Default to sending update
    If sendUpdate = "" Then sendUpdate = "true"

    Dim result
    result = AddAttendeeToMeeting(eventId, attendee, attendeeType, (LCase(sendUpdate) = "true"))

    OutputSuccess result
End Sub

' Add an attendee to an existing meeting
Function AddAttendeeToMeeting(eventId, attendeeEmail, attendeeType, sendUpdate)
    On Error Resume Next

    Dim outlookApp, ns, appt, recip, recipType
    Dim json

    ' Determine recipient type
    Select Case LCase(attendeeType)
        Case "optional"
            recipType = olOptional
        Case "resource"
            recipType = olResource
        Case Else
            recipType = olRequired
    End Select

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    Set ns = outlookApp.GetNamespace("MAPI")

    ' Get the appointment by EntryID
    Set appt = ns.GetItemFromID(eventId)

    If appt Is Nothing Then
        AddAttendeeToMeeting = "{""success"":false,""error"":""Event not found with ID: " & EscapeJSON(eventId) & """}"
        Exit Function
    End If

    If Err.Number <> 0 Then
        AddAttendeeToMeeting = "{""success"":false,""error"":""Failed to get event: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    ' Verify it's an appointment
    If TypeName(appt) <> "AppointmentItem" Then
        AddAttendeeToMeeting = "{""success"":false,""error"":""Item is not an appointment. Type: " & TypeName(appt) & """}"
        Exit Function
    End If

    ' Check if we're the organizer (only organizer can add attendees properly)
    ' Note: For meetings you created, you are the organizer

    ' Add the new attendee
    Err.Clear
    Set recip = appt.Recipients.Add(attendeeEmail)

    If Err.Number <> 0 Then
        AddAttendeeToMeeting = "{""success"":false,""error"":""Failed to add recipient: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    ' Set attendee type
    recip.Type = recipType

    ' Resolve all recipients
    If Not appt.Recipients.ResolveAll Then
        ' Try to resolve just the new one
        If Not recip.Resolve Then
            AddAttendeeToMeeting = "{""success"":false,""error"":""Could not resolve attendee address: " & EscapeJSON(attendeeEmail) & """}"
            Exit Function
        End If
    End If

    ' Get the resolved name
    Dim resolvedName
    resolvedName = recip.Name

    ' Save the appointment - Exchange will send update only to the new attendee
    ' Using Save instead of Send avoids notifying ALL attendees
    Err.Clear
    appt.Save

    If Err.Number <> 0 Then
        AddAttendeeToMeeting = "{""success"":false,""error"":""Failed to save: " & EscapeJSON(Err.Description) & """}"
        Exit Function
    End If

    ' If sendUpdate is true, explicitly send invite to just the new attendee
    ' by forwarding as vCal (Send() would notify everyone)
    If sendUpdate Then
        Err.Clear
        Dim fwdMail
        Set fwdMail = appt.ForwardAsVcal()

        If Err.Number <> 0 Then
            ' Save worked but forward failed - attendee is added but may not have received invite
            AddAttendeeToMeeting = "{""success"":true,""attendeeAdded"":""" & EscapeJSON(attendeeEmail) & """,""resolvedName"":""" & EscapeJSON(resolvedName) & """,""type"":""" & attendeeType & """,""updateSent"":false,""warning"":""Attendee added but invite forward failed: " & EscapeJSON(Err.Description) & """}"
            Exit Function
        End If

        ' Clear recipients and add only the new attendee
        Do While fwdMail.Recipients.Count > 0
            fwdMail.Recipients.Remove 1
        Loop
        fwdMail.Recipients.Add attendeeEmail

        If Not fwdMail.Recipients.ResolveAll Then
            AddAttendeeToMeeting = "{""success"":true,""attendeeAdded"":""" & EscapeJSON(attendeeEmail) & """,""resolvedName"":""" & EscapeJSON(resolvedName) & """,""type"":""" & attendeeType & """,""updateSent"":false,""warning"":""Attendee added but could not resolve for invite""}"
            Set fwdMail = Nothing
            Exit Function
        End If

        fwdMail.Send

        If Err.Number <> 0 Then
            AddAttendeeToMeeting = "{""success"":true,""attendeeAdded"":""" & EscapeJSON(attendeeEmail) & """,""resolvedName"":""" & EscapeJSON(resolvedName) & """,""type"":""" & attendeeType & """,""updateSent"":false,""warning"":""Attendee added but invite send failed: " & EscapeJSON(Err.Description) & """}"
            Set fwdMail = Nothing
            Exit Function
        End If

        Set fwdMail = Nothing
    End If

    ' Build success response
    json = "{"
    json = json & """success"":true,"
    json = json & """attendeeAdded"":""" & EscapeJSON(attendeeEmail) & ""","
    json = json & """resolvedName"":""" & EscapeJSON(resolvedName) & ""","
    json = json & """type"":""" & attendeeType & ""","
    json = json & """updateSent"":" & LCase(CStr(sendUpdate))
    json = json & "}"

    AddAttendeeToMeeting = json

    ' Cleanup
    Set recip = Nothing
    Set appt = Nothing
    Set ns = Nothing
    Set outlookApp = Nothing
End Function

' Run main
Main
