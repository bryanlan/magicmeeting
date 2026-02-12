' getEmailContent.vbs - Gets full email content by EntryID
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    ' Get command line argument
    Dim emailId
    emailId = GetArgument("emailId")

    ' Require emailId
    RequireArgument "emailId"

    ' Get email content
    Dim result
    result = GetEmailById(emailId)

    ' Output result
    OutputSuccess result
End Sub

' Gets email content by EntryID
Function GetEmailById(emailId)
    On Error Resume Next

    Dim outlookApp, namespace, item, json

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()

    ' Get MAPI namespace
    Set namespace = outlookApp.GetNamespace("MAPI")

    If Err.Number <> 0 Then
        OutputError "Failed to get MAPI namespace: " & Err.Description
        WScript.Quit 1
    End If

    ' Get item by EntryID
    Set item = namespace.GetItemFromID(emailId)

    If Err.Number <> 0 Then
        OutputError "Failed to find email with ID: " & Err.Description
        WScript.Quit 1
    End If

    If item Is Nothing Then
        OutputError "Email not found with specified ID"
        WScript.Quit 1
    End If

    ' Build JSON response
    json = "{"
    json = json & """id"":""" & EscapeJSON(item.EntryID) & ""","
    json = json & """subject"":""" & EscapeJSON(item.Subject) & ""","
    json = json & """from"":""" & EscapeJSON(item.SenderName) & ""","
    json = json & """fromEmail"":""" & EscapeJSON(item.SenderEmailAddress) & ""","
    json = json & """to"":""" & EscapeJSON(item.To) & ""","
    json = json & """cc"":""" & EscapeJSON(item.CC) & ""","
    json = json & """received"":""" & FormatDateTime(item.ReceivedTime) & ""","
    json = json & """body"":""" & EscapeJSON(item.Body) & ""","
    json = json & """htmlBody"":""" & EscapeJSON(item.HTMLBody) & """"
    json = json & "}"

    If Err.Number <> 0 Then
        OutputError "Failed to get email content: " & Err.Description
        WScript.Quit 1
    End If

    GetEmailById = json

    ' Clean up
    Set item = Nothing
    Set namespace = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
