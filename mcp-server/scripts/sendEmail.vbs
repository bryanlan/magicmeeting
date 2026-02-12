' sendEmail.vbs - Sends HTML email via Outlook COM
Option Explicit

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Constants for Outlook
Const olMailItem = 0
Const olFormatHTML = 2
Const olFormatPlain = 1

' Decodes base64 string
Function Base64Decode(base64Str)
    On Error Resume Next
    Dim dom, node

    If base64Str = "" Then
        Base64Decode = ""
        Exit Function
    End If

    Set dom = CreateObject("MSXML2.DOMDocument")
    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to create DOM: " & Err.Description
        WScript.Quit 1
    End If

    Set node = dom.createElement("tmp")
    node.DataType = "bin.base64"
    node.Text = base64Str

    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to set base64 text: " & Err.Description
        WScript.Quit 1
    End If

    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to create stream: " & Err.Description
        WScript.Quit 1
    End If

    stream.Type = 1 ' Binary
    stream.Open
    stream.Write node.NodeTypedValue
    stream.Position = 0
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    Base64Decode = stream.ReadText
    stream.Close

    If Err.Number <> 0 Then
        WScript.Echo ERROR_PREFIX & "Failed to decode base64: " & Err.Description
        WScript.Quit 1
    End If

    Set stream = Nothing
    Set node = Nothing
    Set dom = Nothing
End Function

' Main function
Sub Main()
    ' Get command line arguments
    Dim toAddresses, ccAddresses, subject, isHtmlStr
    Dim isHtml, body

    toAddresses = GetArgument("to")
    ccAddresses = GetArgument("cc")
    subject = GetArgument("subject")
    isHtmlStr = GetArgument("isHtml")

    ' Require to and subject
    RequireArgument "to"
    RequireArgument "subject"

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

    ' Parse isHtml (default true)
    If isHtmlStr = "" Or LCase(isHtmlStr) = "true" Then
        isHtml = True
    Else
        isHtml = False
    End If

    ' Send the email
    Dim result
    result = SendOutlookEmail(toAddresses, ccAddresses, subject, body, isHtml)

    ' Output result
    OutputSuccess result
End Sub

' Sends an email using Outlook
Function SendOutlookEmail(toAddresses, ccAddresses, subject, body, isHtml)
    On Error Resume Next

    Dim outlookApp, mailItem, json

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()

    ' Create new mail item
    Set mailItem = outlookApp.CreateItem(olMailItem)

    If Err.Number <> 0 Then
        OutputError "Failed to create mail item: " & Err.Description
        WScript.Quit 1
    End If

    ' Set recipients
    mailItem.To = toAddresses

    If ccAddresses <> "" Then
        mailItem.CC = ccAddresses
    End If

    ' Set subject
    mailItem.Subject = subject

    ' Set body (HTML or plain text)
    If isHtml Then
        mailItem.BodyFormat = olFormatHTML
        mailItem.HTMLBody = body
    Else
        mailItem.BodyFormat = olFormatPlain
        mailItem.Body = body
    End If

    ' Send the email
    mailItem.Send

    If Err.Number <> 0 Then
        OutputError "Failed to send email: " & Err.Description
        WScript.Quit 1
    End If

    ' Build success JSON
    json = "{"
    json = json & """success"":true,"
    json = json & """to"":""" & EscapeJSON(toAddresses) & ""","
    If ccAddresses <> "" Then
        json = json & """cc"":""" & EscapeJSON(ccAddresses) & ""","
    End If
    json = json & """subject"":""" & EscapeJSON(subject) & ""","
    json = json & """isHtml"":" & LCase(CStr(isHtml))
    json = json & "}"

    SendOutlookEmail = json

    ' Clean up
    Set mailItem = Nothing
    Set outlookApp = Nothing
End Function

' Run the main function
Main
