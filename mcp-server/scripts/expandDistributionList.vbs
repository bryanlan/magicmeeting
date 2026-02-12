' expandDistributionList.vbs - Expand a distribution list to get members
Option Explicit

' OlAddressEntryUserType constants
Const olExchangeUserAddressEntry = 0
Const olExchangeDistributionListAddressEntry = 1
Const olExchangeRemoteUserAddressEntry = 2
Const olOutlookContactAddressEntry = 10
Const olOutlookDistributionListAddressEntry = 11
Const olSmtpAddressEntry = 30

' OlDisplayType constants
Const olUser = 0
Const olDistList = 1
Const olPrivateDistList = 5
Const olRemoteUser = 6

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Global visited dictionary for cycle detection
Dim gVisited
Set gVisited = CreateObject("Scripting.Dictionary")

' Main function
Sub Main()
    Dim dlName, recursive, maxDepth
    dlName = GetArgument("name")
    recursive = (LCase(GetArgument("recursive")) = "true")
    maxDepth = GetArgument("maxDepth")

    RequireArgument "name"

    If maxDepth = "" Then
        maxDepth = 3  ' Default max recursion depth
    Else
        maxDepth = CInt(maxDepth)
    End If

    Dim result
    result = ExpandDLToJSON(dlName, recursive, maxDepth)

    OutputSuccess result
End Sub

' Expand a distribution list and return JSON
Function ExpandDLToJSON(dlName, recursive, maxDepth)
    On Error Resume Next

    Dim outlookApp, ns, recipient, ae
    Dim json

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    Set ns = outlookApp.GetNamespace("MAPI")

    ' Create and resolve the recipient
    Set recipient = ns.CreateRecipient(dlName)

    If recipient Is Nothing Then
        ExpandDLToJSON = "{""success"":false,""error"":""Failed to create recipient""}"
        Exit Function
    End If

    ' Try to resolve
    If Not recipient.Resolve Then
        ExpandDLToJSON = "{""success"":false,""error"":""Could not resolve: " & EscapeJSON(dlName) & """}"
        Exit Function
    End If

    ' Get the AddressEntry
    Set ae = recipient.AddressEntry

    If ae Is Nothing Then
        ExpandDLToJSON = "{""success"":false,""error"":""No AddressEntry found""}"
        Exit Function
    End If

    ' Check if it's a distribution list
    If Not IsDistributionList(ae) Then
        ExpandDLToJSON = "{""success"":false,""error"":""Not a distribution list: " & EscapeJSON(ae.Name) & " (type: " & ae.AddressEntryUserType & ")""}"
        Exit Function
    End If

    ' Get DL details
    Dim dlEmail, dlAlias
    dlEmail = GetPrimarySmtp(ae)
    dlAlias = GetAlias(ae)

    ' Build JSON response
    json = "{"
    json = json & """success"":true,"
    json = json & """name"":""" & EscapeJSON(ae.Name) & ""","
    json = json & """email"":""" & EscapeJSON(dlEmail) & ""","

    If dlAlias <> "" Then
        json = json & """alias"":""" & EscapeJSON(dlAlias) & ""","
    End If

    json = json & """recursive"":" & LCase(CStr(recursive)) & ","
    json = json & """maxDepth"":" & maxDepth & ","

    ' Get members
    Dim membersJson
    membersJson = GetMembersJSON(ae, recursive, 0, maxDepth)
    json = json & """members"":" & membersJson

    json = json & "}"

    ExpandDLToJSON = json

    ' Clean up
    Set ae = Nothing
    Set recipient = Nothing
    Set ns = Nothing
    Set outlookApp = Nothing
End Function

' Get members of a DL as JSON array
Function GetMembersJSON(dlAE, recursive, depth, maxDepth)
    On Error Resume Next

    Dim json, exDL, members, i, m
    Dim memberName, memberEmail, memberType, isNested

    json = "["

    ' Check for cycles
    Dim key
    key = dlAE.ID
    If gVisited.Exists(key) Then
        GetMembersJSON = "[]"
        Exit Function
    End If
    gVisited.Add key, True

    ' Get ExchangeDistributionList object
    Set exDL = dlAE.GetExchangeDistributionList()

    If exDL Is Nothing Then
        ' Try using Members property directly (for Outlook DLs)
        Set members = dlAE.Members
    Else
        ' Use Exchange DL method
        Set members = exDL.GetExchangeDistributionListMembers()
    End If

    If members Is Nothing Then
        GetMembersJSON = "[]"
        Exit Function
    End If

    Dim first
    first = True

    For i = 1 To members.Count
        Set m = members.Item(i)

        memberName = m.Name
        memberEmail = GetPrimarySmtp(m)
        memberType = GetEntryTypeName(m.AddressEntryUserType)
        isNested = IsDistributionList(m)

        If Not first Then json = json & ","
        first = False

        json = json & "{"
        json = json & """name"":""" & EscapeJSON(memberName) & ""","
        json = json & """email"":""" & EscapeJSON(memberEmail) & ""","
        json = json & """type"":""" & memberType & ""","
        json = json & """isDistributionList"":" & LCase(CStr(isNested))

        ' Recursively expand nested DLs if requested
        If isNested And recursive And (depth < maxDepth) Then
            Dim nestedMembers
            nestedMembers = GetMembersJSON(m, recursive, depth + 1, maxDepth)
            json = json & ",""members"":" & nestedMembers
        End If

        json = json & "}"
    Next

    json = json & "]"

    GetMembersJSON = json
End Function

' Check if entry is a distribution list
Function IsDistributionList(ae)
    On Error Resume Next
    IsDistributionList = (ae.AddressEntryUserType = olExchangeDistributionListAddressEntry) Or _
                         (ae.AddressEntryUserType = olOutlookDistributionListAddressEntry) Or _
                         (ae.DisplayType = olDistList) Or _
                         (ae.DisplayType = olPrivateDistList)
End Function

' Get human-readable type name
Function GetEntryTypeName(typeCode)
    Select Case typeCode
        Case olExchangeUserAddressEntry
            GetEntryTypeName = "ExchangeUser"
        Case olExchangeDistributionListAddressEntry
            GetEntryTypeName = "ExchangeDistributionList"
        Case olExchangeRemoteUserAddressEntry
            GetEntryTypeName = "ExchangeRemoteUser"
        Case olOutlookContactAddressEntry
            GetEntryTypeName = "OutlookContact"
        Case olOutlookDistributionListAddressEntry
            GetEntryTypeName = "OutlookDistributionList"
        Case olSmtpAddressEntry
            GetEntryTypeName = "SMTP"
        Case Else
            GetEntryTypeName = "Other"
    End Select
End Function

' Get primary SMTP address
Function GetPrimarySmtp(ae)
    On Error Resume Next
    GetPrimarySmtp = ""

    Select Case ae.AddressEntryUserType
        Case olExchangeUserAddressEntry
            Dim exUser
            Set exUser = ae.GetExchangeUser()
            If Not exUser Is Nothing Then
                GetPrimarySmtp = exUser.PrimarySmtpAddress
            End If
        Case olExchangeDistributionListAddressEntry
            Dim exDL
            Set exDL = ae.GetExchangeDistributionList()
            If Not exDL Is Nothing Then
                GetPrimarySmtp = exDL.PrimarySmtpAddress
            End If
        Case olSmtpAddressEntry
            GetPrimarySmtp = ae.Address
        Case Else
            GetPrimarySmtp = ae.Address
    End Select

    ' If still empty or X500, try PropertyAccessor
    If GetPrimarySmtp = "" Or Left(GetPrimarySmtp, 1) = "/" Then
        Err.Clear
        Dim pa
        Set pa = ae.PropertyAccessor
        If Not pa Is Nothing Then
            GetPrimarySmtp = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
            If Err.Number <> 0 Then
                Err.Clear
                GetPrimarySmtp = ae.Address
            End If
        End If
    End If
End Function

' Get alias if available
Function GetAlias(ae)
    On Error Resume Next
    GetAlias = ""

    Select Case ae.AddressEntryUserType
        Case olExchangeUserAddressEntry
            Dim exUser
            Set exUser = ae.GetExchangeUser()
            If Not exUser Is Nothing Then
                GetAlias = exUser.Alias
            End If
        Case olExchangeDistributionListAddressEntry
            Dim exDL
            Set exDL = ae.GetExchangeDistributionList()
            If Not exDL Is Nothing Then
                GetAlias = exDL.Alias
            End If
    End Select
End Function

' Run the main function
Main
