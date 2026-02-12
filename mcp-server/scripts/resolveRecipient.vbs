' resolveRecipient.vbs - Resolve a name/alias/email to a GAL entry
Option Explicit

' OlAddressEntryUserType constants
Const olExchangeUserAddressEntry = 0
Const olExchangeDistributionListAddressEntry = 1
Const olExchangeRemoteUserAddressEntry = 2
Const olExchangeAgentAddressEntry = 3
Const olExchangeOrganizationAddressEntry = 4
Const olExchangePublicFolderAddressEntry = 5
Const olOutlookContactAddressEntry = 10
Const olOutlookDistributionListAddressEntry = 11
Const olLdapAddressEntry = 20
Const olSmtpAddressEntry = 30
Const olOtherAddressEntry = 40

' OlDisplayType constants
Const olUser = 0
Const olDistList = 1
Const olForum = 2
Const olAgent = 3
Const olOrganization = 4
Const olPrivateDistList = 5
Const olRemoteUser = 6

' Include utility functions
Dim fso, scriptDir
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ExecuteGlobal fso.OpenTextFile(fso.BuildPath(scriptDir, "utils.vbs"), 1).ReadAll

' Main function
Sub Main()
    Dim query
    query = GetArgument("query")
    RequireArgument "query"

    Dim result
    result = ResolveRecipientToJSON(query)

    OutputSuccess result
End Sub

' Resolve a recipient and return JSON with details
Function ResolveRecipientToJSON(query)
    On Error Resume Next

    Dim outlookApp, ns, recipient, ae
    Dim json, entryType, primarySmtp, displayName, alias

    ' Create Outlook application
    Set outlookApp = CreateOutlookApplication()
    Set ns = outlookApp.GetNamespace("MAPI")

    ' Create and resolve the recipient
    Set recipient = ns.CreateRecipient(query)

    If recipient Is Nothing Then
        ResolveRecipientToJSON = "{""resolved"":false,""error"":""Failed to create recipient""}"
        Exit Function
    End If

    ' Try to resolve
    If Not recipient.Resolve Then
        ResolveRecipientToJSON = "{""resolved"":false,""error"":""Could not resolve: " & EscapeJSON(query) & """}"
        Exit Function
    End If

    ' Get the AddressEntry
    Set ae = recipient.AddressEntry

    If ae Is Nothing Then
        ResolveRecipientToJSON = "{""resolved"":false,""error"":""No AddressEntry found""}"
        Exit Function
    End If

    ' Determine entry type
    entryType = GetEntryTypeName(ae.AddressEntryUserType)

    ' Get display name
    displayName = ae.Name

    ' Get primary SMTP address based on type
    primarySmtp = GetPrimarySmtp(ae)

    ' Get alias if available
    alias = GetAlias(ae)

    ' Build JSON response
    json = "{"
    json = json & """resolved"":true,"
    json = json & """name"":""" & EscapeJSON(displayName) & ""","
    json = json & """email"":""" & EscapeJSON(primarySmtp) & ""","
    json = json & """type"":""" & entryType & ""","
    json = json & """typeCode"":" & ae.AddressEntryUserType & ","
    json = json & """displayType"":" & ae.DisplayType & ","
    json = json & """isDistributionList"":" & LCase(CStr(IsDistributionList(ae))) & ","

    If alias <> "" Then
        json = json & """alias"":""" & EscapeJSON(alias) & ""","
    End If

    json = json & """id"":""" & EscapeJSON(ae.ID) & """"
    json = json & "}"

    ResolveRecipientToJSON = json

    ' Clean up
    Set ae = Nothing
    Set recipient = Nothing
    Set ns = Nothing
    Set outlookApp = Nothing
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
        Case olExchangeAgentAddressEntry
            GetEntryTypeName = "ExchangeAgent"
        Case olExchangeOrganizationAddressEntry
            GetEntryTypeName = "ExchangeOrganization"
        Case olExchangePublicFolderAddressEntry
            GetEntryTypeName = "ExchangePublicFolder"
        Case olOutlookContactAddressEntry
            GetEntryTypeName = "OutlookContact"
        Case olOutlookDistributionListAddressEntry
            GetEntryTypeName = "OutlookDistributionList"
        Case olLdapAddressEntry
            GetEntryTypeName = "LDAP"
        Case olSmtpAddressEntry
            GetEntryTypeName = "SMTP"
        Case Else
            GetEntryTypeName = "Other"
    End Select
End Function

' Check if entry is a distribution list
Function IsDistributionList(ae)
    On Error Resume Next
    IsDistributionList = (ae.AddressEntryUserType = olExchangeDistributionListAddressEntry) Or _
                         (ae.AddressEntryUserType = olOutlookDistributionListAddressEntry) Or _
                         (ae.DisplayType = olDistList) Or _
                         (ae.DisplayType = olPrivateDistList)
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
            ' Fallback - may be X500 or other format
            GetPrimarySmtp = ae.Address
    End Select

    ' If still empty, try PropertyAccessor for PR_SMTP_ADDRESS
    If GetPrimarySmtp = "" Or Left(GetPrimarySmtp, 1) = "/" Then
        Err.Clear
        Dim pa
        Set pa = ae.PropertyAccessor
        If Not pa Is Nothing Then
            ' PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
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
