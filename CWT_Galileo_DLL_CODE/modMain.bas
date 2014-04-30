Attribute VB_Name = "modMain"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" _
                Alias "GetPrivateProfileStringA" _
                (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
'Public gstrFaresDBSource As String
'Public gstrFQDBSource As String
'Public gstrTProDBSource As String
'Public gstrTProLookupDBSource As String
Public gstrDBInitPath As String
Public gstrAgcyCountryCode As String
Public gstrAgcyCurrencyCode As String
Public GstrAgcyCityCode As String
Public gstrAgcyAirportCode As String
Public gstrAgcyPhone As String
Public gstrPFCode As String
Public GstrPCC As String
Public gstrHQPCC As String
Public gstrPFPCC As String
Public gbolSwitchAcc As Boolean
Public gbolSvcBur As Boolean
Public gbolTEErrorLog As Boolean
Public gbolAgentPCCLog As Boolean
Public gdbConn As ADODB.Connection
Public gstrConn As String
Public gobjHost As GalileoHost
Public gobjLog As CWT_AppLog.AppLog
Public Const gstrID As String = "<Application><VendorId>XMDL</VendorId><VendorType>G</VendorType><SourceId>CARWPR</SourceId><SourceType>G</SourceType></Application>"
Public Const CONNECTION_FAIL As String = "Unable to connect to GDS"
Public Enum SQLType
    Select_ = 0
    Insert_ = 1
    Update_ = 2
    Delete_ = 3
End Enum
Public Const CstrErrorAdvice = "Please report this to your Systems Administrator"
Public Declare Function WNetGetUser& Lib "Mpr" Alias "WNetGetUserA" (lpname As Any, ByVal lpUserName$, lpnLength&)
Public gVPMDIHwnd As Long


Public Function getConnStr(ByRef dbconn As ADODB.Connection) As Boolean
    Dim INITFILE As String
    Dim fnum As Integer
    Dim textline, connStr As String
    
    On Error GoTo conn_err
    connStr = ""
    INITFILE = gstrDBInitPath
    
    fnum = FreeFile
    Open INITFILE For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, textline
        Select Case UCase(Left(textline, InStr(textline, "=") - 1))
            Case "SERVER":
                connStr = connStr & IIf(connStr <> "", ";", "") & textline
            Case "DATABASE":
                connStr = connStr & IIf(connStr <> "", ";", "") & textline
            Case "UID":
                connStr = connStr & IIf(connStr <> "", ";", "") & textline
            Case "PWD":
                connStr = connStr & IIf(connStr <> "", ";", "") & textline
        End Select
    Loop
    Close #fnum
    
    connStr = "Driver={SQL Server};" & connStr
    
    dbconn.ConnectionString = connStr
    getConnStr = True
    Exit Function
    
conn_err:
   
    getConnStr = False

End Function

Private Sub getErrorConfig()
      'Added on 9/6/2005: Log TerminalEntry Errors to Database
Dim rslog As ADODB.Recordset

     Set rslog = gdbConn.Execute("select DbErrorLog,PROCTOLOG from tblTEErrorConfig where ProcToLog='TerminalEntry' OR ProcToLog='AgentPCC' ")
    While Not rslog.EOF
    If rslog!ProcToLog = "TerminalEntry" Then
        gbolTEErrorLog = rslog!DbErrorLog
    End If
    If rslog!ProcToLog = "AgentPCC" Then
        gbolAgentPCCLog = rslog!DbErrorLog
    End If
    rslog.MoveNext
    Wend
End Sub
Public Sub pGetNewStartupValues()
Dim rsConfig As ADODB.Recordset
Dim STRSQL As String
Dim strDefault As String
Dim strMsg As String
strDefault = "0000"

If pSetGlobals(gobjHost.AgentDIV) = False Then
    If pSetGlobals(strDefault) = False Then
        'MsgBox "Unable to retrieve system configuration setting!" & vbCrLf & vbCrLf _
        & "Please ensure that you are signon to GDS or the Agent configurations are setup in" & vbCrLf _
        & "TProData.  If you need assistance," & vbCrLf _
        & "please contact your system administrator!", vbApplicationModal + vbCritical + vbOKOnly
        strMsg = "Unable to retrieve system configuration setting!" & vbCrLf & vbCrLf _
        & "Please ensure that you are signon to GDS or the Agent configurations are setup in" & vbCrLf _
        & "TProData.  If you need assistance," & vbCrLf _
        & "please contact your system administrator!"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End If
End If

End Sub
Public Function pSetGlobals(div As String) As Boolean
Dim rsConfig As ADODB.Recordset
Dim STRSQL As String


STRSQL = "SELECT * FROM tblconfiguration where DIV='" & div & "'"
Set rsConfig = gdbConn.Execute(STRSQL)
If Not rsConfig.EOF Then
        If Not IsNull(rsConfig!CountryCode) Then gstrAgcyCountryCode = UCase(rsConfig!CountryCode)
        If Not IsNull(rsConfig!CurrencyCode) Then gstrAgcyCurrencyCode = UCase(rsConfig!CurrencyCode)
        If Not IsNull(rsConfig!CityCode) Then GstrAgcyCityCode = UCase(rsConfig!CityCode)
        If Not IsNull(rsConfig!AirportCode) Then gstrAgcyAirportCode = UCase(rsConfig!AirportCode)
        If Not IsNull(rsConfig!Tel) Then gstrAgcyPhone = UCase(rsConfig!Tel)
        If Not IsNull(rsConfig!PFCode) Then gstrPFCode = UCase(rsConfig!PFCode)
        If Not IsNull(rsConfig!PFPCC) Then gstrPFPCC = UCase(rsConfig!PFPCC)
        If Not IsNull(rsConfig!BKPCC) Then GstrPCC = UCase(rsConfig!BKPCC)
        gstrHQPCC = GstrPCC
        pSetGlobals = True
End If

End Function

Public Sub GetALLPath()
   Dim strPath As String
   Dim strINIFile As String
   Dim i As Integer
   Dim j As Integer
   Dim strCountry As String
   Dim strMsg As String
   
   'strINIFile = GetSetting("CWTAPP", "APPINI", "FileLoc", "")
   'If strINIFile = "" Then
   '   MsgBox "INI file location info not in registry"
   '   Exit Sub
   'End If
   strINIFile = App.Path & "\cwtapplication.ini"
   If Dir(strINIFile) = "" Then
      'MsgBox "INI file: " & strINIFile & " not found."
      strMsg = "INI file: " & strINIFile & " not found."
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      Exit Sub
   End If
   
   'Tpro DB
   'gstrConn = GetFromINI("TPro", "ConnectionString", strINIFile)
   'gstrConn = decrypt(gstrConn, gintKey)
   gstrConn = GetFromINI("CWTApplication", "ConnectionString", strINIFile)
   gstrConn = decrypt(gstrConn, gintKey)
   openDatabase
   strPath = GetOU(getWinLogon)
   i = InStr(1, strPath, "\")
   j = InStr(i + 1, strPath, "\")
   i = InStr(j + 1, strPath, "\")
   strCountry = Mid(strPath, j + 1, i - j - 1)
 
   If strCountry <> "SG" And strCountry <> "HK" Then
      strCountry = getOthOU(strCountry)
   End If
   gstrConn = getOption("TPro" & strCountry, "ConnectionString", True)
   
   If gdbConn.State = 1 Then
      gdbConn.Close
   End If
      
End Sub

Public Function GetFromINI(Section As String, key As String, Directory As String) As String

 Dim strBuffer As String

 strBuffer = String(750, Chr(0))

 key$ = LCase$(key$)

 GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal key$, "", strBuffer, Len(strBuffer), Directory$))

End Function
Public Function getOption(ByVal strKey, ByVal strType, Optional bolDecrypt As Boolean) As String
Dim rsOption As ADODB.Recordset
Dim STRSQL As String
STRSQL = "Select optionValue from tblOptions Where optionKey='" & strKey & "' AND Type='" & strType & "'"
RunSQLCommand SQLType.Select_, STRSQL, gdbConn, rsOption
If rsOption.EOF = False Then
   If bolDecrypt = False Then
      getOption = Trim(rsOption!optionValue & "")
   Else
      getOption = decrypt(Trim(rsOption!optionValue & ""), gintKey)
   End If
Else
  getOption = ""
End If
End Function

Public Function RunSQLCommand(CType As Integer, STRSQL As String, Conn As ADODB.Connection, Optional ByRef rs As ADODB.Recordset) As Boolean
   Dim intI As Integer
   Dim bolExeSql As Boolean
   Dim strErrDesc As String
   Dim strErrNum As Long
      
   intI = 0
   bolExeSql = False
   Do
     strErrDesc = ""
     strErrNum = 0
     
     If intI <> 0 Then
        Sleep (5000)
     End If
     intI = intI + 1
     
     If CType = SQLType.Select_ Then
        bolExeSql = ExeSQLCommand(CType, STRSQL, Conn, strErrNum, strErrDesc, rs)
     Else
        bolExeSql = ExeSQLCommand(CType, STRSQL, Conn, strErrNum, strErrDesc)
     End If
   Loop Until bolExeSql = True Or intI = 5
   If bolExeSql = False And intI >= 5 Then
      RunSQLCommand = False
   ElseIf bolExeSql = True Then
      RunSQLCommand = True
   End If
End Function

Public Function ExeSQLCommand(CType As Integer, STRSQL As String, Conn As ADODB.Connection, ByRef ErrNum As Long, ByRef ErrDesc As String, Optional ByRef rs As ADODB.Recordset) As Boolean
   On Error GoTo err1
   
   Select Case CType
      Case SQLType.Insert_, SQLType.Update_, SQLType.Delete_
         Conn.Execute STRSQL
      Case SQLType.Select_
         DoEvents
         Set rs = Conn.Execute(STRSQL)
   End Select
   ExeSQLCommand = True
   Exit Function
err1:
   ExeSQLCommand = False
   ErrNum = Err.Number
   ErrDesc = Err.Description
End Function

Private Sub openDatabase()
Dim strMsg As String

On Error GoTo ErrorProc
If gstrConn <> "" Then
    'Open Login database connection
    gdbConn.ConnectionString = gstrConn
    gdbConn.open
Else
    'MsgBox "Error in database connect string!", vbCritical
    strMsg = "Error in database connect string!"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End If
Exit Sub
ErrorProc:
Select Case Err.Number
Case 3044, 3024
    'MsgBox "The database cannot be found" & vbCrLf & _
            "Error Code: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            CstrErrorAdvice, vbOKOnly + vbCritical
    strMsg = "The database cannot be found" & vbCrLf & _
            "Error Code: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            CstrErrorAdvice
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
Case Else
    'MsgBox "There is a problem with the database" & vbCrLf & _
            "Error Code: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            CstrErrorAdvice, vbOKOnly + vbCritical
    strMsg = "There is a problem with the database" & vbCrLf & _
            "Error Code: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
            CstrErrorAdvice
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End Select
End Sub

Private Function GetOU(WinLogin As String) As String

    Dim objConnection As New ADODB.Connection
    Dim objCommand As New ADODB.Command
    Dim objRecordSet As ADODB.Recordset
    Dim strOU As String
    Dim strPath As String
    Dim a As Variant
    Dim arrPath() As String

    objConnection.Provider = "ADsDSOObject"
    objConnection.open "Active Directory Provider"
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = 2
    objCommand.CommandText = "SELECT adsPath FROM 'LDAP://dc=auas,dc=carlson,dc=com' WHERE objectCategory='user' AND sAMAccountName='" & WinLogin & "'"

    Set objRecordSet = objCommand.Execute
    strOU = ""
    If objRecordSet.EOF = False Then
        objRecordSet.MoveFirst
        Do Until objRecordSet.EOF
            strPath = objRecordSet.fields("adsPath").Value
            arrPath = Split(strPath, ",")
            strOU = ""
            For Each a In arrPath
            If Left(a, 2) = "OU" Then
                strOU = "\" & Right(a, Len(a) - 3) & strOU
            End If
            Next
            objRecordSet.MoveNext
        Loop
        End If
    GetOU = strOU
End Function

Private Function getWinLogon() As String

Dim username As String
Dim cbusername As Long
Dim ret As Integer

username = Space(256)
cbusername = Len(username)
ret = WNetGetUser(ByVal 0&, username, cbusername)
If ret = 0 Then
   ' Success - strip off the null.
   username = Left(username, InStr(username, Chr(0)) - 1)
Else
   username = ""
End If
getWinLogon = username

End Function

Public Function getOthOU(ByVal strOUPath As String) As String
Dim rsOUPath As ADODB.Recordset
Dim STRSQL As String
STRSQL = "Select Country from tblOUGroup Where OUPath = '" & strOUPath & "'"
RunSQLCommand SQLType.Select_, STRSQL, gdbConn, rsOUPath
If rsOUPath.EOF = False Then
   getOthOU = Trim(rsOUPath!Country & "")
Else
   getOthOU = ""
End If
End Function

'JiYong - V1.2.1 (Urgent Fix) 201001102 - IR 6 - Pick Up Rule ID in Fare Rules
'Check for alphanumeric
Public Function IsAlphaNum(ByVal sString As String) As Boolean
    If Not sString Like "*[!0-9A-Za-z]*" Then IsAlphaNum = True
End Function



