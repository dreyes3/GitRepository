Attribute VB_Name = "Startup"
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOACTIVATE = &H10
Public Sub ChangeTBBack(hWnd As Long, PNewBack As Long)
    Dim lTBWnd As Long
    'Change the background color of toolbar
    lTBWnd = FindWindowEx(2229850, 0, "ToolbarWindow32", vbNullString) 'Find Hwnd first
    DeleteObject SetClassLong(lTBWnd, GCL_HBRBACKGROUND, PNewBack)
End Sub

Public Sub Main()
    Dim VPHwnd As String
    Dim VPToolbarHwnd As Long
    Dim pID As Long
    Dim intRetry As Integer
    Dim old_parent As Long
    Dim strReaponse As String
    Dim wp As WINDOWPLACEMENT
    Dim bolPNRExist As Boolean
    
    'Get the handler of Galileo Desktop.
    VPHwnd = IsAppRunning("Galileo Desktop")
    If VPHwnd = 0 Then
       'Start Galileo Desktop if handle is 0 as the application is currently not running.
       pID = Shell("C:\fp\swdir\Viewpoint.exe", 3)
       If pID = 0 Then
          MsgBox "Unable to startup Galileo Desktop!", vbCritical, "CWT Desktop - Startup Error"
          End
       Else
Retry:
          Sleep (5000)
          VPHwnd = IsAppRunning("Galileo Desktop")
          If VPHwnd = 0 Then
             'Retries 3 times if still can't get the handler of Galileo Desktop
             If intRetry < 2 Then
                intRetry = intRetry + 1
                Sleep (2000)
                GoTo Retry
             Else
                MsgBox "Unable to startup CWT Desktop!", vbCritical, "CWT Desktop - Startup Error"
                End
             End If
          End If
       End If
    End If
    
    pOpenDatabase
    pSetGlobalObjects
    
    'Find the handler of Galileo Desktop Custom Toolbar to place the TPro toolbar
    VPToolbarHwnd = FindWindowEx(VPHwnd, ByVal 0&, "AfxControlBar70", vbNullString)
    If VPToolbarHwnd = 0 Then
       'Custom Toolbar must be activated in Galileo Desktop
        MsgBox "Custom toolbar must be activated in Galileo Desktop!", vbCritical, "CWT Desktop - Startup Error"
        End
    Else
        'Show up toolbar and embed it into Galileo Desktop
        frmBars.Show
        old_parent = SetParent(frmBars.hWnd, VPToolbarHwnd)
        frmBars.Move 0, 0
        
    End If
    
    glngTargetHwnd = VPHwnd
    'Find the handler of MDIClient in Galileo Desktop to place the panel
    gVPMDIHwnd = FindWindowEx(glngTargetHwnd, ByVal 0&, "MDIClient", vbNullString)
    
    If gVPMDIHwnd <> 0 Then
       Load frmSideBar
       old_parent = SetParent(frmSideBar.hWnd, gVPMDIHwnd)
       frmSideBar.Show
       frmSideBar.Move 0, 0

       'Wrap the text in treeview (Important: Disable these 2 lines in debug mode)
       SetWindowLong frmSideBar.treeViewTraveller.hWnd, GWL_STYLE, GetWindowLong(frmSideBar.treeViewTraveller.hWnd, GWL_STYLE) Or TVS_NOTOOLTIPS Or TVS_HASLINES
       OldProc = SetWindowLong(frmSideBar.hWnd, GWL_WNDPROC, AddressOf WindowProc)
              
       If checkSignOn(True) = True Then
       
          gobjPNR.LoadPNR
          strResponse = gobjHost.TerminalEntry("*R")
          If InStr(1, strResponse, "NO B.F. TO DISPLAY - CREATE OR RETRIEVE FIRST") = 0 Then
             'displayPNRinBar
             bolPNRExist = True
          End If
       End If
       
       'Get highlight keywords on sidebar from DB
       getKeywords
           
       If bolPNRExist Then displayPNRinBar
       'Minimized window and restore back to cater the color faded issue when 1st loaded
       wp.length = Len(wp)
       GetWindowPlacement VPHwnd, wp
       wp.showCmd = SW_SHOWMINIMIZED
       SetWindowPlacement VPHwnd, wp
       wp.showCmd = SW_SHOWMAXIMIZED
       SetWindowPlacement VPHwnd, wp
    Else
       'End application if can't find the handler of gVPMDIHwnd
        MsgBox "Handle of MDI for Galileo Desktop is not found!", vbCritical, "CWT Desktop - Startup Error"
        End
    End If

End Sub

Public Function IsAppRunning(sWindowName As String) As Long
    Dim hWnd As Long, hWndOffline As Long
    
    On Error GoTo IsAppRunning_Eh
    'get handle of the application
    'if handle is 0 the application is currently not running
    gstrTargetName = sWindowName
    glngTargetHwnd = 0
    
    ' Examine the window names.
    EnumWindows AddressOf WindowEnumerator, 0

    ' See if we got an hwnd.
    IsAppRunning = glngTargetHwnd
    Exit Function

IsAppRunning_Eh:
    Call ShowError(sWindowName, "IsAppRunning")
End Function

Public Function WindowEnumerator(ByVal app_hwnd As Long, ByVal lParam As Long) As Long
    Dim buf As String * 256
    Dim title As String
    Dim length As Long

    ' Get the window's title.
    length = GetWindowText(app_hwnd, buf, Len(buf))
    title = Left$(buf, length)

    ' See if the title contains the target.
    If InStr(title, gstrTargetName) > 0 Then
        ' Save the hwnd and end the enumeration.
        glngTargetHwnd = app_hwnd
        WindowEnumerator = False
    Else
        ' Continue the enumeration.
        WindowEnumerator = True
    End If
End Function

Public Function ShowError(sText As String, sProcName As String)
    'this function displays an error that occured
    Dim sMsg As String
    
    sMsg = "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & vbCrLf & Err.Description
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, sMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End Function

Public Sub pSetGlobalObjects()
    pConnectToHost frmDDEOwner.txtDDE
    Set gobjLog = New CWT_AppLog.AppLog
    gobjLog.OpenLog App.Path, App.EXEName, App.title, App.Major, App.Minor, App.Revision
    
    Set gobjHost = New CWT_Galileo3.GalileoHost
    gobjHost.pGetStartupValues gdbConn, gobjLog, gobjHost, gVPMDIHwnd
    Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.pGetStartupValues gdbConn, gobjLog, gobjHost, gVPMDIHwnd
    pGetStartupValues
    pGetAgencyDefaults

End Sub

Public Sub pConnectToHost(ByRef txtLink As TextBox)
Dim strMsg As String
Start:
    On Error Resume Next
    With txtLink
        .LinkTopic = "GalileoIWS|System"
        .LinkItem = "Sessions"
        .LinkMode = gintLINK_MANUAL
        .LinkRequest
        If Err.Number > 0 Then GoTo ErrProc
        gstrHostSession = .Text
        
        .LinkTopic = "GalileoIWS|VisiblePartition"
        .LinkMode = gintLINK_MANUAL
        .LinkTimeout = gintDDE_TIMEOUT
    End With
    Exit Sub
ErrProc:
    Select Case Err.Number
        Case 293
            Call pConnectToHost(frmDDEOwner.txtDDE)
            With frmDDEOwner.txtDDE
                 .LinkItem = "ClearWindow"
                 .Text = "1"
                 .LinkPoke
            End With
        Case 285 'Foreign application won't perform DDE method or operation
             AppActivate "Galileo Desktop"
             PressKey "R", , True
             GoTo Start
        Case Else
            strMsg = "Error " & Err.Number & vbCrLf & Err.Description
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End Select
    
End Sub

Public Sub pDisplayToFP(strEntry As String)
Dim strMsg As String
Start:
    On Error Resume Next
    With frmDDEOwner.txtDDE
        .LinkItem = "Transmit"
        .Text = strEntry
        .LinkPoke
        If Err.Number > 0 Then GoTo ErrProc
    End With
    Exit Sub
ErrProc:
    Select Case Err.Number
        Case 293
            Call pConnectToHost(frmDDEOwner.txtDDE)
            With frmDDEOwner.txtDDE
                 .LinkItem = "ClearWindow"
                 .Text = "1"
                 .LinkPoke
            End With
        Case 285 'Foreign application won't perform DDE method or operation
             AppActivate "Galileo Desktop"
             PressKey "R", , True
             GoTo Start
        Case Else
            strMsg = "Error " & Err.Number & vbCrLf & Err.Description
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End Select
End Sub

Public Sub displayPNRinBar()
    Dim intI As Integer
    Dim intJ As Integer
    Dim strTemp As String
    Dim bolMissingBizPN As Boolean
    Dim bolMissingMobPN As Boolean
    Dim bolExist As Boolean
    
    bolMissingBizPN = True
    bolMissingMobPN = True
    pDisplayToFP ("*R")
    frmSideBar.fraInfo.Caption = " PNR -> " & gobjPNR.RecLoc
    If gobjPNR.RecLoc <> "" Then
       bolExist = False
       For intJ = 0 To frmSideBar.cmbLocator.ListCount - 1
           If frmSideBar.cmbLocator.List(intJ) = gobjPNR.RecLoc Then
              bolExist = True
              Exit For
           End If
       Next
       If bolExist = False Then
          frmSideBar.cmbLocator.AddItem gobjPNR.RecLoc, 0
          frmSideBar.cmbLocator.List(0, 1) = gobjPNR.PassengerName(1).LastName & "/" & gobjPNR.PassengerName(1).FirstName
       End If
    End If
    With frmSideBar.treeViewTraveller
         .Nodes.Clear
         
         'Define all the categories in sidebar
         .Nodes.Add , , "TS", "TRAVELLER SUMMARY"
         If frmBars.sftTabs.Tabs.Current = 0 Then
           .Nodes.Add , , "AP", "AIR POLICY"
           .Nodes.Add , , "APRI", "AIR PRICING"
           .Nodes.Add , , "APRE", "AIR PREFERENCE"
         ElseIf frmBars.sftTabs.Tabs.Current = 1 Then
           .Nodes.Add , , "HP", "HOTEL POLICY & PREFERENCE"
         ElseIf frmBars.sftTabs.Tabs.Current = 2 Then
           .Nodes.Add , , "CP", "CAR POLICY & PREFERENCE"
         End If
         .Nodes.Add , , "VI", "VISA INFORMATION"
         .Nodes.Add , , "OS", "ERROR"
         For intI = 1 To .Nodes.Count
             .Nodes(intI).Bold = True
             .Nodes(intI).Expanded = True
         Next
         
         'Populate the child nodes for each category
         For intI = 1 To gobjPNR.PassengerCount
             .Nodes.Add "TS", tvwChild, , gobjPNR.PassengerName(intI).LastName & ", " & gobjPNR.PassengerName(intI).FirstName & IIf(gobjPNR.PassengerName(intI).PassengerType = "I", " (INFANT)", "")
         Next
                   
         For intI = 1 To gobjPNR.GeneralRemarkCount
             If frmBars.sftTabs.Tabs.Current = 0 And gobjPNR.GeneralRemark(intI).Qualifier = "*B" Then
                .Nodes.Add "AP", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
             ElseIf frmBars.sftTabs.Tabs.Current = 0 And gobjPNR.GeneralRemark(intI).Qualifier = "*Z" Then
                .Nodes.Add "APRI", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
             ElseIf gobjPNR.GeneralRemark(intI).Qualifier = "*G" Then
                 If InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "TRAVELER TYPE: ") > 0 Then
                    intJ = InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "TRAVELER TYPE: ")
                    .Nodes.Add "TS", tvwChild, , "TYPE: " & Mid(gobjPNR.GeneralRemark(intI).RemarkText, intJ + 15)
                 ElseIf frmBars.sftTabs.Tabs.Current = 0 And (InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "SEAT") > 0 _
                        Or InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "SPML") > 0 _
                        Or InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "SSR") > 0) Then
                    .Nodes.Add "APRE", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
                 End If
             ElseIf gobjPNR.GeneralRemark(intI).Qualifier = "*V" Then
                strTemp = Trim(gobjPNR.GeneralRemark(intI).RemarkText)
                strTemp = Replace(strTemp, "VISA:", "")
                strTemp = Replace(strTemp, "-EXP", " EXP")
                strTemp = Trim(strTemp)
                .Nodes.Add "VI", tvwChild, , strTemp
             ElseIf frmBars.sftTabs.Tabs.Current = 1 And gobjPNR.GeneralRemark(intI).Qualifier = "*H" Then
                .Nodes.Add "HP", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
             ElseIf frmBars.sftTabs.Tabs.Current = 2 And gobjPNR.GeneralRemark(intI).Qualifier = "*C" Then
                .Nodes.Add "CP", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
             End If
         Next
                  
         If gobjPNR.FOPType = "" Then
            .Nodes.Add "OS", tvwChild, , "MISSING FOP"
         End If
    
    End With
End Sub

Public Sub controlForms(ByVal bolEnable As Boolean)
       frmSideBar.Enabled = bolEnable
       frmBars.Enabled = bolEnable
End Sub

Public Sub pOpenDatabase()

    Set gdbConn = New ADODB.Connection
    Set gdbEitinConn = New ADODB.Connection
    GetAllPath
  
    If gstrConn <> "" Then
        'Open TPro database connection
        openDatabase gdbConn, gstrConn
    End If
    If gstrEitinConn <> "" Then
        'Open Eitin database connection
        openDatabase gdbEitinConn, gstrEitinConn
    End If
    If gstrEmail <> "" Then
        'Open DBMailServer database connection
        openDatabase gdbEmailConn, gstrEmail
    End If
End Sub

Public Sub GetAllPath()
   Dim strPath As String
   Dim strINIFile As String
   Dim i As Integer
   Dim j As Integer
   Dim strCountry As String
   Dim strMsg As String
   
   strPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
   strINIFile = strPath & "cwtapplication.ini"

   If Dir(strINIFile) = "" Then
      strMsg = "INI file: " & strINIFile & " not found."
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      End
   End If
   
   gstrConn = GetFromINI("CWTApplication", "ConnectionString", strINIFile)
   gstrConn = decrypt(gstrConn, gintKey)
   strPath = GetOU(getWinLogon)
   i = InStr(1, strPath, "\")
   j = InStr(i + 1, strPath, "\")
   i = InStr(j + 1, strPath, "\")
   strCountry = Mid(strPath, j + 1, i - j - 1)
   openDatabase gdbConn, gstrConn
   If strCountry <> "SG" And strCountry <> "HK" Then
      strCountry = getOthOU(strCountry)
   End If
   
   gstrConn = getOption("TPro" & strCountry, "ConnectionString", True)
   gstrEitinConn = getOption("HotelItin", "ConnectionString", True)
   gstrDespatch = getOption("Despatch", "ConnectionString", True)
   gstrDespatchExe = getOption("Despatch", "ApplicationPath")
   gstrEPRptLogin = getOption("EOReport", "ReportLogin", True)
   gstrEOPrtPwd = getOption("EOReport", "ReportPassword", True)
   gstrEORptDB = getOption("EOReport", "Database", True)
   gstrEORptServer = getOption("EOReport", "Server", True)
   gstrReportPath = getOption("EOReport", "ReportPath")
   gstrEmail = getOption("DBEmailServer", "ConnectionString", True)
  
    'get viewpt's window name
    getVPwindows
   
   
   If gdbConn.state = 1 Then
      gdbConn.Close
   End If
     
End Sub

Public Function GetFromINI(Section As String, key As String, Directory As String) As String
    Dim strBuffer As String

    strBuffer = String(750, Chr(0))
    key$ = LCase$(key$)
    GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal key$, "", strBuffer, Len(strBuffer), Directory$))

End Function
Public Sub getVPwindows()

    Dim rs As ADODB.Recordset
    Dim STRSQL As String
    Dim intI As Integer
    intI = 0
    
    STRSQL = "Select * from tblVPWindows"
    RunSQLCommand SQLType.Select_, STRSQL, gdbConn, rs
   
    While Not rs.EOF
        ReDim Preserve gstrVPWindows(intI)
        gstrVPWindows(intI) = rs!WindowName
        intI = intI + 1
        rs.MoveNext
    Wend
    
    rs.Close
    
    
    
End Sub
Public Function getOption(ByVal strKey, ByVal strType, Optional bolDecrypt As Boolean) As String
    Dim rsOption As ADODB.Recordset
    Dim STRSQL As String
    
    STRSQL = "Select optionValue from tblOptions Where optionKey='" & strKey & "' AND Type='" & strType & "'"
    RunSQLCommand SQLType.Select_, STRSQL, gdbConn, rsOption
    If rsOption.EOF = False Then
       If bolDecrypt = False Then
          getOption = Trim(rsOption!optionvalue & "")
       Else
          getOption = decrypt(Trim(rsOption!optionvalue & ""), gintKey)
       End If
    Else
      getOption = ""
    End If
End Function

Private Function GetOU(WinLogin As String) As String

    Dim objConnection As New ADODB.Connection
    Dim objCommand As New ADODB.Command
    Dim objRecordSet As ADODB.Recordset
    Dim strOU As String
    Dim strPath As String
    Dim a As Variant
    Dim arrPath() As String

    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = 2
    objCommand.CommandText = "SELECT adsPath FROM 'LDAP://dc=auas,dc=carlson,dc=com' WHERE objectCategory='user' AND sAMAccountName='" & WinLogin & "'"

    Set objRecordSet = objCommand.Execute
    strOU = ""
    If objRecordSet.EOF = False Then
        objRecordSet.MoveFirst
        Do Until objRecordSet.EOF
            strPath = objRecordSet.Fields("adsPath").value
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

Private Sub openDatabase(ByRef dbConn As ADODB.Connection, strConn As String)
    Dim strMsg As String
    
    If strConn <> "" Then
        dbConn.ConnectionString = strConn
        dbConn.Open
    Else
        strMsg = "Error in database connect string!"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        End
    End If
End Sub

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
      gobjLog.ErrorToLog "", Err.Number, Err.Description & " SQL: " & STRSQL
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

Public Sub pGetStartupValues()
    Dim rsConfig As ADODB.Recordset
    Dim STRSQL As String
    Dim strDefault As String
    Dim strMsg As String
    strDefault = "0000"

    If pSetGlobals(gobjHost.AgentDIV) = False Then
        If pSetGlobals(strDefault) = False Then
          strMsg = "Unable to retrieve system configuration setting!" & vbCrLf & vbCrLf _
          & "Please ensure that you are signon to GDS or the Agent configurations are setup in" & vbCrLf _
          & "TProData.  If you need assistance," & vbCrLf _
          & "please contact your system administrator!"
          modMsgBox.OKMsg = "OK"
          modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
          Call pTerminateApp
        End If
    End If
End Sub

Public Sub pTerminateApp()
    On Error Resume Next
    Dim loadedForm As Form
    pCloseDatabase
    'gobjSQ.Close
    gobjTE.Close
    gobjLog.CloseLog
    Set gobjSQ = Nothing
    Set gobjTE = Nothing
    Set gobjPNR = Nothing
    Set gobjFareQuotes = Nothing
    Set gobjHost = Nothing
    Set gobjLog = Nothing

    For Each loadedForm In Forms
        Unload loadedForm
        Set loadedForm = Nothing
    Next
    End

End Sub

Public Sub pCloseDatabase()
    On Error Resume Next
    gdbConn.Close
    Set gdbConn = Nothing
End Sub

Public Function pSetGlobals(strDiv As String) As Boolean

    Dim rsConfig As ADODB.Recordset
    Dim STRSQL As String

    STRSQL = "SELECT * FROM tblconfiguration where DIV='" & strDiv & "'"
    Set rsConfig = gdbConn.Execute(STRSQL)
    If Not rsConfig.EOF Then
        If Not IsNull(rsConfig!CountryCode) Then gstrAgcyCountryCode = UCase(rsConfig!CountryCode)
        gobjLog.LineTextToLog "gstrAgcyCountryCode =" & gstrAgcyCountryCode
        If Not IsNull(rsConfig!CurrencyCode) Then gstrAgcyCurrCode = UCase(rsConfig!CurrencyCode)
        gobjLog.LineTextToLog "gstrAgcyCurrCode =" & gstrAgcyCurrCode
        If Not IsNull(rsConfig!CityCode) Then gstrAgcyCityCode = UCase(rsConfig!CityCode)
        gobjLog.LineTextToLog "gstrAgcyCityCode =" & gstrAgcyCityCode
        If Not IsNull(rsConfig!AirportCode) Then gstrAgcyAirportCode = UCase(rsConfig!AirportCode)
        gobjLog.LineTextToLog "gstrAgcyAirportCode =" & gstrAgcyAirportCode
        If Not IsNull(rsConfig!Tel) Then gstrAgcyPhone = UCase(rsConfig!Tel)
        gobjLog.LineTextToLog "gstrAgcyPhone =" & gstrAgcyPhone
        If Not IsNull(rsConfig!PFCode) Then gstrPFCode = UCase(rsConfig!PFCode)
        gobjLog.LineTextToLog "gstrPFCode=" & gstrPFCode
        If Not IsNull(rsConfig!PFPCC) Then gstrPFPCC = UCase(rsConfig!PFPCC)
        gobjLog.LineTextToLog "gstrPFPCC=" & gstrPFPCC
        If Not IsNull(rsConfig!BKPCC) Then gstrPCC = UCase(rsConfig!BKPCC)
        gobjLog.LineTextToLog "gstrPCC =" & gstrPCC
        gstrHQPCC = gstrPCC
        gobjLog.LineTextToLog "gstrHQPCC =" & gstrHQPCC
        pSetGlobals = True
    End If

End Function

Public Sub pClearWindow()
Dim strMsg As String
Start:
    On Error Resume Next
    With frmDDEOwner.txtDDE
        .LinkItem = "ClearWindow"
        .Text = "1"
        .LinkPoke
        If Err.Number > 0 Then GoTo ErrProc
    End With
    
    Exit Sub
    
ErrProc:
    Select Case Err.Number
        Case 293
            Call pConnectToHost(frmDDEOwner.txtDDE)
            With frmDDEOwner.txtDDE
                 .LinkItem = "ClearWindow"
                 .Text = "1"
                 .LinkPoke
            End With
        Case 285 'Foreign application won't perform DDE method or operation
             AppActivate "Galileo Desktop"
             PressKey "R", , True
             GoTo Start
        Case Else
            strMsg = "Error " & Err.Number & vbCrLf & Err.Description
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End Select
End Sub

Public Function PNRExistsInGDS(Optional bolExit As Boolean) As Boolean

    Dim strMsg As String
    PNRExistsInGDS = False
    If gobjPNR.CheckPNRStatus <> -1 Then
    
        strMsg = "There is a PNR present in the GDS!" & Chr(13) _
                 & "Would you like to attempt to end the transaction?" & Chr(13) & Chr(13)
        modMsgBox.YESMsg = "End PNR"
        modMsgBox.NOMsg = "Ignore PNR"
        If bolExit = False Then
           modMsgBox.CANCELMsg = "Continue PNR"
        Else
           modMsgBox.CANCELMsg = "Close"
        End If
 
        Select Case modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbYesNoCancel, "CWT Desktop - PNR Exists")
            Case vbYes      'Send an ET
                If gobjHost.EndPNR = False Then
                   strMsg = "Unable to end the PNR!" & Chr(13) & "Is it OK to ignore it?"
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.CANCELMsg = "Cancel"
                    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbOKCancel, "CWT Desktop - PNR Exists") = vbOK Then
                        gobjHost.TerminalEntry "I"
                    Else     'Cancel and back to screen
                        PNRExistsInGDS = True
                    End If
                End If
            Case vbNo        'Ignore transaction
                gobjHost.TerminalEntry "I"
            Case vbCancel
                If bolExit = True Then PNRExistsInGDS = True
        End Select
    End If
End Function

Public Function fWantToQuit() As Boolean
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    
    If modMsgBox.sMsgBox(gVPMDIHwnd, "Are you sure you want to quit?", vbDefaultButton2 + vbYesNo + vbQuestion, "Quit") = vbYes Then
        fWantToQuit = True
    Else
        fWantToQuit = False
    End If
    
End Function

Public Sub pAddToVBILog(pRecLoc As String, pMod As String, startTime As Date, Optional SysStart As Date, Optional SubProc As String, Optional EndTime As Date, Optional FormStart As Date)
    Dim STRSQL As String
    Dim strSysStart As String
    Dim strSubProc As String
    If SysStart = "00:00:00" Then
        strSysStart = "null"
    Else
        strSysStart = "'" & Format(SysStart, "mm/dd/yyyy hh:nn:ss am/pm") & "'"
    End If
    If SubProc = "" Then
        strSubProc = "null"
    Else
        strSubProc = "'" & SubProc & "'"
    End If
    
    STRSQL = "INSERT into tblVBILOG (pcc,recloc,module,ModifiedBy,ModifiedDate,StartDate, SysStart, SubProcess,FormStart)" & _
             " values('" & gstrPCC & "','" & pRecLoc & "','" & pMod & "', '" & gobjHost.AgentName & "', '" & Format(IIf(EndTime = CdatDefaultDate, Now, EndTime), "mm/dd/yyyy hh:nn:ss am/pm") & "','" & _
             Format(startTime, "mm/dd/yyyy hh:nn:ss am/pm") & "', " & strSysStart & "," & strSubProc & ",'" & Format(FormStart, "mm/dd/yyyy hh:nn:ss am/pm") & "')"
    
    gdbConn.Execute STRSQL
    
End Sub

Public Function isLoaded(ByVal FormName As String) As Boolean

   Dim frm As Form
   Dim formCompared As String

   On Error GoTo RoutineExit
   isLoaded = False
   FormName = Trim$(UCase(FormName))
   
   For Each frm In Forms
       formCompared = Trim$(UCase(frm.Name))
       If formCompared = FormName Then
           isLoaded = True
           Exit For
       End If
   Next

   Exit Function
RoutineExit:
   isLoaded = False
   
End Function

Public Function IsSignedOn() As String
   pClearWindow
   pGetFromHost "OP/W*", "CaptureAll"
   If InStr(gstrFPResponse, "SIGN IN") > 0 Then
      IsSignedOn = "You must sign on to Galileo!"
   ElseIf InStr(gstrFPResponse, "DUTY CODE NOT AUTHORISED FOR TERMINAL") > 0 Then
      IsSignedOn = "You must emulate to PCC!"
   Else
      IsSignedOn = ""
   End If
End Function

Public Sub pGetFromHost(entry As String, Request As String)
    Dim intresponse As Integer
    Dim strResponse As String
    Dim strMsg As String
    
Start:
    On Error Resume Next
    With frmDDEOwner.txtDDE
        .LinkItem = "Transmit"
        .Text = entry
        .LinkPoke
        If Err.Number > 0 Then GoTo ErrProc
        
        .LinkItem = Request
        .LinkRequest
        If Err.Number > 0 Then GoTo ErrProc
        gstrFPResponse = .Text
    End With
    Exit Sub
ErrProc:
    Select Case Err.Number
        Case 293
            Call pConnectToHost(frmDDEOwner.txtDDE)
            With frmDDEOwner.txtDDE
                 .LinkItem = "ClearWindow"
                 .Text = "1"
                 .LinkPoke
            End With
        Case 285 'Foreign application won't perform DDE method or operation
             AppActivate "Galileo Desktop"
             PressKey "R", , True
             GoTo Start
        Case Else
            strMsg = "Error " & Err.Number & vbCrLf & Err.Description
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End Select
End Sub

Public Function fAllowAlpha(ByRef AsciiCode As Integer, Optional ByVal OtherAllowedCharacters As String = "") As Integer
    
    Dim lngC As Long
    Select Case AsciiCode
           Case 3, 22, 8, 65 To 90, 97 To 122
               fAllowAlpha = Asc(UCase(Chr(AsciiCode))) ' change valid characters to uppercase
           Case Else
                For lngC = 1 To Len(OtherAllowedCharacters)
                    If Asc(Mid(OtherAllowedCharacters, lngC, 1)) = AsciiCode Then
                        fAllowAlpha = AsciiCode
                        Exit Function
                    End If
                Next
                fAllowAlpha = 0
    End Select

End Function

Public Function fAllowNumeric(ByRef AsciiCode As Integer, Optional ByVal OtherAllowedCharacters As String = "") As Integer
    
    Dim lngC As Long
    Select Case AsciiCode
           Case 3, 22, 8, 48 To 57
               fAllowNumeric = Asc(UCase(Chr(AsciiCode))) ' change valid characters to uppercase USED IN CASE OTHERS ARE ALPHA
           Case Else
                For lngC = 1 To Len(OtherAllowedCharacters)
                    If Asc(Mid(OtherAllowedCharacters, lngC, 1)) = AsciiCode Then
                        fAllowNumeric = AsciiCode
                        Exit Function
                    End If
                Next

               fAllowNumeric = 0
    End Select

End Function

Public Function fAllowAlphaNumeric(ByRef AsciiCode As Integer, Optional ByVal OtherAllowedCharacters As String = "") As Integer
    
    Dim lngC As Long
    Select Case AsciiCode
           Case 3, 22, 8, 48 To 57, 65 To 90, 97 To 122
               fAllowAlphaNumeric = Asc(UCase(Chr(AsciiCode))) ' change valid characters to uppercase
           Case Else
                For lngC = 1 To Len(OtherAllowedCharacters)
                    If Asc(Mid(OtherAllowedCharacters, lngC, 1)) = AsciiCode Then
                        fAllowAlphaNumeric = AsciiCode
                        Exit Function
                    End If
                Next
               fAllowAlphaNumeric = 0
    End Select

End Function

Public Function convertText(t As String) As String
   convertText = t
   convertText = Replace(convertText, "_", "#")
   convertText = Replace(convertText, "-", "/")
   convertText = Replace(convertText, ";", "?")
   convertText = Replace(convertText, "'", ":")
   
End Function

Public Function actualText(t As String) As String
   actualText = t
   actualText = Replace(actualText, "#", "_")
   actualText = Replace(actualText, " ", "_")
   actualText = Replace(actualText, "//", "@")
   actualText = Replace(actualText, "/", "-")
   actualText = Replace(actualText, "?", ";")
   actualText = Replace(actualText, ":", "'")
End Function

Public Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
    
    Dim hMenu As Long
    
    hMenu = GetSystemMenu(hWnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
    
End Sub

Public Function GetAgentEmail(ProfileName As String, AgentSignOn As String, AgentPcc As String, Optional bolSkipEMO As Boolean, Optional bolVendor As Boolean) As String
   Dim rsEmail As ADODB.Recordset
   Dim STRSQL As String
   Dim strReplyEmail As String
   Dim intI As Integer
   'get sender email address based on priority
   'IF can't retrieve it, get sender email address based on CN from tblclients.
   'If can't retrieve it, get agent's email from tblagents
   
   If bolSkipEMO = False Then
        For intI = 1 To gobjPNR.ItinRemarkCount
           If Left(gobjPNR.ItinRemark(intI).RemarkText, 4) = "EMA." And _
              Len(gobjPNR.ItinRemark(intI).RemarkText) > 4 Then
              strReplyEmail = actualText(Mid(gobjPNR.ItinRemark(intI).RemarkText, 5))
              Exit For
           End If
        Next
   End If
   
   If Trim(strReplyEmail) <> "" Then
      GetAgentEmail = Trim(strReplyEmail)
   Else
        'Use different reply email for clients and vendors
        If bolVendor = True Then
           STRSQL = "Select VendorReplyEmail from tblclients where ProName='" & ProfileName & "'"
        Else
           STRSQL = "Select TeamEmail from tblclients where ProName='" & ProfileName & "'"
        End If
        Set rsEmail = gdbConn.Execute(STRSQL)
        If rsEmail.EOF Then
           GetAgentEmail = ""
        Else
           If bolVendor = True Then
              GetAgentEmail = rsEmail!VendorReplyEmail & ""
           Else
              GetAgentEmail = rsEmail!TeamEmail & ""
           End If
        End If

        If Trim(GetAgentEmail) = "" Then
              If Left(AgentPcc, 1) = "0" Then AgentPcc = Mid(AgentPcc, 2)
              STRSQL = "Select Email from tblAgents "
              STRSQL = STRSQL & "Where Sine = '" & AgentSignOn & "' "
              If IsNumeric(AgentSignOn) = False Then
                 STRSQL = STRSQL & "and PCC = '" & AgentPcc & "' "
              End If
              
              Set rsEmail = gdbConn.Execute(STRSQL)
              
              If rsEmail.EOF Then
                 GetAgentEmail = ""
              Else
                 GetAgentEmail = rsEmail!Email & ""
              End If
        End If
        
        rsEmail.Close
        Set rsEmail = Nothing
   End If
   
End Function

Public Function IsInControl(ByVal hWnd As Long) As Boolean
    Dim P As POINTAPI
    GetCursorPos P
    If hWnd = WindowFromPoint(P.X, P.Y) Then IsInControl = -1
End Function

Public Function checkSignOn(Optional noShow As Boolean) As Boolean
    Dim strResponse As String
    strResponse = IsSignedOn
    If strResponse <> "" Then
        checkSignOn = False
        If noShow = False Then
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, strResponse, vbOKOnly + vbDefaultButton1, "Sign On"
        End If
        Exit Function
    End If
    checkSignOn = True
End Function

Public Function GetCountryCode(CountryName As String) As String
    Dim rscountry As ADODB.Recordset
    Dim STRSQL As String
    
    STRSQL = "Select CountryCode from tblCountryCodes " & _
             "Where [CountryName] = '" & CountryName & "'"
    
    Set rscountry = gdbConn.Execute(STRSQL)
    
    If rscountry.EOF = False Then
       GetCountryCode = rscountry![CountryCode]
    Else
       GetCountryCode = ""
    End If
    
    rscountry.Close
    Set rscountry = Nothing
End Function
Public Function GetCountryName(CountryCode As String) As String
    Dim rscountry As ADODB.Recordset
    Dim STRSQL As String
    
    STRSQL = "Select CountryName from tblCountryCodes " & _
             "Where [CountryCode] = '" & CountryCode & "'"
    
    Set rscountry = gdbConn.Execute(STRSQL)
    
    If rscountry.EOF = False Then
       GetCountryName = rscountry![CountryName]
    Else
       GetCountryName = ""
    End If
    
    rscountry.Close
    Set rscountry = Nothing
End Function

Public Sub loadPassengers(ByRef cboPassenger As ComboBox)
    Dim intC As Integer
    Dim strTemp As String
    
    intC = gobjPNR.PassengerCount
    If intC > 0 Then
        For intC = 1 To intC
            With gobjPNR.PassengerName(intC)
                strTemp = Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
            End With
           cboPassenger.AddItem strTemp
           cboPassenger.ItemData(cboPassenger.NewIndex) = intC
        Next
    Else
        cboPassenger.AddItem " 1.1  UNKNOWN"
        cboPassenger.Enabled = False
    End If
    
    cboPassenger.listindex = 0
End Sub

Public Function pAddToQueueLog(ByVal strLocator As String, ByVal strType As String) As Boolean
    Dim STRSQL As String
    Dim strQKey As String
    'Get random number until not found in database
    strQKey = Random
    Do While pCheckQKey(strQKey)
       strQKey = Random
    Loop
    STRSQL = "Insert into tblQueueTime values('" & strQKey & "','" & gobjPNR.RecLoc & "','" _
              & strType & "','" & Format(Now, "mm/dd/yyyy hh:nn:ss am/pm") & "')"
    pAddToQueueLog = RunSQLCommand(SQLType.Insert_, STRSQL, gdbEitinConn)
End Function

Private Function pCheckQKey(ByVal strQKey As String) As Boolean
    Dim rsQueueLog As ADODB.Recordset
    Dim STRSQL As String
    
    STRSQL = "Select * from tblQueueTime Where QKey='" & strQKey & "'"
    Set rsQueueLog = gdbEitinConn.Execute(STRSQL)
    If rsQueueLog.EOF Then
       pCheckQKey = False
    Else
      pCheckQKey = True
    End If
End Function

Public Function Random() As String
Dim i, ch As Integer
Dim strTemp As String

'Use Randomize to initialize the random number generator.
'Then for each letter, use Rnd to pick a number between 0 and 26 + 26 + 10.
'If the number is less than 26 or 2*26, use it to pick a letter between A and Z.
'If the number is 2 * 26 or greater, subtract 2 * 26 and use the result to pick a digit between 0 and 9.
Randomize
For i = 1 To 8
    ch = Int((26 + 26 + 10) * Rnd)
    If ch < 26 Then
        strTemp = strTemp & Chr$(ch + Asc("A"))
    ElseIf ch < 2 * 26 Then
        ch = ch - 26
        strTemp = strTemp & Chr$(ch + Asc("A"))
    Else
        ch = ch - 26 - 26
        strTemp = strTemp & Chr$(ch + Asc("0"))
    End If
Next i
Random = strTemp
End Function

Public Function IsAlphaNumeric(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9A-Za-z]" Then Exit Function
        Next
    
    IsAlphaNumeric = True
    End If
    
End Function

Public Function loadedMinForms() As Boolean
    Dim loadedForm As Form
    
    loadedMinForms = True
    For Each loadedForm In Forms
        If UCase(loadedForm.Name) <> UCase("frmSideBar") And UCase(loadedForm.Name) <> UCase("frmDDEOwner") And UCase(loadedForm.Name) <> UCase("frmBars") Then
           loadedMinForms = False
           Exit For
        End If
    Next
End Function

Public Sub loadSimiliarNames()
      
      Dim strTemp() As String
      Dim item As ListItem
      On Error GoTo test:
      
    
      
      Load frmSimiliarNames
   
      frmSimiliarNames.bolProfile = False
      frmSimiliarNames.lvwNameList.ColumnHeaders.Clear
      frmSimiliarNames.lvwNameList.ColumnHeaders.Add 1, , "Last Name", 3200.12, 0
      frmSimiliarNames.lvwNameList.ColumnHeaders.Add 2, , "First Name", 3200.12, 0
      frmSimiliarNames.lvwNameList.ColumnHeaders.Add 3, , "Segment Date", 1500.06, 0
      frmSimiliarNames.lvwNameList.ColumnHeaders.Add 4, , "Rec Loc", 1200.06, 0
    
   
    
      With frmSimiliarNames
           With .lvwNameList
                .ListItems.Clear
                For lngC = 1 To gobjPNR.SimiliarNamesCount
                    strTemp = Split(gobjPNR.SimiliarName(lngC).PaxName, "/")
                    Set item = .ListItems.Add(, , strTemp(0))
                    If UBound(strTemp) > 0 Then
                       item.SubItems(1) = strTemp(1)
                    End If
                    If gobjPNR.SimiliarName(lngC).Status = True Then
                       item.SubItems(2) = ""
                    Else
                       item.SubItems(2) = Format(gobjPNR.SimiliarName(lngC).FirstSegDate, "DDMMMYY")
                    End If
                    item.SubItems(3) = gobjPNR.SimiliarName(lngC).PNR
                    If item.SubItems(2) <> "" Then
                       item.Tag = Format(gobjPNR.SimiliarName(lngC).FirstSegDate, "yyyy-mm-dd")
                    End If
                    item.Selected = False
                    
                Next
           End With
           .Show
           Do
             DoEvents
           
           Loop Until isLoaded("frmSimiliarNames") = False
          
     End With
     Exit Sub
test:
     
     MsgBox Err.Description & Err.Number
     
End Sub

Public Sub control_LostFocus(ByRef msFlex As MSFlexGrid, ByRef targetForm As Form, ByRef ctrl As Control, Optional strStruc As String = "H", Optional bolSkipFirstCol As Boolean = True)
    Dim i As Long
    Dim intX As Integer
    Dim intY As Integer
    
    If gintX <= msFlex.rows - 1 Then
        intX = msFlex.row
        intY = msFlex.col
        
        i = GetKeyState(VK_TAB) 'Tab Key
        If UCase(TypeName(ctrl)) = "TEXTBOX" Then
           msFlex.TextMatrix(gintX, gintY) = ctrl.Text
           ctrl.Visible = False
        ElseIf UCase(TypeName(ctrl)) = "COMBOBOX" Then
           msFlex.TextMatrix(gintX, gintY) = ctrl.Text
           ctrl.Container.Visible = False
        ElseIf UCase(TypeName(ctrl)) = "DTPICKER" Then
           If IsNull(ctrl.value) = False Then
              msFlex.TextMatrix(gintX, gintY) = ctrl.value
           Else
              msFlex.TextMatrix(gintX, gintY) = ""
           End If
           ctrl.Visible = False
        ElseIf UCase(TypeName(ctrl)) = "CHECKBOX" Then
           If ctrl.value = vbChecked Then
              msFlex.TextMatrix(gintX, gintY) = gstrChecked
           Else
              msFlex.TextMatrix(gintX, gintY) = gstrUnChecked
           End If
           ctrl.Visible = False
        End If
        msFlex.row = intX
        msFlex.col = intY
        
        If i = -127 Or i = -128 Then
            changeCelltext targetForm, msFlex.Name & "_Click", msFlex, strStruc, bolSkipFirstCol
        End If
    End If
End Sub

Public Sub changeCelltext(ByRef targetForm As Form, ByRef strFunction As String, ByRef msFlex As MSFlexGrid, strStyle As String, Optional bolSkipFirstCol As Boolean = True, Optional intSkip As Integer)
    Dim i As Integer
    
    If strStyle = "H" Then  'Horizontal
        For i = msFlex.Cols To 1 Step -1
            If msFlex.ColWidth(i - 1) <> 0 Then
               i = i
               Exit For
            End If
        Next
        
        If gintY < i - 1 Then
           If msFlex.Name = "msFlexSSR" Then
              If msFlex.rows > gintX Then
                 msFlex.row = gintX + 1
              End If
           Else
              If msFlex.ColWidth(gintY + 1) = 0 Then
                 msFlex.col = gintY + 2
              Else
                 msFlex.col = gintY + 1
              End If
           End If
           CallByName targetForm, strFunction, VbMethod
        ElseIf gintY = i - 1 Then
           If gintX < msFlex.rows - 1 Then
              msFlex.col = 0
              If bolSkipFirstCol Then msFlex.col = msFlex.col + 1
              msFlex.row = gintX + 1
              CallByName targetForm, strFunction, VbMethod
           End If
        End If
    ElseIf strStyle = "V" Then  'Vertical
        i = msFlex.rows
        If msFlex.RowHeight(i - 1) = 0 Then
           i = i - 1
        End If
        If gintX < i - 1 Then
           msFlex.row = gintX + 1
           CallByName targetForm, strFunction, VbMethod
        ElseIf gintX = i - 1 And msFlex.Name <> "msFlexPretripMI" And msFlex.Name <> "msFlexPostTripMI" Then
           If gintY < msFlex.Cols - 1 Then
              msFlex.row = 1
              msFlex.col = gintY + 1
              CallByName targetForm, strFunction, VbMethod
           End If
        End If
    End If
End Sub

Public Sub control_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByRef targetForm As Form, ByRef ctrl As Control, Optional bolAddDelete As Boolean = True)
    If bolAddDelete = True Then
        If KeyCode = 65 And Shift = 4 Then
           CallByName targetForm, "subMenuAdd_Click", VbMethod
        ElseIf KeyCode = 68 And Shift = 4 Then
           CallByName targetForm, "subMenuDelete_Click", VbMethod
        End If
    End If

    If KeyCode = 13 Then
       If UCase(TypeName(ctrl)) = "MSFLEXGRID" Then
          CallByName targetForm, ctrl.Name & "_Click", VbMethod
       Else
          ctrl.Visible = False
       End If
    End If
End Sub

Public Sub setControlPosition(ByRef msFlex As MSFlexGrid, ByRef ctrl As Control, ByVal intTop As Integer, ByVal intLeft As Integer, Optional ByRef ctrl2 As Control)
    
    gintX = msFlex.row
    gintY = msFlex.col
    ctrl.Left = intLeft
    ctrl.Top = intTop
    ctrl.Height = msFlex.CellHeight + 50
    ctrl.Width = msFlex.CellWidth
    ctrl.Visible = True
    ctrl.SetFocus
    ctrl.ZOrder 0
     
    If UCase(TypeName(ctrl)) = "TEXTBOX" Then
       ctrl.Text = msFlex.Text
    ElseIf UCase(TypeName(ctrl2)) = "COMBOBOX" Then
       ctrl2.Width = msFlex.CellWidth
       ctrl2.Height = msFlex.CellHeight + 50
       If Trim(msFlex.Text) <> "" Then
          ctrl2.Text = msFlex.Text
       Else
          If ctrl2.ListCount > 0 Then ctrl2.listindex = 0
       End If
       ctrl2.SetFocus
       AutoSizeDropDownWidth ctrl2
    ElseIf UCase(TypeName(ctrl)) = "DTPICKER" Then
       ctrl.Height = ctrl.Height + 30
       If Trim(msFlex.Text) <> "" Then
          ctrl.value = msFlex.Text
       ElseIf ctrl.CheckBox = True Then
          ctrl.value = Null
       ElseIf ctrl.CheckBox = False Then
          ctrl.value = Now
       End If
    ElseIf UCase(TypeName(ctrl)) = "CHECKBOX" Then
       ctrl.Left = (intLeft + (msFlex.CellWidth / 2)) - 100
       ctrl.Top = intTop
       ctrl.Height = msFlex.CellHeight
       ctrl.Width = 200
       i = GetKeyState(VK_TAB) 'Tab Key
       If msFlex.Text = gstrChecked Then
          If i = -127 Or i = -128 Then
             ctrl.value = vbChecked
          Else
             ctrl.value = vbUnchecked
          End If
       Else
          If i = -127 Or i = -128 Then
             ctrl.value = vbUnchecked
          Else
             ctrl.value = vbChecked
          End If
       End If
    End If
End Sub

Public Function ValidCCNum(Vendor As String, CCNum As String) As Boolean

    Dim strCompare As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intZ As Integer
    Dim intCD As Integer
    CCNum = Trim(CCNum)
    Select Case Vendor
        Case "AX"
            If Len(CCNum) <> 15 Or (Left(CCNum, 2) <> "34" And Left(CCNum, 2) <> "37") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "TP"
           If Len(CCNum) <> 15 Or (Left(CCNum, 4) <> "1920" And Left(CCNum, 4) <> "1220") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "VI", "BA"
            If (Len(CCNum) <> 16 And Len(CCNum) <> 13) _
            Or (Left(CCNum, 1) <> "4") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "MC", "CA", "IB"
            If (Len(CCNum) <> 16) _
            Or (Left(CCNum, 2) <> "51" And Left(CCNum, 2) <> "52" And Left(CCNum, 2) <> "53" And Left(CCNum, 2) <> "54" And Left(CCNum, 2) <> "55") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "DS"
            If (Len(CCNum) <> 16) _
            Or (Left(CCNum, 4) <> "6011") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "DC"
            If (Len(CCNum) <> 14) _
            Or (Left(CCNum, 2) <> "30" And Left(CCNum, 2) <> "36" And Left(CCNum, 2) <> "38") Then
                ValidCCNum = False
                Exit Function
            End If
        Case Else
            Err.Raise -1004, "CompanyProfile.ValidCCNum", "Unknown Credit Card Vendor"
    End Select
    strCompare = Format(CCNum, "00000000000000000000")
    
    For intX = 20 To 2 Step -2
        intY = CInt(Mid(strCompare, intX - 1, 1)) * 2
        intZ = CInt(Mid(strCompare, intX, 1))
        intCD = intCD + (intZ + IIf(intY < 10, intY, 1 + (intY - 10)))
    Next
    If (intCD / 10) - Int(intCD / 10) = 0 Then
       ValidCCNum = True
    Else
        ValidCCNum = False
    End If
End Function

Public Function fConvertZero(NumberAmt As String) As Single
    If IsNumeric(NumberAmt) Then
        fConvertZero = CSng(NumberAmt)
    Else
        fConvertZero = 0
    End If
End Function

Public Function sortInt(strText As String) As String

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strTemp() As String
    strTemp = Split(strText, ".")

    If UBound(strTemp) > 0 Then
       For i = 0 To UBound(strTemp)
           For j = i + 1 To UBound(strTemp)
               If CInt(strTemp(i)) > CInt(strTemp(j)) Then
                  k = strTemp(i)
                  strTemp(i) = strTemp(j)
                  strTemp(j) = k
               End If
           Next
       Next
    End If
    sortInt = ""
    For i = 0 To UBound(strTemp)
        sortInt = sortInt & strTemp(i) & "."
    Next
    If Right(sortInt, 1) = "." Then sortInt = Mid(sortInt, 1, Len(sortInt) - 1)
End Function

Private Sub getKeywords()
    Dim rs As ADODB.Recordset
    Dim STRSQL As String
    Dim i As Integer
    
    i = 0
    STRSQL = "Select Count(*) as RecordCount from tblKeywords where Type='SIDEBAR'"

    Set rs = gdbConn.Execute(STRSQL)
   
    Do Until rs.EOF
       ReDim Preserve gstrKeyword(0 To (rs!RecordCount - 1), 3)
       i = rs!RecordCount
       rs.MoveNext
    Loop
   
    If i = 0 Then
       gstrKeyword = Split("", vbCrLf)
    Else
       i = 0
       STRSQL = "Select Keyword, Red, Green, Blue from tblKeywords where Type='SIDEBAR'"
       Set rs = gdbConn.Execute(STRSQL)
   
       Do Until rs.EOF
          gstrKeyword(i, 0) = Trim(rs!Keyword) & ""
          gstrKeyword(i, 1) = rs!red
          gstrKeyword(i, 2) = rs!green
          gstrKeyword(i, 3) = rs!blue
          i = i + 1
          rs.MoveNext
       Loop
    End If
   
    rs.Close
    Set rs = Nothing
End Sub

Public Sub optSelectAll(ByRef msFlex As MSFlexGrid)
    Dim i As Integer
    Dim strChecked As String
    
    With msFlex
        If .TextMatrix(0, 0) = gstrUnChecked Then
           .TextMatrix(0, 0) = gstrChecked
        Else
           .TextMatrix(0, 0) = gstrUnChecked
        End If
        For i = 1 To .rows - 1
            If .TextMatrix(i, 0) <> "" Then
               .TextMatrix(i, 0) = .TextMatrix(0, 0)
               HighlightRow msFlex, i, IIf(.TextMatrix(0, 0) = gstrChecked, True, False)
            End If
        Next
    End With
    
End Sub

Public Sub HighlightRow(ByRef msFlex As MSFlexGrid, iRow As Integer, Optional bolHighlight As Boolean = True)
    
    Dim i As Integer
        
    With msFlex
         For i = .Cols - 1 To 0 Step -1
            .row = iRow
            .col = i
            If bolHighlight = True Then
               .CellBackColor = &HC0FFFF
            Else
               .CellBackColor = vbWhite
            End If
        Next
    End With
End Sub

Public Function PNRisRequired() As Boolean
    PNRisRequired = False
    If gobjPNR.RecLoc = "" Then
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, "PNR is required...", vbOKOnly + vbDefaultButton1, "CWT Desktop - PNR Required"
        PNRisRequired = True
    End If
End Function

Public Sub DataisRequired(ByVal strMsg As String)
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
End Sub

Public Sub setText(ByRef msFlex As MSFlexGrid, ByVal row As Integer, ByVal startCol As Integer, ByVal endCol As Integer)
    
    Dim i As Integer
    
    'Set the format for the column with checked box
    With msFlex
         For i = startCol To endCol
             .row = row
             .col = i
             .CellFontName = "Wingdings"
             .CellFontSize = 12
             .CellAlignment = flexAlignCenterCenter
             .Text = gstrUnChecked
         Next
         .col = 1
    End With
End Sub

Public Function rowSelected(ByRef msFlex As MSFlexGrid) As Integer
    
    Dim i As Integer
    
    rowSelected = 0
    With msFlex
        For i = 1 To .rows - 1
            If .TextMatrix(i, 0) = gstrChecked Then
               rowSelected = rowSelected + 1
            End If
        Next
    End With
    
End Function

Public Sub cmbGetFocus(ByRef cmbDropDown As MSForms.ComboBox)
    With cmbDropDown
         If Len(.Text) > 0 Then
            .SelStart = 0
            .SelLength = Len(.Text)
         End If
    End With
End Sub

Public Function convertPhoneText(t As String) As String
   convertPhoneText = t
   convertPhoneText = Replace(convertPhoneText, "_", "--")
   convertPhoneText = Replace(convertPhoneText, "@", "//")
End Function

Public Function actualPhoneText(t As String) As String
   actualPhoneText = t
   actualPhoneText = Replace(actualPhoneText, "--", "_")
   actualPhoneText = Replace(actualPhoneText, "//", "@")
End Function
Public Function initCap(Text As String) As String
    Dim a As Integer
    Dim na As String
    a = 1
    While (a <= Len(Text))
        If (a = 1) Then
            na = UCase(Mid(Text, a, 1))
            GoTo l1
        End If
        
        na = na + LCase(Mid(Text, a, 1))
l1:
        If (Asc(Mid(Text, a, 1)) = 32) Then
            na = na + UCase(Mid(Text, a + 1, 1))
            a = a + 1
        End If
        
        a = a + 1
    Wend
    'we return Na value back to function call
    initCap = na
End Function

Public Sub preFormLoad()
    
    gobjPNR.LoadPNR
    strResponse = gobjHost.TerminalEntry("*R")
    If InStr(1, strResponse, "NO B.F. TO DISPLAY - CREATE OR RETRIEVE FIRST") = 0 Then
       displayPNRinBar
    End If
End Sub
Public Sub findFPWindow()
    Dim hWnd As Long
    Dim wp As WINDOWPLACEMENT
        
    'Find the handler of focal point window to solve the problem of hanging issue when calling Viewpoint screen
    hWnd = FindWindowEx(gVPMDIHwnd, ByVal 0&, "FocalpointTerminalMDIFrame", vbNullString)
    
   'Minimized FocalpointTerminalMIDFrame and restore back
    wp.length = Len(wp)
    GetWindowPlacement hWnd, wp
    wp.showCmd = SW_SHOWMINIMIZED
    SetWindowPlacement hWnd, wp
    wp.showCmd = SW_SHOWNORMAL
    SetWindowPlacement hWnd, wp
    
End Sub

'*****************************************************************
'Functions below added during migration
'*****************************************************************
Public Sub pSetSelected()
  On Error Resume Next
  Screen.ActiveControl.SelStart = 0
  Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub

Public Sub pErrorReport(Optional ByVal TerminateApp As Boolean = True, Optional ByVal ShowMsgBox As Boolean = True)
    Dim strMsg As String
    gobjLog.ErrorToLog Err.Source, Err.Number, Err.Description
    If ShowMsgBox Then
       strMsg = "ERROR " & Err.Number & Chr(13) & Err.Description & Chr(13)
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End If
    If TerminateApp Then
        gobjLog.CloseLog
        Call pTerminateApp
    End If
End Sub

Public Function getCurrDispNum() As Integer
    Dim i As Integer
    Dim maxNum As Integer
    Dim strDI() As String
    On Error GoTo NoDisplayNum
    maxNum = 0
        
    With gobjPNR
    For i = 1 To .AcctRemarkCount
        If InStr(.AcctRemark(i).RemarkText, "/*") > 0 Then
            strDI = Split(.AcctRemark(i).RemarkText, "/")
            
'modified on 7/2/2005: skip all hotel DI: VLF, VRF, VEC, VFF
            If IsNumeric(Mid(strDI(1), 2)) And UCase(Mid(strDI(0), 1, 1)) <> "V" And UCase(strDI(0)) <> "NOCOMM" Then
                maxNum = IIf(maxNum < Mid(strDI(1), 2), Mid(strDI(1), 2), maxNum)
                
            End If
        End If
    Next i
    End With
    
    getCurrDispNum = maxNum
    Exit Function
NoDisplayNum:
    maxNum = 0
    getCurrDispNum = maxNum
End Function

Public Function isRequireClientMI(CN As String, location As Integer) As Boolean
    Dim rsMI As New ADODB.Recordset
    Dim STRSQL As String
    
    STRSQL = "SELECT tblClientMI.* FROM tblClientMI where cn = '" & CN & "'" & _
             " and location = " & location
    
    'Set rsMI = gdbTPro.OpenRecordset(strSQL)
    Set rsMI = gdbConn.Execute(STRSQL)
    
    If Not rsMI.EOF Then
        isRequireClientMI = True
    Else
        isRequireClientMI = False
    End If
    rsMI.Close
    
End Function

Public Function splitLongMSX(strMSLine As String) As String
Dim lngC As Long
Dim lngLen As Long
Dim strTemp As String

lngC = 0
strTemp = ""
Do Until Len(strMSLine) = 0
    If Len(strMSLine) <= 42 Then
        lngLen = Len(strMSLine)
    Else
        lngLen = InstrLast(Left(strMSLine, 42), "/") - 1
    End If
    
    'strTemp = strTemp & IIf(lngC = 0, "+DI.FT-MS", "+DI.FT-MSX") & Left(strMSLine, lngLen)
    strTemp = strTemp & "+DI.FT-MSX" & Left(strMSLine, lngLen)
    strMSLine = Mid(strMSLine, lngLen + 1)
    lngC = lngC + 1
Loop
splitLongMSX = strTemp
End Function

Public Sub pGetAgencyDefaults()
Dim STRSQL As String
'modified on 21/03/2005
Dim rsCurr As New ADODB.Recordset

STRSQL = "SELECT * FROM tblCurrency WHERE [CurrencyCode] = '" & gstrAgcyCurrCode & "'"
'Set rsCurr = gdbTProLU.OpenRecordset(strSQL)
Set rsCurr = gdbConn.Execute(STRSQL)
If Not rsCurr.EOF Then
    gstrAgcyCurrFormat = fCurrFormat(gstrAgcyCurrCode)
    gstrAgcyCurrRule = rsCurr![RoundRule]
    gbytAgcyCurrDec = rsCurr![Decimal]
    gsngAgcyCurrUnit = rsCurr![RoundUnit]
End If
    
End Sub

Public Function InstrLast(SourceString As String, SearchString As String, Optional ByVal Start As Long = 1, _
    Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Long
    Do
        Start = InStr(Start, SourceString, SearchString, CompareMethod)
        If Start = 0 Then Exit Do
        InstrLast = Start
        Start = Start + 1
    Loop
End Function

Public Function fCurrFormat(ByVal CurrencyCode As String) As String

Dim STRSQL As String
Dim rsCurr As New ADODB.Recordset
Dim intX As Integer
Dim intNumDec As Integer
Dim strFormat As String

STRSQL = "SELECT * FROM tblCurrency WHERE [CurrencyCode] = '" & CurrencyCode & "'"

'Set rsCurr = gdbTProLU.OpenRecordset(strSQL)

Set rsCurr = gdbConn.Execute(STRSQL)
intNumDec = rsCurr![Decimal]

strFormat = "#0" & IIf(intNumDec > 0, ".", "")
For intX = 1 To intNumDec
    strFormat = strFormat & "0"
Next

fCurrFormat = strFormat

End Function

Public Function fCurrRound(ByVal Amount As Single, ByVal CurrencyCode As String, _
    Optional ByVal RoundRule As String = "LOOKUP") As Single
'modified on 21/03/2005
Dim strRule As String
Dim sngUnit As Single
Dim rsCurr As New ADODB.Recordset
Dim STRSQL As String
Dim strMsg As String

strRule = RoundRule

If CurrencyCode = gstrAgcyCurrCode Then
    If RoundRule = "LOOKUP" Then
        strRule = gstrAgcyCurrRule
    End If
    sngUnit = gsngAgcyCurrUnit
Else

    STRSQL = "SELECT * FROM tblCurrency WHERE [CurrencyCode] = '" & CurrencyCode & "'"
    Set rsCurr = gdbConn.Execute(STRSQL)

    With rsCurr
        If .EOF Then
            strMsg = "Currency code not in database!"
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        Else
            If RoundRule = "LOOKUP" Then
                strRule = ![RoundRule]
            End If
            sngUnit = ![RoundUnit]
        End If
    End With
End If

If Left(strRule, 2) = "UP" And ((Amount / sngUnit) - Int(Amount / sngUnit) <> 0) Then
    Amount = Amount + (sngUnit - (((Amount / sngUnit) - Int(Amount / sngUnit)) * sngUnit))
ElseIf Left(strRule, 2) = "DO" And ((Amount / sngUnit) - Int(Amount / sngUnit) <> 0) Then
    Amount = Amount - (((Amount / sngUnit) - Int(Amount / sngUnit)) * sngUnit)
End If

fCurrRound = Round(Amount / sngUnit, 0) * sngUnit

End Function

Public Function getTCMapper(Vendor As String, amt As Double) As String
    Dim rsTC As New ADODB.Recordset
    Dim STRSQL As String
    Dim strAmt As String
    Dim intC As Integer
    Dim strMapList As String
    Dim strTC As String
    Dim strPrefix As String
    
    On Error GoTo getTCMapper_Error
    strAmt = CStr(amt)
    strTC = ""
    STRSQL = "SELECT AlphaCode, Prefix FROM tblAirTourCodes where airVendor = '" & Vendor & "'" & _
             " order by NumCode"
    'Set rsTC = gdbTPro.OpenRecordset(strSQL)
    Set rsTC = gdbConn.Execute(STRSQL)
    strMapList = ""
    While Not rsTC.EOF
        strMapList = strMapList & rsTC!AlphaCode
        strPrefix = rsTC!prefix & ""
        rsTC.MoveNext
    Wend
    rsTC.Close
    
    For intC = 1 To Len(strAmt)
        strTC = strTC & Mid(strMapList, Mid(strAmt, intC, 1) + 1, 1)
    Next
    
    getTCMapper = IIf(strPrefix <> "", strPrefix, "") & strTC

    Exit Function
getTCMapper_Error:
    getTCMapper = ""
    
End Function

Public Function validateCCVendor(cmbFOP As ComboBox) As Boolean

Dim lngC As Long

    validateCCVendor = False

    For lngC = 0 To cmbFOP.ListCount - 1
            If gobjPNR.FOP_CCCode = cmbFOP.List(lngC) Then
               validateCCVendor = True
               Exit For
            End If
    Next
  
End Function

Public Function validateCCDate(ExpDate As Date) As Boolean
    validateCCDate = False
    
    If ExpDate > Date Then
             validateCCDate = True
    End If
End Function

Public Sub promptCCError(Optional ValidCCVendor As Boolean = True, Optional ValidCCDate As Boolean = True)
Dim strError As String

    If ValidCCVendor = False Then
          strError = strError & "Invalid credit card vendor" & vbCrLf
    End If
    
    If ValidCCDate = False Then
        strError = strError & "Invalid credit card expiry date" & vbCrLf
    End If
    
    If strError <> "" Then
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strError & "Please update!", vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End If

End Sub

Public Sub addNP(ByVal intFFNum As Integer, ByVal bolCat35 As Boolean)
Dim i As Integer
Dim bolFound As Boolean
Dim intNPNum As Integer

bolFound = False
For i = 1 To gobjPNR.GeneralRemarkCount
   With gobjPNR.GeneralRemark(i)
        If .Qualifier = "TK" And .RemarkText = "NETTTKT-" & intFFNum & ":Y" Then
           bolFound = True
           intNPNum = .ItemNum
           Exit For
        End If
   End With
Next

If bolCat35 = True Then
   If bolFound = False Then
      gobjHost.TerminalEntry "NP.TK*NETTTKT-" & intFFNum & ":Y"
   End If
Else
   If bolFound = True Then
      gobjHost.TerminalEntry "NP." & intNPNum & "@"
   End If
End If

End Sub

Public Function VendorNum(ProductCode As String, SortKey As String) As String
Dim rs As ADODB.Recordset
Dim STRSQL As String
STRSQL = "Select VendorNumber,ProductCodes from tblVendors Where SortKey ='" & SortKey & "' "
Set rs = gdbConn.Execute(STRSQL)

Do
If InStr(rs!ProductCodes, ProductCode) Then
    VendorNum = "V" & rs!VendorNumber
    Exit Do
End If
rs.MoveNext

Loop While Not rs.EOF
rs.Close
Set rs = Nothing

End Function

Public Function pGetMIFromGDS(FFNum As String, Optional FileFare As Integer = 0, Optional MIFormat As String = "", Optional prodCode As String = "", Optional paxID As Integer = 0, Optional ByRef FFLineNo As Integer) As String
    Dim intFFFound As Integer
    Dim lngC As Long
    Dim lngD As Long
    Dim strAcctRmks() As String
    Dim FFValue As String
    
    On Error GoTo notFound
    FFValue = ""
    intFFFound = 0
    
    If Len(gobjPNR.RecLoc) = 0 Then
        Set gobjPNR = New CWT_GalileoPNR3.PNR
        gobjPNR.LoadPNR
    End If
    
    With gobjPNR
    If MIFormat = "M" Then
        For lngC = 1 To .GASaleRecordCount
            FFValue = ""
            If .GASalesRecord(lngC).ProductCode = prodCode Then
                For lngD = 1 To .GASalesRecord(lngC).FreeFieldTextCount
                    If UCase(Mid(.GASalesRecord(lngC).FreeFieldText(lngD), 1, IIf(InStr(.GASalesRecord(lngC).FreeFieldText(lngD), "-") - 1 > 0, InStr(.GASalesRecord(lngC).FreeFieldText(lngD), "-") - 1, 0))) = "FF" & FFNum Then
                        FFValue = UCase(Mid(.GASalesRecord(lngC).FreeFieldText(lngD), InStr(.GASalesRecord(lngC).FreeFieldText(lngD), "-") + 1))
                        intFFFound = intFFFound + 1
                    End If
                    If intFFFound > 0 Then
                        Exit For
                    End If
                Next lngD
                If intFFFound > 0 Then
                    Exit For
                End If
            End If
        Next lngC
    Else    'for MI format = "D" or "F"
    
        For lngC = 1 To .AcctRemarkCount
            FFValue = ""
            'If UCase(Mid(.AcctRemark(lngC).RemarkText, 1, IIf(InStr(.AcctRemark(lngC).RemarkText, "/") - 1 > 0, InStr(.AcctRemark(lngC).RemarkText, "/") - 1, 0))) = "FF" & FFNum Then
            '   If FileFare > 0 Then
            
             If FFNum <> "AC" Then
                 If UCase(Mid(.AcctRemark(lngC).RemarkText, 1, IIf(InStr(.AcctRemark(lngC).RemarkText, "/") - 1 > 0, InStr(.AcctRemark(lngC).RemarkText, "/") - 1, 0))) = IIf(IsNumeric(FFNum), "FF", "") & FFNum Then
                    If FileFare > 0 And IsNumeric(FFNum) Then
                        strAcctRmks = Split(.AcctRemark(lngC).RemarkText, "/")
                        If CInt(Right(strAcctRmks(1), Len(strAcctRmks(1)) - 1)) = FileFare Then
                            FFValue = UCase(strAcctRmks(2))
                            'FFValue = UCase(Mid(.AcctRemark(lngC).RemarkText, InStr(InStr(.AcctRemark(lngC).RemarkText, "/") + 1, .AcctRemark(lngC).RemarkText, "/") + 1))
                        End If
                    Else
                        FFValue = UCase(Mid(.AcctRemark(lngC).RemarkText, InStr(.AcctRemark(lngC).RemarkText, "/") + 1))
                    End If
                    intFFFound = intFFFound + 1
                End If
            Else
                If .AcctRemark(lngC).RemarkType = "AC" Then
                    FFValue = Replace(.AcctRemark(lngC).RemarkText, "AAA@", "")
                    FFLineNo = .AcctRemark(lngC).ItemNum
                    intFFFound = intFFFound + 1
                End If

            End If
            If intFFFound > 0 Then
                Exit For
            End If
        Next lngC
    End If
    End With
    pGetMIFromGDS = FFValue
    Exit Function
notFound:
    pGetMIFromGDS = ""
End Function

Public Sub pMoveBottomMI()
    Dim strAddDI As String
    Dim strDelDI As String
    Dim lngC As Integer
    Dim STRSQL As String
    Dim strRes As String
    Dim strEntry As String
    Dim strMsg As String
    'Dim intCount As Integer
    'Dim strTempAddDI As String
    'Dim strTempDelDI As String
    'Dim intI As Integer
    'intI = 0
    'intCount = 0
    'ReDim strAddDI(intI)
    'ReDim strDelDI(intI)
    'If Len(gobjPNR.RecLoc) = 0 Then
        'Set gobjPNR = New CWT_GalileoPNR.PNR
       ' If gobjLog.LogOpen Then gobjPNR.OpenLog gobjLog
        Set gobjPNR = New CWT_GalileoPNR3.PNR
        gobjPNR.LoadPNR
    'End If
    
    With gobjPNR
    
    For lngC = 1 To .AcctRemarkCount
        If UCase(Mid(.AcctRemark(lngC).RemarkText, 1, 2)) = "FF" And InStr(.AcctRemark(lngC).RemarkText, "/*") = 0 Then
            strAddDI = strAddDI & IIf(strAddDI <> "", "+", "") & "DI.FT-" & .AcctRemark(lngC).RemarkText
            strDelDI = strDelDI & IIf(strDelDI <> "", ".", "") & .AcctRemark(lngC).ItemNum
            'FFValue = UCase(Mid(.AcctRemark(lngC).RemarkText, InStr(.AcctRemark(lngC).RemarkText, "/") + 1))
            'intCount = intCount + 1
            '    strAddDI(intI) = strTempAddDI
            '    strDelDI(intI) = strTempDelDI
            'If intCount = 23 Then
            '    intI = intI + 1
                'ReDim Preserve strAddDI(intI)
                'ReDim Preserve strDelDI(intI)
           '     strTempAddDI = ""
           '     strTempDelDI = ""
           '     intCount = 0

            'End If
        
        If gbolMoveMILog = True Then
            STRSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
             "VALUES('" & Format(Now, "mm-dd-yyyy hh:mm:ss") & "','" & .AcctRemark(lngC).ItemNum & "','" & .AcctRemark(lngC).RemarkText & "','" & gobjPNR.RecLoc & "', " & _
             "'" & gobjHost.AgentName & "')"
            gdbConn.Execute STRSQL
        End If
        
        End If
    
    Next lngC
    End With
    'If UBound(strAddDI) = 0 Then
   
        If Len(strAddDI) > 0 And Len(strDelDI) > 0 Then
            'gobjHost.TerminalEntry strAddDI
            'gobjHost.TerminalEntry "DI." & strDelDI & "@"
            'strEntry = strAddDI & "+" & "DI." & strDelDI & "@"
            strEntry = strAddDI
            strRes = EntryToFP(strEntry)
            'EntryToFP "R.TPRO TKT QUEUE"
            'EntryToFP "ER"
            'EntryToFP "ER"
            If gbolMoveMILog = True Then
                STRSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
                 "VALUES('" & Format(Now, "mm-dd-yyyy hh:mm:ss") & "','" & strEntry & "','" & strRes & "','" & gobjPNR.RecLoc & "', " & _
                 "'" & gobjHost.AgentName & "')"
                gdbConn.Execute STRSQL
            End If
            
            If InStr(strRes, "*") = 0 Then 'And strRes <> "" Then
             strEntry = strAddDI & "+" & "DI." & strDelDI & "@"
             GoTo MoveError
            Else
                
                strEntry = "DI." & strDelDI & "@"
                strRes = EntryToFP(strEntry)
                
                If gbolMoveMILog = True Then
                STRSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
                 "VALUES('" & Format(Now, "mm-dd-yyyy hh:mm:ss") & "','" & strEntry & "','" & strRes & "','" & gobjPNR.RecLoc & "', " & _
                 "'" & gobjHost.AgentName & "')"
                gdbConn.Execute STRSQL
                End If
                
                If InStr(strRes, "*") = 0 Then 'And strRes <> "" Then
                    GoTo MoveError
                End If
                
                EntryToFP "R.TPRO TKT QUEUE"
                EntryToFP "ER"
                EntryToFP "ER"
                EntryToFP "ER"
            End If
    End If
   
   Exit Sub
MoveError:
                Clipboard.Clear
                Clipboard.setText (strEntry)
                
                strMsg = "Unable to Move MI(DI.FF) to Bottom!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes & _
                       vbCrLf & "The command is stored in clipboard, you have to paste the command to the focalpoint before proceed!"
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                gobjLog.LineTextToLog "Unable to Move MI(DI.FF) to Bottom!" & "Response from GDS is: " & Chr(13) & strRes
                
                If gbolMoveMILog = True Then
                    STRSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
                     "VALUES('" & Format(Now, "dd-MMM-yyyy hh:mm:ss") & "','DI FORMAT ERROR','" & Replace(strRes, " ", "") & "','" & strEntry & "', " & _
                     "'" & gobjHost.AgentName & "')"
                    gdbConn.Execute STRSQL
                End If
                
                EntryToFP "R.TPRO TKT QUEUE"
                EntryToFP "ER"
                EntryToFP "ER"
                EntryToFP "ER"
End Sub

Public Function EntryToFP(entry As String) As String

Dim intTry As Integer
Dim intresponse As Integer
Dim intCount As Integer
Dim strMsg As String
'intTry = intTry + 1
'EntryToFP = SendFPEntry(entry)
EntryToFP = gobjHost.TerminalEntry(entry)

'While SendFPEntry = "" And intTry < Retry
'    intTry = intTry + 1
'    EntryToFP = SendFPEntry(Entry)
'Wend

If EntryToFP = "" Then
     strMsg = CONNECTION_FAIL & " : " & Chr(13) & _
            "Entry: " & """" & entry & """" & " is aborted." & Chr(13) & _
            "Do you want to resend the entry?"
     modMsgBox.RETRYMsg = "Retry"
     modMsgBox.CANCELMsg = "Cancel"
     intresponse = modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbRetryCancel + vbDefaultButton1, "CWT Desktop - Error")
            If intresponse = vbRetry Then
                    Sleep (1000)
                    gobjLog.LineTextToLog "RetryEntry:" & entry
                    intCount = intCount + 1
                    gobjLog.LineTextToLog "Retry Count:" & intCount
                    EntryToFP = SendFPEntry(entry)
                    
            Else
                    strMsg = "Entry: " & """" & entry & """" & " is aborted." & Chr(13) & _
                    "Please toggle to focalpoint to enter/paste(Crtl+V) the entry before continue."
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                    gobjLog.LineTextToLog "CancelEntry:" & entry
                    gobjLog.LineTextToLog "Total Retry Count:" & intCount
                    Clipboard.Clear
                    Clipboard.setText (entry)
            End If
End If
End Function
Public Function SendFPEntry(entry As String) As String
Dim strTemp As String
Dim lngP As Long ' used for recommended pause
Dim strPath As String
Dim intFile As Integer
Dim intRetry As Integer
Dim strRes As String
Dim STRSQL As String

With gobjLog
    If .LogOpen Then
        .LineTextToLog "BEGIN FOCALPOINT ENTRY"
        .LineTextToLog "ENTRY>" & entry
    End If
End With

On Error GoTo Err_MakeEntry

With frmDDEOwner.txtDDE
    pClearWindow
    .LinkItem = "Transmit"
    .Text = entry
    .LinkPoke
    
    .LinkItem = "CaptureAll"
     .LinkRequest
     strRes = .Text
     If InStr(1, strRes, Chr(gintSOM)) <> 0 Then
        strRes = Mid(strRes, 1, InStr(1, strRes, Chr(gintSOM) & Space(5)))
     End If
     'If strRes = "" Then strRes = .Text
    SendFPEntry = strRes
End With

With gobjLog
    If .LogOpen Then
        .LineTextToLog strRes
        .LineTextToLog ">>>END FOCALPOINT RESPONSE> ", 0, 1
        .LineTextToLog String(64, "-"), 0, 1
    End If
End With
'Dim dattemp1 As Date
'dattemp1 = Now
    'For lngP = 0 To 8000000
    'recommended pause
    'Next lngP

Sleep (500)
' MsgBox DateDiff("s", Now, dattemp1)

Exit Function

Err_MakeEntry:

    With gobjLog
        If .LogOpen Then
            .LineTextToLog "SENDFP ERROR:" & " " & Err.Number & " " & Err.Description
            .LineTextToLog "SENDFP COMMAND:" & " " & entry
        End If
    End With
    
    
    If gbolTEErrorLog = True Then
            STRSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
             "VALUES('" & Format(Now, "dd-MMM-yyyy hh:mm:ss") & "','" & Err.Number & "','" & Replace(Err.Description, "'", "''") & " " & gobjPNR.RecLoc & "','" & entry & "', " & _
             "'" & gobjHost.AgentName & "')"
            gdbConn.Execute STRSQL
    End If
    
     SendFPEntry = ""
    'MsgBox "Unable to process command: " & Entry, vbCritical, "Ticket/Invoice Make Entry"
    Exit Function
End Function

Public Function pAddToAQQueueLog() As String
Dim STRSQL, strQKey As String
'Get random number until not found in database
strQKey = Random
Do While pCheckAQKey(strQKey)
   strQKey = Random
Loop
STRSQL = "Insert into lookup.dbo.AQTktLog (PNRQKEY,PNR,QUEUETIME,COUNTRY,PCC) values('" & strQKey & "','" & gobjPNR.RecLoc & "', '" & Format(Now, "mm/dd/yyyy hh:nn:ss am/pm") & "','" & gstrAgcyCountryCode & "','" & gobjHost.AgentPcc & "')"
gdbConn.Execute (STRSQL)
pAddToAQQueueLog = strQKey

End Function
Public Function pCheckAQKey(ByVal strQKey As String) As Boolean
Dim rsQueueLog As ADODB.Recordset
Dim STRSQL As String

STRSQL = "Select * from lookup.dbo.AQTktLog Where PNRQKey='" & strQKey & "'"
Set rsQueueLog = gdbConn.Execute(STRSQL)
If rsQueueLog.EOF Then
   pCheckAQKey = False
Else
   pCheckAQKey = True
End If
End Function

Public Sub AddQKeytoNP(qKey As String)
Dim STRSQL As String
Dim rs As ADODB.Recordset
Dim intI As Integer

For intI = 1 To gobjPNR.GeneralRemarkCount
    If gobjPNR.GeneralRemark(intI).Qualifier = "TK" Then
        If InStr(gobjPNR.GeneralRemark(intI).RemarkText, "AQUA QUEUEKEY:") > 0 Then
            gobjHost.TerminalEntry "NP." & gobjPNR.GeneralRemark(intI).ItemNum & "@"
            Exit For
        End If
    End If
Next
gobjHost.TerminalEntry "NP.TK*AQUA QUEUEKEY:" & qKey
End Sub

Public Function NPExist(Qualifier As String, value As String, Optional ByRef LineNum As Integer = 0) As Boolean
    Dim i As Integer
    
    With gobjPNR
      For i = 1 To .GeneralRemarkCount
          If UCase(.GeneralRemark(i).Qualifier) = UCase(Qualifier) Then
             If UCase(.GeneralRemark(i).RemarkText) = UCase(value) Then
                NPExist = True
                LineNum = .GeneralRemark(i).ItemNum
                Exit Function
             End If
          End If
      Next
    End With
    NPExist = False
End Function

Public Function MMM(MM As Integer) As String
   MMM = Mid(CstrMonths, 3 * (MM - 1) + 1, 3)
End Function

Public Function LastDate(Enterdate As Date) As Date

LastDate = DateSerial(Year(Enterdate), Month(Enterdate) + 1, 0)

End Function

Public Function fGetCNType(strProfile As String) As String
Dim STRSQL As String
Dim rs As New ADODB.Recordset

    STRSQL = "SELECT ClientType FROM tblClients WHERE [ProName] = '" & strProfile & "'"
    Set rs = gdbConn.Execute(STRSQL)

    
If rs.EOF Then
   fGetCNType = ""
Else
   fGetCNType = rs!ClientType & ""
End If
rs.Close
Set rs = Nothing
End Function

Public Sub SendEmail(STo As String, CC As String, BCC As String, SFrom As String, _
                     SenderEmail As String, SenderName As String, Subject As String, _
                     Body As String, HTML As Boolean, SDate As Date, _
                     MailType As String, Attachment As String, STime As Date, _
                     Country As String, PNR As String)
   Dim STRSQL As String
   
   STRSQL = "Insert into tblEmail (STo, CC, BCC, SFrom, SenderEmail, "
   STRSQL = STRSQL & "SenderName, Subject, Body, HTML, SDate, Type, "
   STRSQL = STRSQL & "Attachment, STime, Country, PNR) "
   STRSQL = STRSQL & "Values('" & SQLText(STo) & "' "
   STRSQL = STRSQL & ",'" & SQLText(CC) & "' "
   STRSQL = STRSQL & ",'" & SQLText(BCC) & "' "
   STRSQL = STRSQL & ",'" & SQLText(SFrom) & "' "
   STRSQL = STRSQL & ",'" & SQLText(SenderEmail) & "' "
   STRSQL = STRSQL & ",'" & SQLText(SenderName) & "' "
   STRSQL = STRSQL & ",'" & SQLText(Subject) & "' "
   STRSQL = STRSQL & ",'" & SQLText(Body) & "' "
   STRSQL = STRSQL & "," & IIf(HTML = True, "1", "0") & " "
   STRSQL = STRSQL & ",'" & Format(SDate, "dd/MMM/yyyy") & "' "
   STRSQL = STRSQL & ",'" & SQLText(MailType) & "' "
   STRSQL = STRSQL & ",'" & SQLText(Attachment) & "' "
   STRSQL = STRSQL & ",'" & Format(STime, "HH:mm:ss") & "' "
   STRSQL = STRSQL & ",'" & SQLText(Country) & "' "
   STRSQL = STRSQL & ",'" & SQLText(PNR) & "' "
   STRSQL = STRSQL & ")"
   
   gdbEmailConn.Execute STRSQL
End Sub

Public Function SQLText(Text As String) As String
   SQLText = Replace(Text, "'", "''")
End Function

Public Function fGetCityName(ByVal CityCode As String) As String
Dim STRSQL As String
Dim rsCity As New ADODB.Recordset

fGetCityName = ""

STRSQL = "SELECT tblCityCodes.City, tblCountrySubdivCodes.SubDivName, tblCountryCodes.CountryName " _
    & "FROM (tblCityCodes INNER JOIN tblCountryCodes ON tblCityCodes.CountryCode = tblCountryCodes.CountryCode) LEFT JOIN tblCountrySubdivCodes ON (tblCityCodes.CountrySubdivCode = tblCountrySubdivCodes.SubDivCode) " _
    & "WHERE tblCityCodes.AirportCode = '" & CityCode & "'"


Set rsCity = gdbConn.Execute(STRSQL)
With rsCity
    If .RecordCount > 0 Then fGetCityName = UCase(![City] & IIf(![SubDivName] & "" = "", "", ", " & ![SubDivName]) & ", " & ![CountryName])
End With


End Function

'Callback function for EnumChildWindows
Public Function EnumChildWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim txt As String
Dim Class As String
Dim newentry As String
Dim dummy As Integer
Dim wp As WINDOWPLACEMENT
Dim i As Integer

txt = Space$(MAX_PATH)
Class = Space$(MAX_PATH)

Call GetClassName(hWnd, Class, MAX_PATH)
Call GetWindowText(hWnd, txt, MAX_PATH)
'TrimNull(Class) = "FocalpointTerminalMDIFrame" And
newentry = TrimNull(Class) & vbTab & hWnd & vbTab & TrimNull(txt)
'Debug.Print "windowname:" & TrimNull(txt)

For i = 0 To UBound(gstrVPWindows)

If InStr(UCase(TrimNull(txt)), UCase(gstrVPWindows(i))) > 0 Then
      'rDisableWindow = EnableWindow(hwnd, gbolFPEnable)
      'Minimized window and restore back to cater the color faded issue when 1st loaded
       wp.length = Len(wp)
       GetWindowPlacement hWnd, wp
       wp.showCmd = SW_SHOWMINIMIZED
       SetWindowPlacement hWnd, wp
       'wp.showCmd = SW_SHOWNORMAL
       'SetWindowPlacement hwnd, wp
End If

Next
'frmChildWindows.lstChildWindows.AddItem newentry
'dummy = FindIndex(newentry, frmChildWindows.lstChildWindows)

'If dummy <> -1 Then
'    frmChildWindows.lstChildWindows.ItemData(dummy) = hWnd
'End If
EnumChildWindowProc = 1
End Function
Private Function TrimNull(item As String)

  'remove string before the terminating null(s)
   Dim pos As Integer
   
   pos = InStr(item, Chr$(0))
   
   If pos Then
         TrimNull = Left$(item, pos - 1)
   Else: TrimNull = item
   End If
   
End Function
Public Sub pEndProcessTimeLog(Module As String, SubModule As String, Process As String, Form As String, SubProcess As String, GroupID As String, Optional EndTime As Date, Optional startTime As Date)

    Dim STRSQL As String
    Dim strSysStart As String
    'Dim strSubProc As String
    
    If gdatStartTime <> "00:00:00" Or gdatStartTime <> CdatDefaultDate Then
        startTime = gdatStartTime
    End If
    If EndTime = "00:00:00" Or EndTime = CdatDefaultDate Then
        EndTime = Now
    End If
    '    strSysStart = "'" & Format(SysStart, "mm/dd/yyyy hh:nn:ss am/pm") & "'"
    'End If
    'If SubProc = "" Then
    '    strSubProc = "null"
    'Else
    '    strSubProc = "'" & SubProc & "'"
    'End If
    
    STRSQL = "Insert into tblCWTDTProcessTime " & _
            "(PNR,PCC,SIGNON,[MODULE],GROUPID,SUBMODULE,PROCESS,SUBPROCESS,FORM,STARTTIME,ENDTIME)" & _
             " values('" & gobjPNR.RecLoc & "'," & _
             "'" & gobjHost.AgentPcc & "'," & _
             "'" & gobjHost.AgentSine & "'," & _
             "'" & GroupID & "'," & _
             "'" & Module & "'," & _
             "'" & SubModule & "'," & _
             "'" & Process & "'," & _
             "'" & SubProcess & "'," & _
             "'" & Form & "'," & _
             "'" & Format(startTime, "mm/dd/yyyy hh:nn:ss") & "'," & _
             "'" & Format(EndTime, "mm/dd/yyyy hh:nn:ss") & "')"
             '"'" & Format(IIf(EndTime = CdatDefaultDate, Now, EndTime), "mm/dd/yyyy hh:nn:ss am/pm") & "'," & _

    
    gdbConn.Execute STRSQL
    
    'pResetStartTime
    
End Sub
Public Sub pResetStartTime()

        gdatStartTime = gstrCdatDefaultDate
        gstrModule = ""
        gstrSubModule = ""
        gstrProcess = ""
        gstrForm = ""
        gstrProcessGrpID = ""
        gstrSubProcess = ""
        
End Sub


Public Sub pStartProcessTimeLog(Optional Module As String, Optional SubModule As String, Optional Process As String, Optional Form As String, Optional SubProcess As String, Optional ProcessGrpID As String)

        gdatStartTime = Now
        gstrModule = Module
        gstrSubModule = SubModule
        gstrProcess = Process
        gstrForm = Form
        gstrSubProcess = SubProcess
        gstrProcessGrpID = ProcessGrpID
        
End Sub
Public Function pGetProcessKey() As String
    Dim STRSQL As String
    Dim strQKey As String
    'Get random number until not found in database
    strQKey = Random
    Do While pCheckQLogKey(strQKey)
       strQKey = Random
    Loop
    pGetProcessKey = strQKey
End Function

Private Function pCheckQLogKey(ByVal strQKey As String) As Boolean
    Dim rsQueueLog As ADODB.Recordset
    Dim STRSQL As String
    
    STRSQL = "Select * from tblCWTDTProcessTime Where GroupID='" & strQKey & "'"
    Set rsQueueLog = gdbConn.Execute(STRSQL)
    If rsQueueLog.EOF Then
       pCheckQLogKey = False
    Else
      pCheckQLogKey = True
    End If
End Function


Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
' Example: Call FormOnTop(me.hWnd, True)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2

    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE

    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select


    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub
