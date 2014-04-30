Attribute VB_Name = "Startup"
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOACTIVATE = &H10
' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1

' ZhiSam - Standard Handle Number
Private Const HWND_DESKTOP = 0
' ZhiSam - MsgBox standard class name
Private Const MSGBOX_CLSNAME = "#32770"

'MsgBox - button control
Private Const BM_CLICK = &HF5
Public Sub ChangeTBBack(hwnd As Long, PNewBack As Long)
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
    Dim oDU As Galileo.DesktopUtils
    ' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
    'variable declaration
    Dim SPHwnd As String
    Dim sFO As FileSystemObject
    Dim strLocation As String

    
    
    
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
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    pOpenDatabase
    gIntModuleType = GUISelection
    
    ' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
    'Close the smart point if it is opened because it will conflict with the pSetGlobalObjects function
    If gIntModuleType <> gModuleType.PC Then
        SPHwnd = IsAppRunning("Smartpoint App")
        If SPHwnd > 0 Then CloseApplication ("Travelport.Smartpoint.exe")
    End If
    'JY - V1.2.6 20110916 - CR109 - Startup form for users to select database to be connected (AU ESC)
    'Change the sequence of the code
    
    'JY - 20100604 - Get the width and height of FP (To resize our form based on resolution)
    Set oDU = New Galileo.DesktopUtils
    gFPWidth = (oDU.FocalpointWorkAreaRightPos - oDU.FocalpointWorkAreaLeftPos) * Screen.TwipsPerPixelX
    gFPHeight = (oDU.FocalpointWorkAreaBottomPos - oDU.FocalpointWorkAreaTopPos) * Screen.TwipsPerPixelY
    gPadding = 100
    gSideBarWidth = 250 * Screen.TwipsPerPixelX 'This value must tally with the value in windows.js
    gCustomVPHeight = 185 * Screen.TwipsPerPixelY 'This value must tally with the value in windows.js

    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Assisted Panel Height value mutiple Pixel for Smart Point
    If gIntModuleType <> gModuleType.PC Then
        gAssistedPanelHeight = 147 * Screen.TwipsPerPixelY
    End If
        
    glngTargetHwnd = VPHwnd
    'Find the handler of MDIClient in Galileo Desktop to place the panel
    gVPMDIHwnd = FindWindowEx(glngTargetHwnd, ByVal 0&, "MDIClient", vbNullString)
    
    
    'pOpenDatabase
    pSetGlobalObjects
    
    'Find the handler of Galileo Desktop Custom Toolbar to place the TPro toolbar
    VPToolbarHwnd = FindWindowEx(VPHwnd, ByVal 0&, "AfxControlBar70", vbNullString)
    If VPToolbarHwnd = 0 Then
       'Custom Toolbar must be activated in Galileo Desktop
        MsgBox "Custom toolbar must be activated in Galileo Desktop!", vbCritical, "CWT Desktop - Startup Error"
        End
    Else
        'Show up toolbar and embed it into Galileo Desktop
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
        If gIntModuleType = gModuleType.SYEX Then
            frmSideBar.Width = 0
        End If
        
        frmBars.Show
        old_parent = SetParent(frmBars.hwnd, VPToolbarHwnd)
        frmBars.Move 0, 0
        
    End If
            
    If gVPMDIHwnd <> 0 Then
           
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
             If gIntModuleType = gModuleType.SYEX Then
                  frmCustomVP.Height = 0
                  frmSideBar.Width = 0
                  frmSideBar.Hide
                  
             ElseIf gIntModuleType = gModuleType.SP Then
                  frmCustomVP.Height = 0
                  Load frmSideBar
                 old_parent = SetParent(frmSideBar.hwnd, gVPMDIHwnd)
                 frmSideBar.Show
                 frmSideBar.Move 0, 0
                  
                  'Wrap the text in treeview (Important: Disable these 2 lines in debug mode)
                 SetWindowLong frmSideBar.treeViewTraveller.hwnd, GWL_STYLE, GetWindowLong(frmSideBar.treeViewTraveller.hwnd, GWL_STYLE) Or TVS_NOTOOLTIPS Or TVS_HASLINES
                 OldProc = SetWindowLong(frmSideBar.hwnd, GWL_WNDPROC, AddressOf WindowProc)
                  
             Else
                 'JY - 20100604 - Load the Custom View Point Form
                 Load frmCustomVP
                 frmCustomVP.Show
             
                     
                 Load frmSideBar
                 old_parent = SetParent(frmSideBar.hwnd, gVPMDIHwnd)
                 frmSideBar.Show
                 frmSideBar.Move 0, 0
                 
                 'Wrap the text in treeview (Important: Disable these 2 lines in debug mode)
                 SetWindowLong frmSideBar.treeViewTraveller.hwnd, GWL_STYLE, GetWindowLong(frmSideBar.treeViewTraveller.hwnd, GWL_STYLE) Or TVS_NOTOOLTIPS Or TVS_HASLINES
                 OldProc = SetWindowLong(frmSideBar.hwnd, GWL_WNDPROC, AddressOf WindowProc)
            End If
                        
       If checkSignOn(True) = True Then
       
          gobjPNR.loadPNR
          strResponse = gobjHost.terminalEntry("*R")
          If InStr(1, strResponse, "NO B.F. TO DISPLAY - CREATE OR RETRIEVE FIRST") = 0 Then
             'displayPNRinBar
             bolPNRExist = True
          End If
       End If
       
       'Get highlight keywords on sidebar from DB
       getKeywords
           
       If bolPNRExist Then displayPNRinBar
       
       'JY - 20100617 - Set the Point and Click screen as the default screen
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
            If gIntModuleType = gModuleType.PC Then
                pDisplayToFP (":*R")
            End If
              
       'Minimized window and restore back to cater the color faded issue when 1st loaded
       wp.length = Len(wp)
       GetWindowPlacement VPHwnd, wp
       wp.showCmd = SW_SHOWMINIMIZED
       SetWindowPlacement VPHwnd, wp
       wp.showCmd = SW_SHOWMAXIMIZED
       SetWindowPlacement VPHwnd, wp
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
            If gIntModuleType <> gModuleType.PC Then
                'Open back the Smart Point Application
                Set sFO = New FileSystemObject
                strLocation = "C:\Program Files\Travelport\Smartpoint\Travelport.Smartpoint.exe"
                If sFO.FileExists(strLocation) Then
                        Shell (strLocation)
                End If
                'If gIntModuleType = gModuleType.SP Then
                '    resizePCWindow frmSideBar.cmdReverse.Visible, frmCustomVP.cmdContract.Visible
                'End If
            Else
                resizePCWindow frmSideBar.cmdReverse.Visible, frmCustomVP.cmdContract.Visible
            End If
       
    Else
       'End application if can't find the handler of gVPMDIHwnd
        MsgBox "Handle of MDI for Galileo Desktop is not found!", vbCritical, "CWT Desktop - Startup Error"
        End
    End If
    
End Sub

Public Function IsAppRunning(sWindowName As String) As Long
    Dim hwnd As Long, hWndOffline As Long
    
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
         If frmBars.SftTabs.Tabs.Current = 0 Then
           .Nodes.Add , , "AP", "AIR POLICY"
           .Nodes.Add , , "APRI", "AIR PRICING"
           .Nodes.Add , , "APRE", "AIR PREFERENCE"
         ElseIf frmBars.SftTabs.Tabs.Current = 1 Then
           .Nodes.Add , , "HP", "HOTEL POLICY & PREFERENCE"
         ElseIf frmBars.SftTabs.Tabs.Current = 2 Then
           .Nodes.Add , , "CP", "CAR POLICY & PREFERENCE"
         End If
         .Nodes.Add , , "VI", "VISA INFORMATION"
         .Nodes.Add , , "END", ""
  
         '.Nodes.Add , , "OS", "ERROR"
         For intI = 1 To .Nodes.Count
             .Nodes(intI).Bold = True
             .Nodes(intI).Expanded = True
         Next
         
         'Populate the child nodes for each category
         For intI = 1 To gobjPNR.PassengerCount
             .Nodes.Add "TS", tvwChild, , gobjPNR.PassengerName(intI).LastName & ", " & gobjPNR.PassengerName(intI).FirstName & IIf(gobjPNR.PassengerName(intI).PassengerType = "I", " (INFANT)", "")
         Next
                   
         For intI = 1 To gobjPNR.GeneralRemarkCount
             If frmBars.SftTabs.Tabs.Current = 0 And gobjPNR.GeneralRemark(intI).Qualifier = "*B" Then
                .Nodes.Add "AP", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
             ElseIf frmBars.SftTabs.Tabs.Current = 0 And gobjPNR.GeneralRemark(intI).Qualifier = "*Z" Then
                .Nodes.Add "APRI", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
             ElseIf gobjPNR.GeneralRemark(intI).Qualifier = "*G" Then
                 If InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "TRAVELER TYPE: ") > 0 Then
                    intJ = InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "TRAVELER TYPE: ")
                    .Nodes.Add "TS", tvwChild, , "TYPE: " & Mid(gobjPNR.GeneralRemark(intI).RemarkText, intJ + 15)
                 ElseIf frmBars.SftTabs.Tabs.Current = 0 And (InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "SEAT") > 0 _
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
             ElseIf frmBars.SftTabs.Tabs.Current = 1 And gobjPNR.GeneralRemark(intI).Qualifier = "*H" Then
                .Nodes.Add "HP", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
             ElseIf frmBars.SftTabs.Tabs.Current = 2 And gobjPNR.GeneralRemark(intI).Qualifier = "*C" Then
                .Nodes.Add "CP", tvwChild, , Trim(gobjPNR.GeneralRemark(intI).RemarkText)
             End If
         Next
                  
         'If gobjPNR.FOPType = "" Then
         '   .Nodes.Add "OS", tvwChild, , "MISSING FOP"
         'End If
    
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
   'CC - V1.2.11 20120410 - For HKSG Desktop UAT, Read HBU URL from INI file instead of Database
   'CC - HBU Staging URL will be added to HKSG Desktop UAT server's INI file only
   Dim strHBUStagingURL As String
   
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
   
   'CC - V1.2.11 20120410 - For HKSG Desktop UAT, Read HBU URL from INI file instead of Database
   'CC - HBU Staging URL will be added to HKSG Desktop UAT server's INI file only
   strHBUStagingURL = GetFromINI("HBUStaging", "URL", strINIFile)
   If strHBUStagingURL <> "" Then
        strHBUStagingURL = decrypt(strHBUStagingURL, gintKey)
   End If
      
   strPath = GetOU(getWinLogon)
   i = InStr(1, strPath, "\")
   j = InStr(i + 1, strPath, "\")
   i = InStr(j + 1, strPath, "\")
   strCountry = Mid(strPath, j + 1, i - j - 1)
    
   openDatabase gdbConn, gstrConn
   'CC - V1.2.8 20111028
   openDatabase gdbAPPConn, gstrConn
   If strCountry <> "SG" And strCountry <> "HK" And strCountry <> "IN" Then  '20090202
      strCountry = getOthOU(strCountry)
   End If

   
   'JY - V1.2.6 20110916 - CR109 - Startup form for users to select database to be connected (AU ESC)
   If gstrDatabaseConnected <> "" Then strCountry = gstrDatabaseConnected
   If strCountry = "" Then
      Load frmStartUp
      frmStartUp.Show
      Do
         DoEvents
      Loop Until isLoaded("frmStartUp") = False
      strCountry = gstrDatabaseConnected
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
  
   'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
   'CC - V1.2.11 20120410 - For HKSG Desktop UAT, Read HBU URL from INI file instead of Database
   'CC - HBU Staging URL will be added to HKSG Desktop UAT server's INI file only
   If strHBUStagingURL = "" Then
        gstrHBTURL = getOption("HBT", "WebSitePath")
   Else
        gstrHBTURL = strHBUStagingURL
   End If
   'gstrHBTURL = getOption("HBT", "WebSitePath")
   'gstrHBTURL = "http://10.180.28.173/hbtweb/default.aspx"
   
   'Preethi - V1.2.4 20110712 - CR 76 - Change Validation Logic For ENDPNR
   gstrPNRExpression = getOption("RecLoc", "Expression")
   gintCheckERLineNum = getOption("RecLoc_LineNum", "Expression")
  
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
    Dim strSQL As String
    Dim intI As Integer
    intI = 0
    
    strSQL = "Select * from tblVPWindows"
    RunSQLCommand SQLType.Select_, strSQL, gdbConn, rs
   
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
    Dim strSQL As String
    
    strSQL = "Select optionValue from tblOptions Where optionKey='" & strKey & "' AND Type='" & strType & "'"
    RunSQLCommand SQLType.Select_, strSQL, gdbConn, rsOption
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
    Dim strSQL As String
    strSQL = "Select Country from tblOUGroup Where OUPath = '" & strOUPath & "'"
    RunSQLCommand SQLType.Select_, strSQL, gdbConn, rsOUPath
    If rsOUPath.EOF = False Then
       getOthOU = Trim(rsOUPath!Country & "")
    Else
       getOthOU = ""
    End If
End Function

Public Function RunSQLCommand(CType As Integer, strSQL As String, Conn As ADODB.Connection, Optional ByRef rs As ADODB.Recordset) As Boolean
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
        bolExeSql = ExeSQLCommand(CType, strSQL, Conn, strErrNum, strErrDesc, rs)
     Else
        bolExeSql = ExeSQLCommand(CType, strSQL, Conn, strErrNum, strErrDesc)
     End If
   Loop Until bolExeSql = True Or intI = 5
   If bolExeSql = False And intI >= 5 Then
      RunSQLCommand = False
      gobjLog.ErrorToLog "", Err.Number, Err.Description & " SQL: " & strSQL
   ElseIf bolExeSql = True Then
      RunSQLCommand = True
   End If
End Function

Public Function ExeSQLCommand(CType As Integer, strSQL As String, Conn As ADODB.Connection, ByRef ErrNum As Long, ByRef ErrDesc As String, Optional ByRef rs As ADODB.Recordset) As Boolean
   On Error GoTo err1
   Select Case CType
      Case SQLType.Insert_, SQLType.Update_, SQLType.Delete_
         Conn.Execute strSQL
      Case SQLType.Select_
         DoEvents
         Set rs = Conn.Execute(strSQL)
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
    Dim strSQL As String
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
    Dim strSQL As String

    
    strSQL = "SELECT * FROM tblconfiguration where DIV='" & strDiv & "'"
    Set rsConfig = gdbConn.Execute(strSQL)
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
                If gobjHost.ENDPNR = False Then
                   strMsg = "Unable to end the PNR!" & Chr(13) & "Is it OK to ignore it?"
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.CANCELMsg = "Cancel"
                    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbOKCancel, "CWT Desktop - PNR Exists") = vbOK Then
                        gobjHost.terminalEntry "I"
                    Else     'Cancel and back to screen
                        PNRExistsInGDS = True
                    End If
                End If
            Case vbNo        'Ignore transaction
                gobjHost.terminalEntry "I"
            Case vbCancel
                If bolExit = True Then PNRExistsInGDS = True
        End Select
        'CC - V1.2.4 20110711  - ER01 - Added LoadPNR script (Fix the bug)
        Set gobjPNR = New CWT_GalileoPNR3.PNR
        gobjPNR.loadPNR
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
    Dim strSQL As String
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
    
    strSQL = "INSERT into tblVBILOG (pcc,recloc,module,ModifiedBy,ModifiedDate,StartDate, SysStart, SubProcess,FormStart)" & _
             " values('" & gstrPCC & "','" & pRecLoc & "','" & pMod & "', '" & gobjHost.AgentName & "', '" & Format(IIf(EndTime = CdatDefaultDate, Now, EndTime), "mm/dd/yyyy hh:nn:ss am/pm") & "','" & _
             Format(startTime, "mm/dd/yyyy hh:nn:ss am/pm") & "', " & strSysStart & "," & strSubProc & ",'" & Format(FormStart, "mm/dd/yyyy hh:nn:ss am/pm") & "')"
    
    gdbConn.Execute strSQL
    
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
   'CC - V1.2.20 20130617 - CR215 - Change in E-Docs Translation Table
   actualText = Replace(actualText, "--", "_")
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
    
    hMenu = GetSystemMenu(hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
    
End Sub

Public Function GetAgentEmail(ProfileName As String, AgentSignOn As String, AgentPCC As String, Optional bolSkipEMO As Boolean, Optional bolVendor As Boolean) As String
   Dim rsEmail As ADODB.Recordset
   Dim strSQL As String
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
           strSQL = "Select VendorReplyEmail from tblclients where ProName='" & ProfileName & "'"
        Else
           strSQL = "Select TeamEmail from tblclients where ProName='" & ProfileName & "'"
        End If
        Set rsEmail = gdbConn.Execute(strSQL)
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
              If Left(AgentPCC, 1) = "0" Then AgentPCC = Mid(AgentPCC, 2)
              strSQL = "Select Email from tblAgents "
              strSQL = strSQL & "Where Sine = '" & AgentSignOn & "' "
              If IsNumeric(AgentSignOn) = False Then
                 strSQL = strSQL & "and PCC = '" & AgentPCC & "' "
              End If
              
              Set rsEmail = gdbConn.Execute(strSQL)
              
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

Public Function IsInControl(ByVal hwnd As Long) As Boolean
    Dim P As POINTAPI
    GetCursorPos P
    If hwnd = WindowFromPoint(P.X, P.Y) Then IsInControl = -1
End Function

Public Function checkSignOn(Optional noShow As Boolean) As Boolean
    Dim strResponse As String
    

    strResponse = IsSignedOn
    If strResponse <> "" Then
        checkSignOn = False
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
        'do not populate sign on for Smart Point and SyEx server as smart Point have the logic to auto detect sign on
        If gIntModuleType = gModuleType.PC Then
            If noShow = False Then
               modMsgBox.OKMsg = "OK"
               modMsgBox.sMsgBox gVPMDIHwnd, strResponse, vbOKOnly + vbDefaultButton1, "Sign On"
            End If
        End If
        Exit Function
    End If
    checkSignOn = True

End Function

Public Function GetCountryCode(CountryName As String) As String
    Dim rscountry As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select CountryCode from tblCountryCodes " & _
             "Where [CountryName] = '" & SQLText(CountryName) & "'"
    
    Set rscountry = gdbConn.Execute(strSQL)
    
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
    Dim strSQL As String
    
    strSQL = "Select CountryName from tblCountryCodes " & _
             "Where [CountryCode] = '" & CountryCode & "'"
    
    Set rscountry = gdbConn.Execute(strSQL)
    
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
    Dim strSQL As String
    Dim strQKey As String
    'Get random number until not found in database
    strQKey = Random
    Do While pCheckQKey(strQKey)
       strQKey = Random
    Loop
    strSQL = "Insert into tblQueueTime values('" & strQKey & "','" & gobjPNR.RecLoc & "','" _
              & strType & "','" & Format(Now, "mm/dd/yyyy hh:nn:ss am/pm") & "')"
    pAddToQueueLog = RunSQLCommand(SQLType.Insert_, strSQL, gdbEitinConn)
End Function

Private Function pCheckQKey(ByVal strQKey As String) As Boolean
    Dim rsQueueLog As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select * from tblQueueTime Where QKey='" & strQKey & "'"
    Set rsQueueLog = gdbEitinConn.Execute(strSQL)
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
    Dim strSQL As String
    Dim i As Integer
    
    i = 0
    strSQL = "Select Count(*) as RecordCount from tblKeywords where Type='SIDEBAR'"

    Set rs = gdbConn.Execute(strSQL)
   
    Do Until rs.EOF
       ReDim Preserve gstrKeyword(0 To (rs!RecordCount - 1), 3)
       i = rs!RecordCount
       rs.MoveNext
    Loop
   
    If i = 0 Then
       gstrKeyword = Split("", vbCrLf)
    Else
       i = 0
       strSQL = "Select Keyword, Red, Green, Blue from tblKeywords where Type='SIDEBAR'"
       Set rs = gdbConn.Execute(strSQL)
   
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
        gstrCurrentPNR = Replace(frmSideBar.fraInfo.Caption, " PNR -> ", "")
    
    gobjPNR.loadPNR
    
    If frmSideBar.fraInfo.Caption <> " PNR -> " & gobjPNR.RecLoc Then
        'strResponse = gobjHost.TerminalEntry("*R")
        'If InStr(1, strResponse, "NO B.F. TO DISPLAY - CREATE OR RETRIEVE FIRST") = 0 Then
           displayPNRinBar
        
        'End If
        
        'gstrProcessGrpID = pGetProcessKey
        
    End If
End Sub
Public Sub findFPWindow()
    Dim hwnd As Long
    Dim wp As WINDOWPLACEMENT
        
    'Find the handler of focal point window to solve the problem of hanging issue when calling Viewpoint screen
    hwnd = FindWindowEx(gVPMDIHwnd, ByVal 0&, "FocalpointTerminalMDIFrame", vbNullString)
    
   'Minimized FocalpointTerminalMIDFrame and restore back
    wp.length = Len(wp)
    GetWindowPlacement hwnd, wp
    wp.showCmd = SW_SHOWMINIMIZED
    SetWindowPlacement hwnd, wp
    wp.showCmd = SW_SHOWNORMAL
    SetWindowPlacement hwnd, wp
    
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
    Dim strSQL As String
    
    strSQL = "SELECT tblClientMI.* FROM tblClientMI where cn = '" & CN & "'" & _
             " and location = " & location
    
    'Set rsMI = gdbTPro.OpenRecordset(strSQL)
    Set rsMI = gdbConn.Execute(strSQL)
    
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
Dim strSQL As String
'modified on 21/03/2005
Dim rsCurr As New ADODB.Recordset

strSQL = "SELECT * FROM tblCurrency WHERE [CurrencyCode] = '" & gstrAgcyCurrCode & "'"
'Set rsCurr = gdbTProLU.OpenRecordset(strSQL)
Set rsCurr = gdbConn.Execute(strSQL)
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

Dim strSQL As String
Dim rsCurr As New ADODB.Recordset
Dim intX As Integer
Dim intNumDec As Integer
Dim strFormat As String

strSQL = "SELECT * FROM tblCurrency WHERE [CurrencyCode] = '" & CurrencyCode & "'"

'Set rsCurr = gdbTProLU.OpenRecordset(strSQL)

Set rsCurr = gdbConn.Execute(strSQL)
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
Dim strSQL As String
Dim strMsg As String

strRule = RoundRule

If CurrencyCode = gstrAgcyCurrCode Then
    If RoundRule = "LOOKUP" Then
        strRule = gstrAgcyCurrRule
    End If
    sngUnit = gsngAgcyCurrUnit
Else

    strSQL = "SELECT * FROM tblCurrency WHERE [CurrencyCode] = '" & CurrencyCode & "'"
    Set rsCurr = gdbConn.Execute(strSQL)

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
    Dim strSQL As String
    Dim strAmt As String
    Dim intC As Integer
    Dim strMapList As String
    Dim strTC As String
    Dim strPrefix As String
    
    On Error GoTo getTCMapper_Error
    strAmt = CStr(amt)
    strTC = ""
    strSQL = "SELECT AlphaCode, Prefix FROM tblAirTourCodes where airVendor = '" & Vendor & "'" & _
             " order by NumCode"
    'Set rsTC = gdbTPro.OpenRecordset(strSQL)
    Set rsTC = gdbConn.Execute(strSQL)
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
      gobjHost.terminalEntry "NP.TK*NETTTKT-" & intFFNum & ":Y"
   End If
Else
   If bolFound = True Then
      gobjHost.terminalEntry "NP." & intNPNum & "@"
   End If
End If

End Sub

Public Function VendorNum(ProductCode As String, SortKey As String) As String
Dim rs As ADODB.Recordset
Dim strSQL As String
strSQL = "Select VendorNumber,ProductCodes from tblVendors Where SortKey ='" & SortKey & "' "
Set rs = gdbConn.Execute(strSQL)

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
        gobjPNR.loadPNR
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
    Dim strSQL As String
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
        gobjPNR.loadPNR
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
            strSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
             "VALUES('" & Format(Now, "mm-dd-yyyy hh:mm:ss") & "','" & .AcctRemark(lngC).ItemNum & "','" & .AcctRemark(lngC).RemarkText & "','" & gobjPNR.RecLoc & "', " & _
             "'" & gobjHost.AgentName & "')"
            gdbConn.Execute strSQL
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
                strSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
                 "VALUES('" & Format(Now, "mm-dd-yyyy hh:mm:ss") & "','" & strEntry & "','" & strRes & "','" & gobjPNR.RecLoc & "', " & _
                 "'" & gobjHost.AgentName & "')"
                gdbConn.Execute strSQL
            End If
            
            If InStr(strRes, "*") = 0 Then 'And strRes <> "" Then
             strEntry = strAddDI & "+" & "DI." & strDelDI & "@"
             GoTo MoveError
            Else
                
                strEntry = "DI." & strDelDI & "@"
                strRes = EntryToFP(strEntry)
                
                If gbolMoveMILog = True Then
                strSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
                 "VALUES('" & Format(Now, "mm-dd-yyyy hh:mm:ss") & "','" & strEntry & "','" & strRes & "','" & gobjPNR.RecLoc & "', " & _
                 "'" & gobjHost.AgentName & "')"
                gdbConn.Execute strSQL
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
                    strSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
                     "VALUES('" & Format(Now, "dd-MMM-yyyy hh:mm:ss") & "','DI FORMAT ERROR','" & Replace(strRes, " ", "") & "','" & strEntry & "', " & _
                     "'" & gobjHost.AgentName & "')"
                    gdbConn.Execute strSQL
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
EntryToFP = gobjHost.terminalEntry(entry)

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
Dim strSQL As String

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
            strSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
             "VALUES('" & Format(Now, "dd-MMM-yyyy hh:mm:ss") & "','" & Err.Number & "','" & Replace(Err.Description, "'", "''") & " " & gobjPNR.RecLoc & "','" & entry & "', " & _
             "'" & gobjHost.AgentName & "')"
            gdbConn.Execute strSQL
    End If
    
     SendFPEntry = ""
    'MsgBox "Unable to process command: " & Entry, vbCritical, "Ticket/Invoice Make Entry"
    Exit Function
End Function

Public Function pAddToAQQueueLog() As String
Dim strSQL, strQKey As String
'Get random number until not found in database
strQKey = Random
Do While pCheckAQKey(strQKey)
   strQKey = Random
Loop
strSQL = "Insert into lookup.dbo.AQTktLog (PNRQKEY,PNR,QUEUETIME,COUNTRY,PCC) values('" & strQKey & "','" & gobjPNR.RecLoc & "', '" & Format(Now, "mm/dd/yyyy hh:nn:ss am/pm") & "','" & gstrAgcyCountryCode & "','" & gobjHost.AgentPCC & "')"
gdbConn.Execute (strSQL)
pAddToAQQueueLog = strQKey

End Function
Public Function pCheckAQKey(ByVal strQKey As String) As Boolean
Dim rsQueueLog As ADODB.Recordset
Dim strSQL As String

strSQL = "Select * from lookup.dbo.AQTktLog Where PNRQKey='" & strQKey & "'"
Set rsQueueLog = gdbConn.Execute(strSQL)
If rsQueueLog.EOF Then
   pCheckAQKey = False
Else
   pCheckAQKey = True
End If
End Function

Public Sub AddQKeytoNP(qKey As String)
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim intI As Integer

For intI = 1 To gobjPNR.GeneralRemarkCount
    If gobjPNR.GeneralRemark(intI).Qualifier = "TK" Then
        If InStr(gobjPNR.GeneralRemark(intI).RemarkText, "AQUA QUEUEKEY:") > 0 Then
            gobjHost.terminalEntry "NP." & gobjPNR.GeneralRemark(intI).ItemNum & "@"
            Exit For
        End If
    End If
Next
gobjHost.terminalEntry "NP.TK*AQUA QUEUEKEY:" & qKey
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
Dim strSQL As String
Dim rs As New ADODB.Recordset

    strSQL = "SELECT ClientType FROM tblClients WHERE [ProName] = '" & strProfile & "'"
    Set rs = gdbConn.Execute(strSQL)

    
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
   Dim strSQL As String
   
   strSQL = "Insert into tblEmail (STo, CC, BCC, SFrom, SenderEmail, "
   strSQL = strSQL & "SenderName, Subject, Body, HTML, SDate, Type, "
   strSQL = strSQL & "Attachment, STime, Country, PNR) "
   strSQL = strSQL & "Values('" & SQLText(STo) & "' "
   strSQL = strSQL & ",'" & SQLText(CC) & "' "
   strSQL = strSQL & ",'" & SQLText(BCC) & "' "
   strSQL = strSQL & ",'" & SQLText(SFrom) & "' "
   strSQL = strSQL & ",'" & SQLText(SenderEmail) & "' "
   strSQL = strSQL & ",'" & SQLText(SenderName) & "' "
   strSQL = strSQL & ",'" & SQLText(Subject) & "' "
   strSQL = strSQL & ",'" & SQLText(Body) & "' "
   strSQL = strSQL & "," & IIf(HTML = True, "1", "0") & " "
   strSQL = strSQL & ",'" & Format(SDate, "dd/MMM/yyyy") & "' "
   strSQL = strSQL & ",'" & SQLText(MailType) & "' "
   strSQL = strSQL & ",'" & SQLText(Attachment) & "' "
   strSQL = strSQL & ",'" & Format(STime, "HH:mm:ss") & "' "
   strSQL = strSQL & ",'" & SQLText(Country) & "' "
   strSQL = strSQL & ",'" & SQLText(PNR) & "' "
   strSQL = strSQL & ")"
   
   gdbEmailConn.Execute strSQL
End Sub

Public Function SQLText(Text As String) As String
   SQLText = Replace(Text, "'", "''")
End Function

Public Function fGetCityName(ByVal CityCode As String) As String
Dim strSQL As String
Dim rsCity As New ADODB.Recordset

fGetCityName = ""

strSQL = "SELECT tblCityCodes.City, tblCountrySubdivCodes.SubDivName, tblCountryCodes.CountryName " _
    & "FROM (tblCityCodes INNER JOIN tblCountryCodes ON tblCityCodes.CountryCode = tblCountryCodes.CountryCode) LEFT JOIN tblCountrySubdivCodes ON (tblCityCodes.CountrySubdivCode = tblCountrySubdivCodes.SubDivCode) " _
    & "WHERE tblCityCodes.AirportCode = '" & CityCode & "'"


Set rsCity = gdbConn.Execute(strSQL)
With rsCity
    If .RecordCount > 0 Then fGetCityName = UCase(![City] & IIf(![SubDivName] & "" = "", "", ", " & ![SubDivName]) & ", " & ![CountryName])
End With


End Function

'Callback function for EnumChildWindows
Public Function EnumChildWindowProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim txt As String
Dim Class As String
Dim newentry As String
Dim dummy As Integer
Dim wp As WINDOWPLACEMENT
Dim i As Integer

txt = Space$(MAX_PATH)
Class = Space$(MAX_PATH)

Call GetClassName(hwnd, Class, MAX_PATH)
Call GetWindowText(hwnd, txt, MAX_PATH)
'TrimNull(Class) = "FocalpointTerminalMDIFrame" And
newentry = TrimNull(Class) & vbTab & hwnd & vbTab & TrimNull(txt)
'Debug.Print "windowname:" & TrimNull(txt)

For i = 0 To UBound(gstrVPWindows)

If InStr(UCase(TrimNull(txt)), UCase(gstrVPWindows(i))) > 0 Then
      'rDisableWindow = EnableWindow(hwnd, gbolFPEnable)
      'Minimized window and restore back to cater the color faded issue when 1st loaded
       wp.length = Len(wp)
       GetWindowPlacement hwnd, wp
       wp.showCmd = SW_SHOWMINIMIZED
       SetWindowPlacement hwnd, wp
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
Public Sub pEndProcessTimeLog(CN As String, RequestType As String, Module As String, SubModule As String, Process As String, Form As String, SubProcess As String, GroupID As String, Optional EndTime As Date, Optional startTime As Date)

    Dim strSQL As String
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
    
'Backup on 26 Sept 2008 - Jeremy
'    strSQL = "Insert into tblCWTDTProcessTime " & _
'            "(PNR,PCC,SIGNON,[MODULE],GROUPID,SUBMODULE,PROCESS,SUBPROCESS,FORM,STARTTIME,ENDTIME)" & _
'             " values('" & gobjPNR.RecLoc & "'," & _
'             "'" & gobjHost.AgentPcc & "'," & _
'             "'" & gobjHost.AgentSine & "'," & _
'             "'" & Module & "'," & _
'             "'" & GroupID & "'," & _
'             "'" & SubModule & "'," & _
'             "'" & Process & "'," & _
'             "'" & SubProcess & "'," & _
'             "'" & Form & "'," & _
'             "'" & Format(startTime, "mm/dd/yyyy hh:nn:ss") & "'," & _
'             "'" & Format(EndTime, "mm/dd/yyyy hh:nn:ss") & "')"
'             '"'" & Format(IIf(EndTime = CdatDefaultDate, Now, EndTime), "mm/dd/yyyy hh:nn:ss am/pm") & "'," & _

    strSQL = "Insert into tblCWTDTProcessTime " & _
            "(PNR,CN,PCC,SIGNON,REQUESTTYPE,[MODULE],GROUPID,SUBMODULE,PROCESS,SUBPROCESS,FORM,STARTTIME,ENDTIME)" & _
             " values('" & gobjPNR.RecLoc & "'," & _
             "'" & CN & "'," & _
             "'" & gobjHost.AgentPCC & "'," & _
             "'" & gobjHost.AgentSine & "'," & _
             "'" & RequestType & "'," & _
             "'" & Module & "'," & _
             "'" & GroupID & "'," & _
             "'" & SubModule & "'," & _
             "'" & Process & "'," & _
             "'" & SubProcess & "'," & _
             "'" & Form & "'," & _
             "'" & Format(startTime, "mm/dd/yyyy hh:nn:ss") & "'," & _
             "'" & Format(EndTime, "mm/dd/yyyy hh:nn:ss") & "')"
             '"'" & Format(IIf(EndTime = CdatDefaultDate, Now, EndTime), "mm/dd/yyyy hh:nn:ss am/pm") & "'," & _


    gdbConn.Execute strSQL
    
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
    Dim strSQL As String
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
    Dim strSQL As String
    
    strSQL = "Select * from tblCWTDTProcessTime Where GroupID='" & strQKey & "'"
    Set rsQueueLog = gdbConn.Execute(strSQL)
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

Public Function CheckPreTrip()

    Dim strSQL As String
    Dim rsMI As ADODB.Recordset

    'Check if there's any Pretrip MI, if no then do nothing
    strSQL = "Select a.*, b.Format from tblClientMI a, tblMICategory b Where a.CN='" & gobjPNR.CN & "' AND a.location = b.code AND a.location = '3' Order by FF"
    Set rsMI = gdbConn.Execute(strSQL)

    'CheckPreTrip True means there are PreTrip MI to be entered
    If rsMI.EOF = False Then
        CheckPreTrip = True
    Else
        CheckPreTrip = False
    End If

    Set rsMI = Nothing

End Function
Public Sub UpdatePNR(GrpID As String, RecLoc As String)
    Dim strSQL As String

    
    strSQL = "Update tblcwtdtprocesstime set PNR='" & RecLoc & "' where groupid='" & GrpID & "' and PNR=''"
    
    gdbConn.Execute strSQL
    
    
End Sub


'CS Add Booking Tool FF35
Public Function GetBookingTool() As String
   Dim intI As Integer
   
   With gobjPNR
      For intI = 1 To .GeneralRemarkCount
          If .GeneralRemark(intI).Qualifier = "BT" Then
            'Change format to NP.BT*xxx
            
            ' If Left(.GeneralRemark(intI).RemarkText, 3) = "AIR" Then
            '    GetBookingTool = Mid(.GeneralRemark(intI).RemarkText, 5)
                
            '    Exit Function
            ' End If
            'If InStr(.GeneralRemark(intI).RemarkText, "*") > 0 Then
            'MsgBox .GeneralRemark(intI).RemarkText
                GetBookingTool = .GeneralRemark(intI).RemarkText
            'End If
          End If
      Next
   End With
End Function

Public Function OSNoMF(ProductCode As String, VendorNumber As String) As Boolean

    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    strSQL = "Select ProductCode from tblOSNoMF "
    strSQL = strSQL & "Where ProductCode = '" & ProductCode & "' "
    strSQL = strSQL & "and VendorNumber = '" & VendorNumber & "' "
    
    Set rs = gdbConn.Execute(strSQL)

    If rs.EOF = False Then
        OSNoMF = True
    Else
        OSNoMF = False
    End If
    
    rs.Close
    Set rs = Nothing

End Function

Public Function GenerateSecureFlight() As Collection
    Dim strSSR As String
    Dim strPassportCountry As String
    Dim strPassportNum As String
    Dim strPassportExp As String
    Dim strBirthday As String
    Dim strGender As String
    Dim strNationality As String
    Dim strPaxType As String
    Dim strLastName As String
    Dim strFirstName As String
    Dim intI As Integer
        
    Set GenerateSecureFlight = New Collection
    
    GetPassportInfoFromNP strPassportCountry, strPassportNum, strPassportExp, strBirthday, strGender
    strNationality = strPassportCountry
    
    For intI = 1 To gobjPNR.PassengerCount
        strPaxType = Left(gobjPNR.PassengerName(intI).PassengerType, 1)
        strLastName = gobjPNR.PassengerName(intI).LastName
        strFirstName = gobjPNR.PassengerName(intI).FirstName
        
        If intI > 1 Then
            strPassportCountry = ""
            strPassportNum = ""
            strNationality = ""
            strBirthday = ""
            strGender = ""
            strPaxType = ""
            strPassportExp = ""
        End If
        
        'SI.P2S1/DOCS*P/GB/S12345888/GB/12DEC06/MI/23JAN12/SMITH
        'strSSR = "P" & intI & "S1/"
        'strSSR = strSSR & "DOCS*P/"
        
        'P/GB/S12345888/GB/12DEC06/MI/23JAN12/SMITH
        strSSR = "P/"
        strSSR = strSSR & strPassportCountry & "/"
        strSSR = strSSR & strPassportNum & "/"
        strSSR = strSSR & strNationality & "/"
        If IsDate(strBirthday) Then
            strSSR = strSSR & Format(strBirthday, "ddMMMyy") & "/"
        Else
            strSSR = strSSR & strBirthday & "/"
        End If
        strSSR = strSSR & strGender & strPaxType & "/"
        If IsDate(strPassportExp) Then
            strSSR = strSSR & Format(strPassportExp, "ddMMMyy") & "/"
        Else
            strSSR = strSSR & strPassportExp & "/"
        End If
        strSSR = strSSR & strLastName & "/"
        strSSR = strSSR & strFirstName
        
        GenerateSecureFlight.Add strSSR
    Next
End Function

Public Sub GetPassportInfoFromNP(ByRef PassportCountry As String, ByRef PassportNum As String, _
    PassportExp As String, ByRef Birthday As String, ByRef Gender As String)
    Dim intI As Integer
    Dim strTmp As String
    Dim bolFoundPassport As Boolean
    Dim bolFoundBday As Boolean
    Dim bolFoundGender As Boolean
    
    bolFoundPassport = False
    bolFoundBday = False
    bolFoundGender = False
    
    PassportCountry = ""
    PassportNum = ""
    PassportExp = ""
    Birthday = ""
    Gender = ""
    
    For intI = 1 To gobjPNR.GeneralRemarkCount
        With gobjPNR.GeneralRemark(intI)
            If UCase(.Qualifier) = "*P" Then
                'PASSPORT NO: A14944446-MY-ISS -EXP 22AUG10
                If UCase(Left(.RemarkText, Len("PASSPORT NO: "))) = "PASSPORT NO: " Then
                    If bolFoundPassport = False Then
                        bolFoundPassport = True
                        'A14944446-MY-ISS -EXP 22AUG10
                        strTmp = Mid(.RemarkText, Len("PASSPORT NO: ") + 1)
                        If InStr(1, strTmp, "-") > 1 Then
                            'A14944446
                            PassportNum = Mid(strTmp, 1, InStr(1, strTmp, "-") - 1)
                            'MY-ISS -EXP 22AUG10
                            strTmp = Mid(strTmp, InStr(1, strTmp, "-") + 1)
                            'MY
                            PassportCountry = Mid(strTmp, 1, 2)
                            'MY-ISS -EXP 22AUG10
                            If InStr(1, strTmp, "-EXP") > 1 Then
                                ' 22AUG10
                                strTmp = Mid(strTmp, InStr(1, strTmp, "-EXP") + Len("-EXP"))
                                strTmp = Trim(strTmp)
                                strTmp = Replace(strTmp, "/", "-")
                                PassportExp = strTmp
                            End If
                        End If
                    End If
                ElseIf UCase(Left(.RemarkText, Len("GENDER-"))) = "GENDER-" Then
                    If bolFoundGender = False Then
                        bolFoundGender = True
                        strTmp = Mid(.RemarkText, Len("GENDER-") + 1)
                        strTmp = Trim(strTmp)
                        strTmp = Left(strTmp, 1)
                        Gender = strTmp
                    End If
                End If
            ElseIf UCase(.Qualifier) = "*G" Then
                'BDAY: 12/09/56
                If UCase(Left(.RemarkText, Len("BDAY: "))) = "BDAY: " Then
                    If bolFoundBday = False Then
                        bolFoundBday = True
                        '12/09/56
                        strTmp = Mid(.RemarkText, Len("BDAY: ") + 1)
                        strTmp = Trim(strTmp)
                        strTmp = Replace(strTmp, "/", "-")
                        Birthday = strTmp
                    End If
                End If
'                'GENDER: M or GENDER:M ?
'                If UCase(Left(.RemarkText, Len("GENDER:"))) = "GENDER:" Then
'                    If bolFoundGender = False Then
'                        bolFoundGender = True
'                        ' M
'                        strTmp = Mid(.RemarkText, Len("GENDER:") + 1)
'                        strTmp = Trim(strTmp)
'                        strTmp = Left(strTmp, 1)
'                        Gender = strTmp
'                    End If
'                End If
            End If
        End With
        If bolFoundPassport = True And bolFoundBday = True And bolFoundGender = True Then
            Exit For
        End If
    Next
End Sub

Public Sub resizePCWindow(bolSideBarExpand As Boolean, bolCustomVPExpand)

    Dim oDU As Galileo.DesktopUtils
    Set oDU = New Galileo.DesktopUtils
    Dim pcHdl As Long
    Dim iLeft As Long
    Dim iTop As Long
    Dim iWidth As Long
    Dim iHeight As Long
    Dim intRetry As Integer
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
     If gIntModuleType <> gModuleType.PC Then
            resizeSPWindow bolSideBarExpand, bolCustomVPExpand
     Else
    
Retry:
            pcHdl = oDU.GetGalileoDesktopChildWindowHandle("Point-and-Click")
            If (pcHdl <> 0) Then
                'Resize the height for Window Point & Click
                If bolSideBarExpand Then
                   iLeft = oDU.FocalpointWorkAreaLeftPos + 250
                Else
                   iLeft = oDU.FocalpointWorkAreaLeftPos + 10
                End If
                
                iTop = oDU.FocalpointWorkAreaTopPos
                iWidth = oDU.FocalpointWorkAreaRightPos - iLeft
                
                If bolCustomVPExpand Then
                    iHeight = oDU.FocalpointWorkAreaBottomPos - iTop - 190
                Else
                    iHeight = oDU.FocalpointWorkAreaBottomPos - iTop - 10
                End If
                oDU.RemoveMaxMinButtons pcHdl, 1, 0
                oDU.RemoveSystemMenu pcHdl, 1
                oDU.ResizeWindowTo pcHdl, iLeft, iTop, iWidth, iHeight
            Else
                If intRetry < 2 Then
                   intRetry = intRetry + 1
                   Sleep (2000)
                   GoTo Retry
                End If
            End If
            
    End If

End Sub
' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
Public Sub resizeSPWindow(bolSideBarExpand As Boolean, bolCustomVPExpand)

  Dim oDU As Galileo.DesktopUtils
    Set oDU = New Galileo.DesktopUtils
    Dim pcHdl As Long
    Dim iLeft As Long
    Dim iTop As Long
    Dim iWidth As Long
    Dim iHeight As Long
    Dim iTopPos As Long
    Dim intRetry As Integer
   ' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
    Dim LngX As Long
    Dim lngY As Long
    Dim LngWidth As Long
    Dim LngHeight As Long
    Dim BolIndicator As Boolean
    Dim RectWindow As RECT
    Dim intPNRViewerRetry As Integer
    Dim intAssistedRetry As Integer
    
    ' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
    ' Adjust Smart Point Window Size
    intRetry = 0
    intPNRViewerRetry = 0
    intAssistedRetry = 0

Retry:

     pcHdl = IsAppRunning("Smartpoint App - Window")
  
    If (pcHdl <> 0) Then
        'ZhiSam - 19 Dec 2011 - Get the SmartPoint Current Window Position and Size
        BolIndicator = GetWindowRect(pcHdl, RectWindow)
        LngX = RectWindow.Left
        lngY = RectWindow.Top
        LngWidth = RectWindow.Right - RectWindow.Left
        LngHeight = RectWindow.Bottom - RectWindow.Top
        
        'Resize the height for Window Smart Point
        If bolSideBarExpand Then
            If gIntModuleType = gModuleType.SYEX Then
                iLeft = oDU.FocalpointWorkAreaLeftPos
            Else
                iLeft = oDU.FocalpointWorkAreaLeftPos + 250
                'iLeft = 252
            End If

        Else
            If gIntModuleType = gModuleType.SYEX Then
                iLeft = oDU.FocalpointWorkAreaLeftPos
            Else
                iLeft = oDU.FocalpointWorkAreaLeftPos + 10
                'iLeft = 12
            End If

        End If
        
        iTop = oDU.FocalpointWorkAreaTopPos
        'iTopPos = 131
        iTopPos = iTop + 120
        'iWidth = oDU.FocalpointWorkAreaRightPos - iLeft
        iWidth = LngWidth
        
        If bolCustomVPExpand Then
            'iHeight = oDU.FocalpointWorkAreaBottomPos - iTop - 190
            iHeight = LngHeight
        Else
            'iHeight = oDU.FocalpointWorkAreaBottomPos - iTop - 10
            iHeight = LngHeight
        End If
        
        BolIndicator = MakeWinMove(pcHdl, iLeft, iTopPos, iWidth, iHeight)

        
    Else
        If intRetry < 3 Then
           intRetry = intRetry + 1
           Sleep (2000)
           GoTo Retry
        End If
    End If
        
' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
' ZhiSam - 13 Feb 2012 - SmartPoint: Adjust PNR Viewer Window Size
RetryPNRViewer:
        pcHdl = IsAppRunning("PNR Viewer")
        'pcHdl = IsAppRunning("PNR Viewer - Smartpoint App")
        
        'Debug.Print ("PNR Viewer - Window: " & pcHdl)
        If pcHdl <> 0 Then
            
            iLeft = iLeft + iWidth
            iWidth = oDU.FocalpointWorkAreaRightPos - iLeft
            MakeWinMove pcHdl, iLeft, iTopPos, iWidth, iHeight
        Else
           If intPNRViewerRetry < 3 Then
                intPNRViewerRetry = intPNRViewerRetry + 1
                Sleep (2000)
                GoTo RetryPNRViewer
           End If
        
        End If
        
' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
' Adjust Smart Point Assusted Panel Window Size

RetryAssistedPanel:

        pcHdl = IsAppRunning("Assisted*")
        'Debug.Print ("Assisted* - Window: " & pcHdl)
        
        If pcHdl <> 0 Then
            
            'iLeft = iLeft + iWidth
            'iWidth = oDU.FocalpointWorkAreaRightPos - iLeft
            'MakeWinMove pcHdl, iLeft, iTopPos, iWidth, iHeight
            BolIndicator = GetWindowRect(pcHdl, RectWindow)
            LngX = RectWindow.Left
            lngY = RectWindow.Top
            LngWidth = RectWindow.Right - RectWindow.Left
            LngHeight = RectWindow.Bottom - RectWindow.Top
            
            gAssistedPanelHeight = LngHeight * Screen.TwipsPerPixelY
            gAssistedPanelWidth = LngWidth * Screen.TwipsPerPixelX
            'Debug.Print ("Assisted* - Window Height: " & gAssistedPanelHeight)
            'Debug.Print ("Assisted* - Window Width: " & gAssistedPanelWidth)
            
            
        Else
           If intAssistedRetry < 3 Then
                intAssistedRetry = intAssistedRetry + 1
                Sleep (2000)
                GoTo RetryAssistedPanel
           End If
        
        End If
  

End Sub
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Public Function AddLCCWebFare() As WebFares
    
    Dim objAirSeg As New FareOptionSegment
    Dim objWebFare As New WebFare
    Dim objWebFares As New WebFares
    
    Dim strStartLine As String
    Dim strEndLine As String
    Dim strBFText As String
    Dim strYQTaxText As String
    Dim strOthTaxText As String
    Dim strBookDtText As String
    Dim strBookTmText As String
    Dim strTotalFareText As String
    Dim strConfirmNumText As String

    Dim intI As Integer
    Dim bolFound As Boolean
    Dim bolHeader As Boolean
    
    Dim curBaseFare As Currency
    Dim curTax As Currency
    Dim strCarrier As String
    Dim strRouting As String
    Dim strBookingDate As String
    Dim strCurrency As String
    Dim strConfirmationNum As String
    Dim strTemp As String
    Dim strAry() As String
        
    strStartLine = "------- WEB FARE BOOKING "
    strEndLine = "------- END WEB FARE BOOKING "
    strBookDtText = "ON DATE: "
    strBFText = "BASE FARE: "
    strYQTaxText = "YQ TAX: "
    strOthTaxText = "OTHER TAX: "
    strTotalFareText = "TOTAL FARE PER PASSENGER "
    strConfirmNumText = "WEBSITE CONFIRMATION NUMBER: "
    
    bolFound = False
    bolHeader = False
    
    For intI = 1 To gobjPNR.GeneralRemarkCount
        With gobjPNR.GeneralRemark(intI)
            If bolFound = False And .Qualifier = "HL" _
                And UCase(Mid(.RemarkText, 1, Len(strStartLine))) = strStartLine Then
                'RemarkText = "------- WEB FARE BOOKING 3K SIN-HKG-SIN -------"
                bolFound = True
                bolHeader = True
                
                'Reset all variable
                Set objWebFare = New WebFare
                
                curBaseFare = 0
                curTax = 0
                strCarrier = ""
                strRouting = ""
                strBookingDate = ""
                strCurrency = ""
                strConfirmationNum = ""
                strTemp = ""
                
            Else
                bolHeader = False
            End If
            
            If bolFound And .Qualifier = "HL" Then
                If Mid(.RemarkText, 1, Len(strBFText)) = strBFText Then
                   'RemarkText = "BASE FARE: 252.95 SGD YQ TAX: 102 SGD OTHER TAX: 10 SGD"
                   
                   'strTemp = 252.95 SGD YQ TAX: 102 SGD OTHER TAX: 10 SGD
                   strTemp = Trim(Mid(.RemarkText, Len(strBFText) + 1))
                   'strTemp = 252.95
                   strTemp = GetString(strTemp, " ")
                   curBaseFare = CCur(IIf(IsNumeric(strTemp), strTemp, 0))
                                      
                   'strTemp = YQ TAX: 102 SGD OTHER TAX: 10 SGD
                   If InStr(.RemarkText, strYQTaxText) > 0 Then
                      strTemp = Mid(.RemarkText, InStr(.RemarkText, strYQTaxText))
                      'strtemp = 102 SGD OTHER TAX: 10 SGD
                      strTemp = Replace(strTemp, strYQTaxText, "")
                      'strTemp = 102
                      strTemp = GetString(strTemp, " ")
                      curTax = CCur(IIf(IsNumeric(strTemp), strTemp, 0))
                   End If
                   
                   'strTemp = YQ TAX: 102 SGD OTHER TAX: 10 SGD
                   If InStr(.RemarkText, strOthTaxText) > 0 Then
                      strTemp = Mid(.RemarkText, InStr(.RemarkText, strOthTaxText))
                      'strtemp = 10 SGD
                      strTemp = Replace(strTemp, strOthTaxText, "")
                      'strTemp = 10
                      strTemp = GetString(strTemp, " ")
                      curTax = curTax + CCur(IIf(IsNumeric(strTemp), strTemp, 0))
                   End If
                   
                                       
                ElseIf Mid(.RemarkText, 1, 2) = "0 " And InStr(1, .RemarkText, "/") > 10 Then
                    'RemarkText = "0 3K695Y 20MAY SINHKG/1550 1945"

                    strAry = Split(.RemarkText, " ")
                    
                    'UBound of strAry must be 4
                    If UBound(strAry) = 4 Then
                    
                       Set objAirSeg = New FareOptionSegment
                       With objAirSeg
                             'JY – V1.2.3 20110419 – IR11 - Should capture plating carrier from air segment details
                             If strCarrier = "" Then strCarrier = Mid(strAry(1), 1, 2)
                            .FlightNum = Mid(strAry(1), 3, Len(strAry(1)) - 3)
                            .Class = Right(strAry(1), 1)
                            .DepDate = strAry(2)
                            'strAry(3) = SINHKG/1550
                            .DepCity = Mid(strAry(3), 1, 3)
                            .ArrCity = Mid(strAry(3), 4, 3)
                            .DepTime = Right(strAry(3), 4)
                            .ArrTime = strAry(4)
                       End With
                       objWebFare.AddAirSeg objAirSeg
                    End If
                                                                                                                                                            
                ElseIf bolHeader = True Then
                    'RemarkText = "------- WEB FARE BOOKING 3K SIN-HKG-SIN -------"
                    'JY – V1.2.3 20110419 – IR11 - Should capture plating carrier from air segment details
                    'strCarrier = Mid(.RemarkText, 26, 2)
                    'strRouting ="DEL-BOM-DEL -------"
                    strRouting = Mid(.RemarkText, 29)
                    'strRouting ="DEL-BOM-DEL"
                    strRouting = GetString(strRouting, " -------")
                                    
                ElseIf Mid(.RemarkText, 1, Len(strBookDtText)) = strBookDtText Then
                    'RemarText = "ON DATE: 3/14/2010  AT: 18:45"
                    strBookingDate = Mid(.RemarkText, Len(strBookDtText) + 1)
                    'strBookingDate = 3/14/2010
                    strBookingDate = GetString(strBookingDate, " ")
                                                           
                ElseIf Mid(.RemarkText, 1, Len(strTotalFareText)) = strTotalFareText Then
                    'RemarkText = "TOTAL FARE PER PASSENGER 5448 INR"
                    strCurrency = Right(Trim(.RemarkText), 3)
                    
                '.RemarkText = "WEBSITE CONFIRMATION NUMBER: 012201"
                ElseIf Mid(.RemarkText, 1, Len(strConfirmNumText)) = strConfirmNumText Then
                    strConfirmationNum = Trim(Mid(.RemarkText, Len(strConfirmNumText) + 1))
                                    
                ElseIf Mid(.RemarkText, 1, Len(strEndLine)) = strEndLine Then
                    objWebFare.PlatingCarrier = strCarrier
                    objWebFare.Routing = strRouting
                    objWebFare.BookingDate = IIf(IsDate(strBookingDate), strBookingDate, Date)
                    objWebFare.BaseFare = curBaseFare
                    objWebFare.FareCurrency = strCurrency
                    objWebFare.Tax = curTax
                    objWebFare.TaxCode = "XT"
                    objWebFare.ConfirmationNum = strConfirmationNum
                    objWebFares.AddWebFare objWebFare
                    bolFound = False
                End If
            End If
        End With
    Next
    Set AddLCCWebFare = objWebFares
End Function

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
'FullString = "SG852E 20JUL DELBOM", BeforeChar = " ", Return = "SG852E"
Public Function GetString(FullString As String, BeforeChar As String) As String
    If InStr(1, FullString, BeforeChar) > 1 Then
        GetString = Mid(FullString, 1, InStr(1, FullString, BeforeChar) - 1)
    Else
        GetString = ""
    End If
End Function

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Public Function AddLCCFareOption() As WebFares
 
    Dim objAirSeg As New FareOptionSegment
    Dim objWebFareOption As New WebFare
    Dim objWebFareOptions As New WebFares
    
    Dim strStartLine As String
    Dim strEndLine As String
    Dim strBFText As String
    Dim strYQTaxText As String
    Dim strOthTaxText As String

    Dim intI As Integer
    Dim bolFound As Boolean
    Dim bolHeader As Boolean
    
    Dim curBaseFare As Currency
    Dim curTax As Currency
    Dim strCarrier As String
    Dim strCurrency As String
    Dim strTotalFareText As String
    Dim strTemp As String
    Dim strAry() As String
        
    strStartLine = "------- DECLINED WEB FARE "
    strEndLine = "------- END DECLINED WEB FARE "
    strBFText = "BASE FARE: "
    strYQTaxText = "YQ TAX: "
    strOthTaxText = "OTHER TAX: "
    strTotalFareText = "TOTAL FARE PER PASSENGER "
    
    bolFound = False
    bolHeader = False
    
    For intI = 1 To gobjPNR.GeneralRemarkCount
        With gobjPNR.GeneralRemark(intI)
            If bolFound = False And .Qualifier = "HL" _
                And UCase(Mid(.RemarkText, 1, Len(strStartLine))) = strStartLine Then
                'RemarkText = "------- DECLINED WEB FARE TR SIN-HKG-SIN -------"
                bolFound = True
                bolHeader = True
                
                'Reset all variable
                Set objWebFareOption = New WebFare
                
                curBaseFare = 0
                curTax = 0
                strCarrier = ""
                strCurrency = ""
                strTemp = ""
                
            Else
                bolHeader = False
            End If
            
            If bolFound And .Qualifier = "HL" Then
                If Mid(.RemarkText, 1, Len(strBFText)) = strBFText Then
                   'RemarkText = "BASE FARE: 252.95 SGD YQ TAX: 102 SGD OTHER TAX: 10 SGD"
                   
                   'strTemp = 252.95 SGD YQ TAX: 102 SGD OTHER TAX: 10 SGD
                   strTemp = Trim(Mid(.RemarkText, Len(strBFText) + 1))
                   'strTemp = 252.95
                   strTemp = GetString(strTemp, " ")
                   curBaseFare = CCur(IIf(IsNumeric(strTemp), strTemp, 0))
                                      
                   'strTemp = YQ TAX: 102 SGD OTHER TAX: 10 SGD
                   If InStr(.RemarkText, strYQTaxText) > 0 Then
                      strTemp = Mid(.RemarkText, InStr(.RemarkText, strYQTaxText))
                      'strtemp = 102 SGD OTHER TAX: 10 SGD
                      strTemp = Replace(strTemp, strYQTaxText, "")
                      'strTemp = 102
                      strTemp = GetString(strTemp, " ")
                      curTax = CCur(IIf(IsNumeric(strTemp), strTemp, 0))
                   End If
                   
                   'strTemp = YQ TAX: 102 SGD OTHER TAX: 10 SGD
                   If InStr(.RemarkText, strOthTaxText) > 0 Then
                      strTemp = Mid(.RemarkText, InStr(.RemarkText, strOthTaxText))
                      'strtemp = 10 SGD
                      strTemp = Replace(strTemp, strOthTaxText, "")
                      'strTemp = 10
                      strTemp = GetString(strTemp, " ")
                      curTax = curTax + CCur(IIf(IsNumeric(strTemp), strTemp, 0))
                   End If
                                                                          
                ElseIf Mid(.RemarkText, 1, 2) = "0 " And InStr(1, .RemarkText, "/") > 10 Then
                    'RemarkText = "0 3K695Y 20MAY SINHKG/1550 1945"

                    strAry = Split(.RemarkText, " ")
                    
                    'UBound of strAry must be 4
                    If UBound(strAry) = 4 Then
                    
                       Set objAirSeg = New FareOptionSegment
                       With objAirSeg
                             'JY – V1.2.3 20110419 – IR11 - Should capture plating carrier from air segment details
                             If strCarrier = "" Then strCarrier = Mid(strAry(1), 1, 2)
                            .FlightNum = Mid(strAry(1), 3, Len(strAry(1)) - 3)
                            .Class = Right(strAry(1), 1)
                            .DepDate = strAry(2)
                            'strAry(3) = SINHKG/1550
                            .DepCity = Mid(strAry(3), 1, 3)
                            .ArrCity = Mid(strAry(3), 4, 3)
                            .DepTime = Right(strAry(3), 4)
                            .ArrTime = strAry(4)
                       End With
                       objWebFareOption.AddAirSeg objAirSeg
                    End If
                                                                                                                                                            
                ElseIf bolHeader = True Then
                    'RemarkText = "------- DECLINED WEB FARE TR SIN-HKG-SIN -------"
                    'JY – V1.2.3 20110419 – IR11 - Should capture plating carrier from air segment details
                    'strCarrier = Mid(.RemarkText, 27, 2)
                                    
                ElseIf Mid(.RemarkText, 1, Len(strTotalFareText)) = strTotalFareText Then
                    'RemarkText = "TOTAL FARE PER PASSENGER 5448 INR"
                    strCurrency = Right(Trim(.RemarkText), 3)
                                                        
                ElseIf Mid(.RemarkText, 1, Len(strEndLine)) = strEndLine Then
                    objWebFareOption.PlatingCarrier = strCarrier
                    objWebFareOption.BaseFare = curBaseFare
                    objWebFareOption.FareCurrency = strCurrency
                    objWebFareOption.Tax = curTax
                    objWebFareOption.TaxCode = "XT"
                    objWebFareOptions.AddWebFare objWebFareOption
                    bolFound = False
                End If
            End If
        End With
    Next
    Set AddLCCFareOption = objWebFareOptions
      
End Function

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Public Function fGetCityNameOnly(ByVal CityCode As String) As String

    Dim strSQL As String
    Dim rsCity As New ADODB.Recordset

    fGetCityNameOnly = ""

    strSQL = "SELECT City from tblCityCodes WHERE AirportCode = '" & CityCode & "'"

    Set rsCity = gdbConn.Execute(strSQL)
    
    If Not rsCity.EOF Then
       fGetCityNameOnly = UCase(rsCity![City] & "")
    End If
    
End Function

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Public Function LeftAlign(Text As String, length As Long)
   If Len(Text) >= length Then
      LeftAlign = Left(Text, length)
   Else
      LeftAlign = Text & Space(length - Len(Text))
   End If
End Function
'Preethi - V1.2.4 20110621  - ER01 - Tracking of Touches to CWT Booking
Public Sub GetTouchesContactMethod(ByRef cmbContact As MSForms.ComboBox)

Dim strSQL As String
Dim rsContact As New ADODB.Recordset
Dim intI As Integer

 strSQL = "select ContactCode,ContactMethodDesc from tblPNRTouchContactMethod order by ContactMethodDesc"
 Set rsContact = gdbConn.Execute(strSQL)
 
 cmbContact.Clear
 cmbContact.ColumnCount = 2
 cmbContact.ColumnWidths = "0 cm; 5 cm"
 intI = 0
 While rsContact.EOF = False
   cmbContact.AddItem
   cmbContact.List(intI, 0) = (rsContact![ContactCode])
   cmbContact.List(intI, 1) = (rsContact![ContactMethodDesc])
   rsContact.MoveNext
   intI = intI + 1
 Wend
 rsContact.Close
 Set rsContact = Nothing
End Sub
'Preethi - V1.2.4 20110621  - ER01 - Tracking of Touches to CWT Booking
Public Sub GetTouchesPrimaryCode(ByRef cmbPrimaryReason As MSForms.ComboBox, ByVal CFA As String, ByRef ReasonCodeType As String)

Dim strSQL As String
Dim rsPR As New ADODB.Recordset
Dim intI As Integer

intI = 0
cmbPrimaryReason.Clear
cmbPrimaryReason.ColumnCount = 3
cmbPrimaryReason.ColumnWidths = "0 cm; 5 cm; 0cm"

strSQL = "select PrimaryCode,PrimaryCodeDesc from tblPNRTouchCodePrimary where CFA='" & CFA & "' order by PrimaryCodeDesc"
Set rsPR = gdbConn.Execute(strSQL)

If rsPR.EOF = True Then
  rsPR.Close
  Set rsPR = Nothing
  strSQL = "select PrimaryCode,PrimaryCodeDesc from tblPNRTouchCodePrimary where CFA= '00000' order by PrimaryCodeDesc"
  Set rsPR = gdbConn.Execute(strSQL)
  If rsPR.EOF = False Then
    ReasonCodeType = "S"
  End If
Else
  ReasonCodeType = "C"
End If
While rsPR.EOF = False
  cmbPrimaryReason.AddItem
  cmbPrimaryReason.List(intI, 0) = (rsPR![PrimaryCode])
  cmbPrimaryReason.List(intI, 1) = (rsPR![PrimaryCodeDesc])
  rsPR.MoveNext
  intI = intI + 1
Wend
  
rsPR.Close
Set rsPR = Nothing
End Sub
'Preethi - V1.2.4 20110621  - ER01 - Tracking of Touches to CWT Booking
Public Sub GetTouchesSecondaryCode(ByRef cmbSecondaryReason As MSForms.ComboBox, ByVal CFA As String, ByVal PrimaryCode As String)

Dim strSQL As String
Dim rsSR As New ADODB.Recordset
Dim intI As Integer

intI = 0
cmbSecondaryReason.Clear
cmbSecondaryReason.ColumnCount = 3
cmbSecondaryReason.ColumnWidths = "0 cm; 5 cm; 0cm"
strSQL = "select a.SecondaryCode,a.SecondaryCodeDesc,a.ChargeIndicator from tblPNRTouchCodeSecondary a "
strSQL = strSQL & "join tblPNRTouchCodePrimary b on a.PriCodeID=b.PriCodeID "
strSQL = strSQL & "where b.CFA='" & CFA & "' and b.Primarycode='" & PrimaryCode
strSQL = strSQL & "' order by a.SecondaryCodeDesc"
Set rsSR = gdbConn.Execute(strSQL)

While rsSR.EOF = False
  cmbSecondaryReason.AddItem
  cmbSecondaryReason.List(intI, 0) = (rsSR![SecondaryCode])
  cmbSecondaryReason.List(intI, 1) = (rsSR![SecondaryCodeDesc])
  cmbSecondaryReason.List(intI, 2) = (rsSR![ChargeIndicator])
  rsSR.MoveNext
  intI = intI + 1
Wend

rsSR.Close
Set rsSR = Nothing
End Sub

'Preethi - V1.2.4 20110712 - CR 76 - Change Validation Logic For ENDPNR
Public Function CheckResponse(Response As String, Expression As String, NumberLine As Integer) As Boolean
Dim regEx As RegExp
Dim strRes() As String
Dim intI As Integer

Set regEx = New RegExp
regEx.Pattern = Expression
regEx.IgnoreCase = False
CheckResponse = False

strRes() = Split(Response, vbCrLf)
For intI = 0 To NumberLine - 1
 If intI > UBound(strRes) Then
    Exit For
 Else
    CheckResponse = regEx.test(strRes(intI))
    If CheckResponse = True Then
       Exit For
    End If
 End If
Next
End Function

'CC - 20110816 - HBU
Public Function ApplicationUser(AgentPCC As String, AgentSignOn As String, Application As String, Optional ByRef Module As String = "") As Boolean
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    If IsNumeric(AgentSignOn) Then AgentPCC = ""
    
    strSQL = "Select AgentSignOn, Module from tblApplicationUser "
    strSQL = strSQL & "Where Application = '" & Application & "' and "
    strSQL = strSQL & "((AgentPCC = '" & AgentPCC & "' and AgentSignOn = '" & AgentSignOn & "') or "
    strSQL = strSQL & "(AgentSignOn = 'ALL USERS'))"
    
    Set rs = gdbConn.Execute(strSQL)
    
    If rs.EOF = False Then
        ApplicationUser = True
        Module = rs!Module & ""
    Else
        ApplicationUser = False
        Module = ""
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function
'Preethi - V1.2.6 20110907 - CR 90 - Change OBT Tool Code in FF35
Public Function getFF35OBT(BookingAction As String, BookingTool As String) As String
Dim strSQL As String
Dim rs As New ADODB.Recordset

If BookingAction = "EB" Then
   getFF35OBT = BookingTool
Else
   strSQL = "select Value from tblToolCode where ToolCode='" & BookingTool & "'"
   Set rs = gdbConn.Execute(strSQL)
   If Not rs.EOF Then
      getFF35OBT = rs![value]
   Else
      getFF35OBT = BookingTool
   End If
   rs.Close
   Set rs = Nothing
End If

End Function

'CC - V1.2.8 20111028
'change 1.2.3.5 to 1-3.5
Public Function FormatedLineNum(ByVal DotNumber As String) As String
        Dim Line() As String
        Dim i As Integer
        Dim intStart As Integer
        Dim strNewLine As String
        Dim strCurrent As String
        Dim bolNew As Boolean
        Dim strDotNumber As String

        strDotNumber = DotNumber
        
        If InStr(1, strDotNumber, ".") = 0 Then
            FormatedLineNum = strDotNumber
            Exit Function
        End If
        
        strDotNumber = sortInt(strDotNumber)

        Line = Split(strDotNumber, ".")
        bolNew = True
        For i = 0 To UBound(Line)
            If i = 0 Then
                intStart = Line(i)
                strCurrent = Line(i)
            Else
                If Line(i) = Line(i - 1) + 1 Then
                    strCurrent = intStart & IIf(intStart = Line(i), "", "-" & Line(i))
                Else
                    strNewLine = strNewLine & IIf(strNewLine = "", "", ".") & strCurrent
                    intStart = Line(i)
                    strCurrent = Line(i)
                End If
            End If
        Next
        FormatedLineNum = strNewLine & IIf(strCurrent = "", "", IIf(strNewLine = "", "", ".") & strCurrent)
End Function

'CC - V1.2.8 20111028
'Only Cater for DI and RI.S*, RI, NP line at the moment
Public Function SendGDSCmd(GDSCmd As Collection, CommandType As CmdType, Optional ByRef FailCMD As String = "") As String
    'Max DI line in 1 command is 29, we set the max line to 25.
    'Tested for some cases it only allowed 27
    
    Dim intI As Integer
    Dim intCount As Integer
    Dim intMax As Integer
    Dim intMaxChar As Integer
    Dim strCmd As String
    Dim strRes As String
    Dim strGDSCmd As String
    
    intMaxChar = 0
    If GDSCmd.Count = 0 Then Exit Function
    
    If CommandType = [DI] Then
        intMax = 25
    ElseIf CommandType = [RI.S] Then
        intMax = 9
        intMaxChar = 70 + 6 '6: RI.S1* , 70: Max char of RI
    ElseIf CommandType = [RI] Then
        intMax = 29
    ElseIf CommandType = [NP] Then
        intMax = 29
    ElseIf CommandType = [PE] Then
        intMax = 10
    End If
    intCount = 0
    strRes = ""
    For intI = 1 To GDSCmd.Count
        intCount = intCount + 1
        strGDSCmd = GDSCmd.item(intI)
        If intMaxChar <> 0 Then
            If Len(strGDSCmd) > intMaxChar Then
                strGDSCmd = Mid(strGDSCmd, 1, intMaxChar)
            End If
        End If
        strCmd = strCmd & IIf(strCmd = "", "", "+") & strGDSCmd
        If intCount >= intMax Then
            strRes = gobjHost.terminalEntry(UCase(strCmd))
            If strRes <> "*" And Mid(strRes, 1, 1) <> "*" Then ' (Mid(strRes, 1, 1) <> "*" And Mid(strRes, 7, 1) <> "*") Then 'Command Fail
                FailCMD = strCmd
                SendGDSCmd = strRes
                Exit Function
            End If
            intCount = 0
            strCmd = ""
        End If
    Next
    If strCmd <> "" Then
        strRes = gobjHost.terminalEntry(UCase(strCmd))
        If strRes <> "*" And Mid(strRes, 1, 1) <> "*" Then   'Command Fail
           FailCMD = strCmd
           SendGDSCmd = strRes
           Exit Function
        End If
    End If
    SendGDSCmd = ""
End Function

'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
'Preethi - V1.2.14 20120817  - CR161 - Aqua Itn - Validation in EItin and Queue Screen
Public Function PreLaunchAquaItinValidate() As Boolean
    Dim intI As Integer
    'Dim bolAQExist As Boolean
    Dim bolHZExist As Boolean
    Dim strErr As String
    Dim bolAHExist As Boolean
    
    For intI = 1 To gobjPNR.GeneralRemarkCount
        With gobjPNR.GeneralRemark(intI)
            'If .Qualifier = "" And UCase(Mid(.RemarkText, 1, 3)) = "AQ-" Then
                'bolAQExist = True
            'End If
'            If .Qualifier = "HZ" Then
'                bolHZExist = True
'            End If
            'CC - V1.2.20 20130612 - CR220 - Change to desktop logic for EM itinerary
            If .Qualifier = "HZ" Then
                If UCase(.RemarkText) = "CONF*SEND ITIN" Then
                    bolHZExist = True
                End If
                If UCase(.RemarkText) = "CONF*SEND ETIX" Then
                    bolHZExist = True
                End If
            End If
            If .Qualifier = "" And UCase(Mid(.RemarkText, 1, 3)) = "AH-" Then
                bolAHExist = True
            End If
            If bolHZExist = True And bolAHExist = True Then
                Exit For
            End If
        End With
    Next
    
    'If bolAQExist = False Then
    If bolAHExist = False Then
        'strErr = "Missing NP.AQ- lines and historical remarks."
        strErr = "AQUA Tag has not been added."
        strErr = strErr & vbCrLf & "Please document these remarks with Aqua Itin Rmk module."
    ElseIf bolHZExist = True Then
        strErr = "Your previous document has not been sent."
        strErr = strErr & vbCrLf & "Please try to send this document later."
    End If
    
    If strErr <> "" Then
        'MsgBox strErr
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strErr, vbOKOnly + vbDefaultButton1, "CWT Desktop"
        PreLaunchAquaItinValidate = False
    Else
        PreLaunchAquaItinValidate = True
    End If
End Function

'CC - V1.2.8 20111028
Public Function ENDPNR(IRIfFail As Boolean, Optional Receive As String = "") As Boolean
    Dim strCmd As String
    Dim strR As String
    Dim strTemp As String
    Dim strMsg As String
    
    ENDPNR = False
    If Receive <> "" Then
        strR = "R." & Receive
    Else
        strR = "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
    End If

    strCmd = strR & "+" & "ER"
    strResponse = gobjHost.terminalEntry(strCmd)
    strTemp = strResponse
    If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = False Then
        For i = 0 To 1
            strTemp = gobjHost.terminalEntry("ER")
            If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = True Then
                ENDPNR = True
                Exit Function
            End If
            If i = 1 Then
                ENDPNR = False
                strMsg = "Unable to END PNR. Response from GDS is " & Chr(13) & strTemp
                If IRIfFail = True Then
                    strMsg = strMsg & Chr(13) & "Desktop will Ignore & Retrieve (IR) the PNR"
                Else
                    strMsg = strMsg & Chr(13) & "System will continue without ending this booking."
                End If
                
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop"
                
                If IRIfFail = True Then
                    gobjHost.terminalEntry "IR"
                End If
            End If
        Next
    Else
        ENDPNR = True
    End If
    
End Function


Public Sub closeAll()

    'JY - V1.2.6 20110916 - CR109 - Startup form for users to select database to be connected (AU ESC)
    Dim loadedForm As Form
    
    'Close all the forms
    For Each loadedForm In Forms
        Unload loadedForm
        Set loadedForm = Nothing
    Next
    
    'Close all the database connections
    gdbConn.Close
    gdbEitinConn.Close
    gdbEmailConn.Close
    gdbAPPConn.Close
        
    'Reset all the global variables
    Set gobjPNR = Nothing
    Set gobjSeatMaps = Nothing
    Set gobjSQ = Nothing
    Set gobjTE = Nothing
    Set gobjFareQuotes = Nothing
    Set gobjHost = Nothing
    Set gobjLog = Nothing
    gstrTargetName = ""
    glngTargetHwnd = 0
    gVPMDIHwnd = 0
    gstrHostSession = ""
    Set gdbConn = Nothing
    Set gdbEitinConn = Nothing
    gstrConn = ""
    gstrEitinConn = ""
    gstrAgcyCountryCode = ""
    gstrAgcyCurrCode = ""
    gstrAgcyCurrFormat = ""
    gstrAgcyCurrRule = ""
    gbytAgcyCurrDec = 0
    gsngAgcyCurrUnit = 0
    gstrAgcyCityCode = ""
    gstrPCC = ""
    gstrHQPCC = ""
    gstrPFPCC = ""
    gstrAgcyAirportCode = ""
    gstrAgcyPhone = ""
    gstrPFCode = ""
    gbolCancelMove = False
    gbolMoveProfile = False
    gbolCancelProcess = False
    gbolPerformFF = False
    gbolGetProfileFromDB = False
    gstrFPResponse = ""
    gTrxnType = ""
    hMsVbLibToolBar = 0
    gstrPreviousText = ""
    gintX = 0
    gintY = 0
    Erase gstrKeyword
    gbolBack = False
    gbolBackToRecap = False
    gbolBackToSI = False
    gbolWritingtoPNR = False
    gMode = ""
    gbolSkipAdult = False
    gbolOverrideFare = False
    gbolSelectFare = False
    gbolNetFare = False
    gStartFareQuoteTime = "12:00:00 AM"
    gGetfareStart = "12:00:00 AM"
    gGetfareEnd = "12:00:00 AM"
    gFQSegID = 0
    gstrFOP(0) = ""
    gstrFOP(1) = ""
    gstrFOP(2) = ""
    gdblAmtToCom = 0
    gdblTaxToCom = 0
    gdblAmtToPax = 0
    gdblTaxToPax = 0
    gbolFMR = False
    gstrFOPToCom = ""
    gstrCCVendor = ""
    gstrCCNum = ""
    gstrCCExpDate = "12:00:00 AM"
    gstrPersCCVendor = ""
    gstrPersCCNum = ""
    gstrPersCCExpDate = "12:00:00 AM"
    gstrPersAmt = 0
    gdblTax = 0
    gdblTotAmt = 0
    gdblRebate = 0
    gStartCarTime = "12:00:00 AM"
    gstrProductType = ""
    gSysStartCarTime = "12:00:00 AM"
    gstrDespatchExe = ""
    gstrDespatch = ""
    Set gdbDespatch = Nothing
    gbolReprintEO = False
    gEOID = ""
    gbolRaiseEOReport = False
    gbolPreviewEO = False
    gbolIndEO = False
    gEOReportName = ""
    gstrReportPath = ""
    gstrEORptServer = ""
    gstrEORptDB = ""
    gstrEOPrtPwd = ""
    gstrEOEmailPath = ""
    gstrEOFaxPath = ""
    gbolPreviewCancel = False
    gstrEPRptLogin = ""
    gbolBeginTrans = False
    gbolPrintEO = False
    gbolAcceptEO = False
    Set gdbEmailConn = Nothing
    gstrEmail = ""
    Erase gstrVPWindows
    gbolFPEnable = False
    gdatStartTime = "12:00:00 AM"
    gstrModule = ""
    gstrSubModule = ""
    gstrProcess = ""
    gstrSubProcess = ""
    gstrForm = ""
    gstrProcessGrpID = ""
    gbolCreatPNR = False
    gstrCurrentPNR = ""
    gblnAmend = False
    Erase gstrFQPax
    gFPWidth = 0
    gFPHeight = 0
    gPadding = 0
    gSideBarWidth = 0
    gCustomVPHeight = 0
    gbolStartHBT = False
    gbolExitHBT = False
    gstrHBTURL = ""
    gstrPNRExpression = ""
    gintCheckERLineNum = 0
    Set gobjEO = Nothing
    Set gobjPreEO = Nothing
    gstrPreEOType = ""
    gbolEOAmend = False
    gbolIgnoreEO = False
    gStartOthSvcsTime = "12:00:00 AM"
    gSysStartOthSvcsTime = "12:00:00 AM"
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    gIntModuleType = 0
    gAssistedPanelHeight = 0
    gAssistedPanelWidth = 0
    gbolSPisHiddenByApp = False
    gbolPNRVIewerisHiddenByApp = False
    gbolAssistedisHiddenByApp = False
    
    gstrIPAddress = ""
    gstrHostName = ""
    
    
    
End Sub

Public Sub searchRestrictedArea(strCMC As String, strCountryCode As String, strAirportCode As String, ByRef strMalariaCode As String, ByRef strRestrictedCode As String, ByRef strRule As String)

    'JY - V1.2.9 20120104 - CR117 - EM Prompt for Restricted Countries and Airlines
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    strSQL = "Select TOP 1 Malaria, Restricted, "
    strSQL = strSQL & "case when AirportCode is not null then 1 else 2 end as [Rule] "
    strSQL = strSQL & "from tblRestrictedArea where "
    strSQL = strSQL & "CMC = '" & strCMC & "' and "
    strSQL = strSQL & "((AirportCode = '" & strAirportCode & "') or "
    strSQL = strSQL & "(AirportCode is NULL and CountryCode = '" & strCountryCode & "')) "
    strSQL = strSQL & "order by [Rule]"

    Set rs = gdbConn.Execute(strSQL)
    
    If rs.EOF = False Then
        strMalariaCode = rs!Malaria & ""
        strRestrictedCode = rs!Restricted & ""
        strRule = rs!Rule
    End If
    
    rs.Close
    Set rs = Nothing
                
End Sub

Public Sub searchRestrictedAirline(strCMC As String, strAirline As String, ByRef strRestrictedCode As String)

    'JY - V1.2.9 20120104 - CR117 - EM Prompt for Restricted Countries and Airlines
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    strSQL = "Select Restricted "
    strSQL = strSQL & "from tblRestrictedAirline where "
    strSQL = strSQL & "CMC = '" & strCMC & "' and "
    strSQL = strSQL & "AirlineCode = '" & strAirline & "'"

    Set rs = gdbConn.Execute(strSQL)
    
    If rs.EOF = False Then
        strRestrictedCode = rs!Restricted & ""
    End If
    
    rs.Close
    Set rs = Nothing
    
End Sub

Public Function searchNotification(strCMC As String, strType As String, strValue As String) As String

    'JY - V1.2.9 20120104 - CR117 - EM Prompt for Restricted Countries and Airlines
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    strSQL = "Select Message "
    strSQL = strSQL & "from tblNotification where "
    strSQL = strSQL & "CMC = '" & strCMC & "' and "
    strSQL = strSQL & "Type = '" & strType & "' and "
    strSQL = strSQL & "Value = '" & strValue & "'"

    Set rs = gdbConn.Execute(strSQL)
    
    If rs.EOF = False Then
       searchNotification = rs!Message & ""
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Public Function checkRestrictedRules() As String

    'JY - V1.2.9 20120104 - CR117 - EM Prompt for Restricted Countries and Airlines
    Dim i As Integer
    Dim j As Integer
    Dim strMalariaCode As String
    Dim strAirlineRestrictedCode As String
    Dim strAreaRestrictedCode As String
    Dim strAryMalariaCode() As String
    Dim strAryAirlineRestrictedCode() As String
    Dim strAryAreaRestrictedCode() As String
    Dim bolFound As Boolean
    Dim strMsg As String
    Dim strTemp As String
    Dim strAryAreaRestricted() As String
    Dim strAryAirlineRestricted() As String
    Dim strAryMalaria() As String
    Dim strRule As String

    strAryMalariaCode = Split("", vbCrLf)
    strAryMalaria = Split("", vbCrLf)
    strAryAirlineRestrictedCode = Split("", vbCrLf)
    strAryAirlineRestricted = Split("", vbCrLf)
    strAryAreaRestrictedCode = Split("", vbCrLf)
    strAryAreaRestricted = Split("", vbCrLf)
    
     
    For i = 1 To gobjPNR.AirSegCount
        With gobjPNR.AirSeg(i)
             If .Flown = False Then
                strMalariaCode = ""
                strAirlineRestrictedCode = ""
                strAreaRestrictedCode = ""
                strRule = ""
                searchRestrictedArea gobjPNR.CompInfo.WONum, .ArriveCountry, .ArriveAirport, strMalariaCode, strAreaRestrictedCode, strRule
                                                                                                            
                If strAreaRestrictedCode <> "" Then
                    'Check whether Area Restricted Code is stored in the distinct array before
                    bolFound = False
                    For j = 0 To UBound(strAryAreaRestrictedCode)
                        If strAreaRestrictedCode = strAryAreaRestrictedCode(j) Then
                           bolFound = True
                           Exit For
                        End If
                    Next
                    If bolFound = False Then
                        ReDim Preserve strAryAreaRestrictedCode(UBound(strAryAreaRestrictedCode) + 1)
                        strAryAreaRestrictedCode(UBound(strAryAreaRestrictedCode)) = strAreaRestrictedCode
                        ReDim Preserve strAryAreaRestricted(UBound(strAryAreaRestricted) + 1)
                        strAryAreaRestricted(UBound(strAryAreaRestricted)) = IIf(strRule = "1", .ArriveAirport & ", ", "") & GetCountryName(.ArriveCountry) & vbCrLf
                    Else
                        If InStr(1, strAryAreaRestricted(j), .ArriveAirport & ", ") = 0 Then
                           strAryAreaRestricted(j) = strAryAreaRestricted(j) & IIf(strRule = "1", .ArriveAirport & ", ", "") & GetCountryName(.ArriveCountry) & vbCrLf
                        End If
                    End If
                End If
                
                If strMalariaCode <> "" Then
                    'Check whether Malaria Code is stored in the distinct array before
                    bolFound = False
                    For j = 0 To UBound(strAryMalariaCode)
                        If strMalariaCode = strAryMalariaCode(j) Then
                           bolFound = True
                           Exit For
                        End If
                    Next
                    If bolFound = False Then
                        ReDim Preserve strAryMalariaCode(UBound(strAryMalariaCode) + 1)
                        strAryMalariaCode(UBound(strAryMalariaCode)) = strMalariaCode
                        ReDim Preserve strAryMalaria(UBound(strAryMalaria) + 1)
                        strAryMalaria(UBound(strAryMalaria)) = IIf(strRule = "1", .ArriveAirport & ", ", "") & GetCountryName(.ArriveCountry) & vbCrLf
                    Else
                        If InStr(1, strAryMalaria(j), .ArriveAirport & ", ") = 0 Then
                           strAryMalaria(j) = strAryMalaria(j) & IIf(strRule = "1", .ArriveAirport & ", ", "") & GetCountryName(.ArriveCountry) & vbCrLf
                        End If
                    End If
                End If

                searchRestrictedAirline gobjPNR.CompInfo.WONum, IIf(.OperatedBy = "", .Vendor, .OperatedBy), strAirlineRestrictedCode
                
                If strAirlineRestrictedCode <> "" Then
                    'Check whether Airline Restricted Code is stored in the distinct array before
                    bolFound = False
                    For j = 0 To UBound(strAryAirlineRestrictedCode)
                        If strAirlineRestrictedCode = strAryAirlineRestrictedCode(j) Then
                           bolFound = True
                           Exit For
                        End If
                    Next
                    If bolFound = False Then
                        ReDim Preserve strAryAirlineRestrictedCode(UBound(strAryAirlineRestrictedCode) + 1)
                        strAryAirlineRestrictedCode(UBound(strAryAirlineRestrictedCode)) = strAirlineRestrictedCode
                        ReDim Preserve strAryAirlineRestricted(UBound(strAryAirlineRestricted) + 1)
                        strAryAirlineRestricted(UBound(strAryAirlineRestricted)) = IIf(.OperatedBy = "", .Vendor, .OperatedBy) & vbCrLf
                    Else
                        If InStr(1, strAryAirlineRestricted(j), IIf(.OperatedBy = "", .Vendor, .OperatedBy)) = 0 Then
                           strAryAirlineRestricted(j) = strAryAirlineRestricted(j) & IIf(.OperatedBy = "", .Vendor, .OperatedBy) & vbCrLf
                        End If
                    End If
                End If
             End If
        End With
    Next
    
    For i = 0 To UBound(strAryAreaRestrictedCode)
        strTemp = searchNotification(gobjPNR.CompInfo.WONum, "RESTRICTEDAREA", strAryAreaRestrictedCode(i))
        If strTemp <> "" Then strMsg = strMsg & strTemp & vbCrLf & strAryAreaRestricted(i) & vbCrLf
    Next
                    
    For i = 0 To UBound(strAryAirlineRestrictedCode)
        strTemp = searchNotification(gobjPNR.CompInfo.WONum, "RESTRICTEDAIRLINE", strAryAirlineRestrictedCode(i))
        If strTemp <> "" Then strMsg = strMsg & strTemp & vbCrLf & strAryAirlineRestricted(i) & vbCrLf
    Next
    
    For i = 0 To UBound(strAryMalariaCode)
        strTemp = searchNotification(gobjPNR.CompInfo.WONum, "MALARIA", strAryMalariaCode(i))
        If strTemp <> "" Then strMsg = strMsg & strTemp & vbCrLf & strAryMalaria(i) & vbCrLf
    Next
    
    If strMsg <> "" Then
       strMsg = strMsg & "Do you want to document the NPAQ tags and historical remarks now?"
    End If
    
    checkRestrictedRules = strMsg
    
End Function
'Preethi - V1.2.12 20120528 - CR153 - PNR Touch Tracking - Auto Generate
Public Sub GetAutoAddTouchedCode(Module As String, CFA As String, ByRef PrimaryCode As String, ByRef SecondaryCode As String, ByRef ChargeIndicator As String, ByRef ContactCode As String)
Dim strSQL As String
Dim rsTouch As New ADODB.Recordset

strSQL = "SELECT * FROM [tblPNRTouchCodeAutoGenerate] WHERE [Module] = '" & Module & "' AND [CFA] = '" & CFA & "'"
Set rsTouch = gdbConn.Execute(strSQL)

With rsTouch
    If .EOF = False Then
        PrimaryCode = ![PrimaryCode]
        SecondaryCode = ![SecondaryCode]
        ChargeIndicator = ![ChargeIndicator]
        ContactCode = ![ContactCode]
    Else
       PrimaryCode = ""
       SecondaryCode = ""
       ChargeIndicator = ""
       ContactCode = ""
    End If
End With
rsTouch.Close
End Sub
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
Public Function GetClientMIValue(CN As String, FF As String) As Collection

 Dim strSQL As String
 Dim rs As New ADODB.Recordset
 
 Set GetClientMIValue = New Collection
 strSQL = "select FF, Value, Description from tblClientMIDropdown where CN = '" & CN & "' and FF in (" & FF & ")"
 Set rs = gdbConn.Execute(strSQL)
 
 While rs.EOF = False
    GetClientMIValue.Add rs![FF] & "*" & rs![value] & "/" & rs![Description]
    rs.MoveNext
 Wend
 
 rs.Close
 
End Function

'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
Public Sub PopulatecmbMI(ByRef cmbMI As MSForms.ComboBox, ByRef colMI As Collection, FF As String, strCellValue As String)

Dim intI As Integer
Dim intJ As Integer
Dim intK As Integer
Dim strTemp As String
Dim strTemp2 As String
Dim strValue As String
Dim strDesc As String

Dim bolFFFound As Boolean
Dim bolCellValueFound As Boolean

If cmbMI.style = fmStyleDropDownCombo Then cmbMI.Text = ""
cmbMI.Clear
cmbMI.ColumnCount = 2
cmbMI.ColumnWidths = "1 cm; 5 cm"
intI = 0
 
 
For intJ = 1 To colMI.Count
   'Collection will carry FF value in format 10*RG/Revenue Generating- Billable
   'strTemp = 10*RG/Revenue Generating- Billable
   strTemp = colMI(intJ)
   'strTemp2 = 10
   strTemp2 = Mid(strTemp, 1, InStr(1, strTemp, "*") - 1)
   If InStr(1, strTemp2, FF) And Len(strTemp2) = Len(FF) Then
   
      'strValue = RG
      strValue = ""
      'strTemp = RG/Revenue Generating- Billable
      strTemp = Mid(strTemp, Len(strTemp2) + 2, Len(strTemp))
      For intK = 1 To Len(strTemp)
          If Mid(strTemp, intK, 1) = "/" Then
             Exit For
          Else
             strValue = strValue & Mid(strTemp, intK, 1)
          End If
      Next
   
      'strDesc = Revenue Generating- Billable
      strDesc = ""
      strDesc = Mid(strTemp, Len(strValue) + 2, Len(strTemp))
      
      cmbMI.AddItem
      cmbMI.List(intI, 0) = strValue
      cmbMI.List(intI, 1) = strDesc
      intI = intI + 1
      
      bolFFFound = True
      
      If UCase(strValue) = UCase(strCellValue) Then
         bolCellValueFound = True
      End If

   End If
Next

If colMI.Count > 0 And bolFFFound = True And bolCellValueFound = False Then
   If (Trim(strCellValue) <> "") Then
       cmbMI.AddItem strCellValue
       cmbMI.List(cmbMI.ListCount - 1, 1) = ""
       colMI.Add FF & "*" & strCellValue & "/"
   End If
End If

End Sub


' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
'This function is used to terminate the running application in the task manager
Public Sub CloseApplication(sApplication As String)

    Dim objWMIService, colProcessList, objProcess As Object
    
    'Search task manager and terminate application desired
    Set objWMIService = GetObject("winmgmts:\\" & GetComputerName & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery _
                         ("Select * from Win32_Process Where Name = '" & sApplication & "'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next

End Sub

' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
'This function is used get the computer name

Public Function GetComputerName() As String
    
    Dim sResult As String * 255
    
    GetComputerNameA sResult, 255
    GetComputerName = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
    
End Function

' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
'This function is used to hide form
Public Function FunctHideForm(hwnd As Long)

Dim lngShow As Long
     
     If hwnd <> 0 Then
        lngShow = 0 'invisible the window/form
        If IsWindowVisible(hwnd) Then
            ShowWindow hwnd, lngShow
        End If
    End If
    
  
End Function

' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
'This function is used get show or hide the panels in Smart Point

Public Function showHideSmartPoint(bolShow As Boolean)

     Dim hwnd As String
     Dim lngShow As Long
     
     If bolShow = True Then
        lngShow = 1
     Else
        lngShow = 0
     End If
     
     hwnd = IsAppRunning("Smartpoint App - Window")
     
     If hwnd <> "" Then
            If lngShow = 0 Then
               If IsWindowVisible(hwnd) Then
                  ShowWindow hwnd, lngShow
                  gbolSPisHiddenByApp = True
               End If
            Else
               If gbolSPisHiddenByApp = True Then
                  ShowWindow hwnd, lngShow
                  gbolSPisHiddenByApp = False
               End If
            End If
     End If
     
     hwnd = IsAppRunning("PNR Viewer")
     
     If hwnd <> "" Then
            If lngShow = 0 Then
               If IsWindowVisible(hwnd) Then
                  ShowWindow hwnd, lngShow
                  gbolPNRVIewerisHiddenByApp = True
               End If
            Else
               If gbolPNRVIewerisHiddenByApp = True Then
                  ShowWindow hwnd, lngShow
                  gbolPNRVIewerisHiddenByApp = False
               End If
            End If
     End If
     
     hwnd = IsAppRunning("Assisted*")
            
     If hwnd <> "" Then
            If lngShow = 0 Then
               If IsWindowVisible(hwnd) Then
                  ShowWindow hwnd, lngShow
                  gbolAssistedisHiddenByApp = True
               End If
            Else
               If gbolAssistedisHiddenByApp = True Then
                  ShowWindow hwnd, lngShow
                  gbolAssistedisHiddenByApp = False
               End If
            End If
     End If
     
End Function

' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
Public Function MakeWinMove(ByVal hwnd, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Repaint As Boolean = True)
    Dim boolResult As Boolean
     MakeWinMove = MoveWindow(hwnd, Left, Top, Width, Height, Repaint)
     
    'If Not MoveWindow(hwnd, Left, Top, Width, Height, Repaint) Then
    '    Err.Raise vbObjectError, "MoveWindow", "MoveWindow returned error: H" & Hex$(Err.LastDllError)
    '
    'End If

End Function

' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
' Force the selected windows(by window handle) to appear as top most of the screen -
'Public Function MakeWinTopMost(StrWinName As String)
Public Function MakeWinTopMost(LngWinHwnd As Long)
'ZhiSam - SmartPoint - 02 Feb 2012
 'Dim lngHwnd As Long
  
'lngHwnd = IsAppRunning(StrWinName)
    'If lngHwnd = 0 Then
    If LngWinHwnd = 0 Then
        MsgBox "No application found", vbCritical, "MSG Error"
    Else
        SetWindowPos LngWinHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
    
    'SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS

        


End Function

' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
' Force the selected windows(by window name) to appear as top most of the screen
Public Function MakeWinNameTopMost(strWinName As String)
'Public Function MakeWinTopMost(LngWinName As Long)
'ZhiSam - SmartPoint - 02 Feb 2012
 Dim LngHwnd As Long
  
LngHwnd = IsAppRunning(strWinName)
    If LngHwnd = 0 Then
        MsgBox "No application found", vbCritical, "MSG Error"
    Else
        SetWindowPos LngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
    
    'SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS

        


End Function

' ZhiSam - V1.3.0 20120203 - CR 191 - Replace P&C with Smart Point V2.1
Public Function functRemoveFileFare() As Boolean

'If it is not OBT booking, remove Filed Fare that do not have FOP (which was created by Smart Point)
'If it is OBT booking, just ignore it
'Return true if detected File Fare is removed successfully
'Return false if no File Fare is removed

    Dim sOBTResponse As String
    Dim strResponse As String
    Dim strMsg As String
    Dim bolFXExists As Boolean

    ' ZhiSam - V1.2.20 20130528 - IR-54 - Bug Fix for Desktop to Remove File Fare (SyEx with Tpro) by add the process of refresh PNR
    gobjPNR.loadPNR
    'If sOBTResponse value is not empty(""), this mean it is not OBT booking. Otherwise it is OBT booking
    bolFXExists = False
    sOBTResponse = GetBookingTool()
    If sOBTResponse = "" Then
        'Remove Filed Fare that do not have FOP
        If gobjPNR.FareDataExists = True Then
            For i = 1 To gobjPNR.FiledFareCount
                If gobjPNR.FiledFare(i).PX(1).FOPType = "" Then
                    gobjHost.terminalEntry ("FX" & i)
                    bolFXExists = True
                End If
            Next
            ' enter ER to the GDS
            ' ZhiSam - V1.3.0 20120822 - Replace P&C with Smart Point V2.1
            ' Remove ER for SyEx Flow
            'If bolFXExists = True Then
            '    gobjHost.terminalEntry ("R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine))
            '    strResponse = gobjHost.terminalEntry("ER")
            '    If CheckResponse(strResponse, gstrPNRExpression, gintCheckERLineNum) = False Then
            '        For i = 0 To 1
            '            strResponse = gobjHost.terminalEntry("ER")
            '            If CheckResponse(strResponse, gstrPNRExpression, gintCheckERLineNum) = True Then
            '                Exit For
            '            End If
                        'if fail to enter ER for three times, then prompt error message
            '            If i = 1 Then
            '                strMsg = "Unable to end PNR. Response from GDS is " & Chr(13) & strResponse
            '                strMsg = strMsg & Chr(13) & "System will continue without cancel the fare with 'ER'."
            '                modMsgBox.OKMsg = "OK"
            '                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
            '            End If
            '        Next
            '    End If
            
            'End If
            
            'refresh the gobjPNR object after cancel file fare segment
            gobjPNR.loadPNR
            functRemoveFileFare = True
        End If
    Else
        functRemoveFileFare = False
    End If

End Function

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Public Function SearchGASalesByPC_Pax(ProductCode As String, PaxNum As Integer) As Boolean
    Dim lngI As Long
    
    SearchGASalesByPC_Pax = False
    For lngI = 1 To gobjPNR.GASaleRecordCount
        With gobjPNR.GASalesRecord(lngI)
            If .ProductCode = ProductCode And .PaxNum = PaxNum Then
                SearchGASalesByPC_Pax = True
                Exit For
            End If
        End With
    Next
End Function

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Public Function NonVoidCouponFound() As Boolean
    Dim intI As Integer
    Dim intJ As Integer
    
    NonVoidCouponFound = False
    For intI = 1 To gobjPNR.ETicketCount
        With gobjPNR.ETicket(intI)
            For intJ = 1 To .CouponCount
                If UCase(.Coupon(intJ).CouponStatus) <> "VOID" Then
                    NonVoidCouponFound = True
                    Exit For
                End If
            Next
        End With
    Next

End Function


' ZhiSam - V1.2.18 20120311 - CR-203 - Desktop to Create Retention Line and Update TAW to TAU (SyEx with Tpro)
Public Function bfunctCheckRTLine() As Boolean

Dim i As Integer
Dim item As ListItem
Dim strText As String
Dim strSegText As String
Dim strSegTextDS As String
Dim strRes As String
Dim strCmd As String
Dim bRTLineExist As Boolean
Dim intSegTest As Integer

    bRTLineExist = False
    strSegText = "RETENTION LINE"
    strSegText = UCase(strSegText)
    strSegTextDS = "****"
    
    
    gobjPNR.loadPNR
   
   ' ZhiSam - V1.2.20 20130528 - IR-54 - Bug Fix for Desktop to Create Retention Line (SyEx with Tpro) by read Retention Line from XML
    For i = 1 To gobjPNR.PaidDueCount
        'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
        If gobjPNR.PaidDue(i).SegType = "T" Then
            strRes = gobjPNR.PaidDue(i).FreeText
            If InStr(1, UCase(strRes), strSegText) > 0 Or InStr(1, UCase(strRes), strSegTextDS) Then
                bRTLineExist = True
                Exit For
            End If
        End If
    Next

    bfunctCheckRTLine = bRTLineExist

End Function

'ZhiSam - V1.2.23 20130911 - CR-289 - E-Invoice: Due and Paid Lines Date
'Return Retention line date
Public Function dtfunctRTDate() As Date

Dim i As Integer
Dim item As ListItem
Dim strText As String
Dim strSegText As String
Dim strSegTextDS As String
Dim strRes As String
Dim strCmd As String
Dim bRTLineExist As Boolean
Dim intSegTest As Integer
Dim dtRTdate As Date

    bRTLineExist = False
    strSegText = "RETENTION LINE"
    strSegText = UCase(strSegText)
    strSegTextDS = "****"
    
    
    gobjPNR.loadPNR
   
   ' ZhiSam - V1.2.20 20130528 - IR-54 - Bug Fix for Desktop to Create Retention Line (SyEx with Tpro) by read Retention Line from XML
    For i = 1 To gobjPNR.PaidDueCount
        'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
        If gobjPNR.PaidDue(i).SegType = "T" Then
            strRes = gobjPNR.PaidDue(i).FreeText
            If InStr(1, UCase(strRes), strSegText) > 0 Or InStr(1, UCase(strRes), strSegTextDS) Then
                dtRTdate = gobjPNR.PaidDue(i).SegDate
                'bRTLineExist = True
                Exit For
            End If
        End If
    Next

    'bfunctCheckRTLine = bRTLineExist
    dtfunctRTDate = dtRTdate

End Function
' ZhiSam - V1.2.18 20120311 - CR-203 - Desktop to Create Retention Line and Update TAW to TAU (SyEx with Tpro)
'Public Function strfunctLastTravelDate() As String
Public Function dtFunctLastTravelDate() As Date

Dim strResult As String
Dim i As Integer
Dim dDefaultDate As Date
Dim dDate As Date
Dim dLastestDate As Date
Dim sLastestDate As String
Dim bSegmentExist As Boolean

bSegmentExist = False
'default date
dDefaultDate = "1/1/1900"

dDate = dDefaultDate
dLastestDate = dDefaultDate

 ' ZhiSam - V1.2.20 20130528 - IR-54 - Bug Fix for Desktop to Create Retention Line (SyEx with Tpro) by add refresh PNR process
 gobjPNR.loadPNR
 If gobjPNR.AirSegCount > 0 Then
    bSegmentExist = True
    For i = 1 To gobjPNR.AirSegCount
        dDate = gobjPNR.AirSeg(i).DepartDateTime
        If dDate > dLastestDate Then
            dLastestDate = dDate
        End If
    Next
 End If
 
 If gobjPNR.CarSegCount > 0 Then
    bSegmentExist = True
    For i = 1 To gobjPNR.CarSegCount
        dDate = gobjPNR.CarSeg(i).EndDtTime
        If dDate > dLastestDate Then
            dLastestDate = dDate
        End If
    Next
 End If
 
 If gobjPNR.HotelSegCount > 0 Then
    bSegmentExist = True
    For i = 1 To gobjPNR.HotelSegCount
        dDate = gobjPNR.HotelSeg(i).CheckOutDate
        If dDate > dLastestDate Then
            dLastestDate = dDate
        End If
    Next
 End If

 If bSegmentExist = False Then
    dLastestDate = dDefaultDate
 End If

 dtFunctLastTravelDate = dLastestDate

End Function
'ZhiSam - V1.2.23 20130829 - CR 231 - Desktop SGHK To Disable the X function in Queue Module
Public Function EnableCloseButton(ByVal hwnd As Long, Enable As Boolean) As Integer
    Const xSC_CLOSE As Long = -10
    ' Check that the window handle passed is valid
    EnableCloseButton = -1
    If IsWindow(hwnd) = 0 Then Exit Function
    ' Retrieve a handle to the window's system menu
    Dim hMenu As Long
    hMenu = GetSystemMenu(hwnd, 0)
    ' Retrieve the menu item information for the close menu item/button
    Dim MII As MENUITEMINFO
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    
    If Enable Then
        MII.wID = xSC_CLOSE
    Else
        MII.wID = SC_CLOSE
    End If
        EnableCloseButton = -0
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    ' Switch the ID of the menu item so that VB can not undo the action itself
    Dim lngMenuID As Long
    lngMenuID = MII.wID
    
    If Enable Then
        MII.wID = SC_CLOSE
    Else
        MII.wID = xSC_CLOSE
    End If
        MII.fMask = MIIM_ID
    EnableCloseButton = -2
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Function
    ' Set the enabled / disabled state of the menu item
    If Enable Then
        MII.fState = (MII.fState Or MFS_GRAYED)
        MII.fState = MII.fState - MFS_GRAYED
    Else
        MII.fState = (MII.fState Or MFS_GRAYED)
    End If
    MII.fMask = MIIM_STATE
    EnableCloseButton = -3
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    SendMessage hwnd, WM_NCACTIVATE, True, 0
    EnableCloseButton = 0
End Function
'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration - Validate TMP Card, if is TMP card return true
Public Function IsTMPCard(strCardVendor As String, strCardNum As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim boolIsCard As Boolean
    Dim strTMPCard As String
    Dim strOptionSecCode As String
    
    boolIsCard = False
    strOptionSecCode = UCase(Trim(gstrAgcyCountryCode & "_" & gobjPNR.CompInfo.AgencyName))

    strSQL = "Select [OptionValue] from [tblModOptions] Where [OptionCode] = 'TMPCardNum' and "
    strSQL = strSQL & "[OptionSecCode]= '" & strOptionSecCode & "'"
    Set rs = gdbConn.Execute(strSQL)
    
    Do Until rs.EOF
        strTMPCard = rs![optionvalue]
        strTMPCard = decrypt(strTMPCard, gIntKey_TMP)
        'Check for the case if CardVendor exist and matching
        If UCase(strCardVendor) = Left(UCase(strTMPCard), 2) Then
            If Left(UCase(strCardNum), Len(Mid(UCase(strTMPCard), 3))) = Mid(UCase(strTMPCard), 3) Then
                'it is TMP Card (either CWT or JTB)
                boolIsCard = True
                Exit Do
            End If
        Else
        'Check for the case if CardVendor not exist
            If Left(UCase(strCardNum), Len(strTMPCard)) = UCase(strTMPCard) Then
             'it is TMP Card (either CWT or JTB)
                boolIsCard = True
                Exit Do
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    IsTMPCard = boolIsCard

End Function
 'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
'CR 304 -Test function - delete this after test
Sub TestTMP()

    Dim TestStr As String
    Dim FOP As String
    
    
'Case 1
'TestStr = "CX/DC/3644033-7283"
'TestStr = "CX/EC/3644033-7283"
TestStr = "CX/EC/448488600000-7789"

FOP = TestStr
'CX line - frmDespDI
'"CX/DC/3644033"
If UCase(Left(FOP, 2)) = "CX" Then
    If IsTMPCard(UCase(Mid(FOP, 4, 2)), UCase(Mid(FOP, 7))) Then
        lngMer = 0
    End If
End If
    
End Sub
