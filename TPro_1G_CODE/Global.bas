Attribute VB_Name = "Global"
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
                 ByVal hWnd1 As Long, _
                 ByVal hWnd2 As Long, _
                 ByVal lpsz1 As String, _
                 ByVal lpsz2 As String _
                 ) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" ( _
                ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function InvalidateRect Lib "user32" _
                (ByVal hWnd As Long, lpRect As Long, ByVal bErase As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" _
                Alias "GetPrivateProfileStringA" _
                (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpFileName As String) As Long
Declare Function WNetGetUser& Lib "Mpr" Alias "WNetGetUserA" (lpname As Any, ByVal lpUserName$, lpnLength&)
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Boolean
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Const VK_TAB = &H9
Public Const MF_BYPOSITION = &H400&
Public gobjPNR As CWT_GalileoPNR3.PNR
Public gobjSeatMaps As CWT_Galileo3.SeatMaps
Public gobjSQ As HostAccess.StructuredQuery
Public Const gstrID As String = "<Application><VendorId>XMDL</VendorId><VendorType>G</VendorType><SourceId>CARWPR</SourceId><SourceType>G</SourceType></Application>"
Public Const CONNECTION_FAIL As String = "Unable to connect to GDS"
Public gobjTE As HostAccess.TerminalEmulation
Public gobjFareQuotes As CWT_Galileo3.FareQuotes
Public gobjHost As CWT_Galileo3.GalileoHost
Public gobjLog As CWT_AppLog.AppLog
Public gstrTargetName As String
Public glngTargetHwnd As Long
Public gVPMDIHwnd As Long
Public Const GCL_HBRBACKGROUND = (-10)
Public Const gintLINK_MANUAL As Integer = 2    ' 2 - Manual (controls only)
Public gstrHostSession As String
Public Const gintDDE_TIMEOUT As Integer = 330  ' 'DDE timeout of 60 seconds.
Public gdbConn As New ADODB.Connection
Public gdbEitinConn As New ADODB.Connection
Public gstrConn As String
Public gstrEitinConn As String
Public Enum SQLType
    Select_ = 0
    Insert_ = 1
    Update_ = 2
    Delete_ = 3
End Enum
Type POINTAPI
    X As Long
    Y As Long
End Type
Public gstrAgcyCountryCode As String
Public gstrAgcyCurrCode As String
Public gstrAgcyCurrFormat As String
Public gstrAgcyCurrRule As String
Public gbytAgcyCurrDec As Byte
Public gsngAgcyCurrUnit As Single
Public gstrAgcyCityCode As String
Public gstrPCC As String
Public gstrHQPCC As String
Public gstrPFPCC As String
Public gstrAgcyAirportCode As String
Public gstrAgcyPhone As String
Public gstrPFCode As String
Public gbolCancelMove As Boolean
Public gbolMoveProfile As Boolean
Public gbolCancelProcess As Boolean
Public gbolPerformFF As Boolean
Public gbolGetProfileFromDB As Boolean
Public gstrFPResponse As String
Public gTrxnType As String
Public hMsVbLibToolBar As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public gstrPreviousText As String
Public gintX As Integer
Public gintY As Integer
Public gstrKeyword() As String
Public Type WINDOWPLACEMENT
    length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNORMAL = 1
Public Const gstrChecked = "þ"
Public Const gstrUnChecked = "q"
Public gbolBack As Boolean
Public gbolBackToRecap As Boolean
Public gbolBackToSI As Boolean
Public gbolWritingtoPNR As Boolean
Public gMode As String 'V= Viewpoint, F=Focalpoint

'************************************************
'Added during integration of TPro
'************************************************
Public gbolSkipAdult As Boolean
Public gbolOverrideFare As Boolean
Public gbolSelectFare As Boolean
Public gbolNetFare As Boolean
Public gStartFareQuoteTime As Date
Public gGetfareStart As Date
Public gGetfareEnd As Date
Public Const CdatDefaultDate As Date = #12:00:00 AM#
Public Const CstrdateFormat As String = "mm/dd/yyyy hh:nn:ss am/pm"
Public gFQSegID As Integer

'FMR
Public gstrFOP(2) As String
Public gdblAmtToCom As Double
Public gdblTaxToCom As Double
Public gdblAmtToPax As Double
Public gdblTaxToPax As Double
Public gbolFMR As Boolean
Public gstrFOPToCom As String
Public gstrCCVendor As String
Public gstrCCNum As String
Public gstrCCExpDate As Date
Public gstrPersCCVendor As String
Public gstrPersCCNum As String
Public gstrPersCCExpDate As Date
Public gstrPersAmt As Double
Public gdblTax As Double
Public gdblTotAmt As Double
Public gdblRebate As Double
' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer
' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Public gStartCarTime As Date
Public gstrProductType As String
Public gSysStartCarTime As Date
Public gstrDespatchExe As String
Public gstrDespatch As String
Public gdbDespatch As New ADODB.Connection
Public Const gconSkipChkTktQualifier = "AQ"
Public Const gconSkipChkTkt = "LT-SEND LTT"
Public Const CstrMonths = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC***"
Public gbolReprintEO As Boolean
Public gEOID As String
Public gbolRaiseEOReport As Boolean
Public gbolPreviewEO As Boolean
Public gbolIndEO As Boolean
Public gEOReportName As String
Public gstrReportPath As String
Public gstrEORptServer As String
Public gstrEORptDB As String
Public gstrEOPrtPwd As String
Public gstrEOEmailPath As String
Public gstrEOFaxPath As String
Public gbolPreviewCancel As Boolean
Public gstrEPRptLogin As String
Public gbolBeginTrans As Boolean
Public gbolPrintEO As Boolean
Public gbolAcceptEO As Boolean
Public gdbEmailConn As New ADODB.Connection
Public gstrEmail As String
Public gstrVPWindows() As String

'TEST
'Public Const HWND_NOTOPMOST = -2
'Public Const HWND_TOPMOST = -1
'Public Const MAX_PATH& = 260
'Public Const SW_HIDE = 0
'Public Const SW_MAXIMIZE = 3
'Public Const SW_MINIMIZE = 6
'Public Const SW_SHOW = 5
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOSIZE = &H1
'Public Const WM_CLOSE = &H10
'Public CurrenthWnd As Long
'Public Const flags = SWP_NOSIZE Or SWP_NOMOVE
'Public Declare Function IsWindowVisible& Lib "user32" (ByVal hwnd As Long)
Public Declare Function GetParent& Lib "user32" (ByVal hWnd As Long)
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnableWindow& Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long)
Public gbolFPEnable As Boolean

Public Const MAX_PATH = 260
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Logging Purposes
Public gdatStartTime As Date
Public gstrModule As String
Public gstrSubModule As String
Public gstrProcess As String
Public gstrSubProcess As String
Public gstrForm As String
Public gstrProcessGrpID As String
Public gbolCreatPNR As Boolean
Public gstrCurrentPNR As String

Public Const gconModAir As String = "AIR"
Public Const gconModHtl As String = "HOTEL"
Public Const gconModCar As String = "CAR"
Public Const gconModOthServ As String = "AUXILARY"
Public Const gconModMisc As String = "MISC"
Public Const gconModProfile As String = "PROFILE"

Public Const gconSModSearch As String = "SEARCH"
Public Const gconSModShop As String = "SHOP"
Public Const gconSModAvail As String = "AVAIL"
Public Const gconSModCreatePNR As String = "CREATE PNR"
Public Const gconSModTAU As String = "TAU"
Public Const gconSModTimatic As String = "TIMATIC"
Public Const gconSModRecap As String = "RECAP"
Public Const gconSModSeat As String = "SEAT"
Public Const gconSModServiceInfo As String = "SERVICE INFO"

Public Const gconSModFareQuoteRequest As String = "FARE QUOTE REQUEST"
Public Const gconSModFareQuote As String = "FARE QUOTE"
Public Const gconSModFileFare As String = "FILE FARE"

Public Const gconSModPreTrip As String = "PRETRIP"

Public Const gconSModFareDiff As String = "FARE DIFF"
Public Const gconSModItin As String = "SEND EITIN"
Public Const gconSModQueue As String = "QUEUE"
Public Const gconSModRmk As String = "REMARK"
Public Const gconSModIssueDoc As String = "ISSUE DOCS"

Public Const gconSModBkHtl As String = "BOOK HTL - HARP"
Public Const gconSModHtlSell As String = "HOTEL RENTAL"
Public Const gconSModHtlMI As String = "HOTEL MI"
Public Const gconSModBkCar As String = "CAR RENTAL"
Public Const gconSModCarMI As String = "CAR MI"

Public Const gconSModAux As String = "CREATE ACCT/EO"
Public Const gconSModReprintEO As String = "REPRINT EO"
Public Const gconSModDespatch As String = "DESPATCH VISA PROGRAM"
Public Const gconSModDesUpdate As String = "DESPATCH VISA UPDATE"
Public Const gconSModApproveChq As String = "APPROVE CHEQUE"

Public Const gconSModLTT As String = "LTT REMINDER"
Public Const gconSModPNR As String = "MODIFY PNR"
Public Const gconSModMCO As String = "MCO"
Public Const gconSModETR As String = "ETR"

Public Const gconFormLoad As String = "FORM LOAD"
Public Const gconTouch As String = "AGENT TOUCH"
Public Const gconProcessing As String = "SYSTEM PROCESSING"


Public Enum WindowPos
    vbTopMost = -1&
    vbNotTopMost = -2&
End Enum

Public gblnAmend As Boolean
Public gstrFQPax() As String

'JY - 20100604 - Focal Point screen width and height
Public gFPWidth As Long
Public gFPHeight As Long
Public gPadding As Long
Public gSideBarWidth As Long
Public gCustomVPHeight As Long

'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
Public gbolStartHBT As Boolean
Public gbolExitHBT As Boolean
Public gstrHBTURL As String

'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
Public Const gstrHKWebFareDecimalPoint As String = "#0.00"

Public gstrPNRExpression As String
Public gintCheckERLineNum As Integer

'JY - V1.2.6 20110916 - CR109 - Startup form for users to select database to be connected (AU ESC)
Public gstrDatabaseConnected As String
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'CC - V1.2.7 20111011 - CR114 - Enable Hotel MI for HBU
Public gbolHBUUser As Boolean

'CC - V1.2.8 20111028
Public Enum CmdType
    [NoCmd] = 0
    [DI] = 1
    [RI.S] = 2
    [RI] = 3
    [NP] = 4
    [DI_ITEM_VALUE] = 5 'e.g. DI.FT-MSX/FF34-AB/FF35-OTH/FF36-M/FF47-#CWT, DI_ITEM_VALUE are AB, OTH, M, #CWT
    [PE] = 6
End Enum

Public gdbAPPConn As New ADODB.Connection

' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
Public gIntModuleType As Integer

' ZhiSam - V1.3.1 20120203 - Replace P&C with Smart Point V2.1
' screen width and height for Smart Point
Public gAssistedPanelHeight As Long
Public gAssistedPanelWidth As Long

' ZhiSam - V1.3.1 20120203 - Replace P&C with Smart Point V2.1
' Window Control Declaration
Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As _
Long, ByVal nCmdShow As Long) As Long
Public gbolSPisHiddenByApp As Boolean
Public gbolPNRVIewerisHiddenByApp As Boolean
Public gbolAssistedisHiddenByApp As Boolean
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Boolean
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, rectangle As RECT) As Boolean
Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function PostMessage Lib "user32" (ByVal hWnd As Long, ByVal Msg As Integer, ByVal wParam As Long, ByVal lParam As Long) As Boolean
Declare Function SetActiveWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

' ZhiSam - V1.2.18 20120311 - CR-203 - Desktop to Create Retention Line and Update TAW to TAU (SyEx with Tpro)
Public gbolUpdateTAUDate As Boolean

'ZhiSam - V1.2.23 20130829 - CR 231 - Desktop SGHK To Disable the X function in Queue Module
Public Const SC_CLOSE As Long = &HF060&
Public Const MIIM_STATE As Long = &H1&
Public Const MIIM_ID As Long = &H2&
Public Const MFS_GRAYED As Long = &H3&
Public Const WM_NCACTIVATE As Long = &H86
 
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
 
'Public Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hWnd As Long, ByVal bRevert As Long) As Long
 
Public Declare Function GetMenuItemInfo Lib "user32" Alias _
    "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
 
Public Declare Function SetMenuItemInfo Lib "user32" Alias _
    "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
 
Public Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
 
Public Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long
