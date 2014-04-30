VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOthSvcs 
   Caption         =   "CWT TravelPro - Other Services"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   10455
   Begin MSAdodcLib.Adodc datProducts 
      Height          =   375
      Left            =   240
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEOAmend 
      Caption         =   "Amend EO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   8580
      Picture         =   "frmOthSvcs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Contin&ue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6720
      Picture         =   "frmOthSvcs.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo dbcProducts 
      Bindings        =   "frmOthSvcs.frx":0884
      DataSource      =   "datProducts"
      Height          =   360
      Left            =   2640
      TabIndex        =   6
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc datVendors 
      Height          =   375
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dbcVendors 
      Bindings        =   "frmOthSvcs.frx":089E
      DataSource      =   "datVendors"
      Height          =   360
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc datSelectedVendor 
      Height          =   375
      Left            =   6840
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblStepOne 
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2355
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Other Services - Accounting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10395
   End
   Begin VB.Label lblStepTwo 
      Caption         =   "Vendor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "frmOthSvcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrProductType As String
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date


Private Sub cmdCancel_Click()
    'ReadyToClose = True
    Unload Me
    'Call pRedisplayMenu
End Sub

Private Sub cmdContinue_Click()
Dim objForm As Form
Dim strMsg As String
Dim strFormName As String


datTouchEnd = Now
gbolEOAmend = False
preFormLoad
If dbcProducts.Text = "" Or dbcVendors.Text = "" Then
    'MsgBox "Need to select Product and Vendor", vbApplicationModal + vbExclamation
    strMsg = "Need to select Product and Vendor"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    Exit Sub
End If

cmdContinue.Visible = False
cmdCancel.Visible = False


datSelectedVendor.ConnectionString = gstrConn
datSelectedVendor.Mode = adModeRead
datSelectedVendor.CommandType = adCmdText
datSelectedVendor.RecordSource = "SELECT * FROM tblVendors WHERE [VendorNumber] =  '" & dbcVendors.BoundText & "'"
datSelectedVendor.Refresh


With datProducts.Recordset
    If Not .BOF Then .MoveFirst
    .Find "ProductCode = '" & dbcProducts.BoundText & "'"
    If Not .EOF Then
        Select Case datProducts.Recordset![Type]
            Case "HL"
                Set objForm = frmOSHotel
                strFormName = "frmOSHotel"
            Case "CT", "BT"
                Set objForm = frmOSAirTkt
                strFormName = "frmOSAirTkt"
            Case "CX"
                Set objForm = frmOSCarTxfr
                strFormName = "frmOSCarTxfr"
            Case "MS"
                Set objForm = frmOSMisc
                strFormName = "frmOSMisc"
            Case "VI"
                Set objForm = frmOSVisa
                strFormName = "frmOSVisa"
            Case "TR"
                Set objForm = frmOSOthTkt
                strFormName = "frmOSOthTkt"
            Case Else
                
                'MsgBox "There is an error in the database or this type of product is not supported." & Chr(13) _
                    & "Contact your system administrator for assistance."
                strMsg = "There is an error in the database or this type of product is not supported." & Chr(13) _
                    & "Contact your system administrator for assistance."
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                cmdContinue.Visible = True
                cmdCancel.Visible = True

                Exit Sub
                               
        End Select
        
        'gbolToMainMenu = False
       
        'If gbolToMainMenu = True Then
        '   gbolToMainMenu = False
        '   Unload Me
        '   Exit Sub
        'End If
      
        'objForm.Show 1, Me
        
        '150708JEN
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd

        
        
        
        
        
        Load objForm
        objForm.Show
        Do
            DoEvents
        Loop Until isLoaded(strFormName) = False
        
        

        Set objForm = Nothing
        pDisplayToFP "*DI"
    End If
End With

dbcProducts.Text = ""
dbcVendors.Text = ""

cmdContinue.Visible = True
cmdCancel.Visible = True

              
      


On Error Resume Next
Set frmOSHotel = Nothing
Set frmOSAirTkt = Nothing
Set frmOSCarTxfr = Nothing
Set frmOSMisc = Nothing
Set frmOSVisa = Nothing
Set frmOSOthTkt = Nothing

datFormLoadEnd = Now

End Sub

Private Sub cmdEOAmend_Click()
   frmEOAmend.Show
   'Unload Me
End Sub



Private Sub dbcProducts_Change()

datVendors.Mode = adModeRead
datVendors.CommandType = adCmdText
datVendors.RecordSource = "SELECT * FROM tblVendors WHERE [ProductCodes] LIKE '%" & dbcProducts.BoundText & "%' ORDER BY [VendorName]"
datVendors.Refresh


    If datVendors.Recordset.RecordCount > 0 Then
        Set dbcVendors.DataSource = datVendors
        dbcVendors.Text = ""
        dbcVendors.ListField = "VendorName"
        dbcVendors.BoundColumn = "VendorNumber"
        dbcVendors.Refresh
    End If


End Sub


Private Sub Form_Load()
Dim strTemp As String
Dim strMsg As String
'Dim a As CWT_GalileoPNR.VendorInfo
'Set a = New CWT_GalileoPNR.VendorInfo

'Set gobjPNR = New CWT_GalileoPNR.PNR

  'Me.Top = 25
  'Me.Left = (Screen.Width - Me.Width) - 25


    Dim oldParent As Long

    
    pSetGlobals gobjHost.AgentDIV   '20090202
    
    datFormLoadStart = Now
    gintY = 0
    gintX = 0
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)
    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0



datProducts.ConnectionString = gstrConn
datProducts.Mode = adModeRead
datProducts.CommandType = adCmdText
datProducts.RecordSource = "SELECT * FROM tblProductCodes ORDER BY [Description]"
datProducts.Refresh

'MsgBox (datProduct.Recordset.RecordCount)
'datProduct.Refresh

 Set dbcProducts.DataSource = datProducts
 dbcProducts.ListField = "Description"
 dbcProducts.BoundColumn = "ProductCode"
 dbcProducts.Refresh



'dbcProducts.BoundColumn = "ProductCode"
'datProduct.Refresh
'dbcProducts.Refresh

datVendors.ConnectionString = gstrConn

    RemoveMenus Me, False, False, _
        False, False, False, True, True

datFormLoadEnd = Now
If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart

End Sub

Private Sub RemoveMenus(frm As Form, _
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



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Cancel = Not ReadyToClose
End Sub

