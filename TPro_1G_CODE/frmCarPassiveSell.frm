VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCarPassiveSell 
   Caption         =   "CWT TravelPro - Add Car Passive Segment"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   5535
   Begin VB.Frame Frame1 
      Caption         =   "Fill up all information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   5295
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back to Car Request"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   21
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox lstCarCom 
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "Co&ntinue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   8
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtCarType 
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         ToolTipText     =   "ICAR,ECAR,etc."
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtDailyRate 
         Height          =   300
         Left            =   2040
         TabIndex        =   7
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txtCurrency 
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         ToolTipText     =   "USD,GBP,etc."
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtConfNo 
         Height          =   300
         Left            =   2040
         TabIndex        =   5
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox cmbCarCom 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Text            =   "Combo2"
         Top             =   1800
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpPickUp 
         Height          =   300
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16973825
         CurrentDate     =   38561
      End
      Begin MSComCtl2.DTPicker dtpDropOff 
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16973825
         CurrentDate     =   38561
      End
      Begin VB.Label lblCarCom 
         Height          =   315
         Left            =   3000
         TabIndex        =   20
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Daily Rate:"
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
         Left            =   600
         TabIndex        =   18
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Currency Code:"
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
         TabIndex        =   17
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Conf No:"
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
         Left            =   960
         TabIndex        =   16
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Car Type:"
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
         Left            =   840
         TabIndex        =   15
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Car Company:"
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
         Left            =   480
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Pick-up City:"
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
         Left            =   600
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Drop-off Date:"
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
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Pick-up Date:"
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
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Car Passive Segment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCarPassiveSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date


Private Sub cmbCarCom_Click()
If cmbCarCom.listindex <> -1 Then
lstCarCom.listindex = cmbCarCom.listindex
lblCarCom.Caption = lstCarCom.List(lstCarCom.listindex)
End If
End Sub

Private Sub cmbCarCom_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii)
End Sub

Private Sub cmdBack_Click()
Unload Me
Load frmCarSellRequest
frmCarSellRequest.Show
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdContinue_Click()
Dim strCommand As String
Dim strResp As String
Dim strCarCommand As String
Dim strMsg As String
datTouchEnd = Now
If pValidData = False Then Exit Sub

strCarCommand = cmbCarCom.Text & "AK1" & Trim(txtCity) & Format(dtpPickUp.value, "ddmmm") _
                       & "-" & Format(dtpDropOff.value, "ddmmmyy") & Trim(txtCarType) & "/RT-" & txtCurrency _
                       & txtDailyRate & "/CF-" & txtConfNo
strCommand = "0CAR" & strCarCommand

strResp = gobjHost.terminalEntry(strCommand, True)

If InStr(Replace(Replace(strResp, " ", ""), vbCrLf, ""), UCase(Trim(strCarCommand))) > 0 Then
    gobjHost.terminalEntry "R.TPRO CARSELL+ER"
    gobjHost.terminalEntry "ER"
    gobjHost.terminalEntry "ER"
    pDisplayToFP "*R"
    gstrProductType = "CX"
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModCar, frmSideBar.cmbSelectType.Text, gconSModBkCar, _
    Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModCar, frmSideBar.cmbSelectType.Text, gconSModBkCar, _
    Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModCar, frmSideBar.cmbSelectType.Text, gconSModBkCar, _
    Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
    
    
    
    Unload Me
    Load frmCarRmkMI
    frmCarRmkMI.Show
Else
    'MsgBox "Unable to add passive car segment!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strResp
    strMsg = "Unable to add passive car segment!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strResp
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
    Exit Sub
End If


End Sub

Private Sub Form_Load()
    Dim oldParent As Long
    datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)
    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
    
    Call pPopulateControls
    datFormLoadEnd = Now
    If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
End Sub
Private Sub pPopulateControls()
Dim rs As ADODB.Recordset
Dim intI As Integer

intI = 0


Set rs = gdbConn.Execute("Select * from tblCarVendors order by code")
While Not rs.EOF
    intI = intI + 1
    cmbCarCom.AddItem rs!Code
    lstCarCom.AddItem rs!ShortName
rs.MoveNext
Wend
rs.Close
Set rs = Nothing


cmbCarCom.listindex = 0

dtpPickUp.value = Date
dtpDropOff.value = Date

End Sub
Private Function pValidData() As Boolean

Dim strMsg As String

If dtpPickUp.value < Date Or dtpDropOff.value < Date Then strMsg = strMsg & "Pick up date/Drop off date cannot be past..." & Chr(13)
If Len(txtCity.Text) <> 3 Then strMsg = strMsg & "Invalid City code" & Chr(13)
If Len(cmbCarCom.Text) <> 2 Then strMsg = strMsg & "Invalid Car company code" & Chr(13)
If Len(txtCurrency) <> 3 Then strMsg = strMsg & "Invalid Currency code" & Chr(13)
If Len(txtCarType) <> 4 Then strMsg = strMsg & "Invalid Car Type code" & Chr(13)
If txtDailyRate = "" Then strMsg = strMsg & "Need Rate" & Chr(13)
If txtConfNo = "" Then strMsg = strMsg & "Need Confirmation No" & Chr(13)
If strMsg = "" Then
    pValidData = True
Else
    'MsgBox strMsg, vbApplicationModal + vbExclamation, "TravelPro-Add Passive Car Segment"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    pValidData = False
End If

End Function



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub txtCarType_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii)
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii)
End Sub


Private Sub txtCurrency_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii)
End Sub



Private Sub txtDailyRate_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub
