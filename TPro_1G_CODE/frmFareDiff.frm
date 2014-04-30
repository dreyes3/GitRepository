VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFareDiff 
   Caption         =   "CWT TravelPro - Fare Diff NP.M (Account)"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   5310
   Begin VB.TextBox txtReason 
      Height          =   735
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Frame fraCC 
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   4815
      Begin VB.TextBox txtCCNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   960
         MaxLength       =   18
         TabIndex        =   23
         Tag             =   "NN"
         Top             =   0
         Width           =   2025
      End
      Begin VB.OptionButton optSmallCharge 
         Caption         =   "Small credit charge form"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   22
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox cmbCCType 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         ItemData        =   "frmFareDiff.frx":0000
         Left            =   120
         List            =   "frmFareDiff.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton optUATP 
         Caption         =   "UATP (long charge form)"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   20
         Top             =   840
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpCCExpDate 
         Height          =   360
         Index           =   1
         Left            =   3480
         TabIndex        =   24
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/yy"
         Format          =   63635459
         CurrentDate     =   36526
         MaxDate         =   73050
         MinDate         =   36526
      End
      Begin VB.Label Label6 
         Caption         =   "EXP"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   25
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   27
      Top             =   3120
      Width           =   2655
      Begin VB.OptionButton optINV 
         Caption         =   "INVAGT"
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   29
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton optCC 
         Caption         =   "CC"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame fraCC 
      BorderStyle     =   0  'None
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   4815
      Begin VB.OptionButton optUATP 
         Caption         =   "UATP (long charge form)"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   12
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox cmbCCType 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         ItemData        =   "frmFareDiff.frx":0040
         Left            =   120
         List            =   "frmFareDiff.frx":005C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton optSmallCharge 
         Caption         =   "Small credit charge form"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   11
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtCCNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   960
         MaxLength       =   18
         TabIndex        =   5
         Tag             =   "NN"
         Top             =   0
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker dtpCCExpDate 
         Height          =   360
         Index           =   0
         Left            =   3480
         TabIndex        =   10
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/yy"
         Format          =   63635459
         CurrentDate     =   36526
         MaxDate         =   73050
         MinDate         =   36526
      End
      Begin VB.Label Label6 
         Caption         =   "EXP"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   9
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   16
      Top             =   720
      Width           =   2655
      Begin VB.OptionButton optCC 
         Caption         =   "CC"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1215
      End
      Begin VB.OptionButton optINV 
         Caption         =   "INVAGT"
         Height          =   495
         Index           =   0
         Left            =   1200
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtTransFee 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      MaxLength       =   18
      TabIndex        =   8
      Tag             =   "NN"
      Top             =   2760
      Width           =   1185
   End
   Begin VB.TextBox txtTaxes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      MaxLength       =   18
      TabIndex        =   7
      Tag             =   "NN"
      Top             =   240
      Width           =   1185
   End
   Begin VB.TextBox txtDiffAmt 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      MaxLength       =   18
      TabIndex        =   6
      Tag             =   "NN"
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "Reason:"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "FOP (Trans Fee)"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "FOP (Air Ticket)"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Transaction Fee"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Taxes"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Diff amount"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmFareDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strCmd As String
   Dim strPNR As String
   datTouchEnd = Now
   If validData = False Then Exit Sub
   
   strCmd = "NP.M*DIFF AMOUNT " & Format(txtDiffAmt, "0.00") & " TAXES " & Format(txtTaxes, "0.00")
   gobjHost.terminalEntry strCmd
   
   If optCC(0).value Then
      strCmd = "NP.M*FOP(AIR TICKET)." & cmbCCType(0).Text & txtCCNum(0).Text & " EXP " & Format(dtpCCExpDate(0).value, "MMYY")
      gobjHost.terminalEntry strCmd
      If optSmallCharge(0).value Then
         strCmd = "NP.M*SMALL CREDIT CHARGE FORM"
      Else
         strCmd = "NP.M*UATP LONG CHARGE FORM"
      End If
      gobjHost.terminalEntry strCmd
   Else
      strCmd = "NP.M*FOP(AIR TKT).INVAGT"
      gobjHost.terminalEntry strCmd
   End If
   
   If txtTransFee <> "0" And txtTransFee <> "" Then
      strCmd = "NP.M*TRANSACTION FEE " & Format(txtTransFee, "0.00")
      gobjHost.terminalEntry strCmd

      If optCC(1).value Then
         strCmd = "NP.M*FOP(TRANS FEE)." & cmbCCType(1).Text & txtCCNum(1).Text & " EXP " & Format(dtpCCExpDate(1).value, "MMYY")
         gobjHost.terminalEntry strCmd
         If optSmallCharge(1).value Then
            strCmd = "NP.M*SMALL CREDIT CHARGE FORM"
         Else
            strCmd = "NP.M*UATP LONG CHARGE FORM"
         End If
         gobjHost.terminalEntry strCmd
      Else
         strCmd = "NP.M*FOP(TRANS FEE).INVAGT"
         gobjHost.terminalEntry strCmd
      End If
   
   End If
   'Added on 130807: V46- allow TC to enter reason for fare differences
   Dim lngC As Long
   Dim lngS As Long
   Dim strRest As String
   Dim strTemp As String
   
   strRest = Trim(txtReason) & " "
   strTemp = ""
   lngS = 1
   Do
   
        lngC = InStr(lngS, strRest, " ")
        strTemp = strTemp & Mid(strRest, lngS, lngC)
        strRest = Mid(strRest, lngC + 1)
        
        If Len(strTemp) > 87 Then
            gobjHost.terminalEntry "NP.M*" & Mid(strTemp, 1, InStrRev(strTemp, " ") - 1)
            strRest = strRest
            strTemp = ""
        End If
   Loop While InStr(1, strRest, " ")
   
   If Trim(strTemp) <> "" Then gobjHost.terminalEntry "NP.M*" & Trim(strTemp)
   
   'Added on 5/7/2005: Queue to Irene after FMR
   If UCase(gstrAgcyCountryCode) = "SG" Then
        strPNR = gobjPNR.RecLoc
        gobjHost.terminalEntry "R.TPRO NPM"
        gobjHost.terminalEntry "ER"
        gobjHost.terminalEntry "ER"
        gobjHost.terminalEntry "ER"
        gobjHost.terminalEntry "QEB/781P/77"
        gobjHost.terminalEntry "*" & strPNR
   End If
   
   '*NPM in focal point
    pClearWindow
    pDisplayToFP ("*NPM")
   
   
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareDiff, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareDiff, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareDiff, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
       
   
   
   
   Unload Me
End Sub

Private Sub Form_Load()
    Dim oldParent As Long
    Dim i As Integer
    datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
   
    For i = 0 To 1
      cmbCCType(i).Clear
      cmbCCType(i).AddItem ""
      cmbCCType(i).AddItem "VI"
      cmbCCType(i).AddItem "CA"
      cmbCCType(i).AddItem "AX"
      cmbCCType(i).AddItem "DC"
      cmbCCType(i).AddItem "TP"
      cmbCCType(i).AddItem "JC"
      cmbCCType(i).AddItem "EC"
      cmbCCType(i).listindex = 0
    
       With gobjPNR
          If .FOP_CCCode <> "" Then cmbCCType(i).Text = .FOP_CCCode
          txtCCNum(i).Text = .FOP_CCNum
          If .FOP_CCExpireDate > Now Then dtpCCExpDate(i).value = .FOP_CCExpireDate
       End With
    
       optCC(i).value = True
       optSmallCharge(i).value = True
    Next
   datFormLoadEnd = Now
   If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
End Sub



Private Sub optCC_Click(Index As Integer)
   fraCC(Index).Visible = True
End Sub

Private Sub optINV_Click(Index As Integer)
   fraCC(Index).Visible = False
End Sub


Private Function validData() As Boolean
Dim strErr As String
strErr = ""
   If txtTaxes = "" Then txtTaxes = "0"
      
   If txtDiffAmt = "0" Or txtDiffAmt = "" Then
      strErr = strErr & "Invalid diff amount" & vbCrLf
      'MsgBox "Invalid diff amount"
      'ValidData = False
      'Exit Function
   End If
   If optCC(0).value = True Then
   If fConvertZero(txtDiffAmt) <> 0 Then
    If cmbCCType(0).listindex = 0 Or txtCCNum(0).Text = "" Or dtpCCExpDate(0).value < Now Then
        strErr = strErr & "Incomplete/Invalid Credit Card details for FOP(Air Ticket)" & vbCrLf
    End If
   End If
   End If
   If optCC(1).value = True Then
   If fConvertZero(txtTransFee) <> 0 Then
    If cmbCCType(1).listindex = 0 Or txtCCNum(1).Text = "" Or dtpCCExpDate(1).value < Now Then
        strErr = strErr & "Incomplete/Invalid Credit Card details(Transaction fee)" & vbCrLf
    End If
   End If
   End If
   
   If strErr <> "" Then
    'MsgBox strErr, vbOKOnly, "frmFareDiff_Validation"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strErr, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    validData = False
   Else
    validData = True
   End If
   
End Function
