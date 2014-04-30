VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCarRmkMI 
   Caption         =   "CWT TravelPro - Add Remarks & MI"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   11115
   Begin VB.ComboBox cmbSeg 
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
      ItemData        =   "frmCarRmkMI.frx":0000
      Left            =   240
      List            =   "frmCarRmkMI.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   480
      Width           =   4635
   End
   Begin VB.ComboBox lstFare 
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fraMI 
      Height          =   4440
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   5295
      Begin VB.ComboBox cmbCommission 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cmbBedType 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtMI 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   6
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   13
         Tag             =   "BY-"
         Top             =   3720
         Width           =   465
      End
      Begin VB.ComboBox cmbRmType 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3360
         Width           =   2055
      End
      Begin VB.ComboBox cmbInvoice 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtMI 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   5
         Left            =   2400
         TabIndex        =   10
         Tag             =   "BY-"
         Top             =   3000
         Width           =   1185
      End
      Begin VB.ComboBox cmbBkAction 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtMI 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   3
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   8
         Tag             =   "BY-"
         Top             =   1560
         Width           =   885
      End
      Begin VB.ComboBox cmbBkMtd 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtMI 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   0
         Left            =   2400
         TabIndex        =   6
         Tag             =   "BY-"
         Top             =   360
         Width           =   1425
      End
      Begin VB.TextBox txtMI 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Index           =   1
         Left            =   2400
         TabIndex        =   5
         Tag             =   "BY-"
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label lblBedtype 
         Alignment       =   1  'Right Justify
         Caption         =   "Bed Type :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   4080
         Width           =   2235
      End
      Begin VB.Label lblBedNo 
         Alignment       =   1  'Right Justify
         Caption         =   "No of Bed(s) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   2235
      End
      Begin VB.Label lblRmCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Room Type :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   2235
      End
      Begin VB.Label lblPropNum 
         Alignment       =   1  'Right Justify
         Caption         =   "GDS Property Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   2235
      End
      Begin VB.Label lblBkAction 
         Alignment       =   1  'Right Justify
         Caption         =   "Booking Action:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   22
         Top             =   2280
         Width           =   1875
      End
      Begin VB.Label lblCommAmt 
         Alignment       =   1  'Right Justify
         Caption         =   "Comm(%):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label lblBkMtd 
         Alignment       =   1  'Right Justify
         Caption         =   "Booking Method:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   20
         Top             =   2640
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Prepaid/Referral?:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   480
         TabIndex        =   19
         Top             =   1920
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Commissionable:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   18
         Top             =   1200
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Reference Rate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   420
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Low Rate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   27
         Left            =   480
         TabIndex        =   16
         Top             =   780
         Width           =   1875
      End
   End
   Begin VB.TextBox txtMI 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Index           =   2
      Left            =   7800
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "BY-"
      Top             =   240
      Width           =   645
   End
   Begin VB.TextBox txtMI 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Index           =   4
      Left            =   7800
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "BY-"
      Top             =   2520
      Width           =   645
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
      Height          =   375
      Left            =   8880
      TabIndex        =   1
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   6000
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwRealECodes 
      Height          =   1815
      Left            =   5640
      TabIndex        =   27
      Top             =   600
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMissECodes 
      Height          =   1875
      Left            =   5640
      TabIndex        =   28
      Top             =   2880
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   3307
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Related Segment:"
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
      TabIndex        =   33
      Top             =   240
      Width           =   1785
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Realised Saving Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   5520
      TabIndex        =   30
      Top             =   240
      Width           =   2235
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Missed Saving Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   5760
      TabIndex        =   29
      Top             =   2520
      Width           =   1995
   End
End
Attribute VB_Name = "frmCarRmkMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim productType As String
Dim bookTool As String
'Dim datfrmStartMITime As Date
'Dim datProStartMITime As Date
Dim strHtlItinDB As String
Dim defaultComm As String
Dim msngRate As Single
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date


'Private Sub chkNo_Click()
'If chkNo.value = 1 Then
'    fraRemarks.Enabled = False
'    fraMI.Enabled = False
'    lvwRealECodes.Enabled = False
'    lvwMissECodes.Enabled = False
'    'lvwECodes.Enabled = False
'    txtMI(2).Enabled = False
'    txtMI(4).Enabled = False
'Else
'    fraRemarks.Enabled = True
'    fraMI.Enabled = True
    'lvwECodes.Enabled = True
'    lvwRealECodes.Enabled = True
'    lvwMissECodes.Enabled = True
'    txtMI(2).Enabled = True
'    txtMI(4).Enabled = True
'End If
'End Sub



Private Sub cmbSeg_Click()
Dim sngRate As Single
Dim intPosFare As Integer
Dim intPosPropNum As Integer
Dim intPosStatus As Integer
Dim strStatus As String
Dim strMsg As String

txtMI(0) = ""
If cmbSeg.listindex > -1 Then
lstFare.listindex = cmbSeg.listindex
intPosFare = InStr(lstFare, ".") + 1
intPosPropNum = InStr(lstFare, ";")
If intPosPropNum > 0 Then
   intPosStatus = InStr(intPosPropNum + 1, lstFare, ";")
End If

If intPosFare > 0 And intPosPropNum > 0 Then
    sngRate = Trim(Mid(lstFare, intPosFare, intPosPropNum - intPosFare))
End If
If intPosPropNum > 0 Then
   If intPosStatus > 0 Then
      txtMI(5) = Mid(lstFare, intPosPropNum + 1, intPosStatus - intPosPropNum - 1)
      strStatus = Mid(lstFare, intPosStatus + 1)
      If Trim(UCase(strStatus)) = "HK" Then
         cmbBkMtd.Text = "G-GDS"
      Else
         cmbBkMtd.Text = "M-MANUAL"
      End If
   End If
End If
If sngRate > 0 Then txtMI(1) = sngRate
msngRate = sngRate
End If
'added on 180406: check for passive htl segment confirmation no.
If productType = "HL" Then
    Dim strSegNo As String
    strSegNo = Trim(Left(cmbSeg, InStr(cmbSeg, ".") - 1))
    If pSegmentType(strSegNo) = "P" Then
        If pConfirmationNo(strSegNo) = False Then
            'MsgBox "Missing Confirmation Number on Hotel Segment!", vbCritical, "TravelPro - Confirmation number Check"
            strMsg = "Missing Confirmation Number on Hotel Segment!"
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        End If
    End If
End If

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDone_Click()

'datProStartMITime = Now
datTouchEnd = Now
If pValidData = False Then Exit Sub

pWriteToGDS

'If productType = "CX" Then
 '   If gStartCarTime = CdatDefaultDate Then
        'Call pAddToVBILog(gobjPNR.RecLoc, "Car Sell", datfrmStartMITime, datProStartMITime, "Car MI", , datfrmStartMITime)
 '   Else
 '       Call pAddToVBILog(gobjPNR.RecLoc, "Car Sell", datfrmStartMITime, datProStartMITime, "Car MI", , gStartCarTime)
  '  End If

'Else
 '   Call pAddToVBILog(gobjPNR.RecLoc, "Car Sell", datfrmStartMITime, datProStartMITime, "Hotel MI", , datfrmStartMITime)
'End If


       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       IIf(productType = "CX", gconModCar, gconModHtl), frmSideBar.cmbSelectType.Text, IIf(productType = "CX", gconSModCarMI, gconSModHtlMI), _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       IIf(productType = "CX", gconModCar, gconModHtl), frmSideBar.cmbSelectType.Text, IIf(productType = "CX", gconSModCarMI, gconSModHtlMI), _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       IIf(productType = "CX", gconModCar, gconModHtl), frmSideBar.cmbSelectType.Text, IIf(productType = "CX", gconSModCarMI, gconSModHtlMI), _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd


Unload Me

End Sub


'Private Sub cmdFreeRmkToItin_Click()
'If txtFreeRmk.Text <> "" And ValidIR(txtFreeRmk.Text) = True Then
'    lstItinRmks(1).AddItem txtFreeRmk.Text
'    txtFreeRmk.Text = ""
'End If
'End Sub



'Private Sub cmdItinRmksAddAll_Click()
'Dim strTemp As String

'With lstItinRmks(0)
'    For lngC = 0 To .ListCount - 1
'        strTemp = .List(lngC)
'   If strTemp = "" Then Exit For
'    If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
'                Load frmFareRmkFill
'                With frmFareRmkFill
'                    .lblRmkText = strTemp
'                    .Show 1, Me
'                    strTemp = .lblRmkText.Caption
'                 If strTemp = "" Then
           
'                    Else
'                        If ValidIR(strTemp) = True Then
'                        lstItinRmks(1).AddItem strTemp

'                        End If

'                    End If
'                    Unload frmFareRmkFill
'                End With
              
'            Set frmFareRmkFill = Nothing
            
'     Else
        
'        If ValidIR(strTemp) = True Then
'        lstItinRmks(1).AddItem .List(lngC)
        
'        End If
     
'     End If

'    Next lngC
'End With
'End Sub

'Private Sub cmdItinRmksAddOne_Click()
'Dim strTemp As String

'With lstItinRmks(0)

'If .SelCount > 0 And ValidIR(.Text) = True Then

'strTemp = .Text

'    If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
'                Load frmFareRmkFill
'                With frmFareRmkFill
'                    .lblRmkText = strTemp
'                    .Show 1, Me
'                    strTemp = .lblRmkText.Caption
'                    If strTemp = "" Then

'                    Else
'                        lstItinRmks(1).AddItem strTemp

'                    End If
'                    Unload frmFareRmkFill
'                End With

'            Set frmFareRmkFill = Nothing
'    Else
'    lstItinRmks(1).AddItem .Text

'    End If
'End If
'End With
'End Sub

'Private Sub cmdItinRmksRemove_Click()
'Dim intC As Integer
'With Me.lstItinRmks(1)
'For intC = .ListCount - 1 To 0 Step -1

'If .Selected(intC) = True Then
'    .RemoveItem intC
'End If
'Next intC

'End With
'End Sub

Private Sub Form_Load()
Dim rsRmk As ADODB.Recordset
Dim rsECodes As ADODB.Recordset
Dim strSql As String
Dim lngC As Long
Dim strPdtType As String
Dim oldParent As Long
Dim item As ListItem
Dim strMsg As String
Dim strTimeStamp As String

On Error GoTo LoadError

datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
Me.Move 0, 0
Me.Move frmSideBar.Width, 0

'datfrmStartMITime = Now
'SSTab1.Tab = 0

productType = gstrProductType

If productType = "CX" Then
    If gStartCarTime <> CdatDefaultDate Then
        Call pAddToVBILog(gobjPNR.RecLoc, "Car Sell", gStartCarTime, gSysStartCarTime, "Car Sell", , gStartCarTime)
    End If

End If


If productType = "CX" Then
    strSql = "Select * FROM tblProductRemarks WHERE ProductType = 'CX' and RmkType='C'"
    'chkNo.Caption = "No car reservation requested"
    'lblRmType.Visible = False
    'lblGI.Visible = False
    'cmbRmType.Visible = False
    strPdtType = "CAR"
    lblPropNum.Visible = False
    txtMI(5).Visible = False
    lblRmCode.Visible = False
    cmbRmType.Visible = False
    txtMI(6).Visible = False
    cmbBedType.Visible = False
    'txtMI(3).Visible = False
    lblBedNo.Visible = False
    lblBedtype.Visible = False
ElseIf productType = "HL" Then
    strSql = "Select * FROM tblProductRemarks WHERE ProductType = 'HL' and RmkType='H'"
    'chkNo.Caption = "No hotel reservation requested"
    'lblRmType.Visible = True
    'lblGI.Visible = True
    'cmbRmType.Visible = True
    strPdtType = "HTL"
    lblPropNum.Visible = True
    txtMI(5).Visible = True
    'txtMI(3).Visible = True
    lblRmCode.Visible = True
    lblBedNo.Visible = True
    lblBedtype.Visible = True
    cmbRmType.Visible = True
    txtMI(6).Visible = True
    cmbBedType.Visible = True
    'getHtlItinDB
    populateRmCode
    defaultComm = getDefaultComm
End If

If gobjPNR.CompInfo.MI = False Then
    lblLabels(0).Enabled = False
    lblLabels(27).Enabled = False
    lblLabels(3).Enabled = False
    lblLabels(6).Enabled = False
    lblBkAction.Enabled = False
    lblBkMtd.Enabled = False
    lblPropNum.Enabled = False
    lblRmCode.Enabled = False
    lblBedNo.Enabled = False
    lblBedtype.Enabled = False
    txtMI(0).Enabled = False
    txtMI(1).Enabled = False
    cmbBkAction.Enabled = False
    cmbBkMtd.Enabled = False
    txtMI(5).Enabled = False
    cmbRmType.Enabled = False
    txtMI(6).Enabled = False
    cmbBedType.Enabled = False
    txtMI(2).Enabled = False
    txtMI(4).Enabled = False
    lvwRealECodes.Enabled = False
    lvwMissECodes.Enabled = False
End If


'Set rsRmk = gdbConn.Execute(STRSQL)

'With rsRmk
'    Do Until .EOF
'        lstItinRmks(0).AddItem !Text
'        .MoveNext
'    Loop
'End With
'rsRmk.Close
'Set rsRmk = Nothing
Set gobjPNR = New CWT_GalileoPNR3.PNR
With gobjPNR
    Call .loadPNR
    'CC - V1.2.7 20111011 - CR114 - Enable Hotel MI for HBU
    If productType = "CX" Then
        For lngC = 1 To .CarSegCount
            cmbSeg.AddItem .CarSeg(lngC).TextCarSeg
            lstFare.AddItem UCase(Format(CStr(.CarSeg(lngC).SegNum), "@@. ")) & .CarSeg(lngC).RateAmt & ";;" & .CarSeg(lngC).Status
        Next
    End If
    If productType = "HL" Then
        For lngC = 1 To .HotelSegCount
            strTimeStamp = Right(.HotelSeg(lngC).ServInfo, 9)
            If gbolHBUUser = True And Left(strTimeStamp, 1) = "H" And IsNumeric(Right(strTimeStamp, 8)) Then
                'Do not add HBU hotel segment for HBU user
            Else
                cmbSeg.AddItem .HotelSeg(lngC).TextHtlSeg
                lstFare.AddItem UCase(Format(CStr(.HotelSeg(lngC).SegNum), "@@. ")) & .HotelSeg(lngC).RateAmount & ";" & .HotelSeg(lngC).PropertyNum & ";" & .HotelSeg(lngC).Status
            End If
        Next
    End If
 End With
 
 'cmbRmType.AddItem ""
 'cmbRmType.AddItem "SGLB"
 'cmbRmType.AddItem "TWNB"
 'cmbRmType.AddItem "DBLB"
 'cmbRmType.AddItem "KNGB"
 'cmbRmType.AddItem "STEE"
 'cmbRmType.AddItem "STUD"
 'cmbRmType.AddItem "1BDR"
 'cmbRmType.AddItem "2BDR"
 'cmbRmType.AddItem "3BDR"
 'cmbRmType.listindex = 0
 
 cmbBkMtd.AddItem ""
 cmbBkMtd.AddItem "G-GDS"
 cmbBkMtd.AddItem "M-MANUAL"

 cmbBkAction.AddItem ""
 cmbBkAction.AddItem "EB-Selfed Booked"
 cmbBkAction.AddItem "AC-Car Modified"
 cmbBkAction.AddItem "AH-Hotel Modified"
 
 'Preethi - V1.2.4 20110527 - CR 69 - Change Payment Type Verbiage in Hotel Car MI Screen
 cmbInvoice.AddItem "0-Referral"
 cmbInvoice.AddItem "1-Prepaid"
 cmbInvoice.listindex = 0
 
 cmbCommission.AddItem ""
 cmbCommission.AddItem "Yes"
 cmbCommission.AddItem "No"
 
'strSql = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='HS' OR tblExceptionCodes.ExceptionCodeGroup='HC') ORDER BY CAST(tblClientEC.EC AS integer)"
'strSql = "SELECT distinct(CAST(tblClientEC.EC AS integer)),description,exceptioncodegroup FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC and tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='HS' OR tblExceptionCodes.ExceptionCodeGroup='HC') ORDER BY CAST(tblClientEC.EC AS integer)"
strSql = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND tblExceptionCodes.ProdType='" & strPdtType & "' AND tblClientEC.ProdType='" & strPdtType & "' ORDER BY tblClientEC.EC"
Set rsECodes = gdbConn.Execute(strSql)
If Not rsECodes.EOF Then

     rsECodes.MoveFirst
     Do While Not rsECodes.EOF
        If rsECodes!ECCat = "R" Then
            Set item = lvwRealECodes.ListItems.Add(, , rsECodes!EC)
                  If rsECodes!Remarks = "" Then
                   item.SubItems(1) = rsECodes!Description
                  Else
                   item.SubItems(1) = rsECodes!Remarks
                  End If
            rsECodes.MoveNext
         Else
            Set item = lvwMissECodes.ListItems.Add(, , rsECodes!EC)
                  If rsECodes!Remarks = "" Then
                   item.SubItems(1) = rsECodes!Description
                  Else
                   item.SubItems(1) = rsECodes!Remarks
                  End If
            rsECodes.MoveNext
         End If
      Loop
    rsECodes.Close

Else
   
        rsECodes.Close
        strSql = "SELECT * FROM tblExceptionCodes where ProdType='" & strPdtType & "' and ECInd='C' order by ExceptionCode"
        Set rsECodes = gdbConn.Execute(strSql)
           If Not rsECodes.EOF Then rsECodes.MoveFirst
           
           Do While Not rsECodes.EOF
           If rsECodes!ECCat = "R" Then
              Set item = lvwRealECodes.ListItems.Add(, , rsECodes!exceptioncode)
              item.SubItems(1) = rsECodes!Description
              rsECodes.MoveNext
            Else
              Set item = lvwMissECodes.ListItems.Add(, , rsECodes!exceptioncode)
              item.SubItems(1) = rsECodes!Description
              rsECodes.MoveNext
            End If
            Loop
          rsECodes.Close
     
End If

Set rsECodes = Nothing

'detect self booking
If pSelfBook Then
    lblBkAction.Visible = True
    cmbBkAction.Visible = True
    lblBkMtd.Visible = False
    cmbBkMtd.Visible = False
Else
    lblBkAction.Visible = False
    cmbBkAction.Visible = False
    lblBkMtd.Visible = True
    cmbBkMtd.Visible = True
End If

   datFormLoadEnd = Now
   If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

Exit Sub
LoadError:
    'MsgBox "Error with Form load, system will exit back to menu", vbCritical
    strMsg = "Error with Form load, system will exit back to menu"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Form Load"
    Unload Me
End Sub
Private Function pSelfBook() As Boolean
Dim lngC As Long
Dim strTemp As String
Dim strPdtInd As String

pSelfBook = False
bookTool = ""

If productType = "HL" Then
    strPdtInd = "HOT"
Else
    strPdtInd = "CAR"
End If


For lngC = 1 To gobjPNR.GeneralRemarkCount
        If gobjPNR.GeneralRemark(lngC).Qualifier = "BT" Then
            'If InStr(gobjPNR.GeneralRemark(lngC).RemarkText, ".") - 1 > 0 Then strTemp = Mid(gobjPNR.GeneralRemark(lngC).RemarkText, 1, InStr(gobjPNR.GeneralRemark(lngC).RemarkText, ".") - 1)
            '    If strTemp = strPdtInd Then
            '        bookTool = Mid(gobjPNR.GeneralRemark(lngC).RemarkText, InStr(gobjPNR.GeneralRemark(lngC).RemarkText, ".") + 1)
                    bookTool = gobjPNR.GeneralRemark(lngC).RemarkText
                    pSelfBook = True
                    Exit For
            '    End If
        End If
    Next

End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    productType = Empty
    Set frmCarRmkMI = Nothing
End Sub



Private Sub cmbCommission_Click()
    With cmbCommission
        If .Text = "No" Then
            
            lblCommAmt.Visible = False
            txtMI(3).Visible = False
            txtMI(3).Text = "0"
        ElseIf .Text = "Yes" Then
            
            lblCommAmt.Visible = True
            txtMI(3).Visible = True
            txtMI(3).Text = defaultComm
        End If
    End With
End Sub



Private Sub lvwMissECodes_ItemClick(ByVal item As MSComctlLib.ListItem)
     txtMI(4).Text = lvwMissECodes.SelectedItem
End Sub

Private Sub lvwRealECodes_ItemClick(ByVal item As MSComctlLib.ListItem)
     txtMI(2).Text = lvwRealECodes.SelectedItem
End Sub

Private Sub txtFreeRmk_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii, "#$*()/.: ?@")
End Sub
Private Function pValidData() As Boolean
Dim strMsg As String
Dim bolEC As Boolean
Dim intX As Integer

'If chkNo.value = vbUnchecked Then
    If cmbSeg.listindex = -1 Then
        strMsg = strMsg & "Need to select Segment..." & Chr(13)
    End If
If gobjPNR.CompInfo.MI = True Then
    If txtMI(0).Text = "" Then strMsg = strMsg & "Need Reference Rate (MI)..." & Chr(13)
    If txtMI(1).Text = "" Then strMsg = strMsg & "Need Low Rate (MI)..." & Chr(13)
    If txtMI(2).Text = "" Then strMsg = strMsg & "Need Realised Saving Code (MI)..." & Chr(13)
    If txtMI(4).Text = "" Then strMsg = strMsg & "Need Missed Saving Code (MI)..." & Chr(13)
    End If
If cmbInvoice.Text = "" Then strMsg = strMsg & "Need Invoice/Referral (MI)..." & Chr(13)
If cmbCommission.Text = "" Then strMsg = strMsg & "Need Commissionable (MI)..." & Chr(13)
If cmbCommission.Text = "Yes" Then
    If txtMI(3) = "" Then
        strMsg = strMsg & "Need Commission Percentage (MI)..." & Chr(13)
    ElseIf CDec(txtMI(3)) > 100 Then
        strMsg = strMsg & "Invalid Commission Percentage (MI)..." & Chr(13)
    End If
End If
If gobjPNR.CompInfo.MI = True Then
                If cmbBkAction.Visible = True And cmbBkAction = "" Then strMsg = strMsg & "Need Booking Action (MI)..." & Chr(13)
                If cmbBkMtd.Visible = True And cmbBkMtd = "" Then strMsg = strMsg & "Need Booking Method (MI)..." & Chr(13)
                
                If txtMI(0) <> "" And txtMI(1) <> "" Then
                If fConvertZero(txtMI(0)) < fConvertZero(txtMI(1)) Then
                    strMsg = strMsg & "Reference Rate must be greater or equal to Low Rate (MI)..." & Chr(13)
                End If
                End If
            
            If productType = "HL" Then
                If cmbRmType = "" Then strMsg = strMsg & "Need Room Type (MI)..." & Chr(13)
                If cmbBedType = "" Then strMsg = strMsg & "Need Bed Type (MI)..." & Chr(13)
                If txtMI(6) = "" Then strMsg = strMsg & "Need No of Bed(s) (MI)..." & Chr(13)
                If txtMI(5) = "" Then strMsg = strMsg & "Need Hotel Property Number (MI)..." & Chr(13)
            End If
            
            If txtMI(2).Text <> "" Then
            bolEC = False
            For intX = 1 To lvwRealECodes.ListItems.Count
                If txtMI(2) = lvwRealECodes.ListItems.item(intX) Then
                    bolEC = True
                    Exit For
                End If
            Next intX
            If bolEC = False Then strMsg = strMsg & "Invalid Realised Saving Code..."
            End If
            
            If txtMI(4).Text <> "" Then
            bolEC = False
            For intX = 1 To lvwMissECodes.ListItems.Count
                If txtMI(4) = lvwMissECodes.ListItems.item(intX) Then
                    bolEC = True
                    Exit For
                End If
            Next intX
            If bolEC = False Then strMsg = strMsg & "Invalid Missed Saving Code..."
            End If
End If
'MI Validation
'If txtMI(0) <> "" And txtMI(1) <> "" Then
    'RSA If RF > SF then XX cannot be selected. If RF = SF , then XX must be selected
'    If (txtMI(0) - msngRate) > 0 Then
'       If Trim(txtMI(2).Text) = "XX" Then
'          strMsg = strMsg & "XX code in Realized Saving Code cannot be selected..." & Chr(13)
'       End If
'    ElseIf (txtMI(0) - msngRate) = 0 Then
'      If Trim(txtMI(2).Text) <> "XX" Then
'          strMsg = strMsg & "XX code in Realized Saving Code must be selected..." & Chr(13)
'       End If
'    End If
    'MSA If LF < SF, L cannot be selected. If LF = SF, then L must be selected
'    If (txtMI(1) - msngRate) < 0 Then
'       If Trim(txtMI(4).Text) = "L" Then
'          strMsg = strMsg & "L code in Missing Saving Code cannot be selected..." & Chr(13)
'       End If
'    ElseIf (txtMI(1) - msngRate) = 0 Then
'       If Trim(txtMI(4).Text) <> "L" Then
'          strMsg = strMsg & "L code in Missing Saving Code must be selected..." & Chr(13)
'       End If
'    End If
'End If
'End If

If strMsg <> "" Then
    'MsgBox strMsg
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    pValidData = False
Else
    pValidData = True
End If
End Function

Private Sub pWriteToGDS()
Dim strSegNo As String
Dim lngC As Long
Dim strCmd As String
Dim strDisplayNo As String
Dim strHarpNo As String
frmWait.Show

'If chkNo.value Then
'    gobjHost.terminalEntry "RI." & chkNo.Caption
'Else
    
strSegNo = Trim(Left(cmbSeg, InStr(cmbSeg, ".") - 1))

strCmd = ""
'For lngC = 0 To lstItinRmks(1).ListCount - 1
'   strCmd = strCmd & IIf(strCmd <> "", "+", "") & "RI.S" & strSegNo & "*" & lstItinRmks(1).List(lngC)
'Next

'gobjHost.terminalEntry strCmd

'strDisplayNo = cmbSeg.listindex + 1
strDisplayNo = "S" & Format(Mid(cmbSeg.Text, 1, InStr(cmbSeg.Text, ".")), "00")
        If gobjPNR.CompInfo.MI = True Then
            strCmd = "DI.FT-VLF/*" & strDisplayNo & "/" & txtMI(1)
        
            strCmd = strCmd & "+DI.FT-VRF/*" & strDisplayNo & "/" & txtMI(0).Text
            strCmd = strCmd & "+DI.FT-VEC/*" & strDisplayNo & "/" & txtMI(4).Text
            strCmd = strCmd & "+DI.FT-VFF30/*" & strDisplayNo & "/" & txtMI(2).Text
            'strCmd = strCmd & "+DI.FT-VFF32/*" & strDisplayNo & "/" & chkPreferredVendor.Caption
            'strCmd = strCmd & "+DI.FT-VFF1/*" & strDisplayNo & "/" & chkCommission.Caption
        End If
        
        If cmbCommission.Text = "No" Then
            strCmd = strCmd & IIf(strCmd = "", "", "+") & "DI.FT-NOCOMM/*" & strDisplayNo
        Else
            strCmd = strCmd & IIf(strCmd = "", "", "+") & "DI.FT-VCM/*" & strDisplayNo & "/P" & IIf(txtMI(3) = "", "0", txtMI(3))
        End If
        'strCmd = strCmd & "+DI.FT-VFF24/*" & strDisplayNo & "/" & IIf(chkGDS.Caption = "Y", "A", "P")
        'strCmd = strCmd & IIf(productType = "HL", "+DI.FT-VFF30/*" & strDisplayNo & "/" & cmbRmType.Text, "")
        
        If gobjPNR.CompInfo.MI = True Then
        strCmd = strCmd & "+DI.FT-VFF33/*" & strDisplayNo & "/" & Left(cmbInvoice, InStr(1, cmbInvoice, "-") - 1)
        End If
        
         'Preethi - V1.2.4 20110614 - CR 69 - Change Payment Type Verbiage in Hotel Car MI Screen
        'If Left(cmbInvoice, InStr(1, cmbInvoice, "-") - 1) = "1" Then
            'If productType = "CX" Then
                ' strCmd = strCmd & "+DI.FT-VPC/*" & strDisplayNo & "/19"
            'Else
                'strCmd = strCmd & "+DI.FT-VPC/*" & strDisplayNo & "/16"
            'End If
        'End If
        
        If gobjPNR.CompInfo.MI = True Then
                If cmbBkAction <> "" Then
                    strCmd = strCmd & "+DI.FT-VFF34/*" & strDisplayNo & "/" & Left(cmbBkAction, 2)
                Else
                    strCmd = strCmd & "+DI.FT-VFF34/*" & strDisplayNo & "/AB"
                End If
                If bookTool <> "" Then
                    strCmd = strCmd & "+DI.FT-VFF35/*" & strDisplayNo & "/" & bookTool
                Else
                    strCmd = strCmd & "+DI.FT-VFF35/*" & strDisplayNo & "/GAL"
                End If
                If cmbBkMtd = "" Then
                    strCmd = strCmd & "+DI.FT-VFF36/*" & strDisplayNo & "/S"
                Else
                    strCmd = strCmd & "+DI.FT-VFF36/*" & strDisplayNo & "/" & Left(cmbBkMtd, 1)
                End If
                    strCmd = strCmd & "+DI.FT-VFF39/*" & strDisplayNo & "/" & pSegmentType(strSegNo)
                'strCmd = strCmd & "+DI.FT-VFF33/*" & strDisplayNo & "/" & IIf(chkInvoice.Caption = "Y", "I", "R")
                '& IIf(productType = "HL", "+DI.FT-VFF31/*" & strDisplayNo & "/" & txtMI(3).Text, "")
                If txtMI(5) <> "" Then
                    strHarpNo = pGetHarpNoFromCodiff
                    LogHarp strHarpNo, txtMI(5)
                    If strHarpNo <> "" Then strCmd = strCmd & "+DI.FT-VFF42/*" & strDisplayNo & "/" & strHarpNo
                End If
                If cmbRmType <> "" Then
                    strCmd = strCmd & "+DI.FT-VTYP/*" & strDisplayNo & "/" & Left(cmbRmType, 1) & txtMI(6) & Left(cmbBedType, 1)
                End If
        End If
    gobjHost.terminalEntry strCmd
    gobjHost.terminalEntry "NP.SS*VBI " & gstrProductType & " MI"
    gobjHost.terminalEntry "R.TPRO " & gstrProductType & " MI+ER"
    gobjHost.terminalEntry "ER"
    gobjHost.terminalEntry "ER"
'End If

If frmWait.Visible = True Then Unload frmWait

End Sub
Private Function pGetHarpNoFromCodiff() As String
Dim strGalileoID As String
Dim strSql As String
Dim rs As ADODB.Recordset

strGalileoID = Trim(txtMI(5))
If Len(strGalileoID) < 5 Then
    strGalileoID = Format(strGalileoID, "00000")
End If

strSql = "SELECT * FROM [v_codif] where GDSPropID='" & strGalileoID & "' AND KEYTYPE='1G'"
Set rs = gdbConn.Execute(strSql)

If Not rs.EOF Then
pGetHarpNoFromCodiff = rs!harpno
End If
End Function
Private Sub LogHarp(harpno As String, propno As String)


Dim strSql As String


strSql = "insert into tblharptemp (pnr,searchtime,propertyid,harpid,agent) values('" & gobjPNR.RecLoc & "','" & Now & "','" & propno & "','" & harpno & "','" & gobjHost.AgentSine & "')"
gdbConn.Execute strSql

End Sub
'Private Function pGetHarpNo() As String
'    Dim intI As Integer
    
'    Dim Harp As HARPSearch.SearchServices
'    Dim ContactCode() As String
'    Dim PropID As New HARPSearch.PropertyIdentifier
'    Dim iResp As HARPSearch.HotelPropertyInfoSearchRespo
'    Set PropID = New HARPSearch.PropertyIdentifier
    
'    PropID.szValue = Trim(txtMI(5))
'    PropID.nType = 2
    
'    Set Harp = New HARPSearch.SearchServices
    
'    Harp.ConnDB True
'    Harp.getConn gdbConn
'    Harp.CallingModule "TPro"
'    Harp.strPNRPCC = gobjPNR.PCCOwner
'    Harp.strPNR = gobjPNR.RecLoc
    

'    Set iResp = Harp.HotelPropertyInfoSearch("integerasia", "integerasia", PropID, Format(Date, "yyyymmdd"), "PR", "OHG", "", ContactCode)
 
    
'    pGetHarpNo = ""
'    If iResp.iPropertyInfoOccurs Then
'        With iResp.iPropertyInfo
'           If .PropertyIdentifierCount > 1 Then
'                If .PropertyIdentifierList(1).nType = 0 Then
'                    pGetHarpNo = .PropertyIdentifierList(1).szValue
'                End If
'           End If
'        End With
'    End If
    
'End Function
Private Sub txtMI_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case Index
    Case 0, 1, 3
        KeyAscii = fAllowNumeric(KeyAscii, ".")
    Case 2, 6
        KeyAscii = fAllowNumeric(KeyAscii)
    'Case 3
    '    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Select

End Sub
Private Function pSegmentType(segNo As String) As String
Dim lngC As Long
Dim rsSegType As ADODB.Recordset

If productType = "HL" Then
    For lngC = 1 To gobjPNR.HotelSegCount
        If gobjPNR.HotelSeg(lngC).SegNum = segNo Then
            If gobjPNR.HotelSeg(lngC).Status <> "" Then
                
                Set rsSegType = gdbConn.Execute("select SegmentType from tblStatusCodes where StatusCode='" & gobjPNR.HotelSeg(lngC).Status & "'")
                If Not rsSegType.EOF Then
                    pSegmentType = rsSegType!SegmentType
                End If
                rsSegType.Close
            End If
            Exit For
        End If
    Next
Else
        For lngC = 1 To gobjPNR.CarSegCount
        If gobjPNR.CarSeg(lngC).SegNum = segNo Then
            If gobjPNR.CarSeg(lngC).Status <> "" Then
                Set rsSegType = gdbConn.Execute("select SegmentType from tblStatusCodes where StatusCode='" & gobjPNR.CarSeg(lngC).Status & "'")
                If Not rsSegType.EOF Then
                    pSegmentType = rsSegType!SegmentType
                End If
                rsSegType.Close
            End If
            Exit For
        End If
    Next
End If
End Function
Private Function pConfirmationNo(segNo As String) As Boolean
Dim lngC As Long
pConfirmationNo = True
    For lngC = 1 To gobjPNR.HotelSegCount
        If gobjPNR.HotelSeg(lngC).SegNum = segNo Then
            If gobjPNR.HotelSeg(lngC).ConfNum = "" Then
                pConfirmationNo = False
            End If
            Exit For
        End If
    Next
End Function
Private Sub populateRmCode()
Dim strSql As String
Dim rs As ADODB.Recordset
strSql = "SELECT * FROM [v_htlroomcode]"
Set rs = gdbConn.Execute(strSql)

While Not rs.EOF
    If UCase(Trim(rs!Type)) = "ROOM" And rs!InMI = True Then
        cmbRmType.AddItem rs!Code & "-" & rs!Description
    ElseIf UCase(Trim(rs!Type)) = "BED" And rs!InMI = True Then
         cmbBedType.AddItem rs!Code & "-" & rs!Description
    End If
rs.MoveNext
Wend
End Sub
Private Function getDefaultComm() As String
Dim strSql As String
Dim rs As ADODB.Recordset

strSql = "select Optionvalue from tblModOptions where OptionCode='HtlDefaultComm'"
Set rs = gdbConn.Execute(strSql)

If Not rs.EOF Then
    getDefaultComm = rs!optionvalue
End If

rs.Close
Set rs = Nothing

End Function

