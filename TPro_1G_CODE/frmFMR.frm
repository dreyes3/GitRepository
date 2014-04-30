VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFMR 
   Caption         =   "CWT TravelPro - FMR"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   9765
   Begin VB.CheckBox chkPC50 
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   30
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtTax 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   29
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox chkPC50 
      Height          =   255
      Index           =   2
      Left            =   9120
      TabIndex        =   28
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox chkPC50 
      Height          =   255
      Index           =   0
      Left            =   9120
      TabIndex        =   27
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   21
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtTax 
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   19
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtTax 
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   15
      Top             =   720
      Width           =   1215
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
      Index           =   2
      Left            =   3420
      MaxLength       =   18
      TabIndex        =   14
      Tag             =   "NN"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.ComboBox cmbFOP 
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
      Index           =   2
      ItemData        =   "frmFMR.frx":0000
      Left            =   1080
      List            =   "frmFMR.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1680
      Width           =   1515
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
      Index           =   2
      ItemData        =   "frmFMR.frx":001E
      Left            =   2640
      List            =   "frmFMR.frx":003A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
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
      Index           =   1
      Left            =   3420
      MaxLength       =   18
      TabIndex        =   10
      Tag             =   "NN"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.ComboBox cmbFOP 
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
      ItemData        =   "frmFMR.frx":005E
      Left            =   1080
      List            =   "frmFMR.frx":006B
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1200
      Width           =   1515
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
      ItemData        =   "frmFMR.frx":007C
      Left            =   2640
      List            =   "frmFMR.frx":0098
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
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
      Left            =   3360
      MaxLength       =   18
      TabIndex        =   6
      Tag             =   "NN"
      Top             =   720
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.ComboBox cmbFOP 
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
      ItemData        =   "frmFMR.frx":00BC
      Left            =   1080
      List            =   "frmFMR.frx":00C9
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1515
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
      ItemData        =   "frmFMR.frx":00DA
      Left            =   2640
      List            =   "frmFMR.frx":00F6
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker dtpCCExpDate 
      Height          =   360
      Index           =   0
      Left            =   5520
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
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
      Format          =   63700995
      CurrentDate     =   36526
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin MSComCtl2.DTPicker dtpCCExpDate 
      Height          =   360
      Index           =   1
      Left            =   5520
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
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
      Format          =   63700995
      CurrentDate     =   36526
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin MSComCtl2.DTPicker dtpCCExpDate 
      Height          =   360
      Index           =   2
      Left            =   5520
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
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
      Format          =   63700995
      CurrentDate     =   36526
      MaxDate         =   73050
      MinDate         =   36526
   End
   Begin VB.Label Label6 
      Caption         =   "Rebate/   Discount"
      Height          =   375
      Left            =   8880
      TabIndex        =   26
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Tax"
      Height          =   375
      Left            =   7920
      TabIndex        =   25
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Amount             (Inclusive Tax)"
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblTotTax 
      Height          =   375
      Left            =   7920
      TabIndex        =   23
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblTotAmt 
      Height          =   375
      Left            =   6600
      TabIndex        =   22
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "FOP3"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "FOP2"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "FOP1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmFMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbFOP_Click(Index As Integer)
   If cmbFOP(Index).Text = "CC" Then
      cmbCCType(Index).Visible = True
      txtCCNum(Index).Visible = True
      dtpCCExpDate(Index).Visible = True
      txtAmount(Index).Visible = True
      txtTax(Index).Visible = True
   ElseIf cmbFOP(Index).Text = "" Then
      cmbCCType(Index).Visible = False
      txtCCNum(Index).Visible = False
      dtpCCExpDate(Index).Visible = False
      txtAmount(Index).Visible = False
      txtAmount(Index) = 0
      txtTax(Index).Visible = False
      txtTax(Index) = 0
   Else
      cmbCCType(Index).Visible = False
      txtCCNum(Index).Visible = False
      dtpCCExpDate(Index).Visible = False
      txtAmount(Index).Visible = True
      txtTax(Index).Visible = True
   End If
End Sub

Private Sub cmdCancel_Click()
   gbolFMR = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim i As Integer
   
   gstrCCVendor = ""
   gstrCCNum = ""
   gstrPersCCVendor = ""
   gstrPersCCNum = ""
   
   If validData = False Then Exit Sub
   
   For i = 0 To 2
      gstrFOP(i) = ""
      If cmbFOP(i).Text = "CC" Then
         gstrFOP(i) = cmbCCType(i).Text & txtCCNum(i).Text & "*D" & Format(dtpCCExpDate(i).value, "MMYY") & "$" & Format(txtAmount(i).Text, gstrAgcyCurrFormat)
      ElseIf cmbFOP(i).Text <> "" Then
         gstrFOP(i) = cmbFOP(i).Text & "$" & Format(txtAmount(i).Text, gstrAgcyCurrFormat)
      End If
   Next
   
   gstrFOPToCom = cmbFOP(0).Text
   If gstrFOPToCom = "CC" Then
      gstrCCVendor = cmbCCType(0).Text
      gstrCCNum = txtCCNum(0).Text
      gstrCCExpDate = dtpCCExpDate(0).value
   End If
   If chkPC50(1).value = vbUnchecked And cmbFOP(1).Text = "CC" Then
      gstrPersCCVendor = cmbCCType(1).Text
      gstrPersCCNum = txtCCNum(1).Text
      gstrPersCCExpDate = dtpCCExpDate(1).value
      gstrPersAmt = CDec(txtAmount(1).Text)
   ElseIf chkPC50(2).value = vbUnchecked And cmbFOP(2).Text = "CC" Then
      gstrPersCCVendor = cmbCCType(2).Text
      gstrPersCCNum = txtCCNum(2).Text
      gstrPersCCExpDate = dtpCCExpDate(2).value
      gstrPersAmt = CDec(txtAmount(2).Text)
   End If
   
   'Added on 240806 for JPMC Normal ticket scenario where TF excluded in FMR although Commission>TF

   If chkPC50(1).value = vbChecked Then
        gdblRebate = gdblRebate + CSng(txtAmount(1))
   ElseIf chkPC50(2).value = vbChecked Then
        gdblRebate = gdblRebate + CSng(txtAmount(2))
   End If
   
   
   
   gdblAmtToCom = txtAmount(0)
   gdblAmtToPax = CDec(txtAmount(1).Text) + CDec(txtAmount(2).Text)
   gdblTaxToCom = CDec(txtTax(0).Text)
   gdblTaxToPax = CDec(txtTax(1).Text) + CDec(txtTax(2).Text) 'IIf(txtTax(1) = "", 0, txtTax(1)) + IIf(txtTax(2) = "", 0, txtTax(2))
   gbolFMR = True
   
 
   
   
   Unload Me
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   gbolFMR = False
   For i = 0 To 2
      cmbFOP(i).Clear
      cmbFOP(i).AddItem ""
      cmbFOP(i).AddItem "CC"
      cmbFOP(i).AddItem "INVAGT"
      cmbFOP(i).AddItem "CK"
      cmbFOP(i).listindex = 0
      
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
      
      dtpCCExpDate(i).value = DateSerial(2000, 1, 1)
      With gobjPNR
      If .FOPType = "CC" Then
         If .FOP_CCCode <> "" Then cmbCCType(i).Text = .FOP_CCCode
         txtCCNum(i).Text = .FOP_CCNum
         If .FOP_CCExpireDate > Now Then dtpCCExpDate(i).value = .FOP_CCExpireDate
      End If
      End With
      
   Next
   gdblRebate = 0
   gdblAmtToCom = 0
   gdblAmtToPax = 0
   gdblTaxToCom = 0
   gdblTaxToPax = 0
   
   txtTax(0).Text = gdblTax
   lblTotTax = gdblTax
   lblTotAmt = gdblTotAmt
   

   
End Sub

Private Function validData() As Boolean
   Dim i As Integer
   Dim j As Integer
   Dim strMsg As String
   
   For i = 0 To 2
      If txtAmount(i).Text = "" Then txtAmount(i).Text = 0
      If txtTax(i).Text = "" Then txtTax(i).Text = 0
   Next
   validData = True
   If cmbFOP(0).Text = "" Or cmbFOP(1).Text = "" Then
      'MsgBox "Invalid FOP1 or FOP2"
      strMsg = "Invalid FOP1 or FOP2"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
      validData = False
      Exit Function
   End If
   If cmbFOP(0).Text <> "CC" And cmbFOP(1).Text <> "CC" And cmbFOP(2).Text <> "CC" Then
      'MsgBox "At least 1 CC required"
      strMsg = "At least 1 CC required"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
      validData = False
      Exit Function
   End If
   
   For i = 0 To 2
   If cmbFOP(i).Text = "CC" Then
    If ValidCCNum(cmbCCType(i).Text, txtCCNum(i).Text) = False Then
         'MsgBox "Invalid CC Number"
         strMsg = "Invalid CC Number"
         modMsgBox.OKMsg = "OK"
         modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
         validData = False
         Exit For
         Exit Function
    End If
   End If
   Next
 
    If pSameFOP(0, 1) Or pSameFOP(0, 2) Or pSameFOP(1, 2) Then
        'MsgBox "FOP cannot be the same"
        strMsg = "FOP cannot be the same"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        validData = False
        Exit Function
    End If
 
   If txtAmount(0).Text = 0 Or txtAmount(1).Text = 0 Or (cmbFOP(2).Text <> "" And txtAmount(2).Text = 0) Then
      'MsgBox "Invalid amount"
      strMsg = "Invalid amount"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
      validData = False
      Exit Function
   End If
   If CDec(lblTotAmt) <> CDec(txtAmount(0)) + CDec(txtAmount(1)) + CDec(txtAmount(2)) Then
      'MsgBox "SUM OF FOPS NOT EQUAL TO AMT DUE"
      strMsg = "SUM OF FOPS NOT EQUAL TO AMT DUE"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      validData = False
      Exit Function
   End If
   If CDec(lblTotTax) <> CDec(txtTax(0)) + CDec(txtTax(1)) + CDec(txtTax(2)) Then
      'MsgBox "SUM OF taxes NOT EQUAL TO TOTAL TAXES"
      strMsg = "SUM OF taxes NOT EQUAL TO TOTAL TAXES"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      validData = False
      Exit Function
   End If
   
   For i = 0 To 2
    If cmbFOP(i).Text = "CC" Then
     If cmbCCType(i).listindex = 0 Or txtCCNum(i).Text = "" Or dtpCCExpDate(i).value < Now Then
         'MsgBox "Incomplete/Invalid Credit Card details"
         strMsg = "Incomplete/Invalid Credit Card details"
         modMsgBox.OKMsg = "OK"
         modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
         validData = False
         Exit Function
     End If
    End If
   Next i
End Function
Private Function pSameFOP(i As Integer, j As Integer) As Boolean

If cmbFOP(i) = cmbFOP(j) Then
    Select Case cmbFOP(i)
    Case "CC":
        If cmbCCType(i) = cmbCCType(j) And txtCCNum(i) = txtCCNum(j) Then
        pSameFOP = True
        Exit Function
        End If
    Case Else
            
            pSameFOP = True
            Exit Function
    End Select
Else
    pSameFOP = False
End If

End Function
Public Function ValidCCNum(Vendor As String, CCNum As String) As Boolean

Dim strCompare As String
Dim intX As Integer
Dim intY As Integer
Dim intZ As Integer
Dim intCD As Integer

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



