VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAquaItinRmk 
   Caption         =   "AQUA ITIN Remarks"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   Icon            =   "frmAquaItinRmk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHisRmk2 
      Height          =   285
      Left            =   240
      MaxLength       =   108
      TabIndex        =   10
      Top             =   6120
      Width           =   7455
   End
   Begin VB.TextBox txtHisRmk1 
      Height          =   285
      Left            =   240
      MaxLength       =   108
      TabIndex        =   9
      Top             =   5760
      Width           =   7455
   End
   Begin VB.TextBox txtRmkText 
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   7
      Top             =   4940
      Visible         =   0   'False
      Width           =   2800
   End
   Begin VB.TextBox txtRmkText 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   4940
      Visible         =   0   'False
      Width           =   2800
   End
   Begin MSComctlLib.ListView lvwRemarkText 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   5318
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   531
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Remark Type"
         Object.Width           =   3702
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remark Description"
         Object.Width           =   4940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remark Text"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Historical Remark 1"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Historical Remark 2"
         Object.Width           =   4939
      EndProperty
   End
   Begin MyCommandButton.MyButton cmdPrevious 
      Height          =   360
      Left            =   240
      TabIndex        =   14
      Top             =   6480
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   635
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   14215660
      Caption         =   "&Previous Module"
      Depth           =   1
      GradientType    =   2
   End
   Begin MyCommandButton.MyButton cmdAdd 
      Height          =   360
      Left            =   1800
      TabIndex        =   15
      Top             =   6480
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   635
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   14215660
      Caption         =   "&Add"
      Depth           =   1
      GradientType    =   2
   End
   Begin MyCommandButton.MyButton cmdDelete 
      Height          =   360
      Left            =   2760
      TabIndex        =   16
      Top             =   6480
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   635
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   14215660
      Caption         =   "&Delete"
      Depth           =   1
      GradientType    =   2
   End
   Begin MyCommandButton.MyButton cmdRefresh 
      Height          =   360
      Left            =   3720
      TabIndex        =   17
      Top             =   6480
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   635
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   14215660
      Caption         =   "&Refresh"
      Depth           =   1
      GradientType    =   2
   End
   Begin MyCommandButton.MyButton cmdFinish 
      Height          =   360
      Left            =   5760
      TabIndex        =   18
      Top             =   6480
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   635
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   14215660
      Caption         =   "&Finish"
      Depth           =   1
      GradientType    =   2
   End
   Begin MyCommandButton.MyButton cmdExit 
      Height          =   360
      Left            =   6720
      TabIndex        =   13
      Top             =   6480
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   635
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   14215660
      Caption         =   "&Cancel"
      Depth           =   1
      GradientType    =   2
   End
   Begin MSForms.ComboBox cmbDescription 
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   7455
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "13150;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbRmkType 
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   3360
      Width           =   3135
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "5530;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblHisRmk 
      Caption         =   "Historical Remark 1 &&  2:"
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
      TabIndex        =   8
      Top             =   5400
      Width           =   3585
   End
   Begin VB.Label lblRtext 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   5
      Top             =   5010
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblRtext 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   5010
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblRmkText 
      Caption         =   "Remark Text:"
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
      TabIndex        =   3
      Top             =   4560
      Width           =   1665
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description:"
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
      TabIndex        =   2
      Top             =   3840
      Width           =   1665
   End
   Begin VB.Label lblRmkType 
      Caption         =   "Remark Type:"
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
      Top             =   3360
      Width           =   1665
   End
End
Attribute VB_Name = "frmAquaItinRmk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Preethi - V1.2.8  20111031 - CR101 - Aqua Itin Remark Screen
 Dim mobjAQRmkX As Remarks
 Dim mobjAQRmk As Remarks
 Dim mobjAddedRmk As Remarks
 Const mconMaxRmkLen As Integer = 80
 Const mconMaxRmkIndex As Integer = 1
 Dim flgChange As Boolean
 
Private Sub cmbDescription_Change()
ClearRmkText
End Sub

Private Sub cmbDescription_Click()
flgChange = True
ClearRmkText
If cmbDescription.listindex <> -1 Then
   If cmbDescription.List(cmbDescription.listindex, 1) <> "" Then
      PopulateRmkText cmbDescription.List(cmbDescription.listindex, 1)
   End If
   txtHisRmk1 = cmbDescription.List(cmbDescription.listindex, 2)
   txtHisRmk2 = cmbDescription.List(cmbDescription.listindex, 3)
End If
End Sub

Private Sub cmbRmkType_Click()
flgChange = True
If cmbRmkType.listindex > -1 Then
   If cmbRmkType.List(cmbRmkType.listindex, 1) <> "" Then
      PopulateRmkDesc cmbRmkType.List(cmbRmkType.listindex, 1)
   End If
End If
End Sub

Private Sub cmdAdd_Click()
Dim item As ListItem
Dim intI As Integer
Dim blnValid As Boolean
Dim blnExist As Boolean
Dim strRmk As String

intI = 0
blnExist = False
blnValid = False
strRmk = ""
  
blnValid = validData(strRmk)
If blnValid = True Then
'CC - Remove
'  For intI = 0 To mconMaxRmkIndex
'        If lblRtext(intI).Visible = True Then
'            strRmk = strRmk & lblRtext(intI).Caption
'        End If
'        If txtRmkText(intI).Visible = True Then
'           strRmk = strRmk & txtRmkText(intI).Text
'        End If
'  Next
'  If strRmk <> "" Then
'     For intI = 1 To lvwRemarkText.ListItems.Count
'        If UCase(lvwRemarkText.ListItems.item(intI).SubItems(3)) = UCase(strRmk) Then
'           blnExist = True
'           Exit For
'       End If
'     Next
'     If blnExist = False Then
'       If cmbRmkType.listindex > -1 Then
          Set item = lvwRemarkText.ListItems.Add()
'          If cmbDescription.listindex <> -1 Then
             item.SubItems(1) = cmbRmkType.value
             item.SubItems(2) = cmbDescription.value
             item.SubItems(3) = strRmk
             item.SubItems(4) = txtHisRmk1.Text
             item.SubItems(5) = txtHisRmk2.Text
'          End If
'       End If
       ClearData
'     Else
'       modMsgBox.OKMsg = "OK"
'       modMsgBox.sMsgBox gVPMDIHwnd, "Remark Already Exist", vbOKOnly + vbDefaultButton1, "AQUA Itin Remark - Error"
'     End If
'  End If
End If
End Sub

Private Sub cmdDelete_Click()
Dim intI As Integer
Dim blnDel As Boolean

intI = 0
blnDel = False

For intI = lvwRemarkText.ListItems.Count To 1 Step -1
   If lvwRemarkText.ListItems.item(intI).Checked = True Then
      lvwRemarkText.ListItems.Remove (intI)
       blnDel = True
   End If
   If lvwRemarkText.ListItems.Count = 0 Then Exit For
Next intI
If blnDel = False Then
   modMsgBox.OKMsg = "OK"
   modMsgBox.sMsgBox gVPMDIHwnd, "No Remark selected to Delete", vbOKOnly + vbDefaultButton1, "AQUA Itin Remark - Error"
End If
End Sub

Private Sub cmdExit_Click()
    Set mobjAQRmkX = Nothing
    Set mobjAQRmk = Nothing
    Set mobjAddedRmk = Nothing
    
    gbolCancelProcess = True
    Unload Me
End Sub

Private Sub cmdFinish_Click()

Dim strDelLineForAQX As String
Dim strDelLineForAH As String
Dim colDeletedRmk As Collection
Dim lngI As Long
Dim strDelLineForAQ As String
Dim colAHRmk As Collection
Dim colHistoricalRmk As Collection
Dim colNewlyAddedRmk As Collection
Dim colNPCmd As Collection
Dim strDelLineNum As String
Dim intJ As Integer
Dim strErrMsg As String
Dim strMsg As String
Dim strFailCmd As String
Dim blnExist As Boolean

Set colDeletedRmk = New Collection
Set colAHRmk = New Collection
Set colHistoricalRmk = New Collection
Set colNewlyAddedRmk = New Collection
Set colNPCmd = New Collection

strDelLineForAH = ""
strDelLineForAQX = ""
strDelLineForAQ = ""
strDelLineNum = ""
strResponse = ""
strFailCmd = ""
lngI = 0
strMsg = ""
strErrMsg = ""

strDelLineForAQX = GetLineNumOfAQXToDel(mobjAQRmkX)
For lngI = 1 To mobjAddedRmk.RemarkCount
      With mobjAddedRmk.Remark(lngI)
        strDelLineForAH = IIf(strDelLineForAH = "", .Number, strDelLineForAH & "." & .Number)
    End With
Next

Set colDeletedRmk = GetDeletedRmk(mobjAddedRmk)
Set colNewlyAddedRmk = GetRmkFromListView(mobjAddedRmk, colAHRmk, colHistoricalRmk)

strDelLineForAQ = GetLineNumOfAQToDel(mobjAQRmk, colDeletedRmk)

strDelLineNum = strDelLineForAH
If Len(strDelLineForAQX) > 0 Then
   strDelLineNum = IIf(strDelLineNum = "", strDelLineForAQX, strDelLineNum & "." & strDelLineForAQX)
End If
If Len(strDelLineForAQ) > 0 Then
   strDelLineNum = IIf(strDelLineNum = "", strDelLineForAQ, strDelLineNum & "." & strDelLineForAQ)
End If
strDelLineNum = sortInt(strDelLineNum)
strDelLineNum = FormatedLineNum(strDelLineNum)

If strDelLineNum <> "" Then
   strErrMsg = gobjHost.terminalEntry("NP." & strDelLineNum & "@")
   If (Right(strErrMsg, 1) = "*") = False Then
       strMsg = "Failed to Delete NP Line."
       strMsg = strMsg & vbCrLf & "Desktop will Ignore & Retrieve (IR) the PNR" & vbCrLf
       strMsg = strMsg & vbCrLf & "Galileo Command:" & vbCrLf & "NP." & strDelLineNum & "@"
       strMsg = strMsg & vbCrLf & vbCrLf & "Galileo Response: " & vbCrLf & strErrMsg
       
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strErrMsg, vbOKOnly + vbDefaultButton1, "AQUA Itin Remark - Error"
       gobjHost.terminalEntry "IR"
       Exit Sub
   End If
End If

For intJ = 1 To colDeletedRmk.Count
    blnExist = False
    For lngI = 1 To mobjAQRmk.RemarkCount
        With mobjAQRmk.Remark(lngI)
           If UCase(colDeletedRmk(intJ)) = UCase(.RemarkText) Then
              blnExist = True
              Exit For
          End If
        End With
    Next
    If blnExist = False Then
       colNPCmd.Add ("NP.AQ-" & colDeletedRmk(intJ) & "X")
    End If
Next
For intJ = 1 To colNewlyAddedRmk.Count
    colNPCmd.Add ("NP.AQ-" & colNewlyAddedRmk(intJ))
Next
For intJ = 1 To colHistoricalRmk.Count
    colNPCmd.Add ("NP.H**" & colHistoricalRmk(intJ))
Next
For intJ = 1 To colAHRmk.Count
    colNPCmd.Add ("NP.AH-" & colAHRmk(intJ))
Next
 
strErrMsg = ""
strErrMsg = SendGDSCmd(colNPCmd, NP, strFailCmd)
If strErrMsg <> "" Then
   strMsg = "Failed to Insert NP Line."
   strMsg = strMsg & vbCrLf & "Desktop will Ignore & Retrieve (IR) the PNR" & vbCrLf
   strMsg = strMsg & vbCrLf & "Galileo Command: " & vbCrLf & strFailCmd
   strMsg = strMsg & vbCrLf & vbCrLf & "Galileo Response: " & vbCrLf & strErrMsg
           
  modMsgBox.OKMsg = "OK"
  modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "AQUA Itin Remark - Error"
  gobjHost.terminalEntry "IR"
  Exit Sub
End If

'gobjHost.terminalEntry "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
'gobjHost.terminalEntry "ER"
'gobjHost.terminalEntry "ER"
'gobjHost.terminalEntry "ER"
'loadPNR

If ENDPNR(True) = True Then
    loadPNR
    displayPNRinBar
End If

Set colDeletedRmk = Nothing
Set colNewlyAddedRmk = Nothing
Set colHistoricalRmk = Nothing
Set colAHRmk = Nothing
Set colNPCmd = Nothing

Unload Me

End Sub
Private Sub RefreshData()
  Set mobjAQRmkX = New Remarks
  Set mobjAQRmk = New Remarks
  Set mobjAddedRmk = New Remarks
  GetExistingAQRmk mobjAQRmkX, mobjAQRmk, mobjAddedRmk
  PopulateAQRmk mobjAddedRmk, gobjPNR.CompInfo.WONum
  PopulateRmkType gobjPNR.CompInfo.WONum
  ClearData
End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
  RefreshData
End Sub
Private Sub Form_Load()
   Dim oldParent As Long
   Dim hMenu As Long
   Dim menuItemCount As Long
   gintY = 0
   gintX = 0
   
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
   SwitchWinSetting (Me.hwnd)
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   
   If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
      cmdAdd.Left = cmdPrevious.Left + cmdPrevious.Width + 70
      cmdDelete.Left = cmdAdd.Left + cmdAdd.Width + 70
      cmdRefresh.Left = cmdDelete.Left + cmdDelete.Width + 70
   Else
      cmdPrevious.Visible = False
      cmdAdd.Left = 240
      cmdDelete.Left = cmdAdd.Left + cmdAdd.Width + 70
      cmdRefresh.Left = cmdDelete.Left + cmdDelete.Width + 70
   End If
 
   flgChange = False
   RefreshData
End Sub

Private Sub GetExistingAQRmk(ByRef AQRmkX As Remarks, ByRef AQRmk As Remarks, ByRef AddedRmk As Remarks)
   Dim objAQRmkX As New Remark
   Dim objAQRmk As New Remark
   Dim objAddedRmk As New Remark
   Dim intI As Integer
   Dim strRmkText As String
   Dim strRmkEnd As String
   
   Set AQRmkX = New Remarks
   Set AQRmk = New Remarks
   Set AddedRmk = New Remarks
    
   For intI = 1 To gobjPNR.GeneralRemarkCount
   
       Set objAQRmkX = New Remark
       Set objAQRmk = New Remark
       Set objAddedRmk = New Remark
       
       With gobjPNR.GeneralRemark(intI)
          If .Qualifier = "" Then
             strRmkText = Mid(.RemarkText, 1, 3)
             strRmkEnd = Mid(.RemarkText, Len(.RemarkText), 1)
             If UCase(strRmkText) = "AQ-" And UCase(strRmkEnd) = "X" Then
                objAQRmkX.Number = .ItemNum
                objAQRmkX.RemarkText = Mid(.RemarkText, 4, Len(.RemarkText) - 4)
                AQRmkX.AddRemark objAQRmkX
             ElseIf UCase(strRmkText) = "AQ-" Then
                 objAQRmk.Number = .ItemNum
                 objAQRmk.RemarkText = Mid(.RemarkText, 4)
                 AQRmk.AddRemark objAQRmk
             ElseIf UCase(strRmkText) = "AH-" Then
                 objAddedRmk.Number = .ItemNum
                 objAddedRmk.RemarkText = Mid(.RemarkText, 4)
                 AddedRmk.AddRemark objAddedRmk
             End If
          End If
       End With
   Next
End Sub

Private Sub PopulateAQRmk(ByVal AddedRmk As Remarks, ByVal CMC As String)
  Dim lngI As Long
  Dim intJ As Integer
  Dim strRmkText As String
  Dim strRmk() As String
  Dim colRmk As Collection
  Dim colRmkLike As Collection
  Dim strRemark As String
  Dim strRmkLike As String
  Dim strSql As String
  Dim rsRmk As New ADODB.Recordset
  Dim item As ListItem
  Dim Temp As String
  Dim strTemp As String
  Dim strDesc() As String
  Dim strType() As String
  Dim intK As Integer
  Dim strDescription As String
  Dim strRmkType As String
  
  'Dim strRmkType As String
  Dim strRmkDesc As String
  Dim intCount As Integer
  Dim bolSameRmkType As Boolean
  Dim bolSameRmkDesc As Boolean
  
  Set colRmk = New Collection
  Set colRmkLike = New Collection
  
  lvwRemarkText.ListItems.Clear
  strRmkText = ""
  lngI = 0
  For lngI = 1 To AddedRmk.RemarkCount
     With AddedRmk.Remark(lngI)
          strRmkText = IIf(strRmkText = "", .RemarkText, strRmkText & "*" & .RemarkText)
     End With
  Next
  strRmk() = Split(strRmkText, "*")
  For intJ = 0 To UBound(strRmk)
      If InStr(1, strRmk(intJ), "-") Then
         colRmkLike.Add strRmk(intJ)
      Else
         colRmk.Add strRmk(intJ)
      End If
  Next
  Temp = ""
  If colRmk.Count > 0 Then
     For intJ = 1 To colRmk.Count
         Temp = IIf(Temp = "", "'" & colRmk(intJ) & "'", Temp & ",'" & colRmk(intJ) & "'")
     Next
     strSql = ""
     strSql = "select a.RemarkType,b.RemarkDesc,b.RemarkText from tblAquaItinRmkType a join tblAquaItinRmk b "
     strSql = strSql & "on a.RemarkTypeID = b.RemarkTypeID where a.CMC='"
     strSql = strSql & CMC & "' and b.RemarkText in (" & Temp & ") "
     strSql = strSql & "order by a.RemarkType, b.RemarkDesc"
     Set rsRmk = gdbConn.Execute(strSql)
     While rsRmk.EOF = False
        Set item = lvwRemarkText.ListItems.Add()
        item.SubItems(1) = IIf(IsNull(rsRmk!RemarkType), "", rsRmk!RemarkType)
        item.SubItems(2) = IIf(IsNull(rsRmk!RemarkDesc), "", rsRmk!RemarkDesc)
        item.SubItems(3) = IIf(IsNull(rsRmk!RemarkText), "", rsRmk!RemarkText)
        rsRmk.MoveNext
     Wend
     rsRmk.Close
     Set rsRmk = Nothing
  End If
  
  If colRmkLike.Count > 0 Then
    strRmkText = ""
    For intJ = 1 To colRmkLike.Count
      strRmkText = Mid(colRmkLike(intJ), 1, InStr(1, colRmkLike(intJ), "-"))
      strSql = ""
      strSql = "select a.RemarkType,b.RemarkDesc,b.RemarkText from tblAquaItinRmkType a join tblAquaItinRmk b "
      strSql = strSql & "on a.RemarkTypeID = b.RemarkTypeID where a.CMC='"
      strSql = strSql & CMC & "' and b.RemarkText Like '" & strRmkText & "%' "
      strSql = strSql & "order by a.RemarkType, b.RemarkDesc"
     
      Set rsRmk = gdbConn.Execute(strSql)
      
      'CC - Change the script for easier to understand
        strRmkType = ""
        strRmkDesc = ""
        intCount = 0
        bolSameRmkType = True
        bolSameRmkDesc = True
        Do Until rsRmk.EOF
            intCount = intCount + 1
            If intCount > 1 Then
                If UCase(rsRmk!RemarkType & "") <> UCase(strRmkType) Then
                    bolSameRmkType = False
                End If
                If UCase(rsRmk!RemarkDesc & "") <> UCase(strRmkDesc) Then
                    bolSameRmkDesc = False
                End If
            End If
            strRmkType = rsRmk!RemarkType & ""
            strRmkDesc = rsRmk!RemarkDesc & ""
            rsRmk.MoveNext
        Loop
        rsRmk.Close
        Set rsRmk = Nothing
        
        Set item = lvwRemarkText.ListItems.Add(, , "")
        item.SubItems(1) = IIf(bolSameRmkType, strRmkType, "")
        item.SubItems(2) = IIf(bolSameRmkDesc, strRmkDesc, "")
        item.SubItems(3) = colRmkLike(intJ)
      
'      Temp = ""
'      strTemp = ""
'      While rsRmk.EOF = False
'         Temp = IIf(Temp = "", IIf(IsNull(rsRmk!RemarkDesc), "", rsRmk!RemarkDesc), Temp & _
'                "*" & IIf(IsNull(rsRmk!RemarkDesc), "", rsRmk!RemarkDesc))
'         strTemp = IIf(strTemp = "", IIf(IsNull(rsRmk!RemarkType), "", rsRmk!RemarkType), strTemp & _
'                "*" & IIf(IsNull(rsRmk!RemarkType), "", rsRmk!RemarkType))
'         rsRmk.MoveNext
'      Wend
'      rsRmk.Close
'      Set rsRmk = Nothing
'      strDescription = ""
'      strRmkType = ""
'      If Temp <> "" Then
'         strDesc() = Split(Temp, "*")
'         If UBound(strDesc) > -1 Then
'            strDescription = strDesc(0)
'         End If
'         For intK = 1 To UBound(strDesc)
'             If UCase(strDescription) = UCase(strDesc(intK)) Then
'             Else
'                strDescription = ""
'                Exit For
'             End If
'         Next
'      End If
'
'      If strTemp <> "" Then
'         strType() = Split(strTemp, "*")
'         If UBound(strType) > -1 Then
'            strRmkType = strType(0)
'         End If
'         For intK = 1 To UBound(strType)
'             If UCase(strRmkType) = UCase(strType(intK)) Then
'             Else
'                strRmkType = ""
'                Exit For
'             End If
'         Next
'      End If
'
'     Set item = lvwRemarkText.ListItems.Add()
'     item.SubItems(1) = IIf(strRmkType = "", "", strRmkType)
'     item.SubItems(2) = IIf(strDescription = "", "", strDescription)
'     item.SubItems(3) = colRmkLike(intJ)
          
    Next
  End If
  Set item = Nothing
  Set colRmk = Nothing
  Set colRmkLike = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       gbolCancelProcess = True
    End If
End Sub

Private Sub ClearRmkText()
Dim intI As Integer
 txtHisRmk1 = ""
 txtHisRmk2 = ""
 For intI = 0 To mconMaxRmkIndex
    lblRtext(intI) = ""
    lblRtext(intI).Visible = False
    txtRmkText(intI).Text = ""
    txtRmkText(intI).Visible = False
Next
End Sub

Private Sub lblRtext_Change(Index As Integer)
 flgChange = True
End Sub


Private Sub PopulateRmkType(ByVal CMC As String)
Dim strSql As String
Dim rsRmkType As New ADODB.Recordset
Dim intI As Integer

cmbRmkType.Clear
cmbRmkType.ColumnCount = 2
cmbRmkType.ColumnWidths = "5 cm; 0 cm"
strSql = ""
strSql = "Select RemarkType,RemarkTypeID from tblAquaItinRmkType where CMC='" & CMC & "' order by RemarkType"
Set rsRmkType = gdbConn.Execute(strSql)
intI = 0
While rsRmkType.EOF = False
  cmbRmkType.AddItem
  cmbRmkType.List(intI, 0) = IIf(IsNull(rsRmkType!RemarkType), "", rsRmkType!RemarkType)
  cmbRmkType.List(intI, 1) = IIf(IsNull(rsRmkType!RemarkTypeID), "", rsRmkType!RemarkTypeID)
  rsRmkType.MoveNext
  intI = intI + 1
Wend
rsRmkType.Close
Set rsRmkType = Nothing
cmbRmkType.listindex = -1
End Sub

Private Sub PopulateRmkDesc(RmkTypeID As Long)
Dim strSql As String
Dim rsRmkdesc As New ADODB.Recordset
Dim intI As Integer

cmbDescription.Clear
cmbDescription.ColumnCount = 4
cmbDescription.ColumnWidths = "5 cm; 0 cm; 0cm; 0cm"
strSql = ""
strSql = "Select RemarkDesc,RemarkText,HistoricalRemark1,HistoricalRemark2  from "
strSql = strSql & "tblAquaItinRmk where RemarkTypeID='" & RmkTypeID & "' order by RemarkDesc"
Set rsRmkdesc = gdbConn.Execute(strSql)
intI = 0

While rsRmkdesc.EOF = False
    cmbDescription.AddItem
    cmbDescription.List(intI, 0) = IIf(IsNull(rsRmkdesc!RemarkDesc), "", rsRmkdesc!RemarkDesc)
    cmbDescription.List(intI, 1) = IIf(IsNull(rsRmkdesc!RemarkText), "", rsRmkdesc!RemarkText)
    cmbDescription.List(intI, 2) = IIf(IsNull(rsRmkdesc!HistoricalRemark1), "", rsRmkdesc!HistoricalRemark1)
    cmbDescription.List(intI, 3) = IIf(IsNull(rsRmkdesc!HistoricalRemark2), "", rsRmkdesc!HistoricalRemark2)
    rsRmkdesc.MoveNext
    intI = intI + 1
Wend
  
rsRmkdesc.Close
Set rsRmkdesc = Nothing
'CC - Additional Logic
If cmbDescription.ListCount = 1 Then
    cmbDescription.listindex = 0
Else
    cmbDescription.listindex = -1
End If

End Sub

Private Sub ClearData()
 cmbRmkType.listindex = -1
 cmbDescription.Clear
 ClearRmkText
 flgChange = False
End Sub

Private Sub PopulateRmkText(ByVal RmkText As String)

    Dim intStart As Integer
    Dim intEnd As Integer
    Dim intI As Integer
    Dim bolExceedIndex As Boolean
    
    bolExceedIndex = False
    intI = 0
 
    Do
        intStart = InStr(1, RmkText, "[")
        intEnd = InStr(1, RmkText, "]")
            
        If intStart > 0 And intEnd > 0 And intEnd > intStart Then
            If intStart > 1 Then
                lblRtext(intI).Visible = True
                lblRtext(intI) = Mid(RmkText, 1, intStart - 1)
            End If
           txtRmkText(intI) = Mid(RmkText, intStart, intEnd - intStart + 1)
            RmkText = Mid(RmkText, intEnd + 1)
            
        Else
            lblRtext(intI).Visible = True
            lblRtext(intI) = RmkText
            RmkText = ""
            Exit Do
        End If
        intI = intI + 1
    Loop Until intI >= mconMaxRmkIndex + 1
    
    If RmkText <> "" Then
       bolExceedIndex = True
       txtRmkText(mconMaxRmkIndex).Visible = True
       txtRmkText(mconMaxRmkIndex) = txtRmkText(mconMaxRmkIndex) & RmkText
    End If
    
    For intI = 0 To mconMaxRmkIndex
        If intI = 0 Then
            lblRtext(intI).Left = 240
        Else
            lblRtext(intI).Left = txtRmkText(intI - 1).Left + txtRmkText(intI - 1).Width + 100
        End If
       txtRmkText(intI).Left = lblRtext(intI).Left + lblRtext(intI).Width + 100
        If txtRmkText(intI) = "" Then
           txtRmkText(intI).Visible = False
        Else
           txtRmkText(intI).Visible = True
            If txtRmkText(intI) = "[]" Then
               txtRmkText(intI) = ""
            Else
                If Len(txtRmkText(intI)) - 2 > 0 Then
                    If intI <> mconMaxRmkIndex Then
                       txtRmkText(intI) = Mid(txtRmkText(intI), 2, Len(txtRmkText(intI)) - 2)
                    ElseIf bolExceedIndex = False Then
                       txtRmkText(intI) = Mid(txtRmkText(intI), 2, Len(txtRmkText(intI)) - 2)
                    End If
                End If
            End If
        End If
    Next
End Sub

'CC - Remove ByRef for parameter
Private Function GetLineNumOfAQXToDel(AQRmkX As Remarks) As String
Dim strLineNum As String
Dim lngI As Long
Dim intJ As Integer
Dim intK As Integer
Dim blnExist As Boolean
Dim colNumber As Collection

lngI = 0
intK = 0
intJ = 0
strLineNum = ""
Set colNumber = New Collection

For lngI = 1 To AQRmkX.RemarkCount
    With AQRmkX.Remark(lngI)
         For intJ = 1 To lvwRemarkText.ListItems.Count
             If UCase(lvwRemarkText.ListItems.item(intJ).SubItems(3)) = UCase(.RemarkText) Then
                If colNumber.Count = 0 Then
                      colNumber.Add .Number
                Else
                      blnExist = False
                      For intK = 1 To colNumber.Count
                         If colNumber(intK) = .Number Then
                            blnExist = True
                            Exit For
                         End If
                      Next
                      If blnExist = False Then
                         colNumber.Add .Number
                      End If
                 End If
             End If
        Next
    End With
Next
For intK = 1 To colNumber.Count
      strLineNum = IIf(strLineNum = "", colNumber(intK), strLineNum & "." & colNumber(intK))
Next
Set colNumber = Nothing

GetLineNumOfAQXToDel = strLineNum
End Function

'CC - Remove ByRef for parameter
Private Function GetDeletedRmk(AddedRmk As Remarks) As Collection
Dim lngI As Long
Dim intK As Integer
Dim intJ As Integer
Dim strRmk() As String
Dim flgMatch As Boolean
Dim colDelRmk As Collection

'CC - Added Comment
'NP line: 134. AH-ALT*IAH*OT-EUR900*UPGRADE-999
'AddedRmk.Number = 134
'AddedRmk.RemarkText = ALT*IAH*OT-EUR900*UPGRADE-999

Set colDelRmk = New Collection
For lngI = 1 To AddedRmk.RemarkCount
    With AddedRmk.Remark(lngI)
        strRmk() = Split(.RemarkText, "*")
        For intJ = 0 To UBound(strRmk)
            flgMatch = False
            For intK = 1 To lvwRemarkText.ListItems.Count
                If UCase(lvwRemarkText.ListItems.item(intK).SubItems(3)) = UCase(strRmk(intJ)) Then
                   flgMatch = True
                   Exit For
                Else
                   flgMatch = False
                End If
            Next
            If flgMatch = False Then
               colDelRmk.Add strRmk(intJ)
            End If
        Next
    End With
Next
Set GetDeletedRmk = colDelRmk
End Function

'CC - Remove ByRef for both Parameter
Private Function GetLineNumOfAQToDel(AQRmk As Remarks, DeletedRmk As Collection) As String
Dim strAQToDel As String
Dim lngI As Long
Dim intJ As Integer
Dim intK As Integer
Dim blnExist As Boolean
Dim colNumber As Collection

lngI = 0
intJ = 0
intK = 0
strAQToDel = ""
Set colNumber = New Collection

   For lngI = 1 To AQRmk.RemarkCount
       With AQRmk.Remark(lngI)
            For intJ = 1 To DeletedRmk.Count
                If UCase(DeletedRmk(intJ)) = UCase(.RemarkText) Then
                   If colNumber.Count = 0 Then
                      colNumber.Add .Number
                   Else
                      blnExist = False
                      For intK = 1 To colNumber.Count
                         If colNumber(intK) = .Number Then
                            blnExist = True
                            Exit For
                         End If
                      Next
                      If blnExist = False Then
                         colNumber.Add .Number
                      End If
                   End If
                End If
            Next
       End With
   Next
   For intK = 1 To colNumber.Count
      strAQToDel = IIf(strAQToDel = "", colNumber(intK), strAQToDel & "." & colNumber(intK))
   Next
Set colNumber = Nothing
GetLineNumOfAQToDel = strAQToDel
End Function

'CC - Remove ByRef for both parameter
Private Function IsNewRmk(AddedRmk As Remarks, RmkText As String) As Boolean
    'CC - Change the script for easier to understand
    
    'NP line: 134. AH-ALT*IAH*OT-EUR900*UPGRADE-999
    'AddedRmk.Number = 134
    'AddedRmk.RemarkText = ALT*IAH*OT-EUR900*UPGRADE-999
    
    'RmkText = UPGRADE-999
    
    Dim lngI As Long
    Dim lngJ As Long
    Dim bolFound As Boolean
    Dim strAddedRmk As String
    Dim strAryRmk() As String
    
    strAddedRmk = ""
    For lngI = 1 To AddedRmk.RemarkCount
        strAddedRmk = strAddedRmk & IIf(strAddedRmk = "", "", "*") & AddedRmk.Remark(lngI).RemarkText
    Next
    
    strAryRmk = Split(strAddedRmk, "*")
    
    bolFound = False
    For lngI = 0 To UBound(strAryRmk)
        If UCase(strAryRmk(lngI)) = UCase(RmkText) Then
            bolFound = True
            Exit For
        End If
    Next
    If bolFound = False Then
        IsNewRmk = True
    Else
        IsNewRmk = False
    End If
    

'Dim intI As Integer
'Dim strTemp As String
'Dim strStart As String
'Dim strEnd As String
'
'
'lngI = 0
'IsNewRmk = False
'strStart = ""
'strEnd = ""
'If AddedRmk.RemarkCount > 0 Then
'   For lngI = 1 To AddedRmk.RemarkCount
'       intI = 0
'       strTemp = ""
'       With AddedRmk.Remark(lngI)
'            intI = InStr(1, UCase(.RemarkText), UCase(RmkText))
'            If intI > 0 Then
'               If Len(.RemarkText) > Len(RmkText) Then
'                  If intI = 1 Then
'                     strTemp = Mid(.RemarkText, intI, Len(RmkText) + 2)
'                     strStart = ""
'                  Else
'                    strTemp = Mid(.RemarkText, (intI - 1), Len(RmkText) + 2)
'                    strStart = Mid(strTemp, 1, 1)
'                  End If
'                  If intI + Len(RmkText) = Len(.RemarkText) Then
'                     strEnd = ""
'                  Else
'                    strEnd = Mid(strTemp, Len(strTemp), 1)
'                  End If
'
'                  If (strStart = "*" Or strStart = "") And (strEnd = "*" Or strEnd = "") Then
'                     IsNewRmk = False
'                     Exit For
'                  End If
'               Else
'                  If UCase(.RemarkText) = UCase(RmkText) Then
'                     IsNewRmk = False
'                     Exit For
'                  End If
'               End If
'            Else
'               IsNewRmk = True
'            End If
'       End With
'   Next
'Else
'   IsNewRmk = True
'End If
End Function

'CC - Remove ByRef for AddedRmk
Private Function GetRmkFromListView(AddedRmk As Remarks, ByRef AHRmk As Collection, ByRef HistoricalRmk As Collection) As Collection
Dim intI As Integer
Dim intJ As Integer
Dim strRmk() As String
Dim strTemp As String
Dim strRmkText As String
Dim blnNewRmk As Boolean

strRmkText = ""
strTemp = ""
intI = 0
intJ = 0

Set AHRmk = New Collection
Set HistoricalRmk = New Collection
Set GetRmkFromListView = New Collection

   For intI = 1 To lvwRemarkText.ListItems.Count
       If Len(lvwRemarkText.ListItems.item(intI).SubItems(4)) <> 0 Then
          HistoricalRmk.Add lvwRemarkText.ListItems.item(intI).SubItems(4)
       End If
       If Len(lvwRemarkText.ListItems.item(intI).SubItems(5)) <> 0 Then
          HistoricalRmk.Add lvwRemarkText.ListItems.item(intI).SubItems(5)
       End If
       If Len(lvwRemarkText.ListItems.item(intI).SubItems(3)) <> 0 Then
          If strRmkText = "" Then
             strRmkText = lvwRemarkText.ListItems.item(intI).SubItems(3)
          Else
             strRmkText = strRmkText & "*" & lvwRemarkText.ListItems.item(intI).SubItems(3)
          End If
       End If
       'CC - do not add remark text to GetRmkFromListView if remark text is blank
       If lvwRemarkText.ListItems.item(intI).SubItems(3) <> "" Then
            blnNewRmk = False
            blnNewRmk = IsNewRmk(AddedRmk, lvwRemarkText.ListItems.item(intI).SubItems(3))
            If blnNewRmk = True Then
               GetRmkFromListView.Add lvwRemarkText.ListItems.item(intI).SubItems(3)
            End If
       End If
   Next
If strRmkText <> "" Then
   If Len(strRmkText) > mconMaxRmkLen Then
      strRmk() = Split(strRmkText, "*")
      For intJ = 0 To UBound(strRmk)
          If intJ = 0 Then
             strTemp = strRmk(intJ)
          Else
             If Len(strTemp) + Len(strRmk(intJ)) + 1 < mconMaxRmkLen Then
                strTemp = IIf(strTemp = "", strRmk(intJ), strTemp & "*" & strRmk(intJ))
                If intJ = UBound(strRmk) Then
                   AHRmk.Add strTemp
                End If
             Else
                AHRmk.Add strTemp
                If intJ = UBound(strRmk) Then
                   AHRmk.Add strRmk(intJ)
                Else
                  strTemp = strRmk(intJ)
                End If
             End If
          End If
      Next
   Else
      AHRmk.Add strRmkText
   End If
End If
End Function
Private Sub loadPNR()
Set gobjPNR = New CWT_GalileoPNR3.PNR
gobjPNR.loadPNR
End Sub

Private Sub txtHisRmk_Change(Index As Integer)
flgChange = True
End Sub

Private Sub txtRmkText_Change(Index As Integer)
flgChange = True
End Sub
Private Function validData(ByRef RmkText As String) As Boolean
Dim strError As String
Dim intTotlen As Integer
Dim intI As Integer
Dim blnVisible As Boolean

Dim strRmk As String

strError = ""
intTotlen = 0
blnVisible = False

'CC - Remove flgChange
'If flgChange = True Then
    txtHisRmk1.Text = RTrim(txtHisRmk1.Text)
    txtHisRmk2.Text = RTrim(txtHisRmk2.Text)
    
    strRmk = ""
    'Check if any remark text textbox is blank
    'and combine remark text
    For intI = 0 To mconMaxRmkIndex
        If lblRtext(intI).Visible = True Then
            strRmk = strRmk & lblRtext(intI)
        End If
        If txtRmkText(intI).Visible = True Then
            If txtRmkText(intI).Text = "" Then
                strError = "Please indicate value under Remark Text"
                Exit For
            Else
                strRmk = strRmk & txtRmkText(intI).Text
            End If
        End If
    Next
    
    'Check if combined remark text exceed the max length
    If strError = "" Then
        If Len(strRmk) > mconMaxRmkLen Then
            strError = "Remark Text exceeds " & mconMaxRmkLen & " characters. Please shorten Remark Text"
        End If
    End If
    
    'Check if user key in historical remark if remark text is blank
    If strError = "" Then
        If strRmk = "" And txtHisRmk1 = "" And txtHisRmk2 = "" Then
            If cmbDescription.listindex = -1 Then
                strError = "Please select Remark Description or indicate value under Historical Remark"
            Else
                strError = "Please indicate value under Historical Remark"
            End If
        End If
    End If
    
    'Check if "-" or "*" exist in remark text
    If strError = "" Then
        For intI = 0 To mconMaxRmkIndex
            If txtRmkText(intI).Visible = True Then
                If InStr(1, txtRmkText(intI).Text, "-") > 0 Then
                    strError = strError & "Invalid character '-' found in Remark Text." & vbCrLf
                    Exit For
                End If
            End If
        Next
        For intI = 0 To mconMaxRmkIndex
            If txtRmkText(intI).Visible = True Then
                If InStr(1, txtRmkText(intI).Text, "*") > 0 Then
                    strError = strError & "Invalid character '*' found in Remark Text." & vbCrLf
                    Exit For
                End If
            End If
        Next
        For intI = 0 To mconMaxRmkIndex
            If lblRtext(intI).Visible = True Then
                If InStr(1, lblRtext(intI).Caption, "*") > 0 Then
                    strError = strError & "Invalid character '*' found in Remark Text." & vbCrLf
                    strError = strError & "Please Contact Administrator." & vbCrLf
                    Exit For
                End If
            End If
        Next
    End If
    
    If strError = "" And strRmk <> "" Then
        For intI = 1 To lvwRemarkText.ListItems.Count
            If UCase(strRmk) = UCase(lvwRemarkText.ListItems(intI).SubItems(3)) Then
                strError = "Remark Text exist in grid view box."
                Exit For
            End If
        Next
    End If
    
'CC - Change
'    For intI = 0 To mconMaxRmkIndex
'        If txtRmkText(intI).Visible = True Then
'           blnVisible = True
'           intTotlen = intTotlen + Len(txtRmkText(intI).Text)
'        End If
'    Next
'
'    If blnVisible = True And intTotlen = 0 And Len(txtHisRmk1.Text) = 0 And Len(txtHisRmk2.Text) = 0 Then
'       strError = "Remark Text Cannot Be Empty"
'    Else
'       For intI = 0 To mconMaxRmkIndex
'           If lblRtext(intI).Visible = True Then
'            intTotlen = intTotlen + Len(lblRtext(intI).Caption)
'           End If
'       Next
'       If intTotlen > mconMaxRmkLen Then
'          strError = strError & "Remark Text exceeds " & mconMaxRmkLen & ". Please shorten Remark Text"
'       End If
'    End If
''Else
''    strError = strError & "Please Enter Data"
''End If
If strError <> "" Then
    validData = False
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strError, vbOKOnly + vbDefaultButton1, "AQUA Itin Remark - Data Required"
Else
   validData = True
   RmkText = strRmk
End If
End Function
