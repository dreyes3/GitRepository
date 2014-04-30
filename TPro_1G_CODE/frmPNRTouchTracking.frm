VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmPNRTouchTracking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CWT Desktop - PNR Tracking Touches"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   2520
   ClientWidth     =   10845
   Icon            =   "frmPNRTouchTracking.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   10845
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   3945
      Left            =   0
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6959
      Color           =   16447215
      FinColor        =   16048334
      Caption         =   "ARGradient1"
      ForeColor       =   -2147483630
      GradientSteps   =   65
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MyCommandButton.MyButton cmdFinish 
         Height          =   360
         Left            =   8640
         TabIndex        =   0
         Top             =   3120
         Width           =   1005
         _ExtentX        =   1773
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
         TransparentColor=   15591915
         Caption         =   "&Finish"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   360
         Left            =   9840
         TabIndex        =   1
         Top             =   3120
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   635
         BackColor       =   12648447
         Enabled         =   0   'False
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
         TransparentColor=   15591915
         Caption         =   "&Cancel"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyFramePanel.MyFrame topFrame 
         Height          =   3135
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   5530
         BackColor       =   14342838
         ForeColor       =   15979465
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   13557
         BackgroundMaskColor=   13421772
         BackgroundAlignment=   4
         Caption         =   ""
         CornerTopLeft   =   -1  'True
         CornerTopRight  =   -1  'True
         CornerBottomLeft=   -1  'True
         CornerBottomRight=   -1  'True
         OutSideColor    =   15591915
         HeaderGradientAlign=   5
         HeaderGradientSizeH=   "50%"
         HeaderColorTopLeft=   6973442
         HeaderColorTopRight=   6973442
         HeaderColorBottomLeft=   6973442
         HeaderColorBottomRight=   6973442
         HeaderShow      =   0   'False
         PictureOffsetX  =   5
         Begin MyFramePanel.MyFrame fraTrackingTouches 
            Height          =   3015
            Left            =   120
            Top             =   120
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   5318
            BackColor       =   14342838
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            AppearanceThemes=   3
            BackgroundAlignment=   4
            Caption         =   "Tracking Touches"
            CaptionAlignment=   0
            CornerRadius    =   10
            CornerTopLeft   =   -1  'True
            CornerTopRight  =   -1  'True
            OutSideColor    =   15591915
            HeaderHeight    =   20
            HeaderColorTopLeft=   0
            HeaderColorTopRight=   0
            HeaderColorBottomLeft=   0
            HeaderColorBottomRight=   0
            Begin MyCommandButton.MyButton cmdAdd 
               Height          =   360
               Left            =   8640
               TabIndex        =   5
               Top             =   1530
               Width           =   1005
               _ExtentX        =   1773
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
               TransparentColor=   15591915
               Caption         =   "&Add"
               Depth           =   1
               GradientType    =   2
            End
            Begin MSComctlLib.ListView lvRc 
               Height          =   1050
               Left            =   120
               TabIndex        =   9
               Top             =   1900
               Width           =   9615
               _ExtentX        =   16960
               _ExtentY        =   1852
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16777215
               Appearance      =   1
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Contact Code"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Contact Method"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Primary Code"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Primary Reason"
                  Object.Width           =   5292
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Secondary Code"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Secondary Reason"
                  Object.Width           =   7056
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Charge Indicator"
                  Object.Width           =   0
               EndProperty
            End
            Begin MSForms.ComboBox cmbSecRC 
               Height          =   315
               Left            =   2520
               TabIndex        =   8
               Top             =   1170
               Width           =   7095
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "12515;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbPrimaryRC 
               Height          =   315
               Left            =   2520
               TabIndex        =   7
               Top             =   710
               Width           =   7095
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "12515;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox cmbContactMethod 
               Height          =   315
               Left            =   2520
               TabIndex        =   6
               Top             =   240
               Width           =   7095
               VariousPropertyBits=   746604571
               DisplayStyle    =   3
               Size            =   "12515;556"
               ColumnCount     =   2
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label lblSecRC 
               BackColor       =   &H00DADAB6&
               BackStyle       =   0  'Transparent
               Caption         =   "Secondary Reason Code"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   4
               Top             =   1190
               Width           =   2295
            End
            Begin VB.Label lblPrimaryRC 
               BackColor       =   &H00DADAB6&
               BackStyle       =   0  'Transparent
               Caption         =   "Primary Reason Code"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   3
               Top             =   780
               Width           =   2295
            End
            Begin VB.Label lblContactMethod 
               BackColor       =   &H00DADAB6&
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Method"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   2
               Top             =   360
               Width           =   2055
            End
         End
         Begin vbskpro.Skinner Skinner1 
            Left            =   0
            Top             =   480
            _ExtentX        =   1270
            _ExtentY        =   1270
            CloseButton     =   1
            MaxButton       =   0
            MinButton       =   0
            OldForeColor    =   0
            ChangeSkinButton=   0   'False
            SysDisableSkinCaption=   "&Disable Skin"
            LcK1            =   "..02*-0..*/305*.-2-/"
            LcK2            =   $"frmPNRTouchTracking.frx":038A
            AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
         End
      End
   End
End
Attribute VB_Name = "frmPNRTouchTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Preethi - V1.2.4 20110621  - ER01 - Tracking of Touches to CWT Booking
Dim modPRCtype As String
Dim modI As Integer


Private Sub cmbPrimaryRC_Change()
Dim strClient As String

If modPRCtype = "C" Then
 strClient = gobjPNR.CompInfo.WONum
Else
  If modPRCtype = "S" Then
    strClient = "00000"
  End If
End If

If cmbPrimaryRC.value <> "Null" Then
   GetTouchesSecondaryCode cmbSecRC, strClient, cmbPrimaryRC.value
End If
If cmbSecRC.ListCount > 1 Then
   cmbSecRC.listindex = -1
Else
   If cmbSecRC.ListCount = 1 Then
      cmbSecRC.listindex = 0
   End If
End If
End Sub



Private Sub cmdAdd_Click()
Dim strError As String
Dim intZ As Integer
Dim intI As Integer
Dim strComboText As String
Dim strErr As String
Dim secCmbtext As String

strError = ""
If cmbContactMethod.value = "" Or cmbContactMethod.Text = "" Then
   strError = "Contact method is missing"
End If
If cmbPrimaryRC.value = "" Or cmbPrimaryRC.Text = "" Then
   strError = strError & vbCrLf & "Primary Reason code is missing"
End If
If cmbSecRC.value = "" Or cmbSecRC.Text = "" Then
   strError = strError & vbCrLf & "Secondary Reason code is missing"
End If
If strError <> "" Then
   strError = strError & vbCrLf & "Please select and Click Add button"
End If
If lvRc.ListItems.Count > 0 Then
  For intZ = 1 To lvRc.ListItems.Count
   If lvRc.ListItems(intZ).SubItems(5) = cmbSecRC.Text Then
      strError = strError & vbCrLf & "This secondary code has been added." & vbCrLf & "Please re-select and click Add button"
      Exit For
   End If
  Next
End If

If strError <> "" Then
   MsgBox strError, vbOKOnly, "Tracking Touches - Error"
   Exit Sub
End If

strComboText = ""
strComboText = cmbContactMethod.Text
For intI = 1 To cmbContactMethod.ListCount
    cmbContactMethod.listindex = intI - 1
    If cmbContactMethod.Text = strComboText Then
       strErr = ""
       Exit For
    Else
       strErr = "Contact Method entered doesn't match with available list"
    End If
Next
If strErr <> "" Then
   cmbContactMethod.listindex = -1
   MsgBox strErr, vbOKOnly, "Tracking Touches - Error"
   Exit Sub
End If

strComboText = ""
secCmbtext = ""
secCmbtext = cmbSecRC.Text
strComboText = cmbPrimaryRC.Text
For intI = 1 To cmbPrimaryRC.ListCount
    cmbPrimaryRC.listindex = intI - 1
    If cmbPrimaryRC.Text = strComboText Then
       strErr = ""
       Exit For
    Else
       strErr = "Primary Reason Code entered doesn't match with available list"
    End If
Next
If strErr <> "" Then
   cmbPrimaryRC.listindex = -1
   MsgBox strErr, vbOKOnly, "Tracking Touches - Error"
   Exit Sub
End If

For intI = 1 To cmbSecRC.ListCount
    cmbSecRC.listindex = intI - 1
    If cmbSecRC.Text = secCmbtext Then
       strErr = ""
       Exit For
    Else
       strErr = "Secondary Reason Code entered doesn't match with available list"
    End If
Next
If strErr <> "" Then
   cmbSecRC.listindex = -1
   MsgBox strErr, vbOKOnly, "Tracking Touches - Error"
   Exit Sub
End If

With lvRc
 .ListItems.Add , , cmbContactMethod.value
 .ListItems(modI).SubItems(1) = cmbContactMethod.Text
 .ListItems(modI).SubItems(2) = cmbPrimaryRC.value
 .ListItems(modI).SubItems(3) = cmbPrimaryRC.Text
 .ListItems(modI).SubItems(4) = cmbSecRC.value
 .ListItems(modI).SubItems(5) = cmbSecRC.Text
 .ListItems(modI).SubItems(6) = cmbSecRC.List(cmbSecRC.listindex, 2)
 modI = modI + 1
End With
  
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFinish_Click()

  If lvRc.ListItems.Count = 0 Then
     MsgBox "Please click FINISH button after making at least one selection of" & vbCrLf & "Contact Method/Primary Reason Code and Secondary Reason Code", vbOKOnly, "Tracking Touches - Error"
  Else
     writeNPline
  End If
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
    
    Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
    Skinner1.CloseButton = skNo
    pDisplayToFP "*SD"
    PopulateControls
    modI = 1
   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
End Sub
Private Sub PopulateControls()
   modPRCtype = ""
   GetTouchesContactMethod cmbContactMethod
   If cmbContactMethod.ListCount > 1 Then
      cmbContactMethod.listindex = -1
   Else
     If cmbContactMethod.ListCount = 1 Then
        cmbContactMethod.listindex = 0
     End If
   End If
   GetTouchesPrimaryCode cmbPrimaryRC, gobjPNR.CompInfo.WONum, modPRCtype
   If cmbPrimaryRC.ListCount > 1 Then
      cmbPrimaryRC.listindex = -1
   Else
     If cmbPrimaryRC.ListCount = 1 Then
        cmbPrimaryRC.listindex = 0
     End If
   End If
       
End Sub

Private Sub lvRc_DblClick()

Dim intJ As Integer
Dim intresponse As Integer

intresponse = 0
intJ = 0

If lvRc.ListItems.Count > 0 Then
   intJ = lvRc.SelectedItem.Index
   intresponse = MsgBox("Do you want to remove the selected Reason Code ?", vbYesNo, "Tracking Touches")
   If intresponse = 6 Then
      lvRc.ListItems.Remove (intJ)
      modI = modI - 1
   End If
End If
End Sub
Private Sub writeNPline()
Dim strCmd As String
Dim intZ As Integer
Dim strDate As String
Dim strTime As String
Dim intCount As Integer
Dim strResponse As String
Dim i As Integer
Dim bolFound As Boolean
Dim intNPNum As Integer
Dim strRemark As String
Dim strTemp As String
Dim strMsg As String
Dim strCount As String

strCmd = ""
strDate = Format(DateValue(Now), "DDMMM")
strTime = Format(TimeValue(Now), "hhmm")
If lvRc.ListItems.Count > 0 Then
  For intZ = 1 To lvRc.ListItems.Count
      If lvRc.ListItems(intZ).SubItems(6) = "BY" Or lvRc.ListItems(intZ).SubItems(6) = "BN" Then
         strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP." & lvRc.ListItems(intZ).SubItems(6) & "*" & lvRc.ListItems(intZ).SubItems(2) & lvRc.ListItems(intZ).SubItems(4) & "/D-"
         strCmd = strCmd & UCase(strDate) & "/T-" & strTime & "/P-" & IIf(IsNumeric(gobjPNR.Agent), "XXXX", gobjHost.AgentPCC) & "/CM-" & lvRc.ListItems(intZ).Text
         strCmd = strCmd & "/A-" & gobjPNR.Agent & "/"
      Else
         If lvRc.ListItems(intZ).SubItems(6) = "HI" Then
            bolFound = False
            For i = 1 To gobjPNR.GeneralRemarkCount
                With gobjPNR.GeneralRemark(i)
                  If .Qualifier = "HI" Then
                    strRemark = .RemarkText
                    If InStr(1, strRemark, "C-") Then
                       bolFound = True
                       intNPNum = .ItemNum
                       Exit For
                    End If
                  End If
               End With
            Next
            
            If bolFound = True Then
               intCount = Mid(strRemark, InStr(1, strRemark, "C-") + 2, 3)
                If IsNumeric(intCount) Then
                    If intCount = 999 Then
                       intCount = 999
                    Else
                       intCount = intCount + 1
                    End If
                    strCount = Format(intCount, "000")
                    strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP." & intNPNum & "@"
                    strCmd = strCmd & lvRc.ListItems(intZ).SubItems(6) & "*" & lvRc.ListItems(intZ).SubItems(2) & lvRc.ListItems(intZ).SubItems(4) & "/D-"
                    strCmd = strCmd & UCase(strDate) & "/T-" & strTime & "/P-" & IIf(IsNumeric(gobjPNR.Agent), "XXXX", gobjHost.AgentPCC) & "/CM-" & lvRc.ListItems(intZ).Text
                    strCmd = strCmd & "/A-" & gobjPNR.Agent & "/C-" & strCount & "/"
                 Else
                    GoTo 11
                 End If
            Else
11:
              strCount = "001"
              strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP." & lvRc.ListItems(intZ).SubItems(6) & "*" & lvRc.ListItems(intZ).SubItems(2) & lvRc.ListItems(intZ).SubItems(4) & "/D-"
              strCmd = strCmd & UCase(strDate) & "/T-" & strTime & "/P-" & IIf(IsNumeric(gobjPNR.Agent), "XXXX", gobjHost.AgentPCC) & "/CM-" & lvRc.ListItems(intZ).Text
              strCmd = strCmd & "/A-" & gobjPNR.Agent & "/C-" & strCount & "/"
            End If
         End If
      End If
  Next

End If
If strCmd <> "" Then
  strCmd = strCmd & IIf(strCmd = "", "", "+") & "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
  strCmd = strCmd & "+ER+ER+ER"
  strResponse = gobjHost.terminalEntry(strCmd)
  strTemp = strResponse
    'Preethi - V1.2.4 20110712 - CR 76 - Change Validation Logic For ENDPNR
    'If InStr(strTemp, "1.1") = 0 Then
    If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = False Then
        For i = 0 To 1
           strTemp = gobjHost.terminalEntry("ER")
           strResponse = strTemp
           'If InStr(strTemp, "1.1") > 0 Then
           If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = True Then
               Exit For
           End If
           If i = 1 Then GoTo errorWriting
        Next
    End If
  
    Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.loadPNR
  
End If
Me.cmdCancel.Enabled = True
Unload Me
              
              
Exit Sub

errorWriting:
    'Prompt error message if failed to write to PNR
    gbolWritingtoPNR = False
    Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.loadPNR
    strMsg = "Desktop is unable to end this transaction.. Response from GDS is " & Chr(13) & strResponse
    strMsg = strMsg & Chr(13) & "Please perform IR using Desktop and click on "
    strMsg = strMsg & Chr(13) & "FINISH button again in the TOUCH TRACKING module."
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
    'cmdCancel.Enabled = True
    
  
'  strTemp = gobjHost.EndPNR2(IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine), True, 2)
'  Set gobjPNR = New CWT_GalileoPNR3.PNR
'  gobjPNR.loadPNR
'  pDisplayToFP ("*R")
'  If strTemp <> "True" Then
'     gbolWritingtoPNR = False
'     strMsg = "Unable to write to PNR. Response from GDS is " & Chr(13) & strResponse
'     strMsg = strMsg & Chr(13) & "System will continue without ending this booking."
'     modMsgBox.OKMsg = "OK"
'     modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
'     'Exit Sub
'  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.cmdCancel.Enabled = False Then
       Cancel = True
    End If
End Sub

