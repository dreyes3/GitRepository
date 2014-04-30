VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{4832871B-0993-461C-B983-0EAAA4A43E5C}#5.0#0"; "SftTabs_IX86_U_50.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient.ocx"
Begin VB.Form frmEitin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CWT Desktop - E-Itinerary"
   ClientHeight    =   3420
   ClientLeft      =   2385
   ClientTop       =   3540
   ClientWidth     =   11625
   Icon            =   "frmEitin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   11625
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   4980
      Left            =   0
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   8784
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
      Begin SftTabsLib.SftTabs SftTabs 
         Height          =   2895
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   11415
         PropVer         =   50
         xcx             =   20135
         xcy             =   5106
         PropFile        =   ""
         PropDesignTime  =   1
         DeletePropFile  =   0
         IntVal          =   55
         xBfStyle1       =   63839708
         xBfStyle2       =   -68685137
         xBfStyle3       =   -115751725
         xBfStyle4       =   384187181
         TabCount        =   2
         CurrentTab      =   0
         FlatProperties  =   0   'False
         BeginProperty Tab(0) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Send Info"
            ToolTip         =   ""
            Object.Align           =   0
            BackColor       =   -1
            BackColorActive =   -1
            ClientAreaColor =   -1
            Enabled         =   1
            FlyByColor      =   -1
            ForeColor       =   -1
            ForeColorActive =   -1
            Name            =   ""
            Hidden          =   0
            BackColorStart  =   -1
            BackColorEnd    =   -1
            BackColorActiveStart=   -1
            BackColorActiveEnd=   -1
            ClientAreaColorStart=   -1
            ClientAreaColorEnd=   -1
         EndProperty
         BeginProperty Tab(1) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Remarks"
            ToolTip         =   ""
            Object.Align           =   0
            BackColor       =   -1
            BackColorActive =   -1
            ClientAreaColor =   -1
            Enabled         =   1
            FlyByColor      =   -1
            ForeColor       =   -1
            ForeColorActive =   -1
            Name            =   ""
            Hidden          =   0
            BackColorStart  =   -1
            BackColorEnd    =   -1
            BackColorActiveStart=   -1
            BackColorActiveEnd=   -1
            ClientAreaColorStart=   -1
            ClientAreaColorEnd=   -1
         EndProperty
         BeginProperty Tabs {48328721-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
         EndProperty
         BeginProperty Scrolling {48328724-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            ToolTipCloseButton=   "Close"
            ToolTipLeftButton=   "Scroll Left/Up"
            ToolTipRightButton=   "Scroll Right/Down"
            ToolTipRestoreButton=   "Restore"
            ToolTipMinimizeButton=   "Minimize"
         EndProperty
         Appearance      =   1
         ClientArea      =   1
         Enabled         =   1
         FlatProperties  =   0
         MousePointer    =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         Style           =   21
         UseClientAreaColors=   0
         BackColor       =   -2147483633
         BorderColor     =   -2147483627
         BrightHighlightColor=   -2147483628
         ForeColor       =   -2147483630
         HighlightColor  =   -2147483626
         ShadowColor     =   -2147483632
         ScrollingButtonBackColor=   -2147483633
         ScrollingButtonBorderColor=   -2147483627
         ScrollingButtonForeColor=   -2147483630
         ScrollingButtonForeColorGrayed=   -2147483630
         ScrollingButtonHighlightColor=   -2147483628
         ScrollingButtonShadowColor=   -2147483632
         TabsBoldActive  =   1
         TabsDropText    =   0
         TabsEachRow     =   0
         TabsFillComplete=   0
         TabsFixed       =   0
         TabsFlyby       =   1
         TabsRows        =   1
         TabsRowIndent   =   -1
         TabsTextOnly    =   0
         TabsLeftMargin  =   0
         TabsRightMargin =   0
         ScrollButtonStyle=   5
         ScrollCondScrollButtons=   0
         ScrollFullSize  =   1
         ScrollHideScrollButtons=   0
         ScrollNoTruncate=   0
         ScrollScrollable=   0
         ScrollScrollOnLeft=   0
         AlwaysShowAccel =   0
         UseExactRegion  =   1
         UseThemes       =   -1  'True
         ShowFocusRectangle=   1
         ButtonAlignment =   0
         CloseButton     =   0
         CloseButtonDisabled=   0
         CloseButtonWMCLOSE=   1
         CloseButtonFullSize=   0
         CloseButtonAlignment=   0
         MinimizeButton  =   0
         RestoreButton   =   0
         CustomCode      =   0
         xDesign         =   30
         yDesign         =   345
         List(0)Count    =   0
         List(1)Count    =   1
         List(1)(0)Ctl   =   "EitinRemarkfra"
         List(1)(0)Ena   =   -1
         List(1)(0)x     =   60
         List(1)(0)y     =   420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MyFramePanel.MyFrame topFrame 
            Height          =   2385
            Left            =   15
            Top             =   360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4207
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
            OutSideColor    =   14215660
            HeaderGradientAlign=   5
            HeaderGradientSizeH=   "50%"
            HeaderColorTopLeft=   6973442
            HeaderColorTopRight=   6973442
            HeaderColorBottomLeft=   6973442
            HeaderColorBottomRight=   6973442
            HeaderShow      =   0   'False
            PictureOffsetX  =   5
            Begin VB.Frame fraFaxes 
               BackColor       =   &H00DADAB6&
               Caption         =   "Sending Info"
               Height          =   2055
               Left            =   6120
               TabIndex        =   6
               Top             =   80
               Width           =   4935
               Begin VB.CheckBox chkEFax 
                  BackColor       =   &H00DADAB6&
                  Caption         =   "Send E-Fax"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   7
                  Top             =   240
                  Width           =   1575
               End
               Begin MSFlexGridLib.MSFlexGrid msFlexFax 
                  Height          =   1035
                  Left            =   60
                  TabIndex        =   8
                  Top             =   480
                  Width           =   4605
                  _ExtentX        =   8123
                  _ExtentY        =   1826
                  _Version        =   393216
                  FixedCols       =   0
                  RowHeightMin    =   250
                  BackColor       =   16777215
                  BackColorFixed  =   6973442
                  ForeColorFixed  =   16777215
                  BackColorSel    =   -2147483643
                  BackColorBkg    =   14342838
                  HighLight       =   0
                  AllowUserResizing=   3
                  BorderStyle     =   0
                  Appearance      =   0
               End
            End
            Begin VB.Frame fraEmails 
               BackColor       =   &H00DADAB6&
               Caption         =   "Sending Info"
               Height          =   2055
               Left            =   100
               TabIndex        =   1
               Top             =   80
               Width           =   6015
               Begin VB.CheckBox chkEmail 
                  BackColor       =   &H00DADAB6&
                  Caption         =   "Send Email"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   3
                  Top             =   240
                  Value           =   1  'Checked
                  Width           =   1575
               End
               Begin VB.TextBox txtReplyEmail 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   960
                  TabIndex        =   2
                  Tag             =   "OFFICE"
                  Top             =   1680
                  Width           =   4335
               End
               Begin MSFlexGridLib.MSFlexGrid msFlexEmails 
                  Height          =   1035
                  Left            =   60
                  TabIndex        =   4
                  Top             =   480
                  Width           =   5565
                  _ExtentX        =   9816
                  _ExtentY        =   1826
                  _Version        =   393216
                  Cols            =   6
                  FixedCols       =   0
                  RowHeightMin    =   250
                  BackColor       =   16777215
                  BackColorFixed  =   6973442
                  ForeColorFixed  =   16777215
                  BackColorSel    =   -2147483643
                  BackColorBkg    =   14342838
                  HighLight       =   0
                  AllowUserResizing=   3
                  BorderStyle     =   0
                  Appearance      =   0
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Reply To:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   5
                  Top             =   1740
                  Width           =   690
               End
            End
         End
         Begin MyFramePanel.MyFrame EitinRemarkfra 
            Height          =   2385
            Left            =   -12310
            Top             =   -3805
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4207
            BackColor       =   14342838
            ForeColor       =   15979465
            Enabled         =   0   'False
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
            OutSideColor    =   14215660
            HeaderGradientAlign=   5
            HeaderGradientSizeH=   "50%"
            HeaderColorTopLeft=   6973442
            HeaderColorTopRight=   6973442
            HeaderColorBottomLeft=   6973442
            HeaderColorBottomRight=   6973442
            HeaderShow      =   0   'False
            PictureOffsetX  =   5
            Begin VB.Frame fraEitinRemarks 
               BackColor       =   &H00DADAB6&
               Caption         =   "Additional Remarks"
               Enabled         =   0   'False
               Height          =   2235
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   11115
               Begin VB.TextBox txtAdRI 
                  Height          =   315
                  Index           =   4
                  Left            =   240
                  TabIndex        =   21
                  Top             =   1740
                  Width           =   5235
               End
               Begin VB.TextBox txtAdRI 
                  Height          =   315
                  Index           =   3
                  Left            =   240
                  TabIndex        =   20
                  Top             =   1380
                  Width           =   5235
               End
               Begin VB.TextBox txtAdRI 
                  Height          =   315
                  Index           =   2
                  Left            =   240
                  TabIndex        =   19
                  Top             =   1020
                  Width           =   5235
               End
               Begin VB.TextBox txtAdRI 
                  Height          =   315
                  Index           =   1
                  Left            =   240
                  TabIndex        =   18
                  Top             =   660
                  Width           =   5235
               End
               Begin VB.TextBox txtAdRI 
                  Height          =   315
                  Index           =   0
                  Left            =   240
                  TabIndex        =   17
                  Top             =   300
                  Width           =   5235
               End
            End
         End
      End
      Begin VB.CheckBox chkEntry 
         Height          =   200
         Left            =   3000
         TabIndex        =   12
         Top             =   3060
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.PictureBox cmbContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   1005
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   1000
         Begin MSForms.ComboBox cmbEntry 
            Height          =   375
            Left            =   60
            TabIndex        =   10
            Top             =   0
            Width           =   855
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1508;661"
            ListWidth       =   7055
            ColumnCount     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   34
         End
      End
      Begin vbskpro.Skinner Skinner1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1270
         _ExtentY        =   1270
         CloseButton     =   1
         MaxButton       =   0
         MinButton       =   0
         OldForeColor    =   0
         ChangeSkinButton=   0   'False
         SysDisableSkinCaption=   "&Disable Skin"
         LcK1            =   "..02*-0..*/305*.-2-/"
         LcK2            =   $"frmEitin.frx":038A
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
      Begin MyCommandButton.MyButton cmdFinish 
         Height          =   360
         Left            =   9180
         TabIndex        =   13
         Top             =   3000
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
         TransparentColor=   14215660
         Caption         =   "&Send"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   360
         Left            =   10260
         TabIndex        =   14
         Top             =   3000
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
         TransparentColor=   14215660
         Caption         =   "&Cancel"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdPrevious 
         Height          =   360
         Left            =   7620
         TabIndex        =   15
         Top             =   3000
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
   End
   Begin VB.Menu mnuPopUpFlex 
      Caption         =   "Pop Up Flex"
      Visible         =   0   'False
      Begin VB.Menu subMenuAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu subMenuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "frmEitin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrFlex As String
Dim mstrEmailLines As String
Dim mbolClickBelowRow As Boolean
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date

Private Sub cmbEntry_GotFocus()
    cmbGetFocus cmbEntry
End Sub

Private Sub cmdCancel_Click()
    gbolCancelProcess = True
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    Dim intMsg As Integer
    Dim strMsg As String
    
    'CC - V1.2.24 20140129 - CR 304 - JTB Integration
    If chkEmail.value = vbChecked Then
        strMsg = "Has the E itinerary recipient email address been updated in RI.ITI Field?"
        modMsgBox.YESMsg = "Yes"
        modMsgBox.NOMsg = "No"
        intMsg = modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbYesNo + vbDefaultButton1, "CWT Desktop - E-Itinerary")
        
        If intMsg = vbNo Then
            Exit Sub
        End If
    End If

    'ZhiSam - V1.2.1 20130724 - CR227 - Disable E-DOC & E-ITIN function to sent out twice
    cmdFinish.Enabled = False
    datTouchEnd = Now
    
    If validData Then
       writeDatatoGDS
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModItin, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModItin, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModItin, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
       
       
       
       If gbolWritingtoPNR = False Then Exit Sub
       Unload Me
    'ZhiSam - V1.2.1 20130724 - CR227 - Disable E-DOC & E-ITIN function to sent out twice
    Else
        cmdFinish.Enabled = True
    End If
End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload Me
End Sub

Private Sub Form_Load()

   Dim oldParent As Long
   Dim i As Integer
   datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   pDisplayToFP "*R"
   'Set the header caption and width of msFlexEmails
    With msFlexEmails
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
              .Text = ""
              .ColWidth(i) = 280
            ElseIf i = 1 Then
              .Text = " E-Itin "
              .ColWidth(i) = 600
            ElseIf i = 2 Then
              .Text = " Type"
              .ColWidth(i) = 1000
            ElseIf i = 3 Then
              .Text = " Email"
              .ColWidth(i) = 3500
            ElseIf i = 4 Then
               .ColWidth(i) = 0
            ElseIf i = 5 Then
               .ColWidth(i) = 0
            End If
            .ColAlignment(i) = 0
        Next
        setText msFlexEmails, 0, 0, 0
        setText msFlexEmails, 1, 0, 1
        .row = 1
        .col = 0
   End With
      
   'Set the header caption and width of msFlexFax
   With msFlexFax
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
              .Text = " Attention"
              .ColWidth(i) = 2300
            ElseIf i = 1 Then
              .Text = " Fax Number"
              .ColWidth(i) = 2300
            ElseIf i = 2 Then
               .ColWidth(i) = 0
            End If
            .ColAlignment(i) = 0
        Next
        .row = 1
        .col = 0
   End With
    
   mstrEmailLines = ""
   populatePhones
   populateFromRI
   populateFaxes
   
   'CC - V1.2.24 20140129 - CR 304 - JTB Integration
   EnableDisableFax
   
   If Trim(txtReplyEmail.Text) = "" Then
       txtReplyEmail.Text = UCase(GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjPNR.Agent, gobjPNR.PCCOwner, True, False))
   End If
   SftTabs.Tabs.Current = 0
   If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
   Else
      cmdPrevious.Visible = False
   End If
   datFormLoadEnd = Now
   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
   'If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
   
End Sub
Private Sub populateFaxes()
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    
    For i = 1 To gobjPNR.PhoneCount
        With gobjPNR.Phone(i)
             strTemp = ""
             If InStr(1, .PhoneNum, "F*") Then
                 strTemp = Mid(.PhoneNum, InStr(1, .PhoneNum, "F*") + 2)
             End If
             If strTemp <> "" Then
                With msFlexFax
                     If Trim(.TextMatrix(.rows - 1, 1)) = "" Then
                        .TextMatrix(.rows - 1, 1) = strTemp
                        Exit For
                        '.Rows = .Rows + 1
                     End If
                     '.TextMatrix(.Rows - 1, 1) = strTemp
                End With
             End If
        End With
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       gbolCancelProcess = True
    End If
End Sub

Public Sub msFlexEmails_Click()
    Dim intTop As Integer
    Dim intLeft As Integer

    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If

    intTop = topFrame.Top + fraEmails.Top + msFlexEmails.Top + msFlexEmails.CellTop
    intLeft = topFrame.Left + fraEmails.Left + msFlexEmails.Left + msFlexEmails.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If chkEntry.Visible = True Then chkEntry.Visible = False
  
    mstrFlex = msFlexEmails.Name

    If msFlexEmails.col = 0 And msFlexEmails.row = 0 Then
       optSelectAll msFlexEmails
    ElseIf msFlexEmails.col <= 1 And msFlexEmails.row > 0 Then
       setControlPosition msFlexEmails, chkEntry, intTop, intLeft
       checkedRow msFlexEmails
    ElseIf msFlexEmails.col = 2 And msFlexEmails.row > 0 Then
       cmbEntry.Clear
       'cmbEntry.AddItem "ARG - 2"
       '20081231
       'cmbEntry.AddItem "ARG"
       'cmbEntry.AddItem "DOM"
       'cmbEntry.AddItem "INT"
       'cmbEntry.AddItem "OTR"
       'cmbEntry.AddItem "PAX"
       'cmbEntry.AddItem "PER"
       
       cmbEntry.AddItem "ARG", 0
       cmbEntry.List(0, 1) = "ARRANGER"
       cmbEntry.AddItem "DOM", 1
       cmbEntry.List(1, 1) = "DOMESTIC APPROVER"
       cmbEntry.AddItem "INT", 2
       cmbEntry.List(2, 1) = "INTERATIONAL APPROVER"
       cmbEntry.AddItem "OTR", 3
       cmbEntry.List(3, 1) = "OTHERS"
       cmbEntry.AddItem "PAX", 4
       cmbEntry.List(4, 1) = "PASSENGER"
       cmbEntry.AddItem "PER", 5
       cmbEntry.List(5, 1) = "PERSONAL"
       
       setControlPosition msFlexEmails, cmbContainer, intTop, intLeft, cmbEntry
    ElseIf msFlexEmails.col = 3 And msFlexEmails.row > 0 Then
       setControlPosition msFlexEmails, txtEntry, intTop, intLeft
    End If
End Sub

Private Sub msFlexEmails_KeyDown(KeyCode As Integer, Shift As Integer)
    mstrFlex = msFlexEmails.Name
    control_KeyDown KeyCode, Shift, Me, msFlexEmails
End Sub

Public Sub msFlexFax_Click()
  Dim intTop As Integer
  Dim intLeft As Integer

  If mbolClickBelowRow = True Then
     mbolClickBelowRow = False
     Exit Sub
  End If

  intTop = topFrame.Top + fraFaxes.Top + msFlexFax.Top + msFlexFax.CellTop
  intLeft = topFrame.Left + fraFaxes.Left + msFlexFax.Left + msFlexFax.CellLeft
  If txtEntry.Visible = True Then txtEntry.Visible = False
  If cmbContainer.Visible = True Then cmbContainer.Visible = False
  mstrFlex = msFlexFax.Name

  If msFlexFax.col = 0 Or msFlexFax.col = 1 Then
     setControlPosition msFlexFax, txtEntry, intTop, intLeft
  End If
  
End Sub


Private Sub msFlexEmails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexEmails, Button, Y
End Sub

Private Sub msFlexFax_KeyDown(KeyCode As Integer, Shift As Integer)
    mstrFlex = msFlexFax.Name
    control_KeyDown KeyCode, Shift, Me, msFlexFax, False
End Sub

Private Sub msFlexFax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexFax, vbLeftButton, Y
End Sub

Public Sub subMenuAdd_Click()
    
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If chkEntry.value = True Then chkEntry.Visible = False
    
    If mstrFlex = msFlexEmails.Name Then
       msFlexEmails.rows = msFlexEmails.rows + 1
       setText msFlexEmails, msFlexEmails.rows - 1, 0, 1
    End If
End Sub

Public Sub subMenuDelete_Click()
    Dim i As Integer
    Dim msFlex As MSFlexGrid
    Dim strTemp As String
    Dim bolAddNewRow As Boolean
    Dim preRow As Integer
    
    bolAddNewRow = False
    
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If chkEntry.value = True Then chkEntry.Visible = False
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
    End If
    
    With msFlex
         'Checked how many rows selected
         i = rowSelected(msFlex)
         If i = 0 Then
            If .rows = 2 Then bolAddNewRow = True
         ElseIf i = .rows - 1 Then
            bolAddNewRow = True
         End If
         If bolAddNewRow = True Then
            'Must add new row since fixedRow need at least 1 row
            preRow = .row
            subMenuAdd_Click
            .row = preRow
         End If
         If i = 0 Then
            deleteRow msFlex, .row
         Else
             For i = 1 To .rows - 1
                 If i <= .rows - 1 Then
                    If .TextMatrix(i, 0) = gstrChecked Then
                       deleteRow msFlex, i
                       i = i - 1
                    End If
                 End If
             Next
         End If
    End With
End Sub

Private Sub txtAdRI_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii, "#$*()/.: ?@")
End Sub

Private Sub txtEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    If mstrFlex = msFlexEmails.Name Then
       control_KeyDown KeyCode, Shift, Me, txtEntry
    ElseIf mstrFlex = msFlexFax.Name Then
       control_KeyDown KeyCode, Shift, Me, txtEntry, False
    End If
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
    ElseIf mstrFlex = msFlexFax.Name Then
       Set msFlex = msFlexFax
    End If

    control_LostFocus msFlex, Me, txtEntry, , False
End Sub

Private Sub cmbEntry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If mstrFlex = msFlexEmails.Name Then
       control_KeyDown CInt(KeyCode), Shift, Me, cmbContainer
    ElseIf mstrFlex = msFlexFax.Name Then
       control_KeyDown CInt(KeyCode), Shift, Me, cmbContainer, False
    End If
End Sub

Private Sub cmbEntry_KeyPress(KeyCode As MSForms.ReturnInteger)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmbEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
    ElseIf mstrFlex = msFlexFax.Name Then
       Set msFlex = msFlexFax
    End If

    control_LostFocus msFlex, Me, cmbEntry, , False
End Sub

Private Sub txtReplyEmail_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub writeDatatoGDS()
    Dim strCmd As String
    Dim strCmd2 As String
    Dim strCmdRemove As String
    Dim strField() As String
    Dim strTemp As String
    Dim strTemp2 As String
    Dim strResponse As String
    Dim strQueue As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim strMsg As String

    Dim strQKey As String
        
    strQKey = pAddToAQQueueLog
    strCmd = ""
    strCmd2 = ""
    strQueue = ""
    If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
    gbolWritingtoPNR = True
    
    'Email Itinerary
    If chkEmail.value = 1 Then
        'Email Fields (RI & P.)
        With msFlexEmails
             For i = 1 To .rows - 1
                 If Len(.TextMatrix(i, 4)) >= 2 Then
                    If Right(.TextMatrix(i, 4), 1) = "." Then .TextMatrix(i, 4) = Mid(.TextMatrix(i, 4), 1, Len(.TextMatrix(i, 4)) - 1)
                 End If
                 .TextMatrix(i, 5) = Replace(.TextMatrix(i, 5), "PHONE-", "")
                 If .TextMatrix(i, 5) = "" Then
                    strCmd = strCmd & IIf(strCmd = "", "", "+") & "P." & gstrAgcyCityCode & "E*" & .TextMatrix(i, 2) & "-" & convertPhoneText(.TextMatrix(i, 3))
                 Else
                    strTemp = gstrAgcyCityCode & "E*" & .TextMatrix(i, 2) & "-" & convertPhoneText(.TextMatrix(i, 3))
                    strTemp2 = gobjPNR.Phone(CInt(.TextMatrix(i, 5))).PhoneNum
                    If UCase(strTemp) <> UCase(strTemp2) Then
                       strCmd = strCmd & IIf(strCmd = "", "", "+") & "P." & .TextMatrix(i, 5) & "@" & strTemp
                    End If
                 End If
                 strField = Split(.TextMatrix(i, 4), ".")
                 strTemp = Trim(.TextMatrix(i, 3))
                 If strTemp <> "" Then
                    For j = 1 To 1
                        bolExist = False
                        strTemp = convertText(Trim(.TextMatrix(i, 3)))
                        strTemp2 = Trim(.TextMatrix(i, j))
                        strTemp3 = ""
                        If j = 1 Then
                           strTemp3 = "ITI"
                        End If
                        If strTemp2 = gstrChecked Then
                           For k = 0 To UBound(strField)
                               l = InStr(1, strField(k), strTemp3)
                               If l > 0 Then
                                  l = Mid(strField(k), l + 4)
                                  bolExist = True
                                  With gobjPNR.ItinRemark(l)
                                       If UCase(strTemp3 & "." & strTemp) <> UCase(.RemarkText) Then
                                          strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & l & "@" & strTemp3 & "." & strTemp
                                       End If
                                  End With
                                  Exit For
                               End If
                           Next
                           If bolExist = False Then
                              strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI./0*" & strTemp3 & "." & strTemp
                           End If
                        Else
                           'Delete existing RI if field is blank
                           For k = 0 To UBound(strField)
                               l = InStr(1, strField(k), strTemp3)
                               If l > 0 Then
                                  strCmd2 = strCmd2 & IIf(strCmd2 = "", "", ".") & Mid(strField(k), l + 4)
                                  Exit For
                               End If
                           Next
                        End If
                    Next
                 End If
            Next
        End With
                       
        'Reply email address
        If Trim(UCase(txtReplyEmail.Text)) <> Trim(UCase(GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjPNR.Agent, gobjPNR.PCCOwner, True, False))) Then
           updateReplyEmail Trim(txtReplyEmail.Text), "EMA", strCmd, strCmd2
        Else
           updateReplyEmail "", "EMA", strCmd, strCmd2
        End If
             
        'Add this for eItin remarks comparison
        'gobjHost.TerminalEntry "NP.I*EITINQUEUETIME:" & Format(Now, "ddmmmhh:ss")
             
        'This is for additional remarks
        splitLongRI "I", strQKey
             
        strQueue = "QEB/5E4P/10"
        'strQueue = "QEB/5E4P/72"
        
    End If
    
    'Fax Itinerary
    If chkEFax.value = 1 Then
       RemoveRI strCmd2
       AddRI strCmd
       strQueue = strQueue & IIf(strQueue = "", "QEB/5E4P/30", "+5E4P/30")
    End If
    
    If strCmd2 <> "" Or mstrEmailLines <> "" Then
      'Delete existing RI that stored in strCmd2
      strCmd2 = mstrEmailLines & strCmd2
      If Len(strCmd2) >= 2 Then
         If Right(strCmd2, 1) = "." Then strCmd2 = Mid(strCmd2, 1, Len(strCmd2) - 1)
      End If
      If strCmd2 <> "" Then
         strCmd2 = "RI." & sortInt(strCmd2) & "@"
         strCmdRemove = strCmd2
         'strCmd = strCmd & IIf(strCmd = "", "", "+") & strCmd2
         strCmd2 = ""
      End If
    End If

    'Preethi - V1.2.4 20110614 - CR 76 - Change Validation Logic For ENDPNR
    'send entries, received & end the PNR
    strCmd = strCmd & IIf(strCmd = "", "", "+") & "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine) & "+ER+ER+ER"
    
    If strCmdRemove <> "" Then
        strResponse = gobjHost.terminalEntry(strCmdRemove)
    End If
    
    strResponse = gobjHost.terminalEntry(strCmd)
    strTemp = strResponse
    'Preethi - V1.2.4 20110712 - CR 76 - Change Validation Logic For ENDPNR
    If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = False Then
    'If InStr(strTemp, "1.1") = 0 Then
       For i = 0 To 1
           strTemp = gobjHost.terminalEntry("ER")
           'If InStr(strTemp, "1.1") > 0 Then
           If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = True Then
               Exit For
           End If
           If i = 1 Then GoTo errorWriting
       Next
    End If
'    strTemp = gobjHost.EndPNR2(IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine), True, 2)
'    If strTemp <> "True" Then GoTo errorWriting
        
    If chkEmail.value = 1 Then
       If pAddToQueueLog(gobjPNR.RecLoc, "EItin") = False Then
          strMsg = "Cannot send E-Itinerary." & vbCrLf & "Cannot add queue key to database."
          modMsgBox.OKMsg = "OK"
          modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Add Queue Key"
          'ZhiSam - V1.2.1 20130724 - CR227 - Disable E-DOC & E-ITIN function to sent out twice
          cmdFinish.Enabled = True
          Exit Sub
       End If
    End If
    
    If chkEFax.value = 1 Then
       If pAddToQueueLog(gobjPNR.RecLoc, "EFAX") = False Then
          strMsg = "Cannot send E-Fax." & vbCrLf & "Cannot add queue key to database."
          modMsgBox.OKMsg = "OK"
          modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Add Queue Key"
          'ZhiSam - V1.2.1 20130724 - CR227 - Disable E-DOC & E-ITIN function to sent out twice
          cmdFinish.Enabled = True
          Exit Sub
       End If
    End If
    
    strResponse = gobjHost.terminalEntry(strQueue)
    If InStr(strResponse, "ON QUEUE") > 0 Or InStr(strResponse, gobjPNR.RecLoc) = 0 Then
       gobjLog.LineTextToLog "PNR QUEUED"
       frmSideBar.fraInfo.Caption = " PNR -> "
       frmSideBar.treeViewTraveller.Nodes.Clear
    Else
       strMsg = "Cannot send E-Itinerary/E-Fax. Response from GDS: " & strResponse & vbCrLf & "Please recify PNR problem and re-send."
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Queue PNR"
       'ZhiSam - V1.2.1 20130724 - CR227 - Disable E-DOC & E-ITIN function to sent out twice
       cmdFinish.Enabled = True
       Exit Sub
    End If
    
    Exit Sub
    
errorWriting:
    'Prompt error message if failed to write to PNR
    gbolWritingtoPNR = False
    strMsg = "Unable to write to PNR. Response from GDS is " & Chr(13) & strResponse
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
    'gobjHost.TerminalEntry "IR"
    'ZhiSam - V1.2.1 20130724 - CR227 - Disable E-DOC & E-ITIN function to sent out twice
    cmdFinish.Enabled = True
    
End Sub

Private Function validData() As Boolean
    Dim strMsg As String
    Dim bolEmpty As Boolean
    Dim i As Integer
    
    bolEmpty = True
    validData = True
    
    If chkEmail.value = 0 And chkEFax.value = 0 Then
      strMsg = strMsg & "No sending option is selected ..." & Chr(13)
    End If
    
    If chkEmail.value = 1 Then
        'Email is mandatory field
        With msFlexEmails
             For i = 1 To .rows - 1
                 If .TextMatrix(i, 1) = gstrChecked Then bolEmpty = False
                 If Trim(.TextMatrix(i, 2)) = "" Then
                     strMsg = strMsg & "Missing email type for record " & i & " ..." & Chr(13)
                 End If
                 If Trim(.TextMatrix(i, 3)) = "" Then
                     strMsg = strMsg & "Missing email address for record " & i & " ..." & Chr(13)
                 End If
             Next
        End With
                
        If bolEmpty = True And InStr(1, strMsg, "Missing email address") = 0 Then strMsg = strMsg & "Email info is required ..." & Chr(13)
        
        If Trim(txtReplyEmail.Text) = "" Then
           strMsg = strMsg & "Missing reply address by client ..." & Chr(13)
        ElseIf InStr(Trim(txtReplyEmail.Text), ";") > 0 Then
           strMsg = strMsg & "Only 1 email address is allowed in reply address by client ..." & Chr(13)
        ElseIf InStr(Trim(txtReplyEmail.Text), "@") = 0 And Trim(txtReplyEmail.Text) <> "" Then
           strMsg = strMsg & "Invalid email address in reply address by client ..." & Chr(13)
        End If
 
    End If
    If chkEFax.value = 1 Then
       bolEmpty = True
       With msFlexFax
            For i = 1 To .rows - 1
                 If Trim(.TextMatrix(i, 0)) <> "" Or Trim(.TextMatrix(i, 1)) <> "" Then
                     If Trim(.TextMatrix(i, 0)) = "" Then
                        strMsg = strMsg & "Missing fax recipient for record " & i & " ..." & Chr(13)
                     ElseIf Trim(.TextMatrix(i, 1)) = "" Then
                        strMsg = strMsg & "Missing fax number for record " & i & " ..." & Chr(13)
                     End If
                 End If
                 If Trim(.TextMatrix(i, 0)) <> "" And Trim(.TextMatrix(i, 1)) <> "" Then
                     bolEmpty = False
                 End If
            Next
            If bolEmpty = True And InStr(1, strMsg, "Missing fax") = 0 Then strMsg = strMsg & "Fax info is required ..." & Chr(13)
       End With
    End If
    If strMsg <> "" Then
       validData = False
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    End If
End Function

Public Sub updateReplyEmail(strReplyEmail As String, strQualifier As String, ByRef strCmd As String, ByRef strCmd2 As String)
    Dim intI As Integer
    For intI = 1 To gobjPNR.ItinRemarkCount
        If Left(gobjPNR.ItinRemark(intI).RemarkText, 4) = strQualifier & "." And _
           Len(gobjPNR.ItinRemark(intI).RemarkText) > 4 Then
           Exit For
        End If
    Next
    If intI <= gobjPNR.ItinRemarkCount Then
       If strReplyEmail = "" Then
          strCmd2 = strCmd2 & IIf(strCmd2 = "", "", ".") & intI
       Else
          strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & intI & "@" & strQualifier & "." & convertText(strReplyEmail)
       End If
    Else
       If strReplyEmail <> "" Then
          strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & strQualifier & "." & convertText(strReplyEmail)
       End If
    End If
End Sub

Private Sub RemoveRI(ByRef strCmd2 As String)
   Dim i As Integer
   For i = 1 To gobjPNR.ItinRemarkCount
       With gobjPNR.ItinRemark(i)
            If .SegNum = 0 Then
               If UCase(Left(.RemarkText, 7)) = "FAXNAME" Or _
                  UCase(Left(.RemarkText, 5)) = "FAXNO" Then
                  strCmd2 = strCmd2 & IIf(strCmd2 = "", "", ".") & i
               End If
            End If
       End With
   Next
End Sub

Private Sub AddRI(ByRef strCmd As String)
     strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI./0*" & "FAXNAME " & msFlexFax.TextMatrix(1, 0)
     strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI./0*" & "FAXNO " & msFlexFax.TextMatrix(1, 1)
End Sub

Private Sub mouseDown(ByRef msFlex As MSFlexGrid, ByVal Button As Integer, ByVal Y As Single)
    With msFlex
         If txtEntry.Visible = True Then txtEntry.Visible = False
         If cmbContainer.Visible = True Then cmbContainer.Visible = False
         If chkEntry.Visible = True Then chkEntry.Visible = False
         
         mstrFlex = .Name
         .row = .MouseRow
         .col = .MouseCol
        
         If Button = vbRightButton Then
             PopupMenu mnuPopUpFlex
         ElseIf Button = vbLeftButton Then
             If Y > .RowPos(.rows - 1) + .RowHeight(.rows - 1) Then
                mbolClickBelowRow = True
             Else
                mbolClickBelowRow = False
             End If
         End If
    End With
End Sub

Private Sub populatePhones()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim phoneType As Integer
    Dim strTemp As String
    Dim bolInPhone As Boolean
    
    For i = 1 To gobjPNR.PhoneCount
        With gobjPNR.Phone(i)
             phoneType = 0
             If InStr(1, .PhoneNum, "E*") Then
                 phoneType = 5 'Email
                 strTemp = Mid(.PhoneNum, InStr(1, .PhoneNum, "E*") + 2)
                 strTemp = Replace(strTemp, "PAX-", "")
                 strTemp = Replace(strTemp, "PER-", "")
                 strTemp = Replace(strTemp, "OTR-", "")
                 strTemp = Replace(strTemp, "ARG-", "")
                 strTemp = Replace(strTemp, "DOM-", "")
                 strTemp = Replace(strTemp, "INT-", "")
                 strTemp = actualPhoneText(strTemp)
             End If
             If phoneType = 5 Then
                 If Trim(msFlexEmails.TextMatrix(1, 3)) <> "" Then
                    msFlexEmails.rows = msFlexEmails.rows + 1
                    msFlexEmails.row = msFlexEmails.rows - 1
                    setText msFlexEmails, msFlexEmails.row, 0, 1
                 Else
                    msFlexEmails.row = 1
                 End If
                 If InStr(1, .PhoneNum, "E*PAX-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 2) = "PAX"
                 ElseIf InStr(1, .PhoneNum, "E*PER-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 2) = "PER"
                 ElseIf InStr(1, .PhoneNum, "E*OTR-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 2) = "OTR"
                 ElseIf InStr(1, .PhoneNum, "E*ARG-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 2) = "ARG"
                 ElseIf InStr(1, .PhoneNum, "E*DOM-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 2) = "DOM"
                 ElseIf InStr(1, .PhoneNum, "E*INT-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 2) = "INT"
                 End If
                 msFlexEmails.TextMatrix(msFlexEmails.row, 3) = strTemp
                 msFlexEmails.TextMatrix(msFlexEmails.row, 5) = "PHONE-" & i
             End If
        End With
    Next
    
    
     For i = 1 To gobjPNR.ItinRemarkCount
    
    If Left(gobjPNR.ItinRemark(i).RemarkText, 4) = "ITI." And _
         Len(gobjPNR.ItinRemark(i).RemarkText) > 4 Then
         strTemp = actualText(Mid(gobjPNR.ItinRemark(i).RemarkText, 5))
         bolInPhone = False
         For j = 1 To msFlexEmails.rows - 1
            If Trim(msFlexEmails.TextMatrix(j, 3)) = strTemp Then
               bolInPhone = True
            End If
         Next
    
    
    If bolInPhone = False Then
    If Trim(msFlexEmails.TextMatrix(1, 3)) <> "" Then
       msFlexEmails.rows = msFlexEmails.rows + 1
       msFlexEmails.row = msFlexEmails.rows - 1
       setText msFlexEmails, msFlexEmails.row, 0, 1
    Else
       msFlexEmails.row = 1
    End If
    msFlexEmails.TextMatrix(msFlexEmails.row, 3) = strTemp
    End If
           
    End If
    
    
    Next
    
End Sub

Private Sub populateFromRI()
    Dim i As Integer
    Dim bolExist As Boolean
    
    For i = 1 To gobjPNR.ItinRemarkCount
        With gobjPNR.ItinRemark(i)
            If Left(.RemarkText, 4) = "ITI." And _
               Len(.RemarkText) > 4 Then
               bolExist = dataExistInCell(msFlexEmails, actualText(Mid(.RemarkText, 5)))
               If bolExist = True Then
                  insertIntoEmailFlex bolExist, i, 1, gstrChecked, Mid(.RemarkText, 5)
               Else
                  mstrEmailLines = mstrEmailLines & i & "."
               End If
            ElseIf Left(.RemarkText, 5) = "ITIX." And _
               Len(.RemarkText) > 5 Then
               bolExist = dataExistInCell(msFlexEmails, actualText(Mid(.RemarkText, 6)))
               If bolExist = True Then
                  insertIntoEmailFlex bolExist, i, 1, gstrUnChecked, Mid(.RemarkText, 6)
               Else
                  mstrEmailLines = mstrEmailLines & i & "."
               End If
            ElseIf Left(.RemarkText, 5) = "ITID." And _
               Len(.RemarkText) > 5 Then
               bolExist = dataExistInCell(msFlexEmails, actualText(Mid(.RemarkText, 6)))
               If bolExist = True Then
                  insertIntoEmailFlex bolExist, i, 1, gstrChecked, Mid(.RemarkText, 6)
               Else
                  mstrEmailLines = mstrEmailLines & i & "."
               End If
            ElseIf Left(.RemarkText, 6) = "ITIDX." And _
               Len(.RemarkText) > 6 Then
               bolExist = dataExistInCell(msFlexEmails, actualText(Mid(.RemarkText, 7)))
               If bolExist = True Then
                  insertIntoEmailFlex bolExist, i, 1, gstrUnChecked, Mid(.RemarkText, 7)
               Else
                  mstrEmailLines = mstrEmailLines & i & "."
               End If
            ElseIf Left(.RemarkText, 4) = "EMA." And _
               Len(.RemarkText) > 4 Then
               txtReplyEmail.Text = actualText(Mid(.RemarkText, 5))
            End If
        End With
   Next
End Sub

Private Function dataExistInCell(ByRef msFlex As MSFlexGrid, strValue As String) As Boolean
    Dim i As Integer
    
    For i = 1 To msFlex.rows - 1
        If UCase(Trim(msFlex.TextMatrix(i, 3))) = UCase(Trim(strValue)) Then
           dataExistInCell = True
           msFlex.row = i
           Exit For
        End If
    Next

End Function

Private Sub insertIntoEmailFlex(bolExist As Boolean, intRI As Integer, colYesNo As Integer, strChecked As String, strEmail As String)
    Dim strTemp As String
    
    If colYesNo = 1 Then
       strTemp = "ITI"
    End If
    msFlexEmails.TextMatrix(msFlexEmails.row, colYesNo) = strChecked
    msFlexEmails.TextMatrix(msFlexEmails.row, 4) = msFlexEmails.TextMatrix(msFlexEmails.row, 4) & strTemp & "-" & intRI & "."

End Sub

Private Sub deleteRow(ByRef msFlex As MSFlexGrid, ByVal i As Integer)
           
   Dim strTemp As String
       
   With msFlex
        If .Name = msFlexEmails.Name Then
            If .TextMatrix(i, 4) <> "" Then
                strTemp = Replace(.TextMatrix(i, 4), "ITI-", "")
                mstrEmailLines = mstrEmailLines & strTemp
            End If
        End If
       msFlex.RemoveItem (i)
       msFlex.col = 1
       If i <= .rows - 1 Then
          msFlex.row = i
       Else
          msFlex.row = i - 1
       End If
   End With
End Sub

Private Sub checkedRow(ByRef msFlex As MSFlexGrid)
    If chkEntry.value = vbChecked And gintY = 0 Then
       HighlightRow msFlex, gintX
    ElseIf chkEntry.value = vbUnchecked And gintY = 0 Then
       HighlightRow msFlex, gintX, False
    End If
End Sub

Private Sub chkEntry_Click()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
    End If
    checkedRow msFlex
End Sub

Private Sub chkEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
       'Clement - 20080812
       'msFlexEmails.SetFocus
    End If
    control_LostFocus msFlex, Me, chkEntry
End Sub

'Public Function pGetQueueKey() As String
'
'Dim strQKey As String
''Get random number until not found in database
'strQKey = Random
'Do While pCheckQKey(strQKey)
'   strQKey = Random
'Loop
'
'pGetQueueKey = strQKey
'End Function
'
'Private Function pCheckQKey(ByVal strQKey As String) As Boolean
'Dim rsQueueLog As ADODB.Recordset
'Dim strSQL As String
'
'strSQL = "Select * from tblQueueTime Where QKey='" & strQKey & "'"
'Set rsQueueLog = gdbEitinConn.Execute(strSQL)
'If rsQueueLog.EOF Then
'   pCheckQKey = False
'Else
'  pCheckQKey = True
'End If
'End Function

Public Function splitLongRI(strQualifier As String, qKey As String) As String
Dim lngC As Long
Dim lngLen As Long
Dim strTemp As String
Dim strChar As String
Dim strEntry As String
Dim strEntryText As String
Dim strNow As String

lngC = 0
strTemp = ""
strEntry = ""
strNow = Format(Now, "ddmmmhh:ss")
'& qKey

'For lngC = 0 To txtAdRI.Count - 1
'
'    If Trim(txtAdRI(lngC)) <> "" Then
'        strEntry = strEntry & IIf(strEntry = "", "", "+") & "RI." & strQualifier & "." & strNow & "." & txtAdRI(lngC)
'
'    End If
'
'Next

'Change eitin remarks from RI to NP.HI on 24 Sep 2008 by Jeremy
For lngC = 0 To txtAdRI.Count - 1

    If Trim(txtAdRI(lngC)) <> "" Then
        strEntry = strEntry & IIf(strEntry = "", "", "+") & "NP.HI*" & strQualifier & "." & strNow & "." & txtAdRI(lngC)
    End If

Next

gobjHost.terminalEntry strEntry

For lngC = 1 To gobjPNR.GeneralRemarkCount

    If InStr(gobjPNR.GeneralRemark(lngC).RemarkText, "EITINQUEUETIME:") And gobjPNR.GeneralRemark(lngC).Qualifier = "*I" Then
        gobjHost.terminalEntry "NP." & gobjPNR.GeneralRemark(lngC).ItemNum & "@"
    End If
    
Next

gobjHost.terminalEntry "NP.I*EITINQUEUETIME:" & strNow
    
End Function

'CC - V1.2.24 20140129 - CR 304 - JTB Integration
Private Sub EnableDisableFax()
    If UCase(gobjPNR.CompInfo.AgencyName) = "JTB" Then
        fraFaxes.Visible = False
    Else
        fraFaxes.Visible = True
    End If
End Sub
