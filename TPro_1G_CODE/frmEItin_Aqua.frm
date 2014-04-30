VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{4832871B-0993-461C-B983-0EAAA4A43E5C}#5.0#0"; "SftTabs_IX86_U_50.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient.ocx"
Begin VB.Form frmEItin_Aqua 
   Caption         =   "CWT Desktop - Aqua Itinerary"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   11595
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
         xBfStyle2       =   -69159233
         xBfStyle3       =   -116221709
         xBfStyle4       =   384657181
         TabCount        =   1
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
            Left            =   0
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
            Begin VB.Frame fraEmails 
               BackColor       =   &H00DADAB6&
               Caption         =   "Sending Info"
               Height          =   2295
               Left            =   100
               TabIndex        =   1
               Top             =   80
               Width           =   10815
               Begin MSFlexGridLib.MSFlexGrid msFlexEmails 
                  Height          =   1995
                  Left            =   60
                  TabIndex        =   2
                  Top             =   240
                  Width           =   5565
                  _ExtentX        =   9816
                  _ExtentY        =   3519
                  _Version        =   393216
                  Cols            =   7
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
         End
      End
      Begin MSComctlLib.ListView lvwMappingTable 
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "To"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox cmbContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   1005
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   1000
         Begin MSForms.ComboBox cmbEntry 
            Height          =   375
            Left            =   60
            TabIndex        =   6
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
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CheckBox chkEntry 
         Height          =   200
         Left            =   3000
         TabIndex        =   3
         Top             =   3060
         Visible         =   0   'False
         Width           =   195
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
         LcK2            =   $"frmEItin_Aqua.frx":0000
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
      Begin MyCommandButton.MyButton cmdFinish 
         Height          =   360
         Left            =   9180
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
Attribute VB_Name = "frmEItin_Aqua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'CC - V1.2.16 20121018 - CR163 - EM trigger generation removal in QUEUE module
'Public mstrCallFrom As String
Public mstrActivateFrom As String

Dim mstrFlex As String
Dim mbolClickBelowRow As Boolean
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date

Enum mEmailCol
    CheckBox_mEmail = 0         '0
    EItin_mEmail = 1            '1
    ETkt_mEmail = 2
    Type_mEmail = 3             '2
    Email_mEmail = 4            '3
    LineNum_mEmail = 5        '4
    PNRLoc = 6            '5
End Enum

Private Sub cmbEntry_GotFocus()
    cmbGetFocus cmbEntry
End Sub

Private Sub cmdCancel_Click()
    gbolCancelProcess = True
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    Dim bolPass As Boolean
    Dim intMsg As Integer
    Dim strMsg As String
    
    'CC - V1.2.24 20140129 - CR 304 - JTB Integration
    'intMsg = MsgBox("Has the E itinerary recipient email address been updated in RI.ITI Field?", vbYesNo, "CWT Desktop - E-Itinerary")
    strMsg = "Has the E itinerary recipient email address been updated in RI.ITI Field?"
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    intMsg = modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbYesNo + vbDefaultButton1, "CWT Desktop - E-Itinerary")
    
    If intMsg = vbNo Then
        Exit Sub
    End If
    
    'ZhiSam - V1.2.1 20130724 - CR227 - Disable E-DOC & E-ITIN function to sent out twice
    cmdFinish.Enabled = False
    
    If ValidEmail Then
        bolPass = UpdatePNR
        If bolPass = True Then
            gobjHost.terminalEntry "I"
            Set gobjPNR = New CWT_GalileoPNR3.PNR
            gobjPNR.loadPNR
            displayPNRinBar
            
            Unload Me
        Else
            'gobjHost.terminalEntry "IR"
            'ZhiSam - V1.2.1 20130724 - CR227 - Disable E-DOC & E-ITIN function to sent out twice
            cmdFinish.Enabled = True
        End If
        '?????
        'If gbolWritingtoPNR = False Then Exit Sub
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
   Dim lngW As Long
   
   Dim objEmails As New EmailAddresses
   
   datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   
   'CC - V1.2.16 20121018 - CR163 - EM trigger generation removal in QUEUE module
   mstrActivateFrom = ""
   
   pDisplayToFP "*R"
   
   GetMappingTable
   
   cmbEntry.ColumnCount = 2
   cmbEntry.ColumnWidths = "30,150"
   
   lngW = 0
   'Set the header caption and width of msFlexEmails
    With msFlexEmails
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = mEmailCol.CheckBox_mEmail Then
              .Text = ""
              .ColWidth(i) = 280
            ElseIf i = mEmailCol.EItin_mEmail Then
              .Text = " E-Itin "
              .ColWidth(i) = 600
            ElseIf i = mEmailCol.ETkt_mEmail Then
              .Text = " E-Tkt "
              .ColWidth(i) = 600
            ElseIf i = mEmailCol.Type_mEmail Then
              .Text = " Type"
              .ColWidth(i) = 1200   '1000
            ElseIf i = mEmailCol.Email_mEmail Then
              .Text = " Email"
              .ColWidth(i) = 5500   '3500
            ElseIf i = mEmailCol.LineNum_mEmail Then
               .ColWidth(i) = 0
            ElseIf i = mEmailCol.PNRLoc Then
               .ColWidth(i) = 0
            End If
            .ColAlignment(i) = 0
            lngW = lngW + .ColWidth(i)
        Next
        .Width = lngW + 500
        setText msFlexEmails, 0, 0, 0
        setText msFlexEmails, 1, mEmailCol.CheckBox_mEmail, mEmailCol.CheckBox_mEmail
        setText msFlexEmails, 1, mEmailCol.EItin_mEmail, mEmailCol.EItin_mEmail
        setText msFlexEmails, 1, mEmailCol.ETkt_mEmail, mEmailCol.ETkt_mEmail
        .row = 1
        .col = 0
   End With
       
   Set objEmails = GetEmailFromPNR
   PopulateEmailFlex objEmails
   
   'SftTabs.Tabs.Current = 0
   If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
   Else
      cmdPrevious.Visible = False
   End If
   datFormLoadEnd = Now
'   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
   
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

    If msFlexEmails.col = mEmailCol.CheckBox_mEmail And msFlexEmails.row = 0 Then
       optSelectAll msFlexEmails
'    ElseIf msFlexEmails.col <= 1 And msFlexEmails.row > 0 Then
'       setControlPosition msFlexEmails, chkEntry, intTop, intLeft
'       checkedRow msFlexEmails
    ElseIf (msFlexEmails.col = mEmailCol.CheckBox_mEmail Or _
        msFlexEmails.col = mEmailCol.EItin_mEmail Or _
        msFlexEmails.col = mEmailCol.ETkt_mEmail) _
        And msFlexEmails.row > 0 Then
       setControlPosition msFlexEmails, chkEntry, intTop, intLeft
       checkedRow msFlexEmails
    ElseIf msFlexEmails.col = mEmailCol.Type_mEmail And msFlexEmails.row > 0 Then
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
    ElseIf msFlexEmails.col = mEmailCol.Email_mEmail And msFlexEmails.row > 0 Then
       setControlPosition msFlexEmails, txtEntry, intTop, intLeft
    End If
End Sub

Private Sub msFlexEmails_KeyDown(KeyCode As Integer, Shift As Integer)
    mstrFlex = msFlexEmails.Name
    control_KeyDown KeyCode, Shift, Me, msFlexEmails
End Sub

Private Sub msFlexEmails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexEmails, Button, Y
End Sub

Public Sub subMenuAdd_Click()
    
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If chkEntry.value = True Then chkEntry.Visible = False
    
    If mstrFlex = msFlexEmails.Name Then
       msFlexEmails.rows = msFlexEmails.rows + 1
'       setText msFlexEmails, msFlexEmails.rows - 1, 0, 1
       setText msFlexEmails, msFlexEmails.rows - 1, mEmailCol.CheckBox_mEmail, mEmailCol.CheckBox_mEmail
       setText msFlexEmails, msFlexEmails.rows - 1, mEmailCol.EItin_mEmail, mEmailCol.EItin_mEmail
       setText msFlexEmails, msFlexEmails.rows - 1, mEmailCol.ETkt_mEmail, mEmailCol.ETkt_mEmail
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
                    If .TextMatrix(i, mEmailCol.CheckBox_mEmail) = gstrChecked Then
                       deleteRow msFlex, i
                       i = i - 1
                    End If
                 End If
             Next
         End If
    End With
End Sub

'Private Sub txtAdRI_KeyPress(Index As Integer, KeyAscii As Integer)
'    KeyAscii = fAllowAlphaNumeric(KeyAscii, "#$*()/.: ?@")
'End Sub

Private Sub txtEntry_KeyDown(KeyCode As Integer, Shift As Integer)
'    If mstrFlex = msFlexEmails.Name Then
       control_KeyDown KeyCode, Shift, Me, txtEntry
'    ElseIf mstrFlex = msFlexFax.Name Then
'       control_KeyDown KeyCode, Shift, Me, txtEntry, False
'    End If
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
'    ElseIf mstrFlex = msFlexFax.Name Then
'       Set msFlex = msFlexFax
    End If

    control_LostFocus msFlex, Me, txtEntry, , False
End Sub

Private Sub cmbEntry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If mstrFlex = msFlexEmails.Name Then
       control_KeyDown CInt(KeyCode), Shift, Me, cmbContainer
'    ElseIf mstrFlex = msFlexFax.Name Then
'       control_KeyDown CInt(KeyCode), Shift, Me, cmbContainer, False
    End If
End Sub

Private Sub cmbEntry_KeyPress(KeyCode As MSForms.ReturnInteger)
   KeyCode = Asc(UCase(Chr(KeyCode)))
End Sub

Private Sub cmbEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
'    ElseIf mstrFlex = msFlexFax.Name Then
'       Set msFlex = msFlexFax
    End If

    control_LostFocus msFlex, Me, cmbEntry, , False
End Sub

'Private Sub txtReplyEmail_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

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

Private Sub deleteRow(ByRef msFlex As MSFlexGrid, ByVal i As Integer)
   Dim strTemp As String
       
   With msFlex
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

Private Function GetEmailFromPNR() As EmailAddresses
    Dim intI As Integer
    Dim strEMailText As String
    Dim strPaxType As String
    Dim strEmailAddr As String
    Dim strEmailType As String
    Dim intLineNum As Integer
    Dim objEmails As New EmailAddresses
    Dim strTemp As String
    
                                '1234567890
    'gobjPNR.Phone(8).PhoneNum = SINE*PAX-CPIYADA//CARLSONWAGONLIT.COM
    'gobjPNR.Phone(8).PhoneNum = SINE*PAX-CPIYADA//CARLSONWAGONLIT.COM ITI
    For intI = 1 To gobjPNR.PhoneCount
        strEMailText = gobjPNR.Phone(intI).PhoneNum
        If Mid(UCase(strEMailText), 4, 2) = "E*" And Mid(strEMailText, 9, 1) = "-" Then
            strPaxType = Mid(strEMailText, 6, 3)
            strTemp = Mid(strEMailText, 10)
            intLineNum = gobjPNR.Phone(intI).ItemNum
            If InStr(1, strTemp, " ") > 1 Then
                strEmailAddr = Mid(strTemp, 1, InStr(1, strTemp, " ") - 1)
                strEmailType = Mid(strTemp, InStr(1, strTemp, " ") + 1)
            Else
                strEmailAddr = strTemp
                strEmailType = ""
            End If
            strEmailAddr = actualPhoneText(strEmailAddr)
            
            AddEmail strEmailAddr, strEmailType, strPaxType, "P", intLineNum, objEmails
        End If
    Next
    
    'gobjPNR.ItinRemark(1).RemarkText = TKT.CCHING@CARLSONWAGONLIT.COM
    For intI = 1 To gobjPNR.ItinRemarkCount
        With gobjPNR.ItinRemark(intI)
            strEmailType = UCase(Mid(.RemarkText, 1, 3))
            If (strEmailType = "ITI" Or strEmailType = "TKT") And _
                Mid(.RemarkText, 4, 1) = "." Then
                strEmailAddr = Mid(.RemarkText, 5)
                strEmailAddr = actualText(strEmailAddr)
                intLineNum = .ItemNum
                strPaxType = ""
                 
                AddEmail strEmailAddr, strEmailType, strPaxType, "RI", intLineNum, objEmails
            End If
        End With
    Next
    
    Set GetEmailFromPNR = objEmails
End Function

Private Sub AddEmail(EmailAddr As String, EmailType As String, PaxType As String, PNRLoc As String, LineNum As Integer, ByRef Emails As EmailAddresses)
    Dim intI As Integer
    Dim bolFound As Boolean
    Dim objEmail As New EmailAddress
    
    bolFound = False
    EmailAddr = UCase(EmailAddr)
    For intI = 1 To Emails.EmailCount
        With Emails.Email(intI)
            If UCase(.EmailAddress) = EmailAddr Then
                If .EInv = False And .EItin = False And .ETkt = False Then
                    Select Case UCase(EmailType)
                        Case "ITI"
                            .EItin = True
                        Case "TKT"
                            .ETkt = True
                        Case "INV"  'For future, INV is not using at the moment
                            .EInv = True
                    End Select
                End If
                If .PaxType = "" And PaxType <> "" Then
                    .PaxType = UCase(PaxType)
                End If
                bolFound = True
                Exit For
            End If
        End With
    Next
    
    If bolFound = False Then
        With objEmail
            Select Case UCase(EmailType)
                Case "ITI"
                    .EItin = True
                Case "INV"
                    .EInv = True
                Case "TKT"
                    .ETkt = True
            End Select
            .EmailAddress = EmailAddr
            .LineNum = LineNum
            .PaxType = UCase(PaxType)
            .PNRLoc = UCase(PNRLoc)
            
            Emails.AddEmail objEmail
        End With
    End If
End Sub

Private Sub PopulateEmailFlex(Emails As EmailAddresses)
    Dim intI As Integer
    
    If Emails.EmailCount = 0 Then Exit Sub
    
'    If Trim(msFlexEmails.TextMatrix(1, mEmailCol.Email_mEmail)) <> "" Then
'        msFlexEmails.rows = msFlexEmails.rows + 1
'        msFlexEmails.row = msFlexEmails.rows - 1
'    Else
'        msFlexEmails.row = 1
'    End If
    
    For intI = 1 To Emails.EmailCount
        If intI = 1 Then
            msFlexEmails.row = 1
        Else
            msFlexEmails.rows = msFlexEmails.rows + 1
            msFlexEmails.row = msFlexEmails.rows - 1
        End If
        With Emails.Email(intI)

            setText msFlexEmails, msFlexEmails.row, mEmailCol.CheckBox_mEmail, mEmailCol.CheckBox_mEmail
            setText msFlexEmails, msFlexEmails.row, mEmailCol.EItin_mEmail, mEmailCol.EItin_mEmail
            setText msFlexEmails, msFlexEmails.row, mEmailCol.ETkt_mEmail, mEmailCol.ETkt_mEmail
            
            If .EItin Then
                msFlexEmails.TextMatrix(msFlexEmails.row, mEmailCol.EItin_mEmail) = gstrChecked
            ElseIf .ETkt Then
                msFlexEmails.TextMatrix(msFlexEmails.row, mEmailCol.ETkt_mEmail) = gstrChecked
            End If
             
            msFlexEmails.TextMatrix(msFlexEmails.row, mEmailCol.Type_mEmail) = .PaxType
            msFlexEmails.TextMatrix(msFlexEmails.row, mEmailCol.Email_mEmail) = .EmailAddress
            msFlexEmails.TextMatrix(msFlexEmails.row, mEmailCol.LineNum_mEmail) = .LineNum
            msFlexEmails.TextMatrix(msFlexEmails.row, mEmailCol.PNRLoc) = .PNRLoc
        End With
    Next
End Sub

Private Function GetEmailFromEmailGrid() As EmailAddresses
    Dim objEmails As New EmailAddresses
    Dim objEmail As New EmailAddress
    Dim strITI As String
    Dim strTKT As String
    Dim intI As Integer
    
    With msFlexEmails
        For intI = 1 To .rows - 1
            Set objEmail = New EmailAddress
            strITI = .TextMatrix(intI, mEmailCol.EItin_mEmail)
            strTKT = .TextMatrix(intI, mEmailCol.ETkt_mEmail)
            
            If strITI = gstrChecked Then objEmail.EItin = True
            If strTKT = gstrChecked Then objEmail.ETkt = True
            objEmail.PaxType = .TextMatrix(intI, mEmailCol.Type_mEmail)
            objEmail.EmailAddress = .TextMatrix(intI, mEmailCol.Email_mEmail)
            'objEmail.LineNum = .TextMatrix(intI, mEmailCol.LineNum_mEmail)
            'objEmail.PNRLoc = .TextMatrix(intI, mEmailCol.PNRLoc)
            
            objEmails.AddEmail objEmail
        Next
    End With
    
    Set GetEmailFromEmailGrid = objEmails
End Function

Private Function ValidEmail() As Boolean
    Dim intI As Integer
    Dim bolIti As Boolean
    Dim bolTkt As Boolean
    Dim strErr As String
    Dim objEmails As New EmailAddresses
    
    Set objEmails = GetEmailFromEmailGrid
    bolIti = False
    bolTkt = False
    For intI = 1 To objEmails.EmailCount
        With objEmails.Email(intI)
            If .PaxType = "" Then
                strErr = "Missing Pax Type.."
                Exit For
            End If
            If .EmailAddress = "" Then
                strErr = "Missing Email Address.."
                Exit For
            End If
            If .EItin Then
                bolIti = True
            End If
            If .ETkt Then
                bolTkt = True
            End If
        End With
    Next
    
    If strErr = "" Then
        If bolIti = True And bolTkt = True Then
            strErr = "Only one document type should be selected."
            strErr = strErr & vbCrLf & "Please check your selection"
        End If
        If bolIti = False And bolTkt = False Then
            strErr = "No selection has been made."
            strErr = strErr & vbCrLf & "Please check your selection"
        End If
    End If
    
    If strErr <> "" Then
        MsgBox strErr, , "CWT Desktop"
        ValidEmail = False
    Else
        ValidEmail = True
    End If
End Function

'CC - V1.2.20 20130612 - CR220 - Change to desktop logic for EM itinerary
'CC - Remark unused function
'Private Function PreLaunchValidate() As Boolean
'    Dim intI As Integer
'    Dim bolAQExist As Boolean
'    Dim bolHZExist As Boolean
'    Dim strErr As String
'
'    For intI = 1 To gobjPNR.GeneralRemarkCount
'        With gobjPNR.GeneralRemark(intI)
'            If .Qualifier = "" And UCase(Mid(.RemarkText, 1, 3)) = "AQ-" Then
'                bolAQExist = True
'            End If
'            If .Qualifier = "HZ" Then
'                bolHZExist = True
'            End If
'            If bolAQExist = True And bolHZExist = True Then
'                Exit For
'            End If
'        End With
'    Next
'
'    If bolAQExist = False Then
'        strErr = "Missing NP.AQ- lines and historical remarks."
'        strErr = strErr & vbCrLf & "Please document these remarks with Aqua Itin Rmk module."
'    ElseIf bolHZExist = True Then
'        strErr = "Your previous document has not been sent."
'        strErr = strErr & vbCrLf & "Please try to send this document later."
'    End If
'
'    If strErr <> "" Then
'        MsgBox strErr, , "CWT Desktop"
'        PreLaunchValidate = False
'    Else
'        PreLaunchValidate = True
'    End If
'End Function

Private Sub GetMappingTable()
   Dim strSQL As String
   Dim rs As ADODB.Recordset
   Dim item As ListItem
      
   lvwMappingTable.ListItems.Clear
   
   strSQL = "Select * from tblMapping "
   strSQL = strSQL & "Where MappingType = 'AquaItinEMail' "
   strSQL = strSQL & "Order By Sequence "
   
   Set rs = gdbAPPConn.Execute(strSQL)
   
   Do Until rs.EOF
        Set item = lvwMappingTable.ListItems.Add(, , rs!MapFrom & "")
        item.SubItems(1) = rs!MapTo & ""
        rs.MoveNext
   Loop
   
   rs.Close
   Set rs = Nothing

End Sub


Private Function ExactEmailToNPEmail(Email As String) As String
    Dim intI As Integer
    Dim strFrom As String
    Dim strTo As String
    Dim strEmail As String
    
    strEmail = Email
    For intI = 1 To lvwMappingTable.ListItems.Count
        strFrom = lvwMappingTable.ListItems(intI).Text
        strTo = lvwMappingTable.ListItems(intI).SubItems(1)
            
        strEmail = Replace(strEmail, strFrom, strTo)
    Next
    ExactEmailToNPEmail = strEmail
End Function

Private Function NPEmailToExactEmail(Email As String) As String
    Dim intI As Integer
    Dim strFrom As String
    Dim strTo As String
    Dim strEmail As String
    
    strEmail = Email
    For intI = lvwMappingTable.ListItems.Count To 1 Step -1
        strFrom = lvwMappingTable.ListItems(intI).SubItems(1)
        strTo = lvwMappingTable.ListItems(intI).Text
            
        strEmail = Replace(strEmail, strFrom, strTo)
    Next
    NPEmailToExactEmail = strEmail
End Function

'Get line number of RI.ITID, RI.ITIX, RI.ITIDX, RI.EMA
Private Function GetEDocRILine() As String
    Dim intI As Integer
    Dim strLineNum As String
    
    For intI = 1 To gobjPNR.ItinRemarkCount
        With gobjPNR.ItinRemark(intI)
'            If UCase(Mid(.RemarkText, 1, 4)) = "ITI." Or
'               UCase(Mid(.RemarkText, 1, 4)) = "TKT." Or
             If UCase(Mid(.RemarkText, 1, 4)) = "EMA." Or _
               UCase(Mid(.RemarkText, 1, 5)) = "ITID." Or _
               UCase(Mid(.RemarkText, 1, 5)) = "ITIX." Or _
               UCase(Mid(.RemarkText, 1, 6)) = "ITIDX." Then
               strLineNum = strLineNum & IIf(strLineNum = "", "", ".") & .ItemNum
            End If
        End With
    Next
    If strLineNum <> "" Then
        strLineNum = FormatedLineNum(strLineNum)
    End If
    GetEDocRILine = strLineNum
End Function

'Get line number of P.xxxE*
Private Function GetPELine() As String
    Dim intI As Integer
    Dim strPhoneType As String
    Dim strLineNum As String
    
                                '1234567890
    'gobjPNR.Phone(8).PhoneNum = SINE*PAX-CPIYADA//CARLSONWAGONLIT.COM
    For intI = 1 To gobjPNR.PhoneCount
        With gobjPNR.Phone(intI)
            strPhoneType = UCase(Mid(.PhoneNum, 4, 2))
            If strPhoneType = "E*" Then
                strLineNum = strLineNum & IIf(strLineNum = "", "", ".") & .ItemNum
            End If
        End With
    Next
    If strLineNum <> "" Then
        strLineNum = FormatedLineNum(strLineNum)
    End If
    GetPELine = strLineNum
End Function

Private Function UpdatePNR() As Boolean
    Dim objEmails As New EmailAddresses
    Dim strEDocRILine As String
    Dim strNPMLine As String
    Dim strPELine As String
    Dim colPECmd As New Collection
    Dim colAQNPCmd As New Collection
    Dim colRmkCmd As New Collection
    Dim strRes As String
    Dim strCmd As String
    Dim strMsg As String
    Dim strErrMsg As String
    Dim strFailCmd As String
    Dim intI As Integer
    
    'Load PNR again, in case Aqua updated NP line
    Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.loadPNR
    
    Set objEmails = GetEmailFromEmailGrid
    
    strEDocRILine = GetEDocRILine
    strPELine = GetPELine
    strNPMLine = GetLineNumNPM
    Set colPECmd = GetPECommand(objEmails)
    Set colAQNPCmd = GetAQNPCommand(objEmails)
'    Set colRmkCmd = GetRemarkCommand
    
    'Delete existing RI.ITID, RI.ITIX, RI.ITIDX, RI.EMA
    If strEDocRILine <> "" Then
        strCmd = "RI." & strEDocRILine & "@"
        strRes = gobjHost.terminalEntry(strCmd)
        If strRes <> "*" Then
           strMsg = "Failed to remove RI Remark."
           strMsg = strMsg & vbCrLf & "Desktop will Ignore & Retrieve (IR) the PNR" & vbCrLf
           strMsg = strMsg & vbCrLf & "Galileo Command: " & strCmd
           strMsg = strMsg & vbCrLf & "Galileo Response: " & strRes
           MsgBox strMsg, , "CWT Desktop"
           
           gobjHost.terminalEntry "IR"
           UpdatePNR = False
           Exit Function
        End If
    End If
    
    'Delete existing P.xxxE*xxxxxx
    If strPELine <> "" Then
        strCmd = "P." & strPELine & "@"
        strRes = gobjHost.terminalEntry(strCmd)
        If strRes <> "*" Then
           strMsg = "Failed to remove Phone Field."
           strMsg = strMsg & vbCrLf & "Desktop will Ignore & Retrieve (IR) the PNR" & vbCrLf
           strMsg = strMsg & vbCrLf & "Galileo Command: " & strCmd
           strMsg = strMsg & vbCrLf & "Galileo Response: " & strRes
           MsgBox strMsg, , "CWT Desktop"
           
           gobjHost.terminalEntry "IR"
           UpdatePNR = False
           Exit Function
        End If
    End If
    
    'Delete existing NP.M*MAIL-xxxxxxx
    If strNPMLine <> "" Then
        strCmd = "NP." & strNPMLine & "@"
        strRes = gobjHost.terminalEntry(strCmd)
        If strRes <> "*" Then
           strMsg = "Failed to remove NP.M*MAIL- line"
           strMsg = strMsg & vbCrLf & "Desktop will Ignore & Retrieve (IR) the PNR" & vbCrLf
           strMsg = strMsg & vbCrLf & "Galileo Command: " & strCmd
           strMsg = strMsg & vbCrLf & "Galileo Response: " & strRes
           MsgBox strMsg, , "CWT Desktop"
           
           gobjHost.terminalEntry "IR"
           UpdatePNR = False
           Exit Function
        End If
    End If
    
    'Insert P.xxxE*[email address]
    strErrMsg = SendGDSCmd(colPECmd, PE, strFailCmd)
    If strErrMsg <> "" Then
        strMsg = "Failed to Insert Phone Field."
        strMsg = strMsg & vbCrLf & "Desktop will Ignore & Retrieve (IR) the PNR" & vbCrLf
        strMsg = strMsg & vbCrLf & "Galileo Command: " & vbCrLf & strFailCmd
        strMsg = strMsg & vbCrLf & vbCrLf & "Galileo Response: " & vbCrLf & strErrMsg
        MsgBox strMsg, , "CWT Desktop"
        
        gobjHost.terminalEntry "IR"
        UpdatePNR = False
        Exit Function
    End If
    
    'Insert NP.M*MAIL-[email address]//TVLR
    'Insert NP.HZ*CONF*SEND ITIN
    strErrMsg = SendGDSCmd(colAQNPCmd, NP, strFailCmd)
    If strErrMsg <> "" Then
        strMsg = "Failed to Insert NP Line."
        strMsg = strMsg & vbCrLf & "Desktop will Ignore & Retrieve (IR) the PNR" & vbCrLf
        strMsg = strMsg & vbCrLf & "Galileo Command: " & vbCrLf & strFailCmd
        strMsg = strMsg & vbCrLf & vbCrLf & "Galileo Response: " & vbCrLf & strErrMsg
        MsgBox strMsg, , "CWT Desktop"
        
        gobjHost.terminalEntry "IR"
        UpdatePNR = False
        Exit Function
    End If
     
    
'    strErrMsg = SendGDSCmd(colRmkCmd, NP, strFailCmd)
'    If strErrMsg <> "" Then
'        strMsg = "Failed to Insert NP Line (Remarks)."
'        strMsg = strMsg & vbCrLf & "Desktop will Ignore & Retrieve (IR) the PNR" & vbCrLf
'        strMsg = strMsg & vbCrLf & "Galileo Command: " & vbCrLf & strFailCmd
'        strMsg = strMsg & vbCrLf & vbCrLf & "Galileo Response: " & vbCrLf & strErrMsg
'
'        MsgBox strMsg, , "CWT Desktop"
'        UpdatePNR = False
'        Exit Function
'    End If
    
    If ENDPNR(True) = True Then
        UpdatePNR = True
    Else
        UpdatePNR = False
    End If
End Function

Private Function GetPECommand(objEmails As EmailAddresses) As Collection
    Dim strCmd As String
    Dim colCmd As New Collection
    Dim intI As Integer
    Dim strDocType As String
    
    For intI = 1 To objEmails.EmailCount
        With objEmails.Email(intI)
             If .EItin = True Then
                strDocType = " " & "ITI"
             ElseIf .ETkt = True Then
                strDocType = " " & "TKT"
             Else
                strDocType = ""
             End If
             strCmd = "P." & gstrAgcyCityCode & "E*" & UCase(.PaxType) & "-" & convertPhoneText(.EmailAddress) & strDocType
             colCmd.Add strCmd
        End With
    Next
    
    Set GetPECommand = colCmd
End Function

Private Function GetAQNPCommand(objEmails As EmailAddresses) As Collection
    Dim strCmd As String
    Dim colCmd As New Collection
    Dim intI As Integer
    Dim strDocType As String
    
    For intI = 1 To objEmails.EmailCount
        With objEmails.Email(intI)
            If .EmailAddress <> "" And (.EItin = True Or .ETkt = True) Then
                If .EItin = True Then
                   strDocType = "ITIN"
                ElseIf .ETkt = True Then
                   strDocType = "ETIX"
                End If
                strCmd = "NP.M*MAIL-" & ExactEmailToNPEmail(.EmailAddress) & "//TVLR"
                colCmd.Add strCmd
            End If
        End With
    Next
    
    If strDocType <> "" Then
        'CC - V1.2.16 20121018 - CR163 - EM trigger generation removal in QUEUE module
        If mstrActivateFrom <> "QueueScreen" Then
            strCmd = "NP.HZ*CONF*SEND " & strDocType
            colCmd.Add strCmd
        End If
    End If
    
    Set GetAQNPCommand = colCmd
End Function

'Private Function GetRemarkCommand() As Collection
'    Dim lngC As Long
'    Dim colCmd As New Collection
'    Dim strNow As String
'
'    strNow = Format(Now, "ddmmmhh:ss")
'    For lngC = 0 To txtAdRI.Count - 1
'        If Trim(txtAdRI(lngC)) <> "" Then
'            colCmd.Add ("NP.HI*" & "I" & "." & strNow & "." & txtAdRI(lngC))
'        End If
'    Next
'
'    Set GetRemarkCommand = New Collection
'    Set GetRemarkCommand = colCmd
'End Function

'Get line number of NP.M*MAIL-
Private Function GetLineNumNPM() As String
    Dim intI As Integer
    Dim strLineNum As String
    
    For intI = 1 To gobjPNR.GeneralRemarkCount
        With gobjPNR.GeneralRemark(intI)
            If UCase(.Qualifier) = "*M" And UCase(Mid(.RemarkText, 1, 5)) = "MAIL-" Then
                strLineNum = strLineNum & IIf(strLineNum = "", "", ".") & .ItemNum
            End If
        End With
    Next
    If strLineNum <> "" Then
        strLineNum = FormatedLineNum(strLineNum)
    End If
    GetLineNumNPM = strLineNum
End Function
