VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmAdvancedSearch 
   BackColor       =   &H00FAF6EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " CWT Desktop - Advanced Search"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdvancedSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6450
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   3900
      Left            =   0
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   6879
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
      Begin MyFramePanel.MyFrame MyFrame2 
         Height          =   2865
         Left            =   120
         Top             =   100
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5054
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
         HeaderGradientAlign=   5
         HeaderGradientSizeH=   "50%"
         HeaderColorTopLeft=   6973442
         HeaderColorTopRight=   6973442
         HeaderColorBottomLeft=   6973442
         HeaderColorBottomRight=   6973442
         HeaderShow      =   0   'False
         PictureOffsetX  =   5
         Begin VB.OptionButton optSearchBy 
            BackColor       =   &H00DADAB6&
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton optSearchBy 
            BackColor       =   &H00DADAB6&
            Caption         =   "Rec Locator: *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkFlight 
            BackColor       =   &H00DADAB6&
            Caption         =   "Include Flight Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   390
            TabIndex        =   6
            Top             =   1560
            Width           =   3135
         End
         Begin VB.CheckBox chkAllBranches 
            BackColor       =   &H00DADAB6&
            Caption         =   "All Branches"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3885
            TabIndex        =   5
            Top             =   1250
            Width           =   1575
         End
         Begin VB.TextBox txtPCC 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1485
            TabIndex        =   4
            Top             =   1250
            Width           =   2025
         End
         Begin VB.TextBox txtFlightNum 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4035
            TabIndex        =   3
            Top             =   1920
            Width           =   1545
         End
         Begin VB.TextBox txtFirstName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1515
            TabIndex        =   2
            Top             =   915
            Width           =   2025
         End
         Begin VB.TextBox txtLastName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1515
            TabIndex        =   1
            Top             =   585
            Width           =   2025
         End
         Begin VB.TextBox txtLocator 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1515
            TabIndex        =   0
            Top             =   250
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker dtpDepDate 
            Height          =   375
            Left            =   4035
            TabIndex        =   9
            Top             =   2280
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483647
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   64880641
            CurrentDate     =   38285
         End
         Begin VB.Label lblPCC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pseudo: *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   390
            TabIndex        =   17
            Top             =   1245
            Width           =   690
         End
         Begin VB.Label lblBoardPoint 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Board Point:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   390
            TabIndex        =   16
            Top             =   2345
            Width           =   870
         End
         Begin VB.Label lblFlightNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Flight Number: *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   15
            Top             =   1970
            Width           =   1125
         End
         Begin VB.Label lblAirline 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Airline: *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   390
            TabIndex        =   14
            Top             =   1970
            Width           =   570
         End
         Begin VB.Label lblDepDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Depart Date: *"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   13
            Top             =   2345
            Width           =   1020
         End
         Begin VB.Label lblFirstName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   390
            TabIndex        =   12
            Top             =   960
            Width           =   795
         End
         Begin MSForms.ComboBox cmbAirline 
            Height          =   315
            Left            =   1440
            TabIndex        =   11
            Top             =   1920
            Width           =   1095
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1931;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox cmbBoardPoint 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   2280
            Width           =   1095
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1931;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   34
         End
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   360
         Left            =   5280
         TabIndex        =   18
         Top             =   3050
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
         MouseIcon       =   "frmAdvancedSearch.frx":038A
         MousePointer    =   99
         Picture         =   "frmAdvancedSearch.frx":06A4
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   16447215
         Caption         =   "&Cancel"
         Depth           =   1
         PictureDisabled =   "frmAdvancedSearch.frx":09CF
         PictureOffsetX  =   2
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdSearch 
         Height          =   360
         Left            =   4200
         TabIndex        =   19
         Top             =   3050
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
         MouseIcon       =   "frmAdvancedSearch.frx":0D21
         MousePointer    =   99
         Picture         =   "frmAdvancedSearch.frx":103B
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   16447215
         Caption         =   "&Search"
         Depth           =   1
         PictureDisabled =   "frmAdvancedSearch.frx":1344
         PictureOffsetX  =   2
         GradientType    =   2
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
         LcK2            =   $"frmAdvancedSearch.frx":1696
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
   End
End
Attribute VB_Name = "frmAdvancedSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datTouchEnd As Date

Private Sub chkAllBranches_Click()
    If chkAllBranches.value = 1 Then
       txtPCC.Enabled = False
    Else
       txtPCC.Enabled = True
    End If
End Sub

Private Sub chkFlight_Click()
    If chkFlight.value = 1 Then
       'PCC details must be disabled
       txtPCC.Enabled = False
       chkAllBranches.Enabled = False
       cmbAirline.Enabled = True
       txtFlightNum.Enabled = True
       cmbBoardPoint.Enabled = True
       dtpDepDate.Enabled = True
    Else
       txtPCC.Enabled = True
       chkAllBranches.Enabled = True
       cmbAirline.Enabled = False
       txtFlightNum.Enabled = False
       cmbBoardPoint.Enabled = False
       dtpDepDate.Enabled = False
    End If
End Sub

Private Sub cmbAirline_GotFocus()
    cmbGetFocus cmbAirline
End Sub

Private Sub cmbAirline_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 13 Then
      cmdSearch_Click
   Else
      KeyAscii = fAllowAlphaNumeric(CInt(KeyAscii))
   End If
End Sub

Private Sub cmbBoardPoint_GotFocus()
    cmbGetFocus cmbBoardPoint
End Sub

Private Sub cmbBoardPoint_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 13 Then
      cmdSearch_Click
   Else
      KeyAscii = fAllowAlphaNumeric(CInt(KeyAscii))
   End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim strResponse As String
    Dim strName As String
    datTouchEnd = Now
    
    If validData Then
       If optSearchBy(0).value = True Then
       Set gobjPNR = New CWT_GalileoPNR3.PNR
          gobjPNR.loadPNR Trim(txtLocator)
          displayPNRinBar
          Unload Me
       Else
          strName = Trim(txtLastName.Text) & IIf(Trim(txtFirstName.Text) <> "", "/" & Trim(txtFirstName.Text), "")
          If chkAllBranches.value = 0 And chkFlight.value = 0 Then
          Set gobjPNR = New CWT_GalileoPNR3.PNR
             strResponse = gobjPNR.loadPNR("", strName, Trim(txtPCC))
          ElseIf chkAllBranches.value = 1 And chkFlight.value = 0 Then
          Set gobjPNR = New CWT_GalileoPNR3.PNR
             strResponse = gobjPNR.loadPNR("", strName, "", True)
          ElseIf chkFlight.value = 1 Then
          Set gobjPNR = New CWT_GalileoPNR3.PNR
             strResponse = gobjPNR.loadPNR("", strName, "", False, Trim(cmbAirline.Text), Trim(txtFlightNum.Text), Format(dtpDepDate.value, "yyyymmdd"), Trim(cmbBoardPoint.Text))
          End If
          
          If strResponse = "SNL" Then
             Unload Me
             loadSimiliarNames
          ElseIf strResponse = "PRO" Then
             Unload Me
             Set gobjPNR = New CWT_GalileoPNR3.PNR
             gobjPNR.loadPNR
             displayPNRinBar
          End If
       End If
    End If
    
   gstrProcessGrpID = pGetProcessKey
   pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, gconModProfile, frmSideBar.cmbSelectType.Text, _
      gconSModSearch, Me.Name, "SEARCHPNR", gstrProcessGrpID, _
      , datTouchEnd
    
End Sub

Private Sub dtpDepDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdSearch_Click
   Else
      KeyAscii = fAllowAlphaNumeric(KeyAscii)
   End If
End Sub

Private Sub Form_Load()
   Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
   optSearchBy(0).value = True
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   
   loadAirline
   loadBoardPoint
   txtPCC.Text = gobjHost.AgentPCC
   cmbAirline.Enabled = False
   cmbBoardPoint.Enabled = False
End Sub

Private Sub optSearchBy_Click(Index As Integer)
    If Index = 0 Then
       txtLocator.Visible = True
       lblFirstName.Visible = False
       txtFirstName.Visible = False
       optSearchBy(1).Caption = "Name:"
       txtLastName.Visible = False
       lblAirline.Visible = False
       cmbAirline.Visible = False
       lblFlightNum.Visible = False
       txtFlightNum.Visible = False
       lblBoardPoint.Visible = False
       cmbBoardPoint.Visible = False
       lblDepDate.Visible = False
       dtpDepDate.Visible = False
       lblPCC.Visible = False
       txtPCC.Visible = False
       chkAllBranches.Visible = False
       chkFlight.Visible = False
    ElseIf Index = 1 Then
       txtLocator.Visible = False
       txtFirstName.Visible = True
       optSearchBy(1).Caption = "Last Name: *"
       txtLastName.Visible = True
       lblFirstName.Visible = True
       lblAirline.Visible = True
       cmbAirline.Visible = True
       lblFlightNum.Visible = True
       txtFlightNum.Visible = True
       lblBoardPoint.Visible = True
       cmbBoardPoint.Visible = True
       lblDepDate.Visible = True
       dtpDepDate.Visible = True
       lblPCC.Visible = True
       txtPCC.Visible = True
       chkAllBranches.Visible = True
       chkFlight.Visible = True
    End If
End Sub


Private Function validData() As Boolean
    
    Dim strMsg As String

    If optSearchBy(0).value = True Then
       If Trim(txtLocator) = "" Then strMsg = strMsg & "Missing Record Locator..." & Chr(13)
    Else
        If Trim(txtLastName) = "" Then strMsg = strMsg & "Missing Last Name..." & Chr(13)
        If Trim(txtPCC) = "" And chkAllBranches.value = 0 And chkFlight.value = 0 Then
           strMsg = strMsg & "Missing Pseudo City Code..." & Chr(13)
        Else
           If Len(Trim(txtPCC)) > 4 And chkAllBranches.value = 0 And chkFlight.value = 0 Then
              strMsg = strMsg & "Pseudo City Code must not more than 4 characters..." & Chr(13)
           End If
        End If
        If chkFlight.value = 1 Then
           If Trim(cmbAirline) = "" Then strMsg = strMsg & "Missing Airline..." & Chr(13)
           If Trim(txtFlightNum) = "" Then strMsg = strMsg & "Missing Flight Number..." & Chr(13)
           If IsNull(dtpDepDate.value) Then strMsg = strMsg & "Missing Departure Date..." & Chr(13)
        End If
    End If
    
    If strMsg = "" Then
        validData = True
    Else
        validData = False
        DataisRequired strMsg
    End If
    
End Function

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdSearch_Click
   Else
      KeyAscii = fAllowAlpha(KeyAscii, " ")
   End If
End Sub

Private Sub txtFlightNum_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdSearch_Click
   Else
      KeyAscii = fAllowAlphaNumeric(KeyAscii, " ")
   End If
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdSearch_Click
   Else
      KeyAscii = fAllowAlpha(KeyAscii, " ")
   End If
End Sub

Private Sub txtLocator_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdSearch_Click
   Else
      KeyAscii = fAllowAlphaNumeric(KeyAscii)
   End If
End Sub

Private Sub txtPCC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdSearch_Click
   Else
      KeyAscii = fAllowAlphaNumeric(KeyAscii)
   End If
End Sub

Private Sub loadAirline()
    
    Dim rsRecord As ADODB.Recordset
    Dim strSql As String

    strSql = "Select code,description from tblAirVendors Where Type = 'AIR' order by code"
    Set rsRecord = gdbConn.Execute(strSql)
    cmbAirline.Clear
    cmbAirline.ColumnCount = 2
    cmbAirline.ColumnWidths = "30,200"
    cmbAirline.ListWidth = 250
    While Not rsRecord.EOF
        cmbAirline.AddItem rsRecord!Code & ""
        cmbAirline.List(cmbAirline.ListCount - 1, 1) = rsRecord!Description & ""
        rsRecord.MoveNext
    Wend
    Set rsRecord = Nothing
    If cmbAirline.ListCount > 0 Then
       cmbAirline.listindex = 0
    End If
    
End Sub

Private Sub loadBoardPoint()
    
    Dim rsRecord As ADODB.Recordset
    Dim strSql As String

    strSql = "select AirportCode,Airport,City from tblCityCodes order by airportCode"
    Set rsRecord = gdbConn.Execute(strSql)
    cmbBoardPoint.Clear
    cmbBoardPoint.ColumnCount = 3
    cmbBoardPoint.ColumnWidths = "30,200,200"
    cmbBoardPoint.ListWidth = 450
    While Not rsRecord.EOF
        cmbBoardPoint.AddItem rsRecord!AirportCode & ""
        cmbBoardPoint.List(cmbBoardPoint.ListCount - 1, 1) = rsRecord!Airport & ""
        cmbBoardPoint.List(cmbBoardPoint.ListCount - 1, 2) = rsRecord!City & ""
        rsRecord.MoveNext
    Wend
    Set rsRecord = Nothing
    If cmbBoardPoint.ListCount > 0 Then
       cmbBoardPoint.listindex = 0
    End If
End Sub


