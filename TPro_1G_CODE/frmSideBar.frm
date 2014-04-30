VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmSideBar 
   BorderStyle     =   0  'None
   Caption         =   "CWT Desktop - SideBar"
   ClientHeight    =   9615
   ClientLeft      =   2415
   ClientTop       =   4080
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView treeViewTraveller 
      Height          =   4695
      Left            =   240
      TabIndex        =   30
      Top             =   3960
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   8281
      _Version        =   393217
      Indentation     =   441
      Style           =   7
      Appearance      =   1
   End
   Begin MyCommandButton.MyButton cmdExpand 
      Height          =   975
      Left            =   0
      TabIndex        =   29
      Top             =   3000
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   1720
      BackColor       =   15523541
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSideBar.frx":0000
      BackColorDown   =   15523541
      BackColorOver   =   15523541
      BackColorFocus  =   15523541
      BackColorDisabled=   15523541
      BorderColor     =   8540205
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmSideBar.frx":00A6
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton cmdReverse 
      Height          =   975
      Left            =   3625
      TabIndex        =   28
      Top             =   3000
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   1720
      BackColor       =   15523541
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmSideBar.frx":014C
      BackColorDown   =   15523541
      BackColorOver   =   15523541
      BackColorFocus  =   15523541
      BackColorDisabled=   15523541
      BorderColor     =   8540205
      TransparentColor=   14215660
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmSideBar.frx":01F2
      ShowFocus       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   4800
   End
   Begin VB.Frame frmProfile 
      BackColor       =   &H00DADAB6&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1680
      Left            =   4080
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   3420
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1125
         TabIndex        =   15
         Top             =   960
         Width           =   2200
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1125
         TabIndex        =   14
         Top             =   640
         Width           =   2200
      End
      Begin VB.OptionButton optTrvType 
         BackColor       =   &H00DADAB6&
         Caption         =   "Group"
         Height          =   255
         Index           =   2
         Left            =   2280
         MaskColor       =   &H0080C0FF&
         TabIndex        =   18
         Top             =   50
         Width           =   1095
      End
      Begin VB.OptionButton optTrvType 
         BackColor       =   &H00DADAB6&
         Caption         =   "Corporate"
         Height          =   255
         Index           =   1
         Left            =   1080
         MaskColor       =   &H0080C0FF&
         TabIndex        =   12
         Top             =   50
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optTrvType 
         BackColor       =   &H00DADAB6&
         Caption         =   "Leisure"
         Height          =   255
         Index           =   0
         Left            =   120
         MaskColor       =   &H0080C0FF&
         TabIndex        =   11
         Top             =   50
         Width           =   1815
      End
      Begin MyCommandButton.MyButton cmdReset 
         Height          =   360
         Index           =   1
         Left            =   2280
         TabIndex        =   17
         Top             =   1300
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
         MouseIcon       =   "frmSideBar.frx":0298
         MousePointer    =   99
         Picture         =   "frmSideBar.frx":05B2
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   14215660
         Caption         =   "&Reset"
         Depth           =   1
         PictureDisabled =   "frmSideBar.frx":08DC
         PictureOffsetX  =   2
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdSearch 
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   1300
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
         MouseIcon       =   "frmSideBar.frx":0B12
         MousePointer    =   99
         Picture         =   "frmSideBar.frx":0E2C
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   14215660
         Caption         =   "&Search"
         Depth           =   1
         PictureDisabled =   "frmSideBar.frx":1135
         PictureOffsetX  =   2
         GradientType    =   2
      End
      Begin MSForms.ComboBox cmbBar 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   280
         Width           =   3240
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "5715;556"
         ColumnCount     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   975
         Width           =   795
      End
   End
   Begin VB.Frame frmLocator 
      BackColor       =   &H00DADAB6&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1680
      Left            =   4080
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   3420
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Top             =   470
         Width           =   2025
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   7
         Top             =   800
         Width           =   2025
      End
      Begin MyCommandButton.MyButton cmdReset 
         Height          =   360
         Index           =   0
         Left            =   2280
         TabIndex        =   10
         Top             =   1200
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
         MouseIcon       =   "frmSideBar.frx":1487
         MousePointer    =   99
         Picture         =   "frmSideBar.frx":17A1
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   14215660
         Caption         =   "&Reset"
         Depth           =   1
         PictureDisabled =   "frmSideBar.frx":1ACB
         PictureOffsetX  =   2
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdSearch 
         Height          =   360
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Top             =   1200
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
         MouseIcon       =   "frmSideBar.frx":1D01
         MousePointer    =   99
         Picture         =   "frmSideBar.frx":201B
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   14215660
         Caption         =   "&Search"
         Depth           =   1
         PictureDisabled =   "frmSideBar.frx":2324
         PictureOffsetX  =   2
         GradientType    =   2
      End
      Begin VB.OptionButton optSearchBy 
         BackColor       =   &H00DADAB6&
         Caption         =   "Last Name: *"
         Height          =   255
         Index           =   1
         Left            =   50
         TabIndex        =   4
         Top             =   480
         Width           =   1450
      End
      Begin VB.OptionButton optSearchBy 
         BackColor       =   &H00DADAB6&
         Caption         =   "Rec Loc : *"
         Height          =   255
         Index           =   0
         Left            =   50
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin MSForms.ComboBox cmbLocator 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   2025
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "3572;556"
         ListWidth       =   8819
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   34
         Object.Width           =   "1762;7055"
      End
      Begin VB.Label lblAdvancedSearch 
         BackColor       =   &H00DADAB6&
         Caption         =   "&Advanced Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006A6802&
         Height          =   495
         Left            =   120
         MouseIcon       =   "frmSideBar.frx":2676
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblFirstName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   195
         Left            =   320
         TabIndex        =   22
         Top             =   840
         Width           =   795
      End
   End
   Begin VB.TextBox txtHidden 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4080
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   9570
      Left            =   0
      Top             =   0
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   16880
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
      Begin MyFramePanel.MyFrame fraInfo 
         Height          =   5475
         Left            =   120
         Top             =   3405
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   9657
         BackColor       =   14342838
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   14333622
         Picture         =   "frmSideBar.frx":2980
         BackgroundMaskColor=   13421772
         BackgroundAlignment=   4
         BodyShow        =   -1  'True
         BodyColorTopLeft=   14342838
         BodyColorTopRight=   14342838
         BodyColorBottomRight=   14342838
         BodyColorBottomLeft=   14342838
         BodyGradientSizeV=   "80%"
         BorderColor     =   6973442
         Caption         =   " Passenger Info"
         CornerTopLeft   =   -1  'True
         CornerTopRight  =   -1  'True
         CornerBottomLeft=   -1  'True
         CornerBottomRight=   -1  'True
         OutSideColor    =   14215660
         HeaderHeight    =   27
         HeaderGradientAlign=   5
         HeaderColorTopLeft=   6973442
         HeaderColorTopRight=   6973442
         HeaderColorBottomLeft=   6973442
         HeaderColorBottomRight=   6973442
         PictureOffsetX  =   5
      End
      Begin MyFramePanel.MyFrame fraSearch 
         Height          =   3225
         Left            =   120
         Top             =   80
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   5689
         BackColor       =   14342838
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   14333622
         Picture         =   "frmSideBar.frx":2D46
         BackgroundMaskColor=   13421772
         BackgroundAlignment=   4
         BodyShow        =   -1  'True
         BodyColorTopLeft=   14342838
         BodyColorTopRight=   14342838
         BodyColorBottomRight=   14342838
         BodyColorBottomLeft=   14342838
         BodyGradientSizeV=   "80%"
         BorderColor     =   6973442
         Caption         =   " Search"
         CornerTopLeft   =   -1  'True
         CornerTopRight  =   -1  'True
         CornerBottomLeft=   -1  'True
         CornerBottomRight=   -1  'True
         OutSideColor    =   14215660
         HeaderGradientAlign=   5
         HeaderColorTopLeft=   6973442
         HeaderColorTopRight=   6973442
         HeaderColorBottomLeft=   6973442
         HeaderColorBottomRight=   6973442
         PictureOffsetX  =   5
         Begin VB.ComboBox cmbSelectReq 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmSideBar.frx":3188
            Left            =   1320
            List            =   "frmSideBar.frx":318A
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1170
            Width           =   2020
         End
         Begin VB.TextBox txtRequestor 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   1
            Top             =   840
            Width           =   2025
         End
         Begin VB.ComboBox cmbSelectType 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmSideBar.frx":318C
            Left            =   1320
            List            =   "frmSideBar.frx":3196
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   480
            Width           =   2020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Request Type: *"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1230
            Width           =   1155
         End
         Begin VB.Label lblRequestor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Requestor: *"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   900
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Type:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   540
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmSideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim datTouchEnd As Date

Private Sub cmbBar_GotFocus()
    cmbGetFocus cmbBar
End Sub

Private Sub cmbBar_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 13 Then
      cmdSearch_Click (1)
   Else
      KeyAscii = fAllowAlphaNumeric(CInt(KeyAscii), " ")
   End If
End Sub

Private Sub cmbLocator_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   If KeyCode = 13 Then
      datTouchEnd = Now
      retrievePNR
   End If
End Sub

Private Sub cmbLocator_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = fAllowAlphaNumeric(CInt(KeyAscii))
End Sub

Private Sub cmbSelectReq_Click()

' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    If gIntModuleType = gModuleType.SYEX Then
         ' ZhiSam - V1.3.2 - 20121019 - BugFix - Always Enable the Tab wehre do not disableAllSubTab for HKSG SyEx Module
         'do nothing for enableTab
    Else
        enableTab
    End If

End Sub

Private Sub cmbSelectType_Click()
Dim i As Integer

    'Load respective panel
    If cmbSelectType.Text = "Profile" Then
       reloadPanel frmProfile
       
       'default to new booking
       For i = 0 To cmbSelectReq.ListCount - 1
        If cmbSelectReq.ItemData(i) = 6 Then
            cmbSelectReq.listindex = i
            lblRequestor.Caption = "Requestor:*"
            Exit For
        End If
       Next
       
       
    ElseIf cmbSelectType.Text = "PNR" Then
       reloadPanel frmLocator
       optSearchBy(0).value = True
        'default to amendment
       For i = 0 To cmbSelectReq.ListCount - 1
        If cmbSelectReq.ItemData(i) = 7 Then
            cmbSelectReq.listindex = i
            lblRequestor.Caption = "Requestor:"
            Exit For
        End If
       Next
      
    End If

' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    If gIntModuleType = gModuleType.SYEX Then
         ' ZhiSam - V1.3.2 - 20121019 - BugFix - Always Enable the Tab wehre do not disableAllSubTab for HKSG SyEx Module
         'do nothing for enableTab
    Else
        enableTab
    End If
    
End Sub

'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
'Change to Public sub
Public Sub cmdExpand_Click()
   Me.Width = gSideBarWidth
   Me.Move 0, 0
   cmdExpand.Visible = False
   cmdReverse.Visible = True
   resizePCWindow frmSideBar.cmdReverse.Visible, frmCustomVP.cmdContract.Visible
End Sub

Private Sub cmdReset_Click(Index As Integer)
    If Index = 0 Then
       cmbLocator.Text = ""
    End If
    txtRequestor.Text = ""
    txtLastName(Index).Text = ""
    txtFirstName(Index).Text = ""
End Sub

Public Sub searchProfile()
    Dim strMsg As String
    
    If checkSignOn = False Then
       Exit Sub
    End If
    
    If PNRExistsInGDS() = True Then
       Exit Sub
    End If
    
    If treeViewTraveller.Nodes.Count > 0 Then treeViewTraveller.Nodes.Clear
    If cmbSelectType.Text = "Profile" And cmbSelectReq.Text = "New Booking - 6" Then
        
        If Trim(txtRequestor.Text) = "" Then
           strMsg = "Missing requestor name ..." & Chr(13)
        End If
    End If
    If Trim(cmbSelectReq.Text) = "" Then
       strMsg = strMsg & "Missing requestor type ..." & Chr(13)
    End If
    'Preethi - V1.2.4 20110527 - CR68 - Default BAR Drop List to Please Select One
    If Trim(cmbBar.Text) = "" Or Trim(cmbBar.Text) = "Please Select One" Then
       strMsg = strMsg & "Must select a company profile ..." & Chr(13)
    End If
    If strMsg <> "" Then
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Sub
    End If
    
    gbolMoveProfile = False
    gbolCancelMove = False
    moveProfile True

End Sub
'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
'Change to Public sub
Public Sub cmdReverse_Click()
   'JY - 20100604 - Change width to pixel. default is in twips
   Me.Width = 9 * Screen.TwipsPerPixelX
   Me.Move 0, 0
   cmdExpand.Visible = True
   cmdReverse.Visible = False
       resizePCWindow frmSideBar.cmdReverse.Visible, frmCustomVP.cmdContract.Visible
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    If Index = 0 Then
       datTouchEnd = Now
       retrievePNR  'Retrieve PNR
    ElseIf Index = 1 Then
       datTouchEnd = Now
       searchProfile 'Search Profile
    End If
End Sub

Private Sub Form_Load()
   
   Dim old_parent As Long
   
   old_parent = SetParent(frmLocator.hwnd, fraSearch.hwnd)
   old_parent = SetParent(frmProfile.hwnd, fraSearch.hwnd)
   
   cmdExpand.Visible = False
   cmdReverse.Visible = True
   
   frmSideBar.Width = gSideBarWidth
   
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Adjust size for the side bar
    If gIntModuleType <> gModuleType.PC Then
        ' if it is Smart Point Window
        frmSideBar.Height = gFPHeight - gAssistedPanelHeight
        ARGradient1.Height = gFPHeight - gAssistedPanelHeight
        fraInfo.Height = gFPHeight - gAssistedPanelHeight - fraInfo.Top - gPadding
        
    Else
        ' if it is Point & Click Window
        frmSideBar.Height = gFPHeight - frmCustomVP.Height
        ARGradient1.Height = gFPHeight - frmCustomVP.Height
        fraInfo.Height = gFPHeight - frmCustomVP.Height - fraInfo.Top - gPadding
    End If
    
   
   treeViewTraveller.Height = fraInfo.Height - (treeViewTraveller.Top - fraInfo.Top) - gPadding
   
   optTrvType(1).value = True
   optTrvType_Click (1)
   cmbSelectType.listindex = 0
   pConnectToHost txtHidden
   loadRequestType
   
End Sub

Private Sub retrievePNR()
    Dim lngC As Long
    Dim item As ListItem
    Dim oldParent As Long
    Dim strMsg As String
    Dim strTemp As String
    
    On Error GoTo test
    
    If checkSignOn = False Then
       Exit Sub
    End If

    If PNRExistsInGDS() = True Then
       Exit Sub
    End If
    
    If treeViewTraveller.Nodes.Count > 0 Then treeViewTraveller.Nodes.Clear
    If cmbSelectType.Text = "Profile" And cmbSelectReq.Text = "New Booking - 6" Then
        If Trim(txtRequestor.Text) = "" Then
           strMsg = "Missing requestor name ..." & Chr(13)
        End If
    End If
    If Trim(cmbSelectReq.Text) = "" Then
       strMsg = strMsg & "Missing requestor type ..." & Chr(13)
    End If
    If optSearchBy(0).value = True And Trim(cmbLocator.Text) = "" Then
       strMsg = strMsg & "Missing record locator ..." & Chr(13)
    ElseIf optSearchBy(1).value = True And Trim(txtLastName(0).Text) = "" Then
       strMsg = strMsg & "Missing last name ..." & Chr(13)
    End If
  
    If strMsg <> "" Then
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Sub
    End If
  
    If optSearchBy(0).value = True Then
    Set gobjPNR = New CWT_GalileoPNR3.PNR
       strTemp = gobjPNR.loadPNR(Trim(cmbLocator))
       If strTemp = "PRO" Then
          displayPNRinBar
          frmBars.searchPanel
          'Preethi - V1.2.4 20110621  - ER01 - Tracking of Touches to CWT Booking
          If gobjPNR.CompInfo.PNRTrackingTouches = True Then
             Load frmPNRTouchTracking
             frmPNRTouchTracking.Show
          End If
       End If
    
    ElseIf optSearchBy(1).value = True Then
       Set gobjPNR = New CWT_GalileoPNR3.PNR
       strTemp = gobjPNR.loadPNR("", Trim(txtLastName(0).Text) & IIf(Trim(txtFirstName(0).Text) <> "", "/" & Trim(txtFirstName(0).Text), ""))
    
       If strTemp = "SNL" Then
          loadSimiliarNames
     
       ElseIf strTemp = "PRO" Then
           displayPNRinBar
           frmBars.searchPanel
           'Preethi - V1.2.4 20110621  - ER01 - Tracking of Touches to CWT Booking
           If gobjPNR.CompInfo.PNRTrackingTouches = True Then
              Load frmPNRTouchTracking
              frmPNRTouchTracking.Show
          End If
       End If
    End If

   gstrProcessGrpID = pGetProcessKey
   pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, gconModProfile, frmSideBar.cmbSelectType.Text, _
      gconSModSearch, Me.Name, "SEARCHPNR", gstrProcessGrpID, _
      , datTouchEnd

Exit Sub

test:
MsgBox Err.Description & Err.Number

End Sub

Private Sub lblAdvancedSearch_Click()
   Dim strMsg As String
   
   If checkSignOn = False Then
       Exit Sub
   End If
    
   If PNRExistsInGDS() = True Then
       Exit Sub
   End If
   If cmbSelectType.Text = "Profile" And cmbSelectReq.Text = "New Booking - 6" Then

    If Trim(txtRequestor.Text) = "" Then
        strMsg = "Missing requestor name ..." & Chr(13)
    End If
   
   End If
   If strMsg <> "" Then
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Sub
   End If
    
   If isLoaded("frmAdvancedSearch") = False Then
       frmAdvancedSearch.Show
   End If
End Sub

Private Sub optSearchBy_Click(Index As Integer)
    'Search PNR either by record locator or name
    cmbLocator.Text = ""
    txtFirstName(0).Text = ""
    txtLastName(0).Text = ""
    If Index = 0 Then
       cmbLocator.Visible = True
       lblFirstName.Visible = False
       txtFirstName(0).Visible = False
       txtLastName(0).Visible = False
       optSearchBy(1).Caption = "Name:"
    ElseIf Index = 1 Then
       cmbLocator.Visible = False
       lblFirstName.Visible = True
       txtFirstName(0).Visible = True
       txtLastName(0).Visible = True
       optSearchBy(1).Caption = "Last Name: *"
    End If
End Sub

Private Sub optTrvType_Click(Index As Integer)
  Select Case Index
    Case 0
      gTrxnType = "L"
      loadProName "L"
    Case 1
      gTrxnType = "B"
      loadProName "B"
    Case 2
      gTrxnType = "M"
      loadProName "M"
  End Select
End Sub

Private Sub reloadPanel(ByRef frmPanel As Frame)
    'Unload all panels and show the selected panel
    frmProfile.Visible = False
    frmLocator.Visible = False
    frmPanel.Top = 1480
    frmPanel.Left = 10
    frmPanel.Visible = True
End Sub


Private Sub timer1_Timer()
Start:
    On Error Resume Next
    If loadedMinForms = True Then
        With txtHidden
            .LinkItem = "CaptureAll"
            .LinkRequest
            If Err.Number > 0 Then GoTo ErrProc
            If IsAlphaNumeric(Mid(.Text, 1, 6)) = True And Mid(.Text, 7, 1) = "/" Then
               If gobjPNR Is Nothing = False Then
                  If gobjPNR.RecLoc <> Mid(.Text, 1, 6) Or InStr(1, gstrPreviousText, "IGNORED") > 0 Then
                     If InStr(1, .Text, "IGNORED") = 0 Then
                     
                        gobjPNR.loadPNR
                        displayPNRinBar
                        gstrPreviousText = ""
                        Exit Sub
                     End If
                  End If
               End If
            End If
            If IsAlphaNumeric(Mid(.Text, 67, 6)) = True And Mid(.Text, 73, 1) = "/" Then
               If gobjPNR Is Nothing = False Then
                  If gobjPNR.RecLoc <> Mid(.Text, 67, 6) Then
                     gobjPNR.loadPNR
                     displayPNRinBar
                     Exit Sub
                  End If
               End If
            End If
            If (InStr(1, gstrPreviousText, "»ER") > 0 Or InStr(1, gstrPreviousText, "+ER") > 0 Or InStr(1, gstrPreviousText, "»IR                                       ") > 0) And _
               ((IsAlphaNumeric(Mid(.Text, 1, 6)) = True And Mid(.Text, 7, 1) = "/") Or (IsAlphaNumeric(Mid(.Text, 67, 6)) = True And Mid(.Text, 73, 1) = "/")) Then
               If gobjPNR Is Nothing = False Then
                  gobjPNR.loadPNR
                  displayPNRinBar
                  gstrPreviousText = ""
               End If
            End If
            If InStr(1, .Text, "IGNORED") > 0 Or InStr(1, .Text, "EOK -") > 0 Then
               If treeViewTraveller.Nodes.Count > 0 Then
                  treeViewTraveller.Nodes.Clear
                  fraInfo.Caption = " No PNR"
               End If
            End If
            gstrPreviousText = .Text
        End With
    End If
    Exit Sub
ErrProc:
    Select Case Err.Number
        Case 293
            Call pConnectToHost(txtHidden)
            With frmDDEOwner.txtDDE
                 .LinkItem = "ClearWindow"
                 .Text = "1"
                 .LinkPoke
            End With
        Case 285 'Foreign application won't perform DDE method or operation
             AppActivate "Galileo Desktop"
             PressKey "R", , True
             GoTo Start
        Case Else
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, "Error " & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error Connect to GDS"
    End Select
End Sub

Private Sub txtFirstName_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index = 0 Then
        datTouchEnd = Now
        retrievePNR
      ElseIf Index = 1 Then
        datTouchEnd = Now
        searchProfile
      End If
   Else
      KeyAscii = fAllowAlpha(KeyAscii, " ")
   End If
End Sub

Private Sub txtLastName_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index = 0 Then
        datTouchEnd = Now
        retrievePNR
      ElseIf Index = 1 Then
        datTouchEnd = Now
        searchProfile
      End If
   Else
      KeyAscii = fAllowAlphaNumeric(KeyAscii, " ")
   End If
End Sub

Private Sub loadProName(strType As String)

    Dim rsProfile As ADODB.Recordset
    Dim strSql As String

    'Load Profile Name from database
    strSql = "SELECT ProName,CN FROM tblClients WHERE ProName <>'' AND Category='" & strType & "' ORDER BY [ProName]"
    Set rsProfile = gdbConn.Execute(strSql)
    cmbBar.Clear
    cmbBar.ColumnCount = 2
    cmbBar.ColumnWidths = "200,100"
    cmbBar.ListWidth = 300
    'Preethi - V1.2.4 20110527 - CR68 - Default BAR Drop List to Please Select One
    cmbBar.AddItem "Please Select One"
    While Not rsProfile.EOF
        cmbBar.AddItem rsProfile!ProName & ""
        cmbBar.List(cmbBar.ListCount - 1, 1) = rsProfile!CN & ""
        rsProfile.MoveNext
    Wend
    Set rsProfile = Nothing
    If cmbBar.ListCount > 0 Then
       cmbBar.listindex = 0
    End If
End Sub

Public Function moveProfile(ByVal bolSearch As Boolean) As Boolean
   Dim bolFound As Boolean
   Dim bolAdHoc As Boolean
   
   'Preethi - V1.2.12 20120528 - CR153 - PNR Touch Tracking - Auto Generate
   Dim strPrimaryCode As String
   Dim strSecondaryCode As String
   Dim strChargeIndicator As String
   Dim strContactCode As String
   
   ' This is a search profile function if bolSearch is true, else is move profile function
   
   pSetGlobals gobjHost.AgentDIV    '20090202
   
   'Ji Yong - Hold back this request due to the the concerns raised by BA
   'Preethi - V1.2.2 20110223 - CR37 - Removal Of Retention Line(Fastrak project-AQUA)
   'Preethi - V1.2.13 20120703 - CR167 - Change the PNR's Retention Period
   'gobjHost.terminalEntry "RT.T/" & Format(DateAdd("M", 6, Date), "ddmmm") & "*RETENTION LINE"
   gobjHost.terminalEntry "RT.T/" & Format(DateAdd("M", 3, Date), "ddmmm") & "*RETENTION LINE"
   
   gobjHost.terminalEntry "P." & gstrAgcyCityCode & "T*CARLSON WAGONLIT TRAVEL - " & gstrAgcyPhone _
                          & " - " & gobjHost.AgentName & "-Q" & gobjHost.AgentQueue & "+DI.AR-" & gobjHost.AgentName
   bolAdHoc = True
   If Trim(txtLastName(1)) = "" Then
        
        If gobjHost.moveProfile(gstrHQPCC, Trim(cmbBar.Text)) = "PRO" Then
           bolFound = True
        Else
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, "No Company Profile Found", vbOKOnly + vbDefaultButton1, "CWT Desktop - No Profile"
           Exit Function
        End If
   Else
        
        bolFound = GetTravPro(bolSearch)
        If gbolCancelMove = True Then
           
           Exit Function
        End If
        
        If bolFound = True Then bolAdHoc = False
        
        If bolFound = False And bolSearch = True Then
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, "No Traveller Profile Found", vbOKOnly + vbDefaultButton1, "CWT Desktop - No Profile"
           If gobjHost.moveProfile(gstrHQPCC, Trim(cmbBar.Text)) = "PRO" Then
              bolFound = True
           Else
              modMsgBox.OKMsg = "OK"
              modMsgBox.sMsgBox gVPMDIHwnd, "No Company Profile Found", vbOKOnly + vbDefaultButton1, "CWT Desktop - No Profile"
              Exit Function
           End If
        ElseIf bolFound = False And bolSearch = False Then
            If gobjHost.moveProfile(gstrHQPCC, Trim(cmbBar.Text)) = "PRO" Then
              bolFound = True
            Else
              modMsgBox.OKMsg = "OK"
              modMsgBox.sMsgBox gVPMDIHwnd, "No Company Profile Found", vbOKOnly + vbDefaultButton1, "CWT Desktop - No Profile"
              Exit Function
            End If
        End If
   End If

   If bolFound = True Then
   gobjHost.terminalEntry "D.@/0*" & " ATTN " & txtRequestor.Text
   Set gobjPNR = New CWT_GalileoPNR3.PNR
       gobjPNR.loadPNR
       displayPNRinBar
       frmBars.searchPanel
   'Preethi - V1.2.12 20120528 - CR153 - PNR Touch Tracking - Auto Generate
   If gobjPNR.CompInfo.PNRTrackingTouches = True Then
      GetAutoAddTouchedCode "ProfileMove", gobjPNR.CompInfo.WONum, strPrimaryCode, strSecondaryCode, strChargeIndicator, strContactCode
      If strPrimaryCode <> "" And strSecondaryCode <> "" And strChargeIndicator <> "" And strContactCode <> "" Then
         gobjHost.terminalEntry "NP." & strChargeIndicator & "*" & strPrimaryCode & strSecondaryCode & "/D-" & _
                                Format(Date, "ddmmm") & "/T-" & Format(Time, "hhmm") & "/P-" & _
                                gobjHost.AgentPCC & "/CM-" & strContactCode & "/A-" & gobjHost.AgentSine & "/"
      End If
   End If
   End If
   moveProfile = bolFound
   
   gstrProcessGrpID = pGetProcessKey
   pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, gconModProfile, frmSideBar.cmbSelectType.Text, _
      gconSModSearch, Me.Name, IIf(bolAdHoc = True, "MOVEBAR", "MOVEPAR"), gstrProcessGrpID, _
      , datTouchEnd
   
   
End Function

Public Function GetTravPro(ByVal bolSearch As Boolean) As Boolean
    
    Dim strTemp As String
    Dim lngC As Long

    Select Case gobjHost.moveProfile(gstrHQPCC, Trim(cmbBar.Text), Trim(txtLastName(1).Text) & IIf(Trim(txtFirstName(1).Text) <> "", " " & Trim(txtFirstName(1).Text), ""))
    
        Case "PRO"
            GetTravPro = True
            gbolGetProfileFromDB = False
        Case "SNL"
             Load frmSimiliarNames
             frmSimiliarNames.bolProfile = True
             frmSimiliarNames.lvwNameList.ColumnHeaders.Clear
             frmSimiliarNames.lvwNameList.ColumnHeaders.Add 1, , "Title List", 5000, 0
          
             With frmSimiliarNames
                With .lvwNameList
                     .ListItems.Clear
                      For lngC = 1 To gobjHost.SimNameListCount
                          Set item = .ListItems.Add(, , gobjHost.SimNameList(lngC))
                          item.Selected = False
                      Next
                End With
                .Show
                Do
                   DoEvents
                Loop Until isLoaded("frmSimiliarNames") = False
                If gbolCancelMove = True Then
                   Exit Function
                Else
                   If gbolMoveProfile = False Then
                      Select Case gobjHost.moveProfile(gstrHQPCC, Trim(cmbBar.Text), Trim(txtLastName(1).Text) & IIf(Trim(txtFirstName(1).Text) <> "", " " & Trim(txtFirstName(1).Text), ""))
                             Case "PRO"
                                  GetTravPro = True
                      End Select
                   End If
                End If
             End With
    End Select
        
End Function

Private Sub loadRequestType()
    Dim rsRequestType As ADODB.Recordset
    Dim strSql As String
    Dim i As Integer
    
    'Load request type from database
    strSql = "SELECT Type, ID FROM tblRequestType order by Type"
    Set rsRequestType = gdbConn.Execute(strSql)
    cmbSelectReq.Clear
    While Not rsRequestType.EOF
        cmbSelectReq.AddItem rsRequestType.Fields(0) & " - " & rsRequestType.Fields(1)
        cmbSelectReq.ItemData(i) = rsRequestType.Fields(1)
        rsRequestType.MoveNext
        i = i + 1
    Wend
    Set rsRequestType = Nothing
    If cmbSelectReq.ListCount > 0 Then
       cmbSelectReq.listindex = 0
    End If
    AutoSizeDropDownWidth cmbSelectReq
End Sub


Private Sub txtRequestor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If frmLocator.Visible = True Then
        cmdSearch_Click (0)
      ElseIf frmProfile.Visible = True Then
        cmdSearch_Click (1)
      End If
   Else
      KeyAscii = fAllowAlphaNumeric(KeyAscii, " /")
   End If
End Sub

Private Function profileInDB(ByVal strName As String) As Boolean
    Dim rsRecord As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select * from tblAdHocTraveller Where Passenger like '%" & strName & "%'"
    Set rsRecord = gdbConn.Execute(strSql)
    
    If rsRecord.EOF = False Then
        profileInDB = True
    End If
    
    rsRecord.Close
    Set rsRecord = Nothing
    
End Function

Private Sub enableTab()
    Dim i As Integer
    Dim j As Integer
    Dim disableAllSubTabs As Boolean

' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    If gIntModuleType = gModuleType.SYEX Then
        ' ZhiSam - V1.3.2 - 20121019 - BugFix - Always Enable the Tab where do not disableAllSubTab for HKSG SyEx Module
        
    Else
            With frmBars.SftSubTabs(0)
                 If cmbSelectType.Text = "Profile" And InStr(1, cmbSelectReq.Text, "New Booking") > 0 Then
                    lblRequestor.Caption = "Requestor:*"
                    disableAllSubTabs = True
                    For i = 0 To .Tabs.Count - 1
                       If i = 0 Or i = 1 Or i = 2 Or i = 3 Then
                          .Tab(i).Enabled = True
                       Else
                          .Tab(i).Enabled = False
                       End If
                    Next
                 Else
                    lblRequestor.Caption = "Requestor:"
                    disableAllSubTabs = False
                    For i = 0 To .Tabs.Count - 1
                       If i = 3 Then
                          .Tab(i).Enabled = False
                       Else
                          .Tab(i).Enabled = True
                       End If
                    Next
                 End If
            End With
    
    End If
    
    
    
    'For j = 3 To frmBars.SftSubTabs.Count - 1
    '    With frmBars.SftSubTabs(j)
    '         For i = 0 To .Tabs.Count - 1
    '             .Tab(i).Enabled = Not disableAllSubTabs
    '         Next
    '    End With
    'Next
End Sub
