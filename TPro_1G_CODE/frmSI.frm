VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{4832871B-0993-461C-B983-0EAAA4A43E5C}#5.0#0"; "SftTabs_IX86_U_50.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmSI 
   BackColor       =   &H00FAF6EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CWT Desktop - Service Information"
   ClientHeight    =   4275
   ClientLeft      =   1965
   ClientTop       =   3360
   ClientWidth     =   11370
   Icon            =   "frmSI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   11370
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   4260
      Left            =   0
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   7514
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
      Begin SftTabsLib.SftTabs sftTabs 
         Height          =   3720
         Left            =   120
         TabIndex        =   0
         Top             =   45
         Width           =   11145
         PropVer         =   50
         xcx             =   19659
         xcy             =   6562
         PropFile        =   ""
         PropDesignTime  =   1
         DeletePropFile  =   0
         IntVal          =   55
         xBfStyle1       =   63808725
         xBfStyle2       =   -733169
         xBfStyle3       =   -47768774
         xBfStyle4       =   316204230
         TabCount        =   2
         CurrentTab      =   0
         FlatProperties  =   0   'False
         BeginProperty Tab(0) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Special Service Request"
            ToolTip         =   ""
            Object.Align           =   0
            BackColor       =   -1
            BackColorActive =   -1
            ClientAreaColor =   -1
            Enabled         =   1
            FlyByColor      =   0
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
            Text            =   "Other Services Info"
            ToolTip         =   ""
            Object.Align           =   0
            BackColor       =   -1
            BackColorActive =   -1
            ClientAreaColor =   -1
            Enabled         =   1
            FlyByColor      =   0
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
         BackColor       =   -2147483628
         BorderColor     =   6973442
         BrightHighlightColor=   -2147483628
         ForeColor       =   -2147483630
         HighlightColor  =   6973442
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
         xDesign         =   15
         yDesign         =   345
         List(0)Count    =   0
         List(1)Count    =   1
         List(1)(0)Ctl   =   "sharedFra(1)"
         List(1)(0)Ena   =   -1
         List(1)(0)x     =   120
         List(1)(0)y     =   360
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MyFramePanel.MyFrame sharedFra 
            Height          =   3300
            Index           =   1
            Left            =   -12025
            Top             =   -4660
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   5821
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
            HeaderGradientAlign=   5
            HeaderGradientSizeH=   "50%"
            HeaderColorTopLeft=   6973442
            HeaderColorTopRight=   6973442
            HeaderColorBottomLeft=   6973442
            HeaderColorBottomRight=   6973442
            HeaderShow      =   0   'False
            PictureOffsetX  =   5
            Begin VB.Frame fraOSI 
               BackColor       =   &H00DADAB6&
               Caption         =   "Other Service Info"
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   3225
               Left            =   120
               TabIndex        =   14
               Top             =   0
               Width           =   10650
               Begin MSFlexGridLib.MSFlexGrid msFlexOSI 
                  Height          =   2895
                  Left            =   120
                  TabIndex        =   9
                  Top             =   240
                  Width           =   10395
                  _ExtentX        =   18336
                  _ExtentY        =   5106
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   0
                  RowHeightMin    =   280
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
         Begin MyFramePanel.MyFrame sharedFra 
            Height          =   3300
            Index           =   0
            Left            =   120
            Top             =   360
            Width           =   10905
            _ExtentX        =   19235
            _ExtentY        =   5821
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
            Begin VB.Frame fraSSR 
               BackColor       =   &H00DADAB6&
               Caption         =   "Special Service Request"
               ForeColor       =   &H00000000&
               Height          =   3225
               Left            =   120
               TabIndex        =   15
               Top             =   0
               Width           =   10650
               Begin VB.ComboBox cmbSSRStatus 
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmSI.frx":038A
                  Left            =   5760
                  List            =   "frmSI.frx":038C
                  TabIndex        =   5
                  Top             =   600
                  Width           =   825
               End
               Begin VB.TextBox txtSSRText 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Left            =   960
                  TabIndex        =   4
                  Top             =   960
                  Width           =   8385
               End
               Begin VB.ComboBox cmbSSRCode 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmSI.frx":038E
                  Left            =   960
                  List            =   "frmSI.frx":0390
                  Style           =   2  'Dropdown List
                  TabIndex        =   3
                  Top             =   600
                  Width           =   3705
               End
               Begin VB.ComboBox cmbSegment 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmSI.frx":0392
                  Left            =   5760
                  List            =   "frmSI.frx":0394
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   240
                  Width           =   3585
               End
               Begin VB.ComboBox cmbPassenger 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00000000&
                  Height          =   315
                  ItemData        =   "frmSI.frx":0396
                  Left            =   960
                  List            =   "frmSI.frx":0398
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   240
                  Width           =   3705
               End
               Begin MSFlexGridLib.MSFlexGrid msFlexSSR 
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   8
                  Top             =   1560
                  Width           =   10395
                  _ExtentX        =   18336
                  _ExtentY        =   2778
                  _Version        =   393216
                  Cols            =   7
                  FixedCols       =   0
                  RowHeightMin    =   280
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
               Begin MyCommandButton.MyButton btnUpdate 
                  Height          =   300
                  Left            =   9480
                  TabIndex        =   7
                  Top             =   600
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
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
                  TransparentColor=   16447215
                  Caption         =   "&Update"
                  Depth           =   1
                  GradientType    =   2
               End
               Begin MyCommandButton.MyButton bthAdd 
                  Height          =   300
                  Left            =   9480
                  TabIndex        =   6
                  Top             =   240
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
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
                  TransparentColor=   16447215
                  Caption         =   "&Add"
                  Depth           =   1
                  PictureDisabled =   "frmSI.frx":039A
                  GradientType    =   2
               End
               Begin VB.Label lblDOCSSample 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sample format: P/ISS COUNTRY CODE/PPT NBR/NATIONALITY CODE/DOB/GENDER/PPT EXP/LAST NM/FIRST NM/MIDDLE NM"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   180
                  Left            =   960
                  TabIndex        =   26
                  Top             =   1320
                  Width           =   8745
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Status:"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   5040
                  TabIndex        =   23
                  Top             =   660
                  Width           =   495
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "SSRText:"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   22
                  Top             =   960
                  Width           =   690
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "SSR Code:"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   21
                  Top             =   660
                  Width           =   795
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Segment:"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   4920
                  TabIndex        =   20
                  Top             =   300
                  Width           =   675
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Passenger:"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   19
                  Top             =   300
                  Width           =   795
               End
            End
         End
      End
      Begin VB.CheckBox chkEntry 
         Height          =   200
         Left            =   3120
         TabIndex        =   24
         Top             =   3960
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.PictureBox cmbContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   855
         TabIndex        =   17
         Top             =   3900
         Visible         =   0   'False
         Width           =   855
         Begin MSForms.ComboBox cmbEntry 
            Height          =   375
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   855
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1508;661"
            ListWidth       =   7055
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
         Left            =   240
         TabIndex        =   16
         Top             =   3960
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MyCommandButton.MyButton cmdFinish 
         Height          =   360
         Left            =   9000
         TabIndex        =   12
         Top             =   3840
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
         TransparentColor=   16447215
         Caption         =   "&Finish"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdNext 
         Height          =   360
         Left            =   7920
         TabIndex        =   11
         Top             =   3840
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
         TransparentColor=   16447215
         Caption         =   "&Next"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   360
         Left            =   10080
         TabIndex        =   13
         Top             =   3840
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
         TransparentColor=   16447215
         Caption         =   "&Cancel"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdBack 
         Height          =   360
         Left            =   6840
         TabIndex        =   10
         Top             =   3840
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
         TransparentColor=   16447215
         Caption         =   "&Back"
         Depth           =   1
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
         LcK2            =   $"frmSI.frx":06EC
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
      Begin MyCommandButton.MyButton cmdPrevious 
         Height          =   360
         Left            =   5280
         TabIndex        =   25
         Top             =   3840
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
         TransparentColor=   16447215
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
Attribute VB_Name = "frmSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrFlex As String
Dim strSSR As String
Dim strOSI As String
Dim bol1stTab As Boolean
Dim bol2ndTab As Boolean
Dim mbolClickBelowRow As Boolean
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date
Dim mbolCTCMExist As Boolean

Private Sub bthAdd_Click()
    Dim strTemp As String
    Dim strMsg As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim bolExist As Boolean
    'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
    Dim blnDocs As Boolean
    
    With msFlexSSR
        'Perform validation for new SSR (Code, FF Text, Passenger, Segment)
        
        If Trim(cmbPassenger) = "" Then
           strMsg = strMsg & "Missing passenger (SSR)" & Chr(13)
        End If
        If Trim(cmbSegment) = "" Then
           strMsg = strMsg & "Missing segment for record (SSR)" & Chr(13)
        End If
        strTemp = Trim(cmbSSRCode.Text)
        If InStr(strTemp, "-----") > 0 Then
           strTemp = ""
        Else
           i = InStr(strTemp, "-")
           If i > 0 Then
              strTemp = Trim(Mid(strTemp, 1, i - 1))
           End If
        End If

        If strTemp = "" Then
           strMsg = strMsg & "Missing SS code for record (SSR)" & Chr(13)
        Else
           If mandateFFText(strTemp) = True And Trim(txtSSRText) = "" Then
              strMsg = strMsg & "Missing free form text (SSR)" & Chr(13)
           End If
        End If
        
        If UCase(Mid(cmbSSRCode.Text, 1, 4)) = "DOCS" Then
            strMsg = strMsg & ValidTSA(txtSSRText)
        End If
        
        

        If strMsg = "" Then
           If InStr(1, cmbPassenger.Text, "All") > 0 Then
              i = 0
           Else
              i = cmbPassenger.listindex
           End If
           For j = i To IIf(InStr(1, cmbPassenger.Text, "All") > 0, cmbPassenger.ListCount - 2, i)
               If InStr(1, cmbSegment.Text, "All") > 0 Then
                  i = 0
               Else
                  i = cmbSegment.listindex
               End If
               
               'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
                blnDocs = False
                For k = i To IIf(InStr(1, cmbSegment.Text, "All") > 0, cmbSegment.ListCount - 2, i)
                    bolExist = False
                   'Check whether exist in the flexGrid or not
                    For i = 1 To .rows - 1
                        If .TextMatrix(i, 2) = strTemp And .TextMatrix(i, 5) = cmbPassenger.List(j) Then
                           If UCase(Mid(cmbSSRCode.Text, 1, 4)) <> "DOCS" Then
                              If .TextMatrix(i, 6) = cmbSegment.List(k) Then
                                  bolExist = True
                                  Exit For
                              End If
                           Else
                              bolExist = True
                              Exit For
                           End If
                        End If
                    Next
                        If bolExist = False And blnDocs = False Then
                        If .TextMatrix(1, 2) <> "" Then
                            .rows = .rows + 1
                            setText msFlexSSR, .rows - 1, 0, 0
                        End If
                        .TextMatrix(.rows - 1, 2) = strTemp
                        .TextMatrix(.rows - 1, 3) = Trim(txtSSRText)
                        .TextMatrix(.rows - 1, 5) = cmbPassenger.List(j)
                        If UCase(Mid(cmbSSRCode.Text, 1, 4)) <> "DOCS" Then
                           .TextMatrix(.rows - 1, 6) = cmbSegment.List(k)
                        Else
                           blnDocs = True
                        End If
                       End If
                Next
           Next
        Else
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        End If
    End With
End Sub

Private Sub btnUpdate_Click()
    Dim strMsg As String
    Dim i As Integer
    
    i = btnUpdate.Tag
    
    
    With msFlexSSR
        'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
           
         If .TextMatrix(i, 5) <> cmbPassenger.Text Or (.TextMatrix(i, 6) <> cmbSegment.Text And UCase(Mid(cmbSSRCode.Text, 1, 4)) <> "DOCS") Or InStr(1, cmbSSRCode.Text, .TextMatrix(i, 2)) = 0 Then
            strMsg = strMsg & "Change only permitted on status code / free format text" & Chr(13)
         End If
                 
         'If .TextMatrix(i, 5) <> cmbPassenger.Text Or .TextMatrix(i, 6) <> cmbSegment.Text Or InStr(1, cmbSSRCode.Text, .TextMatrix(i, 2)) = 0 Then
         '   strMsg = strMsg & "Change only permitted on status code / free format text" & Chr(13)
         'End If
         
         If mandateFFText(.TextMatrix(i, 2)) And Trim(txtSSRText.Text) = "" Then
            strMsg = strMsg & "Missing free form text (SSR)" & Chr(13)
         End If
         If .TextMatrix(i, 4) <> "" And Trim(cmbSSRStatus.Text) = "" Then
            strMsg = strMsg & "Missing status (SSR)" & Chr(13)
         End If
         
         If UCase(Mid(cmbSSRCode.Text, 1, 4)) = "DOCS" Then
            strMsg = strMsg & ValidTSA(txtSSRText)
         End If
        
         If strMsg = "" Then
            .TextMatrix(i, 3) = Trim(txtSSRText.Text)
            .TextMatrix(i, 4) = Trim(cmbSSRStatus.Text)
         Else
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
         End If
    End With
End Sub

Private Sub cmbEntry_Click()
    Dim i As Integer
    
    If sftTabs.Tabs.Current = 0 And gintY = 1 Then
       If InStr(cmbEntry.Text, "-----") > 0 Then
          cmbEntry.Text = ""
       Else
          i = InStr(cmbEntry.Text, "-")
          If i > 0 Then
             cmbEntry.Text = Trim(Mid(cmbEntry.Text, 1, i - 1))
          End If
       End If
       If allowFFText(cmbEntry.Text) = False Then msFlexSSR.TextMatrix(gintX, 2) = ""
    End If
End Sub

Private Sub cmbEntry_GotFocus()
    cmbGetFocus cmbEntry
End Sub

Private Sub cmbEntry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    control_KeyDown CInt(KeyCode), Shift, Me, cmbEntry.Container
End Sub

Private Sub cmbEntry_KeyPress(KeyCode As MSForms.ReturnInteger)
    KeyAscii = fAllowAlphaNumeric(CInt(KeyAscii), " -")
End Sub

Private Sub cmbEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexSSR.Name Then
       Set msFlex = msFlexSSR
    ElseIf mstrFlex = msFlexOSI.Name Then
       Set msFlex = msFlexOSI
    End If
    If sftTabs.Tabs.Current = 0 Then
       'Clement - 20080812
       'msFlexSSR.SetFocus
    ElseIf sftTabs.Tabs.Current = 1 Then
       'Clement - 20080812
       'msFlexOSI.SetFocus
    End If
    control_LostFocus msFlex, Me, cmbEntry
End Sub

Private Sub cmbSSRCode_Click()
    Dim strTemp As String
    
    strTemp = cmbSSRCode.Text
    If InStr(strTemp, "-----") > 0 Then
       strTemp = ""
    Else
       i = InStr(strTemp, "-")
       If i > 0 Then
          strTemp = Trim(Mid(strTemp, 1, i - 1))
       End If
    End If
    If allowFFText(strTemp) = False Then
       txtSSRText.Enabled = False
       txtSSRText.Text = ""
    Else
       txtSSRText.Enabled = True
    End If
    
    'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
    If UCase(Mid(cmbSSRCode.Text, 1, 4)) = "DOCS" Then
        lblDOCSSample.Visible = True
        cmbSegment.Enabled = False
    Else
        lblDOCSSample.Visible = False
        cmbSegment.Enabled = True
    End If
    
End Sub

Private Sub cmdCancel_Click()
   gbolCancelProcess = True
   Unload Me
End Sub

Private Sub loadSSList()
    Dim i As Integer
    Dim j As Integer
    Dim intC As Integer
    Dim intD As Integer
    Dim intTemp As Integer
    Dim strTemp As String
    Dim strTemp2 As String
    Dim item As ListItem
    Dim bolFound As Boolean
    Dim bolDOCSFound As Boolean
    Dim bolSSRExist As Boolean
    
    bolDOCSFound = False
    bolSSRExist = False
    'Loading from existing SSR
    For i = 1 To gobjPNR.SSRCount
        With gobjPNR.SSR(i)
              strTemp = ""
              intTemp = .SegNum
             'Search Air Segment
             'Preethi - V1.2.2 20110111 - CR28 - Short Format For DOCS
             If Not (.SSCode = "DOCS" And Mid(.Text, 1, 4) = "////") Then
              For intC = 1 To gobjPNR.AirSegCount
                  With gobjPNR.AirSeg(intC)
                       If .segnumber = intTemp And .Flown = False Then
                          strTemp = .segnumber & ". " & .Vendor & " " & Format(.DepartDateTime, "DDMMM HH:NN") & " " & .DepartCityCode & "-" & .ArriveCityCode
                          Exit For
                       End If
                  End With
              Next intC
              If strTemp <> "" Then
                 If i > msFlexSSR.rows - 1 Then
                    msFlexSSR.rows = msFlexSSR.rows + 1
                    setText msFlexSSR, msFlexSSR.rows - 1, 0, 0
                 End If
                 msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 1) = .GFax
                 msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 2) = .SSCode
                 msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 3) = .Text
                 msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 4) = .Status
                 msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 6) = strTemp
                 bolSSRExist = True
                 If .SSCode = "DOCS" Then
                    bolDOCSFound = True
                 End If
                 'Search Passenger Name
                 intTemp = .PsgNum
                 For intC = 1 To gobjPNR.PassengerCount
                     With gobjPNR.PassengerName(intC)
                          If .PassengerNum = intTemp Then
                             msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 5) = Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
                             Exit For
                          End If
                     End With
                 Next intC
              End If
             End If
        End With
    Next
    strTemp = ""
    'Loading preference from NP.G
    For i = 1 To gobjPNR.GeneralRemarkCount
        With gobjPNR.GeneralRemark(i)
             If .Qualifier = "*G" Then
                 intC = InStr(1, .RemarkText, "SPML")
                 If intC > 0 Then
                    If Len(Trim(Mid(.RemarkText, intC + 4))) = 4 Then
                       strTemp = Trim(Mid(.RemarkText, intC + 4))
                    Else
                       strTemp2 = Trim(Mid(.RemarkText, intC + 4))
                    End If
                 Else
                    intC = InStr(1, .RemarkText, "SSR")
                    If intC > 0 Then
                       If Len(Trim(Mid(.RemarkText, intC + 3))) = 4 Then
                            For j = 1 To gobjPNR.AirSegCount
                                bolFound = False
                                If gobjPNR.AirSeg(j).Flown = False Then
                                    For intD = 1 To gobjPNR.SSRCount
                                        If gobjPNR.SSR(intD).SegNum = gobjPNR.AirSeg(j).segnumber And gobjPNR.SSR(intD).SSCode = Trim(Mid(.RemarkText, intC + 3)) Then
                                           bolFound = True
                                           Exit For
                                        End If
                                    Next
                                    If bolFound = False Then
                                       If msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 2) <> "" Then
                                          msFlexSSR.rows = msFlexSSR.rows + 1
                                          setText msFlexSSR, msFlexSSR.rows - 1, 0, 0
                                       End If
                                       If gobjPNR.PassengerCount > 0 Then
                                           msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 2) = Trim(Mid(.RemarkText, intC + 3))
                                           msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 5) = Format(gobjPNR.PassengerName(1).GDSNum, "@@@@ ") & gobjPNR.PassengerName(1).LastName & "/" & gobjPNR.PassengerName(1).FirstName
                                           With gobjPNR.AirSeg(j)
                                                msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 6) = .segnumber & ". " & .Vendor & " " & Format(.DepartDateTime, "DDMMM HH:NN") & " " & .DepartCityCode & "-" & .ArriveCityCode
                                           End With
                                       End If
                                    End If
                                End If
                            Next
                       End If
                    End If
                 End If
             End If
        End With
    Next
    
    If strTemp <> "" Then
        For i = 1 To gobjPNR.AirSegCount
            bolFound = False
            If gobjPNR.AirSeg(i).Flown = False Then
                For intC = 1 To gobjPNR.SSRCount
                    If gobjPNR.SSR(intC).SegNum = gobjPNR.AirSeg(i).segnumber And gobjPNR.SSR(intC).SSCode = strTemp Then
                       bolFound = True
                       Exit For
                    End If
                Next
                If bolFound = False Then
                   If msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 2) <> "" Then
                      msFlexSSR.rows = msFlexSSR.rows + 1
                      setText msFlexSSR, msFlexSSR.rows - 1, 0, 0
                   End If
                   If gobjPNR.PassengerCount > 0 Then
                       msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 2) = strTemp
                       msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 5) = Format(gobjPNR.PassengerName(1).GDSNum, "@@@@ ") & gobjPNR.PassengerName(1).LastName & "/" & gobjPNR.PassengerName(1).FirstName
                       With gobjPNR.AirSeg(i)
                            msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 6) = .segnumber & ". " & .Vendor & " " & Format(.DepartDateTime, "DDMMM HH:NN") & " " & .DepartCityCode & "-" & .ArriveCityCode
                       End With
                   End If
                End If
            End If
        Next
    End If
    
    If bolDOCSFound = False Then
        AddSecureFlightSSR bolSSRExist
    End If
End Sub

Private Sub AddSecureFlightSSR(SSRExist As Boolean)
    Dim colSSR As Collection
    Dim ingI As Integer
    Dim intJ As Integer
    
    Set colSSR = New Collection
    Set colSSR = GenerateSecureFlight
    'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
    For intI = 1 To colSSR.Count
        'For intJ = 1 To gobjPNR.AirSegCount
            'If gobjPNR.AirSeg(intJ).Flown = False Then
                If SSRExist Then
                    msFlexSSR.rows = msFlexSSR.rows + 1
                End If
                SSRExist = True
                setText msFlexSSR, msFlexSSR.rows - 1, 0, 0
                msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 1) = ""
                msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 2) = "DOCS"
                msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 3) = colSSR(intI)
                msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 4) = "NN"
                With gobjPNR.PassengerName(intI)
                    msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 5) = Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
                End With
                'With gobjPNR.AirSeg(intJ)
                    'msFlexSSR.TextMatrix(msFlexSSR.rows - 1, 6) = .segnumber & ". " & .Vendor & " " & Format(.DepartDateTime, "DDMMM HH:NN") & " " & .DepartCityCode & "-" & .ArriveCityCode
                'End With
            'End If
        'Next
    Next
   
End Sub

Private Sub loadOSIList()
    Dim i As Integer

    mbolCTCMExist = False
    For i = 1 To gobjPNR.OSICount
       With gobjPNR.OSI(i)
            If i > msFlexOSI.rows - 1 Then
               msFlexOSI.rows = msFlexOSI.rows + 1
               setText msFlexOSI, msFlexOSI.rows - 1, 0, 0
            End If
            msFlexOSI.TextMatrix(msFlexOSI.rows - 1, 1) = .GFax
            msFlexOSI.TextMatrix(msFlexOSI.rows - 1, 2) = .Vendor
            msFlexOSI.TextMatrix(msFlexOSI.rows - 1, 3) = .Text
            If Mid(UCase(.Text), 1, 4) = "CTCM" Then
                mbolCTCMExist = True
            End If
       End With
    Next
    
    
End Sub



Private Sub cmdFinish_Click()
    If validData Then

       datTouchEnd = Now
       writeDatatoGDS
       'If gbolWritingtoPNR = False Then Exit Sub
       Set gobjPNR = New CWT_GalileoPNR3.PNR
       gobjPNR.loadPNR
       'displayPNRinBar
       
        'Log formload
        'Backup on 26 Sept - Jeremy
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModServiceInfo), _
'       IIf(gbolCreatPNR = True, gconSModServiceInfo, ""), Me.Name, gconFormLoad, gstrProcessGrpID, _
'       datFormLoadEnd, datFormLoadStart
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModServiceInfo), _
'       IIf(gbolCreatPNR = True, gconSModServiceInfo, ""), Me.Name, gconTouch, gstrProcessGrpID, _
'       datTouchEnd, datFormLoadEnd
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModServiceInfo), _
'       IIf(gbolCreatPNR = True, gconSModServiceInfo, ""), Me.Name, gconProcessing, gstrProcessGrpID, _
'        , datTouchEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModServiceInfo, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModServiceInfo, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModServiceInfo, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
       
       Unload Me
    End If
End Sub

Private Sub cmdNext_Click()
    If sftTabs.Tabs.Current + 1 < sftTabs.Tabs.Count Then
       sftTabs.Tabs.Current = sftTabs.Tabs.Current + 1
    End If
End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload Me
End Sub

Private Sub Form_Load()
   Dim oldParent As Long
   Dim strSQL As String
   Dim strTemp As String
   Dim rs As ADODB.Recordset
   datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

   
   Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   pDisplayToFP "*SI"
   
   
   
   lblDOCSSample.Visible = False
   lblDOCSSample.Caption = "Sample format: P/ISS COUNTRY CODE/PPT NBR/NATIONALITY CODE/DOB/GENDER/PPT EXP/LAST NM/FIRST NM/MIDDLE NM"
   
   'Set the header caption and width of msFlexSSR
   With msFlexSSR
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
               .Text = ""
               .ColWidth(i) = 300
            ElseIf i = 1 Then
               .Text = ""
               .ColWidth(i) = 0
            ElseIf i = 2 Then
               .Text = " Code"
               .ColWidth(i) = 800
            ElseIf i = 3 Then
               .Text = " Free Form Text"
               .ColWidth(i) = 2000
            ElseIf i = 4 Then
               .Text = " Status"
               .ColWidth(i) = 600
            ElseIf i = 5 Then
               .Text = " Name"
               .ColWidth(i) = 3500
            ElseIf i = 6 Then
               .Text = " Segment"
               .ColWidth(i) = 2800
            End If
            .ColAlignment(i) = 1
        Next
        setText msFlexSSR, 0, 0, 0
        setText msFlexSSR, 1, 0, 0
        .row = 1
        .col = 1
   End With

   'Set the header caption and width of msFlexOSI
   With msFlexOSI
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
               .Text = ""
               .ColWidth(i) = 300
            ElseIf i = 1 Then
               .Text = " Num"
               .ColWidth(i) = 0
            ElseIf i = 2 Then
               .Text = " Airline"
               .ColWidth(i) = 1000
            ElseIf i = 3 Then
               .Text = " Free Form Text"
               .ColWidth(i) = 8500
            End If
            .ColAlignment(i) = 1
        Next
        setText msFlexOSI, 0, 0, 0
        setText msFlexOSI, 1, 0, 0
        .row = 1
        .col = 1
   End With
    
   strSSR = ""
   strOSI = ""
   'Load Passenger
   For i = 1 To gobjPNR.PassengerCount
        With gobjPNR.PassengerName(i)
             cmbPassenger.AddItem Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
        End With
   Next i
   If cmbPassenger.ListCount > 1 Then cmbPassenger.AddItem "All PAssengers"
   If cmbPassenger.ListCount > 0 Then cmbPassenger.listindex = 0
   
   'Load Segments
   For i = 1 To gobjPNR.AirSegCount
       With gobjPNR.AirSeg(i)
            If .Flown = False Then
                cmbSegment.AddItem .segnumber & ". " & .Vendor & " " & Format(.DepartDateTime, "DDMMM HH:NN") & " " & .DepartCityCode & "-" & .ArriveCityCode
            End If
       End With
   Next i
   'If cmbSegment.ListCount > 1 Then cmbSegment.AddItem "All Segments"
   cmbSegment.AddItem "All Segments"
   If cmbSegment.ListCount > 0 Then cmbSegment.listindex = cmbSegment.ListCount - 1
   'If cmbSegment.ListCount > 0 Then cmbSegment.listindex = 0
   
   'Load SSR Code
   strSQL = "Select * from tblRemarksType Where Type='SS' order by SubType1, SubType2 "
   Set rs = gdbConn.Execute(strSQL)
   i = 0
   While Not rs.EOF
        If 1 = 0 Then
           strTemp = rs!subType1
           cmbSSRCode.AddItem "-- " & rs!subType1 & "-----"
        Else
           If rs!subType1 <> strTemp Then
              strTemp = rs!subType1
              cmbSSRCode.AddItem "-- " & rs!subType1 & "-----"
           End If
        End If
        cmbSSRCode.AddItem rs!subType2
        i = i + 1
        rs.MoveNext
   Wend
   rs.Close
   Set rs = Nothing
   cmbSSRStatus.AddItem "HK"
   cmbSSRStatus.AddItem "HL"
   
   loadSSList 'load existing SSR code and auto capturing SSR preference from NP line
   loadOSIList
   bol1stTab = True
   bol2ndTab = False
   sftTabs.Tabs.Current = 0
   If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
   Else
      cmdPrevious.Visible = False
   End If
   cmdBack.Enabled = False
   
   If mbolCTCMExist Or gbolBackToSI Then
      cmdFinish.Enabled = True
   Else
      cmdFinish.Enabled = False
   End If
   
   
   
   datFormLoadEnd = Now
   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       gbolCancelProcess = True
    End If
End Sub

Public Sub msFlexOSI_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
        
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
        
    intTop = sftTabs.Top + sharedFra(1).Top + fraSSR.Top + msFlexOSI.Top + msFlexOSI.CellTop
    intLeft = sftTabs.Left + sharedFra(1).Left + fraSSR.Left + msFlexOSI.Left + msFlexOSI.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If chkEntry.Visible = True Then chkEntry.Visible = False
    
    mstrFlex = msFlexOSI.Name
    
    If msFlexOSI.col = 0 And msFlexOSI.row = 0 Then
       optSelectAll msFlexOSI
    ElseIf msFlexOSI.col = 0 And msFlexOSI.row > 0 Then
       setControlPosition msFlexOSI, chkEntry, intTop, intLeft
       checkedRow msFlexOSI
       chkEntry_Click
    ElseIf msFlexOSI.col = 3 And msFlexOSI.row > 0 Then
       setControlPosition msFlexOSI, txtEntry, intTop, intLeft
    ElseIf msFlexOSI.col = 2 And msFlexOSI.row > 0 Then
       cmbEntry.Clear
       For i = 1 To gobjPNR.AirSegCount
          With gobjPNR.AirSeg(i)
               If InStr(1, strTemp, .Vendor) = 0 Then
                  strTemp = strTemp & .Vendor & ";"
                  cmbEntry.AddItem .Vendor
               End If
          End With
       Next i
       cmbEntry.AddItem "YY"
       setControlPosition msFlexOSI, cmbContainer, intTop, intLeft, cmbEntry
    End If

End Sub

Private Sub msFlexOSI_KeyDown(KeyCode As Integer, Shift As Integer)
    mstrFlex = msFlexOSI.Name
    control_KeyDown KeyCode, Shift, Me, msFlexOSI
End Sub

Private Sub msFlexOSI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexOSI, Button, Y
End Sub

Public Sub msFlexSSR_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
        
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
        
    intTop = sftTabs.Top + sharedFra(0).Top + fraSSR.Top + msFlexSSR.Top + msFlexSSR.CellTop
    intLeft = sftTabs.Left + sharedFra(0).Left + fraSSR.Left + msFlexSSR.Left + msFlexSSR.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If chkEntry.Visible = True Then chkEntry.Visible = False
    mstrFlex = msFlexSSR.Name
    
    If msFlexSSR.col = 0 And msFlexSSR.row = 0 Then
       optSelectAll msFlexSSR
       btnUpdate.Enabled = False
       cmbSSRStatus.Enabled = False
    ElseIf msFlexSSR.col = 0 And msFlexSSR.row > 0 And msFlexSSR.TextMatrix(msFlexSSR.row, 2) <> "" Then
       setControlPosition msFlexSSR, chkEntry, intTop, intLeft
       checkedRow msFlexSSR
       chkEntry_Click
    End If
End Sub

Private Sub msFlexSSR_KeyDown(KeyCode As Integer, Shift As Integer)
    mstrFlex = msFlexSSR.Name
    control_KeyDown KeyCode, Shift, Me, msFlexSSR
End Sub

Private Sub msFlexSSR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexSSR, Button, Y
End Sub

Private Sub cmdBack_Click()
    If sftTabs.Tabs.Current - 1 >= 0 Then
       sftTabs.Tabs.Current = sftTabs.Tabs.Current - 1
    End If
End Sub

Private Sub SftTabs_Switching(NextTab As Integer, Allow As Boolean, Refresh As Boolean)
    Dim strMsg As String
    
    If sftTabs.Tabs.Current = 1 Then
        strMsg = validateOSI
    End If
    
    If strMsg <> "" Then
       Allow = False
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Sub
    End If
    
    If NextTab = 0 Then
       bol1stTab = True
       cmdBack.Enabled = False
       cmdNext.Enabled = True
    ElseIf NextTab = 1 Then
       bol2ndTab = True
       cmdBack.Enabled = True
       cmdNext.Enabled = False
    End If
    If bol1stTab = True And bol2ndTab = True Then
       cmdFinish.Enabled = True
    Else
       cmdFinish.Enabled = False
    End If
End Sub
Public Sub subMenuAdd_Click()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexOSI.Name Then
       Set msFlex = msFlexOSI
       msFlex.rows = msFlex.rows + 1
       setText msFlex, msFlex.rows - 1, 0, 0
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
    
    If mstrFlex = msFlexSSR.Name Then
       Set msFlex = msFlexSSR
    ElseIf mstrFlex = msFlexOSI.Name Then
       Set msFlex = msFlexOSI
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
            If msFlex.Name = msFlexSSR.Name Then
               msFlex.rows = msFlex.rows + 1
               setText msFlex, msFlex.rows - 1, 0, 0
            Else
               subMenuAdd_Click
            End If
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
         btnUpdate.Enabled = False
         cmbSSRStatus.Enabled = False
    End With
End Sub

Private Sub txtEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, txtEntry
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(CInt(KeyAscii), " -")
End Sub

Private Sub txtEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexSSR.Name Then
       Set msFlex = msFlexSSR
    ElseIf mstrFlex = msFlexOSI.Name Then
       Set msFlex = msFlexOSI
    End If
    If sftTabs.Tabs.Current = 0 Then
       'Clement - 20080812
       'msFlexSSR.SetFocus
    ElseIf sftTabs.Tabs.Current = 1 Then
       'Clement - 20080812
       'msFlexOSI.SetFocus
    End If
    control_LostFocus msFlex, Me, txtEntry, , False
End Sub

Public Function allowFFText(strTemp As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    If Trim(strTemp) = "" Then Exit Function
    strSQL = "Select Additional from tblRemarksType Where SubType2 like '%" & strTemp & " -" & "%'"
    Set rs = gdbConn.Execute(strSQL)
    If Not rs.EOF Then
       ' 0-Must Not Have Additional Text
       ' 1-Must Have Additional Text
       ' 2-Optional Additional Text
       If rs!additional > 0 Then allowFFText = True
    End If
    Set rs = Nothing
End Function

Private Function mandateFFText(strTemp As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    If Trim(strTemp) = "" Then Exit Function
    strSQL = "Select Additional from tblRemarksType Where SubType2 like '%" & strTemp & " -" & "%'"
    Set rs = gdbConn.Execute(strSQL)
    If Not rs.EOF Then
       ' 0-Must Not Have Additional Text
       ' 1-Must Have Additional Text
       ' 2-Optional Additional Text
       If rs!additional = 1 Then mandateFFText = True
    End If
    Set rs = Nothing
End Function

Private Sub writeDatatoGDS()
    Dim i As Integer
    Dim j As Integer
    Dim strCmd As String
    Dim strTemp() As String
    Dim strTemp2 As String
    Dim strMsg As String
    Dim strResponse As String
    Dim intPsgNumForSI As Integer
    Dim intSegNumForSI As Integer
    Dim strRemDOCS As String
    
    'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
    Dim strCmdSSRDOCS As String

    If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
    gbolWritingtoPNR = True

    For i = 1 To msFlexSSR.rows - 1
        With msFlexSSR
        
             If .TextMatrix(i, 1) <> "" Then
                'Existing SSR (Check on Free Form Text & Status
                If Trim(.TextMatrix(i, 3)) <> gobjPNR.SSR(CInt(.TextMatrix(i, 1))).Text Then
                  'Preethi - V1.2.2 20110111 - CR28 - Short Format For DOCS
                  If gobjPNR.SSR(CInt(.TextMatrix(i, 1))).SSCode = "DOCS" Then
                  
                      intPsgNumForSI = gobjPNR.SSR(CInt(.TextMatrix(i, 1))).PsgNum
                      intSegNumForSI = gobjPNR.SSR(CInt(.TextMatrix(i, 1))).SegNum
                      strRemDOCS = RemoveSSRDocs(intPsgNumForSI, intSegNumForSI)
                      gobjHost.terminalEntry strRemDOCS
                      strCmd = strCmd & IIf(strCmd = "", "", "+") & AddFormat(Trim(.TextMatrix(i, 5)), Trim(.TextMatrix(i, 6)), Trim(.TextMatrix(i, 3)), Trim(.TextMatrix(i, 2)))
                      'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
                      strCmdSSRDOCS = strCmdSSRDOCS & IIf(strCmdSSRDOCS = "", "", "+") & AddFormat(Trim(.TextMatrix(i, 5)), Trim(.TextMatrix(i, 6)), Trim(.TextMatrix(i, 3)), Trim(.TextMatrix(i, 2)))
                  Else
                    strCmd = strCmd & IIf(strCmd = "", "", "+") & "SI.P" & gobjPNR.SSR(CInt(.TextMatrix(i, 1))).PsgNum & _
                             "S" & gobjPNR.SSR(CInt(.TextMatrix(i, 1))).SegNum & "/" & gobjPNR.SSR(CInt(.TextMatrix(i, 1))).SSCode & _
                            "@*" & Trim(.TextMatrix(i, 3))
                   End If
                End If

                If Trim(.TextMatrix(i, 4)) <> gobjPNR.SSR(CInt(.TextMatrix(i, 1))).Status Then
                   strCmd = strCmd & IIf(strCmd = "", "", "+") & "SI.P" & gobjPNR.SSR(CInt(.TextMatrix(i, 1))).PsgNum & _
                            "S" & gobjPNR.SSR(CInt(.TextMatrix(i, 1))).SegNum & "/" & gobjPNR.SSR(CInt(.TextMatrix(i, 1))).SSCode & _
                            "@" & Trim(.TextMatrix(i, 4))
                End If
             Else
                 'Preethi - V1.2.2 20110111 - CR28 - Short Format For DOCS
                 'New SSR
                 If Trim(.TextMatrix(i, 5)) <> "" Then
                 'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
                   If Trim(.TextMatrix(i, 2)) = "DOCS" Then
                        strCmdSSRDOCS = strCmdSSRDOCS & IIf(strCmdSSRDOCS = "", "", "+") & AddFormat(Trim(.TextMatrix(i, 5)), Trim(.TextMatrix(i, 6)), Trim(.TextMatrix(i, 3)), Trim(.TextMatrix(i, 2)))
                    Else
                       strCmd = strCmd & IIf(strCmd = "", "", "+") & AddFormat(Trim(.TextMatrix(i, 5)), Trim(.TextMatrix(i, 6)), Trim(.TextMatrix(i, 3)), Trim(.TextMatrix(i, 2)))
                    End If
                 End If
             End If
        End With
    Next
    
    'Must Delete SSR 1st
    If strSSR <> "" Then
       strCmd = strSSR & IIf(strCmd <> "", "+" & strCmd, "")
    End If

    For i = 1 To msFlexOSI.rows - 1
        With msFlexOSI
             If .TextMatrix(i, 1) <> "" Then
                'Existing OSI (Check on carrier and free form text)
                If Trim(.TextMatrix(i, 2)) <> gobjPNR.OSI(CInt(.TextMatrix(i, 1))).Vendor _
                   Or Trim(.TextMatrix(i, 3)) <> gobjPNR.OSI(CInt(.TextMatrix(i, 1))).Text Then
                   strCmd = strCmd & IIf(strCmd = "", "", "+") & "SI." & .TextMatrix(i, 1) & "@" & .TextMatrix(i, 2) & "*" & .TextMatrix(i, 3)
                End If
             Else
                If Trim(.TextMatrix(i, 2)) <> "" Then
                   strCmd = strCmd & IIf(strCmd <> "", "+", "") & "SI." & .TextMatrix(i, 2) & "*" & .TextMatrix(i, 3)
                End If
             End If
        End With
    Next
    
    'Delete OSI at the last
    If strOSI <> "" Then
       If Right(strOSI, 1) = "." Then strOSI = Mid(strOSI, 1, Len(strOSI) - 1)
       strCmd = strCmd & IIf(strCmd <> "", "+", "") & "SI." & sortInt(strOSI) & "@"
    End If

    'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
    
    If strCmdSSRDOCS <> "" Then
        functSIEndPNR (strCmdSSRDOCS)
    End If
    If strCmd <> "" Then
        functSIEndPNR (strCmd)
    End If
    
    'Preethi - V1.2.4 20110614 - CR 76 - Change Validation Logic For ENDPNR
    'send entries, received & end the PNR
'    If strCmd <> "" Then
'         If gbolCreatPNR = True And gobjPNR.RecLoc = "" Then
'        Else
'        strCmd = strCmd & IIf(strCmd = "", "", "+") & "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
'        End If
'        strCmd = strCmd & "+ER+ER+ER"
'        strResponse = gobjHost.terminalEntry(strCmd)
'        strTemp2 = strResponse
'        'Preethi - V1.2.4 20110712 - CR 76 - Change Validation Logic For ENDPNR
'        'If InStr(strTemp2, "1.1") = 0 Then
'        If CheckResponse(strTemp2, gstrPNRExpression, gintCheckERLineNum) = False Then
'           For i = 0 To 1
'               strTemp2 = gobjHost.terminalEntry("ER")
'               'If InStr(strTemp2, "1.1") > 0 Then
'               If CheckResponse(strTemp2, gstrPNRExpression, gintCheckERLineNum) = True Then
'                   Exit For
'               End If
'               If i = 1 Then GoTo errorWriting
'           Next
'        End If
''        strTemp2 = gobjHost.EndPNR2(IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine), True, 2)
''        If strTemp2 <> "True" Then GoTo errorWriting
'        pDisplayToFP ("*SI")
'    End If
'
'    Exit Sub
'
'errorWriting:
'    'Prompt error message if failed to write to PNR
'    gbolWritingtoPNR = False
'    strMsg = "Unable to write to PNR. Response from GDS is " & Chr(13) & strResponse
'    strMsg = strMsg & Chr(13) & "System will continue without ending this booking."
'
'    modMsgBox.OKMsg = "OK"
'    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
    'gobjHost.TerminalEntry "IR"
    
End Sub

Private Function validData() As Boolean
    Dim i As Integer
    Dim strMsg As String
    validData = True
            
    strMsg = validateOSI
    
    If strMsg <> "" Then
       validData = False
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Function
    End If
    
    validData = ValidTSAInMsFlex
    

End Function

Private Sub chkEntry_Click()
    Dim msFlex As MSFlexGrid
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    If mstrFlex = msFlexSSR.Name Then
       Set msFlex = msFlexSSR
    ElseIf mstrFlex = msFlexOSI.Name Then
       Set msFlex = msFlexOSI
    End If
    checkedRow msFlex
    If mstrFlex = msFlexSSR.Name Then
        j = 0
        For i = 1 To msFlex.rows - 1
            If i <> gintX Then
               If msFlex.TextMatrix(i, 0) = gstrChecked Then
                  j = j + 1
                  k = i
               End If
            Else
               If chkEntry.value = vbChecked Then
                  j = j + 1
                  k = i
               End If
            End If
        Next
        If j = 1 Then
           btnUpdate.Enabled = True
           btnUpdate.Tag = k
           cmbPassenger.Text = msFlex.TextMatrix(k, 5)
           'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
           If msFlex.TextMatrix(k, 2) <> "DOCS" Then
              cmbSegment.Text = msFlex.TextMatrix(k, 6)
           End If
           For i = 0 To cmbSSRCode.ListCount - 1
               If InStr(1, cmbSSRCode.List(i), msFlex.TextMatrix(k, 2)) > 0 Then
                  cmbSSRCode.listindex = i
                  cmbSSRCode_Click
                  Exit For
               End If
           Next
           txtSSRText.Text = msFlex.TextMatrix(k, 3)
           If msFlex.TextMatrix(k, 4) <> "" Then
              cmbSSRStatus.Text = msFlex.TextMatrix(k, 4)
              cmbSSRStatus.Enabled = True
           Else
              cmbSSRStatus.Text = ""
              cmbSSRStatus.Enabled = False
           End If
        Else
           btnUpdate.Enabled = False
           cmbSSRStatus.Enabled = False
        End If
    End If
End Sub

Private Sub chkEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexSSR.Name Then
       Set msFlex = msFlexSSR
       'Clement - 20080812
       'If sftTabs.Tabs.Current = 0 Then msFlexSSR.SetFocus
    ElseIf mstrFlex = msFlexOSI.Name Then
       Set msFlex = msFlexOSI
       'Clement - 20080812
       'If sftTabs.Tabs.Current = 1 Then msFlexOSI.SetFocus
    End If
    control_LostFocus msFlex, Me, chkEntry
End Sub

Private Sub checkedRow(ByRef msFlex As MSFlexGrid)
    If chkEntry.value = vbChecked And gintY = 0 Then
       HighlightRow msFlex, gintX
    ElseIf chkEntry.value = vbUnchecked And gintY = 0 Then
       HighlightRow msFlex, gintX, False
    End If
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
         Else
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
   'Preethi - V1.2.2 20110111 - CR28 - Short Format For DOCS
   Dim strRemDOCS As String
   Dim intPaxNumForSI As Integer
   Dim intSegNumForSI As Integer
       
   With msFlex
        If .Name = msFlexSSR.Name Then
            If .TextMatrix(i, 1) <> "" Then
                With gobjPNR.SSR(CInt(msFlex.TextMatrix(i, 1)))
                    'Preethi - V1.2.2 20110111 - CR28 - Short Format For DOCS
                     If .SSCode = "DOCS" Then
                        intPaxNumForSI = .PsgNum
                        intSegNumForSI = .SegNum
                        strRemDOCS = RemoveSSRDocs(intPaxNumForSI, intSegNumForSI)
                        strSSR = strSSR & IIf(strSSR = "", "", "+") & strRemDOCS
                     Else
                       strSSR = strSSR & IIf(strSSR = "", "", "+") & _
                       "SI.P" & .PsgNum & "S" & .SegNum & "/" & .SSCode & "@"
                     End If
                End With
            End If
        ElseIf mstrFlex = msFlexOSI.Name Then
            If msFlex.TextMatrix(msFlex.row, 1) <> "" Then
               strOSI = strOSI & msFlex.TextMatrix(msFlex.row, 1) & "."
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

Private Function validateOSI() As String
    Dim strMsg As String
     
    With msFlexOSI
         For i = 1 To .rows - 1
            If Trim(.TextMatrix(i, 2)) <> "" Or Trim(.TextMatrix(i, 3)) <> "" Then
               If Trim(.TextMatrix(i, 2)) = "" Then
                  strMsg = strMsg & "Missing airline for record " & i & " (OSI)" & Chr(13)
               End If
               If Trim(.TextMatrix(i, 3)) = "" Then
                  strMsg = strMsg & "Missing free form text for record " & i & " (OSI)" & Chr(13)
               End If
            End If
         Next
    End With
    validateOSI = strMsg

End Function

Private Sub txtSSRText_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function ValidTSAInMsFlex() As Boolean
    Dim intI As Integer
    Dim strMsg As String
    Dim strErr As String
    
    For intI = 1 To msFlexSSR.rows - 1
        If msFlexSSR.TextMatrix(intI, 2) = "DOCS" Then
            strMsg = ValidTSA(msFlexSSR.TextMatrix(intI, 3))
            If strMsg <> "" Then
                strErr = "Invalid DOCS - DHS format in SSR line " & intI & " : " & vbCrLf & strMsg
                ValidTSAInMsFlex = False
                sftTabs.Tabs.Current = 0
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strErr, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
                Exit Function
            End If
        End If
    Next
    
    ValidTSAInMsFlex = True
End Function


Private Function ValidTSA(DHS As String) As String
    Dim strTmp() As String
    Dim strErr As String
    Dim intI As Integer
    Dim strTemp As String
    
    'P/GB/S12345888/GB/12DEC06/MI/23JAN12/HO/LINDA
    strTmp = Split(DHS, "/")
    If UBound(strTmp) <> 8 And UBound(strTmp) <> 9 Then
        strErr = "Invalid DHS format"
    Else
        For intI = 0 To 9
            Select Case intI
               Case 0  'P
                   If strTmp(intI) = "" Then
                       strErr = strErr & Space(2) & "Missing P in 1st letter" & vbCrLf
                   End If
               Case 1  'Country Code
                   If Len(strTmp(intI)) <> 2 Then
                       strErr = strErr & Space(2) & "Invalid Issue Country Code" & vbCrLf
                   End If
               Case 2  'Passport Num
                   If Trim(strTmp(intI)) = "" Then
                       strErr = strErr & Space(2) & "Invalid Passport Number" & vbCrLf
                   End If
               Case 3  'Country Code
                   If strTmp(intI) = "" Then
                       'Allow empty
                   ElseIf Len(strTmp(intI)) <> 2 Then
                       strErr = strErr & Space(2) & "Invalid Nationality Country Code" & vbCrLf
                   End If
               Case 4  'Birth Date (ddMMMyy)
                   'If strTmp(intI) = "" Then
                   '    'Allow empty
                   'Else
                   If Len(strTmp(intI)) <> 7 Then
                       strErr = strErr & Space(2) & "Invalid Birthdate (e.g. 05FEB80)" & vbCrLf
                   Else
                       strTemp = Mid(strTmp(intI), 1, 2) & "/" & Mid(strTmp(intI), 3, 3) & "/" & Mid(strTmp(intI), 6, 2)
                       If IsDate(strTemp) = False Then
                           strErr = strErr & Space(2) & "Invalid Birthdate (e.g. 05FEB80)" & vbCrLf
                       End If
                   End If
               Case 5  'M, MI, F, FI
                   'If strTmp(intI) = "" Then
                   '    'Allow empty
                   'Else
                   If strTmp(intI) <> "F" And strTmp(intI) <> "FI" _
                       And strTmp(intI) <> "M" And strTmp(intI) <> "MI" Then
                       strErr = strErr & Space(2) & "Invalid Gender" & vbCrLf
                   End If
              Case 6   'Passport Exp Date (ddMMMyy)
                   If strTmp(intI) = "" Then
                       'Allow empty
                   ElseIf Len(strTmp(intI)) <> 7 Then
                       strErr = strErr & Space(2) & "Invalid Passport Exp Date Format (e.g. 05DEC12)" & vbCrLf
                   Else
                       strTemp = Mid(strTmp(intI), 1, 2) & "/" & Mid(strTmp(intI), 3, 3) & "/" & Mid(strTmp(intI), 6, 2)
                       If IsDate(strTemp) = False Then
                           strErr = strErr & Space(2) & "Invalid Passport Exp Date Format (e.g. 05DEC12)" & vbCrLf
                       End If
                   End If
              Case 7    'Last Name
                   If Trim(strTmp(intI)) = "" Then
                       strErr = strErr & Space(2) & "Invalid Last Name" & vbCrLf
                   End If
              Case 8    'First Name
                   If Trim(strTmp(intI)) = "" Then
                       strErr = strErr & Space(2) & "Invalid First Name" & vbCrLf
                   End If
              Case 9    'Middle Name
                   'No need validation, can be empty
            End Select
        Next
        If strErr <> "" Then
            strErr = "Invalid DHS format: " & vbCrLf & strErr
        End If
    End If
        
    If strErr <> "" Then
        ValidTSA = strErr
    End If
    
End Function
'Preethi - V1.2.2 20110111 - CR28 - Short Format For DOCS
Public Function RemoveSSRDocs(intPsgNum As Integer, intSegNum As Integer) As String
Dim strSSR As String
Dim intI As Integer

For intI = 1 To gobjPNR.SSRCount
     With gobjPNR.SSR(intI)
         If .SSCode = "DOCS" And .PsgNum = intPsgNum And .SegNum = intSegNum Then
           strSSR = strSSR & IIf(strSSR = "", "", "+") & "SI.P" & .PsgNum & "S" & .SegNum & "/" & "DOCS@"
         End If
     End With
Next intI
RemoveSSRDocs = strSSR

End Function
'Preethi - V1.2.2 20110112 - CR28 - Short Format For DOCS
Public Function AddFormat(strName As String, strSegmnts As String, strText As String, strSScode As String) As String
 
 Dim strCmd As String
 Dim intJ As Integer
 Dim strShortTemp() As String
 Dim strTemp() As String
 Dim intPsgNum As Integer
 Dim intSegNum As Integer
 
'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
 Dim strSSRCmd As String

 strSSRCmd = "SSRDOCSYYHK1"
 
 strTemp = Split(Trim(strName), " ")
 For intJ = 1 To gobjPNR.PassengerCount
     With gobjPNR.PassengerName(intJ)
         If .GDSNum = strTemp(0) Then
             intPsgNum = .PassengerNum
             strCmd = strCmd & IIf(strCmd <> "", "+", "") & "SI.P" & .PassengerNum
             Exit For
         End If
      End With
 Next

'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format

 If Trim(strSScode) = "DOCS" Then
 
    strTemp = Split(Trim(strName), " ")
    strShortTemp = Split(Trim(strText), "/")
    strCmd = strCmd & "/" & strSSRCmd & "/" & IIf(Trim(strText) <> "", "" & Trim(strText), "")
    strCmd = strCmd & IIf(strCmd <> "", "+", "") & "SI.P" & intPsgNum
    ' new format doesnt include segment number
    strCmd = strCmd & "/" & strSSRCmd & "/////" & strShortTemp(4) & _
             "/" & strShortTemp(5) & "//" & strShortTemp(7) & "/" & strShortTemp(8)
'    strCmd = strCmd & "S" & intSegNum & "/" & Trim(strSScode) & "*" & "////" & strShortTemp(4) & _
'             "/" & strShortTemp(5) & "//" & strShortTemp(7) & "/" & strShortTemp(8)
                       
 Else
    strTemp = Split(Trim(strSegmnts), ".")
    intSegNum = Trim(strTemp(0))
    strCmd = strCmd & "S" & intSegNum & "/" & Trim(strSScode) & _
             IIf(Trim(strText) <> "", "*" & Trim(strText), "")
 End If

 AddFormat = strCmd
  
End Function
'Preethi - V1.2.10 20120319 - CR143 - Change in DOCS format
Private Function functSIEndPNR(strCmd As String) As Boolean
    Dim i As Integer
    Dim strResponse As String
    Dim strTemp As String
    Dim strMsg As String

   If strCmd <> "" Then
         If gbolCreatPNR = True And gobjPNR.RecLoc = "" Then
         Else
            strCmd = strCmd & IIf(strCmd = "", "", "+") & "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
         End If
        strCmd = strCmd & "+ER+ER+ER"
        strResponse = gobjHost.terminalEntry(strCmd)
        strTemp = strResponse
        
        'Preethi - V1.2.4 20110712 - CR 76 - Change Validation Logic For ENDPNR
        'If InStr(strTemp2, "1.1") = 0 Then
        If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = False Then
           For i = 0 To 1
               strTemp = gobjHost.terminalEntry("ER")
               'If InStr(strTemp2, "1.1") > 0 Then
               If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = True Then
                   Exit For
               End If
               If i = 1 Then GoTo errorWriting
           Next
        End If
'        strTemp2 = gobjHost.EndPNR2(IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine), True, 2)
'        If strTemp2 <> "True" Then GoTo errorWriting
        pDisplayToFP ("*SI")
    End If
    
    functSIEndPNR = True
    Exit Function
    
errorWriting:
    'Prompt error message if failed to write to PNR
    gbolWritingtoPNR = False
    functSIEndPNR = False
    strMsg = "Unable to write to PNR. Response from GDS is " & Chr(13) & strResponse
    strMsg = strMsg & Chr(13) & "System will continue without ending this booking."
 
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
    
    'gobjHost.TerminalEntry "IR"


End Function
