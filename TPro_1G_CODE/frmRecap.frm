VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4832871B-0993-461C-B983-0EAAA4A43E5C}#5.0#0"; "SftTabs_IX86_U_50.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmRecap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " CWT Desktop - Recapitulate"
   ClientHeight    =   3420
   ClientLeft      =   3690
   ClientTop       =   2775
   ClientWidth     =   11370
   Icon            =   "frmRecap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   11370
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   3800
      Left            =   0
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   6694
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
         Height          =   3000
         Left            =   60
         TabIndex        =   0
         Top             =   45
         Width           =   11205
         PropVer         =   50
         xcx             =   19764
         xcy             =   5292
         PropFile        =   ""
         PropDesignTime  =   1
         DeletePropFile  =   0
         IntVal          =   55
         xBfStyle1       =   63747964
         xBfStyle2       =   -396575009
         xBfStyle3       =   -443545741
         xBfStyle4       =   443545757
         TabCount        =   7
         CurrentTab      =   6
         FlatProperties  =   0   'False
         BeginProperty Tab(0) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Contact && FOP"
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
            Text            =   "Emails"
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
         BeginProperty Tab(2) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Passport Details"
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
         BeginProperty Tab(3) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Address"
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
         BeginProperty Tab(4) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Frequent Flyer"
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
         BeginProperty Tab(5) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Pre-trip Reporting Field"
            ToolTip         =   ""
            Object.Align           =   0
            BackColor       =   -1
            BackColorActive =   -1
            ClientAreaColor =   -1
            Enabled         =   0
            FlyByColor      =   0
            ForeColor       =   -1
            ForeColorActive =   -1
            Name            =   ""
            Hidden          =   1
            BackColorStart  =   -1
            BackColorEnd    =   -1
            BackColorActiveStart=   -1
            BackColorActiveEnd=   -1
            ClientAreaColorStart=   -1
            ClientAreaColorEnd=   -1
         EndProperty
         BeginProperty Tab(6) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Post-trip Reporting Field"
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
         BorderColor     =   -2147483628
         BrightHighlightColor=   -2147483628
         ForeColor       =   0
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
         List(0)Count    =   1
         List(0)(0)Ctl   =   "sharedFra(0)"
         List(0)(0)Ena   =   -1
         List(0)(0)x     =   105
         List(0)(0)y     =   345
         List(1)Count    =   1
         List(1)(0)Ctl   =   "sharedFra(1)"
         List(1)(0)Ena   =   -1
         List(1)(0)x     =   105
         List(1)(0)y     =   360
         List(2)Count    =   1
         List(2)(0)Ctl   =   "sharedFra(2)"
         List(2)(0)Ena   =   -1
         List(2)(0)x     =   105
         List(2)(0)y     =   405
         List(3)Count    =   1
         List(3)(0)Ctl   =   "sharedFra(3)"
         List(3)(0)Ena   =   -1
         List(3)(0)x     =   105
         List(3)(0)y     =   405
         List(4)Count    =   1
         List(4)(0)Ctl   =   "sharedFra(4)"
         List(4)(0)Ena   =   -1
         List(4)(0)x     =   105
         List(4)(0)y     =   405
         List(5)Count    =   1
         List(5)(0)Ctl   =   "sharedFra(5)"
         List(5)(0)Ena   =   -1
         List(5)(0)x     =   120
         List(5)(0)y     =   405
         List(6)Count    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MyFramePanel.MyFrame sharedFra 
            Height          =   2505
            Index           =   2
            Left            =   -12040
            Top             =   -3910
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4419
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
            Begin VB.Frame fraPassport 
               BackColor       =   &H00DADAB6&
               Caption         =   "Passport"
               Enabled         =   0   'False
               Height          =   2350
               Left            =   120
               TabIndex        =   12
               Top             =   50
               Width           =   10740
               Begin MSFlexGridLib.MSFlexGrid msFlexPassport 
                  Height          =   2055
                  Left            =   120
                  TabIndex        =   2
                  Top             =   195
                  Width           =   10425
                  _ExtentX        =   18389
                  _ExtentY        =   3625
                  _Version        =   393216
                  Rows            =   8
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
         Begin MyFramePanel.MyFrame sharedFra 
            Height          =   2505
            Index           =   3
            Left            =   -12040
            Top             =   -3910
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4419
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
            Begin VB.Frame fraBillAddress 
               BackColor       =   &H00DADAB6&
               Caption         =   " Billing Address"
               Enabled         =   0   'False
               Height          =   2300
               Left            =   5640
               TabIndex        =   14
               Top             =   50
               Width           =   5175
               Begin VB.TextBox txtBillAddr 
                  Height          =   1800
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   4
                  Top             =   360
                  Width           =   4905
               End
            End
            Begin VB.Frame fraDelivAddress 
               BackColor       =   &H00DADAB6&
               Caption         =   " Delivery Address"
               Enabled         =   0   'False
               Height          =   2300
               Left            =   240
               TabIndex        =   13
               Top             =   50
               Width           =   5295
               Begin VB.TextBox txtDeliveryAddr 
                  Height          =   1800
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   3
                  Top             =   360
                  Width           =   5025
               End
            End
         End
         Begin MyFramePanel.MyFrame sharedFra 
            Height          =   2505
            Index           =   5
            Left            =   -12040
            Top             =   -3910
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4419
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
            Begin VB.Frame fraPretripMI 
               BackColor       =   &H00DADAB6&
               Caption         =   " Pre-Trip Reporting Field"
               Enabled         =   0   'False
               Height          =   2350
               Left            =   100
               TabIndex        =   15
               Top             =   50
               Width           =   10740
               Begin MSFlexGridLib.MSFlexGrid msFlexPretripMI 
                  Height          =   2055
                  Left            =   120
                  TabIndex        =   6
                  Top             =   240
                  Width           =   10455
                  _ExtentX        =   18441
                  _ExtentY        =   3625
                  _Version        =   393216
                  Cols            =   5
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
            Height          =   2505
            Index           =   6
            Left            =   105
            Top             =   405
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4419
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
            Begin VB.Frame fraPostTripMI 
               BackColor       =   &H00DADAB6&
               Caption         =   " Post-Trip Reporting Field"
               Height          =   2350
               Left            =   100
               TabIndex        =   16
               Top             =   50
               Width           =   10740
               Begin MSFlexGridLib.MSFlexGrid msFlexPostTripMI 
                  Height          =   2055
                  Left            =   120
                  TabIndex        =   7
                  Top             =   240
                  Width           =   10455
                  _ExtentX        =   18441
                  _ExtentY        =   3625
                  _Version        =   393216
                  Cols            =   5
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
            Height          =   2580
            Index           =   0
            Left            =   -12040
            Top             =   -3925
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4551
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
            Begin VB.Frame fraFOP 
               BackColor       =   &H00DADAB6&
               Caption         =   " Form of Payment"
               Enabled         =   0   'False
               Height          =   1815
               Left            =   5160
               TabIndex        =   25
               Top             =   50
               Width           =   5655
               Begin VB.TextBox txtCCNum 
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   2
                  Left            =   2280
                  MaxLength       =   16
                  TabIndex        =   35
                  Top             =   1080
                  Width           =   2000
               End
               Begin VB.ComboBox cboCCCode 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   2
                  ItemData        =   "frmRecap.frx":038A
                  Left            =   1260
                  List            =   "frmRecap.frx":038C
                  Style           =   2  'Dropdown List
                  TabIndex        =   34
                  Top             =   1080
                  Width           =   1000
               End
               Begin VB.TextBox txtCCNum 
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   1
                  Left            =   2280
                  MaxLength       =   16
                  TabIndex        =   31
                  Top             =   720
                  Width           =   2000
               End
               Begin VB.ComboBox cboCCCode 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   1
                  ItemData        =   "frmRecap.frx":038E
                  Left            =   1260
                  List            =   "frmRecap.frx":0390
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   720
                  Width           =   1000
               End
               Begin VB.TextBox txtCCNum 
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   0
                  Left            =   2280
                  MaxLength       =   16
                  TabIndex        =   27
                  Top             =   360
                  Width           =   2000
               End
               Begin VB.ComboBox cboCCCode 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   0
                  ItemData        =   "frmRecap.frx":0392
                  Left            =   1260
                  List            =   "frmRecap.frx":0394
                  Style           =   2  'Dropdown List
                  TabIndex        =   26
                  Top             =   360
                  Width           =   1000
               End
               Begin MSComCtl2.DTPicker dtpCCExpire 
                  Height          =   315
                  Index           =   0
                  Left            =   4320
                  TabIndex        =   28
                  Top             =   360
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   556
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
                  CalendarTitleBackColor=   -2147483647
                  CalendarTitleForeColor=   16777215
                  CustomFormat    =   "M/yyyy"
                  Format          =   63832067
                  CurrentDate     =   37987
                  MaxDate         =   73050
                  MinDate         =   36526
               End
               Begin MSComCtl2.DTPicker dtpCCExpire 
                  Height          =   315
                  Index           =   1
                  Left            =   4320
                  TabIndex        =   32
                  Top             =   720
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   556
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
                  CalendarTitleBackColor=   -2147483647
                  CalendarTitleForeColor=   16777215
                  CustomFormat    =   "M/yyyy"
                  Format          =   63832067
                  CurrentDate     =   37987
                  MaxDate         =   73050
                  MinDate         =   36526
               End
               Begin MSComCtl2.DTPicker dtpCCExpire 
                  Height          =   315
                  Index           =   2
                  Left            =   4320
                  TabIndex        =   36
                  Top             =   1080
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   556
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
                  CalendarTitleBackColor=   -2147483647
                  CalendarTitleForeColor=   16777215
                  CustomFormat    =   "M/yyyy"
                  Format          =   63832067
                  CurrentDate     =   37987
                  MaxDate         =   73050
                  MinDate         =   36526
               End
               Begin VB.Label lblFOP 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Car Preferred:"
                  Height          =   195
                  Index           =   2
                  Left            =   255
                  TabIndex        =   37
                  Top             =   1125
                  Width           =   975
               End
               Begin VB.Label lblFOP 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hotel Preferred:"
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   33
                  Top             =   765
                  Width           =   1110
               End
               Begin VB.Label lblFOP 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Air Preferred:"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   29
                  Top             =   400
                  Width           =   1110
               End
            End
            Begin VB.Frame fraContacts 
               BackColor       =   &H00DADAB6&
               Caption         =   " Contact Details"
               Enabled         =   0   'False
               Height          =   1815
               Left            =   120
               TabIndex        =   17
               Top             =   50
               Width           =   4935
               Begin MSFlexGridLib.MSFlexGrid msFlexContacts 
                  Height          =   1335
                  Left            =   120
                  TabIndex        =   1
                  Top             =   285
                  Width           =   4695
                  _ExtentX        =   8281
                  _ExtentY        =   2355
                  _Version        =   393216
                  Rows            =   5
                  Cols            =   6
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
         Begin MyFramePanel.MyFrame sharedFra 
            Height          =   2505
            Index           =   4
            Left            =   -12040
            Top             =   -3910
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4419
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
            Begin VB.Frame fraFFlyer 
               BackColor       =   &H00DADAB6&
               Caption         =   " Frequent Flyer"
               Enabled         =   0   'False
               Height          =   2350
               Left            =   120
               TabIndex        =   20
               Top             =   50
               Width           =   10740
               Begin MSFlexGridLib.MSFlexGrid msFlexFFlyer 
                  Height          =   2055
                  Left            =   120
                  TabIndex        =   5
                  Top             =   240
                  Width           =   10455
                  _ExtentX        =   18441
                  _ExtentY        =   3625
                  _Version        =   393216
                  Cols            =   5
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
            Height          =   2580
            Index           =   1
            Left            =   -12040
            Top             =   -3940
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4551
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
            Begin VB.Frame fraEmails 
               BackColor       =   &H00DADAB6&
               Caption         =   "Emails *"
               Enabled         =   0   'False
               Height          =   2400
               Left            =   120
               TabIndex        =   23
               Top             =   120
               Width           =   8415
               Begin MSFlexGridLib.MSFlexGrid msFlexEmails 
                  Height          =   1995
                  Left            =   120
                  TabIndex        =   24
                  Top             =   285
                  Width           =   8205
                  _ExtentX        =   14473
                  _ExtentY        =   3519
                  _Version        =   393216
                  Cols            =   8
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
      Begin VB.CheckBox chkEntry 
         Height          =   200
         Left            =   4680
         TabIndex        =   38
         Top             =   3120
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.PictureBox cmbContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   1005
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   1000
         Begin MSForms.ComboBox cmbEntry 
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   855
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1508;661"
            ListWidth       =   7055
            ColumnCount     =   2
            cColumnInfo     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   34
            Object.Width           =   "3527;0"
         End
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker dtpEntry 
         Height          =   315
         Left            =   3120
         TabIndex        =   19
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CheckBox        =   -1  'True
         CustomFormat    =   "ddMMMyy"
         Format          =   63832065
         CurrentDate     =   36161
         MaxDate         =   109574
         MinDate         =   21916
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
         LcK2            =   $"frmRecap.frx":0396
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
      Begin MyCommandButton.MyButton cmdFinish 
         Height          =   360
         Left            =   9120
         TabIndex        =   10
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
         TransparentColor=   14215660
         Caption         =   "&Finish"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdNext 
         Height          =   360
         Left            =   8040
         TabIndex        =   9
         Top             =   3100
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
         Caption         =   "&Next"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   360
         Left            =   10200
         TabIndex        =   11
         Top             =   3100
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
      Begin MyCommandButton.MyButton cmdBack 
         Height          =   360
         Left            =   6960
         TabIndex        =   8
         Top             =   3100
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
         Caption         =   "&Back"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdPrevious 
         Height          =   360
         Left            =   5400
         TabIndex        =   39
         Top             =   3105
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
         Caption         =   "&Add Row"
      End
      Begin VB.Menu subMenuDelete 
         Caption         =   "&Delete Row"
      End
   End
End
Attribute VB_Name = "frmRecap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrFlex As String
Dim bol1stTab As Boolean
Dim bol2ndTab As Boolean
Dim bol3rdTab As Boolean
Dim bol4thTab As Boolean
Dim bol5thTab As Boolean
Dim bol6thTab As Boolean
Dim bol7thTab As Boolean
Dim mstrEmailLines As String
Dim mbolClickBelowRow As Boolean
Dim promptReminderBefore As Boolean
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date
'Dim mbolBack As Boolean
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
Dim colFFValue As Collection

Private Sub cboCCCode_Change(Index As Integer)
    If cboCCCode(Index).Text = "INVAGT" Or cboCCCode(Index).Text = "" Then
       txtCCNum(Index).Visible = False
       dtpCCExpire(Index).Visible = False
    Else
       txtCCNum(Index).Visible = True
       dtpCCExpire(Index).Visible = True
    End If
End Sub

Private Sub cboCCCode_Click(Index As Integer)
    cboCCCode_Change (Index)
End Sub

Private Sub cboCCCode_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii)
End Sub

Private Sub chkEntry_Click()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
    ElseIf mstrFlex = msFlexFFlyer.Name Then
       Set msFlex = msFlexFFlyer
    End If
    checkedRow msFlex
    
End Sub

Private Sub chkEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 1 Then msFlexEmails.SetFocus
    ElseIf mstrFlex = msFlexFFlyer.Name Then
       Set msFlex = msFlexFFlyer
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 4 Then msFlexFFlyer.SetFocus
    End If
    control_LostFocus msFlex, Me, chkEntry
    cmbEntry.style = fmStyleDropDownCombo
End Sub

Private Sub cmbEntry_GotFocus()
    cmbGetFocus cmbEntry
End Sub

Private Sub cmbEntry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If mstrFlex = msFlexEmails.Name Or mstrFlex = msFlexFFlyer.Name Then
       control_KeyDown CInt(KeyCode), Shift, Me, cmbEntry.Container
    Else
       control_KeyDown CInt(KeyCode), Shift, Me, cmbEntry.Container, False
    End If
End Sub

Private Sub cmbEntry_KeyPress(KeyCode As MSForms.ReturnInteger)
    KeyCode = Asc(UCase(Chr(CInt(KeyCode))))
End Sub

Private Sub cmdCancel_Click()
    gbolCancelProcess = True
    Unload Me
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
        'Back up on 26 Sept - Jeremy
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModRecap), _
'       IIf(gbolCreatPNR = True, gconSModRecap, ""), Me.Name, gconFormLoad, gstrProcessGrpID, _
'       datFormLoadEnd, datFormLoadStart
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModRecap), _
'       IIf(gbolCreatPNR = True, gconSModRecap, ""), Me.Name, gconTouch, gstrProcessGrpID, _
'       datTouchEnd, datFormLoadEnd
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModRecap), _
'       IIf(gbolCreatPNR = True, gconSModRecap, ""), Me.Name, gconProcessing, gstrProcessGrpID, _
'        , datTouchEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModRecap, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModRecap, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModRecap, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd

       
       Unload Me
    End If
End Sub

Private Sub cmdNext_Click()
    If SftTabs.Tabs.Current + 1 < SftTabs.Tabs.Count Then
       If SftTabs.Tab(SftTabs.Tabs.Current + 1).Enabled = True Then
          SftTabs.Tabs.Current = SftTabs.Tabs.Current + 1
       Else
          If SftTabs.Tabs.Current + 2 < SftTabs.Tabs.Count Then
             If SftTabs.Tab(SftTabs.Tabs.Current + 2).Enabled = True Then
                SftTabs.Tabs.Current = SftTabs.Tabs.Current + 2
             End If
           End If
       End If
    End If
End Sub

Private Sub cmbEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
    cmbEntry.style = fmStyleDropDownCombo
    cmbEntry.ColumnWidths = "100 pt;0 pt"
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 1 Then msFlexEmails.SetFocus
    ElseIf mstrFlex = msFlexPassport.Name Then
       Set msFlex = msFlexPassport
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 2 Then msFlexPassport.SetFocus
       control_LostFocus msFlex, Me, cmbEntry, "V"
       Exit Sub
    ElseIf mstrFlex = msFlexFFlyer.Name Then
       Set msFlex = msFlexFFlyer
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 4 Then msFlexFFlyer.SetFocus
    ElseIf mstrFlex = msFlexPretripMI.Name Then
       Set msFlex = msFlexPretripMI
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 5 Then msFlexPretripMI.SetFocus
       control_LostFocus msFlex, Me, cmbEntry, "V"
       Exit Sub
    ElseIf mstrFlex = msFlexPostTripMI.Name Then
       Set msFlex = msFlexPostTripMI
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 6 Then msFlexPostTripMI.SetFocus
       control_LostFocus msFlex, Me, cmbEntry, "V"
       Exit Sub
    End If
    control_LostFocus msFlex, Me, cmbEntry
        
End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload Me
End Sub

Private Sub dtpEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, dtpEntry, False
End Sub

Private Sub dtpEntry_LostFocus()
  Dim msFlex As MSFlexGrid
  
  If mstrFlex = msFlexPassport.Name Then
     Set msFlex = msFlexPassport
     'Clement - 20080812
     'If SftTabs.Tabs.Current = 2 Then msFlexPassport.SetFocus
     control_LostFocus msFlex, Me, dtpEntry, "V"
     Exit Sub
  End If
End Sub

Private Sub Form_Load()
   Dim oldParent As Long
   Dim i As Integer
   
  'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
   Dim strFF As String
   Dim intC As Integer
   
   datFormLoadStart = Now

    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   
   'mbolBack = gbolBack
   
   pDisplayToFP "*R"
   'Set the header caption and width of msFlexContacts
   With msFlexContacts
        .col = 0
        For i = 1 To .rows - 1
            .row = i
            If i = 1 Then
               .Text = " Business *"
            ElseIf i = 2 Then
               .Text = " Mobile *"
            ElseIf i = 3 Then
               .Text = " Home"
            ElseIf i = 4 Then
               .Text = " Fax"
            End If
        Next
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 1 Then
               .Text = " Contact 1"
               .ColWidth(i) = 1200
            ElseIf i = 2 Then
               .Text = " Contact 2"
               .ColWidth(i) = 1200
            ElseIf i = 3 Then
               .Text = " Contact 3"
               .ColWidth(i) = 1200
            ElseIf i = 4 Then
               .Text = ""
               .ColWidth(i) = 0
            ElseIf i = 5 Then
               .Text = ""
               .ColWidth(i) = 0
            End If
            .ColAlignment(i) = 0
        Next
        .row = 1
        .col = 1
   End With
      
   'Set the header caption and width of msFlexEmails
   With msFlexEmails
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
               .ColWidth(i) = 300
            ElseIf i = 1 Then
              .Text = " E-Itin"
              .ColWidth(i) = 600
            ElseIf i = 2 Then
              .Text = " E-Tkt"
              .ColWidth(i) = 600
            ElseIf i = 3 Then
              .Text = " E-Inv"
              .ColWidth(i) = 600
            ElseIf i = 4 Then
               .Text = " Type"
               .ColWidth(i) = 1000
            ElseIf i = 5 Then
              .Text = " Email"
              .ColWidth(i) = 4500
            ElseIf i = 6 Then
               .Text = "" 'Store ITI, INV & TKT line number
               .ColWidth(i) = 0
            ElseIf i = 7 Then
               .Text = "" 'Store Phone line number
               .ColWidth(i) = 0
            End If
            .ColAlignment(i) = 0
        Next
        setText msFlexEmails, 0, 0, 0
        setText msFlexEmails, 1, 0, 3
        .row = 1
        .col = 0
   End With
   
   'Set the header caption and width of msFlexPassport
   With msFlexPassport
        For i = 0 To .Cols - 1
            If i = 0 Then
               .ColWidth(i) = 1200
            ElseIf i = 1 Then
               .ColWidth(i) = 3000
            End If
            .ColAlignment(i) = 0
        Next
        .col = 0
        For i = 1 To .rows - 1
            .row = i
            If i = 1 Then
               .Text = " Passport Num "
            ElseIf i = 2 Then
                .Text = " Expiry Date "
            ElseIf i = 3 Then
                .Text = " Birth Date"
            ElseIf i = 4 Then
                .Text = " Issue Country "
            ElseIf i = 5 Then
                .Text = " Nationality *"
            ElseIf i = 6 Then
                .Text = " Citizenship"
            ElseIf i = 7 Then
                .RowHeight(7) = 0
            End If
        Next
   End With
   
   'Set the header caption and width of msFlexFFlyer
   With msFlexFFlyer
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
               .ColWidth(i) = 300
            ElseIf i = 1 Then
               .Text = " Name"
               .ColWidth(i) = 4000
            ElseIf i = 2 Then
              .Text = " Vendor"
              .ColWidth(i) = 800
            ElseIf i = 3 Then
              .Text = " Frequent Flyer Number"
              .ColWidth(i) = 1800
            ElseIf i = 4 Then
               .Text = " Cross Accrual (Example TG/SQ)"
               .ColWidth(i) = 3200
            End If
            .ColAlignment(i) = 0
        Next
        setText msFlexFFlyer, 0, 0, 0
        setText msFlexFFlyer, 1, 0, 0
        .row = 1
        .col = 0
   End With
   
   
   'Set the header caption and width of msFlexPretripMI and msFlexPosttripMI
   'Hide Pretrip because Pretrip move to after Fare Quote
   'setMIHeader msFlexPretripMI
   setMIHeader msFlexPostTripMI
   
   'Populate Preset Controls
   getFOP
   mstrEmailLines = ""
   bol1stTab = True
   bol2ndTab = False
   bol3rdTab = False
   bol4thTab = False
   bol5thTab = False
   bol6thTab = False
   bol7thTab = False
   
   populatePsgr    'from Passenger Name
   populateAirFOP  'from FOP
   populatePhones  'from Email and Phone Field
   populateFromNP
   populateFromRI
   populateAddr
   populateFFlyer
   'Hide PreTrip because Pretrip move to after Fare Quote
   'getReportingField msFlexPretripMI, "3"    'Pretrip MI with location 3
   getReportingField msFlexPostTripMI, "1,2"   'PostTrip MI with location 1
   SftTabs.Tabs.Current = 0
   If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
   Else
      cmdPrevious.Visible = False
   End If
   cmdBack.Enabled = False
   If gbolBackToRecap Then
      cmdFinish.Enabled = True
   Else
      cmdFinish.Enabled = False
   End If
   promptReminderBefore = False
   
   datFormLoadEnd = Now
   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
   
  'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
  Set colFFValue = New Collection
    strFF = ""
    With msFlexPostTripMI
            For intC = 1 To .rows - 1
                If strFF = "" Then
                   strFF = "'" & Trim(.TextMatrix(intC, 2)) & "'"
                Else
                   strFF = strFF & ",'" & Trim(.TextMatrix(intC, 2)) & "'"
                End If
            Next
    End With
    
    'post trip MI location 1
    Set colFFValue = GetClientMIValue(gobjPNR.CN, strFF)
   
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       gbolCancelProcess = True
    End If
End Sub

Public Sub msFlexContacts_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
        
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
    
    intTop = sharedFra(0).Top + fraContacts.Top + SftTabs.Top + msFlexContacts.Top + msFlexContacts.CellTop
    intLeft = sharedFra(0).Left + fraContacts.Left + SftTabs.Left + msFlexContacts.Left + msFlexContacts.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    mstrFlex = msFlexContacts.Name
    setControlPosition msFlexContacts, txtEntry, intTop, intLeft
    
End Sub

Private Sub msFlexContacts_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, msFlexContacts, False
End Sub

Private Sub msFlexContacts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexContacts, vbLeftButton, Y
End Sub

Public Sub msFlexEmails_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
    
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
    intTop = sharedFra(1).Top + fraEmails.Top + SftTabs.Top + msFlexEmails.Top + msFlexEmails.CellTop
    intLeft = sharedFra(1).Left + fraEmails.Left + SftTabs.Left + msFlexEmails.Left + msFlexEmails.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    If chkEntry.Visible = True Then chkEntry.Visible = False
    
    mstrFlex = msFlexEmails.Name
    
    If msFlexEmails.col = 0 And msFlexEmails.row = 0 Then
       optSelectAll msFlexEmails
    ElseIf msFlexEmails.col <= 3 And msFlexEmails.row > 0 Then
       setControlPosition msFlexEmails, chkEntry, intTop, intLeft
       checkedRow msFlexEmails
    ElseIf msFlexEmails.col = 4 And msFlexEmails.row > 0 Then
       cmbEntry.Clear
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
    ElseIf msFlexEmails.col = 5 And msFlexEmails.row > 0 Then
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

Public Sub msFlexFFlyer_Click()
  Dim rsRecord As ADODB.Recordset
  Dim strSql As String
  Dim intTop As Integer
  Dim intLeft As Integer
        
  If mbolClickBelowRow = True Then
     mbolClickBelowRow = False
     Exit Sub
  End If
  intTop = sharedFra(4).Top + fraFFlyer.Top + SftTabs.Top + msFlexFFlyer.Top + msFlexFFlyer.CellTop
  intLeft = sharedFra(4).Left + fraFFlyer.Left + SftTabs.Left + msFlexFFlyer.Left + msFlexFFlyer.CellLeft
  If txtEntry.Visible = True Then txtEntry.Visible = False
  If cmbContainer.Visible = True Then cmbContainer.Visible = False
  If dtpEntry.Visible = True Then dtpEntry.Visible = False
  If chkEntry.Visible = True Then chkEntry.Visible = False
  
  mstrFlex = msFlexFFlyer.Name
  
  If msFlexFFlyer.col = 0 And msFlexFFlyer.row = 0 Then
     optSelectAll msFlexFFlyer
  ElseIf msFlexFFlyer.col = 0 And msFlexFFlyer.row > 0 Then
     setControlPosition msFlexFFlyer, chkEntry, intTop, intLeft
     checkedRow msFlexFFlyer
  ElseIf msFlexFFlyer.col <= 2 And msFlexFFlyer.row > 0 Then
    cmbEntry.Clear
    If msFlexFFlyer.col = 2 Then
       strSql = "Select code from tblAirVendors Where Type = 'AIR' order by code"
       Set rsRecord = gdbConn.Execute(strSql)
       While Not rsRecord.EOF
          cmbEntry.AddItem rsRecord!Code & ""
          rsRecord.MoveNext
       Wend
    ElseIf msFlexFFlyer.col = 1 Then
       For i = 1 To gobjPNR.PassengerCount
            With gobjPNR.PassengerName(i)
                 cmbEntry.AddItem Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
            End With
       Next i
    End If
    setControlPosition msFlexFFlyer, cmbContainer, intTop, intLeft, cmbEntry
  ElseIf msFlexFFlyer.col > 2 And msFlexFFlyer.row > 0 Then
    setControlPosition msFlexFFlyer, txtEntry, intTop, intLeft
  End If

End Sub

Private Sub msFlexFFlyer_KeyDown(KeyCode As Integer, Shift As Integer)
    mstrFlex = msFlexFFlyer.Name
    control_KeyDown KeyCode, Shift, Me, msFlexFFlyer
End Sub

Private Sub msFlexFFlyer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexFFlyer, Button, Y
End Sub

Public Sub msFlexPassport_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
    
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
        
    intTop = sharedFra(2).Top + fraPassport.Top + SftTabs.Top + msFlexPassport.Top + msFlexPassport.CellTop
    intLeft = sharedFra(2).Left + fraPassport.Left + SftTabs.Left + msFlexPassport.Left + msFlexPassport.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    mstrFlex = msFlexPassport.Name

    If msFlexPassport.row = 1 Then
       setControlPosition msFlexPassport, txtEntry, intTop, intLeft
    ElseIf msFlexPassport.row = 2 Or msFlexPassport.row = 3 Then
       setControlPosition msFlexPassport, dtpEntry, intTop, intLeft
    ElseIf msFlexPassport.row = 4 Or msFlexPassport.row = 5 Or msFlexPassport.row = 6 Then
       cmbEntry.Clear
       If msFlexPassport.row = 4 Or msFlexPassport.row = 5 Or msFlexPassport.row = 6 Then
           GetCountry
       End If
       setControlPosition msFlexPassport, cmbContainer, intTop, intLeft, cmbEntry
    End If
End Sub

Private Sub msFlexPassport_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, msFlexPassport, False
End Sub

Private Sub msFlexPassport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexPassport, vbLeftButton, Y
End Sub

Public Sub msFlexPostTripMI_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
    
    Dim intI As Integer
        
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
        
    intTop = sharedFra(6).Top + fraPostTripMI.Top + SftTabs.Top + msFlexPostTripMI.Top + msFlexPostTripMI.CellTop
    intLeft = sharedFra(6).Left + fraPostTripMI.Left + SftTabs.Left + msFlexPostTripMI.Left + msFlexPostTripMI.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    mstrFlex = msFlexPostTripMI.Name

    If msFlexPostTripMI.col = 1 Then
       cmbEntry.Clear
       'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
       cmbEntry.style = fmStyleDropDownCombo
'       If colFFValue.Count > 0 Then
'          cmbEntry.style = fmStyleDropDownList
'       Else
'          cmbEntry.style = fmStyleDropDownCombo
'       End If
       'If cmbEntry.style <> 2 Then
          cmbEntry.Text = ""
        'End If
      
        PopulatecmbMI cmbEntry, colFFValue, CStr(msFlexPostTripMI.TextMatrix(msFlexPostTripMI.row, 2)), msFlexPostTripMI.Text
       setControlPosition msFlexPostTripMI, cmbContainer, intTop, intLeft, cmbEntry
       
       If cmbEntry.ListCount > 0 Then
          cmbEntry.style = fmStyleDropDownList
       Else
          cmbEntry.style = fmStyleDropDownCombo
       End If
       
    End If
End Sub

Private Sub msFlexPostTripMI_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, msFlexPostTripMI, False
End Sub

Private Sub msFlexPostTripMI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexPostTripMI, vbLeftButton, Y
End Sub

Public Sub msFlexPretripMI_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
    
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
        
    intTop = sharedFra(5).Top + fraPretripMI.Top + SftTabs.Top + msFlexPretripMI.Top + msFlexPretripMI.CellTop
    intLeft = sharedFra(5).Left + fraPretripMI.Left + SftTabs.Left + msFlexPretripMI.Left + msFlexPretripMI.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    mstrFlex = msFlexPretripMI.Name
  
    If msFlexPretripMI.col = 1 Then
       cmbEntry.Clear
       cmbEntry.Text = ""
       setControlPosition msFlexPretripMI, cmbContainer, intTop, intLeft, cmbEntry
    End If
End Sub

Private Sub msFlexPretripMI_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, msFlexPretripMI, False
End Sub

Private Sub cmdBack_Click()
    If SftTabs.Tabs.Current - 1 >= 0 Then
       If SftTabs.Tab(SftTabs.Tabs.Current - 1).Enabled = True Then
          SftTabs.Tabs.Current = SftTabs.Tabs.Current - 1
       Else
           If SftTabs.Tabs.Current - 2 >= 0 Then
              If SftTabs.Tab(SftTabs.Tabs.Current - 2).Enabled = True Then
                 SftTabs.Tabs.Current = SftTabs.Tabs.Current - 2
              End If
           End If
       End If
    End If
End Sub

Private Sub msFlexPretripMI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexPretripMI, vbLeftButton, Y
End Sub

Private Sub SftTabs_Switching(NextTab As Integer, Allow As Boolean, Refresh As Boolean)
    
    Dim strMsg As String
    
    'Validate tab by tab
    If SftTabs.Tabs.Current = 0 Then
       strMsg = validateContact_FOP
    ElseIf SftTabs.Tabs.Current = 1 Then
      strMsg = validateEmail
    ElseIf SftTabs.Tabs.Current = 2 Then
      strMsg = validatePassport
    ElseIf SftTabs.Tabs.Current = 3 Then
      strMsg = validateAddress
    ElseIf SftTabs.Tabs.Current = 4 Then
      strMsg = validateFFlyer
    ElseIf SftTabs.Tabs.Current = 5 Then
      strMsg = validatePretripMI
    ElseIf SftTabs.Tabs.Current = 6 Then
      strMsg = validatePostTripMI
    End If

    If strMsg <> "" Then
       Allow = False
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Sub
    ElseIf strMsg = "" And SftTabs.Tabs.Current = 2 Then
        If promptReminderBefore = False Then
           promptReminderBefore = True
           With msFlexPassport
                'Prompt reminder for passport expiry date
                If Trim(.TextMatrix(2, 1)) <> "" Then
                   If gobjPNR.AirSegCount > 0 Then
                      If gobjPNR.AirSeg(gobjPNR.AirSegCount).DepartDateTime > dtMax Then
                         dtMax = gobjPNR.AirSeg(gobjPNR.AirSegCount).DepartDateTime
                      End If
                   End If
                   If gobjPNR.HotelSegCount > 0 Then
                      If gobjPNR.HotelSeg(gobjPNR.HotelSegCount).CheckInDate > dtMax Then
                         dtMax = gobjPNR.HotelSeg(gobjPNR.HotelSegCount).CheckInDate
                      End If
                   End If
                   If gobjPNR.CarSegCount > 0 Then
                      If gobjPNR.CarSeg(gobjPNR.CarSegCount).StartDtTime > dtMax Then
                         dtMax = gobjPNR.CarSeg(gobjPNR.CarSegCount).StartDtTime
                      End If
                   End If
                   If Now > dtMax Then dtMax = Now
                   
                   If DateDiff("d", DateAdd("m", 6, dtMax), CDate(Trim(.TextMatrix(2, 1)))) < 0 Then
                      strMsg = "Passport must be valid 6 months before travel ..." & Chr(13)
                   End If
                Else
                    strMsg = "No passport expiry date for passenger 1 ..." & Chr(13)
                End If
           End With
           If strMsg <> "" Then
                strMsg = strMsg & "Do you still want to continue?"
                modMsgBox.YESMsg = "Yes"
                modMsgBox.NOMsg = "No"
                If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop - Reminder") = vbNo Then
                   Allow = False
                   Exit Sub
                End If
           End If
        End If
    End If
    
    cmdBack.Enabled = True
    cmdNext.Enabled = True
    
    If NextTab = 0 Then
       bol1stTab = True
       cmdBack.Enabled = False
    ElseIf NextTab = 1 Then
       bol2ndTab = True
    ElseIf NextTab = 2 Then
       bol3rdTab = True
    ElseIf NextTab = 3 Then
       bol4thTab = True
    ElseIf NextTab = 4 Then
       bol5thTab = True
       If SftTabs.Tab(5).Enabled = False And SftTabs.Tab(6).Enabled = False Then
          cmdNext.Enabled = False
       End If
    ElseIf NextTab = 5 Then
       bol6thTab = True
       If SftTabs.Tab(6).Enabled = False Then
          cmdNext.Enabled = False
       End If
    ElseIf NextTab = 6 Then
       bol7thTab = True
       cmdNext.Enabled = False
    End If
    If gbolBackToRecap = False Then
        'If bol1stTab = True And bol2ndTab = True And bol3rdTab = True And bol4thTab = True And bol5thTab = True And bol6thTab = True And bol7thTab = True Then
        'bol6thTab is for PreTrip MI which is hidden so it will be False
        If bol1stTab = True And bol2ndTab = True And bol3rdTab = True And bol4thTab = True And bol5thTab = True And bol6thTab = False And bol7thTab = True Then
           cmdFinish.Enabled = True
        Else
           cmdFinish.Enabled = False
        End If
    End If
End Sub

Public Sub subMenuAdd_Click()
    
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    If chkEntry.value = True Then chkEntry.Visible = False
    
    If mstrFlex = msFlexEmails.Name Then
       msFlexEmails.rows = msFlexEmails.rows + 1
       setText msFlexEmails, msFlexEmails.rows - 1, 0, 3
    ElseIf mstrFlex = msFlexFFlyer.Name Then
       msFlexFFlyer.rows = msFlexFFlyer.rows + 1
       setText msFlexFFlyer, msFlexFFlyer.rows - 1, 0, 0
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
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    If chkEntry.value = True Then chkEntry.Visible = False
    
    If mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
    ElseIf mstrFlex = msFlexFFlyer.Name Then
       Set msFlex = msFlexFFlyer
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

Private Sub txtBillAddr_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCCNum_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii)
End Sub

Private Sub txtDeliveryAddr_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    If mstrFlex = msFlexEmails.Name Or mstrFlex = msFlexFFlyer.Name Then
       control_KeyDown KeyCode, Shift, Me, txtEntry
    Else
       control_KeyDown KeyCode, Shift, Me, txtEntry, False
    End If
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexContacts.Name Then
       Set msFlex = msFlexContacts
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 0 Then msFlexContacts.SetFocus
    ElseIf mstrFlex = msFlexEmails.Name Then
       Set msFlex = msFlexEmails
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 1 Then msFlexEmails.SetFocus
       control_LostFocus msFlex, Me, txtEntry, , False
       Exit Sub
    ElseIf mstrFlex = msFlexPassport.Name Then
       Set msFlex = msFlexPassport
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 2 Then msFlexPassport.SetFocus
       control_LostFocus msFlex, Me, txtEntry, "V"
       Exit Sub
    ElseIf mstrFlex = msFlexFFlyer.Name Then
       Set msFlex = msFlexFFlyer
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 4 Then msFlexFFlyer.SetFocus
       control_LostFocus msFlex, Me, txtEntry, , False
       Exit Sub
    End If
    control_LostFocus msFlex, Me, txtEntry
End Sub

Private Sub getFOP()
   Dim i As Integer
   
   For i = 0 To cboCCCode.Count - 1
       cboCCCode(i).Clear
       cboCCCode(i).AddItem ""
       If i = 0 Then
          'Only apply to Air FOP only
          cboCCCode(i).AddItem "INVAGT"
       End If
       cboCCCode(i).AddItem "AX"
       cboCCCode(i).AddItem "CA"
       cboCCCode(i).AddItem "VI"
       cboCCCode(i).AddItem "DC"
       cboCCCode(i).AddItem "JP"
       cboCCCode(i).AddItem "TP"
       cboCCCode(i).listindex = 0
       dtpCCExpire(i).value = Now
   Next

End Sub

Private Sub GetCountry()
   Dim rscountry As ADODB.Recordset
   Dim strSql As String
      
   strSql = "Select CountryName from tblCountryCodes " & _
            "order by CountryName"

   Set rscountry = gdbConn.Execute(strSql)
   
   Do Until rscountry.EOF
      cmbEntry.AddItem rscountry!CountryName
      rscountry.MoveNext
   Loop
   
   rscountry.Close
   Set rscountry = Nothing
   
End Sub

Private Sub populatePsgr()
    Dim i As Integer
    
    'Populate passengers from PNR
    For i = 1 To gobjPNR.PassengerCount
        If i = 1 Then
           msFlexPassport.ColWidth(i) = 3000
           msFlexPassport.TextMatrix(0, i) = gobjPNR.PassengerName(i).LastName & "/" & gobjPNR.PassengerName(i).FirstName
        Else
           Exit For
        End If
    Next
End Sub
Private Sub populateFFlyer()
    Dim i As Integer
    Dim intTemp As Integer
    
    'Populate frquent flyer numbers from PNR
    For i = 1 To gobjPNR.FreqCustCount
        With gobjPNR.FreqCustNumber(i)
            If i > msFlexFFlyer.rows - 1 Then
               msFlexFFlyer.rows = msFlexFFlyer.rows + 1
               setText msFlexFFlyer, msFlexFFlyer.rows - 1, 0, 0
            End If
            'Search Passenger Name
            intTemp = .PassengerNum
            For intC = 1 To gobjPNR.PassengerCount
                With gobjPNR.PassengerName(intC)
                     If .PassengerNum = intTemp Then
                        msFlexFFlyer.TextMatrix(msFlexFFlyer.rows - 1, 1) = Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
                        Exit For
                     End If
                End With
            Next intC
            msFlexFFlyer.TextMatrix(msFlexFFlyer.rows - 1, 2) = .Vendor
            msFlexFFlyer.TextMatrix(msFlexFFlyer.rows - 1, 3) = .FreqCustNum
            msFlexFFlyer.TextMatrix(msFlexFFlyer.rows - 1, 4) = .CrossAccrual
        End With
    Next
End Sub

Private Sub populateAirFOP()
    If gobjPNR.FOPType = "CC" Then
       cboCCCode(0).Text = gobjPNR.FOP_CCCode
       txtCCNum(0).Text = gobjPNR.FOP_CCNum
       dtpCCExpire(0).value = gobjPNR.FOP_CCExpireDate
    ElseIf gobjPNR.FOPType = "INV" Then
       cboCCCode(0).Text = "INVAGT"
    End If
End Sub

Private Sub populateFromNP()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intPsgr As Integer
    Dim strTemp As String
    Dim strArray() As String
    
    For i = 1 To gobjPNR.GeneralRemarkCount
        With gobjPNR.GeneralRemark(i)
            If .Qualifier = "*H" Or .Qualifier = "*C" Then
               If .Qualifier = "*H" Then
                  k = 1 'Hotel FOP
               ElseIf .Qualifier = "*C" Then
                  k = 2 'Car FOP
               End If
               
               j = InStr(1, .RemarkText, "CC GUARANTEE:")
               If j > 0 Then
                  cboCCCode(k).Tag = i 'Store NP line number for persoal credit card
                  strTemp = Trim(Replace(.RemarkText, "CC GUARANTEE:", ""))
                  cboCCCode(k).Text = Mid(strTemp, 1, 2)
                  strArray = Split(Mid(strTemp, 3), "EXP")
                  If UBound(strArray) = 1 Then
                     txtCCNum(k).Text = strArray(0)
                     If Len(Trim(strArray(1))) = 4 Or Len(Trim(strArray(1))) = 2 Then
                        strTemp = Mid(Trim(strArray(1)), 1, 4)
                        strTemp = MonthName(CLng(Left(strTemp, 2))) & "/25/" & IIf(Len(Trim(strArray(1))) = 2, Format(Now, "yy"), Right(strTemp, 2))
                        dtpCCExpire(k).value = DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", CDate(strTemp)) + 1, CDate(strTemp))))
                     End If
                  End If
               End If
            ElseIf .Qualifier = "*G" Then
                j = InStr(1, .RemarkText, "BDAY: ")
                If j > 0 Then
                   strTemp = .RemarkText
                   strTemp = Trim(Mid(strTemp, 7))
                   If IsDate(strTemp) Then
                      msFlexPassport.TextMatrix(3, 1) = Format(CDate(strTemp), "mm/dd/yyyy")
                      msFlexPassport.TextMatrix(7, 1) = msFlexPassport.TextMatrix(7, 1) & "BIRTH-" & i & "."
                   End If
                End If
            ElseIf .Qualifier = "*P" And InStr(1, .RemarkText, "PASSPORT NO: ") > 0 Then
                j = InStr(1, .RemarkText, "PASSPORT NO: ")
                If j > 0 Then
                   strTemp = .RemarkText
                   strArray = Split(strTemp, "-")
                   For j = 0 To UBound(strArray)
                       If InStr(strArray(j), "PASSPORT NO: ") > 0 Then 'Passport Num
                          strTemp = Trim(Replace(strArray(j), "PASSPORT NO: ", ""))
                          msFlexPassport.TextMatrix(1, 1) = strTemp
                       ElseIf Len(Trim(strArray(j))) = 2 Then 'Nationality
                          strTemp = Trim(strArray(j))
                          msFlexPassport.TextMatrix(5, 1) = GetCountryName(strTemp)
                       ElseIf InStr(strArray(j), "ISS") > 0 Then 'Issue Country
                          strTemp = Trim(Replace(strArray(j), "ISS", ""))
                          If Len(strTemp) = 2 Then
                             msFlexPassport.TextMatrix(4, 1) = GetCountryName(strTemp)
                          End If
                       ElseIf InStr(strArray(j), "EXP") > 0 Then 'Expiration Date
                          strTemp = Trim(Replace(strArray(j), "EXP", ""))
                          If Len(strTemp) = 7 Then
                             msFlexPassport.TextMatrix(2, 1) = Format(Mid(strTemp, 1, 2) & "-" & Mid(strTemp, 3, 3) & "-" & Right(strTemp, 2), "mm/dd/yyyy")
                          End If
                       End If
                   Next
                   msFlexPassport.TextMatrix(7, 1) = msFlexPassport.TextMatrix(7, 1) & "DETAILS-" & i & "."
                End If
            ElseIf .Qualifier = "*P" And InStr(1, .RemarkText, "CITIZENSHIP: ") > 0 Then
                j = InStr(1, .RemarkText, "CITIZENSHIP: ")
                If j > 0 Then
                   strTemp = .RemarkText
                   strTemp = Mid(strTemp, 14)
                   msFlexPassport.TextMatrix(6, 1) = GetCountryName(strTemp)
                   msFlexPassport.TextMatrix(7, 1) = msFlexPassport.TextMatrix(7, 1) & "CITIZEN-" & i & "."
                End If
            End If
        End With
   Next
   
End Sub

Private Sub populatePhones()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim phoneType As Integer
    Dim strTemp As String
    Dim bolInPhone As Boolean
    Dim strDocType As String
    
    For i = 1 To gobjPNR.PhoneCount
        With gobjPNR.Phone(i)
             phoneType = 0
             If .phoneType = "B" Then
                 phoneType = 1 'Business
                 strTemp = .PhoneNum
             ElseIf .phoneType = "R" Then
                 phoneType = 3 'Home
                 strTemp = .PhoneNum
             ElseIf InStr(1, .PhoneNum, "M*") Then
                 phoneType = 2 'Mobile
                 strTemp = Mid(.PhoneNum, InStr(1, .PhoneNum, "M*") + 2)
             ElseIf InStr(1, .PhoneNum, "F*") Then
                 phoneType = 4 'Fax
                 strTemp = Mid(.PhoneNum, InStr(1, .PhoneNum, "F*") + 2)
             ElseIf InStr(1, .PhoneNum, "E*") Then
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
             If phoneType > 0 And phoneType < 5 Then
                For j = 1 To msFlexContacts.Cols - 1
                    If Trim(msFlexContacts.TextMatrix(phoneType, j)) = "" Then
                          msFlexContacts.TextMatrix(phoneType, j) = strTemp
                          msFlexContacts.TextMatrix(phoneType, 4) = msFlexContacts.TextMatrix(phoneType, 4) & i & "."
                          If phoneType = 2 Then
                             For k = 1 To gobjPNR.OSICount
                                 With gobjPNR.OSI(k)
                                      If .Vendor = "YY" And Trim(.Text) = "CTCM " & msFlexContacts.TextMatrix(phoneType, j) Then
                                          msFlexContacts.TextMatrix(phoneType, 5) = msFlexContacts.TextMatrix(phoneType, 5) & k & "."
                                          Exit For
                                      End If
                                 End With
                             Next
                          End If
                       Exit For
                    End If
                Next
             ElseIf phoneType = 5 Then
                 If Trim(msFlexEmails.TextMatrix(1, 5)) <> "" Then
                    msFlexEmails.rows = msFlexEmails.rows + 1
                    msFlexEmails.row = msFlexEmails.rows - 1
                    setText msFlexEmails, msFlexEmails.row, 0, 3
                 Else
                    msFlexEmails.row = 1
                 End If
                 If InStr(1, .PhoneNum, "E*PAX-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 4) = "PAX"
                 ElseIf InStr(1, .PhoneNum, "E*PER-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 4) = "PER"
                 ElseIf InStr(1, .PhoneNum, "E*OTR-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 4) = "OTR"
                 ElseIf InStr(1, .PhoneNum, "E*ARG-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 4) = "ARG"
                 ElseIf InStr(1, .PhoneNum, "E*DOM-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 4) = "DOM"
                 ElseIf InStr(1, .PhoneNum, "E*INT-") > 0 Then
                    msFlexEmails.TextMatrix(msFlexEmails.row, 4) = "INT"
                 End If
                 'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
                 'SINE*ARG-CCHING//CARLSONWAGONLIT.COM ITI to CCHING@CARLSONWAGONLIT.COM
                 If gobjPNR.CompInfo.AquaItin Then
                    strDocType = UCase(Right(strTemp, 4))
                    If strDocType = " ITI" Or strDocType = " TKT" Or strDocType = " INV" Then
                        If Len(strTemp) - 4 > 0 Then
                            strTemp = Mid(strTemp, 1, Len(strTemp) - 4)
                        End If
                    End If
                    If strDocType = " ITI" Then
                        msFlexEmails.TextMatrix(msFlexEmails.row, 1) = gstrChecked
                    ElseIf strDocType = " TKT" Then
                        msFlexEmails.TextMatrix(msFlexEmails.row, 2) = gstrChecked
                    ElseIf strDocType = " INV" Then 'For future, INV is not using at the moment
                        msFlexEmails.TextMatrix(msFlexEmails.row, 3) = gstrChecked
                    End If
                 End If
                 msFlexEmails.TextMatrix(msFlexEmails.row, 5) = strTemp
                 msFlexEmails.TextMatrix(msFlexEmails.row, 7) = "PHONE-" & i
             End If
        End With
    Next
    
    
    For i = 1 To gobjPNR.ItinRemarkCount
    
    If (Left(gobjPNR.ItinRemark(i).RemarkText, 4) = "ITI." Or Left(gobjPNR.ItinRemark(i).RemarkText, 4) = "TKT." Or Left(gobjPNR.ItinRemark(i).RemarkText, 4) = "INV.") And _
         Len(gobjPNR.ItinRemark(i).RemarkText) > 4 Then
         strTemp = actualText(Mid(gobjPNR.ItinRemark(i).RemarkText, 5))
         bolInPhone = False
         For j = 1 To msFlexEmails.rows - 1
            If Trim(msFlexEmails.TextMatrix(j, 5)) = strTemp Then
               bolInPhone = True
            End If
         Next
    
    
    If bolInPhone = False Then
    If Trim(msFlexEmails.TextMatrix(1, 5)) <> "" Then
       msFlexEmails.rows = msFlexEmails.rows + 1
       msFlexEmails.row = msFlexEmails.rows - 1
       setText msFlexEmails, msFlexEmails.row, 0, 3
    Else
       msFlexEmails.row = 1
    End If
    msFlexEmails.TextMatrix(msFlexEmails.row, 5) = strTemp
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
            ElseIf Left(.RemarkText, 4) = "TKT." And _
               Len(.RemarkText) > 4 Then
               bolExist = dataExistInCell(msFlexEmails, actualText(Mid(.RemarkText, 5)))
               If bolExist = True Then
                  insertIntoEmailFlex bolExist, i, 2, gstrChecked, Mid(.RemarkText, 5)
               Else
                  mstrEmailLines = mstrEmailLines & i & "."
               End If
            ElseIf Left(.RemarkText, 5) = "TKTX." And _
               Len(.RemarkText) > 5 Then
               bolExist = dataExistInCell(msFlexEmails, actualText(Mid(.RemarkText, 6)))
               If bolExist = True Then
                  insertIntoEmailFlex bolExist, i, 2, gstrUnChecked, Mid(.RemarkText, 6)
               Else
                  mstrEmailLines = mstrEmailLines & i & "."
               End If
            ElseIf Left(.RemarkText, 4) = "INV." And _
               Len(.RemarkText) > 4 Then
               bolExist = dataExistInCell(msFlexEmails, actualText(Mid(.RemarkText, 5)))
               If bolExist = True Then
                  insertIntoEmailFlex bolExist, i, 3, gstrChecked, Mid(.RemarkText, 5)
               Else
                  mstrEmailLines = mstrEmailLines & i & "."
               End If
            ElseIf Left(.RemarkText, 5) = "INVX." And _
               Len(.RemarkText) > 5 Then
               bolExist = dataExistInCell(msFlexEmails, actualText(Mid(.RemarkText, 6)))
               If bolExist = True Then
                  insertIntoEmailFlex bolExist, i, 3, gstrUnChecked, Mid(.RemarkText, 6)
               Else
                  mstrEmailLines = mstrEmailLines & i & "."
               End If
            End If
        End With
   Next
End Sub

Private Function dataExistInCell(ByRef msFlex As MSFlexGrid, strValue As String) As Boolean
    Dim i As Integer
    
    For i = 1 To msFlex.rows - 1
        If UCase(Trim(msFlex.TextMatrix(i, 5))) = UCase(Trim(strValue)) Then
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
    ElseIf colYesNo = 2 Then
       strTemp = "TKT"
    ElseIf colYesNo = 3 Then
       strTemp = "INV"
    End If
    msFlexEmails.TextMatrix(msFlexEmails.row, colYesNo) = strChecked
    msFlexEmails.TextMatrix(msFlexEmails.row, 6) = msFlexEmails.TextMatrix(msFlexEmails.row, 6) & strTemp & "-" & intRI & "."

End Sub

Private Sub populateAddr()
   txtDeliveryAddr.Text = Replace(gobjPNR.DeliveryAddress, "@", vbCrLf)
   txtBillAddr.Text = Replace(gobjPNR.BillingAddress, "@", vbCrLf)
End Sub

Private Sub getReportingField(ByRef msFlex As MSFlexGrid, strLocation As String, Optional bolFromDB As Boolean)
   Dim rsMI As ADODB.Recordset
   Dim strSql As String
   Dim i As Integer
   Dim j As Integer
   Dim intLocation() As Integer
   Dim strTemp() As String
   
   strTemp = Split(strLocation, ",")
   
   For j = LBound(strTemp) To UBound(strTemp)
    ReDim Preserve intLocation(j)
    intLocation(j) = CInt(strTemp(j))
   Next
   
   strSql = "Select a.*, b.Format from tblClientMI a, tblMICategory b Where a.CN='" & gobjPNR.CN & "' AND a.location = b.code AND "
   
   strSql = strSql & "("
   For j = LBound(intLocation) To UBound(intLocation)
   strSql = strSql & IIf(j > 0, " or ", "") & "a.location='" & intLocation(j) & "'"
   Next
   strSql = strSql & ")"
   strSql = strSql & " Order by FF"
   Set rsMI = gdbConn.Execute(strSql)
   
   If rsMI.EOF = True Then
      If msFlex.Name = "msFlexPretripMI" Then
         bol6thTab = True
         SftTabs.Tab(5).Enabled = False
      ElseIf msFlex.Name = "msFlexPostTripMI" Then
         bol7thTab = True
         SftTabs.Tab(6).Enabled = False
      End If
   End If
   Do Until rsMI.EOF
      i = i + 1
      If i = 1 Then
         msFlex.row = 1
      Else
         msFlex.rows = msFlex.rows + 1
         msFlex.row = msFlex.rows - 1
      End If
      msFlex.TextMatrix(msFlex.row, 2) = rsMI!FF & ""
      msFlex.TextMatrix(msFlex.row, 0) = rsMI!Description & ""
      msFlex.TextMatrix(msFlex.row, 1) = IIf(bolFromDB = False, pGetMIFromGDS(rsMI!FF & "", , rsMI!Format & ""), "")
      msFlex.TextMatrix(msFlex.row, 3) = rsMI!dataType & ""
      msFlex.TextMatrix(msFlex.row, 4) = rsMI!length & ""
      rsMI.MoveNext
   Loop
   rsMI.Close
   Set rsMI = Nothing
End Sub

Private Sub writeDatatoGDS()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim strTemp As String
    Dim strTemp2 As String
    Dim strTemp3 As String
    Dim strCmp As String
    Dim strCmd As String
    Dim strCmd2 As String
    Dim strCmd3 As String
    Dim strResponse As String
    Dim strDelNP As String
    Dim strField() As String
    Dim strField2() As String
    Dim bolAdd As Boolean
    Dim bolUpdate As Boolean
    Dim bolDelete As Boolean
    Dim bolExist As Boolean
    Dim strMsg As String
    
    strDelNP = ""
    strCmd = ""
    strCmd2 = ""
    strCmd3 = ""
    
    If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
    gbolWritingtoPNR = True
    
    'FOP Fields (Air, Hotel & Car)
    If cboCCCode(0).Visible = True Then
        If cboCCCode(0).Text <> "INVAGT" Then
            strTemp = Trim(cboCCCode(0)) & Trim(txtCCNum(0)) & "/D" & Format(dtpCCExpire(0), "mmyy")
            If gobjPNR.FOPType <> "" Then
               strCmp = gobjPNR.FOP_CCCode & gobjPNR.FOP_CCNum & "/D" & gobjPNR.FOP_CCExpireDate
               If UCase(Trim(strTemp)) <> UCase(Trim(strCmp)) Then
                  strCmd = strCmd & IIf(strCmd = "", "", "+") & "F.@" & strTemp
               End If
            Else
               strCmd = strCmd & IIf(strCmd = "", "", "+") & "F." & strTemp
            End If
        ElseIf cboCCCode(0).Text = "INVAGT" Then
            If gobjPNR.FOPType <> "INV" Then
               If gobjPNR.FOPType <> "" Then
                  strCmd = strCmd & IIf(strCmd = "", "", "+") & "F.@INVAGT"
               Else
                  strCmd = strCmd & IIf(strCmd = "", "", "+") & "F.INVAGT"
               End If
             End If
        End If
    End If

    For j = 1 To cboCCCode.Count - 1
        If j = 1 Then
           strTemp2 = "H*" 'Qualifier for hotel
        ElseIf j = 2 Then
           strTemp2 = "C*" 'Qualifier for car
        End If
        If cboCCCode(j).Text <> "INVAGT" And cboCCCode(j).Text <> "" Then
           strTemp = "CC GUARANTEE: " & Trim(cboCCCode(j)) & Trim(txtCCNum(j)) & " EXP " & Format(dtpCCExpire(j), "mmyy")
           If cboCCCode(j).Tag <> "" Then
              i = cboCCCode(j).Tag
              strCmp = gobjPNR.GeneralRemark(i).RemarkText
              If UCase(Trim(strCmp)) <> UCase(Trim(strTemp)) Then
                 strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP." & i & "@" & strTemp2 & strTemp
              End If
           Else
              strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP." & strTemp2 & strTemp
           End If
        ElseIf cboCCCode(j).Text = "" Then
           If cboCCCode(j).Tag <> "" Then
              i = cboCCCode(j).Tag
              strDelNP = strDelNP & IIf(strDelNP = "", "", ".") & i
           End If
        End If
    Next
            
    'Email Fields (RI & P.)
    With msFlexEmails
        For i = 1 To .rows - 1
             If Len(.TextMatrix(i, 6)) >= 2 Then
                If Right(.TextMatrix(i, 6), 1) = "." Then .TextMatrix(i, 6) = Mid(.TextMatrix(i, 6), 1, Len(.TextMatrix(i, 6)) - 1)
             End If
             .TextMatrix(i, 7) = Replace(.TextMatrix(i, 7), "PHONE-", "")
             If .TextMatrix(i, 7) = "" Then
                strCmd = strCmd & IIf(strCmd = "", "", "+") & "P." & gstrAgcyCityCode & "E*" & .TextMatrix(i, 4) & "-" & convertPhoneText(.TextMatrix(i, 5))
             Else
                strTemp = gstrAgcyCityCode & "E*" & .TextMatrix(i, 4) & "-" & convertPhoneText(.TextMatrix(i, 5))
                strTemp2 = gobjPNR.Phone(CInt(.TextMatrix(i, 7))).PhoneNum
                If UCase(strTemp) <> UCase(strTemp2) Then
                   strCmd = strCmd & IIf(strCmd = "", "", "+") & "P." & .TextMatrix(i, 7) & "@" & strTemp
                End If
             End If
             strField = Split(.TextMatrix(i, 6), ".")
             strTemp = Trim(.TextMatrix(i, 5))
             If strTemp <> "" Then
                For j = 1 To msFlexEmails.Cols - 5
                    bolExist = False
                    strTemp = convertText(Trim(.TextMatrix(i, 5)))
                    strTemp2 = Trim(.TextMatrix(i, j))
                    strTemp3 = ""
                    If j = 1 Then
                       strTemp3 = "ITI"
                    ElseIf j = 2 Then
                       strTemp3 = "TKT"
                    ElseIf j = 3 Then
                       strTemp3 = "INV"
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
    
    If strCmd2 <> "" Or mstrEmailLines <> "" Then
       'Delete existing RI that stored in strCmd2
       strCmd2 = mstrEmailLines & strCmd2
       If Len(strCmd2) >= 2 Then
          If Right(strCmd2, 1) = "." Then strCmd2 = Mid(strCmd2, 1, Len(strCmd2) - 1)
       End If
       If strCmd2 <> "" Then
          strCmd2 = "RI." & sortInt(strCmd2) & "@"
          strCmd = strCmd & IIf(strCmd = "", "", "+") & strCmd2
          strCmd2 = ""
       End If
    End If
        
    'Phone fields (Contacts)
    With msFlexContacts
         For i = 1 To .rows - 1
             If Len(.TextMatrix(i, 4)) >= 2 Then
                If Right(.TextMatrix(i, 4), 1) = "." Then .TextMatrix(i, 4) = Mid(.TextMatrix(i, 4), 1, Len(.TextMatrix(i, 4)) - 1)
             End If
             If i = 2 Then
                'Mobile in SI field
                If Len(.TextMatrix(i, 5)) >= 2 Then
                   If Right(.TextMatrix(i, 5), 1) = "." Then .TextMatrix(i, 5) = Mid(.TextMatrix(i, 5), 1, Len(.TextMatrix(i, 5)) - 1)
                End If
                strField2 = Split(.TextMatrix(i, 5), ".")
             End If

             strField = Split(.TextMatrix(i, 4), ".")
             For j = 1 To msFlexContacts.Cols - 3
                 bolAdd = False
                 bolUpdate = False
                 bolDelete = False
                 strTemp = Trim(.TextMatrix(i, j))
                 strTemp2 = ""
                 
                 If UBound(strField) >= j - 1 Then
                    If strTemp = "" Then
                       bolDelete = True
                    Else
                       If i = 1 Or i = 3 Then
                          strTemp2 = gobjPNR.Phone(CInt(strField(j - 1))).PhoneNum
                       ElseIf i = 2 Then
                          With gobjPNR.Phone(CInt(strField(j - 1)))
                               strTemp2 = Mid(.PhoneNum, InStr(1, .PhoneNum, "M*") + 2)
                          End With
                       ElseIf i = 4 Then
                          With gobjPNR.Phone(CInt(strField(j - 1)))
                               strTemp2 = Mid(.PhoneNum, InStr(1, .PhoneNum, "F*") + 2)
                          End With
                       End If
                       If UCase(strTemp) <> UCase(strTemp2) Then bolUpdate = True
                    End If
                 Else
                    If strTemp <> "" Then bolAdd = True
                 End If
                 
                 If i = 2 Then
                    'Mobile in SI
                    If UBound(strField2) >= j - 1 Then
                       If strTemp = "" Then
                          'Stored the SI fields to be deleted
                          strCmd2 = strCmd2 & IIf(strCmd2 = "", "", ".") & strField2(j - 1)
                       Else
                          If Trim(gobjPNR.OSI(CInt(strField2(j - 1))).Text) <> "CTCM " & strTemp Then
                             strCmd = strCmd & IIf(strCmd = "", "", "+") & "SI." & strField2(j - 1) & "@" & "YY*CTCM " & strTemp
                          End If
                       End If
                    Else
                       If strTemp <> "" Then
                          strCmd = strCmd & IIf(strCmd = "", "", "+") & "SI.YY*CTCM " & strTemp
                       End If
                    End If
                 End If
                 If bolAdd = True Or bolUpdate = True Then
                    If i = 1 Then
                       strTemp = gstrAgcyCityCode & "B" & "*" & strTemp
                    ElseIf i = 2 Then
                       strTemp = gstrAgcyCityCode & "M" & "*" & strTemp
                    ElseIf i = 3 Then
                       strTemp = gstrAgcyCityCode & "H" & "*" & strTemp
                    ElseIf i = 4 Then
                       strTemp = gstrAgcyCityCode & "F" & "*" & strTemp
                    End If
                    If bolAdd = True Then
                       strCmd = strCmd & IIf(strCmd = "", "", "+") & "P." & strTemp
                    ElseIf bolUpdate = True Then
                       strCmd = strCmd & IIf(strCmd = "", "", "+") & "P." & strField(j - 1) & "@" & strTemp
                    End If
                    
                 ElseIf bolDelete = True Then
                    strCmd3 = strCmd3 & IIf(strCmd3 = "", "", ".") & strField(j - 1)
                 End If
             Next
         Next
    End With
    
    If strCmd2 <> "" Or strCmd3 <> "" Then
       'Delete SI Field
       If strCmd2 <> "" Then
          strCmd2 = "SI." & sortInt(strCmd2) & "@"
          strCmd = strCmd & IIf(strCmd = "", "", "+") & strCmd2
          strCmd2 = ""
       End If
       'Delete Phone Field
       If strCmd3 <> "" Then
          strCmd3 = "P." & strCmd3 & "@"
          strCmd = strCmd & IIf(strCmd = "", "", "+") & strCmd3
          strCmd3 = ""
       End If
    End If
        
    'Address fields
    strTemp = Replace(Trim(txtDeliveryAddr), vbCrLf, "*")
    If strTemp <> "" Then
       strDeliveryAddress = strTemp
       If gobjPNR.DeliveryAddress <> "" Then
          strCmp = Replace(gobjPNR.DeliveryAddress, "@", "*")
          If UCase(Trim(strTemp)) <> UCase(Trim(strCmp)) Then
              strCmd = strCmd & IIf(strCmd = "", "", "+") & "D.@" & strTemp
          End If
       Else
          strCmd = strCmd & IIf(strCmd = "", "", "+") & "D." & strTemp
       End If
    Else
       If gobjPNR.DeliveryAddress <> "" Then
          strCmd = strCmd & IIf(strCmd = "", "", "+") & "D.@"
       End If
    End If
    
    strTemp = Replace(Trim(txtBillAddr), vbCrLf, "*")
    If strTemp <> "" Then
       strBillingAddress = strTemp
       If gobjPNR.BillingAddress <> "" Then
          strCmp = Replace(gobjPNR.BillingAddress, "@", "*")
          If UCase(Trim(strTemp)) <> UCase(Trim(strCmp)) Then
              strCmd = strCmd & IIf(strCmd = "", "", "+") & "W.@" & strTemp
          End If
       Else
          strCmd = strCmd & IIf(strCmd = "", "", "+") & "W." & strTemp
       End If
    Else
       If gobjPNR.BillingAddress <> "" Then
          strCmd = strCmd & IIf(strCmd = "", "", "+") & "W.@"
       End If
    End If
           
    With msFlexPassport
         For i = 1 To .Cols - 1
             If i = 1 Then
                If Len(.TextMatrix(7, i)) >= 2 Then
                   If Right(.TextMatrix(7, i), 1) = "." Then .TextMatrix(7, 1) = Mid(.TextMatrix(7, i), 1, Len(.TextMatrix(7, i)) - 1)
                End If
                strField = Split(.TextMatrix(7, i), ".")
                
                'Birth date
                strTemp = "BDAY: " & Format(Trim(.TextMatrix(3, i)), "mm/dd/yy")
                strTemp2 = "BIRTH"
                bolExist = False
                If strTemp <> "BDAY: " Then
                   For j = 0 To UBound(strField)
                       k = InStr(1, strField(j), strTemp2)
                       If k > 0 Then
                          bolExist = True
                          With gobjPNR.GeneralRemark(CInt(Mid(strField(j), Len(strTemp2) + 2)))
                               If Trim(UCase(strTemp)) <> Trim(UCase(.RemarkText)) Then
                                  strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP." & Mid(strField(j), Len(strTemp2) + 2) & "@G*" & strTemp
                               End If
                          End With
                          Exit For
                        End If
                   Next
                   If bolExist = False Then
                      strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP.G*" & strTemp
                   End If
                Else
                   For j = 0 To UBound(strField)
                       k = InStr(1, strField(j), strTemp2)
                       If k > 0 Then
                          strDelNP = strDelNP & IIf(strDelNP = "", "", ".") & Mid(strField(j), Len(strTemp2) + 2)
                          Exit For
                        End If
                   Next
                End If
                
                'Passport Details
                strTemp = "PASSPORT NO: " & Trim(.TextMatrix(1, i)) & "-" & GetCountryCode(Trim(.TextMatrix(5, i))) & _
                          "-ISS " & GetCountryCode(Trim(.TextMatrix(4, i))) & "-EXP " & Format(.TextMatrix(2, i), "ddmmmyy")
                strTemp2 = "DETAILS"
                bolExist = False
                If strTemp <> "" Then
                   For j = 0 To UBound(strField)
                       k = InStr(1, strField(j), strTemp2)
                       If k > 0 Then
                          bolExist = True
                          With gobjPNR.GeneralRemark(CInt(Mid(strField(j), Len(strTemp2) + 2)))
                               If Trim(UCase(strTemp)) <> Trim(UCase(.RemarkText)) Then
                                  strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP." & Mid(strField(j), Len(strTemp2) + 2) & "@P*" & strTemp
                               End If
                          End With
                          Exit For
                        End If
                   Next
                   If bolExist = False Then
                      strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP.P*" & strTemp
                   End If
                End If
                
                'Citizenship
                strTemp = "CITIZENSHIP: " & GetCountryCode(Trim(.TextMatrix(6, i)))
                strTemp2 = "CITIZEN"
                bolExist = False
                If strTemp <> "CITIZENSHIP: " Then
                   For j = 0 To UBound(strField)
                       k = InStr(1, strField(j), strTemp2)
                       If k > 0 Then
                          bolExist = True
                          With gobjPNR.GeneralRemark(CInt(Mid(strField(j), Len(strTemp2) + 2)))
                               If Trim(UCase(strTemp)) <> Trim(UCase(.RemarkText)) Then
                                  strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP." & Mid(strField(j), Len(strTemp2) + 2) & "@P*" & strTemp
                               End If
                          End With
                          Exit For
                        End If
                   Next
                   If bolExist = False Then
                      strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP.P*" & strTemp
                   End If
                Else
                   For j = 0 To UBound(strField)
                       k = InStr(1, strField(j), strTemp2)
                       If k > 0 Then
                          strDelNP = strDelNP & IIf(strDelNP = "", "", ".") & Mid(strField(j), Len(strTemp2) + 2)
                          Exit For
                        End If
                   Next
                End If
             End If
         Next
    End With
    
    If strDelNP <> "" Then
       If Len(strDelNP) >= 2 Then
          If Right(strDelNP, 1) = "." Then strDelNP = Mid(strDelNP, 1, Len(strDelNP) - 1)
       End If
       If strDelNP <> "" Then
         'Delete Notepad lines (CC FOP and also passport details)
         strDelNP = "NP." & sortInt(strDelNP) & "@"
         strCmd = strCmd & IIf(strCmd = "", "", "+") & strDelNP
         strDelNP = ""
       End If
    End If
        
    'Frequent Flyer
    If gobjPNR.FreqCustCount > 0 Then strCmd = strCmd & IIf(strCmd = "", "", "+") & "M.@"
    With msFlexFFlyer
         For i = 1 To .rows - 1
             If Trim(.TextMatrix(i, 3)) <> "" Then
                strTemp = Mid(Trim(.TextMatrix(i, 1)), 1, InStr(1, Trim(.TextMatrix(i, 1)), " "))
                strTemp = Trim(strTemp)
                For j = 1 To gobjPNR.PassengerCount
                    With gobjPNR.PassengerName(j)
                         If .GDSNum = strTemp Then
                            strCmd = strCmd & IIf(strCmd <> "", "+", "") & "M.P" & .PassengerNum
                            Exit For
                         End If
                    End With
                Next
                strCmd = strCmd & "/" & Trim(.TextMatrix(i, 2)) & Trim(.TextMatrix(i, 3)) & IIf(Trim(.TextMatrix(i, 4)) = "", "", "/" & Trim(.TextMatrix(i, 4)))
             End If
         Next
    End With
    
    'MI for Pre-Trip
    'Hide PreTrip because Pretrip is move after Fare Quote
'    strCmd2 = addMI(msFlexPretripMI)
'    If strCmd2 <> "" Then
'       strCmd = strCmd & IIf(strCmd <> "", "+", "") & strCmd2
'       strCmd2 = ""
'    End If
    
    'MI for Post-Trip
    strCmd2 = addMI(msFlexPostTripMI)
    If strCmd2 <> "" Then
       strCmd = strCmd & IIf(strCmd <> "", "+", "") & strCmd2
       strCmd2 = ""
    End If
    'Preethi - V1.2.4 20110614 - CR 76 - Change Validation Logic For ENDPNR
    'send entries, received & end the PNR
    If gbolCreatPNR = True And gobjPNR.RecLoc = "" Then
    Else
    strCmd = strCmd & IIf(strCmd = "", "", "+") & "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
    End If
    strCmd = strCmd & "+ER+ER+ER"
    strResponse = gobjHost.terminalEntry(strCmd)
    strTemp = strResponse
    'Preethi - V1.2.4 20110712 - CR 76 - Change Validation Logic For ENDPNR
    'If InStr(strTemp, "1.1") = 0 Then
    If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = False Then
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
    
    Exit Sub
    
errorWriting:
    'Prompt error message if failed to write to PNR
    gbolWritingtoPNR = False
    strMsg = "Unable to write to PNR. Response from GDS is " & Chr(13) & strResponse
    strMsg = strMsg & Chr(13) & "System will continue without ending this booking."
 
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
    'gobjHost.TerminalEntry "IR"
    
End Sub

Private Function addMI(ByRef msFlex As MSFlexGrid) As String
    Dim i As Integer
    Dim strFF As String
    Dim strFFValue As String
    Dim intDILine As Integer
    Dim bolFound As Boolean
    
    With msFlex
        For i = 1 To msFlex.rows - 1
            strFF = Trim(.TextMatrix(i, 2))
            If strFF <> "" Then
               If IsNumeric(strFF) Then
                  strFF = "FF" & strFF
               End If
               strFFValue = Trim(.TextMatrix(i, 1))
               intDILine = 0
               bolFound = False
               bolFound = updateMI(strFF, strFFValue, intDILine)
               If bolFound And intDILine > 0 Then
                  addMI = addMI & IIf(addMI = "", "", "+") & "DI." & intDILine & "@FT-" & strFF & "/" & strFFValue
               ElseIf bolFound = False And intDILine = 0 Then
                  addMI = addMI & IIf(addMI = "", "", "+") & "DI.FT-" & strFF & "/" & strFFValue
               End If
            End If
        Next
    End With
End Function

Private Sub setMIHeader(ByRef msFlex As MSFlexGrid)
   Dim i As Integer
   With msFlex
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
              .Text = " MIS Description"
              .ColWidth(i) = 3500
            ElseIf i = 1 Then
              .Text = " MIS Value"
              .ColWidth(i) = 2000
            ElseIf i = 2 Then
              .Text = " MIS Field"
              .ColWidth(i) = 1500
            ElseIf i = 3 Then
              .Text = " MIS Datatype"
              .ColWidth(i) = 1700
            ElseIf i = 4 Then
              .Text = " MIS Length"
              .ColWidth(i) = 1000
            End If
            .ColAlignment(i) = 0
        Next
   End With
End Sub

Private Function updateMI(strFF As String, strFFValue As String, intDILine As Integer) As Boolean
    Dim i As Integer
    Dim strTemp() As String
    
    For i = 1 To gobjPNR.AcctRemarkCount
        strTemp = Split(gobjPNR.AcctRemark(i).RemarkText, "/")
        If UBound(strTemp) >= 1 Then
           If UCase(Trim(strTemp(0))) = UCase(strFF) Then
              If UCase(Trim(strTemp(1))) <> UCase(strFFValue) Then
                 updateMI = True
                 intDILine = i
              Else
                 updateMI = False
                 intDILine = i
              End If
              Exit For
           End If
        End If
    Next
End Function

Private Function validData() As Boolean

    Dim strMsg As String
    
    validData = True
    
    'Validate all the tabs
    strMsg = validateContact_FOP
    strMsg = strMsg & validateEmail
    strMsg = strMsg & validatePassport
    strMsg = strMsg & validateFFlyer
    strMsg = strMsg & validatePretripMI
    strMsg = strMsg & validatePostTripMI
    
    If strMsg <> "" Then
       validData = False
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    End If
    
End Function

Private Function incompleteMI(ByRef msFlex As MSFlexGrid) As String
    Dim i As Integer
    Dim strMsg As String
        
    incompleteMI = ""
    With msFlex
         For i = 1 To msFlex.rows - 1
             If Trim(.TextMatrix(i, 1)) = "" Then
                strMsg = strMsg & "Incomplete MI for " & IIf(IsNumeric(.TextMatrix(i, 2)), "FF " & .TextMatrix(i, 2), .TextMatrix(i, 2)) & " ..." & Chr(13)
             Else
                If IsNumeric(.TextMatrix(i, 4)) Then
                    If Len(Trim(.TextMatrix(i, 1))) <> CInt(.TextMatrix(i, 4)) Then
                       strMsg = strMsg & "Incompatible length detected for " & IIf(IsNumeric(.TextMatrix(i, 2)), "FF " & .TextMatrix(i, 2), .TextMatrix(i, 2)) & " ..." & Chr(13)
                    Else
                       If UCase(.TextMatrix(i, 3)) = "NUMERIC" And Not IsNumeric(fConvertZero(.TextMatrix(i, 1))) Then
                          strMsg = strMsg & "Invalid data detected for " & IIf(IsNumeric(.TextMatrix(i, 2)), "FF " & .TextMatrix(i, 2), .TextMatrix(i, 2)) & " ..." & Chr(13)
                       End If
                    End If
                Else
                    If UCase(.TextMatrix(i, 3)) = "NUMERIC" And Not IsNumeric(fConvertZero(.TextMatrix(i, 1))) Then
                       strMsg = strMsg & "Invalid data detected for " & IIf(IsNumeric(.TextMatrix(i, 2)), "FF " & .TextMatrix(i, 2), .TextMatrix(i, 2)) & " ..." & Chr(13)
                    End If
                End If
             End If
         Next
    End With
    
    incompleteMI = strMsg
    
End Function

Private Sub deleteRow(ByRef msFlex As MSFlexGrid, ByVal i As Integer)
           
   Dim strTemp As String
       
   With msFlex
        If .Name = msFlexEmails.Name Then
            If .TextMatrix(i, 6) <> "" Then
                strTemp = Replace(.TextMatrix(i, 6), "ITI-", "")
                strTemp = Replace(strTemp, "INV-", "")
                strTemp = Replace(strTemp, "TKT-", "")
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

Private Sub mouseDown(ByRef msFlex As MSFlexGrid, ByVal Button As Integer, ByVal Y As Single)
         
    With msFlex
         If txtEntry.Visible = True Then txtEntry.Visible = False
         If cmbContainer.Visible = True Then cmbContainer.Visible = False
         If dtpEntry.Visible = True Then dtpEntry.Visible = False
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

Private Function validateContact_FOP() As String
        
    Dim strMsg As String
    Dim i As Integer
    
    If txtEntry.Visible Then
        txtEntry_LostFocus
    End If
    'Business & mobile number are mandatory fields
    With msFlexContacts
         If Trim(.TextMatrix(1, 1)) = "" And Trim(.TextMatrix(1, 2)) = "" And Trim(.TextMatrix(1, 3)) = "" Then
            strMsg = strMsg & "Business number is required ..." & Chr(13)
         End If
         If Trim(.TextMatrix(2, 1)) = "" And Trim(.TextMatrix(2, 2)) = "" And Trim(.TextMatrix(2, 3)) = "" Then
            strMsg = strMsg & "Mobile number is required ..." & Chr(13)
         End If
    End With
    
    'Validate FOP
    For i = 0 To cboCCCode.Count - 1
        If i = 0 Then
           If cboCCCode(i).Text = "" Then
              strMsg = strMsg & "Missing FOP for " & Replace(lblFOP(i), ":", "") & Chr(13)
           End If
        End If
        If cboCCCode(i).Text <> "INVAGT" And cboCCCode(i).Text <> "" Then
            If Trim(txtCCNum(i).Text) = "" Then
                strMsg = strMsg & "Missing " & Replace(lblFOP(i), ":", "") & " CC number..." & Chr(13)
            ElseIf ValidCCNum(cboCCCode(i).Text, txtCCNum(i).Text) = False Then
                strMsg = strMsg & "Invalid or incomplete " & Replace(lblFOP(i), ":", "") & " CC number..." & Chr(13)
            End If
            If dtpCCExpire(i).value < Now Then strMsg = strMsg & Replace(lblFOP(i), ":", "") & " CC has expired" & Chr(13)
        End If
    Next
    validateContact_FOP = strMsg
    
End Function

Private Function validateEmail() As String
    Dim strMsg As String
    Dim i As Integer
    
    If txtEntry.Visible Then
        txtEntry_LostFocus
    End If
    If cmbEntry.Visible Then
        cmbEntry_LostFocus
    End If
    
    'Email is mandatory field
    With msFlexEmails
         For i = 1 To .rows - 1
             If Trim(.TextMatrix(i, 4)) = "" Then
                 strMsg = strMsg & "Missing email type for record " & i & " ..." & Chr(13)
             End If
             If Trim(.TextMatrix(i, 5)) = "" Then
                 strMsg = strMsg & "Missing email address for record " & i & " ..." & Chr(13)
             End If
         Next
    End With
    validateEmail = strMsg

End Function

Private Function validatePassport() As String
    Dim strMsg As String
    Dim dtMax As Date
    
    If txtEntry.Visible Then
        txtEntry_LostFocus
    End If
    If cmbEntry.Visible Then
        cmbEntry_LostFocus
    End If
    
    'Validate Passport Details
    With msFlexPassport
         If Trim(.TextMatrix(2, 1)) <> "" Then
            If CDate(Trim(.TextMatrix(2, 1))) < Now Then
               strMsg = strMsg & "Invalid passport expiry date for passenger 1 ..." & Chr(13)
            End If
         End If
         If Trim(.TextMatrix(5, 1)) = "" Then
            strMsg = strMsg & "Missing nationality for passenger 1 ..." & Chr(13)
         End If
    End With
        
    validatePassport = strMsg
    
End Function

Private Function validateAddress() As String
    Dim strMsg As String
    Dim strTemp() As String
    Dim intLen As Integer
    
    intLen = 0
                
    'Validate delivery address (Max 6 subfields and max 37 characters in each subfield)
    If Trim(txtDeliveryAddr) <> "" Then
       Do While Len(txtDeliveryAddr) >= 2
          If Right(txtDeliveryAddr, 2) = vbCrLf Then
             If Len(txtDeliveryAddr) > 2 Then
                txtDeliveryAddr.Text = Mid(txtDeliveryAddr, 1, Len(txtDeliveryAddr) - 2)
             Else
                txtDeliveryAddr.Text = ""
             End If
          Else
             Exit Do
          End If
       Loop
    End If
    strTemp = Split(txtDeliveryAddr, vbCrLf)
    If UBound(strTemp) > 5 Then
       strMsg = strMsg & "Max 6 subfields in delivery address field ..." & Chr(13)
    Else
        For i = 0 To UBound(strTemp)
            If Len(strTemp(i)) > 37 Then
               strMsg = strMsg & "Only 37 characters are allowed in each delivery address subfield ..." & Chr(13)
               Exit For
            End If
        Next
    End If

    'Validate billing address (Max 5 subfields, 37 characters in each subfield, mandatory P/, 119 characters in entire field)
    If Trim(txtBillAddr) <> "" Then
       Do While Len(txtBillAddr) >= 2
          If Right(txtBillAddr, 2) = vbCrLf Then
             If Len(txtBillAddr) > 2 Then
                txtBillAddr.Text = Mid(txtBillAddr, 1, Len(txtBillAddr) - 2)
             Else
                txtBillAddr.Text = ""
             End If
          Else
             Exit Do
          End If
       Loop
       
       strTemp = Split(txtBillAddr, vbCrLf)
       If UBound(strTemp) > 4 Then
           strMsg = strMsg & "Max 5 subfields in billing address field ..." & Chr(13)
       Else
            For i = 0 To UBound(strTemp)
                intLen = intLen + Len(strTemp(i))
                If Len(strTemp(i)) > 37 Then
                   strMsg = strMsg & "Only 37 characters are allowed in each billing address subfield ..." & Chr(13)
                   Exit For
                End If
            Next
       End If
       If InStr(1, strMsg, "billing") = 0 Then
           If intLen > 119 Then
              strMsg = strMsg & "Only 119 characters are allowed in entire billing address field ..." & Chr(13)
           Else
              If InStr(1, txtBillAddr, "P/") = 0 Then
                 strMsg = strMsg & "Missing identifier P/ (Post Code) in billing address field ..." & Chr(13)
              End If
           End If
       End If
    End If
    validateAddress = strMsg

End Function

Private Function validateFFlyer() As String
    Dim i As Integer
    Dim strTemp() As String
    Dim strMsg As String
    
    If txtEntry.Visible Then
        txtEntry_LostFocus
    End If
    If cmbEntry.Visible Then
        cmbEntry_LostFocus
    End If
    
   'Validate frequent flyer fields
    With msFlexFFlyer
        For i = 1 To msFlexFFlyer.rows - 1
            If Trim(.TextMatrix(i, 1)) <> "" Or Trim(.TextMatrix(i, 2)) <> "" Or Trim(.TextMatrix(i, 3)) <> "" Then
               If Trim(.TextMatrix(i, 1)) = "" Then
                  strMsg = strMsg & "Missing passenger for record " & i & " (Frequent Flyer)" & Chr(13)
               End If
               If Trim(.TextMatrix(i, 2)) = "" Then
                  strMsg = strMsg & "Missing vendor for record " & i & " (Frequent Flyer)" & Chr(13)
               End If
               If Trim(.TextMatrix(i, 3)) = "" Then
                  strMsg = strMsg & "Missing frequent flyer number for record " & i & " (Frequent Flyer)" & Chr(13)
               End If
               If Trim(.TextMatrix(i, 4)) <> "" Then
                  'Validate cross accrual
                  If Right(.TextMatrix(i, 4), 1) = "/" Then .TextMatrix(i, 4) = Mid(.TextMatrix(i, 4), 1, Len(.TextMatrix(i, 4)) - 1)
                  If Trim(.TextMatrix(i, 4)) <> "" Then
                     strTemp = Split(Trim(.TextMatrix(i, 4)), "/")
                     For j = 0 To UBound(strTemp)
                         If Len(strTemp(j)) <> 2 Then
                            strMsg = strMsg & "Invalid cross accrual carrier " & strTemp(j) & " for record " & i & " (Frequent Flyer)" & Chr(13)
                         End If
                     Next
                  End If
               End If
            End If
        Next
    End With
    validateFFlyer = strMsg

End Function

Private Function validatePretripMI() As String
    Dim strMsg As String
        
    If cmbEntry.Visible Then
        cmbEntry_LostFocus
    End If
    
    'Validate incompleteMI
    If SftTabs.Tab(5).Enabled = True Then
       strMsg = strMsg & incompleteMI(msFlexPretripMI)
    End If
    validatePretripMI = strMsg
    
End Function

Private Function validatePostTripMI() As String
    Dim strMsg As String
    
    If cmbEntry.Visible Then
        cmbEntry_LostFocus
    End If
    
    'Validate incompleteMI
    If SftTabs.Tab(6).Enabled = True Then
       strMsg = strMsg & incompleteMI(msFlexPostTripMI)
    End If
    validatePostTripMI = strMsg

End Function



