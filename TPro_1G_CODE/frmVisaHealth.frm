VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{4832871B-0993-461C-B983-0EAAA4A43E5C}#5.0#0"; "SftTabs_IX86_U_50.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmVisaHealth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CWT Desktop - Timatic"
   ClientHeight    =   3420
   ClientLeft      =   1380
   ClientTop       =   2520
   ClientWidth     =   11370
   Icon            =   "frmVisaHealth.frx":0000
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
         Height          =   3060
         Left            =   120
         TabIndex        =   0
         Top             =   15
         Width           =   11205
         PropVer         =   50
         xcx             =   19764
         xcy             =   5398
         PropFile        =   ""
         PropDesignTime  =   1
         DeletePropFile  =   0
         IntVal          =   55
         xBfStyle1       =   63747965
         xBfStyle2       =   719519087
         xBfStyle3       =   672544242
         xBfStyle4       =   -672544242
         TabCount        =   2
         CurrentTab      =   0
         FlatProperties  =   0   'False
         BeginProperty Tab(0) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Visa && Health Info"
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
            Text            =   "Remarks"
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
         List(0)Count    =   0
         List(1)Count    =   1
         List(1)(0)Ctl   =   "sharedFra(1)"
         List(1)(0)Ena   =   -1
         List(1)(0)x     =   105
         List(1)(0)y     =   405
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
            Height          =   2580
            Index           =   1
            Left            =   -12040
            Top             =   -3985
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
            Begin VB.Frame Frame4 
               BackColor       =   &H00DADAB6&
               Caption         =   "Itinerary Remarks"
               Enabled         =   0   'False
               Height          =   2535
               Left            =   5400
               TabIndex        =   35
               Top             =   0
               Width           =   5415
               Begin VB.ListBox lstRmks 
                  Height          =   2200
                  Left            =   120
                  TabIndex        =   28
                  Top             =   240
                  Width           =   5175
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00DADAB6&
               Caption         =   "Preset Remarks"
               Enabled         =   0   'False
               Height          =   2500
               Left            =   120
               TabIndex        =   33
               Top             =   0
               Width           =   5295
               Begin MyCommandButton.MyButton cmdMoveAll 
                  Height          =   405
                  Left            =   4800
                  TabIndex        =   48
                  Top             =   1080
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   714
                  BackColor       =   15523541
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Small Fonts"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Picture         =   "frmVisaHealth.frx":038A
                  BackColorDown   =   15523541
                  BackColorOver   =   15523541
                  BackColorFocus  =   15523541
                  BackColorDisabled=   15523541
                  BorderColor     =   8540205
                  TransparentColor=   14215660
                  Caption         =   "ALL"
                  CaptionAlignment=   7
                  CaptionPosition =   3
                  DepthEvent      =   1
                  PictureDisabled =   "frmVisaHealth.frx":05FC
                  PictureAlignment=   1
                  ShowFocus       =   -1  'True
               End
               Begin MyCommandButton.MyButton cmdMoveRmk 
                  Height          =   405
                  Left            =   4800
                  TabIndex        =   47
                  Top             =   600
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   714
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
                  Picture         =   "frmVisaHealth.frx":08EE
                  BackColorDown   =   15523541
                  BackColorOver   =   15523541
                  BackColorFocus  =   15523541
                  BackColorDisabled=   15523541
                  BorderColor     =   8540205
                  TransparentColor=   14215660
                  Caption         =   ""
                  DepthEvent      =   1
                  PictureDisabled =   "frmVisaHealth.frx":0B60
                  PictureAlignment=   4
                  ShowFocus       =   -1  'True
               End
               Begin VB.ListBox lstPreset 
                  Height          =   1425
                  ItemData        =   "frmVisaHealth.frx":0E52
                  Left            =   120
                  List            =   "frmVisaHealth.frx":0E54
                  MultiSelect     =   2  'Extended
                  TabIndex        =   26
                  Top             =   600
                  Width           =   4575
               End
               Begin VB.ComboBox cmbCountry 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   25
                  Top             =   240
                  Width           =   2655
               End
               Begin VB.TextBox txtFreeText 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Left            =   120
                  TabIndex        =   27
                  Top             =   2100
                  Width           =   4575
               End
               Begin MyCommandButton.MyButton cmdAdd 
                  Height          =   405
                  Left            =   4800
                  TabIndex        =   49
                  Top             =   2040
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   714
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
                  Picture         =   "frmVisaHealth.frx":0E56
                  BackColorDown   =   15523541
                  BackColorOver   =   15523541
                  BackColorFocus  =   15523541
                  BackColorDisabled=   15523541
                  BorderColor     =   8540205
                  TransparentColor=   14215660
                  Caption         =   ""
                  DepthEvent      =   1
                  PictureDisabled =   "frmVisaHealth.frx":10C8
                  PictureAlignment=   4
                  ShowFocus       =   -1  'True
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Country:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   34
                  Top             =   240
                  Width           =   855
               End
            End
         End
         Begin MyFramePanel.MyFrame sharedFra 
            Height          =   2650
            Index           =   0
            Left            =   105
            Top             =   345
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4683
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
            Begin VB.Frame Frame2 
               BackColor       =   &H00DADAB6&
               Caption         =   "Visa && Health Info"
               Height          =   2475
               Left            =   4320
               TabIndex        =   36
               Top             =   60
               Width           =   6495
               Begin RichTextLib.RichTextBox txtRules 
                  Height          =   1815
                  Left            =   120
                  TabIndex        =   24
                  Top             =   600
                  Width           =   6255
                  _ExtentX        =   11033
                  _ExtentY        =   3201
                  _Version        =   393217
                  BorderStyle     =   0
                  Enabled         =   -1  'True
                  ScrollBars      =   2
                  TextRTF         =   $"frmVisaHealth.frx":13BA
               End
               Begin VB.CheckBox chkNoVisa 
                  BackColor       =   &H00DADAB6&
                  Caption         =   "Visa is not required travelling to all cities"
                  Height          =   350
                  Left            =   120
                  TabIndex        =   44
                  Top             =   240
                  Width           =   3255
               End
               Begin MyCommandButton.MyButton cmdGetInfo 
                  Height          =   360
                  Left            =   3480
                  TabIndex        =   45
                  Top             =   200
                  Width           =   1815
                  _ExtentX        =   3201
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
                  Picture         =   "frmVisaHealth.frx":143C
                  AppearanceThemes=   1
                  BackColorDown   =   3968251
                  BackColorOver   =   6805503
                  BackColorFocus  =   16765357
                  BackColorDisabled=   16765357
                  TransparentColor=   14215660
                  Caption         =   "Get &Info"
                  Depth           =   1
                  PictureOffsetX  =   5
                  GradientType    =   2
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00DADAB6&
               Caption         =   "Timatic Request"
               Height          =   2475
               Left            =   120
               TabIndex        =   37
               Top             =   60
               Width           =   4095
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   5
                  Left            =   1560
                  MaxLength       =   3
                  TabIndex        =   9
                  Top             =   1320
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   6
                  Left            =   2040
                  MaxLength       =   3
                  TabIndex        =   10
                  Top             =   1320
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   7
                  Left            =   2520
                  MaxLength       =   3
                  TabIndex        =   11
                  Top             =   1320
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   8
                  Left            =   3000
                  MaxLength       =   3
                  TabIndex        =   12
                  Top             =   1320
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   9
                  Left            =   3480
                  MaxLength       =   3
                  TabIndex        =   13
                  Top             =   1320
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   4
                  Left            =   3480
                  MaxLength       =   3
                  TabIndex        =   8
                  Top             =   960
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   3
                  Left            =   3000
                  MaxLength       =   3
                  TabIndex        =   7
                  Top             =   960
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   2
                  Left            =   2520
                  MaxLength       =   3
                  TabIndex        =   6
                  Top             =   960
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   1
                  Left            =   2040
                  MaxLength       =   3
                  TabIndex        =   5
                  Top             =   960
                  Width           =   500
               End
               Begin VB.TextBox txtDes 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   0
                  Left            =   1560
                  MaxLength       =   3
                  TabIndex        =   4
                  Top             =   960
                  Width           =   500
               End
               Begin VB.TextBox txtEmbark 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Left            =   3480
                  MaxLength       =   3
                  TabIndex        =   3
                  Top             =   600
                  Width           =   500
               End
               Begin VB.TextBox txtVisitCity 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   4
                  Left            =   3480
                  MaxLength       =   3
                  TabIndex        =   23
                  Top             =   2040
                  Width           =   500
               End
               Begin VB.TextBox txtVisitCity 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   3
                  Left            =   3000
                  MaxLength       =   3
                  TabIndex        =   22
                  Top             =   2040
                  Width           =   500
               End
               Begin VB.TextBox txtVisitCity 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   2
                  Left            =   2520
                  MaxLength       =   3
                  TabIndex        =   21
                  Top             =   2040
                  Width           =   500
               End
               Begin VB.TextBox txtVisitCity 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   1
                  Left            =   2040
                  MaxLength       =   3
                  TabIndex        =   20
                  Top             =   2040
                  Width           =   500
               End
               Begin VB.TextBox txtVisitCity 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   0
                  Left            =   1560
                  MaxLength       =   3
                  TabIndex        =   19
                  Top             =   2040
                  Width           =   500
               End
               Begin VB.TextBox txtTransit 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   4
                  Left            =   3480
                  MaxLength       =   3
                  TabIndex        =   18
                  Top             =   1680
                  Width           =   500
               End
               Begin VB.TextBox txtTransit 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   3
                  Left            =   3000
                  MaxLength       =   3
                  TabIndex        =   17
                  Top             =   1680
                  Width           =   500
               End
               Begin VB.TextBox txtTransit 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   2
                  Left            =   2520
                  MaxLength       =   3
                  TabIndex        =   16
                  Top             =   1680
                  Width           =   500
               End
               Begin VB.TextBox txtTransit 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   1
                  Left            =   2040
                  MaxLength       =   3
                  TabIndex        =   15
                  Top             =   1680
                  Width           =   500
               End
               Begin VB.TextBox txtTransit 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Index           =   0
                  Left            =   1560
                  MaxLength       =   3
                  TabIndex        =   14
                  Top             =   1680
                  Width           =   500
               End
               Begin VB.ComboBox cmbPassengers 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   1080
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   240
                  Width           =   2895
               End
               Begin VB.TextBox txtNation 
                  BackColor       =   &H00FFFFFF&
                  Height          =   280
                  Left            =   1080
                  MaxLength       =   3
                  TabIndex        =   2
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Destination:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   43
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Embarkation City:"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   42
                  Top             =   600
                  Width           =   1455
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nationality:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   41
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label Label5 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Visited Cities:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   40
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Transit Cities:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   39
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Passenger:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   38
                  Top             =   240
                  Width           =   855
               End
            End
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
         LcK2            =   $"frmVisaHealth.frx":1745
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
      Begin MyCommandButton.MyButton cmdFinish 
         Height          =   360
         Left            =   9120
         TabIndex        =   31
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
         Caption         =   "&Finish"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdNext 
         Height          =   360
         Left            =   8040
         TabIndex        =   30
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
         TabIndex        =   32
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
         TabIndex        =   29
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
         TabIndex        =   46
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
End
Attribute VB_Name = "frmVisaHealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bol1stTab As Boolean
Dim bol2ndTab As Boolean
Dim datFormLoadEnd As Date
Dim datFormLoadStart As Date
Dim datTouchEnd As Date
Dim datGetGDSInfoStart As Date
Dim datGetGDSInfoStop As Date
Dim bolFormLoaded As Boolean



Private Sub chkNoVisa_Click()
    Dim strTemp As String
    Dim strSql As String
    Dim i As Integer
    Dim intK As Integer
    Dim rscountry As ADODB.Recordset
    
    If chkNoVisa.value = 1 Then
          
          strTemp = ""
          
          For intK = 0 To txtDes.Count - 1
             If Trim(txtDes(intK)) <> "" Then
             strTemp = strTemp & IIf(strTemp <> "", "," & "'" & Trim(txtDes(intK)) & "'", "'" & Trim(txtDes(intK)) & "'")
             End If
          Next
          If strTemp <> "" Then
             strSql = "select distinct countryname from tblcountrycodes where countrycode in (select distinct countrycode from tblcitycodes where airportcode in (" & strTemp & ")) order by countryname"
             Set rscountry = gdbConn.Execute(strSql)
             strTemp = ""
             Do Until rscountry.EOF
                strTemp = strTemp & Trim(rscountry!CountryName) & ","
                rscountry.MoveNext
             Loop
             rscountry.Close
             Set rscountry = Nothing
         End If
         
        If Right(strTemp, 1) = "," Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        If Trim(strTemp) <> "" Then
           lstRmks.AddItem "Visa is not required travelling to " & Replace(strTemp, ",", "/")
        End If
        
    Else
        For i = 0 To lstRmks.ListCount - 1
            If InStr(1, lstRmks.List(i), "Visa is not required travelling to") > 0 Then
               lstRmks.RemoveItem i
               Exit For
            End If
        Next
    End If
End Sub

Private Sub cmbCountry_Click()
    If cmbCountry.Text <> "" Then
       GetCountryRemarks cmbCountry.Text
    End If
End Sub

Private Sub cmbPassengers_Click()
    Dim lngC As Long
    Dim strNation As String
    
    txtNation.Text = ""
    If cmbPassengers.ListCount > 0 Then
       If cmbPassengers.listindex = 0 Then
            For lngC = 1 To gobjPNR.GeneralRemarkCount
                If gobjPNR.GeneralRemark(lngC).Qualifier = "*P" Then
                    If InStr(gobjPNR.GeneralRemark(lngC).RemarkText, "PASSPORT NO:") > 0 Then
                        If InStr(gobjPNR.GeneralRemark(lngC).RemarkText, "-") > 0 Then strNation = Trim(Mid(gobjPNR.GeneralRemark(lngC).RemarkText, InStr(gobjPNR.GeneralRemark(lngC).RemarkText, "-") + 1))
                        If InStr(strNation, "-") > 0 Then strNation = Trim(Mid(strNation, 1, InStr(strNation, "-") - 1))
                        If Len(strNation) < 3 Then txtNation.Text = strNation
                    End If
                End If
            Next
        End If
    End If
End Sub



Private Sub cmdAdd_Click()
If Trim(txtFreeText.Text) <> "" Then
            lstRmks.AddItem txtFreeText.Text
            txtFreeText.Text = ""
End If
End Sub

'Private Sub cmdAddRI_Click()
'    If txtFreeText.Text <> "" Then
'       lstRmks.AddItem txtFreeText.Text
'    End If
'End Sub

Private Sub cmdBack_Click()
    If SftTabs.Tabs.Current - 1 >= 0 Then
       SftTabs.Tabs.Current = SftTabs.Tabs.Current - 1
    End If
End Sub

Private Sub cmdCancel_Click()
    gbolCancelProcess = True
    Unload Me
End Sub

Private Sub cmdMoveAll_Click()
Dim intI As Integer
Dim strTemp As String

For intI = 0 To lstPreset.ListCount - 1
    'lstRmks.AddItem lstPreset.List(intI)
        strTemp = lstPreset.List(intI)
        If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
        Load frmFareRmkFill
        With frmFareRmkFill
            .lblRmkText = strTemp
            '.Show 1, Me
            .Show
            .FormatRemark
            Do
              DoEvents
            Loop Until frmFareRmkFill.Visible = False
            strTemp = .lblRmkText.Caption
            If strTemp = "" Then
                'CANCEL
            Else
                lstRmks.AddItem strTemp
            End If
            Unload frmFareRmkFill
        End With
    Else
        lstRmks.AddItem strTemp
    End If
    Set frmFareRmkFill = Nothing

    
    
    
Next intI


End Sub

Private Sub cmdMoveRmk_Click()
Dim intI As Integer
Dim strTemp As String

For intI = 0 To lstPreset.ListCount - 1
    If lstPreset.Selected(intI) Then
        'lstRmks.AddItem lstPreset.List(intI)
        strTemp = lstPreset.List(intI)
        If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
        Load frmFareRmkFill
        With frmFareRmkFill
            .lblRmkText = strTemp
            '.Show 1, Me
            .Show
            .FormatRemark
            Do
              DoEvents
            Loop Until frmFareRmkFill.Visible = False
            strTemp = .lblRmkText.Caption
            If strTemp = "" Then
                'CANCEL
            Else
                lstRmks.AddItem strTemp
            End If
            Unload frmFareRmkFill
        End With
    Else
        lstRmks.AddItem strTemp
    End If
    Set frmFareRmkFill = Nothing

        
        
        
    End If
Next intI
End Sub

Private Sub cmdNext_Click()
    If SftTabs.Tabs.Current + 1 < SftTabs.Tabs.Count Then
       SftTabs.Tabs.Current = SftTabs.Tabs.Current + 1
    End If
End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload Me
End Sub

Private Sub Form_Load()
   Dim oldParent As Long
   
   chkNoVisa.Visible = False
   datFormLoadStart = Now

    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)
    
   Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   
   pDisplayToFP "*R"
   loadPassengers cmbPassengers
   PopulateCtrls
   bol1stTab = True
   bol2ndTab = False
   If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
   Else
      cmdPrevious.Visible = False
   End If
   cmdBack.Enabled = False
   cmdFinish.Enabled = False
   GetVisaHealthInfo False
   SftTabs.Tabs.Current = 0
   'Preethi - V1.2.4 20110613 - CR 19 - Generate Standard Remarks For Visa and Fare Quotes
   lstPreset.AddItem ("VISA IS NOT REQUIRED")
   
   datFormLoadEnd = Now
   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
   bolFormLoaded = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       gbolCancelProcess = True
    End If
End Sub

Private Sub lstPreset_DblClick()
    'If lstPreset.Text <> "" Then lstRmks.AddItem lstPreset.Text
Dim strTemp As String
    
If lstPreset.Text <> "" Then
     
    strTemp = lstPreset.Text
    If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
        Load frmFareRmkFill
        With frmFareRmkFill
            .lblRmkText = strTemp
            '.Show 1, Me
            .Show
            .FormatRemark
            Do
              DoEvents
            Loop Until frmFareRmkFill.Visible = False
            strTemp = .lblRmkText.Caption
            If strTemp = "" Then
                'CANCEL
            Else
                lstRmks.AddItem strTemp
            End If
            Unload frmFareRmkFill
        End With
    Else
        lstRmks.AddItem strTemp
    End If
    Set frmFareRmkFill = Nothing

End If
        
    
    
'    If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
'                Load frmFareRmkFill
'                With frmFareRmkFill
'                    .lblRmkText = strTemp
'                    '.Show 1, Me
'                    .Show
'                    .FormatRemark
'                     Do
'                      DoEvents
'                    Loop Until frmFareRmkFill.Visible = False
'                    strTemp = .lblRmkText.Caption
'                 If strTemp = "" Then
'
'                    Else
'                        lstEORmks(1).AddItem strTemp
'                        'lstEORmks(0).RemoveItem lngC
'                        'lngC = lngC - 1
'
'                    End If
'                    Unload frmFareRmkFill
'                End With
'
'            Set frmFareRmkFill = Nothing
'
'    End If
    
End Sub





Private Sub SftTabs_Switched()
    If SftTabs.Tabs.Current = 1 Then
       GetCountry
       cmbCountry.SetFocus
    End If
End Sub

Private Sub SftTabs_Switching(NextTab As Integer, Allow As Boolean, Refresh As Boolean)
    If NextTab = 0 Then
       bol1stTab = True
       cmdBack.Enabled = False
       cmdNext.Enabled = True
    ElseIf NextTab = 1 Then
       bol2ndTab = True
       cmdBack.Enabled = True
       cmdNext.Enabled = False
       'If bolFormLoaded Then Set cmbCountry.SetFocus = True
    End If
    If bol1stTab = True And bol2ndTab = True Then
       cmdFinish.Enabled = True
    Else
       cmdFinish.Enabled = False
    End If
End Sub

Private Sub txtFreeText_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Trim(txtFreeText.Text) <> "" Then
            lstRmks.AddItem txtFreeText.Text
            txtFreeText.Text = ""
        End If
    End If
End Sub

Private Sub txtNation_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub GetVisaHealthInfo(Optional showMsg As Boolean = True)
    Dim intC As Integer
    Dim intD As Integer
    Dim strCmd As String
    Dim strRes As String
    Dim blnNextPg As Boolean
    Dim strTemp As String

    If validData(showMsg) = False Then Exit Sub
    
    cmdGetInfo.Enabled = False
    cmdGetInfo.Caption = "Loading Info..."
    
    strCmd = ""
    strCmd = "TI-RA" & Space(12) & "TIMATIC VISA AND HEALTH INFORMATION REQUEST"
    strCmd = strCmd & Space(3) & "3 NATIONALITY     :NA·"
    strCmd = strCmd & Trim(txtNation.Text)
    strCmd = strCmd & Space(34) & "2 EMBARKATION CITY:EM·" & GetValue(txtEmbark)
    strCmd = strCmd & Space(39) & "1 DESTINATION     :DE·"
    For intC = 0 To txtDes.Count - 1
        strCmd = strCmd & GetValue(txtDes(intC)) & "/"
    Next
    If Right(strCmd, 1) = "/" Then strCmd = Mid(strCmd, 1, Len(strCmd) - 1)
    strCmd = strCmd & Space(3) & "0 TRANSIT CITIES  :TR·"
    For intC = 0 To txtTransit.Count - 1
        strCmd = strCmd & GetValue(txtTransit(intC)) & "/"
    Next
    strCmd = strCmd & Dot(3) & "/" & Dot(3) & "/" & Dot(3) & "/" & Dot(3) & "/" & Dot(3)
    strCmd = strCmd & Space(3) & "0 CITIES VISITED  :VT·"
    For intC = 0 To txtVisitCity.Count - 1
        strCmd = strCmd & GetValue(txtVisitCity(intC)) & "/"
    Next
    strCmd = strCmd & Dot(3) & "/" & Dot(3) & "/" & Dot(3) & "/" & Dot(3) & "/" & Dot(3)
    If gobjHost Is Nothing Then Set gobjHost = New CWT_Galileo3.GalileoHost
    strRes = gobjHost.terminalEntry(strCmd, True)

    If InStr(strRes, "TIMATIC") = 0 Then
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, "Unable to display visa information, GDS response:" & vbCrLf & strRes, vbOKOnly + vbDefaultButton1, "Get Visa Info"
    
    ElseIf InStr(strRes, "TIPN") > 0 Then
        txtRules = ""
        strTemp = strRes
         Do
             strRes = NextPage
             strTemp = strTemp & strRes
             If InStr(strRes, "TIPN") > 0 Then
                blnNextPg = True
             Else
                blnNextPg = False
             End If
         Loop While blnNextPg
         txtRules = Replace(strTemp, "TIPN·", vbCrLf)
    Else
        txtRules = strRes
    End If

    cmdGetInfo.Enabled = True
    cmdGetInfo.Caption = "Get &Info"
    If Trim(txtRules.Text) <> "" Then highlightKeywords
End Sub

Private Sub cmdFinish_Click()
    Dim strMsg As String
    
    If lstRmks.ListCount > 0 Then
    
       datTouchEnd = Now
       
       cmdFinish.Enabled = False
       writeDatatoGDS
       'If gbolWritingtoPNR = False Then Exit Sub
       Set gobjPNR = New CWT_GalileoPNR3.PNR
       gobjPNR.loadPNR
       'displayPNRinBar
       
       'Log formload
       'Back up on 26 Sept - Jeremy
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModTimatic), _
'       IIf(gbolCreatPNR = True, gconSModTimatic, ""), Me.Name, gconFormLoad, gstrProcessGrpID, _
'       datFormLoadEnd, datFormLoadStart
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModTimatic), _
'       IIf(gbolCreatPNR = True, gconSModTimatic, ""), Me.Name, gconTouch, gstrProcessGrpID, _
'       datTouchEnd, datFormLoadEnd
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModTimatic), _
'       IIf(gbolCreatPNR = True, gconSModTimatic, ""), Me.Name, gconProcessing, gstrProcessGrpID, _
'        , datTouchEnd
        
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModTimatic, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModTimatic, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModTimatic, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd

       '-----------------------------
       
       Unload Me
    Else
       strMsg = "Itinerary remark is required ..." & Chr(13)
       If strMsg <> "" Then
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       End If
    End If
End Sub

Private Sub cmdGetInfo_Click()
    datGetGDSInfoStart = Now
    Call GetVisaHealthInfo
    datGetGDSInfoStop = Now
    
    'Logging
  
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModTimatic), _
'        IIf(gbolCreatPNR = True, gconSModTimatic, ""), Me.Name, "GET VISA/HEALTH FORM GDS", gstrProcessGrpID, _
'        datGetGDSInfoStop, datGetGDSInfoStart
  
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, gconModAir, frmSideBar.cmbSelectType.Text, _
        gconSModTimatic, Me.Name, "GET VISA/HEALTH FORM GDS", gstrProcessGrpID, _
        datGetGDSInfoStop, datGetGDSInfoStart
    
    
End Sub
Private Sub PopulateCtrls()
    
    Dim lngC As Long
    Dim lngK As Long
    Dim blnExist As Boolean
    
    If gobjPNR.AirSegCount > 0 Then
        txtEmbark = gobjPNR.AirSeg(1).DepartAirport
        txtDes(0) = gobjPNR.AirSeg(1).ArriveAirport
   
        For lngC = 2 To gobjPNR.AirSegCount
            If gobjPNR.AirSeg(lngC).ArriveAirport <> txtEmbark Then
                If lngC = gobjPNR.AirSegCount Then
                    If gobjPNR.AirSeg(1).DepartAirport = gobjPNR.AirSeg(lngC).ArriveAirport Then
                        Exit For
                    End If
                End If
                blnExist = False
                For lngK = 0 To txtDes.Count - 1
                        If txtDes(lngK) = "" Then
                            Exit For
                        Else
                            If txtDes(lngK) = gobjPNR.AirSeg(lngC).ArriveAirport Then
                                blnExist = True
                                Exit For
                            End If
                        End If
                Next
                
                If blnExist = False Then
                    For lngK = 0 To txtDes.Count - 1
                        If txtDes(lngK) = "" Then
                            txtDes(lngK) = gobjPNR.AirSeg(lngC).ArriveAirport
                            Exit For
                        End If
                    Next
                End If
            End If
        Next
    End If
    GetCountry
End Sub
Private Function validData(showMsg As Boolean) As Boolean

    Dim strMsg As String
    Dim intI As Integer
    Dim blnFill As Boolean
    
    blnFill = False
        If Trim(txtNation.Text) <> "" Then
        blnFill = True
            If Len(Trim(txtNation.Text)) < 2 Or Len(Trim(txtNation.Text)) > 3 Then
                strMsg = strMsg & "NATIONALITY CODE MUST BE 3 CHAR AIRPORT/CITY CODE OR 2 CHAR COUNTRY CODE" & vbCrLf
            End If
        End If
    
    If blnFill <> True Then strMsg = strMsg & "NEED NATIONALITY" & vbCrLf
    
    
    If Trim(txtEmbark) = "" Then strMsg = strMsg & "NEED EMBARKATION CITY" & vbCrLf
    If Trim(txtEmbark) <> "" And Len(Trim(txtEmbark)) < 2 Then strMsg = strMsg & "EMBARKATION CITY CODE MUST BE 3 CHAR AIRPORT/CITY CODE OR 2 CHAR COUNTRY CODE" & vbCrLf
    
    blnFill = False
    For intI = 0 To txtDes.Count - 1
        If Trim(txtDes(intI)) <> "" Then
            blnFill = True
            If Len(Trim(txtDes(intI))) < 2 Then
                strMsg = strMsg & "DESTINATION CODE MUST BE 3 CHAR AIRPORT/CITY CODE OR 2 CHAR COUNTRY CODE" & vbCrLf
                Exit For
            End If
        End If
    Next
    If blnFill <> True Then strMsg = strMsg & "NEED AT LEAST 1 DESTINATION" & vbCrLf
    
    For intI = 0 To txtTransit.Count - 1
        If Trim(txtTransit(intI)) <> "" And Len(Trim(txtTransit(intI))) < 2 Then
            strMsg = strMsg & "TRANSIT CITIES CODE MUST BE 3 CHAR AIRPORT/CITY CODE OR 2 CHAR COUNTRY CODE" & vbCrLf
            Exit For
        End If
    Next
    
    For intI = 0 To txtVisitCity.Count - 1
        If Trim(txtVisitCity(intI)) <> "" And Len(Trim(txtVisitCity(intI))) < 2 Then
            strMsg = strMsg & "CITIES VSISITED CODE MUST BE 3 CHAR AIRPORT/CITY CODE OR 2 CHAR COUNTRY CODE" & vbCrLf
            Exit For
        End If
    Next
    
    If strMsg <> "" Then
        If showMsg Then
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "Get Visa Info"
        End If
        validData = False
    Else
        validData = True
    End If
    
End Function
Private Function GetValue(textctrl As TextBox) As String

    If Trim(textctrl) <> "" Then
        If textctrl.Name = "txtNation" Then
            If textctrl.Index = 1 Then
                GetValue = Trim(textctrl) & Dot(4 - Len(Trim(textctrl)))
            Else
                GetValue = Trim(textctrl) & Dot(3 - Len(Trim(textctrl)))
            End If
        Else
            GetValue = Trim(textctrl) & Dot(3 - Len(Trim(textctrl)))
        End If
     Else
        If textctrl.Name = "txtNation" Then
            If textctrl.Index = 1 Then
                GetValue = Dot(4)
            Else
                GetValue = Dot(3)
            End If
        Else
            GetValue = Dot(3)
        End If
    End If
   
End Function
Private Function Dot(DotNo As Integer) As String
    Dim intI As Integer

    For intI = 1 To DotNo
        Dot = Dot & "."
    Next intI
End Function

Private Function NextPage() As String
    NextPage = gobjHost.terminalEntry("TIPN·", True)
End Function

Private Sub lstRmks_DblClick()
If lstRmks.listindex <> -1 Then
    txtFreeText.Text = lstRmks.Text
    lstRmks.RemoveItem lstRmks.listindex
    
End If
End Sub

Private Sub txtDes_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtEmbark_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTransit_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtVisitCity_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub writeDatatoGDS()

    Dim lngC As Long
    Dim strCmd As String
    Dim strResponse As String
    Dim strMsg As String
    Dim i As Integer
    Dim strTemp As String
    
    Dim strNation As String
    Dim PaxNation As String
    'Preethi - V1.2.4 20110531 - CR 19 - Generate Standard Remarks For Visa and Fare Quotes
    Dim rs As ADODB.Recordset
    Dim strSql As String
    Dim intI As Integer
    Dim blnVisa As Boolean

    blnVisa = False
    If lstRmks.ListCount > 0 Then
      With lstRmks
         For intI = 0 To .ListCount - 1
               .listindex = intI
               If .Text = "VISA IS NOT REQUIRED" Then
                  blnVisa = True
                  Exit For
               End If
         Next
      End With
    End If
    
    If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
    gbolWritingtoPNR = True
        
    strCmd = "RI.************* VISA AND PASSPORT ADVICE ***************"
    
    If lstRmks.ListCount > 0 Then
         With lstRmks
              For lngC = 0 To .ListCount - 1
                  .listindex = lngC
                  strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & .Text
              Next
         End With
    Else
        strCmd = strCmd + "+RI.VISA NOT REQUIRED FOR THIS ITINERARY"
    End If

    strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI.******************************************************"
    'Preethi - V1.2.4 20110531 - CR 19 - Generate Standard Remarks For Visa and Fare Quotes
    strSql = "Select OptionValue  from tblModOptions where OptionCode = 'StandardPassportRmk' order by OptionSecCode"
    Set rs = gdbConn.Execute(strSql)
    Do Until rs.EOF
      strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & rs!optionvalue & ""
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    strSql = ""
    If blnVisa = True Then
       strSql = "Select OptionValue  from tblModOptions where OptionCode = 'StandardNoVisaRmk' order by OptionSecCode"
    Else
       strSql = "Select OptionValue  from tblModOptions where OptionCode = 'StandardVisaRmk' order by OptionSecCode"
    End If
    Set rs = gdbConn.Execute(strSql)
    Do Until rs.EOF
      strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & rs!optionvalue & ""
     
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI.******************************************************"
    strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP.SS*VBITI"
        
    'New function to check if Nationality exist. if no, write into NP*P*CITIZENSHIP
    For lngC = 1 To gobjPNR.GeneralRemarkCount
        If gobjPNR.GeneralRemark(lngC).Qualifier = "*P" Then
            If InStr(gobjPNR.GeneralRemark(lngC).RemarkText, "PASSPORT NO:") > 0 Then
                If InStr(gobjPNR.GeneralRemark(lngC).RemarkText, "-") > 0 Then strNation = Trim(Mid(gobjPNR.GeneralRemark(lngC).RemarkText, InStr(gobjPNR.GeneralRemark(lngC).RemarkText, "-") + 1))
                If InStr(strNation, "-") > 0 Then strNation = Trim(Mid(strNation, 1, InStr(strNation, "-") - 1))
                If Len(strNation) < 3 Then PaxNation = strNation
            End If
        End If
    Next
    
    If PaxNation = "" Then
        '"NP.I*EITINQUEUETIME:" & strNow
        'NP.P*PASSPORT NO: 22258964B-SG-ISS -EXP 19MAY15
        '"NP.P.CITIZENSHIP:" & "txtNation"
    
        'This will do
        '"NP.P.PASSPORT NO: -" & "txtNation" & "-"
        strCmd = strCmd + "+NP.P*PASSPORT NO: -" + txtNation + "-"
    End If
    
    
    'Preethi - V1.2.4 20110614 - CR 76 - Change Validation Logic For ENDPNR
    'send entries, received & end the PNR
    If gbolCreatPNR = True And gobjPNR.RecLoc = "" Then
    Else
        strCmd = strCmd & IIf(strCmd = "", "", "+") & "R." & IIf(Trim(frmSideBar.txtRequestor.Text) = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
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
    
    Exit Sub

    
End Sub

Private Sub GetCountry()
   Dim rscountry As ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   Dim intK As Integer
   
   cmbCountry.Clear
   strTemp = ""
   For intK = 0 To txtDes.Count - 1
       If Trim(txtDes(intK)) <> "" Then
          strTemp = strTemp & IIf(strTemp <> "", "," & "'" & Trim(txtDes(intK)) & "'", "'" & Trim(txtDes(intK)) & "'")
       End If
   Next
   If strTemp <> "" Then
       strSql = "select distinct countryname from tblcountrycodes where countrycode in (select distinct countrycode from tblcitycodes where airportcode in (" & strTemp & ")) order by countryname"
       Set rscountry = gdbConn.Execute(strSql)
       Do Until rscountry.EOF
          cmbCountry.AddItem rscountry!CountryName
          rscountry.MoveNext
       Loop
       rscountry.Close
       Set rscountry = Nothing
   End If
   
   'Add Others in Country dropdown for general remarks
   cmbCountry.AddItem "OTHERS"
   
End Sub

Private Sub GetCountryRemarks(strCountry As String)
   Dim rsVisaRemarks As ADODB.Recordset
   Dim rsOtherRemarks As ADODB.Recordset
   Dim strSql As String
   
   lstPreset.Clear
   
   'strSQL = "Select Remarks from tblVisaRemarks Where CountryCode='" & GetCountryCode(strCountry) & "'" & _
            " order by Remarks"

''   strSql = "Select Remarks from tblVisaRemarks Where CountryCode='" & GetCountryCode(strCountry) & "' or CountryCode='ALL'"
''
''   Set rsVisaRemarks = gdbConn.Execute(strSql)
''
''   Do Until rsVisaRemarks.EOF
''      lstPreset.AddItem rsVisaRemarks!Remarks
''      rsVisaRemarks.MoveNext
''   Loop
''
''    rsVisaRemarks.Close
''    Set rsVisaRemarks = Nothing

    strSql = "Select Remarks from tblVisaRemarks Where CountryCode='" & GetCountryCode(strCountry) & "'"
    
    Set rsVisaRemarks = gdbConn.Execute(strSql)
    
    If rsVisaRemarks.EOF Then
        'No remarks found this country, get general remarks
        strSql = "Select Remarks from tblVisaRemarks Where CountryCode='ALL'"
        Set rsOtherRemarks = gdbConn.Execute(strSql)
        
        Do Until rsOtherRemarks.EOF
           lstPreset.AddItem rsOtherRemarks!Remarks
           rsOtherRemarks.MoveNext
        Loop
        
        rsOtherRemarks.Close
        Set rsOtherRemarks = Nothing
        
    Else
        Do Until rsVisaRemarks.EOF
           lstPreset.AddItem rsVisaRemarks!Remarks
           rsVisaRemarks.MoveNext
        Loop
    
        rsVisaRemarks.Close
        Set rsVisaRemarks = Nothing
    
    End If

End Sub

Private Sub highlightKeywords()
   Dim rs As ADODB.Recordset
   Dim strSql As String
   Dim strHighlight As String
   Dim i As Integer
   Dim intRed As Integer
   Dim intGreen As Integer
   Dim intBlue As Integer
      
   strSql = "Select Keyword,Red,Green,Blue from tblKeywords where Type='TIMATIC'"

   Set rs = gdbConn.Execute(strSql)
   
   Do Until rs.EOF
      With txtRules
         strHighlight = Trim(rs!Keyword) & ""
         intRed = rs!red
         intGreen = rs!green
         intBlue = rs!blue
         i = .Find(strHighlight, 1, Len(.Text))
         If i <> -1 Then
            .SelStart = i
            .SelLength = Len(strHighlight)
            .SelBold = True
            .SelColor = RGB(intRed, intGreen, intBlue)
         End If
      End With
      rs.MoveNext
   Loop
   
   rs.Close
   Set rs = Nothing
   With txtRules
        .SelStart = 0
        .SelBold = False
        .SelColor = RGB(0, 0, 0)
   End With
End Sub


