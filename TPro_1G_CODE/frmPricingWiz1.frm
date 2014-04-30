VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPricingWiz1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CWT TravelPro - PNR Pricing Wizard"
   ClientHeight    =   7605
   ClientLeft      =   1860
   ClientTop       =   3510
   ClientWidth     =   10125
   Icon            =   "frmPricingWiz1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10125
   Begin VB.ComboBox cmbPx 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   215
      Text            =   "cmbPx"
      Top             =   480
      Width           =   5055
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Finished"
      Default         =   -1  'True
      DisabledPicture =   "frmPricingWiz1.frx":08CA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4920
      MaskColor       =   &H8000000B&
      Picture         =   "frmPricingWiz1.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   6840
      Width           =   2052
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   600
      MaskColor       =   &H8000000B&
      Picture         =   "frmPricingWiz1.frx":114E
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   6840
      Width           =   2052
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7080
      MaskColor       =   &H8000000B&
      Picture         =   "frmPricingWiz1.frx":1590
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   6840
      Width           =   2052
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Co&ntinue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2760
      MaskColor       =   &H8000000B&
      Picture         =   "frmPricingWiz1.frx":19D2
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   6840
      Width           =   2052
   End
   Begin TabDlg.SSTab sstTabs 
      Height          =   5775
      Left            =   45
      TabIndex        =   126
      Top             =   960
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   573
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Segment Info"
      TabPicture(0)   =   "frmPricingWiz1.frx":1E14
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblLabels(7)"
      Tab(0).Control(1)=   "lblLabels(6)"
      Tab(0).Control(2)=   "lblLabels(5)"
      Tab(0).Control(3)=   "lblLabels(4)"
      Tab(0).Control(4)=   "lblLabels(10)"
      Tab(0).Control(5)=   "lblLabels(22)"
      Tab(0).Control(6)=   "lblLabels(23)"
      Tab(0).Control(7)=   "lblLabels(24)"
      Tab(0).Control(8)=   "chkConnection(7)"
      Tab(0).Control(9)=   "chkConnection(6)"
      Tab(0).Control(10)=   "chkConnection(5)"
      Tab(0).Control(11)=   "chkConnection(4)"
      Tab(0).Control(12)=   "chkConnection(3)"
      Tab(0).Control(13)=   "chkConnection(2)"
      Tab(0).Control(14)=   "chkConnection(1)"
      Tab(0).Control(15)=   "chkConnection(0)"
      Tab(0).Control(16)=   "txtTktDesig(7)"
      Tab(0).Control(17)=   "txtTktDesig(6)"
      Tab(0).Control(18)=   "txtTktDesig(5)"
      Tab(0).Control(19)=   "txtTktDesig(4)"
      Tab(0).Control(20)=   "txtTktDesig(3)"
      Tab(0).Control(21)=   "txtTktDesig(2)"
      Tab(0).Control(22)=   "txtTktDesig(1)"
      Tab(0).Control(23)=   "txtTktDesig(0)"
      Tab(0).Control(24)=   "txtFBC(7)"
      Tab(0).Control(25)=   "txtFBC(6)"
      Tab(0).Control(26)=   "txtFBC(5)"
      Tab(0).Control(27)=   "txtFBC(4)"
      Tab(0).Control(28)=   "txtFBC(3)"
      Tab(0).Control(29)=   "txtFBC(2)"
      Tab(0).Control(30)=   "txtFBC(1)"
      Tab(0).Control(31)=   "txtFBC(0)"
      Tab(0).Control(32)=   "txtFlightInfo(0)"
      Tab(0).Control(33)=   "txtFlightInfo(1)"
      Tab(0).Control(34)=   "txtFlightInfo(2)"
      Tab(0).Control(35)=   "txtFlightInfo(3)"
      Tab(0).Control(36)=   "txtFlightInfo(4)"
      Tab(0).Control(37)=   "txtFlightInfo(5)"
      Tab(0).Control(38)=   "txtFlightInfo(6)"
      Tab(0).Control(39)=   "txtFlightInfo(7)"
      Tab(0).Control(40)=   "txtValue(0)"
      Tab(0).Control(41)=   "txtValue(1)"
      Tab(0).Control(42)=   "txtValue(2)"
      Tab(0).Control(43)=   "txtValue(3)"
      Tab(0).Control(44)=   "txtValue(4)"
      Tab(0).Control(45)=   "txtValue(5)"
      Tab(0).Control(46)=   "txtValue(6)"
      Tab(0).Control(47)=   "txtValue(7)"
      Tab(0).Control(48)=   "txtPriceFBC(0)"
      Tab(0).Control(49)=   "txtPriceFBC(1)"
      Tab(0).Control(50)=   "txtPriceFBC(2)"
      Tab(0).Control(51)=   "txtPriceFBC(3)"
      Tab(0).Control(52)=   "txtPriceFBC(4)"
      Tab(0).Control(53)=   "txtPriceFBC(5)"
      Tab(0).Control(54)=   "txtPriceFBC(6)"
      Tab(0).Control(55)=   "txtPriceFBC(7)"
      Tab(0).Control(56)=   "txtPriceFBC(8)"
      Tab(0).Control(57)=   "txtPriceFBC(9)"
      Tab(0).Control(58)=   "txtPriceFBC(10)"
      Tab(0).Control(59)=   "txtValue(8)"
      Tab(0).Control(60)=   "txtValue(9)"
      Tab(0).Control(61)=   "txtValue(10)"
      Tab(0).Control(62)=   "txtFlightInfo(8)"
      Tab(0).Control(63)=   "txtFlightInfo(9)"
      Tab(0).Control(64)=   "txtFlightInfo(10)"
      Tab(0).Control(65)=   "txtFBC(8)"
      Tab(0).Control(66)=   "txtFBC(9)"
      Tab(0).Control(67)=   "txtFBC(10)"
      Tab(0).Control(68)=   "txtTktDesig(8)"
      Tab(0).Control(69)=   "txtTktDesig(9)"
      Tab(0).Control(70)=   "txtTktDesig(10)"
      Tab(0).Control(71)=   "chkConnection(8)"
      Tab(0).Control(72)=   "chkConnection(9)"
      Tab(0).Control(73)=   "chkConnection(10)"
      Tab(0).Control(74)=   "txtPriceFBC(11)"
      Tab(0).Control(75)=   "txtValue(11)"
      Tab(0).Control(76)=   "txtFlightInfo(11)"
      Tab(0).Control(77)=   "txtFBC(11)"
      Tab(0).Control(78)=   "txtTktDesig(11)"
      Tab(0).Control(79)=   "chkConnection(11)"
      Tab(0).Control(80)=   "txtNVB(0)"
      Tab(0).Control(81)=   "txtNVA(0)"
      Tab(0).Control(82)=   "txtNVB(1)"
      Tab(0).Control(83)=   "txtNVA(1)"
      Tab(0).Control(84)=   "txtNVB(2)"
      Tab(0).Control(85)=   "txtNVA(2)"
      Tab(0).Control(86)=   "txtNVB(3)"
      Tab(0).Control(87)=   "txtNVA(3)"
      Tab(0).Control(88)=   "txtNVB(4)"
      Tab(0).Control(89)=   "txtNVA(4)"
      Tab(0).Control(90)=   "txtNVB(5)"
      Tab(0).Control(91)=   "txtNVA(5)"
      Tab(0).Control(92)=   "txtNVB(6)"
      Tab(0).Control(93)=   "txtNVA(6)"
      Tab(0).Control(94)=   "txtNVB(7)"
      Tab(0).Control(95)=   "txtNVA(7)"
      Tab(0).Control(96)=   "txtNVB(8)"
      Tab(0).Control(97)=   "txtNVA(8)"
      Tab(0).Control(98)=   "txtNVB(9)"
      Tab(0).Control(99)=   "txtNVA(9)"
      Tab(0).Control(100)=   "txtNVB(10)"
      Tab(0).Control(101)=   "txtNVA(10)"
      Tab(0).Control(102)=   "txtNVB(11)"
      Tab(0).Control(103)=   "txtNVA(11)"
      Tab(0).Control(104)=   "txtBag(0)"
      Tab(0).Control(105)=   "txtBag(1)"
      Tab(0).Control(106)=   "txtBag(2)"
      Tab(0).Control(107)=   "txtBag(3)"
      Tab(0).Control(108)=   "txtBag(4)"
      Tab(0).Control(109)=   "txtBag(5)"
      Tab(0).Control(110)=   "txtBag(6)"
      Tab(0).Control(111)=   "txtBag(7)"
      Tab(0).Control(112)=   "txtBag(8)"
      Tab(0).Control(113)=   "txtBag(9)"
      Tab(0).Control(114)=   "txtBag(10)"
      Tab(0).Control(115)=   "txtBag(11)"
      Tab(0).Control(116)=   "txtSegNum(11)"
      Tab(0).Control(117)=   "txtSegNum(10)"
      Tab(0).Control(118)=   "txtSegNum(9)"
      Tab(0).Control(119)=   "txtSegNum(8)"
      Tab(0).Control(120)=   "txtSegNum(7)"
      Tab(0).Control(121)=   "txtSegNum(6)"
      Tab(0).Control(122)=   "txtSegNum(5)"
      Tab(0).Control(123)=   "txtSegNum(4)"
      Tab(0).Control(124)=   "txtSegNum(3)"
      Tab(0).Control(125)=   "txtSegNum(2)"
      Tab(0).Control(126)=   "txtSegNum(1)"
      Tab(0).Control(127)=   "txtSegNum(0)"
      Tab(0).ControlCount=   128
      TabCaption(1)   =   "Fare Info"
      TabPicture(1)   =   "frmPricingWiz1.frx":1E30
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblLabels(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLabels(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblLabels(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblLabels(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblLabels(8)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblLabels(9)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblLabels(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblLabels(20)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblLabels(21)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblFareType"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblLabels(37)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblExcTax"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblAComm"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblTransFee"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkTransFee"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtFareInfo(0)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtFareInfo(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtFareInfo(2)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtFareInfo(3)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtFareInfo(4)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtFareInfo(5)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtFareInfo(6)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmbFareOnTkt"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtFareInfo(7)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Frame1"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Frame2"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmbFareType"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "chkNRCC"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtFareInfo(8)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "chkDocFee"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "chkShowASF"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "cmbCountry"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtAComm"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtTransFee"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "chkNRCCAC"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Frame4"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "Ticket Modifiers"
      TabPicture(2)   =   "frmPricingWiz1.frx":1E4C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLabels(12)"
      Tab(2).Control(1)=   "lblLabels(13)"
      Tab(2).Control(2)=   "lblLabels(14)"
      Tab(2).Control(3)=   "lblLabels(15)"
      Tab(2).Control(4)=   "lblLabels(16)"
      Tab(2).Control(5)=   "lblLabels(17)"
      Tab(2).Control(6)=   "lblLabels(18)"
      Tab(2).Control(7)=   "lblLabels(19)"
      Tab(2).Control(8)=   "lblLabels(25)"
      Tab(2).Control(9)=   "lblLabels(26)"
      Tab(2).Control(10)=   "lblLabels(38)"
      Tab(2).Control(11)=   "lblLabels(29)"
      Tab(2).Control(12)=   "dtpCCExpDate(1)"
      Tab(2).Control(13)=   "txtTktMod(0)"
      Tab(2).Control(14)=   "txtTktMod(1)"
      Tab(2).Control(15)=   "txtTktMod(3)"
      Tab(2).Control(16)=   "cmbFOP(0)"
      Tab(2).Control(17)=   "cmbFOP(1)"
      Tab(2).Control(18)=   "dtpCCExpDate(0)"
      Tab(2).Control(19)=   "cmbFOP(2)"
      Tab(2).Control(20)=   "txtTktMod(2)"
      Tab(2).Control(21)=   "cmbFOP(3)"
      Tab(2).Control(22)=   "cmbValCarrier"
      Tab(2).Control(23)=   "txtTktMod(4)"
      Tab(2).Control(24)=   "txtTktMod(5)"
      Tab(2).Control(25)=   "txtTktMod(6)"
      Tab(2).Control(26)=   "txtTktMod(7)"
      Tab(2).Control(27)=   "txtTktMod(8)"
      Tab(2).Control(28)=   "txtTktMod(9)"
      Tab(2).Control(29)=   "chkPaperTkt"
      Tab(2).Control(30)=   "txtTktMod(10)"
      Tab(2).Control(31)=   "chkUATP"
      Tab(2).Control(32)=   "txtDeliveryDate"
      Tab(2).Control(33)=   "cmdFCEndorse"
      Tab(2).ControlCount=   34
      TabCaption(3)   =   "MI"
      TabPicture(3)   =   "frmPricingWiz1.frx":1E68
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraEC"
      Tab(3).Control(1)=   "cmdClientMI"
      Tab(3).Control(2)=   "fraMI"
      Tab(3).Control(3)=   "lvwECodes"
      Tab(3).ControlCount=   4
      Begin VB.Frame fraEC 
         Height          =   5055
         Left            =   -70440
         TabIndex        =   236
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txtMS 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   6
            Tag             =   "BY-"
            Top             =   2520
            Width           =   645
         End
         Begin VB.TextBox txtRS 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   7
            Tag             =   "BY-"
            Top             =   240
            Width           =   645
         End
         Begin MSComctlLib.ListView lvwRealECodes 
            Height          =   1815
            Left            =   240
            TabIndex        =   237
            Top             =   600
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   7056
            EndProperty
         End
         Begin MSComctlLib.ListView lvwMissECodes 
            Height          =   1875
            Left            =   240
            TabIndex        =   238
            Top             =   2880
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   3307
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   7056
            EndProperty
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Missed Saving Code:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   40
            Left            =   360
            TabIndex        =   240
            Top             =   2520
            Width           =   1995
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Realised Saving Code:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   41
            Left            =   120
            TabIndex        =   239
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Add/Override Taxes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         TabIndex        =   228
         Top             =   2160
         Width           =   3615
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   234
            Top             =   600
            Width           =   615
         End
         Begin VB.ListBox lstTax 
            Height          =   645
            Left            =   2160
            TabIndex        =   231
            Top             =   220
            Width           =   1215
         End
         Begin VB.TextBox txtTaxCode 
            Height          =   360
            Left            =   840
            TabIndex        =   230
            Tag             =   "Tax Code"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtTaxAmt 
            Height          =   345
            Left            =   840
            TabIndex        =   229
            Tag             =   "Tax Amt"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Amt"
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
            TabIndex        =   233
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
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
            TabIndex        =   232
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CheckBox chkNRCCAC 
         Caption         =   "NRCC"
         Height          =   375
         Left            =   6480
         TabIndex        =   223
         Top             =   5280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtTransFee 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   221
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAComm 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7440
         TabIndex        =   219
         Tag             =   "NN."
         Top             =   5400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmbCountry 
         Height          =   315
         Left            =   2100
         TabIndex        =   217
         ToolTipText     =   "For VUSA Fare"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton cmdClientMI 
         Caption         =   "Client &MI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -73440
         TabIndex        =   216
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmdFCEndorse 
         Caption         =   "Preloaded EN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68760
         TabIndex        =   214
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtDeliveryDate 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69720
         TabIndex        =   213
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CheckBox chkShowASF 
         Caption         =   "Show ASF?"
         Height          =   255
         Left            =   7320
         TabIndex        =   211
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkDocFee 
         Caption         =   "Return to Client?"
         Height          =   255
         Left            =   7320
         TabIndex        =   210
         Top             =   2340
         Width           =   1575
      End
      Begin VB.TextBox txtFareInfo 
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
         Index           =   8
         Left            =   5520
         TabIndex        =   209
         Tag             =   "NN."
         Top             =   2280
         Width           =   1485
      End
      Begin VB.CheckBox chkUATP 
         Caption         =   "UATP     (Long Charge)"
         Height          =   495
         Left            =   -66960
         TabIndex        =   206
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkNRCC 
         Caption         =   "NRCC"
         Height          =   375
         Left            =   6480
         TabIndex        =   205
         Top             =   4800
         Width           =   855
      End
      Begin VB.ComboBox cmbFareType 
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
         ItemData        =   "frmPricingWiz1.frx":1E84
         Left            =   2040
         List            =   "frmPricingWiz1.frx":1E86
         Style           =   2  'Dropdown List
         TabIndex        =   204
         Top             =   4800
         Visible         =   0   'False
         Width           =   4395
      End
      Begin VB.Frame fraMI 
         Height          =   2880
         Left            =   -74880
         TabIndex        =   197
         Top             =   480
         Width           =   4455
         Begin VB.ComboBox cboBookingAction 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2350
            Width           =   2055
         End
         Begin VB.ComboBox cboTrip 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   225
            Top             =   3120
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtMI 
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
            Index           =   6
            Left            =   2220
            MaxLength       =   2
            TabIndex        =   4
            Tag             =   "BY-"
            Top             =   1920
            Width           =   885
         End
         Begin VB.ComboBox cboClassServ 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1560
            Width           =   3900
         End
         Begin VB.ComboBox cboTripType 
            Height          =   315
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtMI 
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
            Index           =   3
            Left            =   2220
            MaxLength       =   3
            TabIndex        =   2
            Tag             =   "BY-"
            Top             =   1000
            Width           =   885
         End
         Begin VB.TextBox txtMI 
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
            Left            =   2220
            TabIndex        =   0
            Tag             =   "BY-"
            Top             =   180
            Width           =   1665
         End
         Begin VB.TextBox txtMI 
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
            Left            =   2220
            TabIndex        =   1
            Tag             =   "BY-"
            Top             =   600
            Width           =   1665
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Booking Action:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   39
            Left            =   240
            TabIndex        =   226
            Top             =   2350
            Width           =   1875
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Trip Type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   34
            Left            =   1800
            TabIndex        =   224
            Top             =   3120
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Low Fare Carrier:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   36
            Left            =   240
            TabIndex        =   207
            Top             =   1920
            Width           =   1875
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Trip Type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   33
            Left            =   720
            TabIndex        =   202
            Top             =   3720
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Class of Services:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   32
            Left            =   120
            TabIndex        =   201
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Final Destination:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   30
            Left            =   240
            TabIndex        =   200
            Top             =   960
            Width           =   1875
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Reference Fare:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   28
            Left            =   240
            TabIndex        =   199
            Top             =   180
            Width           =   1875
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Low Fare:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   27
            Left            =   240
            TabIndex        =   198
            Top             =   600
            Width           =   1875
         End
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   196
         Top             =   1020
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   195
         Top             =   1380
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   194
         Top             =   1740
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   193
         Top             =   2100
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   192
         Top             =   2460
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   191
         Top             =   2820
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   190
         Top             =   3180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   189
         Top             =   3540
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   188
         Top             =   3900
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   187
         Top             =   4260
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   186
         Top             =   4620
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtSegNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -65160
         MaxLength       =   6
         TabIndex        =   185
         Top             =   4980
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtTktMod 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   -66360
         TabIndex        =   119
         Tag             =   "NN."
         Top             =   2280
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   93
         Tag             =   "BN"
         Top             =   4920
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   86
         Tag             =   "BN"
         Top             =   4560
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   79
         Tag             =   "BN"
         Top             =   4200
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   72
         Tag             =   "BN"
         Top             =   3840
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   65
         Tag             =   "BN"
         Top             =   3480
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -66240
         MaxLength       =   6
         TabIndex        =   58
         Tag             =   "BN"
         Top             =   3120
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   51
         Tag             =   "BN"
         Top             =   2760
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   44
         Tag             =   "BN"
         Top             =   2400
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   37
         Tag             =   "BN"
         Top             =   2040
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   30
         Tag             =   "BN"
         Top             =   1680
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   23
         Tag             =   "BN"
         Top             =   1320
         Width           =   585
      End
      Begin VB.TextBox txtBag 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -66270
         MaxLength       =   6
         TabIndex        =   16
         Tag             =   "BN"
         Top             =   960
         Width           =   585
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -67260
         MaxLength       =   7
         TabIndex        =   92
         Tag             =   "BN"
         Top             =   4920
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   91
         Tag             =   "BN"
         Top             =   4920
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   85
         Tag             =   "BN"
         Top             =   4560
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   84
         Tag             =   "BN"
         Top             =   4560
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   78
         Tag             =   "BN"
         Top             =   4200
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   77
         Tag             =   "BN"
         Top             =   4200
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -67260
         MaxLength       =   7
         TabIndex        =   71
         Tag             =   "BN"
         Top             =   3840
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   70
         Tag             =   "BN"
         Top             =   3840
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   64
         Tag             =   "BN"
         Top             =   3480
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   63
         Tag             =   "BN"
         Top             =   3480
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   57
         Tag             =   "BN"
         Top             =   3120
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   56
         Tag             =   "BN"
         Top             =   3120
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   50
         Tag             =   "BN"
         Top             =   2760
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   49
         Tag             =   "BN"
         Top             =   2760
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   43
         Tag             =   "BN"
         Top             =   2400
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   42
         Tag             =   "BN"
         Top             =   2400
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   36
         Tag             =   "BN"
         Top             =   2040
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   35
         Tag             =   "BN"
         Top             =   2040
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   29
         Tag             =   "BN"
         Top             =   1680
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   28
         Tag             =   "BN"
         Top             =   1680
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   22
         Tag             =   "BN"
         Top             =   1320
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   21
         Tag             =   "BN"
         Top             =   1320
         Width           =   940
      End
      Begin VB.TextBox txtNVA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -67245
         MaxLength       =   7
         TabIndex        =   14
         Tag             =   "BN"
         Top             =   960
         Width           =   940
      End
      Begin VB.TextBox txtNVB 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -68190
         MaxLength       =   7
         TabIndex        =   13
         Tag             =   "BN"
         Top             =   960
         Width           =   940
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   7080
         TabIndex        =   179
         Top             =   540
         Visible         =   0   'False
         Width           =   1035
         Begin VB.OptionButton optCommissionType 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   0
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton optCommissionType 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   7080
         TabIndex        =   178
         Top             =   1800
         Visible         =   0   'False
         Width           =   1035
         Begin VB.OptionButton optDiscType 
            Caption         =   "NF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton optDiscType 
            Caption         =   "PUB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   0
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.TextBox txtFareInfo 
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
         Index           =   7
         Left            =   5520
         TabIndex        =   102
         Tag             =   "NN."
         Text            =   "0"
         Top             =   1800
         Width           =   1485
      End
      Begin VB.CheckBox chkPaperTkt 
         Caption         =   "ET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72480
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cmbFareOnTkt 
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
         ItemData        =   "frmPricingWiz1.frx":1E88
         Left            =   2040
         List            =   "frmPricingWiz1.frx":1E8A
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   4380
         Width           =   4395
      End
      Begin VB.TextBox txtTktMod 
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
         Index           =   9
         Left            =   -69360
         TabIndex        =   109
         Tag             =   "BY-"
         Top             =   1020
         Width           =   1665
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   89
         Tag             =   "BN"
         Top             =   4920
         Width           =   1065
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   88
         Tag             =   "BN"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   11
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   4920
         Width           =   2565
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   150
         Top             =   4980
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   90
         Tag             =   "BN"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   4560
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   4200
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   82
         Tag             =   "BN"
         Top             =   4560
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   75
         Tag             =   "BN"
         Top             =   4200
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   68
         Tag             =   "BN"
         Top             =   3840
         Width           =   1065
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   81
         Tag             =   "BN"
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   74
         Tag             =   "BN"
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   67
         Tag             =   "BN"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   10
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   4560
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   9
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   4200
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   8
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   3840
         Width           =   2565
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   144
         Top             =   3900
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   146
         Top             =   4260
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   148
         Top             =   4620
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   69
         Tag             =   "BN"
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   76
         Tag             =   "BN"
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   83
         Tag             =   "BN"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtTktMod 
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
         Index           =   8
         Left            =   -72480
         MaxLength       =   45
         TabIndex        =   125
         Tag             =   "BY@/.-#*()"
         Top             =   4560
         Width           =   5025
      End
      Begin VB.TextBox txtTktMod 
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
         Index           =   7
         Left            =   -72480
         MaxLength       =   45
         TabIndex        =   124
         Tag             =   "BY@/.-#*()"
         Top             =   4200
         Width           =   5025
      End
      Begin VB.TextBox txtTktMod 
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
         Index           =   6
         Left            =   -72480
         MaxLength       =   45
         TabIndex        =   123
         Tag             =   "BY@/.-#*()"
         Top             =   3840
         Width           =   5025
      End
      Begin VB.TextBox txtTktMod 
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
         Index           =   5
         Left            =   -72480
         MaxLength       =   29
         TabIndex        =   122
         Tag             =   "BY$.-"
         Top             =   3480
         Width           =   3585
      End
      Begin VB.TextBox txtTktMod 
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
         Index           =   4
         Left            =   -72480
         MaxLength       =   29
         TabIndex        =   121
         Tag             =   "BY$.-"
         Top             =   3120
         Width           =   3585
      End
      Begin VB.ComboBox cmbValCarrier 
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
         ItemData        =   "frmPricingWiz1.frx":1E8C
         Left            =   -72480
         List            =   "frmPricingWiz1.frx":1E8E
         TabIndex        =   107
         Top             =   600
         Width           =   855
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
         Index           =   3
         ItemData        =   "frmPricingWiz1.frx":1E90
         Left            =   -70920
         List            =   "frmPricingWiz1.frx":1EAC
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   1860
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTktMod 
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
         Left            =   -70140
         MaxLength       =   18
         TabIndex        =   116
         Tag             =   "NN"
         Top             =   1860
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
         ItemData        =   "frmPricingWiz1.frx":1ED0
         Left            =   -72480
         List            =   "frmPricingWiz1.frx":1EDA
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   1800
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtpCCExpDate 
         Height          =   360
         Index           =   0
         Left            =   -68040
         TabIndex        =   113
         Top             =   1440
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
         Format          =   61997059
         CurrentDate     =   36526
         MaxDate         =   73050
         MinDate         =   36526
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
         ItemData        =   "frmPricingWiz1.frx":1EE7
         Left            =   -70920
         List            =   "frmPricingWiz1.frx":1F03
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
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
         ItemData        =   "frmPricingWiz1.frx":1F27
         Left            =   -72480
         List            =   "frmPricingWiz1.frx":1F37
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   1440
         Width           =   1515
      End
      Begin VB.TextBox txtTktMod 
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
         Index           =   3
         Left            =   -72480
         TabIndex        =   120
         Tag             =   "BY*.#()-"
         Top             =   2700
         Width           =   1665
      End
      Begin VB.TextBox txtTktMod 
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
         Left            =   -70140
         MaxLength       =   18
         TabIndex        =   112
         Tag             =   "NN"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtTktMod 
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
         Left            =   -72480
         TabIndex        =   108
         Tag             =   "BY-"
         Top             =   1020
         Width           =   1665
      End
      Begin VB.TextBox txtFareInfo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Index           =   6
         Left            =   2040
         MaxLength       =   225
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         Tag             =   "BY@/.-#* ()"
         Top             =   3360
         Width           =   5445
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   62
         Tag             =   "BN"
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   55
         Tag             =   "BN"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   48
         Tag             =   "BN"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   41
         Tag             =   "BN"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   34
         Tag             =   "BN"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "BN"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "BN"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtPriceFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -69525
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "BN"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtFareInfo 
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
         Index           =   5
         Left            =   5520
         TabIndex        =   101
         Tag             =   "NN."
         Top             =   1380
         Width           =   1665
      End
      Begin VB.TextBox txtFareInfo 
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
         Index           =   4
         Left            =   5520
         TabIndex        =   100
         Tag             =   "NN."
         Top             =   960
         Width           =   1665
      End
      Begin VB.TextBox txtFareInfo 
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
         Index           =   3
         Left            =   5520
         MaxLength       =   4
         TabIndex        =   97
         Tag             =   "NN."
         Text            =   "0"
         Top             =   540
         Width           =   1485
      End
      Begin VB.TextBox txtFareInfo 
         Enabled         =   0   'False
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
         Left            =   2100
         TabIndex        =   96
         Text            =   "PENDING"
         Top             =   1380
         Width           =   1665
      End
      Begin VB.TextBox txtFareInfo 
         Enabled         =   0   'False
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
         Left            =   2100
         TabIndex        =   95
         Text            =   "PENDING"
         Top             =   960
         Width           =   1665
      End
      Begin VB.TextBox txtFareInfo 
         Enabled         =   0   'False
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
         Left            =   2100
         TabIndex        =   94
         Text            =   "PENDING"
         Top             =   540
         Width           =   1665
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   142
         Top             =   3540
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   140
         Top             =   3180
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   138
         Top             =   2820
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   136
         Top             =   2460
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   134
         Top             =   2100
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   132
         Top             =   1740
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   130
         Top             =   1380
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -64980
         MaxLength       =   6
         TabIndex        =   128
         Top             =   1020
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   7
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   3480
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   6
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3120
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   5
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2760
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   4
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2400
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Index           =   3
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2040
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1680
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1320
         Width           =   2565
      End
      Begin VB.TextBox txtFlightInfo 
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
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   2565
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "BN"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "BN"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "BN"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "BN"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "BN"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   46
         Tag             =   "BN"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   53
         Tag             =   "BN"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtFBC 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -71940
         MaxLength       =   10
         TabIndex        =   60
         Tag             =   "BN"
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -70560
         MaxLength       =   6
         TabIndex        =   11
         Tag             =   "BN"
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "BN"
         Top             =   1320
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   26
         Tag             =   "BN"
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   33
         Tag             =   "BN"
         Top             =   2040
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   40
         Tag             =   "BN"
         Top             =   2400
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   47
         Tag             =   "BN"
         Top             =   2760
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   54
         Tag             =   "BN"
         Top             =   3120
         Width           =   1065
      End
      Begin VB.TextBox txtTktDesig 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -70590
         MaxLength       =   6
         TabIndex        =   61
         Tag             =   "BN"
         Top             =   3480
         Width           =   1065
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   960
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   2040
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   2400
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   3120
         Width           =   375
      End
      Begin VB.CheckBox chkConnection 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   3480
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtpCCExpDate 
         Height          =   360
         Index           =   1
         Left            =   -68040
         TabIndex        =   117
         Top             =   1860
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
         Format          =   61997059
         CurrentDate     =   36526
         MaxDate         =   73050
         MinDate         =   36526
      End
      Begin MSComctlLib.ListView lvwECodes 
         Height          =   4275
         Left            =   -70320
         TabIndex        =   15
         Top             =   5520
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   7541
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6553
         EndProperty
      End
      Begin VB.CheckBox chkTransFee 
         Caption         =   "Include Trans Fee in ASF?"
         Height          =   375
         Left            =   7080
         TabIndex        =   235
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Fare Ladder:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   29
         Left            =   -73920
         TabIndex        =   227
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblTransFee 
         Caption         =   "Transaction Fee"
         Height          =   375
         Left            =   7440
         TabIndex        =   222
         Top             =   4560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblAComm 
         Caption         =   "A. Comm."
         Height          =   255
         Left            =   7440
         TabIndex        =   220
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblExcTax 
         Alignment       =   1  'Right Justify
         Caption         =   "Exclude Tax"
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
         Left            =   840
         TabIndex        =   218
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Delivery Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   38
         Left            =   -71280
         TabIndex        =   212
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Doc Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   37
         Left            =   4500
         TabIndex        =   208
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label lblFareType 
         Alignment       =   1  'Right Justify
         Caption         =   "Fare Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   203
         Top             =   4800
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "PT/ET:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   26
         Left            =   -74040
         TabIndex        =   184
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "PT Surcharge:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   25
         Left            =   -67880
         TabIndex        =   183
         Top             =   2280
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Bag"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   24
         Left            =   -66300
         TabIndex        =   182
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   -64980
         TabIndex        =   181
         Top             =   660
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "NVA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   -67200
         TabIndex        =   180
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount ($)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   3900
         TabIndex        =   177
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Fare on Ticket:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   360
         TabIndex        =   176
         Top             =   4380
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Value Code:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   -70800
         TabIndex        =   175
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "(appended to INV FOP ONLY!)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   -70740
         TabIndex        =   174
         Top             =   2760
         Width           =   3195
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Validating Carrier:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   -74400
         TabIndex        =   173
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "FOP from PAX:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   -74040
         TabIndex        =   172
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Endorsements:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   -74040
         TabIndex        =   171
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "FOP Code:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   -74040
         TabIndex        =   170
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "FOP on Ticket:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   -74040
         TabIndex        =   169
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Tour Code Box:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   -74040
         TabIndex        =   168
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Fare Calculation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   167
         Top             =   3360
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Price FBC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -69420
         TabIndex        =   166
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Base Fare"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   480
         TabIndex        =   165
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Taxes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   480
         TabIndex        =   164
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Fare"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   163
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Commission(%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   3900
         TabIndex        =   162
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Fare (NF)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3900
         TabIndex        =   161
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Selling Fare"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3900
         TabIndex        =   160
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "TD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -70560
         TabIndex        =   159
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Ticket FBC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -71820
         TabIndex        =   158
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "NVB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -68160
         TabIndex        =   157
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Flight Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   -73860
         TabIndex        =   156
         Top             =   600
         Width           =   1515
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "PNR Pricing Wizard"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   155
      Top             =   60
      Width           =   9360
   End
End
Attribute VB_Name = "frmPricingWiz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbytNumSegs As Byte
Dim mbolFormLoaded As Boolean
Dim mbolORPriceFBC As Boolean
Dim mbolORTktFBC As Boolean
Dim mstrCommType As String
Dim mstrDiscType As String
Dim mstrIT_BT As String
Dim mstrFOPType As String
Dim mstrFOPCCInfo As String
Dim mstrFOPCode As String
Dim mbolFareStored As Boolean
Dim mstrFBUFields As String
Dim mbytFFNum As Byte
Dim mbolValidData As Boolean
Dim mstrBaseCurr As String
Dim msngFareDiff As Single
Dim mstrTotalCurr As String
Dim mstrFareCalc As String
Dim mbolFBUMode As Boolean
Dim mbolSysEndo As Boolean
Dim strRDLine() As String
Dim intNIndex() As Integer
Dim strNPxName() As String
Dim intI As Integer
Dim mintDisplayNo As Integer
'29122004
Dim mintPxNum As Integer
Dim mintStartPxNum As Integer
Dim mbolFQ As Boolean
'FMR
Dim mstrFMRCmd As String
'Timer
Dim startTime As Date
Dim blnLoaded As Boolean
Dim blnFirstPx As Boolean
Dim blnChgPCarrier As Boolean
Dim blnCancelFF As Boolean
Dim mstrFFConn As String
Dim SysStart As Date
Dim sngASF As Single
Dim strSegmentNo As String
Dim mbolActivate As Boolean

'CS Add Booking Method
Dim mstrBookingTool As String
Dim mstrFirstSeg As String

Dim datFormLoadEnd As Date
Dim datFormLoadStart As Date
Dim datTouchEnd As Date

Private Sub chkConnection_Click(Index As Integer)
With chkConnection(Index)
    If .value = vbChecked Then
        .Caption = "X"
    Else
        .Caption = "O"
    End If
End With
If mbolFormLoaded Then mstrFFConn = "/CNX"
End Sub



Private Sub chkNRCCAC_Click()
   If chkNRCCAC.value = 1 Then
      txtAComm.Enabled = True
   Else
      txtAComm = ""
      txtAComm.Enabled = False
   End If
End Sub

Private Sub chkPaperTkt_Click()
With chkPaperTkt
    If .value = vbChecked Then
        .Caption = "PT"
    Else
        .Caption = "ET"
    End If
End With

End Sub

Private Sub chkUATP_Click()
   If chkUATP.value = 1 Then
      cmbFOP(2).Text = "CC"
   End If
End Sub

Private Sub cmbFareOnTkt_Click()
    Dim FCEndorseText As String
    'modified on 09/11/04: EB1 will be auto-populated in FF when FareOnTicket <> PUB
    If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PrivateFare And cmbFareOnTkt.listindex < 1 And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
        FCEndorseText = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(1).Endorsement(1)
        If InStr(FCEndorseText, "^") > 0 Then
            txtTktMod(4).Text = RemoveChar(Mid(FCEndorseText, 1, InStr(FCEndorseText, "^") - 1), "B", True, Mid(txtTktMod(4).Tag, 3)) 'Endors1
        Else
            txtTktMod(4).Text = RemoveChar(FCEndorseText, "B", True, Mid(txtTktMod(4).Tag, 3)) 'Endors1
        End If
        cmdFCEndorse.Enabled = False
    Else
        txtTktMod(4).Text = ""
        cmdFCEndorse.Enabled = True
    End If
        
        
End Sub

Private Sub cmbFareType_Click()
If gstrAgcyCountryCode = "HK" Then
    chkNRCC.Visible = False
    chkNRCC.Enabled = False
    lblAComm.Visible = False
    txtAComm.Visible = False
    txtAComm = ""
Else
    If UCase(cmbFareType.Text) = "APF - SPECIAL FARE" Then
        chkNRCC.Visible = True
        chkNRCC.Enabled = True
        lblAComm.Visible = True
        txtAComm.Visible = True
        chkNRCCAC.Visible = True
           
    Else
        chkNRCC.Visible = False
        chkNRCC.Enabled = False
        chkNRCC.value = 0
        chkNRCCAC.value = 0
        chkNRCCAC.Visible = False
        txtAComm = ""
        lblAComm.Visible = False
        txtAComm.Visible = False
        
    End If
End If
End Sub

Private Sub cmbFOP_Click(Index As Integer)
    Dim strMsg As String

    Select Case cmbFOP(Index).Text
        Case "INV", "MS"
            'FMR
            'gbolFMR = False
            chkUATP.Visible = False
            'chkInHouse.Visible = False
            chkUATP.value = 0
            Select Case Index
                Case 0
                    cmbFOP(1).Visible = False
                    txtTktMod(1).Visible = False
                    dtpCCExpDate(0).Visible = False
                    cmbFOP(2).Enabled = True
                    If cmbFOP(2).Text = "CC" Then
                        txtTktMod(2).Visible = True
                        cmbFOP(3).Visible = True
                        dtpCCExpDate(1).Visible = True
                    End If
                Case 2
                    txtTktMod(2).Visible = False
                    cmbFOP(3).Visible = False
                    dtpCCExpDate(1).Visible = False
            End Select
        Case "CC"
'            chkUATP.Visible = True
'            chkUATP.Value = 1
            'FMR
            'gbolFMR = False
            'chkInHouse.Visible = True
            Select Case Index
                Case 0
                    chkUATP.Visible = True
                    chkUATP.value = 1
                    
                    cmbFOP(1).Visible = True
                    cmbFOP(2).Enabled = False
                    txtTktMod(1).Visible = True
                    dtpCCExpDate(0).Visible = True
                    txtTktMod(2).Visible = False
                    cmbFOP(3).Visible = False
                    dtpCCExpDate(1).Visible = False
                    


                Case 2
                    If chkUATP.value <> 1 Then
                        txtTktMod(2).Visible = True
                        cmbFOP(3).Visible = True
                        dtpCCExpDate(1).Visible = True
                    End If
                    If cmbFOP(0).Text = "FMR" Then
                        chkUATP.Visible = True
                        chkUATP.value = 1
                    End If
                Case Else
                    'nothing
                End Select
        Case "MULTI"
            ' can only be on index 0
                'FMR
                'gbolFMR = False
                    cmbFOP(2).Enabled = False
                    cmbFOP(1).Visible = False
                    txtTktMod(1).Visible = False
                    dtpCCExpDate(0).Visible = False

        'FMR
        Case "FMR"
            'gbolFMR = True
            'Added on 14/01/2005: Added for FMR checking, prompt if 2 Px
            If cmbPx.ListCount > 1 Then
                'MsgBox "FMR is not allow for more than 1 passenger", vbOKOnly, "FMR Warning"
                strMsg = "FMR is not allow for more than 1 passenger"
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                cmbFOP(0).listindex = 0
                cmbFOP(0).SetFocus
            Exit Sub
            End If
            'chkUATP.Visible = True
            'chkInHouse.Visible = True
            'chkUATP.Value = 0
            Select Case Index
                Case 0
                    cmbFOP(1).Visible = False
                    txtTktMod(1).Visible = False
                    dtpCCExpDate(0).Visible = False
                    cmbFOP(2).Enabled = True
                    If cmbFOP(2).Text = "CC" Then
                        txtTktMod(2).Visible = True
                        cmbFOP(3).Visible = True
                        dtpCCExpDate(1).Visible = True
                        chkUATP.Visible = True
                        chkUATP.value = 1
                    End If
                                      
                'Case 2
                '    txtTktMod(2).Visible = False
                '    cmbFOP(3).Visible = False
                '    dtpCCExpDate(1).Visible = False
            End Select

    End Select
    'only check for changes in cmbFOP(1)
    Select Case Index
        Case 0, 1
            UATPControl (1)
    End Select
    
End Sub

Private Sub cmbValCarrier_Click()
If mbolFormLoaded Then blnChgPCarrier = True
End Sub

Private Sub cmdAdd_Click()
Dim strTemp As String
Dim strMsg As String
Dim intI As Integer
Dim strTax() As String
Dim bolFound As Boolean

If Len(Trim(txtTaxAmt)) = 0 Then strMsg = "Need tax amount" & vbCrLf
If Len(Trim(txtTaxCode)) = 0 Then strMsg = strMsg & "Need tax code" & vbCrLf
If Len(Trim(txtTaxAmt)) > 0 And CSng(txtTaxAmt) <= 0 Then
   strMsg = strMsg & "Tax Amount must greater than 0" & vbCrLf
End If
For intI = 0 To lstTax.ListCount - 1
    strTax = Split(lstTax.List(intI), " ")
    If UBound(strTax) > 0 Then
       If UCase(Trim(txtTaxCode)) = strTax(1) Then
          bolFound = True
          Exit For
       End If
    End If
Next
If bolFound = True Then
   strMsg = strMsg & "Duplicate tax code in list" & vbCrLf
End If

If strMsg = "" Then
    strTemp = Format(txtTaxAmt, gstrAgcyCurrFormat) & " " & UCase(txtTaxCode)
    lstTax.AddItem strTemp
    txtTaxAmt = ""
    txtTaxCode = ""
Else

    'MsgBox strMsg
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
End If

End Sub

Private Sub cmdBack_Click()
    sstTabs.Tab = sstTabs.Tab - 1
End Sub

Private Sub cmdCancel_Click()
Dim MsgBoxResponse As Integer

On Error Resume Next

If fWantToQuit Then
    Unload Me
    'Call pRedisplayMenu
Else
    Exit Sub
End If
    
End Sub

Private Sub cmdClientMI_Click()
    Call loadClientMI
End Sub

Private Sub cmdDone_Click()

Dim strMsg As String
Dim strFareCalc As String

   ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    
    If gIntModuleType = gModuleType.SYEX Then
        ' do not run GetVL checking to avoid error prompting
    Else
        If GetVL = False Then Exit Sub
    End If

datTouchEnd = Now
SysStart = Now
'If UCase(GetSetting("TPro", "Startup", "CountryCode", "NOT FOUND")) = "SG" Then
If gstrAgcyCountryCode = "SG" Then

   If cmbFOP(0).Text = "INV" Then
      If cmbFareType.Text = "APF - SPECIAL FARE" And _
         (chkNRCC.value = 1 Or fConvertZero(txtAComm) <> 0) Then
         sstTabs.Tab = 1
         'MsgBox "Cannot tick 'NRCC' check box or fill in A. Comm amount if FOP on tkt is not CC"
         strMsg = "Cannot tick 'NRCC' check box or fill in A. Comm amount if FOP on tkt is not CC"
         modMsgBox.OKMsg = "OK"
         modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
         Exit Sub
      End If
      
   'ElseIf cmbFOP(0).Text = "CC" And _
          (Left(UCase(cmbFOP(1).Text), 2) = "DC" And _
          Left(UCase(txtTktMod(1).Text), 7) = "3644033") And _
          (chkNRCC.value = 1 Or fConvertZero(txtAComm) <> 0) Then
     'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    ElseIf cmbFOP(0).Text = "CC" And _
          IsTMPCard(Left(UCase(cmbFOP(1).Text), 2), UCase(txtTktMod(1).Text)) And _
          (chkNRCC.value = 1 Or fConvertZero(txtAComm) <> 0) Then
          sstTabs.Tab = 1
          'MsgBox "Cannot tick 'NRCC' check box or fill in A. Comm amount if FOP on tkt is not CC"
          strMsg = "Cannot tick 'NRCC' check box or fill in A. Comm amount if FOP on tkt is not CC"
          modMsgBox.OKMsg = "OK"
          modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
          Exit Sub
   End If
 
   If cmbFareType.listindex = -1 Then
      sstTabs.Tab = 1
      'MsgBox "Missing Fare Type.."
      strMsg = "Missing Fare Type.."
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
      Exit Sub
   End If
   SGFF
   'Write to SQL Log
   'WriteToLog
   Exit Sub
End If

gobjLog.EventToLog "frmPricingWiz1.cmdDone_Click"

'On Error GoTo ProcErr
' #### need to move the mbolsysendo to bottom

    If Not validData Then Exit Sub
    frmWait.Show
    If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
    
    Call pCheckFBC
    gobjLog.LineTextToLog "mstrFBUFields = " & mstrFBUFields
   
    'FQ   1 time only
    'Use 1st Px Account code
    '29122004
    'If cmbPx.ListIndex = 0 Then
    If mbolFQ = False Then
       If Not FileFare Then
           Unload frmWait
           Exit Sub
       End If
       gobjLog.LineTextToLog "FileFare = True"
    End If
    
    gobjPNR.LoadFiledFare (CStr(mbytFFNum))
    
    'If Not validMI Then
    '    If cmbPx.listindex = mintStartPxNum Then
    '       gobjHost.terminalEntry "FX" & mbytFFNum
    '    If blnFirstPx = True Then mbolFQ = False
    '       strMsg = gobjHost.terminalEntry("R.TPRO PRICE")
    '       strMsg = gobjHost.terminalEntry("ER")
    '       strMsg = gobjHost.terminalEntry("ER")
    '       strMsg = gobjHost.terminalEntry("ER")
    '    End If
    '    Unload frmWait
    '    Exit Sub
    'End If
    
    'Call UpdateTktData
    With txtFareInfo(6)
        strFareCalc = gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).FareConstructText
        .Text = Mid(strFareCalc, 1, IIf((InStr(1, strFareCalc, "END ROE") - 1) > 0, InStr(1, strFareCalc, "END ROE") - 1, Len(strFareCalc)))
        '.Text = gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).FareConstructText
        If txtTktMod(6).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(6).Text)
        If txtTktMod(7).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(7).Text)
        If txtTktMod(8).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(8).Text)
    End With
    
    If Not UpdateTktData Then
       gobjHost.terminalEntry "FBF", True
       If cmbPx.listindex = mintStartPxNum Then
          gobjHost.terminalEntry "FX" & mbytFFNum
          If blnFirstPx = True Then mbolFQ = False
          blnCancelFF = True
           ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
          If gIntModuleType = gModuleType.SYEX Then
                ' do not ER for SyEx flow
          Else
                'gobjHost.terminalEntry ("R.TPRO PRICE")
                'gobjHost.terminalEntry ("ER")
                'gobjHost.terminalEntry ("ER")
                'gobjHost.terminalEntry ("ER")
                ' ZhiSam - V1.2.20 20130528 - IR-54 - Bug Fix for Desktop ER
                If Not gobjHost.ENDPNR("TPRO PRICE", True) Then
                    strMsg = "Unable to end transaction. Please end transaction before proceeding."
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - End PNR"
                    'MsgBox "Unable to end transaction. Please end transaction before proceeding."
                End If
          End If
          

        End If
        Unload frmWait
        Exit Sub
    End If
    gobjPNR.LoadFiledFare (CStr(mbytFFNum)), True
    
    If Not AddTMU Then
        gobjHost.terminalEntry "FX" & mbytFFNum
        If blnFirstPx = True Then mbolFQ = False
        'blnCancelFF = True
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
          If gIntModuleType = gModuleType.SYEX Then
                ' do not ER for SyEx flow
          Else
                'strMsg = gobjHost.terminalEntry("R.TPRO PRICE")
                'strMsg = gobjHost.terminalEntry("ER")
                'strMsg = gobjHost.terminalEntry("ER")
                'strMsg = gobjHost.terminalEntry("ER")
                ' ZhiSam - V1.2.20 20130528 - IR-54 - Bug Fix for Desktop ER
                If Not gobjHost.ENDPNR("TPRO PRICE", True) Then
                    strMsg = "Unable to end transaction. Please end transaction before proceeding."
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - End PNR"
                    'MsgBox "Unable to end transaction. Please end transaction before proceeding."
                End If
          End If

        Unload frmWait
        Exit Sub
    End If
    gobjLog.LineTextToLog "AddTMU = True"
        
    Call pFillInFFData
        
    'If Me.optCommissionType(1).Value Or Me.optDiscType(1).Value Then
    '    Call AddNF
    'End If
   
    'Call UpdateTktData
    mbolFareStored = True
    
    Call writeDatatoGDS
    
    Unload frmClientMI
    
    If cmbPx.listindex <> cmbPx.ListCount - 1 Then
       cmbPx.listindex = cmbPx.listindex + 1
       ClearVar
       clearControls
       Call pSetInitialValues
       blnFirstPx = False
       blnCancelFF = False
       Unload frmWait
       Exit Sub
    End If
    
    
    'Added on 31 Oct 2007. Requested by HK to has delivery address feature as SG
    Call updateDeliveryAddr
    'Added on 1/2/05: flag NRCC for Aqua process
    flagNRCC (mbytFFNum)
    gobjHost.ENDPNR "TPRO PRICE", True
    
    'pSendToFP "*FF" & mbytFFNum
    'pDisplayToFP "*FF" & mbytFFNum
    
    'Added on 22/07/04 - include remarks for AQUA checking
    gobjHost.terminalEntry "NP.S*APMI SCRIPT COMPLETED+NP.SS*VBIFF"

    pDisplayToFP "*FF" & mbytFFNum
    
    'Added on 1/2/05 - indicate file fare completed by VBI
    'gobjHost.TerminalEntry "NP.SS*VBIFF"

    'Added on 14/10/04: add to VBI log table
    'Timer
    Call pAddToVBILog(gobjPNR.RecLoc, "File Fare", startTime, SysStart, "File Fare", , startTime)

    
    'added on 6/1/05: delete FQ record after file fare
    If cmbPx.listindex = cmbPx.ListCount - 1 And Not gobjFareQuotes(1).FQ(1).StoreFare Then
        Call delFQRec
    End If
    
    'Write to SQL Log
    WriteToLog
    
    
    'Added on 26/08/04: Queue to ticketing after file fare
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
          If gIntModuleType = gModuleType.SYEX Then
                ' Remove Tiket Queue Form from the SyEx flow
          Else
                Load frmTktQueue
                frmTktQueue.Show
                Do
                        DoEvents
                Loop Until isLoaded("frmTktQueue") = False
          End If
          

        
    Unload Me
    Unload frmWait
    

    
    'If Pretrip exist, load frmPreTtrip
    'If CheckPreTrip = True Then
    '    frmPreTrip.Show
    'End If
    'Call pRedisplayMenu

Exit Sub
ProcErr:
Select Case Err.Number
    Case vbObjectError + 615
        'MsgBox "Unexpected response from GDS!" & vbCrLf & "You will need to finish manually."
        strMsg = "Unexpected response from GDS!" & vbCrLf & "You will need to finish manually."
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Case Else
        'MsgBox "ERROR " & Err.Number & vbCrLf _
            & Err.Description, "RUN TIME ERROR"
        strMsg = "ERROR " & Err.Number & vbCrLf _
            & Err.Description
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
            Resume Next
    End Select

Unload frmClientMI
Unload Me
'Call pRedisplayMenu
'Exit Function

End Sub

Private Sub cmdFCEndorse_Click()
    Load frmEndorseList
    frmEndorseList.txtEndorse.Text = Replace(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(1).Endorsement(1), "^", vbCrLf)
    frmEndorseList.txtEndorse.Enabled = False
    frmEndorseList.Show
    Do
       DoEvents
    Loop Until isLoaded("frmEndorseList") = False
End Sub

Private Sub cmdNext_Click()
If mbolFareStored Then
    cmdDone_Click
Else
    sstTabs.Tab = sstTabs.Tab + 1
End If
End Sub

Private Sub Form_Activate()
    
'Added on 26/10/04: Check for existing File Fares
Dim strMsg As String

If mbolActivate Then Exit Sub

datFormLoadStart = Now

mbolActivate = True

txtTransFee = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TransactionFee

If gstrAgcyCountryCode = "SG" Then
   lblTransFee.Visible = True
   txtTransFee.Visible = True
   chkNRCC.Visible = False
   chkNRCC.Enabled = False
   chkNRCCAC.Visible = False
   lblAComm.Visible = False
   txtAComm.Visible = False
   txtAComm = ""
End If

If cmbPx.listindex = 0 Then
If FFExist() Then
    strMsg = "Existing Filed Fares exist, Do you want to continue?"
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop - Reminder") = vbNo Then
    'If Not MsgBox("Existing Filed Fares exist, Do you want to continue?", vbYesNo, "CWT TravelPro") = vbYes Then
        Unload Me
        'Call pRedisplayMenu
        Exit Sub
    End If
End If
     
    
End If
mbolFormLoaded = True
 sstTabs.Tab = 0
    '29122004
    If cmbPx.listindex = mintStartPxNum And blnCancelFF <> True Then mbytFFNum = gobjPNR.FiledFareCount + 1
    If blnCancelFF = True Then blnCancelFF = False
    
    'If cmbPx.ListIndex = 0 Then mbytFFNum = gobjPNR.FiledFareCount + 1
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim strAL() As String
Dim strSQL As String
Dim rsRec1 As ADODB.Recordset
Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
     SwitchWinSetting (Me.hwnd)
    
Me.Move 0, 0
Me.Move frmSideBar.Width, 0
blnChgPCarrier = False
'Timer
startTime = Now

'added on 25/1/2005
gbolFMR = False

'added on 22/12
mintDisplayNo = getCurrDispNum() + 1
blnLoaded = True
blnFirstPx = True
blnCancelFF = False


Screen.MousePointer = vbDefault

strAL = GetALCodes
With cmbValCarrier
    For intX = 0 To UBound(strAL)
        .AddItem strAL(intX)
    Next
    .listindex = 0
End With

With dtpCCExpDate(0)
    .MinDate = DateSerial(2000, 1, 1)
    .MaxDate = DateAdd("yyyy", 10, Date)
    .value = DateSerial(2000, 1, 1)
End With
With dtpCCExpDate(1)
    .MinDate = DateSerial(2000, 1, 1)
    .MaxDate = DateAdd("yyyy", 10, Date)
    .value = DateSerial(2000, 1, 1)
End With

frmPricingWiz1.cmbPx.Clear
'Modified on 10/1/05: split farequote & filefare
If Not gobjFareQuotes(1).FQ(1).StoreFare Then
    For intX = 1 To gobjFareQuotes.PxCount
        For intY = 1 To gobjPNR.PassengerCount
         With gobjPNR.PassengerName(intY)
         
            If .PassengerNum = gobjFareQuotes(intX).FQ(1).PxNum Then
               cmbPx.AddItem Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
                If gobjFareQuotes(intX).FQ(1).PIC = "" Or gobjFareQuotes(intX).FQ(1).PIC = "VAC" Then
                     frmPricingWiz1.cmbPx.ItemData(frmPricingWiz1.cmbPx.NewIndex) = 1
                Else
                     frmPricingWiz1.cmbPx.ItemData(frmPricingWiz1.cmbPx.NewIndex) = 0
                End If

               Exit For
            End If
         End With
        Next intY
    Next intX
       
Else
    For intX = 0 To frmFareQuote.cmbPx.ListCount - 1
        'frmPricingWiz1.cmbPx.AddItem frmFareQuote.cmbPx.List(intX)
        '29122004
        frmPricingWiz1.cmbPx.AddItem frmFareQuote.cmbPx.List(intX)
        frmPricingWiz1.cmbPx.ItemData(frmPricingWiz1.cmbPx.NewIndex) = frmFareQuote.cmbPx.ItemData(intX)
    Next
End If
frmPricingWiz1.cmbPx.listindex = 0

'29122004
mintStartPxNum = 0
If gbolSkipAdult Then
   For intX = 0 To cmbPx.ListCount - 1
       If cmbPx.ItemData(intX) = 0 Then
          frmPricingWiz1.cmbPx.listindex = intX
          mintStartPxNum = intX
          Exit For
       End If
   Next
End If

'enabling or disabling value code box
'using Select case to add future countries
Select Case gstrAgcyCountryCode
    Case "HK"
        txtTktMod(9).Visible = True
        lblLabels(19).Visible = True
        txtFareInfo(8).Visible = False
        chkDocFee.Visible = False
        lblLabels(37).Visible = False
        'chkInHouse.Visible = True
    Case Else
        txtTktMod(9).Visible = False
        lblLabels(19).Visible = False
        'chkInHouse.Visible = False
End Select

With cmbFareOnTkt
    .AddItem "PUB - PUBLISHED FARE"
    .AddItem "APF - PRIVATE FARE"
    .AddItem "ITN - IT IN FARE BOX WITH NO FARE CALC"
    .AddItem "BTN - BT IN FARE BOX WITH NO FARE CALC"
    .AddItem "ITC - IT IN FARE BOX WITH FARE CALC"
    .AddItem "BTC - BT IN FARE BOX WITH FARE CALC"
    
End With

'    cmbFOP(0).AddItem "MULTI"
'Added on 21/07/04
'populate combo box values for Class of Services, Trip Type
'cboClassServ.AddItem ""
cboClassServ.AddItem "FF - First Class Full Fare "
cboClassServ.AddItem "FD - First Class Discounted Fare"
'cboClassServ.AddItem "FN"   'First Class - Nett Fare
cboClassServ.AddItem "FC - First Class Corporate Fare"
cboClassServ.AddItem "FW - First Class CWT Negotiated Fare"
cboClassServ.AddItem "CF - Business Class Full Fare"
cboClassServ.AddItem "CD - Business Class Discounted Fare"
'cboClassServ.AddItem "CN"
cboClassServ.AddItem "CC - Business Class Corporate Fare"
cboClassServ.AddItem "CW - Business Class CWT Negotiated Fare"
cboClassServ.AddItem "YF - Economy Class Full Fare"
cboClassServ.AddItem "YD - Economy Class Discounted Fare"
'cboClassServ.AddItem "YN"
cboClassServ.AddItem "YC - Economy Class Corporate Fare"
cboClassServ.AddItem "YW - Economy Class CWT Negotiated Fare"
cboClassServ.listindex = 0
'cboMIFareType.AddItem ""
'cboMIFareType.AddItem "W- CWT Negotiated Fares"
'cboMIFareType.AddItem "C- Client Negotiated Fares"
'cboMIFareType.listindex = 0

'CS - Remove FF26 (Trip Type)
'cboTripType.AddItem ""
'cboTripType.AddItem "Round"
'cboTripType.AddItem "One Way"

'CS - Add International or Domestic
cboTrip.AddItem "International"
cboTrip.AddItem "Domestic"
cboTrip.listindex = 0
 
'CS - Add Booking Method
'cboBookingMethod.AddItem "GDS"
'cboBookingMethod.AddItem "Manual"
'cboBookingMethod.AddItem "Self Booking"
'cboBookingMethod.listindex = 0

'CS Add Booking Tool
'cboBookingAction.AddItem "Agent Booked"
'cboBookingAction.AddItem "Self Booked"
'cboBookingAction.AddItem "Air Modified"

mstrBookingTool = GetBookingTool
If mstrBookingTool <> "" Then
   'cboBookingMethod.Text = "Self Booking"
   'cboBookingMethod.Locked = True
   'cboBookingAction.Locked = False
   'cboBookingAction.AddItem "Agent Booked"
   cboBookingAction.AddItem "EB - Self Booked"
   cboBookingAction.AddItem "AA - Air Modified"
   cboBookingAction.AddItem "AM - Multiple Modification"
   cboBookingAction.Text = "EB - Self Booked"
   

Else
   'cboBookingMethod.Locked = False
   'cboBookingAction.Text = "Agent Booked"
   'cboBookingAction.Locked = True
  cboBookingAction.AddItem "AB - Agent Booked"
  'cboBookingAction.AddItem "Self Booked"
  'cboBookingAction.AddItem "Air Modified"
  cboBookingAction.Text = "AB - Agent Booked"
  
End If
'cboBookingAction.listindex = -1


If UCase(gstrAgcyCountryCode) = "HK" Then
    lblFareType.Visible = False
    cmbFareType.Visible = False
    lblExcTax.Visible = False
    cmbCountry.Visible = False
    
    chkNRCC.Visible = False
    'chkNRCC.Left = 7320
    'chkNRCC.Top = 2280
    'txtTransFee.Visible = True
    'txtTransFee.Left = 5520
    'txtTransFee.Top = 2280
    'lblTransFee.Visible = True
    'lblTransFee.Top = 2280
    'lblTransFee.Left = 4300
    'lblTransFee.Font = "Tahoma"
    'lblTransFee.FontBold = True
    'lblTransFee.FontSize = 10
    'lblTransFee.Alignment = 2
    'lblTransFee.Caption = "TransFee"
    
    'txtAComm.Visible = True
    'lblAComm.Visible = True
    'chkNRCCAC.Visible = True
    'chkNRCCAC.Left = 7320
    'chkNRCCAC.Top = 2640
    'txtAComm.Left = 5520
    'txtAComm.Top = 2640
    'lblAComm.Left = 4500
    'lblAComm.Top = 2640
    'lblAComm.Font = "Tahoma"
    'lblAComm.FontBold = True
    'lblAComm.FontSize = 10
    'lblAComm.Alignment = 2
    
    chkTransFee.Visible = True
Else
    With cmbFareType
       .Visible = True
       .Clear
    '   .AddItem "SQ/MI Published Nett Fare"
    '   .AddItem "SQ/BA/QF Corporate Fare"
    '   .AddItem "SQ/MI Published Fare"
    '   .AddItem "Published Fare"
    '   .AddItem "Special Fare"
  
    strSQL = "select * from tblFareType"
    Set rsRec1 = gdbConn.Execute(strSQL)
    
    While Not rsRec1.EOF
        .AddItem Trim(rsRec1!FareType) & " - " & Trim(rsRec1!FareDesc)
        rsRec1.MoveNext
    Wend
    
    rsRec1.Close
    
    End With
    
    lblFareType.Visible = True
    lblExcTax.Visible = True
    cmbCountry.Visible = True
    cmbCountry.AddItem ""
    cmbCountry.AddItem "US"
    cmbCountry.listindex = 0
    chkTransFee.Visible = False
    
End If


Call pSetInitialValues
'Added on 18/10/04: Temporary disable this column
txtTktMod(10).Enabled = False

'If gobjPNR.GASaleRecordCount = 0 Then
'   mintDisplayNo = 0
'Else
'   mintDisplayNo = gobjPNR.GASalesRecord(gobjPNR.GASaleRecordCount).DisplanyNo
'End If

datFormLoadEnd = Now
If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    If Not fWantToQuit Then
        Cancel = 1
    Else
        'Call pRedisplayMenu
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPricingWiz1 = Nothing
'Preethi - V1.1.1 20100831 - IR2 - Client MI screen is populated with old data
Unload frmClientMI
End Sub

Private Sub lstTax_Click()
Dim strTemp() As String
Dim intI As Integer
If lstTax.ListCount = 0 Then Exit Sub
For intI = 0 To lstTax.ListCount - 1
    If lstTax.Selected(intI) = True Then
        strTemp = Split(lstTax.List(intI), " ")
    End If
    
    If UBound(strTemp) >= 0 Then txtTaxAmt = strTemp(0)
    If UBound(strTemp) >= 1 Then txtTaxCode = strTemp(1)
    Exit For
Next
End Sub

Private Sub lstTax_DblClick()
Dim intI As Integer
If lstTax.ListCount = 0 Then Exit Sub
For intI = 0 To lstTax.ListCount - 1
    If lstTax.Selected(intI) = True Then
        lstTax.RemoveItem (intI)
        txtTaxAmt = ""
        txtTaxCode = ""
    End If

    Exit For
Next
End Sub

'CS Change EC
'Private Sub lvwECodes_DblClick()
'    txtMI(2).Text = lvwECodes.SelectedItem
'End Sub

Private Sub lvwMissECodes_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lvwMissECodes_ItemClick(ByVal item As MSComctlLib.ListItem)
   txtMS = lvwMissECodes.SelectedItem
End Sub

Private Sub lvwRealECodes_ItemClick(ByVal item As MSComctlLib.ListItem)
   txtRS = lvwRealECodes.SelectedItem
End Sub

Private Sub optCommissionType_Click(Index As Integer)

If Index = 1 Then
    txtFareInfo(4).Enabled = False
    txtFareInfo(3).MaxLength = 0
ElseIf optDiscType(1).value = False Then
    txtFareInfo(4).Enabled = True
    txtFareInfo(3).MaxLength = 2
    txtFareInfo(4).Text = Trim(Left(txtFareInfo(4).Text, 2))
    
End If

End Sub

Private Sub optDiscType_Click(Index As Integer)

If Index = 1 Then
    txtFareInfo(4).Enabled = False
ElseIf optDiscType(1).value = False Then
    txtFareInfo(4).Enabled = True
End If

End Sub

Private Sub sstTabs_Click(PreviousTab As Integer)
        cmdNext.Enabled = sstTabs.Tab - 3
        cmdBack.Enabled = sstTabs.Tab
End Sub

Private Function FileFare() As Boolean
Dim strCmd As String
Dim strRes As String
Dim strSegs() As String
Dim lngC As Long
Dim strCNX As String
Dim strSegX As String
Dim strSegO As String
Dim strPx As String
Dim strSeg As String
Dim strMsg As String


strPx = ""
For lngC = 1 To gobjFareQuotes.PxCount
    With gobjFareQuotes(lngC).FQ(1)
       strPx = strPx & IIf(strPx = "", "P", ".") & .PxNum & IIf((.PIC = "" Or .PIC = "AD" Or .PIC = "ADT"), "", "*" & .PIC)
    End With
Next
strCmd = "FQ" & "C" & cmbValCarrier.Text

'11012005 move -SGCWT (or LC) to infront (by Sok Leng)
If gobjFareQuotes(1).FQ(1).PrivateFare = True And cmbFareOnTkt.listindex >= 1 Then
    If gobjFareQuotes(1).FQ(1).PFAccountCode <> "" Then
        strCmd = strCmd & "-" & gobjFareQuotes(1).FQ(1).PFAccountCode
        If gobjFareQuotes(1).FQ(1).RuleID <> "" Then strCmd = strCmd & "@@" & gobjFareQuotes(1).FQ(1).RuleID
        If gobjFareQuotes(1).FQ(1).FQPCC <> "" And gobjFareQuotes(1).FQ(1).FQPCC <> gobjHost.AgentPCC Then strCmd = strCmd & "*" & gobjFareQuotes(1).FQ(1).FQPCC
    End If
Else
    If gobjFareQuotes(1).FQ(1).FQPCC <> "" And gobjFareQuotes(1).FQ(1).FQPCC <> gobjHost.AgentPCC Then strCmd = strCmd & "-*" & gobjFareQuotes(1).FQ(1).FQPCC
End If
'strCmd = strCmd & "/" & chkPaperTkt.Caption

'Modified on 17/2/2005: use FareType to differentiate PF cmd

'11012005 requested by Sok Leng
'If UCase(gstrAgcyCountryCode) = "SG" Then
'   If cmbFareType.Text = "Special Fare" And _
'      gobjFareQuotes(1).FQ(1).PrivateFare = True And _
'      cmbFareOnTkt.ListIndex >= 1 And _
'      UCase(Left(gobjFareQuotes(1).FQ(1).PFAccountCode, 2)) = "SG" Then
'      strCmd = strCmd & ":C::" & UCase(gstrAgcyCurrCode)
'   End If
'End If

If UCase(gstrAgcyCountryCode) = "SG" Then
   'If cmbFareType.Text = "Special Fare" And
   If gobjFareQuotes(1).FQ(1).PrivateFare = True And _
      cmbFareOnTkt.listindex >= 1 Then
      
      'If UCase(gobjFareQuotes(1).FQ(1).PFFareType) = "APF" Then
      '      strCmd = strCmd & ":C::" & UCase(gstrAgcyCurrCode)
      'ElseIf UCase(gobjFareQuotes(1).FQ(1).PFFareType) = "SQC" Then
            strCmd = strCmd & ":P" ' ::" & UCase(gstrAgcyCurrCode)
      'End If
   End If
End If


With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
      
    'added on 17/01/2005: to get the exact seg no.
    If .SegmentSelected = True And .StoreFare = False Then
        .SegmentSelectString = getSegmentSelected
    End If
    
    'If .PIC <> "AD" Then
    '    strCmd = strCmd & "*" & .PIC
    'End If
    strCmd = strCmd & "/" & strPx
    
    'If gbolOverrideFare Then
    '   GoTo SkipPF
    'End If
    
    'If .PrivateFare = True And cmbFareOnTkt.ListIndex >= 1 Then strCmd = strCmd & "-" & .PFAccountCode
    
    'If .PrivateFare = True Then
'SkipPF:
    
         If cmbFareOnTkt.listindex < 1 Then   'PUB
            
            If .PrivateFare = True Then
                'added on 6/12: for HKG - VUSA
                If gstrAgcyCountryCode = "HK" And txtPriceFBC(0) = "VU" Then
                    strCmd = strCmd & "*VU"
                End If
            End If
            strSeg = ""
            'strCmd = strCmd & "/S"
            
            For lngC = 0 To mbytNumSegs - 1
                'added on 15/09/04: check for missing FBC
                If txtPriceFBC(lngC) = "" Then
                    'modified on 19/5/2005: Default PriceFBC to same as FBC
                    If .PrivateFare = True Then txtPriceFBC(lngC).Text = txtFBC(lngC).Text
                    strSeg = strSeg & IIf(strSeg = "", "/S", ".") & txtSegNum(lngC).Text & IIf(txtPriceFBC(lngC).Text <> "", "@" & txtPriceFBC(lngC).Text, "")
                    'MsgBox "Missing Price FBC."
                    'FileFare = False
                    'Exit Function
                Else
                    'added on 6/12: for HKG - VUSA
                    If gstrAgcyCountryCode = "HK" And txtPriceFBC(0) = "VU" Then
                        If .PrivateFare = True Then
                            strSeg = strSeg & IIf(strSeg = "", "/S", ".") & txtSegNum(lngC).Text
                            'strCmd = strCmd & IIf(lngC > 0, ".", "") & txtSegNum(lngC).Text
                        End If
                    Else
                        'strCmd = strCmd & IIf(lngC > 0, ".", "") & txtSegNum(lngC).Text & "@" & txtPriceFBC(lngC)
                        strSeg = strSeg & IIf(strSeg = "", "/S", ".") & txtSegNum(lngC).Text & "@" & txtPriceFBC(lngC)
                    End If
                End If
            Next
            
            
            'If .SegmentSelected = True Then
            'strCmd = strCmd & "/S" & .SegmentSelectString
            'End If
            
            
         'End If
         
         strCmd = strCmd & strSeg
            'Exclude for VUSA Fare
            If Not (gstrAgcyCountryCode = "HK" And txtPriceFBC(0) = "VU") Then
                strCmd = strCmd & "/:N"
            End If
            'If .SegmentSelected = True Then
            '    strCmd = strCmd & "/S" & .SegmentSelectString
            'End If
        'ElseIf .SegmentSelected = True Then
        '    strCmd = strCmd & "/S" & .SegmentSelectString
        'End If
    ElseIf .SegmentSelected = True Then
        strCmd = strCmd & "/S" & .SegmentSelectString
    End If
End With


'13/08/04 Request from HK & SG users: this option(/*RP) is not required
'If txtFareInfo(7).Text <> "0" And txtFareInfo(7).Text <> "" Then
'    strCmd = strCmd & "/*RP" & txtFareInfo(7).Text
'End If

'Added on 01/10/04 for Stopover/Transit
strSegX = ""
strSegO = ""
strCNX = ""

If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).OverrideConx Then
'If InStr(mstrFFConn, "/CNX") > 0 Then

    For lngC = 0 To mbytNumSegs - 1
        If chkConnection(lngC).value = 1 Then
            strSegX = strSegX & IIf(Len(strSegX) > 0, ".", "") & txtSegNum(lngC)
        Else
            strSegO = strSegO & IIf(Len(strSegO) > 0, ".", "") & txtSegNum(lngC)
        End If
    Next
    strCNX = IIf(Len(strSegX) > 0, "X" & strSegX, "") & IIf(Len(strSegO) > 0, "O" & strSegO, "")
    strCmd = strCmd & IIf(Len(strCNX) > 0, "/" & strCNX, "")
End If


'added on 26/7/2005: for VUSA fare to exclude tax
If cmbCountry.Text <> "" Then
    strCmd = strCmd & "/TE-" & Trim(cmbCountry.Text)
End If



strCmd = strCmd & "/" & chkPaperTkt.Caption

'11012005 move -SGCWT (or LC) to infront (by Sok Leng)
'If gobjFareQuotes(1).FQ(1).PrivateFare = True And cmbFareOnTkt.ListIndex >= 1 Then strCmd = strCmd & "-" & gobjFareQuotes(1).FQ(1).PFAccountCode
''
''strCmd = strCmd & "/" & chkPaperTkt.Caption

''11012005 requested by Sok Leng
'If UCase(gstrAgcyCountryCode) = "SG" Then
'   If cmbFareType.Text = "Special Fare" And _
'      gobjFareQuotes(1).FQ(1).PrivateFare = True And _
'      cmbFareOnTkt.ListIndex >= 1 And _
'      UCase(Left(gobjFareQuotes(1).FQ(1).PFAccountCode, 2)) = "SG" Then
'      strCmd = strCmd & ":C::" & UCase(gstrAgcyCurrCode)
'   End If
'End If

strRes = gobjHost.terminalEntry(strCmd, True)
lngC = Len(strCmd)
'29122004
'MsgBox strCmd
If Left(UCase(Replace(strRes, vbCrLf, "")), lngC) = UCase(strCmd) And valError(strRes) Then
    FileFare = True
    mbolFQ = True
Else
    FileFare = False
    mbolFQ = False
    'MsgBox "Unable to file fare!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
    strMsg = "Unable to file fare!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End If

End Function

Private Function AddNF() As Boolean
Dim sngBF As Single
Dim sngNF As Single
Dim sngASF As Single
Dim sngDisc As Single
Dim sngComm As Single
Dim strCmd As String
Dim strRes As String
Dim strMsg As String

sngBF = CSng(txtFareInfo(0).Text)
sngASF = sngBF

If txtFareInfo(3).Text = "" Then txtFareInfo(3).Text = "0"
If txtFareInfo(7).Text = "" Then txtFareInfo(7).Text = "0"

If optCommissionType(1).value = True Then
    sngNF = sngBF - CSng(txtFareInfo(3).Text)
Else
    sngComm = fCurrRound(sngBF - (sngBF * (1 - (CSng(txtFareInfo(3).Text) * 0.01))), gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency, "DOWN")
    sngNF = sngBF - sngComm
End If

If optDiscType(1).value = True Then
    sngDisc = fCurrRound(sngBF - (sngBF * (1 - (CSng(txtFareInfo(7).Text) * 0.01))), gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency, "DOWN")
    sngNF = sngNF - sngDisc
    sngASF = sngBF - sngDisc
End If

'Removed on 14/10/04: this is done in AddTMU
'strCmd = "TMU" & mbytFFNum & "/NF" & gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).TotalCurrency & Format(sngNF, gstrAgcyCurrFormat)

'Removed on 14/09/04: this is done in AddTMU
'If cmbFOP(0).Text = "CC" Then
'    strCmd = strCmd & "/ASF" & Format(sngASF, gstrAgcyCurrFormat)
'End If

strRes = gobjHost.terminalEntry(strCmd)

If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
    AddNF = True
ElseIf InStr(strRes, "ERROR 8516 - INVALID FORMAT/DATA - MODIFIER ALREADY EXISTS") > 0 Then
    gobjHost.terminalEntry "TMU" & mbytFFNum & "/NF@"
    gobjHost.terminalEntry "TMU" & mbytFFNum & "/ASF@"
    strRes = gobjHost.terminalEntry(strCmd)
    If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
        AddNF = True
    Else
        AddNF = False
        'MsgBox "Unable to add NF/ASF TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
        strMsg = "Unable to add NF/ASF TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End If
Else
    AddNF = False
    'MsgBox "Unable to add NF/ASF TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
    strMsg = "Unable to add NF/ASF TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End If

End Function

Private Function AddTMU() As Boolean
Dim strCmd As String
Dim strRes As String
Dim strMsg As String

'29122004
'If cmbPx.ListIndex = 0 Then
If cmbPx.listindex = mintStartPxNum And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
    If cmbFareOnTkt.listindex > 0 Then gobjHost.terminalEntry "TMU" & mbytFFNum & "/TC@"
End If

strCmd = ""
strCmd = "TMU" & mbytFFNum

If optCommissionType(0).value And ((gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False) Or (gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = True And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).CommType = False)) Then
    strCmd = strCmd & "/Z" & txtFareInfo(3).Text
End If

If txtFareInfo(4).Text <> "" And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
   strCmd = strCmd & "/NF" & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency & Format(txtFareInfo(4).Text, gstrAgcyCurrFormat)
End If

If txtTktMod(0).Text <> "" And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
    Select Case gstrAgcyCountryCode 'using select case to allow for more countries
        Case "HK"
            strCmd = strCmd & "/AI-" & txtTktMod(0).Text
    
            If txtTktMod(9).Text <> "" Then
                strCmd = strCmd & "/VC-" & txtTktMod(9).Text
            End If
        
        Case Else
            
            strCmd = strCmd & "/TC" & txtTktMod(0).Text
    End Select
End If

Select Case cmbFOP(0).Text
    Case "INV"
        strCmd = strCmd & "/F" & cmbFOP(0).Text & IIf(txtTktMod(3) <> "", txtTktMod(3), "AGT")
    Case "MS"
        strCmd = strCmd & "/F" & cmbFOP(0).Text & IIf(txtTktMod(3) <> "", txtTktMod(3), "")
    Case "CC"
        strCmd = strCmd & "/F" & cmbFOP(1).Text & txtTktMod(1).Text & "*D" & Format(dtpCCExpDate(0).value, "mmyy")
            'If txtFareInfo(4) <> "" And txtFareInfo(4).Enabled And chkShowASF.value = 1 Then
            If txtFareInfo(4) <> "" And chkShowASF.value = 1 Then
                'Added on 2/3/2005: If NRCC- ASF deduct disc
                With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
                    If .NRCC Then
                        If .TransactionFee > 0 Then
                            If .MerchAmt > 0 Then
                                If gobjPNR.CompInfo.TFIncMF Then
                                    strCmd = strCmd & "/ASF" & Format((fConvertZero(txtFareInfo(5).Text) - .DiscountAmt + .TransactionFee), gstrAgcyCurrFormat)
                                Else
                                    strCmd = strCmd & "/ASF" & Format((fConvertZero(txtFareInfo(5).Text) - .DiscountAmt + .TransactionFee - .MerchAmt), gstrAgcyCurrFormat)
                                End If
                            Else
                                If chkTransFee.value = 1 Then
                                    strCmd = strCmd & "/ASF" & Format((fConvertZero(txtFareInfo(5).Text) - .DiscountAmt + .TransactionFee), gstrAgcyCurrFormat)
                                Else
                                    strCmd = strCmd & "/ASF" & Format((fConvertZero(txtFareInfo(5).Text) - .DiscountAmt), gstrAgcyCurrFormat)
                                End If
                            End If
                        Else
                            strCmd = strCmd & "/ASF" & Format((fConvertZero(txtFareInfo(5).Text) - .DiscountAmt), gstrAgcyCurrFormat)
                        End If
                    Else
                        strCmd = strCmd & "/ASF" & Format(txtFareInfo(5).Text, gstrAgcyCurrFormat)
                    End If
                End With
                If gstrAgcyCountryCode = "HK" And cmbValCarrier.Text = "SQ" Then gobjHost.terminalEntry "N.P" & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PxNum & "@*NF"
            End If
End Select

If txtTktMod(4).Text <> "" Then
    strCmd = strCmd & "/EB" & txtTktMod(4).Text
End If
If txtTktMod(4).Text = "" And txtTktMod(5).Text <> "" Then
    strCmd = strCmd & "/EB" & txtTktMod(5).Text
End If
If txtTktMod(4).Text <> "" And txtTktMod(5).Text <> "" Then
    strCmd = strCmd & "*EB" & txtTktMod(5).Text
End If

'removed on 13/12: PT/ET is specified during FQ command
'strCmd = strCmd & IIf(chkPaperTkt.Value = vbChecked, "/PT", "/ET")

If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
    Select Case cmbFareOnTkt.listindex
        Case 2, 4
            strCmd = strCmd & "/IT" & IIf(cmbFareOnTkt.listindex = 4, "*PC", "")
        Case 3, 5
            strCmd = strCmd & "/BT" & IIf(cmbFareOnTkt.listindex = 5, "*PC", "")
    End Select
End If

'29122004
'If cmbPx.ListIndex = 0 Then
If cmbPx.listindex = mintStartPxNum Then

        strRes = gobjHost.terminalEntry(strCmd)
        If InStr(1, strRes, "DBI AIRPLUS INTERNATIONAL DESCRIPTIVE BILLING") > 0 Then
           AddTMU = False
           blnCancelFF = True
           'MsgBox "Unable to add TMU!" & Chr(13) & Chr(13) & "Please inform operation manager to disable the DBI screen in Galileo"
           strMsg = "Unable to add TMU!" & Chr(13) & Chr(13) & "Please inform operation manager to disable the DBI screen in Galileo"
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
           Exit Function
        End If
        If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
            AddTMU = True
        ElseIf InStr(strRes, "ERROR 8516 - INVALID FORMAT/DATA - MODIFIER ALREADY EXISTS") > 0 Then
            gobjHost.terminalEntry "TMU" & mbytFFNum & "/TC@"
            strRes = gobjHost.terminalEntry(strCmd)
            If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
                AddTMU = True
            Else
                AddTMU = False
                blnCancelFF = True
                'MsgBox "Unable to add TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
                strMsg = "Unable to add TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                Exit Function
            End If
        Else
            AddTMU = False
            blnCancelFF = True
            'MsgBox "Unable to add TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
            strMsg = "Unable to add TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
            Exit Function
        End If
        
    'FMR
    If cmbFOP(0).Text = "FMR" Then
       If Not FMR Then
          AddTMU = False
          blnCancelFF = True
          Exit Function
       End If
    End If
   
Else
            'Add to NP.TM*
            AddTMU = True
            Dim intPos As Integer
            Dim intStart As Integer
            Dim intLength As Integer
            Dim intPrevious As Integer
            Dim strPrevious As String
            Dim i As Integer
            Dim j As Integer
           'FMR
           If cmbFOP(0).Text = "FMR" Then
              strCmd = strCmd & "/FMR"
           End If
           
            
            If Len(strCmd) > 72 Then
            i = 0
            Do While Len(strCmd) > 72
                intPrevious = 0
                intPos = 0
                intStart = 0
                j = 0
                Do Until intPos > 72
                intPrevious = intPos
                    intPos = InStr(IIf(j = 0, 1, intPos + 1), strCmd, "/")
                    If intPos = 0 Then Exit Do
                    If j = 0 And intPos > 75 Then
                    intPrevious = 72
                    Exit Do
                    End If
                    j = j + 1
                Loop
                
        
                intPos = IIf(j = 0, intPrevious, intPrevious - 1)
                intLength = intPos - intStart
                strPrevious = strCmd
                strCmd = Mid(strCmd, 1, intLength)
                
                strCmd = "NP.TM*" & "FF" & Format(mbytFFNum, "00") & "PX" & Format(cmbPx.listindex + 1, "00") & ":" & ConvertNPText(strCmd)
                strRes = gobjHost.terminalEntry(strCmd)
                
                strCmd = Mid(strPrevious, IIf(j = 0, intPos + 1, intPos + 2))
        
                i = i + 1
                Loop
                
                strCmd = "NP.TM*" & "FF" & Format(mbytFFNum, "00") & "PX" & Format(cmbPx.listindex + 1, "00") & ":" & ConvertNPText(strCmd)
                strRes = gobjHost.terminalEntry(strCmd)
        
            
            Else
                strCmd = "NP.TM*" & "FF" & Format(mbytFFNum, "00") & "PX" & Format(cmbPx.listindex + 1, "00") & ":" & ConvertNPText(strCmd)
                strRes = gobjHost.terminalEntry(strCmd)
            End If
        End If
End Function

Private Function FFExist() As Boolean
    If gobjPNR.FiledFareCount > 0 Then
        FFExist = True
    Else
        FFExist = False
    End If
End Function
Private Function SellingFare() As Double
    SellingFare = CDec(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TaxTotal) + CDec(fConvertZero(txtFareInfo(5).Text)) - CDec(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).DiscountAmt)
    If SellingFare < 0 Then SellingFare = 0

End Function

Private Function validData() As Boolean
Dim strMsg As String
Dim intX As Integer
Dim bolPriceFBC As Boolean
Dim bolEC As Boolean
Dim intY As Integer
If Len(cmbValCarrier.Text) <> 2 Then strMsg = strMsg & "Need Validating Carrier..." & vbCrLf

'Added on 2/3/2005: Mandatory check for FOP from Pax
'Added on 8/11/2005: Validate CC no. in DI
If cmbFOP(2).listindex = -1 Then
    strMsg = strMsg & "Need FOP from Pax..." & Chr(13)
ElseIf cmbFOP(2).Text = "CC" And cmbFOP(0) <> "CC" Then
    
    If cmbFOP(3).Text = "" Or txtTktMod(2).Text = "" Then
        strMsg = strMsg & "Need Card Number for FOP from PAX..." & Chr(13)
    Else
       If ValidCCNum(cmbFOP(3).Text, txtTktMod(2).Text) = False Then strMsg = strMsg & "Credit card number is invalid(FOP from Pax)..." & Chr(13)
    End If

End If

If dtpCCExpDate(0).Visible = True Then
    If dtpCCExpDate(0).value <= Date Or dtpCCExpDate(0).value > dtpCCExpDate(0).MaxDate Then
        strMsg = strMsg & "Verify CC Expiration date..." & Chr(13)
    End If
End If

If dtpCCExpDate(1).Visible = True Then
    If dtpCCExpDate(1).value <= Date Or dtpCCExpDate(1).value > dtpCCExpDate(1).MaxDate Then
        strMsg = strMsg & "Verify CC Expiration date..." & Chr(13)
    End If
End If

'300908 Detect MI by CN
If gobjPNR.CompInfo.MI = True Then
        If txtMI(0).Text = "" Then strMsg = strMsg & "Need Reference Fare (MI)..." & Chr(13)
        If txtMI(1).Text = "" Then strMsg = strMsg & "Need Low Fare (MI)..." & Chr(13)
        If txtMI(0).Text <> "" And txtMI(1) <> "" Then
            If CDec(txtMI(0)) < CDec(txtMI(1)) Then
                strMsg = strMsg & "Reference Fare(MI) must be greater than Low Fare(MI)..." & Chr(13)
            End If
        End If
        'Chnage EC
        'If txtMI(2).Text = "" Then strMsg = strMsg & "Need Exception Code (MI)..." & Chr(13)
        If txtMS = "" Then strMsg = strMsg & "Need Missed Saving (MI)..." & Chr(13)
        If txtRS = "" Then strMsg = strMsg & "Need Realised Saving (MI)..." & Chr(13)
        'If txtRS <> "" And txtRS = "MC" Then strMsg = strMsg & "Realised Saving Code(MI) cannot be MC..." & Chr(13)
        'If txtMS <> "" And txtMS = "M" Then strMsg = strMsg & "Missed Saving Code(MI) cannot be M..." & Chr(13)
        If txtMI(1).Text <> "" And txtMS.Text = "L" Then
            If SellingFare - CDec(txtMI(1)) > 0 Then
                strMsg = strMsg & "Missed Saving Amount Exist(MI), incorrect Miss Saving Code(L)..." & Chr(13)
            End If
        End If
        'Preethi - V1.1.1 20100831 - CR15 - Reference Fare and Low Fare Validation
        If txtMI(0).Text <> "" Then
            If txtMI(0).Text < SellingFare Then
                strMsg = strMsg & "Reference Fare must be greater than or equal to Selling Fare" & Chr(13)
            End If
        End If
        If txtMI(1).Text <> "" Then
            If txtMI(1).Text > SellingFare Then
                'Preethi - V1.1.1 20100831 - CR15 - Reference Fare and Low Fare Validation
                strMsg = strMsg & "Low Fare must be lower than or equal to Selling Fare" & Chr(13)
            End If
        End If
        
        
        If txtMI(3).Text = "" Then strMsg = strMsg & "Need Final Desitnation (MI)..." & Chr(13)
        'If txtMI(6).Text = "" Then strMsg = strMsg & "Need Low Fare Carrier (MI)..." & Chr(13)
        'CS - Remove FF26 (Trip Type)
        'If cboTripType.Text = "" Then strMsg = strMsg & "Need Trip Type (MI)..." & Chr(13)
        'CS - Add International or Domestic
        If cboTrip.Text = "" Then strMsg = strMsg & "Need Trip Type (MI)..." & Chr(13)
        'CS Add Booking Method
        'If cboBookingMethod.Text = "" Then strMsg = strMsg & "Need Booking Method (MI)..." & Chr(13)
        'CS Add Booking Action
        If cboBookingAction.Text = "" Then strMsg = strMsg & "Need Booking Action (MI)..." & Chr(13)
        'If cboClassServ.Text = "" Then strMsg = strMsg & "Need Class Service (MI)..." & Chr(13)

End If

If cmbFOP(0).Text = "CC" Then
    If cmbFOP(1).Text = "" Or txtTktMod(1).Text = "" Then
        strMsg = strMsg & "Need Card Number..."
    End If
End If

If gstrAgcyCountryCode = "SG" And cmbFareType.Text = "APF - SPECIAL FARE" And NFNotFound_SG Then
   strMsg = strMsg & "Need Net Fare..." & Chr(13)
End If

'Added on 7/2/2005: checking for SQ/MI Published Fare, required Doc Fee and Nett Fare
'Modified on 27/9/2007: For SQ paper ticket, no need Nett Fare and show ASF
If gstrAgcyCountryCode = "SG" Then
    
    If cmbFareType.Text = "SQP - SQ/MI PUBLISHED FARE" Then
        If Trim(cmbValCarrier.Text) = "SQ" And chkPaperTkt.Caption = "PT" Then
           If txtFareInfo(4).Text <> "" Then strMsg = strMsg & "SQ Published Fare with PT do not require Nett Fare..." & Chr(13)
           If chkShowASF Then strMsg = strMsg & "SQ Published Fare with PT do not require to Show ASF..." & Chr(13)
           If txtFareInfo(8).Text <> "" Then strMsg = strMsg & "SQ Published Fare with PT do not require Doc Fee..." & Chr(13)
        Else
            If fConvertZero(txtFareInfo(4).Text) = 0 Then strMsg = strMsg & "Need Nett Fare..." & Chr(13)
            If fConvertZero(txtFareInfo(8).Text) = 0 And mstrFirstSeg = "SIN" Then
               strMsg = strMsg & "Need Doc Fee..." & Chr(13)
            ElseIf fConvertZero(txtFareInfo(8).Text) > 0 And mstrFirstSeg <> "SIN" Then
               strMsg = strMsg & "Doc Fee is not required..." & Chr(13)
            End If
        End If
        If fConvertZero(txtFareInfo(3).Text) > 0 Then strMsg = strMsg & "SQ/MI Published Fare do not require Commission(%)..." & Chr(13)
        If fConvertZero(txtFareInfo(7).Text) > 0 Then strMsg = strMsg & "SQ/MI Published Fare do not require Discount($)..." & Chr(13)
    End If
    
    If cmbFareType.Text = "PUB - PUBLISHED FARE" And fConvertZero(txtFareInfo(3).Text) > 0 Then
        If fConvertZero(txtFareInfo(4).Text) > 0 Then strMsg = strMsg & "Published Fare with commission do not require Nett Fare..." & Chr(13)
        If chkShowASF Then strMsg = strMsg & "Published Fare with commission do not require to Show ASF..." & Chr(13)
    End If

    If cmbFareType.Text = "APF - SPECIAL FARE" And chkNRCCAC.value = 1 And fConvertZero(txtAComm) = 0 Then
       strMsg = strMsg & "Need A.Comm. Amount..." & Chr(13)
    End If
    If cmbFareType.Text = "APF - SPECIAL FARE" And fConvertZero(txtFareInfo(4).Text) = 0 Then
       If fConvertZero(txtAComm) > 0 Then strMsg = strMsg & "Need Nett Fare..." & Chr(13)
    End If
    
    If cmbFareType.Text = "SQC - CORPORATE FARE (A)" And txtTktMod(0) = "" Then
       strMsg = strMsg & "Need Tour Code for SQ/BA/QF Corporate Fare..." & Chr(13)
    End If
Else
    If chkTransFee.value = 1 And chkUATP.value = 0 And cmbFOP(0).Text = "CC" Then
        strMsg = strMsg & "Cannot include Trans Fee in ASF if Airline FOP is not UATP..." & Chr(13)
    End If
End If

If cmbCountry.ListCount > 1 And cmbCountry.Text <> "" Then
    If Len(cmbCountry.Text) <> 2 Then
        strMsg = strMsg & "Invalid Tax Country Code..." & Chr(13)
    End If
End If

'Added on 4/2/2005: check for invalid EC
'CS Change EC
'bolEC = False
'For intX = 1 To lvwECodes.ListItems.count
'    If txtMI(2) = lvwECodes.ListItems.Item(intX) Then
'        bolEC = True
'        Exit For
'    End If
'Next intX
'If bolEC = False Then strMsg = strMsg & "Invalid Exception Code..."
'300908 Detect MI by CN
If gobjPNR.CompInfo.MI = True Then
        bolEC = False
        For intX = 1 To lvwRealECodes.ListItems.Count
           If txtRS = lvwRealECodes.ListItems.item(intX) Then
              bolEC = True
              Exit For
           End If
        Next
        If bolEC = False Then strMsg = strMsg & "Invalid Realised Saving Code..."
        bolEC = False
        For intX = 1 To lvwMissECodes.ListItems.Count
           If txtMS = lvwMissECodes.ListItems.item(intX) Then
              bolEC = True
              Exit For
           End If
        Next
        If bolEC = False Then strMsg = strMsg & "Invalid Missed Saving Code..."
End If
'Added on 19/08/04
'Removed on 15/10/04
'If gstrAgcyCountryCode = "SG" And NFNotFound_SG Then
'    strMsg = strMsg & "Need NF/SF ..."
'End If
'

'For intX = 0 To mbytNumSegs - 1
'    If txtPriceFBC(intX) <> "" Then
'        If bolPriceFBC Or txtPriceFBC(0) <> "" Then strMsg = strMsg & "Must fill in or delete all Price FBC..." & vbCrLf
'        bolPriceFBC = True
'    End If
' Next

'Added on 31 Oct 2007. Only 6 subfields are allow in delivery address and each subfield allows 37 characters only
If Len(Trim(txtDeliveryDate.Text)) > 0 Then
    'Detect total number of subfields in delivery date field and existing delivery address field
    intX = 0
    intY = 0
    Do While InStr(intX + 1, Trim(txtDeliveryDate.Text), "*") > 0
       intX = InStr(intX + 1, Trim(txtDeliveryDate.Text), "*")
       intY = intY + 1
    Loop
    intX = 0
    Do While InStr(intX + 1, gobjPNR.DeliveryAddress, "@") > 0
       intX = InStr(intX + 1, gobjPNR.DeliveryAddress, "@")
       intY = intY + 1
    Loop
    If intY > 5 Then strMsg = strMsg & "Only 6 subfields are allow in delivery address field..." & Chr(13)
    
   'Detect number of characters between * (Cannot more than 37 characters)
   intX = 0
   intY = 0
   Do While InStr(intX + 1, Trim(txtDeliveryDate.Text), "*") > 0
      intY = intX + 1
      intX = InStr(intX + 1, Trim(txtDeliveryDate.Text), "*")
      
      If intX > 0 Then
         If Len(Mid(Trim(txtDeliveryDate.Text), intY, intX - intY)) > 37 Then
            strMsg = strMsg & "Only 36 characters are allow in each subfield of delivery address field..." & Chr(13)
            Exit Do
         End If
      Else
         Exit Do
      End If
   Loop
   
   'Detect number of charaters after * that will be appended with existing delivery address field
   intX = 0
   intY = 0
   If Right(Trim(txtDeliveryDate.Text), 1) <> "*" Then
      Do While InStr(intX + 1, Trim(txtDeliveryDate.Text), "*") > 0
         intX = InStr(intX + 1, Trim(txtDeliveryDate.Text), "*")
      Loop
      If intX = 0 Then
         intY = Len(Trim(txtDeliveryDate.Text))
      Else
         intY = Len(Trim(Mid(txtDeliveryDate.Text, intX + 1)))
      End If
   
      intX = InStr(1, gobjPNR.DeliveryAddress, "@")
      If intX = 0 Then
         If intY + Len(" " & gobjPNR.DeliveryAddress) > 37 Then
            strMsg = strMsg & "Only 36 characters are allow in each subfield of delivery address field..." & Chr(13)
         End If
      Else
         If intY + Len(" " & Mid(gobjPNR.DeliveryAddress, 1, intX - 1)) > 37 Then
            strMsg = strMsg & "Only 36 characters are allow in each subfield of delivery address field..." & Chr(13)
         End If
      End If
   End If
End If
If strMsg <> "" Then
    'MsgBox strMsg
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    validData = False
Else
    validData = True
    'Check for completion of client MI
    If isRequireClientMI(gobjPNR.CN, 4) Then
        If isLoaded("frmClientMI") Then
            'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
            'If frmClientMI.bolCheck = False Then
                strMsg = frmClientMI.incompleteMI
                If strMsg <> "" Then
                    validData = False
                    'MsgBox strMsg, vbCritical
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
                    loadClientMI
                End If
           ' End If
        Else
                validData = False
                'MsgBox "Client MI data is incomplete...", vbCritical
                strMsg = "Client MI data is incomplete..."
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
                loadClientMI
        End If
    End If
End If


End Function

Private Sub pSetInitialValues()
Dim strSQL As String
Dim rsECodes As New ADODB.Recordset
Dim strContract As String
Dim lngC As Long
Dim lngY As Long
Dim item As ListItem
Dim intFFFound As Integer
Dim blnChkVendor As Boolean
Dim blnChkDate As Boolean
Dim strClass() As String

Dim strPFBC() As String
Dim bolFound As Boolean
strSegmentNo = "/SG"
mbytNumSegs = 0
'Modified on 12/1/2005: for Fare Fare Only Option
If Not gobjFareQuotes(1).FQ(1).StoreFare Then
    For lngY = 1 To gobjFareQuotes(1).FQ(1).FareSegCount
        For lngC = 1 To gobjPNR.AirSegCount
            With gobjFareQuotes(1).FQ(1).FareSeg(lngY)
                If gobjPNR.AirSeg(lngC).ArriveAirport = .ArrCityCode And gobjPNR.AirSeg(lngC).DepartAirport = .DepCityCode And gobjPNR.AirSeg(lngC).Vendor = .Vendor And gobjPNR.AirSeg(lngC).FlightNumber = .FlightNum And gobjPNR.AirSeg(lngC).DepartDateTime = .DepDate And gobjPNR.AirSeg(lngC).Class = .Cos Then
                        If mbytNumSegs = 0 Then
                           mstrFirstSeg = .DepCityCode
                        End If
                        gobjPNR.AirSeg(lngC).SelectedForPricing = True
                        txtFlightInfo(mbytNumSegs).Text = Mid(gobjPNR.AirSeg(lngC).TextAirSeg, 5, 23)
                        txtSegNum(mbytNumSegs) = gobjPNR.AirSeg(lngC).segnumber
                        strSegmentNo = strSegmentNo & Format(gobjPNR.AirSeg(lngC).segnumber, "00")
                        mbytNumSegs = mbytNumSegs + 1
                End If
            End With
        Next lngC
    Next lngY
Else
    For lngC = 1 To gobjPNR.AirSegCount
        With gobjPNR.AirSeg(lngC)
            If .SelectedForPricing = True Then
                If mbytNumSegs = 0 Then
                    mstrFirstSeg = .DepartCityCode
                End If
                txtFlightInfo(mbytNumSegs).Text = Mid(.TextAirSeg, 5, 23)
                txtSegNum(mbytNumSegs) = .segnumber
                strSegmentNo = strSegmentNo & Format(gobjPNR.AirSeg(lngC).segnumber, "00")
                mbytNumSegs = mbytNumSegs + 1
            End If
        End With
    Next lngC
End If

lngY = 1

strPFBC() = Split(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).PriceFBC, "/")


For lngC = 0 To gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareSegCount - 1
    With gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareSeg(lngC + 1)
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).FareSeg(" & lngC + 1 & ") VALUES:"
        If .Stopover = False Then chkConnection(lngC).value = vbChecked
        chkConnection(lngC).Enabled = False
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "  .Stopover = " & .Stopover
        txtFBC(lngC) = .FBC
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "  .FBC = " & .FBC
        txtTktDesig(lngC) = .TD
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "  .TD = " & .TD
        If .NVA >= (Date - 30) Then txtNVA(lngC) = UCase(Format(.NVA, "ddmmmyy"))
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "  .NVA = " & UCase(Format(.NVA, "ddmmmyy"))
        If .NVB >= (Date - 30) Then txtNVB(lngC) = UCase(Format(.NVB, "ddmmmyy"))
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "  .NVB = " & UCase(Format(.NVB, "ddmmmyy"))
        txtBag(lngC) = .BagInfo
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "  .BagInfo = " & .BagInfo
        'Modified on 05/10/05: capture PFBC in Farequote request if no APF Price FBC
        'If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).PriceFBC = "" Then
        ''txtPriceFBC (lngC)
        '    txtPriceFBC(lngC) = .OverridePFBC
        'Else
        '    txtPriceFBC(lngC) = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).PriceFBC
        'End If
        
  
        
        If InStr(1, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).PriceFBC, "/") = 0 Then
           If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).PriceFBC = "" Then
              txtPriceFBC(lngC) = .OverridePFBC
           Else
              txtPriceFBC(lngC) = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).PriceFBC
           End If
        Else
           If UBound(strPFBC) >= lngC Then
              txtPriceFBC(lngC) = strPFBC(lngC)
           End If
        End If
              
        
        
        
        
        ''11012005 Copy over Ticket FBC to Price FBC if Price FBC = "" requested by Sok Leng
        'If txtPriceFBC(lngC).Text = "" Then
        '   txtPriceFBC(lngC).Text = txtFBC(lngC).Text
        'End If
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).FareComponent(lngY).PriceFBC = " & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).PriceFBC
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "  ."
        If .ArrCityCode = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).Destinantion Then
            txtValue(lngC) = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngY).Amount
            lngY = lngY + 1
        End If
        If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = True Then
           txtFlightInfo(lngC).Enabled = False
           txtFBC(lngC).Enabled = False
           txtTktDesig(lngC).Enabled = False
           txtPriceFBC(lngC).Enabled = False
           txtNVB(lngC).Enabled = False
           txtNVA(lngC).Enabled = False
           txtBag(lngC).Enabled = False
        Else
           txtFlightInfo(lngC).Enabled = True
           txtFBC(lngC).Enabled = True
           txtTktDesig(lngC).Enabled = True
           txtPriceFBC(lngC).Enabled = True
           txtNVB(lngC).Enabled = True
           txtNVA(lngC).Enabled = True
           txtBag(lngC).Enabled = True
        End If
    End With
Next
'''??
For lngC = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareSegCount To 11
    chkConnection(lngC).Enabled = False
    txtFlightInfo(lngC).Enabled = False
    txtFBC(lngC).Enabled = False
    txtTktDesig(lngC).Enabled = False
    txtPriceFBC(lngC).Enabled = False
    txtNVB(lngC).Enabled = False
    txtNVA(lngC).Enabled = False
    txtBag(lngC).Enabled = False
Next lngC


    With gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(1)
        If .CommissionOnTicket <> 0 Then
            txtFareInfo(3).Text = .CommissionOnTicket
            
            optCommissionType(0).value = True
        End If
        
        
        
        'added on 01/10/04: PT Surcharge
        If .PaperTktSurcharge > 0 Then
            txtTktMod(10).Text = .PaperTktSurcharge
        End If
        'Added on 7/02/05: Generate TC for BI,CA,MH
        If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ITNum = "" And gstrAgcyCountryCode = "SG" And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PFFareType = "APF" And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PrivateFare = True Then
            txtTktMod(0).Text = getTCMapper(cmbValCarrier.Text, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).BaseAmount)
        Else
            txtTktMod(0).Text = RemoveChar(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ITNum, "B", True, Mid(txtTktMod(0).Tag, 3)) ' TourCode
        End If
        
        txtTktMod(9).Text = RemoveChar(.ValueCode, "B", True, Mid(txtTktMod(9).Tag, 3))   'ValueCode
        txtTktMod(3).Text = RemoveChar(.FOPCode, "B", True, Mid(txtTktMod(3).Tag, 3))
        'txtTktMod(4).Text = RemoveChar(.Endorsement(1), "B", True, Mid(txtTktMod(4).Tag, 3)) 'Endors1
        txtTktMod(5).Text = RemoveChar(.Endorsement(2), "B", True, Mid(txtTktMod(5).Tag, 3))     'Endors2
        txtTktMod(6).Text = RemoveChar(.Endorsement(3), "B", True, Mid(txtTktMod(6).Tag, 3))     'Endors3
        txtTktMod(7).Text = RemoveChar(.Endorsement(4), "B", True, Mid(txtTktMod(7).Tag, 3))  'Endors4
        txtTktMod(8).Text = RemoveChar(.Endorsement(5), "B", True, Mid(txtTktMod(8).Tag, 3))     'Endors5
        
        Select Case .FareOnTicket
            Case "APF"
                cmbFareOnTkt.listindex = 1
            Case "ITN"
                cmbFareOnTkt.listindex = 2
            Case "BTN"
                cmbFareOnTkt.listindex = 3
            Case "ITC"
                cmbFareOnTkt.listindex = 4
            Case "BTC"
                cmbFareOnTkt.listindex = 5
            Case Else
                .FareOnTicket = "PUB"
                cmbFareOnTkt.listindex = 0
        End Select
        
        'modified on 09/11/04: EB1 will be auto-populated in FF when FareOnTicket <> PUB
        If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PrivateFare And cmbFareOnTkt.listindex < 1 Then
            If InStr(.Endorsement(1), "^") > 0 Then
                txtTktMod(4).Text = RemoveChar(Mid(.Endorsement(1), 1, InStr(.Endorsement(1), "^") - 1), "B", True, Mid(txtTktMod(4).Tag, 3)) 'Endors1
            Else
                txtTktMod(4).Text = RemoveChar(.Endorsement(1), "B", True, Mid(txtTktMod(4).Tag, 3)) 'Endors1
            End If
            cmdFCEndorse.Enabled = False
        Else
            txtTktMod(4).Text = ""
            cmdFCEndorse.Enabled = True
        End If
        
    End With

 
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount > 0 Then
   txtFareInfo(4).Text = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount
   'Preethi - V1.2.1 20101011 - CR21 - Nett Fare Mark Up
   'If gstrAgcyCountryCode = "HK" Then
      'Added by JiYong to get the actual net fare
      If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount > gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ActualNetAmount And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ActualNetAmount > 0 Then
         txtFareInfo(4).Text = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ActualNetAmount
      End If
   'End If
End If
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).SellAmount > 0 Then txtFareInfo(5).Text = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).SellAmount
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).CommissionPt > 0 Then txtFareInfo(3).Text = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).CommissionPt
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PTkt = True Then chkPaperTkt.value = vbChecked
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = True Then
   cmbFareOnTkt.listindex = 1
   cmbFareOnTkt.Enabled = False
End If

txtMI(1).Text = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).LowFare
cmbFOP(0).listindex = 0


With gobjPNR
    If .FOPType = "CC" Then
    
       ' fill in cc details on both sets of controls
       'modified on 27/1/05: checking for invalid cc vendor
        blnChkVendor = True
        blnChkDate = True
        If .FOP_CCCode <> "" Then
            If validateCCVendor(cmbFOP(1)) = True Then
                cmbFOP(1).Text = .FOP_CCCode
            Else
                blnChkVendor = False
            End If
        End If
        txtTktMod(1).Text = .FOP_CCNum
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog ".FOP_CCExpireDate = " & CStr(.FOP_CCExpireDate) & " /" & CDec(.FOP_CCExpireDate)
        If validateCCDate(.FOP_CCExpireDate) Then
            dtpCCExpDate(0).value = .FOP_CCExpireDate
            dtpCCExpDate(1).value = .FOP_CCExpireDate
        Else
           blnChkDate = False
        End If
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog ".FOP_CCExpireDate < Now = " & CStr(.FOP_CCExpireDate < Now)
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "dtpCCExpDate(0).Value = " & dtpCCExpDate(0).value & " /" & CDec(dtpCCExpDate(0).value)
                
        If .FOP_CCCode <> "" Then
            If validateCCVendor(cmbFOP(3)) = True Then
               cmbFOP(3).Text = .FOP_CCCode
            Else
                blnChkVendor = False
            End If
        End If
        txtTktMod(2).Text = .FOP_CCNum
        
        If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "dtpCCExpDate(1).Value = " & dtpCCExpDate(1).value
        
        If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PrivateFare = True Then
            cmbFOP(2).listindex = 1
        End If
        
        If blnChkVendor = False Or blnChkDate = False Then
           promptCCError blnChkVendor, blnChkDate
        End If
    End If
    
    'Added on 04/10/04: Retrieve FF10,FF11
    '14/01/05: FF10,11 handled by Client MI screen
    'intFFFound = 0
    'For lngC = 1 To .AcctRemarkCount
    '    If UCase(Mid(.AcctRemark(lngC).RemarkText, 1, IIf(InStr(.AcctRemark(lngC).RemarkText, "/") - 1 > 0, InStr(.AcctRemark(lngC).RemarkText, "/") - 1, 0))) = "FF10" Then
    '        If Mid(.AcctRemark(lngC).RemarkText, InStr(.AcctRemark(lngC).RemarkText, "/") + 1, 1) = "*" Then
    '            txtMI(4).Text = UCase(Mid(.AcctRemark(lngC).RemarkText, InStr(InStr(.AcctRemark(lngC).RemarkText, "/") + 1, .AcctRemark(lngC).RemarkText, "/") + 1))
    '        Else
    '            txtMI(4).Text = UCase(Mid(.AcctRemark(lngC).RemarkText, InStr(.AcctRemark(lngC).RemarkText, "/") + 1))
    '        End If
    '        intFFFound = intFFFound + 1
    '    ElseIf UCase(Mid(.AcctRemark(lngC).RemarkText, 1, IIf(InStr(.AcctRemark(lngC).RemarkText, "/") - 1 > 0, InStr(.AcctRemark(lngC).RemarkText, "/") - 1, 0))) = "FF11" Then
    '        If Mid(.AcctRemark(lngC).RemarkText, InStr(.AcctRemark(lngC).RemarkText, "/") + 1, 1) = "*" Then
    '            txtMI(5).Text = UCase(Mid(.AcctRemark(lngC).RemarkText, InStr(InStr(.AcctRemark(lngC).RemarkText, "/") + 1, .AcctRemark(lngC).RemarkText, "/") + 1))
    '        Else
    '            txtMI(5).Text = UCase(Mid(.AcctRemark(lngC).RemarkText, InStr(.AcctRemark(lngC).RemarkText, "/") + 1))
    '        End If
    '        intFFFound = intFFFound + 1
    '    End If
    '    If intFFFound > 1 Then
    '        Exit For
    '    End If
    'Next lngC
    '
End With

'Added on 15/10/04: TMP card: uncheck UATP + FOP on PAX set to INV


UATPControl (1)
'Modified on 1/2/05: add on client specifie EC
'strSQL = "SELECT * FROM tblEXceptionCodes where ExceptionCodeGroup='C' order by cint(exceptioncode) "

'Set rsECodes = gdbTPro.OpenRecordset(strSQL)

'With rsECodes
'    .MoveFirst
'   Do While Not .EOF
'      Set item = lvwECodes.ListItems.Add(, , !exceptioncode)
'        item.SubItems(1) = !Description
      
'      .MoveNext
'    Loop
'  End With

'rsECodes.Close

'Added on 15/03/05: for HK, retrieve NRCC from Farequote
bolFound = False
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NRCC Then chkNRCC.value = 1
For lngC = 0 To cmbValCarrier.ListCount - 1
    If cmbValCarrier.List(lngC) = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PlatCarrier Then
       cmbValCarrier.listindex = lngC
       bolFound = True
    End If
Next
If bolFound = False Then
   cmbValCarrier.AddItem Trim(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PlatCarrier)
   cmbValCarrier.listindex = cmbValCarrier.ListCount - 1
End If

'300908 Detect MI by CN
If gobjPNR.CompInfo.MI = False Then
    fraMI.Enabled = False
    fraEC.Enabled = False
    lblLabels(40).Enabled = False
    lblLabels(41).Enabled = False
    lblLabels(39).Enabled = False
    lblLabels(36).Enabled = False
    lblLabels(32).Enabled = False
    lblLabels(27).Enabled = False
    lblLabels(30).Enabled = False
    lblLabels(28).Enabled = False
    For lngC = 0 To txtMI.Count - 1
        If lngC <> 2 Then
        txtMI(lngC).Enabled = False
        End If
    Next
    txtRS.Enabled = False
    txtMS.Enabled = False
    lvwRealECodes.Enabled = False
    lvwMissECodes.Enabled = False
    cboClassServ.Enabled = False
    cboBookingAction.Enabled = False
    
End If

'strSql = "SELECT distinct(CAST(tblClientEC.EC AS integer)),description,exceptioncodegroup,REMARKS FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC and tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='AS' OR tblExceptionCodes.ExceptionCodeGroup='AC') ORDER BY CAST(tblClientEC.EC AS integer)"

'CS Change EC
'strSql = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='AS' OR tblExceptionCodes.ExceptionCodeGroup='AC') ORDER BY CAST(tblClientEC.EC AS integer)"
'CS Change EC
'strSQL = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND tblExceptionCodes.ProdType='" & "AIR" & "' AND tblExceptionCodes.ECInd='S' ORDER BY tblClientEC.EC"
strSQL = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.ClientID & " AND tblExceptionCodes.ProdType='AIR' AND tblClientEC.ProdType='AIR' ORDER BY tblClientEC.EC"
Set rsECodes = gdbConn.Execute(strSQL)
If Not rsECodes.EOF Then

     rsECodes.MoveFirst
     Do While Not rsECodes.EOF
        If rsECodes!ECCat = "R" Then
            Set item = lvwRealECodes.ListItems.Add(, , rsECodes!EC)
                  If rsECodes!Remarks = "" Then
                   item.SubItems(1) = rsECodes!Description
                  Else
                   item.SubItems(1) = rsECodes!Remarks
                  End If
            rsECodes.MoveNext
         Else
            Set item = lvwMissECodes.ListItems.Add(, , rsECodes!EC)
                  If rsECodes!Remarks = "" Then
                   item.SubItems(1) = rsECodes!Description
                  Else
                   item.SubItems(1) = rsECodes!Remarks
                  End If
            rsECodes.MoveNext
         End If
      Loop
    rsECodes.Close

Else
   
        rsECodes.Close
        strSQL = "SELECT * FROM tblExceptionCodes where ProdType='" & "AIR" & "' and ECInd='C' order by ExceptionCode"
        Set rsECodes = gdbConn.Execute(strSQL)
           If Not rsECodes.EOF Then rsECodes.MoveFirst
           
           Do While Not rsECodes.EOF
           If rsECodes!ECCat = "R" Then
              Set item = lvwRealECodes.ListItems.Add(, , rsECodes!exceptioncode)
              item.SubItems(1) = rsECodes!Description
              rsECodes.MoveNext
            Else
              Set item = lvwMissECodes.ListItems.Add(, , rsECodes!exceptioncode)
              item.SubItems(1) = rsECodes!Description
              rsECodes.MoveNext
            End If
            Loop
          rsECodes.Close
     
End If

Set rsECodes = Nothing

''Set rsECodes = gdbTPro.OpenRecordset(strSQL)
'Set rsECodes = gdbConn.Execute(strSql)
'If Not rsECodes.EOF Then
'
'     rsECodes.MoveFirst
'     Do While Not rsECodes.EOF
'        Set Item = lvwECodes.ListItems.Add(, , rsECodes!EC)
'              If rsECodes!Remarks = "" Then
'               Item.SubItems(1) = rsECodes!Description
'              Else
'               Item.SubItems(1) = rsECodes!Remarks
'              End If
'        rsECodes.MoveNext
'      Loop
'    rsECodes.Close'
'
'Else
'
'        rsECodes.Close
'        strSql = "SELECT * FROM tblExceptionCodes where ExceptionCodeGroup='AC' order by CAST(ExceptionCode AS integer) "
'        'Set rsECodes = gdbTPro.OpenRecordset(strSQL)
'        Set rsECodes = gdbConn.Execute(strSql)
'           If Not rsECodes.EOF Then rsECodes.MoveFirst
'
'           Do While Not rsECodes.EOF
'              Set Item = lvwECodes.ListItems.Add(, , rsECodes!exceptioncode)
'              Item.SubItems(1) = rsECodes!Description
'              rsECodes.MoveNext
'            Loop
'          rsECodes.Close
'
'End If
'
'Set rsECodes = Nothing

'Added on 14 Dec fro Cat 35
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = True Then
   'able to amend commission if the commission is not being filed
   If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).CommType = True Then
      txtFareInfo(3).Enabled = False
   Else
      txtFareInfo(3).Enabled = True
   End If
   txtFareInfo(4).Enabled = False
   txtTktMod(0).Enabled = False
   txtTktMod(9).Enabled = False
   txtTktMod(4).Text = ""
   txtTktMod(5).Text = ""
   txtTktMod(6).Text = ""
   txtTktMod(6).Enabled = False
   txtTktMod(7).Text = ""
   txtTktMod(7).Enabled = False
   txtTktMod(8).Text = ""
   txtTktMod(8).Enabled = False
Else
   txtFareInfo(3).Enabled = True
   txtFareInfo(4).Enabled = True
   txtTktMod(0).Enabled = True
   txtTktMod(9).Enabled = True
   txtTktMod(6).Enabled = True
   txtTktMod(7).Enabled = True
   txtTktMod(8).Enabled = True
End If

        
End Sub

Private Sub pFillInFFData()
Dim strFareCalc As String

'gobjPNR.LoadFiledFare (CStr(mbytFFNum))

'Call pGetFFInfo

With txtFareInfo(6)
    .Enabled = True
     strFareCalc = gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).FareConstructText
    .Text = Mid(strFareCalc, 1, IIf((InStr(1, strFareCalc, "END ROE") - 1) > 0, InStr(1, strFareCalc, "END ROE") - 1, Len(strFareCalc)))
    '.Text = gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).FareConstructText
    'If txtTktMod(6).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(6).Text)
    'If txtTktMod(7).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(7).Text)
    'If txtTktMod(8).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(8).Text)
    
End With

With txtFareInfo(0)
    If gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).BaseCurrency = gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).TotalCurrency Then
        .Text = CStr(gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).BaseAmount)
    Else
        .Text = CStr(gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).EquivAmount)
    End If
    .Enabled = True
    msngFareDiff = fConvertZero(.Text) - fConvertZero(txtFareInfo(5).Text)
    .Locked = True
End With

With txtFareInfo(1)
    .Text = CStr(gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).TaxTotal)
    .Enabled = True
    .Locked = True
End With

With txtFareInfo(2)
    .Text = CStr(gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).TotAmount)
    .Enabled = True
    .Locked = True
End With

sstTabs.Tab = 1

End Sub

Private Function UpdateTktData() As Boolean
Dim lngTS As Long 'ticket segments
Dim lngC As Long
Dim strCmd As String
Dim strResp As String
Dim strMsg As String
Dim strLogLine As String
Dim intPxNum As Integer
Dim bolFound As Boolean
Dim strTemp() As String
Dim intI As Integer
UpdateTktData = True

'Added on 01/10/04: PT Surcharge
'If fConvertZero(txtTktMod(10).Text) > 0 Then
'    mstrFBUFields = mstrFBUFields & "/PTS"
'End If

intPxNum = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PxNum
'Change on 22 Dec to Entry *FB command with Px Number
If mstrFBUFields <> "" Or lstTax.ListCount > 0 Then
    If InStr(gobjHost.terminalEntry("*FB" & mbytFFNum & "P" & intPxNum), "*FB" & mbytFFNum & "P" & intPxNum) > 0 Then
End If

If mstrFBUFields <> "" And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
        mbolFBUMode = True
        strCmd = "FBU"
        For lngTS = 0 To mbytNumSegs - 1
        'remove on 12/4/05: Already specified in Filefare
            'If InStr(mstrFBUFields, "/CNX") > 0 Then
            '    strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "X" & CStr(lngTS + 1) & "/" & chkConnection(lngTS).Caption
            '    If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "chkConnection(" & lngTS & ").Caption = " & chkConnection(lngTS).Caption
            'End If
            
            If InStr(mstrFBUFields, "/FBC") > 0 Or InStr(mstrFBUFields, "/TKD") > 0 Then
                'added on 13/12:
                'if PFBC is overriden, the FareCal EB (for the overriden fare basis) should be removed
                If InStr(mstrFBUFields, "/FBC") > 0 Then
                '    strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "EB/"
                    strResp = gobjHost.terminalEntry("FBUEB/", True)
                End If

                strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "FB" & CStr(lngTS + 1) & "/" _
                    & txtFBC(lngTS).Text & IIf(txtTktDesig(lngTS).Text <> "", "/" & txtTktDesig(lngTS).Text, "")
                If gobjLog.LogOpen = True Then
                    gobjLog.LineTextToLog "txtFBC(" & lngTS & ").Text = " & txtFBC(lngTS).Text
                    gobjLog.LineTextToLog "txtTktDesig(" & lngTS & ").Text = " & txtTktDesig(lngTS).Text
                End If
            End If
            
            'If InStr(mstrFBUFields, "/NVB") > 0 Then
                'strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "NVB" & CStr(lngTS + 1) & "/" & Left(txtNVB(lngTS).Text, 5)
                'If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "txtNVB(" & lngTS & ").Text = " & txtNVB(lngTS).Text
                strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "NVB" & CStr(lngTS + 1) & "/" & txtNVB(lngTS).Text
                If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "txtNVB(" & lngTS & ").Text = " & txtNVB(lngTS).Text
            
            'End If

            'If InStr(mstrFBUFields, "/NVA") > 0 Then
                'strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "NVA" & CStr(lngTS + 1) & "/" & Left(txtNVA(lngTS).Text, 5)
                'If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "txtNVA(" & lngTS & ").Text = " & txtNVA(lngTS).Text
                strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "NVA" & CStr(lngTS + 1) & "/" & txtNVA(lngTS).Text
                If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "txtNVA(" & lngTS & ").Text = " & txtNVA(lngTS).Text
            
            'End If
            
            If InStr(mstrFBUFields, "/BAG") > 0 Then
                strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "BG" & CStr(lngTS + 1) & "/" & txtBag(lngTS).Text
                If gobjLog.LogOpen = True Then gobjLog.LineTextToLog "txtBag(" & lngTS & ").Text = " & txtBag(lngTS).Text
            End If
            
            If Len(strCmd) > 180 Then
                strResp = gobjHost.terminalEntry(strCmd, True)
                gobjLog.EventToLog "SENT TO HOST: " & strCmd
                If InStr(strResp, "DATA ACCEPTED") = 0 Then
                    strMsg = "UNABLE TO UPDATE TICKET DATA (FBU)" & Chr(13) & Chr(13) _
                        & "RESPONSE FROM GALILEO WAS:" & Chr(13) & strResp
                    'MsgBox strMsg
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                
                    'pSendToFP "*FB"
                    pDisplayToFP "*FB"
                    UpdateTktData = False
                    Exit Function
                End If
                strCmd = ""
                If lngTS <> mbytNumSegs - 1 Then strCmd = "FBU"
            End If
            
        Next

        'Added on 01/10/04: PT Surcharge
        'If fConvertZero(txtTktMod(10).Text) > 0 Then
        '    strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "FARE/" & gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).TotalCurrency & Format(txtTktMod(10).Text, gstrAgcyCurrFormat)
        'End If
        '
        If InStr(mstrFBUFields, "/FCN") > 0 Then
            strCmd = strCmd & IIf(strCmd <> "FBU", "+", "") & "FC/" & txtFareInfo(6).Text
        End If
        
        
        If strCmd <> "" Then
        strResp = gobjHost.terminalEntry(strCmd, True)
            If InStr(strResp, "DATA ACCEPTED") = 0 Then
                strMsg = "Command: " & strCmd & Chr(13) & Chr(13)
                
                strMsg = strMsg & "UNABLE TO UPDATE TICKET DATA (FBU)" & Chr(13) & Chr(13) _
                            & "RESPONSE FROM GALILEO WAS:" & Chr(13) & strResp
                'MsgBox strMsg
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                'pSendToFP "*FB"
                pDisplayToFP "*FB"
                UpdateTktData = False
                Exit Function
            End If
        End If
        
        
        
        
    End If
End If
        
        'Added on 140607: YY taxes (Amend tax amount if same tax code)
        Dim intTaxcount As Integer
        
        With gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1)
            If lstTax.ListCount > 0 Then
                intTaxcount = .TaxCount
                For intI = 0 To lstTax.ListCount - 1
                    bolFound = False
                    strTemp = Split(lstTax.List(intI), " ")
                    If UBound(strTemp) > 0 Then
                       For lngC = 1 To .TaxCount
                           If UCase(.Tax(lngC).TaxCode) = UCase(strTemp(1)) Then
                              bolFound = True
                              Exit For
                           End If
                       Next
                    End If
                    If bolFound = True Then
                       strCmd = IIf(strCmd = "", "FBUTAX", strCmd & "+TAX") & lngC & "/" & strTemp(0) & strTemp(1)
                    Else
                        intTaxcount = intTaxcount + 1
                        strCmd = IIf(strCmd = "", "FBUTAX", strCmd & "+TAX") & intTaxcount & "/" & strTemp(0) & strTemp(1)
                    End If
                Next
                strCmd = strCmd & "+TTL/"
            End If
        End With
        
        If strCmd <> "" Then
        strResp = gobjHost.terminalEntry(strCmd, True)
            If InStr(strResp, "DATA ACCEPTED") = 0 Then
                strMsg = "Command: " & strCmd & Chr(13) & Chr(13)
                
                strMsg = strMsg & "UNABLE TO UPDATE TICKET DATA (FBU)" & Chr(13) & Chr(13) _
                            & "RESPONSE FROM GALILEO WAS:" & Chr(13) & strResp
                'MsgBox strMsg
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                pDisplayToFP "*FB"
                UpdateTktData = False
                Exit Function
            End If
        End If
        
        If strCmd <> "" Then
            strResp = gobjHost.terminalEntry("FBF", True)
    
            If InStr(strResp, "MANUAL FARE FILED") = 0 Then
                strMsg = "UNABLE TO UPDATE TICKET DATA (FBF)" & Chr(13) & Chr(13) _
                            & "RESPONSE FROM GALILEO WAS:" & Chr(13) & strResp
                'MsgBox strMsg
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                pDisplayToFP "*FB"
                Exit Function
            End If
            mstrFBUFields = ""
        Else
           gobjHost.terminalEntry "FBF"
        End If
End Function


Private Sub writeDatatoGDS()
Dim strTemp As String
Dim strEntry As String
Dim dtmNewDate As Date
Dim sngTotCharge As Single
Dim sngTmp As Single
Dim rs As ADODB.Recordset
Dim strTrxnFeeCode As String
Dim booIsTMPCard As Boolean
Dim sngTotChargeAC As Single
Dim strSecLine As String
'Preethi - V1.2.6 20110907 - CR 90 - Change OBT Tool Code in FF35
Dim strBookingTool As String
'strEntry = ""
If blnLoaded = False Then
    mintDisplayNo = mintDisplayNo + 1
End If
'Add a flag in NP line to indicate this is a Cat35 (for AQUA use)
addNP mintDisplayNo, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35
'dtmNewDate = DateAdd("d", 90, Date)
'ZhiSam - V1.2.23 20130911 - CR-289 - E-Invoice: Due and Paid Lines Date
If bfunctCheckRTLine = True Then
    dtmNewDate = dtfunctRTDate
Else
    dtmNewDate = DateAdd("d", 90, Date)
End If

'FMR
If gbolFMR = False Then
   sngTotCharge = fConvertZero(txtFareInfo(5).Text) + fConvertZero(txtFareInfo(1).Text)
   If fConvertZero(txtAComm) > 0 Then sngTotChargeAC = fConvertZero(txtFareInfo(4).Text) + fConvertZero(txtFareInfo(1).Text)

   
Else
   sngTotCharge = gdblAmtToCom
   sngTotChargeAC = gdblAmtToCom
End If

Set rs = gdbConn.Execute("Select sortkey from tblProductCodes where productcode = '35'")
If rs.EOF Then
    strTrxnFeeCode = ""
Else
    strTrxnFeeCode = rs!SortKey & ""
End If
rs.Close
Set rs = Nothing

If Me.cmbFOP(2).Text = "CC" Then
    If gbolFMR = False Then
        If chkUATP.value = 1 Or cmbFOP(0).Text = "CC" Then
            cmbFOP(3).Text = cmbFOP(1).Text
            txtTktMod(2).Text = txtTktMod(1).Text
            dtpCCExpDate(1).value = dtpCCExpDate(0).value
        End If
    End If
    
        Select Case cmbFOP(3).Text
            Case "AX"
                strTemp = "CX2"
            Case "DC"
                strTemp = "CX3"
            Case "VI", "CA"
                strTemp = "CX4"
            Case "TP"
                strTemp = "CX5"
        End Select
    Else
        strTemp = "CC"
End If

    'Check if it is TMP Card
    'If (Left(UCase(cmbFOP(3).Text), 2) = "DC" And _
        Left(UCase(txtTktMod(2).Text), 7) = "3644033") Then
     'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    If IsTMPCard(Left(UCase(cmbFOP(3).Text), 2), UCase(txtTktMod(2).Text)) Then
        booIsTMPCard = True
    Else
        booIsTMPCard = False
    End If


'mintDisplayNo = mintDisplayNo + 1

With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
    'FMR
    If gbolFMR = False Then
       'If fConvertZero(txtAComm) = 0 Then
            strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format(txtFareInfo(5).Text, gstrAgcyCurrFormat)
       'Else
       '     strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format(txtFareInfo(4).Text, gstrAgcyCurrFormat)
       'End If
    Else
       'Modified on 240806 - Get the rebate from FMR, no need to compute
        strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format(gdblAmtToCom - gdblTaxToCom + gdblRebate, gstrAgcyCurrFormat)
       'If (.DiscountAmt > .TransactionFee) And .TransactionFee > 0 Then
       ' strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format(gdblAmtToCom - gdblTaxToCom - .TransactionFee + .DiscountAmt, gstrAgcyCurrFormat)
       'Else
       ' strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format(gdblAmtToCom - gdblTaxToCom + .DiscountAmt, gstrAgcyCurrFormat)
       'End If
    End If

 If cmbFOP(2).Text = "CC" And Not booIsTMPCard Then
 'modified on 7/4/2005: change FOP include TF if disc>tf
        If gbolFMR = True Then
        '    If (.DiscountAmt > .TransactionFee) And .TransactionFee > 0 Then
        '       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/" & IIf(chkUATP.Value = 1, "CC", strTemp) & "/" & cmbFOP(3).Text & txtTktMod(2).Text _
        '           & "/" & Format(sngTotCharge, gstrAgcyCurrFormat)
       
        '    Else
               strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/" & IIf(chkUATP.value = 1, "CC", strTemp) & "/" & cmbFOP(3).Text & txtTktMod(2).Text _
                   & "/" & Format(sngTotCharge, gstrAgcyCurrFormat)
        '    End If
        
        Else
            'If chkTransFee.value = 1 Then
            '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/" & IIf(chkUATP.value = 1, "CC", strTemp) & "/" & cmbFOP(3).Text & txtTktMod(2).Text _
             '   & "/" & Format(sngTotCharge + .TransactionFee - .DiscountAmt, gstrAgcyCurrFormat)
            'Else
            'If fConvertZero(txtAComm) > 0 Then
            '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/" & IIf(chkUATP.value = 1, "CC", strTemp) & "/" & cmbFOP(3).Text & txtTktMod(2).Text _
            '    & "/" & Format(sngTotChargeAC - .DiscountAmt, gstrAgcyCurrFormat)
            'Else
                strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/" & IIf(chkUATP.value = 1, "CC", strTemp) & "/" & cmbFOP(3).Text & txtTktMod(2).Text _
                & "/" & Format(sngTotCharge - .DiscountAmt, gstrAgcyCurrFormat)
            'End If
            'End If
        End If
Else
    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/CASH"
End If


    'Added on 21/07/04 for FF8 and FF26
    If cboClassServ.Text <> "" Then
       If gobjPNR.CompInfo.MI = True Then strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF8/*" & mintDisplayNo & "/" & Trim(Left(cboClassServ.Text, 2))
    End If
    'CS - Remove FF26 (Trip Type)
    'If cboTripType.Text <> "" Then
    '    Select Case UCase(cboTripType.Text)
    '        Case "ROUND"
    '            strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF26/*" & mintDisplayNo & "/" & "R"
    '        Case "ONE WAY"
    '            strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF26/*" & mintDisplayNo & "/" & "O"
    '    End Select
    'End If
    'CS - Add International or Domestic
    'If cboTrip.Text <> "" Then
    '   strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF41/*" & mintDisplayNo & "/" & Left(cboTrip.Text, 1)
    'End If
   
    'CS Add Booking Action
    If cboBookingAction.Text <> "" And gobjPNR.CompInfo.MI = True Then
       Select Case cboBookingAction.Text
          Case "AB - Agent Booked"
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF34/*" & mintDisplayNo & "/" & "AB"
          Case "EB - Self Booked"
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF34/*" & mintDisplayNo & "/" & "EB"
          Case "AA - Air Modified"
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF34/*" & mintDisplayNo & "/" & "AA"
          Case "AM - Multiple Modification"
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF34/*" & mintDisplayNo & "/" & "AM"

       End Select
    End If
    
    'CS Add Booking Tool
    If mstrBookingTool <> "" And gobjPNR.CompInfo.MI = True Then
       'Preethi - V1.2.6 20110907 - CR 90 - Change OBT Tool Code in FF35
       strBookingTool = getFF35OBT(Mid(cboBookingAction.Text, 1, 2), mstrBookingTool)
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF35/*" & mintDisplayNo & "/" & strBookingTool
    Else
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF35/*" & mintDisplayNo & "/" & "GAL"
    End If
    
    'CS Add Booking Method
    'modified on 190106: Set SBT -S, GDS-G , no need user to select
    If mstrBookingTool <> "" And gobjPNR.CompInfo.MI = True Then
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "S"
    Else
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "G"
    End If
    'If cboBookingMethod.Text <> "" Then
       'Select Case UCase(cboBookingMethod.Text)
       '   Case "GDS"
       '      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "G"
       '   Case "MANUAL"
       '      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "M"
       '   Case "SELF BOOKING"
       '      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "S"
       'End Select
    'End If
    
    
    
    'CS - Add FF31 FF32
    
    If .TransactionFee > 0 Then sngTmp = Format(CStr(.TransactionFee), gstrAgcyCurrFormat)
    
    If gobjPNR.CompInfo.MI = True Then
    If .TransactionFee > 0 Then
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF31/*" & mintDisplayNo & "/Y"
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF32/*" & mintDisplayNo & "/" & sngTmp
    Else
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF31/*" & mintDisplayNo & "/N"
    End If
    End If

    
    
    'Added FF38 for ticket type
    If gobjPNR.CompInfo.MI = True Then strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF38/*" & mintDisplayNo & "/" & Left(chkPaperTkt.Caption, 1)
    
    'Added on 02/08/04 for FF10,FF11,FF19
    '14/01/05: FF10, 11 handled by Client MI screen
    'If txtMI(4).Text <> "" Then
    '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF10/*" & mintDisplayNo & "/" & Trim(txtMI(4).Text)
    'End If
    'If txtMI(5).Text <> "" Then
    '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF11/*" & mintDisplayNo & "/" & Trim(txtMI(5).Text)
    'End If
    If txtMI(6).Text <> "" And gobjPNR.CompInfo.MI = True Then
        strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF81/*" & mintDisplayNo & "/" & UCase(Trim(txtMI(6).Text))
    End If
    If gobjPNR.CompInfo.MI = True Then strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF7/*" & mintDisplayNo & "/" & txtMI(3).Text
    'CS Add txtRS (ff30)
    If gobjPNR.CompInfo.MI = True Then strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF30/*" & mintDisplayNo & "/" & txtRS
    'CS Change EC txtMI(2) --> txtMS
    If gobjPNR.CompInfo.MI = True Then
        strEntry = strEntry & IIf(strEntry <> "", "+", "") _
            & "DI.FT-LF/*" & mintDisplayNo & "/" & txtMI(1).Text _
            & "+DI.FT-RF/*" & mintDisplayNo & "/" & txtMI(0).Text _
            & "+DI.FT-EC/*" & mintDisplayNo & "/" & txtMS.Text
    End If
    '19/01/05: Write client MI data into DI
    If frmClientMI.DIfreefields <> "" Then
        strEntry = strEntry & IIf(strEntry <> "", "+", "") & frmClientMI.DIfreefields
    End If
    
    'CS
    'If .TransactionFee > 0 Then sngTmp = Format(CStr(.TransactionFee), gstrAgcyCurrFormat)


Dim strTFFOPCharge As String
If chkTransFee.value = 1 Then
    strTFFOPCharge = "CC"
Else
    strTFFOPCharge = strTemp
End If
'IIf(Len(strTemp) > 0 And Not booIsTMPCard, _
'Modified on 6/4/2005: if FOP from Pax is INV, TF FOP will be FS
    If .TransactionFee > 0 Then
       ' strEntry = strEntry & IIf(strEntry <> "", "+", "") & _
       '      "DI.FT-MS/PC" & strTrxnFeeCode & "/" & VendorNum("35", "CWT") & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & _
             "DI.FT-MS/PC" & strTrxnFeeCode & "/" & VendorNum(strTrxnFeeCode, IIf(chkTransFee.value = False, "CWT", "MER")) & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum
      
               'Preethi - V1.2.4 20110527 - CR 59 - Remove Round Up Function on Transaction Fee
                'strSecLine = "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & Format(sngTmp, "#0.00") & "/SF" & Format(sngTmp, "#0.00") & "/C" & Format(sngTmp, "#0.00") & strSegmentNo
                'Preethi - V1.2.4 20110613 - CR 59 - Remove Round Up Function on Transaction Fee
                strSecLine = "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & strSegmentNo

                strEntry = strEntry & splitLongMSX(strSecLine)
                
             If cmbFOP(2).Text = "CC" Then
                'If (.DiscountAmt > .TransactionFee) And gbolFMR Then
                
                If (gdblRebate = .DiscountAmt - .TransactionFee) And gbolFMR Then
                       strEntry = strEntry & "+DI.FT-MSX/FS"
                Else
                      'Preethi - V1.2.4 20110527 - CR 59 - Remove Round Up Function on Transaction Fee
                      'strEntry = strEntry & "+DI.FT-MSX/F" & strTFFOPCharge & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "EXP" & Format(dtpCCExpDate(1).value, "MMYY") & "/D" & Format(CStr(.TransactionFee), "#0.00")
                        'Preethi - V1.2.4 20110613 - CR 59 - Remove Round Up Function on Transaction Fee
                        strEntry = strEntry & "+DI.FT-MSX/F" & strTFFOPCharge & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "EXP" & Format(dtpCCExpDate(1).value, "MMYY") & "/D" & Format(CStr(.TransactionFee), gstrAgcyCurrFormat)

                       'If chkTransFee.value = 0 Then
                       '   strSecLine = strEntry & IIf(strEntry <> "", "+", "") & "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp
                       'Else
                       '   strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & "/FS"
                       'End If
                End If
             Else
                    strEntry = strEntry & "+DI.FT-MSX/FS"
             End If
             
             
             'strEntry = strEntry & splitLongMSX(strSecLine)
            
             'If (gdblRebate = .DiscountAmt - .TransactionFee) And gbolFMR And cmbFOP(2).Text = "CC" Then
             
            '    strEntry = strEntry & "+DI.FT-MSX/F" & strTFFOPCharge & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "EXP" & Format(dtpCCExpDate(1).value, "MMYY") & "/D" & Format(CStr(.TransactionFee), gstrAgcyCurrFormat)
             
            ' End If
    'added on 08/12: include FF10,11 in MS line for BTA clients
    '14/01/05: FF10,11 handled by Client MI screen
    'If txtMI(4).Text <> "" And txtMI(5).Text <> "" Then
    '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/FF10-" & txtMI(4).Text & "/FF11-" & txtMI(5).Text
    'End If
    
    'added on 17/01/05: copy all file-fare related MI to MSX line
    strEntry = strEntry & splitLongMSX(getMSLineforMI())
    End If
    
        'modified 121007: request by hk, MN client may have hidden TF also
        'If fConvertZero(txtAComm) <> 0 Then
        '   strEntry = strEntry & IIf(strEntry <> "", "+", "") & _
        '              "DI.FT-MS/PC" & strTrxnFeeCode & "/" & VendorNum(strTrxnFeeCode, "MER") & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum & _
        '              IIf(Len(strTemp) > 0 And Not booIsTMPCard, _
        '              "+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & txtAComm & "/SF" & txtAComm & "/C" & txtAComm & "+DI.FT-MSX/F" & "CC" & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "EXP" & Format(dtpCCExpDate(1).value, "MMYY") & "/D" & Format(CStr(txtAComm), gstrAgcyCurrFormat), _
        '              "+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & txtAComm & "/SF" & txtAComm & "/C" & txtAComm & "/FS")
        '   strEntry = strEntry & splitLongMSX(getMSlineForMI())
        '   strEntry = strEntry & "+DI.FT-MSX/FF TRANSACTION FEE"
        'End If
    
'Removed on 26/07/04 - do not need this hard-coded entry
'"+DI./0+DI.FT-MS/VCWT/TK0000000000000/PC35"

'230108
Dim strFuelChargeCode As String

Set rs = gdbConn.Execute("Select sortkey from tblProductCodes where productcode = '41'")
If rs.EOF Then
    strFuelChargeCode = ""
Else
    strFuelChargeCode = rs!SortKey & ""
End If
rs.Close
Set rs = Nothing

    If .FuelSurcharge > 0 Then
           strEntry = strEntry & IIf(strEntry <> "", "+", "") & _
             "DI.FT-MS/PC" & strFuelChargeCode & "/" & VendorNum(strFuelChargeCode, "CWT") & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum
      
             If cmbFOP(2).Text = "CC" Then
           
                    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & .FuelSurcharge & "/SF" & .FuelSurcharge & "/C" & .FuelSurcharge & "+DI.FT-MSX/F" & strTemp & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "EXP" & Format(dtpCCExpDate(1).value, "MMYY") & "/D" & Format(CStr(.FuelSurcharge), gstrAgcyCurrFormat)
      
             Else
                    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & .FuelSurcharge & "/SF" & .FuelSurcharge & "/C" & .FuelSurcharge & "/FS"
             End If
             strEntry = strEntry & splitLongMSX(getMSLineforMI())
    End If
    


    End With

'Added by JiYong to add NF to DI line
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount > 0 Then
   If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).HiddenComm = 0 Then
      If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount > gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ActualNetAmount And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ActualNetAmount > 0 Then
         strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-NF/*" & mintDisplayNo & "/" & Format(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount, gstrAgcyCurrFormat)
         'Preethi - V1.2.1 20101015 - CR21 - Nett Fare Mark Up
         strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF90/*" & mintDisplayNo & "/CWTF"
      End If
   Else
      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-NF/*" & mintDisplayNo & "/" & Format(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount + gobjFareQuotes(cmbPx.listindex + 1).FQ(1).HiddenComm, gstrAgcyCurrFormat)
   End If
   
End If

'gobjHost.terminalEntry strEntry
SendDIEntry (strEntry)

'Added on 29/07/04 for DI line entries
If UCase(gstrAgcyCountryCode) = "HK" Then
    Call pDIEntryHKG(strTemp)
'Added on 29/07/04 for RD line entries
    Call pRDEntryHKG
'
End If


'230108: Add Filefare agent for productivity tracking for HK
Dim strFFAgtSignon As String
Dim OSLineNo As Integer
    OSLineNo = OSLineNum
    If gobjHost.AgentGACode <> "" Then
        
        strFFAgtSignon = IIf(OSLineNo > 0, "DI@" & OSLineNo & ".FT-OS/", "DI.FT-OS/") & gobjHost.AgentGACode
       
    Else
        strFFAgtSignon = IIf(OSLineNo > 0, "DI@" & OSLineNo & ".FT-OS/", "DI.FT-OS/") & gobjHost.AgentSine
    End If
    
    gobjHost.terminalEntry strFFAgtSignon
    

blnLoaded = False
End Sub
Private Function OSLineNum() As Integer

Dim lngC As Long

For lngC = 1 To gobjPNR.AcctRemarkCount
        With gobjPNR.AcctRemark(lngC)
            If Left(.RemarkText, 2) = "OS" Then
                    OSLineNum = .ItemNum
                    Exit Function
            End If
            
        End With

Next

End Function
Private Sub pRDEntryHKG()
Dim strTemp As String
Dim strPaidDueDescription As String
Dim strCCPayDescription As String
Dim strFMRCorpCCPayDescription As String
Dim strFMRPersCCPayDescription As String
Dim dtAcctDate As Date
Dim PaidByCC As Boolean
Dim curTaxTotal As Currency
Dim curGross As Currency
Dim curCWTDisc As Currency
Dim curPubFare As Currency

'29122004
If cmbPx.listindex = mintStartPxNum Then
'If cmbPx.ListIndex = 0 Then
    intI = 1
    ReDim strRDLine(0)
End If

If gbolFMR = True Then
    If gstrCCVendor <> "" Then
        strFMRCorpCCPayDescription = gstrCCVendor & "XXXXXXXXXXX" & Right(gstrCCNum, 4)
    End If
    If gstrPersCCVendor <> "" Then
        strFMRPersCCPayDescription = gstrPersCCVendor & "XXXXXXXXXXX" & Right(gstrPersCCNum, 4)
    End If
End If

With gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1)
    If cmbFOP(2).Text = "CC" Then
        PaidByCC = True
        'strCCPayDescription = "CREDIT CARD PAYMENT-" & .FOP_CCCode
        strCCPayDescription = cmbFOP(3).Text & "XXXXXXXXXXX" & Right(txtTktMod(2).Text, 4)
    Else
        PaidByCC = False
        strCCPayDescription = ""
    End If
    If .Cat35 = False Then
       curTaxTotal = .TaxTotal
    Else
       curTaxTotal = .TktTaxTotal
    End If
    
    If .Cat35 = False Then
        'modified on 04/05/06 use equivamount, if base different with total
        If .BaseCurrency = .TotalCurrency Then
            curPubFare = .BaseAmount
        Else
            curPubFare = .EquivAmount
        End If
    Else
        If .TktBaseCurrency = .TktTotalCurrency Then
            curPubFare = .TktBaseAmount
        Else
            curPubFare = .TktEquivAmount
        End If
    End If
End With

With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
    'dtAcctDate = DateAdd("M", 3, gobjPNR.AirSeg(gobjPNR.AirSegCount).DepartDateTime)
    'ZhiSam - V1.2.23 20130911 - CR-289 - E-Invoice: Due and Paid Lines Date
    If bfunctCheckRTLine = True Then
        dtAcctDate = dtfunctRTDate
    Else
        dtAcctDate = DateAdd("d", 90, Date)
    End If
    'modified on 31/03/2005: use base fare value
    'modified on 28 mar
    'modified on 30/05/2005: add merchant fee to base fare for normal ticket
    ReDim Preserve strRDLine(intI)
    If gbolFMR = False Then
    'If fConvertZero(txtAComm) > 0 Then
    '  If .NetAmount = 0 Then
    '    strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(txtFareInfo(0).Text) + .MerchAmt + fConvertZero(txtAComm)), gstrAgcyCurrFormat) & _
     '           " TAXES " & Format(txtFareInfo(1).Text, gstrAgcyCurrFormat) & "*" & _
     '           Format(fConvertZero(txtFareInfo(0).Text) + fConvertZero(txtFareInfo(1).Text) + .MerchAmt + fConvertZero(txtAComm), gstrAgcyCurrFormat)
    '  Else
     '   strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(txtFareInfo(0).Text) + fConvertZero(txtAComm)), gstrAgcyCurrFormat) & _
    '            " TAXES " & Format(txtFareInfo(1).Text, gstrAgcyCurrFormat) & "*" & _
     '           Format(fConvertZero(txtFareInfo(0).Text) + fConvertZero(txtFareInfo(1).Text) + fConvertZero(txtAComm), gstrAgcyCurrFormat)
    '  End If
    'Else
      If .Cat35 = False Then
         If .NetAmount = 0 Then
            strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(txtFareInfo(0).Text) + .MerchAmt), gstrAgcyCurrFormat) & _
                  " TAXES " & Format(txtFareInfo(1).Text, gstrAgcyCurrFormat) & "*" & _
                  Format(fConvertZero(txtFareInfo(0).Text) + fConvertZero(txtFareInfo(1).Text) + .MerchAmt, gstrAgcyCurrFormat)
         Else
           strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(txtFareInfo(0).Text)), gstrAgcyCurrFormat) & _
                  " TAXES " & Format(txtFareInfo(1).Text, gstrAgcyCurrFormat) & "*" & _
                  Format(fConvertZero(txtFareInfo(0).Text) + fConvertZero(txtFareInfo(1).Text), gstrAgcyCurrFormat)
         End If
      Else
         If .NetAmount = 0 Then
            strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(.TktBaseAmount) + .MerchAmt), gstrAgcyCurrFormat) & _
                  " TAXES " & Format(.TktTaxTotal, gstrAgcyCurrFormat) & "*" & _
                  Format(fConvertZero(.TktBaseAmount) + fConvertZero(.TktTaxTotal) + .MerchAmt, gstrAgcyCurrFormat)
         Else
           strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(.TktBaseAmount)), gstrAgcyCurrFormat) & _
                  " TAXES " & Format(.TktTaxTotal, gstrAgcyCurrFormat) & "*" & _
                  Format(fConvertZero(.TktBaseAmount) + fConvertZero(.TktTaxTotal), gstrAgcyCurrFormat)
         End If
      End If
    'End If
    Else
        'If (.DiscountAmt > .TransactionFee) And .TransactionFee > 0 Then
        '   strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format(gdblAmtToCom - gdblTaxToCom - .TransactionFee + .DiscountAmt, gstrAgcyCurrFormat) & _
        '            " TAXES " & Format(gdblTaxToCom, gstrAgcyCurrFormat) & "*" & _
        '            Format(gdblAmtToCom + .DiscountAmt - .TransactionFee, gstrAgcyCurrFormat)
        'Else
        '    strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format(gdblAmtToCom - gdblTaxToCom + .DiscountAmt, gstrAgcyCurrFormat) & _
        '            " TAXES " & Format(gdblTaxToCom, gstrAgcyCurrFormat) & "*" & _
        '            Format(gdblAmtToCom + .DiscountAmt, gstrAgcyCurrFormat)
        'End If
        
        'Modified on 240806 - Get the rebate from FMR mask instead of auto compute
            strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "DDMMM") & "*AIR TICKET " & Format(gdblAmtToCom - gdblTaxToCom + gdblRebate, gstrAgcyCurrFormat) & _
                              " TAXES " & Format(gdblTaxToCom, gstrAgcyCurrFormat) & "*" & _
                              Format(gdblAmtToCom + gdblRebate, gstrAgcyCurrFormat)

    End If
    intI = intI + 1
    
    
    
    If .NetAmount > 0 Then
        strPaidDueDescription = "CWT FARE DISCOUNT"
        curCWTDisc = 0
        curGross = .BaseAmount + .Commission
        If curGross <> IIf(.Cat35 = False, .BaseAmount, .TktBaseAmount) Then
            'Added on 24 Nov 08 by Jeremy to minus Hidden Comm
            curCWTDisc = curPubFare - curGross - .MerchAmt - .HiddenComm
            'curCWTDisc = curPubFare - curGross - .MerchAmt
        Else
            'Add on 24 Nov 08 by Jeremy to minus Hidden Comm
            curCWTDisc = curPubFare - .NetAmount - .MerchAmt - .HiddenComm
            'curCWTDisc = curPubFare - .NetAmount - .MerchAmt
        End If
        If curCWTDisc > 0 Then
            
            ReDim Preserve strRDLine(intI)
            
            strRDLine(intI) = "RP.T/" & Format(dtAcctDate, "ddmmm") & "*" & strPaidDueDescription & "*" & _
                Format(curCWTDisc, gstrAgcyCurrFormat)
            intI = intI + 1
            'gobjHost.TerminalEntry strTemp
        End If
    End If

    
    If .DiscountAmt > 0 Then
        strPaidDueDescription = "CLIENT DISCOUNT"

            ReDim Preserve strRDLine(intI)

        strRDLine(intI) = "RP.T/" & Format(dtAcctDate, "ddmmm") & "*" & strPaidDueDescription & "*" & _
            Format(.DiscountAmt, gstrAgcyCurrFormat)
        intI = intI + 1
        'gobjHost.TerminalEntry strTemp
    End If
        
    If PaidByCC Then

        ReDim Preserve strRDLine(intI)
        If gbolFMR = False Then
            strRDLine(intI) = "RP.T/" & Format(dtAcctDate, "ddmmm") & "*" & strCCPayDescription & "*" & _
                Format(.SellAmount - .DiscountAmt + curTaxTotal, gstrAgcyCurrFormat)
            intI = intI + 1
            'gobjHost.TerminalEntry strTemp
        Else
            'FMR
            'Print the Split FOP here for Reference (except for Rebate/Discount)
             'modified on 28 mar
            If gstrCCVendor <> "" Then
                 'Modified on 240806 - Get the rebate from FMR mask instead of auto compute
                If (gdblRebate = .DiscountAmt - .TransactionFee) And .TransactionFee > 0 Then
                'If (.DiscountAmt > .TransactionFee) And .TransactionFee > 0 Then
                strRDLine(intI) = "RP.T/" & Format(dtAcctDate, "ddmmm") & "*" & strFMRCorpCCPayDescription & "*" & _
                        Format(gdblAmtToCom - .TransactionFee, gstrAgcyCurrFormat)
    
                Else
                    strRDLine(intI) = "RP.T/" & Format(dtAcctDate, "ddmmm") & "*" & strFMRCorpCCPayDescription & "*" & _
                        Format(gdblAmtToCom, gstrAgcyCurrFormat)
                End If
                  
            End If
            'If gstrPersCCVendor <> "" Then
            '    strRDLine(intI) = "RT.T/" & Format(dtAcctDate, "ddmmm") & "*" & strFMRPersCCPayDescription & " " & _
            '       Format(gstrPersAmt, gstrAgcyCurrFormat)
                intI = intI + 1
            'End If
        End If
        
    End If
    
    If .TransactionFee > 0 Then
        strPaidDueDescription = "TRANSACTION FEE"

            ReDim Preserve strRDLine(intI)

        strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "ddmmm") & "*" & strPaidDueDescription & "*" & _
            Format(.TransactionFee, gstrAgcyCurrFormat)
        'gobjHost.TerminalEntry strTemp
        intI = intI + 1
        
        If PaidByCC Then
         
                ReDim Preserve strRDLine(intI)

            strRDLine(intI) = "RP.T/" & Format(dtAcctDate, "ddmmm") & "*" & strCCPayDescription & "*" & _
                Format(.TransactionFee, gstrAgcyCurrFormat)
            'gobjHost.TerminalEntry strTemp
            intI = intI + 1
            
    
        End If
    End If
    
    
     If .FuelSurcharge > 0 Then
        strPaidDueDescription = "FUEL CHARGE SVC FEE"

            ReDim Preserve strRDLine(intI)

        strRDLine(intI) = "RD.T/" & Format(dtAcctDate, "ddmmm") & "*" & strPaidDueDescription & "*" & _
            Format(.FuelSurcharge, gstrAgcyCurrFormat)
        'gobjHost.TerminalEntry strTemp
        intI = intI + 1
        
        If PaidByCC Then
         
                ReDim Preserve strRDLine(intI)

            strRDLine(intI) = "RP.T/" & Format(dtAcctDate, "ddmmm") & "*" & strCCPayDescription & "*" & _
                Format(.FuelSurcharge, gstrAgcyCurrFormat)
            'gobjHost.TerminalEntry strTemp
            intI = intI + 1
            
    
        End If
    End If
    
    
    
    If cmbPx.listindex = cmbPx.ListCount - 1 Then
        If UBound(strRDLine) <> 0 Then
            For intI = 1 To UBound(strRDLine)
        
                gobjHost.terminalEntry (strRDLine(intI))
        
            Next intI
        End If
    End If
    
End With


End Sub
Private Sub pDIEntryHKG(strFOP As String)
Dim strEntry As String
Dim strDiscCode As String
Dim curDiscAmt As Currency
Dim rs As ADODB.Recordset
Dim strSecLine As String

Set rs = gdbConn.Execute("Select sortkey from tblProductCodes where productcode = '50'")
If rs.EOF Then
    strDiscCode = ""
Else
    strDiscCode = rs!SortKey & ""
End If
rs.Close
Set rs = Nothing

With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
    If .DiscountAmt > 0 Then
        curDiscAmt = -Format(CStr(.DiscountAmt), gstrAgcyCurrFormat)
        
        'strEntry = "DI.FT-MS/PC" & strDiscCode & "/VCWT" & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum & _
         'IIf(strFOP = "", "+DI.FT-MSX/A" & curDiscAmt & "/SF" & curDiscAmt & "/C" & curDiscAmt & "/FS", _
                           "+DI.FT-MSX/A" & curDiscAmt & "/SF" & curDiscAmt & "/C" & curDiscAmt & "+DI.FT-MSX/F" & strFOP & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text _
        '    & "/D" & curDiscAmt)
        
         strEntry = "DI.FT-MS/PC" & strDiscCode & "/" & VendorNum("50", "CWT") & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum
         'strSecLine = "/A" & curDiscAmt & "/SF" & curDiscAmt & "/C" & curDiscAmt & "/FS" & "/TCW" & strSegmentNo
         strSecLine = "/A" & curDiscAmt & "/SF" & curDiscAmt & "/C" & curDiscAmt & "/FS" & strSegmentNo
         strEntry = strEntry & splitLongMSX(strSecLine)
        'added on 08/12: include FF10,11 in MS line for BTA clients
        '14/01/05: FF10,11 handled by Client MI screen
        'If txtMI(4).Text <> "" And txtMI(5).Text <> "" Then
        '    strEntry = strEntry & "+DI.FT-MSX/FF10-" & txtMI(4).Text & "/FF11-" & txtMI(5).Text
        'End If
        
        'added on 17/01/05: copy all file-fare related MI to MSX line
        strEntry = strEntry & splitLongMSX(getMSLineforMI())

        gobjHost.terminalEntry strEntry
    End If
End With

End Sub
Private Function NFNotFound_SG()
Select Case cmbFareType.Text
    Case "SQP - SQ/MI PUBLISHED FARE", "APF - SPECIAL FARE", "PUB - PUBLISHED FARE"
        If fConvertZero(txtFareInfo(4).Text) = 0 Or fConvertZero(txtFareInfo(5).Text) = 0 Then
            NFNotFound_SG = True
        Else
            NFNotFound_SG = False
        End If
    End Select
End Function
Private Sub pCheckFBC()
Dim lngC As Long

If cmbFareType.Text <> "SQP - SQ/MI PUBLISHED FARE" _
   And cmbFareType.Text <> "PUB - PUBLISHED FARE" Then
   If cmbFareOnTkt.listindex = 0 Then
      For lngC = 0 To mbytNumSegs - 1
          If Len(txtPriceFBC(lngC).Text) > 0 And Len(txtFBC(lngC).Text) > 0 And _
             txtPriceFBC(lngC).Text <> txtFBC(lngC) Then
             mstrFBUFields = mstrFBUFields & "/FBC"
          End If
      Next lngC
   End If
End If
End Sub
Private Sub UATPControl(Index As Integer)

Select Case Index

Case 1:
'Added on 15/10/04: TMP card: uncheck UATP + FOP on PAX set to INV
    If cmbFOP(0).Text = "CC" Then
        'If Left(UCase(cmbFOP(1).Text), 2) = "DC" And Left(UCase(txtTktMod(1).Text), 7) = "3644033" Then
         'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
        If IsTMPCard(Left(UCase(cmbFOP(1).Text), 2), UCase(txtTktMod(1).Text)) Then
            chkUATP.value = 0
            chkUATP.Visible = False
            cmbFOP(2).Text = "INV"
        
        Else
            chkUATP.value = 1
            chkUATP.Visible = True
        End If
   End If

Case 2:
   If cmbFOP(2).Text = "CC" And cmbFOP(0) = "FMR" Then
           'If Left(UCase(cmbFOP(3).Text), 2) = "DC" And Left(UCase(txtTktMod(2).Text), 7) = "3644033" Then
            'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
            If IsTMPCard(Left(UCase(cmbFOP(3).Text), 2), UCase(txtTktMod(2).Text)) Then
                chkUATP.value = 0
                chkUATP.Visible = False
            Else
                chkUATP.value = 1
                chkUATP.Visible = True
            End If
    End If

End Select
End Sub



Private Sub txtAComm_KeyPress(KeyAscii As Integer)
   'KeyAscii = fAllowNumeric(KeyAscii, Mid(txtAComm.Tag, 3))
If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8 Or KeyAscii = 46 Then
   If KeyAscii = 46 Then
     If InStr(txtAComm.Text, ".") Then
            KeyAscii = 0
            Exit Sub
        Else
            txtAComm.Text = txtAComm.Text
        End If
    Else
    End If
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtBag_Change(Index As Integer)
If mbolFormLoaded Then
    mstrFBUFields = mstrFBUFields & "/BAG"
    gobjLog.EventToLog "frmPricingWiz1.txtBag_Change"
End If
End Sub

Private Sub txtBag_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtBag_GotFocus(Index As Integer)
Call pSetSelected
End Sub



Private Sub txtDeliveryDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 32, 35, 42, 45, 47, 48 To 57, 65 To 90, 97 To 122
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select
    
End Sub

Private Sub txtFareInfo_Change(Index As Integer)
Select Case Index
    Case 3
        If txtFareInfo(3) = "" Then txtFareInfo(3) = "0"
    Case 6
        If mbolFareStored = True Then mstrFBUFields = mstrFBUFields & "/FCN"
        
End Select
End Sub

Private Sub txtFareInfo_GotFocus(Index As Integer)
Call pSetSelected

End Sub

Private Sub txtFareInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 4, 5, 8
            KeyAscii = fAllowNumeric(KeyAscii, Mid(txtFareInfo(Index).Tag, 3))
        Case 6
            KeyAscii = fAllowAlphaNumeric(KeyAscii, Mid(txtFareInfo(Index).Tag, 3))
        Case 3
           KeyAscii = fAllowNumeric(KeyAscii, Mid(txtFareInfo(Index).Tag, 3))
    End Select
End Sub

Private Sub txtFBC_Change(Index As Integer)
If mbolFormLoaded Then
    mstrFBUFields = mstrFBUFields & "/FBC"
    gobjLog.EventToLog "frmPricingWiz1.txtFBC_Change"
End If
End Sub

Private Sub txtFBC_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtFBC_GotFocus(Index As Integer)
Call pSetSelected

End Sub



Private Sub txtMI_GotFocus(Index As Integer)
Call pSetSelected

End Sub

Private Sub txtMI_KeyPress(Index As Integer, KeyAscii As Integer)

Select Case Index
    Case 0, 1
        KeyAscii = fAllowNumeric(KeyAscii, ".")
    Case 2
        KeyAscii = fAllowNumeric(KeyAscii)
    Case 3
        KeyAscii = fAllowAlpha(KeyAscii)
    Case 4, 5, 6
        KeyAscii = fAllowAlphaNumeric(KeyAscii)
    'Case 6
    '    KeyAscii = fAllowAlpha(KeyAscii)
End Select

End Sub

Private Sub txtNVA_Change(Index As Integer)
'If mbolFormLoaded Or gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).PrivateFare Then
If mbolFormLoaded Then
    
    mstrFBUFields = mstrFBUFields & "/NVA"
    gobjLog.EventToLog "frmPricingWiz1.txtNVA_Change"
End If
End Sub

Private Sub txtNVA_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtNVA_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtNVB_Change(Index As Integer)
'If mbolFormLoaded Or gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).PrivateFare Then
If mbolFormLoaded Then
    mstrFBUFields = mstrFBUFields & "/NVB"
    gobjLog.EventToLog "frmPricingWiz1.txtNVB_Change"
End If

End Sub

Private Sub txtNVB_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtNVB_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtPriceFBC_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtPriceFBC_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtTaxAmt_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtTktDesig_Change(Index As Integer)
    If mbolFormLoaded Then
    mstrFBUFields = mstrFBUFields & "/TKD"
    gobjLog.EventToLog "frmPricingWiz1.txtTktDesig_Change"
    End If
End Sub

Private Sub txtTktDesig_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtTktDesig_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtTktMod_Change(Index As Integer)
Select Case Index
    'Added on 15/10/04: TMP card: uncheck UATP + FOP on PAX set to INV
    Case 1, 2
        UATPControl (Index)
    Case 6, 7, 8
    'If mbolFormLoaded Then
        mstrFBUFields = mstrFBUFields & "/FCN"
        gobjLog.EventToLog "frmPricingWiz1.txtTktMod_Change"
    'End If
    End Select
    
End Sub

Private Sub txtTktMod_GotFocus(Index As Integer)

If Index = 9 Then txtTktMod(9).MaxLength = 15 - Len(txtTktMod(0).Text)
Call pSetSelected

End Sub

Private Sub txtTktMod_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index

    Case 0, 9
        Select Case KeyAscii
            Case 32, 45
                'allow
            Case Else
                KeyAscii = fAllowAlphaNumeric(KeyAscii)
        End Select
    
    Case 1, 2
        KeyAscii = fAllowNumeric(KeyAscii)
        
    Case 3
        Select Case KeyAscii
            Case 32, 35, 40 To 42, 45, 46
                'allow
            Case Else
                KeyAscii = fAllowAlphaNumeric(KeyAscii)
        End Select
            
    Case 4, 5
        Select Case KeyAscii
            Case 32, 36, 45, 46
                'allow
            Case Else
                KeyAscii = fAllowAlphaNumeric(KeyAscii)
        End Select

    Case 6 To 8
        Select Case KeyAscii
            Case 32, 35, 40 To 42, 45, 46, 47, 64
                'allow
            Case Else
                KeyAscii = fAllowAlphaNumeric(KeyAscii)
        End Select
        
    Case 10
        Select Case KeyAscii
            Case 46
                'allow
            Case Else
                KeyAscii = fAllowNumeric(KeyAscii)
        End Select
   End Select

End Sub

Private Sub txtValue_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Function GetALCodes() As String()
Dim lngC As Long
Dim strAL As String
Dim strSQL As String
Dim rsRec As New ADODB.Recordset
'Modified on 12/1/2005: for Fare Fare Only Option
If Not gobjFareQuotes(1).FQ(1).StoreFare Then

    strSQL = "Select distinct(Vendor) from tblFareSeg where [RecLoc] = '" & gobjPNR.RecLoc & "' and SegID=" & frmFareQuoteRequest.cboFQ.ItemData(frmFareQuoteRequest.cboFQ.listindex) & ""
    'Set rsRec = gdbFQ.OpenRecordset(strSQL)
    Set rsRec = gdbConn.Execute(strSQL)
        If Not rsRec.EOF Then rsRec.MoveFirst
            While Not rsRec.EOF
                strAL = strAL & IIf(strAL = "", "", "/") & rsRec!Vendor
                rsRec.MoveNext
            Wend
         GetALCodes = Split(strAL, "/")
    rsRec.Close
    Set rsRec = Nothing
  
Else
    For lngC = 1 To gobjPNR.AirSegCount
        With gobjPNR.AirSeg(lngC)
            If .SelectedForPricing = True Then
                If InStr(1, strAL, .Vendor) = 0 Then strAL = strAL & IIf(strAL = "", "", "/") & .Vendor
            End If
        End With
    GetALCodes = Split(strAL, "/")
   Next
End If
End Function

Private Function CurrencyFormat(ByVal CurrencyCode As String) As String
Dim rsCurr As New ADODB.Recordset
Dim strSQL As String
Dim strFormat As String
Dim intX As Integer
Dim intND As Integer 'Number Decimal places

strSQL = "SELECT * FROM tblCurrency WHERE [CurrencyCode] = '" & CurrencyCode & "'"

'Set rsCurr = gdbTProLU.OpenRecordset(strSQL)
Set rsCurr = gdbConn.Execute(strSQL)
rsCurr.MoveFirst
intND = rsCurr![Decimal]

strFormat = IIf(intND > 0, "#0.", "#0")
For intX = 1 To intND
    strFormat = strFormat & "0"
Next

strFormat = CurrencyFormat

End Function

Private Function RemoveChar(TestString As String, ByVal AlphaNumericBoth As String, ByVal AllowSpace As Boolean, ByVal SpecChar As String) As String
Dim strTemp As String
Dim strChar As String
Dim lngC As Long

strTemp = UCase(TestString)

For lngC = 1 To Len(strTemp)
    strChar = Mid(strTemp, lngC, 1)
    Select Case Asc(strChar)
        Case 48 To 57                       'numeric
            If AlphaNumericBoth = "A" Then
                'remove
                 strTemp = Left(strTemp, (lngC - 1)) & " " & Mid(strTemp, (lngC + 1))
            End If
        
        Case 65 To 90       'alpha
            If AlphaNumericBoth = "N" Then
                'remove
                strTemp = Left(strTemp, (lngC - 1)) & " " & Mid(strTemp, (lngC + 1))
            End If

        Case Else
            If InStr(SpecChar, strChar) = 0 Then
                'remove
                strTemp = Left(strTemp, (lngC - 1)) & " " & Mid(strTemp, (lngC + 1))
            End If
    End Select
Next

If AllowSpace = False And InStr(strTemp, " ") <> 0 Then
    While (InStr(strTemp, " ") <> 0)
        strTemp = Left(strTemp, InStr(strTemp, " ") - 1) & Mid(strTemp, InStr(strTemp, " ") + 1)
    Wend
End If

RemoveChar = strTemp
           
End Function
Private Sub FOPCForMI()
    'Silk Air FOP Code:
    '1. FOP = INVAGT:
    '         QSCode to be placed in FOP -> FINVAGT.QS1234
    '
    '2. FOP = CC:
    '         QSCode to be placed in Tour Code (if no existing TC)
    '         QSCode to be placed in Endorsement (if TC exists)

    If cmbValCarrier.Text = "MI" And txtTktMod(3).Text <> "" And cmbFOP(0).Text = "CC" Then
        If txtTktMod(0).Text <> "" Then
            If txtTktMod(4).Text <> "" Then
                If txtTktMod(5).Text <> "" Then
                    txtTktMod(5).Text = txtTktMod(5).Text & "." & txtTktMod(3).Text
                Else
                    txtTktMod(5).Text = txtTktMod(3).Text   'EB#3
                End If
            Else
                txtTktMod(4).Text = txtTktMod(3).Text       'EB#2
            End If
        Else
            txtTktMod(0).Text = txtTktMod(3).Text           'Tour Code
        End If
    End If
        
End Sub

Private Sub SGFF()
Dim strMsg As String
Dim strFareCalc As String

'gobjLog.EventToLog "frmPricingWiz1.cmdDone_Click"

'On Error GoTo ProcErr
' #### need to move the mbolsysendo to bottom

    If Not validData Then Exit Sub
    
    'Removed 10/Nov/05
    'If cmbFareOnTkt.listindex = 0 And cmbFareType.Text = "Special Fare" Then
    '   mstrFBUFields = mstrFBUFields & "/FBC"
    '   mstrFBUFields = mstrFBUFields & "/NVB"
    '   mstrFBUFields = mstrFBUFields & "/NVA"
    '   mstrFBUFields = mstrFBUFields & "/BAG"
    'End If
    
    frmWait.Show
    If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
    
    Call pCheckFBC
    gobjLog.LineTextToLog "mstrFBUFields = " & mstrFBUFields
    
    'FQ   1 time only
    'Use 1st Px Account code
    '29122004
    'If cmbPx.ListIndex = 0 Then
    If mbolFQ = False Then
       If Not FileFare Then
           Unload frmWait
           Exit Sub
       End If
       gobjLog.LineTextToLog "FileFare = True"
    End If
    
    gobjPNR.LoadFiledFare (CStr(mbytFFNum))
    
    'If Not validMI Then
    '   If cmbPx.listindex = mintStartPxNum Then
    '      gobjHost.terminalEntry "FX" & mbytFFNum
    '      If blnFirstPx = True Then mbolFQ = False
    '      blnCancelFF = True
    '      gobjHost.terminalEntry ("R.TPRO PRICE")
    '      gobjHost.terminalEntry ("ER")
    '      gobjHost.terminalEntry ("ER")
    '      gobjHost.terminalEntry ("ER")
    '    End If
    '    Unload frmWait
    '    Exit Sub
    'End If

    'Call UpdateTktData
    With txtFareInfo(6)
        strFareCalc = gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).FareConstructText
        .Text = Mid(strFareCalc, 1, IIf((InStr(1, strFareCalc, "END ROE") - 1) > 0, InStr(1, strFareCalc, "END ROE") - 1, Len(strFareCalc)))
        '.Text = gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).FareConstructText
        If txtTktMod(6).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(6).Text)
        If txtTktMod(7).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(7).Text)
        If txtTktMod(8).Text <> "" Then .Text = Trim(.Text) & " " & Trim(txtTktMod(8).Text)
    End With
        
    If Not UpdateTktData Then
       gobjHost.terminalEntry "FBF", True
       If cmbPx.listindex = mintStartPxNum Then
          gobjHost.terminalEntry "FX" & mbytFFNum
          If blnFirstPx = True Then mbolFQ = False
          blnCancelFF = True
       ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
        If gIntModuleType = gModuleType.SYEX Then
                ' do not run ER for SyEx flow
        Else
                gobjHost.terminalEntry ("R.TPRO PRICE")
                gobjHost.terminalEntry ("ER")
                gobjHost.terminalEntry ("ER")
                gobjHost.terminalEntry ("ER")
        End If

        End If
        Unload frmWait
        Exit Sub
    End If
    
    gobjPNR.LoadFiledFare (CStr(mbytFFNum)), True
            
    'TMU add for 1Px the rest add to NP.TM*, NP cannot accept '@'
    If Not AddTMU_SG Then
        gobjHost.terminalEntry "FX" & mbytFFNum
        If blnFirstPx = True Then mbolFQ = False
        blnCancelFF = True
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    
        If gIntModuleType = gModuleType.SYEX Then
                ' do not run ER for SyEx flow
        Else
                gobjHost.terminalEntry ("R.TPRO PRICE")
                gobjHost.terminalEntry ("ER")
                gobjHost.terminalEntry ("ER")
                gobjHost.terminalEntry ("ER")
        End If
    
 
        Unload frmWait
        Exit Sub
    End If
    gobjLog.LineTextToLog "AddTMU = True"
    
    'If cmbPx.ListIndex = 0 Then
    Call pFillInFFData
    'End If
    
    'Call pChkETNoSurcharge
    
    If Me.optCommissionType(1).value Or Me.optDiscType(1).value Then
        'Amend TMU 'Can remove
        Call AddNF_SG
    End If
    
    'Added on 16/2/2005: handle ASF
    'If gbolFMR = True Then
    '    If Not AddASF Then
    '        gobjHost.TerminalEntry "FX" & mbytFFNum
    '        If blnFirstPx = True Then mbolFQ = False
    '        'blnCancelFF = True
    '        gobjHost.TerminalEntry ("R.TPRO PRICE+ER")
    '        gobjHost.TerminalEntry ("ER")
    '        gobjHost.TerminalEntry ("ER")
    '        Unload frmWait
    '        Exit Sub
    '    End If
        
    '    gobjLog.LineTextToLog "Add_ASFTMU = True"
    'End If
    
    'FB build/override fare
    'Call UpdateTktData
    mbolFareStored = True
    
    'DI and RD line
    
    Call WriteDataToGDS_SG
    
    Unload frmClientMI
    
    If cmbPx.listindex <> cmbPx.ListCount - 1 Then
       cmbPx.listindex = cmbPx.listindex + 1
       ClearVar
       clearControls
       blnChgPCarrier = False
       Call pSetInitialValues
       blnFirstPx = False
       blnCancelFF = False
       Unload frmWait
       Exit Sub
    End If
    
    Call updateDeliveryAddr
    
    'Added on 1/2/05: flag NRCC for Aqua process
    flagNRCC (mbytFFNum)
    
    
    Call pAddFareTypeNP
    
    gobjHost.ENDPNR "TPRO PRICE", True
    'pSendToFP "R.File Fare"
    'pSendToFP "ER"
    'pSendToFP "*FF" & mbytFFNum
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    
        If gIntModuleType = gModuleType.SYEX Then
                ' do not run ER for SyEx flow
        Else
                gobjHost.terminalEntry "R.File Fare+ER"
        End If

    'gobjHost.TerminalEntry "ER"
    
    
    'Added on 22/07/04 - include remarks for AQUA checking
    gobjHost.terminalEntry "NP.S*APMI SCRIPT COMPLETED+NP.SS*VBIFF"
    pDisplayToFP "*FF" & mbytFFNum
    'Added on 1/2/05 - indicate file fare completed by VBI
    'gobjHost.TerminalEntry "NP.SS*VBIFF"
    
    ''FMR
    'If gbolFMR = True Then
    '   Call FMR
    '   strRes = gobjHost.TerminalEntry(mstrFMRCmd)
    '   If InStr(1, strRes, "TICKET MODIFIERS UPDATED") = 0 Then
    '      MsgBox "Unable to add FMR." & vbCrLf & strRes
    '   End If
    'End If

'Added on 14/10/04: add to VBI log table
'Timer
Call pAddToVBILog(gobjPNR.RecLoc, "File Fare", startTime, SysStart, "File Fare", , startTime)


    'added on 6/1/05: delete FQ record after file fare
    If cmbPx.listindex = cmbPx.ListCount - 1 And Not gobjFareQuotes(1).FQ(1).StoreFare Then
        Call delFQRec
    End If
    
    WriteToLog

    'Added on 26/08/04: Queue to ticketing after file fare
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    
        If gIntModuleType = gModuleType.SYEX Then
                ' Remove Tiket Queue Form from the SyEx flow
        Else
                Load frmTktQueue
                frmTktQueue.Show
                Do
                        DoEvents
                Loop Until isLoaded("frmTktQueue") = False
        End If
        
 
    Unload Me
    Unload frmWait
        
    'If Pretrip exist, load frmPreTtrip
    'If CheckPreTrip = True Then
    '    frmPreTrip.Show
    'End If
    'Call pRedisplayMenu
    
Exit Sub

ProcErr:
Select Case Err.Number
    Case vbObjectError + 615
        'MsgBox "Unexpected response from GDS!" & vbCrLf & "You will need to finish manually."
        strMsg = "Unexpected response from GDS!" & vbCrLf & "You will need to finish manually."
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Case Else
        'MsgBox "ERROR " & Err.Number & vbCrLf _
            & Err.Description, "RUN TIME ERROR"
        strMsg = "ERROR " & Err.Number & vbCrLf _
            & Err.Description
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
            Resume Next
    End Select

Unload Me
'Call pRedisplayMenu
Exit Sub

End Sub
Private Sub updateDeliveryAddr()
Dim strResp As String

strResp = ""
'Added on 26/10/04: Update Delivery Date/time in Delivery field
If Len(Trim(txtDeliveryDate.Text)) > 0 Then
    strResp = gobjHost.terminalEntry("D.@")
    strResp = gobjHost.terminalEntry("D." & Trim(txtDeliveryDate.Text) & IIf(Right(Trim(txtDeliveryDate.Text), 1) = "*", "", " ") & Replace(gobjPNR.DeliveryAddress, "@", "*"))
End If
End Sub
Private Function AddTMU_SG() As Boolean
Dim strCmd As String
Dim strRes As String
Dim sngDocFee As Single
Dim strTC As String
Dim strMsg As String
'''???
'29122004
'If cmbPx.ListIndex = 0 Then
If cmbPx.listindex = mintStartPxNum And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
   If cmbFareOnTkt.listindex > 0 Then gobjHost.terminalEntry "TMU" & mbytFFNum & "/TC@"
End If

'added on 14/12: Interpret FOP Code for Silk Air
Call FOPCForMI

strCmd = ""
strCmd = "TMU" & mbytFFNum

If optCommissionType(0).value And ((gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False) Or (gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = True And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).CommType = False)) Then
    strCmd = strCmd & "/Z" & txtFareInfo(3).Text
End If

If fConvertZero(txtFareInfo(8).Text) > 0 And chkDocFee.value = 1 Then
    sngDocFee = fConvertZero(txtFareInfo(8).Text)
Else
    sngDocFee = 0
End If

'If txtFareInfo(4).Text <> "" Then
'    If cmbFareType.Text <> "SQ/MI Published Nett Fare" And _
'       cmbFareType.Text <> "SQ Corporate Fare" Then
'       strCmd = strCmd & "/NF" & gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).TotalCurrency & Format(txtFareInfo(4).Text, gstrAgcyCurrFormat)
'    End If
'End If
'11012005 if IT should not have NF and ASF Requested by Sok Leng
If txtFareInfo(4).Text <> "" And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
    If cmbFareType.Text <> "SQN - SQ/MI PUBLISHED NETT FARE" And _
       cmbFareType.Text <> "SQC - CORPORATE FARE (A)" Then
       If cmbFareOnTkt.listindex <> 2 And cmbFareOnTkt.listindex <> 4 Then
          strCmd = strCmd & "/NF" & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency & Format(txtFareInfo(4).Text, gstrAgcyCurrFormat)
       End If
    End If
End If



'If txtTktMod(0).Text <> "" Then
    '''Select Case gstrAgcyCountryCode 'using select case to allow for more countries
    '''    Case "HK"
    '''        strCmd = strCmd & "/AI-" & txtTktMod(0).Text
    '''
    '''        If txtTktMod(9).Text <> "" Then
    '''            strCmd = strCmd & "/VC-" & txtTktMod(9).Text
    '''        End If
    '''
    '''    Case Else
  
'Modified on 7/02/05: Generate TC for BI,CA,MH
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
    If blnChgPCarrier = True Then
        If cmbFareType.Text = "APF - SPECIAL FARE" Then strTC = getTCMapper(cmbValCarrier.Text, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).BaseAmount)
            If strTC <> "" Then
                strCmd = strCmd & "/TC" & strTC
            ElseIf txtTktMod(0).Text <> "" Then
                strCmd = strCmd & "/TC" & txtTktMod(0).Text
            End If
    ElseIf txtTktMod(0).Text <> "" Then
    'ElseIf txtTktMod(0).Text <> "" Then
        strCmd = strCmd & "/TC" & txtTktMod(0).Text
    End If
End If
            
    '''End Select
'End If

Select Case cmbFOP(0).Text
    Case "INV"
        'modified on 14/09/04:BCODE format: /FINVAGT.BCODE...
        strCmd = strCmd & "/F" & cmbFOP(0).Text & IIf(txtTktMod(3) <> "", "AGT." & txtTktMod(3), "AGT")
    Case "MS"
        strCmd = strCmd & "/F" & cmbFOP(0).Text & IIf(txtTktMod(3) <> "", txtTktMod(3), "")
    Case "CC"
        strCmd = strCmd & "/F" & cmbFOP(1).Text & txtTktMod(1).Text & "*D" & Format(dtpCCExpDate(0).value, "mmyy")
            If cmbFareType.Text <> "SQN - SQ/MI PUBLISHED NETT FARE" And _
                cmbFareType.Text <> "SQC - CORPORATE FARE (A)" Then
                'Added on 15/10/04
                If chkShowASF.value = 1 Then
                    'Added on 14/09/04
                    'If UCase(cmbFOP(1).Text) = "DC" And Left(UCase(txtTktMod(1).Text), 7) = "3644033" Then
                    '    strCmd = strCmd & "/ASF" & Format(txtFareInfo(4).Text, gstrAgcyCurrFormat)
                    'Else
                    '    strCmd = strCmd & "/ASF" & Format(CSng(txtFareInfo(5).Text) + sngDocFee, gstrAgcyCurrFormat)
                    '    '''If gstrAgcyCountryCode = "HK" Then gobjHost.TerminalEntry "N.P1@*NF"
                    'End If
                    '11012005 if IT should not have NF and ASF Requested by Sok Leng
                    If cmbFareOnTkt.listindex <> 2 And cmbFareOnTkt.listindex <> 4 Then
                       'If UCase(cmbFOP(1).Text) = "DC" And Left(UCase(txtTktMod(1).Text), 7) = "3644033" Then
                       If IsTMPCard(Left(UCase(cmbFOP(1).Text), 2), UCase(txtTktMod(1).Text)) Then
                           sngASF = fConvertZero(txtFareInfo(4).Text)
                           strCmd = strCmd & "/ASF" & Format(sngASF, gstrAgcyCurrFormat)
                            
                       Else
                       'modified on 16/2/2005: doc fee
                            'strCmd = strCmd & "/ASF" & Format(fConvertZero(txtFareInfo(5).Text) + sngDocFee, gstrAgcyCurrFormat)
                            sngASF = fConvertZero(txtFareInfo(5).Text) - sngDocFee
                            strCmd = strCmd & "/ASF" & Format(sngASF, gstrAgcyCurrFormat)
                            'strCmd = strCmd & "/ASF" & Format(CSng(txtFareInfo(5).Text), gstrAgcyCurrFormat)
                           '''If gstrAgcyCountryCode = "HK" Then gobjHost.TerminalEntry "N.P1@*NF"
                       End If
                    End If
                End If
            End If

    
     
End Select
'Modified on 22 Feb: capture second EB in TMU

'If txtTktMod(4).Text <> "" Then
'    strCmd = strCmd & "/EB" & txtTktMod(4).Text
'    If txtTktMod(5).Text <> "" Then
'        strCmd = strCmd & "*EB" & txtTktMod(5).Text
'    End If
'End If

If txtTktMod(4).Text <> "" Then
    strCmd = strCmd & "/EB" & txtTktMod(4).Text
End If
If txtTktMod(4).Text = "" And txtTktMod(5).Text <> "" Then
    strCmd = strCmd & "/EB" & txtTktMod(5).Text
End If
If txtTktMod(4).Text <> "" And txtTktMod(5).Text <> "" Then
    strCmd = strCmd & "*EB" & txtTktMod(5).Text
End If

'removed on 13/12: PT/ET is specified during FQ command
'strCmd = strCmd & IIf(chkPaperTkt.Value = vbChecked, "/PT", "/ET")
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35 = False Then
    Select Case cmbFareOnTkt.listindex
        Case 2, 4
            strCmd = strCmd & "/IT" & IIf(cmbFareOnTkt.listindex = 4, "*PC", "")
        Case 3, 5
            strCmd = strCmd & "/BT" & IIf(cmbFareOnTkt.listindex = 5, "*PC", "")
    End Select
End If
'29122004
'If cmbPx.ListIndex = 0 Then
If cmbPx.listindex = mintStartPxNum Then

strRes = gobjHost.terminalEntry(strCmd)
If InStr(1, strRes, "DBI AIRPLUS INTERNATIONAL DESCRIPTIVE BILLING") > 0 Then
   AddTMU_SG = False
   blnCancelFF = True
   'MsgBox "Unable to add TMU!" & Chr(13) & Chr(13) & "Please inform operation manager to disable the DBI screen in Galileo"
   strMsg = "Unable to add TMU!" & Chr(13) & Chr(13) & "Please inform operation manager to disable the DBI screen in Galileo"
   modMsgBox.OKMsg = "OK"
   modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
   Exit Function
End If
If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
    AddTMU_SG = True
ElseIf InStr(strRes, "ERROR 8516 - INVALID FORMAT/DATA - MODIFIER ALREADY EXISTS") > 0 Then
    gobjHost.terminalEntry "TMU" & mbytFFNum & "/TC@"
    strRes = gobjHost.terminalEntry(strCmd)
    If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
        AddTMU_SG = True
    Else
        AddTMU_SG = False
        blnCancelFF = True
        'MsgBox "Unable to add TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
        strMsg = "Unable to add TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        Exit Function
    End If
Else
    AddTMU_SG = False
    blnCancelFF = True
    'MsgBox "Unable to add TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
    strMsg = "Unable to add TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Exit Function
End If

    'FMR
    If cmbFOP(0).Text = "FMR" Then
   
        If Not AddASF Then
            'gobjHost.TerminalEntry "FX" & mbytFFNum
            'If blnFirstPx = True Then mbolFQ = False
            AddTMU_SG = False
            blnCancelFF = True
            'gobjHost.TerminalEntry ("R.TPRO PRICE+ER")
            'gobjHost.TerminalEntry ("ER")
            'gobjHost.TerminalEntry ("ER")
            'Unload frmWait
            Exit Function
        End If
        
        gobjLog.LineTextToLog "Add_ASFTMU = True"
       
       If Not FMR Then
          AddTMU_SG = False
          blnCancelFF = True
          Exit Function
       End If
    End If

Else
    'Add to NP.TM*
    AddTMU_SG = True
    Dim intPos As Integer
    Dim intStart As Integer
    Dim intLength As Integer
    Dim intPrevious As Integer
    Dim strPrevious As String
    Dim i As Integer
    Dim j As Integer
    'FMR
    If cmbFOP(0).Text = "FMR" Then
       strCmd = strCmd & "/FMR"
    End If
    
    If Len(strCmd) > 72 Then
    i = 0
    Do While Len(strCmd) > 72
        intPrevious = 0
        intPos = 0
        intStart = 0
        j = 0
        Do Until intPos > 72
        intPrevious = intPos
            intPos = InStr(IIf(j = 0, 1, intPos + 1), strCmd, "/")
            If intPos = 0 Then Exit Do
            If j = 0 And intPos > 75 Then
            intPrevious = 72
            Exit Do
            End If
            j = j + 1
        Loop
        

        intPos = IIf(j = 0, intPrevious, intPrevious - 1)
        intLength = intPos - intStart
        strPrevious = strCmd
        strCmd = Mid(strCmd, 1, intLength)
        
        strCmd = "NP.TM*" & "FF" & Format(mbytFFNum, "00") & "PX" & Format(cmbPx.listindex + 1, "00") & ":" & ConvertNPText(strCmd)
        strRes = gobjHost.terminalEntry(strCmd)
        
        strCmd = Mid(strPrevious, IIf(j = 0, intPos + 1, intPos + 2))

        i = i + 1
        Loop
        
        strCmd = "NP.TM*" & "FF" & Format(mbytFFNum, "00") & "PX" & Format(cmbPx.listindex + 1, "00") & ":" & ConvertNPText(strCmd)
        strRes = gobjHost.terminalEntry(strCmd)

    
    Else
        strCmd = "NP.TM*" & "FF" & Format(mbytFFNum, "00") & "PX" & Format(cmbPx.listindex + 1, "00") & ":" & ConvertNPText(strCmd)
        strRes = gobjHost.terminalEntry(strCmd)
    End If
End If

End Function

Private Function AddNF_SG() As Boolean
Dim sngBF As Single
Dim sngNF As Single
Dim sngASF As Single
Dim sngDisc As Single
Dim sngComm As Single
Dim strCmd As String
Dim strRes As String
Dim strMsg As String

sngBF = CSng(txtFareInfo(0).Text)
sngASF = sngBF

If txtFareInfo(3).Text = "" Then txtFareInfo(3).Text = "0"
If txtFareInfo(7).Text = "" Then txtFareInfo(7).Text = "0"

If optCommissionType(1).value = True Then
    sngNF = sngBF - CSng(txtFareInfo(3).Text)
Else
    sngComm = fCurrRound(sngBF - (sngBF * (1 - (CSng(txtFareInfo(3).Text) * 0.01))), gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency, "DOWN")
    sngNF = sngBF - sngComm
End If

If optDiscType(1).value = True Then
    sngDisc = fCurrRound(sngBF - (sngBF * (1 - (CSng(txtFareInfo(7).Text) * 0.01))), gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency, "DOWN")
    sngNF = sngNF - sngDisc
    sngASF = sngBF - sngDisc
End If

'Removed on 14/10/04: this is done in AddTMU
'If cmbFareType.Text <> "SQ/MI Published Nett Fare" And _
'   cmbFareType.Text <> "SQ Corporate Fare" Then
'   strCmd = "TMU" & mbytFFNum & "/NF" & gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).TotalCurrency & Format(sngNF, gstrAgcyCurrFormat)
'End If

'Removed on 14/09/04: this is done in AddTMU
'If cmbFareType.Text <> "SQ/MI Published Nett Fare" And _
'   cmbFareType.Text <> "SQ Corporate Fare" Then
'   If cmbFOP(0).Text = "CC" Then
'       strCmd = strCmd & "/ASF" & Format(sngASF, gstrAgcyCurrFormat)
'   End If
'End If

If Trim(strCmd) = "" Then
   AddNF_SG = True
   Exit Function
End If

strRes = gobjHost.terminalEntry(strCmd)

If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
    AddNF_SG = True
ElseIf InStr(strRes, "ERROR 8516 - INVALID FORMAT/DATA - MODIFIER ALREADY EXISTS") > 0 Then
    gobjHost.terminalEntry "TMU" & mbytFFNum & "/NF@"
    gobjHost.terminalEntry "TMU" & mbytFFNum & "/ASF@"
    strRes = gobjHost.terminalEntry(strCmd)
    If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
        AddNF_SG = True
    Else
        AddNF_SG = False
        'MsgBox "Unable to add NF/ASF TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
        strMsg = "Unable to add NF/ASF TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End If
Else
    AddNF_SG = False
    'MsgBox "Unable to add NF/ASF TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
    strMsg = "Unable to add NF/ASF TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End If

End Function

Private Sub WriteDataToGDS_SG()
Dim strTemp As String
Dim strEntry As String
Dim dtmNewDate As Date
Dim strNewDate As String
Dim sngTotCharge As Single
Dim sngTotChargeAC As Single
Dim sngTmp As Single
Dim rs As New ADODB.Recordset
Dim strTrxnFeeCode As String
Dim strTFVendor As String
Dim sngCM As Single
Dim strDiscCode As String

Dim intCount As Integer
Dim sngTotalInvDue As Single
Dim sngTotalAmtToCC As Single
Dim sngTktBaseAmt As Single
Dim booIsTMPCard As Boolean
Dim sngDocFee As Single
Dim sngDisc As Single
Dim blnFOP1IsTMP As Boolean
Dim strTFFOPCharge As String
Dim sngTransFee As Single
Dim strSecLine As String

intCount = 0
sngTktBaseAmt = 0

'Preethi - V1.2.6 20110907 - CR 90 - Change OBT Tool Code in FF35
Dim strBookingTool As String

If blnLoaded = False Then
    mintDisplayNo = mintDisplayNo + 1
End If
'If chkNRCC.Value = False Then
'   strTFVendor = "VCWT"
'Else
'   strTFVendor = "VMER"
'End If
'Add a flag in NP line to indicate this is a Cat35 (for AQUA use)
addNP mintDisplayNo, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Cat35
'dtmNewDate = DateAdd("M", 6, Date)
'ZhiSam - V1.2.23 20130911 - CR-289 - E-Invoice: Due and Paid Lines Date
If bfunctCheckRTLine = True Then
    dtmNewDate = dtfunctRTDate
Else
    dtmNewDate = DateAdd("d", 90, Date)
End If
sngCM = 0
sngDisc = 0

If IsNumeric(txtTransFee) Then
   If CSng(txtTransFee) > 0 And chkNRCC.value = 1 Then
      sngTransFee = Format(txtTransFee, gstrAgcyCurrFormat)
   End If
End If

'If fConvertZero(txtFareInfo(8).Text) > 0 And chkDocFee.Value = 0 Then
'    sngDocFee = fConvertZero(txtFareInfo(8).Text)
'Else
'    sngDocFee = 0
'End If
If fConvertZero(txtFareInfo(8).Text) > 0 And chkDocFee.value = 1 Then
    sngDocFee = fConvertZero(txtFareInfo(8).Text)
Else
    sngDocFee = 0
End If

If fConvertZero(txtFareInfo(7).Text) > 0 Then
    sngDisc = fConvertZero(txtFareInfo(7).Text)
End If

'FMR
'Modified on 7/2/2005: Doc fee, use Nett Fare + Doc Fee
If gbolFMR = False Then
   'If fConvertZero(txtFareInfo(4).Text) > 0 And fConvertZero(txtFareInfo(8).Text) > 0 Then
    'sngTotCharge = fConvertZero(txtFareInfo(5).Text) + fConvertZero(txtFareInfo(1).Text) + sngDocFee - sngDisc
    sngTotCharge = fConvertZero(txtFareInfo(5).Text) + fConvertZero(txtFareInfo(1).Text) - sngDocFee - sngDisc - sngTransFee
   If fConvertZero(txtAComm) = 0 Then
   
   Else
      'sngTotChargeAC = fConvertZero(txtFareInfo(4).Text) + fConvertZero(txtFareInfo(1).Text) + sngDocFee - sngDisc
      sngTotChargeAC = fConvertZero(txtFareInfo(4).Text) + fConvertZero(txtFareInfo(1).Text) - sngDocFee - sngDisc
   End If
   'Else
    'sngTotCharge = CSng(txtFareInfo(0).Text) + CSng(txtFareInfo(1).Text) + sngDocFee - sngDisc
   'End If
Else
   sngTotCharge = gdblAmtToCom
   sngTotChargeAC = gdblAmtToCom
   'sngTotCharge = gdblAmtToCom + sngDocFee
End If
'sngTotCharge = CSng(txtFareInfo(5).Text) + CSng(txtFareInfo(1).Text) + sngDocFee - sngDisc

Set rs = gdbConn.Execute("Select sortkey from tblProductCodes where productcode = '35'")
If rs.EOF Then
    strTrxnFeeCode = ""
Else
    strTrxnFeeCode = rs!SortKey & ""
End If
rs.Close

Set rs = gdbConn.Execute("Select sortkey from tblProductCodes where productcode = '50'")
If rs.EOF Then
    strDiscCode = ""
Else
    strDiscCode = rs!SortKey & ""
End If
rs.Close
Set rs = Nothing


If Me.cmbFOP(2).Text = "CC" Then
    If gbolFMR = False Then
        If chkUATP.value = 1 Or cmbFOP(0).Text = "CC" Then
           cmbFOP(3).Text = cmbFOP(1).Text
           txtTktMod(2).Text = txtTktMod(1).Text
           dtpCCExpDate(1).value = dtpCCExpDate(0).value
        End If
    End If
    
    Select Case cmbFOP(3).Text
        Case "AX"
            strTemp = "CX2"
        Case "DC"
            strTemp = "CX3"
        Case "VI", "CA"
            strTemp = "CX4"
        Case "TP"
            strTemp = "CX5"
    End Select
    
    'Check if it is TMP Card
    'If (Left(UCase(cmbFOP(3).Text), 2) = "DC" And _
        Left(UCase(txtTktMod(2).Text), 7) = "3644033") Then
     'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    If IsTMPCard(Left(UCase(cmbFOP(3).Text), 2), UCase(txtTktMod(2).Text)) Then
        booIsTMPCard = True
    Else
        booIsTMPCard = False
    End If

End If
'Added on 16/02/2005: check for TMP card in FOP1 for FMR
If gbolFMR = True Then
    'If (Left(UCase(gstrCCVendor), 2) = "DC" And _
        Left(UCase(gstrCCNum), 7) = "3644033") Then
     'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    If IsTMPCard(Left(UCase(gstrCCVendor), 2), UCase(gstrCCNum)) Then
        blnFOP1IsTMP = True
    Else
        blnFOP1IsTMP = False
    End If
End If

'mintDisplayNo = mintDisplayNo + 1
'Modified on 7/2/2005: Doc fee, use Nett Fare + Doc Fee
With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
    'FMR
    If gbolFMR = False Then
       'If (fConvertZero(txtFareInfo(4).Text) > 0 And fConvertZero(txtFareInfo(8).Text) > 0) Then
       If fConvertZero(txtAComm) = 0 Then
          'strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format((fConvertZero(txtFareInfo(5).Text)) + sngDocFee, gstrAgcyCurrFormat)
          'If chkNRCC.Value = 1 Then
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format((fConvertZero(txtFareInfo(5).Text)) - sngDocFee - sngTransFee, gstrAgcyCurrFormat)
          'Else
          '   strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format((fConvertZero(txtFareInfo(5).Text)) - sngDocFee, gstrAgcyCurrFormat)
          'End If
       Else
          'strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format((fConvertZero(txtFareInfo(4).Text)) + sngDocFee, gstrAgcyCurrFormat)
          strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format((fConvertZero(txtFareInfo(4).Text)) - sngDocFee, gstrAgcyCurrFormat)
       End If
       'Else
       'strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format((CSng(txtFareInfo(0).Text)) + sngDocFee, gstrAgcyCurrFormat)
      
       'End If
    Else
       'strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format(gdblAmtToCom - gdblTaxToCom , gstrAgcyCurrFormat)
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format(gdblAmtToCom - gdblTaxToCom + fConvertZero(txtFareInfo(7)), gstrAgcyCurrFormat)
    End If
     If .TransactionFee > 0 Then sngTmp = Format(CStr(.TransactionFee), gstrAgcyCurrFormat)

    'modified on 16/09/04: apply the following codes to all fares
    'If .PrivateFare = True Then
        '''strEntry = "DI.FT-SF/*" & mbytFFNum & "/" & Format(txtFareInfo(5).Text, gstrAgcyCurrFormat) & "+"
     
    If cmbFOP(2).Text = "CC" And Not booIsTMPCard Then
       If fConvertZero(txtAComm) = 0 Then
          strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/" & IIf(chkUATP.value = 1, "CC", strTemp) & "/" & cmbFOP(3).Text & txtTktMod(2).Text _
                     & "/" & Format(CStr(sngTotCharge), gstrAgcyCurrFormat)
       Else
          strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/" & IIf(chkUATP.value = 1, "CC", strTemp) & "/" & cmbFOP(3).Text & txtTktMod(2).Text _
                     & "/" & Format(CStr(sngTotChargeAC), gstrAgcyCurrFormat)
       End If
    'Else
    Else
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/CASH"
    End If
    
    'End If
    'Else
    '   strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-SF/*" & mintDisplayNo & "/" & Format(gdblAmtToCom - gdblTaxToPax + sngDocFee, gstrAgcyCurrFormat)
    '   If cmbFOP(2).Text = "CC" And Not booIsTMPCard Then
    '      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/" & IIf(chkUATP.Value = 1, "CC", strTemp) & "/" & cmbFOP(3).Text & txtTktMod(2).Text _
    '                & "/" & Format(sngTotCharge, gstrAgcyCurrFormat)
    '   Else
    '      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FOP/*" & mintDisplayNo & "/CASH"
    '   End If
    'End If
    'Added on 21/07/04 for FF8 and FF26
    If gobjPNR.CompInfo.MI = True Then
    If cboClassServ.Text <> "" Then
        strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF8/*" & mintDisplayNo & "/" & Trim(Left(cboClassServ, 2))
    End If
    'CS - Remove FF26 (Trip Type)
    'If cboTripType.Text <> "" Then
    '    Select Case UCase(cboTripType.Text)
    '        Case "ROUND"
    '            strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF26/*" & mintDisplayNo & "/" & "R"
    '        Case "ONE WAY"
    '            strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF26/*" & mintDisplayNo & "/" & "O"
    '    End Select
    'End If
    
    'CS - Add International or Domestic
    'If cboTrip.Text <> "" Then
    '   strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF41/*" & mintDisplayNo & "/" & Left(cboTrip.Text, 1)
    'End If
    
    'CS Add Booking Action
    If cboBookingAction.Text <> "" Then
       Select Case cboBookingAction.Text
          Case "AB - Agent Booked"
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF34/*" & mintDisplayNo & "/" & "AB"
          Case "EB - Self Booked"
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF34/*" & mintDisplayNo & "/" & "EB"
          Case "AA - Air Modified"
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF34/*" & mintDisplayNo & "/" & "AA"
          Case "AM - Multiple Modification"
             strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF34/*" & mintDisplayNo & "/" & "AM"
 
       End Select
    End If
    
    'CS Add Booking Tool
    If mstrBookingTool <> "" Then
    'Preethi - V1.2.6 20110907 - CR 90 - Change OBT Tool Code in FF35
       strBookingTool = getFF35OBT(Mid(cboBookingAction.Text, 1, 2), mstrBookingTool)
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF35/*" & mintDisplayNo & "/" & strBookingTool
    Else
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF35/*" & mintDisplayNo & "/" & "GAL"
    End If
    
    'CS Add Booking Method
    'If cboBookingMethod.Text <> "" Then
    '   Select Case UCase(cboBookingMethod.Text)
    '      Case "GDS"
    '         strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "G"
    '      Case "MANUAL"
    '         strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "M"
    '      Case "SELF BOOKING"
    '         strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "S"
    '   End Select
    'End If
    If mstrBookingTool <> "" Then
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "S"
    Else
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF36/*" & mintDisplayNo & "/" & "G"
    End If
    
    
    'CS - Add FF31 FF32
    If .TransactionFee > 0 Then
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF31/*" & mintDisplayNo & "/Y"
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF32/*" & mintDisplayNo & "/" & sngTmp
    Else
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF31/*" & mintDisplayNo & "/N"
    End If
    
    'Added FF38 for ticket type
    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF38/*" & mintDisplayNo & "/" & Left(chkPaperTkt.Caption, 1)
    
    'Added on 02/08/04 for FF10,FF11,FF19
    '14/01/05: FF10,11 handled by Client MI screen
    'If txtMI(4).Text <> "" Then
    '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF10/*" & mintDisplayNo & "/" & Trim(txtMI(4).Text)
    'End If
    'If txtMI(5).Text <> "" Then
    '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF11/*" & mintDisplayNo & "/" & Trim(txtMI(5).Text)
    'End If
    If txtMI(6).Text <> "" Then
        strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF81/*" & mintDisplayNo & "/" & UCase(Trim(txtMI(6).Text))
    End If
    
    'CS Add txtRS (ff30)
    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF30/*" & mintDisplayNo & "/" & txtRS
    'CS Change EC txtMI(2) --> txtMS
    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF7/*" & mintDisplayNo & "/" & txtMI(3).Text
    strEntry = strEntry & IIf(strEntry <> "", "+", "") _
        & "DI.FT-LF/*" & mintDisplayNo & "/" & txtMI(1).Text _
        & "+DI.FT-RF/*" & mintDisplayNo & "/" & txtMI(0).Text _
        & "+DI.FT-EC/*" & mintDisplayNo & "/" & txtMS
    End If
    '19/01/05: Write client MI data into DI
    If frmClientMI.DIfreefields <> "" Then
        strEntry = strEntry & IIf(strEntry <> "", "+", "") & frmClientMI.DIfreefields
    End If
    
'CS
'If .TransactionFee > 0 Then sngTmp = Format(CStr(.TransactionFee), gstrAgcyCurrFormat)
   
'Clement
If .TransactionFee > 0 Then
'added on 5/07/2005: if ASF < FF TKT Value then is TF FOP Charge CC

'If .BaseAmount > sngASF And sngASF > 0 Then
'    strTFFOPCharge = "CC"
'Else
'    strTFFOPCharge = strTemp
'End If
If chkNRCC.value = 1 Then
   strTFFOPCharge = "CC"
Else
   strTFFOPCharge = strTemp

End If



        'modified on 14/09/04: FOP for Trans Fee = /FS if using TMP card
   'FMR
   'If gbolFMR = False Then
   
         strEntry = strEntry & IIf(strEntry <> "", "+", "") & _
         "DI.FT-MS/PC" & strTrxnFeeCode & "/" & VendorNum(strTrxnFeeCode, IIf(chkNRCC.value = False, "CWT", "MER")) & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum

'"+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & strSegmentNo & "+DI.FT-MSX/F" & strTFFOPCharge & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "EXP" & Format(dtpCCExpDate(1).value, "MMYY") & "/D" & Format(CStr(.TransactionFee), gstrAgcyCurrFormat), _
         'Preethi - V1.2.4 20110527 - CR 59 - Remove Round Up Function on Transaction Fee
         strSecLine = "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & Format(sngTmp, "#0.00") & "/SF" & Format(sngTmp, "#0.00") & "/C" & Format(sngTmp, "#0.00") & strSegmentNo
         
               
       ' strSecLine = IIf(Len(strTemp) > 0 And Not booIsTMPCard, _
       '         "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & strSegmentNo, _
       '         "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & "/FS" & strSegmentNo)

        
        strSecLine = splitLongMSX(strSecLine)
        strEntry = strEntry & strSecLine & IIf(Len(strTemp) > 0 And Not booIsTMPCard, "+DI.FT-MSX/F" & strTFFOPCharge & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "EXP" & Format(dtpCCExpDate(1).value, "MMYY") & "/D" & Format(CStr(.TransactionFee), gstrAgcyCurrFormat), "+DI.FT-MSX/FS")
   'Else
   '      strEntry = strEntry & IIf(strEntry <> "", "+", "") & _
   '      "DI.FT-MS/PC" & strTrxnFeeCode & "/" & strTFVendor & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum & _
   '      IIf(Len(strTemp) > 0 And Not booIsTMPCard, _
   '             "+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & "+DI.FT-MSX/F" & strTemp & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "/D" & Format(CStr(.TransactionFee), gstrAgcyCurrFormat), _
   '             "+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & "/FS")
   'End If
         'added on 08/12: include FF10,11 in MS line for BTA clients
         '14/01/05: FF10,11 handled by Client MI screen
         'If txtMI(4).Text <> "" And txtMI(5).Text <> "" Then
         '   strEntry = strEntry & "+DI.FT-MSX/FF10-" & txtMI(4).Text & "/FF11-" & txtMI(5).Text
         'End If
            
        'added on 17/01/05: copy all file-fare related MI to MSX line
        strEntry = strEntry & splitLongMSX(getMSLineforMI())
       
        
        strEntry = strEntry & "+DI.FT-MSX/FF TRANSACTION FEE"
        
        '16082005
        'If fConvertZero(txtAComm) <> 0 Then
        '   strEntry = strEntry & IIf(strEntry <> "", "+", "") & _
        '              "DI.FT-MS/PC" & strTrxnFeeCode & "/" & VendorNum(strTrxnFeeCode, "MER") & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum & _
        '              IIf(Len(strTemp) > 0 And Not booIsTMPCard, _
        '              "+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & txtAComm & "/SF" & txtAComm & "/C" & txtAComm & "+DI.FT-MSX/F" & "CC" & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "/D" & Format(CStr(txtAComm), gstrAgcyCurrFormat), _
        '              "+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & txtAComm & "/SF" & txtAComm & "/C" & txtAComm & "/FS")
        '   strEntry = strEntry & splitLongMSX(getMSlineForMI())
        '   strEntry = strEntry & "+DI.FT-MSX/FF TRANSACTION FEE"
        
        'End If
         'IIf(strTemp = "" Or booIsTMPCard, "+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & "/FS", _
         '                  "+DI.FT-MSX/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & sngTmp & "/SF" & sngTmp & "/C" & sngTmp & "+DI.FT-MSX/" & IIf(strTFVendor = "VMER", "CC", "F" & strTemp) & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text _
         '   & "/D" & Format(CStr(.TransactionFee), gstrAgcyCurrFormat)) & _
         '   "+DI.FT-MSX/FF TRANSACTION FEE"
End If
'modified 180406: request by Sharon, MN client may have hidden TF also
        If fConvertZero(txtAComm) <> 0 Then
           strEntry = strEntry & IIf(strEntry <> "", "+", "") & _
                      "DI.FT-MS/PC" & strTrxnFeeCode & "/" & VendorNum(strTrxnFeeCode, "MER") & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum
          '      strSecLine = IIf(Len(strTemp) > 0 And Not booIsTMPCard, _
          '            "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & txtAComm & "/SF" & txtAComm & "/C" & txtAComm & strSegmentNo, _
          '            "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & txtAComm & "/SF" & txtAComm & "/C" & txtAComm & "/FS" & strSegmentNo)
           
          'Preethi - V1.2.4 20110527 - CR 59 - Remove Round Up Function on Transaction Fee
           strSecLine = "/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & Format(txtAComm, "#0.00") & "/SF" & Format(txtAComm, "#0.00") & "/C" & Format(txtAComm, "#0.00") & strSegmentNo
                     '"/" & IIf(gstrAgcyCountryCode = "HK", "A", "S") & txtAComm & "/SF" & txtAComm & "/C" & txtAComm & "/FS" & strSegmentNo)
           'Preethi - V1.2.4 20110527 - CR 59 - Remove Round Up Function on Transaction Fee
           strEntry = strEntry & splitLongMSX(strSecLine) & IIf(Len(strTemp) > 0 And Not booIsTMPCard, "+DI.FT-MSX/F" & "CC" & "/CCN" & cmbFOP(3).Text & txtTktMod(2).Text & "EXP" & Format(dtpCCExpDate(1).value, "MMYY") & "/D" & Format(CStr(txtAComm), "#0.00"), "+DI.FT-MSX/FS")
           strEntry = strEntry & splitLongMSX(getMSLineforMI())
           strEntry = strEntry & "+DI.FT-MSX/FF TRANSACTION FEE"
        End If
        
'Removed on 26/07/04 - do not need this hard-coded entry
'"+DI./0+DI.FT-MS/VCWT/TK0000000000000/PC35"

'Added on 04/10/04: Retrieve Ticket Face Value for CM calculation
'For intCount = 1 To gobjPNR.FiledFareCount
'    With gobjPNR.FiledFare(intCount).PX(cmbPx.ListIndex + 1)
'        sngTktBaseAmt = .BaseAmount
'    End With
'Next
'Modified on 22/12: retrieve base amount from current file fare/pax
sngTktBaseAmt = IIf(IsNumeric(txtFareInfo(0).Text), txtFareInfo(0).Text, 0)


If cmbFareType.Text = "APF - SPECIAL FARE" Then
    'sngCM = IIf(.EquivAmount > 0, .EquivAmount, .BaseAmount) - CSng(txtFareInfo(4).Text)
    sngCM = sngTktBaseAmt - fConvertZero(txtFareInfo(4).Text)
    If sngCM > 0 Then
       strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-CM/*" & mintDisplayNo & "/" & sngCM
    End If
End If

sngTotalInvDue = fConvertZero(txtFareInfo(5).Text) + fConvertZero(txtFareInfo(1).Text)
If strTemp = "" Then sngTotalInvDue = sngTotalInvDue + CSng(.TransactionFee)
If strTemp <> "" Then
   sngTotalAmtToCC = fConvertZero(txtFareInfo(5).Text) + fConvertZero(txtFareInfo(1).Text) + CSng(.TransactionFee)
   sngTotalAmtToCC = Format(sngTotalAmtToCC, gstrAgcyCurrFormat)
End If
  
'DI for Doc Fee not required when using TMP card
If cmbFareType.Text = "SQP - SQ/MI PUBLISHED FARE" Or cmbFareType.Text = "PUB - PUBLISHED FARE" Then
   If (fConvertZero(txtFareInfo(8).Text) > 0 And chkDocFee.value = 0 And Not booIsTMPCard And cmbFOP(2).Text = "CC") Or (fConvertZero(txtFareInfo(8).Text) > 0 And chkDocFee.value = 0 And gbolFMR And Not blnFOP1IsTMP) Then
      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MS/PC70/" & VendorNum("70", "DOCACM") & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum
      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/C" & Format(fConvertZero(txtFareInfo(8).Text), gstrAgcyCurrFormat) & "/FS" & "/TCW"
        'added on 08/12: include FF10,11 in MS line for BTA clients
        '14/01/05: FF10,11 handled by Client MI screen
        'If txtMI(4).Text <> "" And txtMI(5).Text <> "" Then
        '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/FF10-" & txtMI(4).Text & "/FF11-" & txtMI(5).Text
        'End If
        
        'added on 17/01/05: copy all file-fare related MI to MSX line
        strEntry = strEntry & splitLongMSX(getMSLineforMI())
        
      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/FF DOCUMENTATION FEE"
   End If
End If

'Added on 27/10/04: Commission Rebate
If sngDisc > 0 Then
   strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MS/PC" & strDiscCode & "/" & VendorNum(strDiscCode, "REBATE") & "/TKFF" & Format(mbytFFNum, "00") & "/PX" & .PxNum
   'IIf(strEntry <> "", "+", "") & "DI.FT-MSX/S" & Format(-sngDisc, gstrAgcyCurrFormat) & "/SF" & Format(-sngDisc, gstrAgcyCurrFormat) & "/C" & Format(-sngDisc, gstrAgcyCurrFormat) & "/FS" & "/TCW"
  'Preethi - V1.2.4 20110527 - CR 59 - Remove Round Up Function on Transaction Fee
   strSecLine = "/S" & Format(-sngDisc, "#0.00") & "/SF" & Format(-sngDisc, "#0.00") & "/C" & Format(-sngDisc, "#0.00") & strSegmentNo
    'added on 08/12: include FF10,11 in MS line for BTA clients
    '14/01/05: FF10,11 handled by Client MI screen
    'If txtMI(4).Text <> "" And txtMI(5).Text <> "" Then
    '    strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/FF10-" & txtMI(4).Text & "/FF11-" & txtMI(5).Text
    'End If
    'strEntry = strEntry & splitLongMSX(strSecLine) & "+DI.FT-MSX/FS" & "/TCW"
    strEntry = strEntry & splitLongMSX(strSecLine) & "+DI.FT-MSX/FS"
    'added on 17/01/05: copy all file-fare related MI to MSX line
    strEntry = strEntry & splitLongMSX(getMSLineforMI())
    
   strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-MSX/FF REBATE RETURN"
   
End If

'Preethi - V1.2.1 20101011 - CR21 - Nett Fare Mark Up
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount > 0 Then
   If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).HiddenComm = 0 Then
      If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount > gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ActualNetAmount And gobjFareQuotes(cmbPx.listindex + 1).FQ(1).ActualNetAmount > 0 Then
         strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-NF/*" & mintDisplayNo & "/" & Format(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount, gstrAgcyCurrFormat)
         'Preethi - V1.2.1 20101015 - CR21 - Nett Fare Mark Up
         strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-FF90/*" & mintDisplayNo & "/CWTF"
      End If
   Else
      strEntry = strEntry & IIf(strEntry <> "", "+", "") & "DI.FT-NF/*" & mintDisplayNo & "/" & Format(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount + gobjFareQuotes(cmbPx.listindex + 1).FQ(1).HiddenComm, gstrAgcyCurrFormat)
   End If
End If



SendDIEntry strEntry
'gobjHost.TerminalEntry strEntry
blnLoaded = False

'29122004
If cmbPx.listindex = mintStartPxNum Then
'If cmbPx.ListIndex = 0 Then
    intI = 1
    ReDim strRDLine(0)
End If
'Due/Paid line for Air Fares, CC Payment
'Modified on 18/09/04
'strRDLine = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*AIR TICKET " & Format(txtFareInfo(5).Text, gstrAgcyCurrFormat) & _
'            " TAXES " & Format(txtFareInfo(1).Text, gstrAgcyCurrFormat) & "*" & _
'            Format(CSng(sngTotCharge), gstrAgcyCurrFormat)


ReDim Preserve strRDLine(intI)
'FMR
If gbolFMR = False Then
'Modified on 7/02/04: DOC FEE
'If fConvertZero(txtFareInfo(4).Text) > 0 And fConvertZero(txtFareInfo(8).Text) > 0 Then
   If fConvertZero(txtAComm) = 0 Then
      'strRDLine(intI) = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(txtFareInfo(5).Text)) + sngDocFee - sngDisc, gstrAgcyCurrFormat) & _
      '         " TAXES " & Format(fConvertZero(txtFareInfo(1).Text), gstrAgcyCurrFormat) & "*" & _
      '         Format(CSng(sngTotCharge), gstrAgcyCurrFormat)
      strRDLine(intI) = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(txtFareInfo(5).Text)) - sngDocFee - sngDisc - sngTransFee, gstrAgcyCurrFormat) & _
               " TAX " & Format(fConvertZero(txtFareInfo(1).Text), gstrAgcyCurrFormat) & "*" & _
               Format(CSng(sngTotCharge), gstrAgcyCurrFormat)
   Else
      'strRDLine(intI) = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(txtFareInfo(4).Text)) + cdec(txtAComm) + sngDocFee - sngDisc, gstrAgcyCurrFormat) & _
      '         " TAXES " & Format(fConvertZero(txtFareInfo(1).Text), gstrAgcyCurrFormat) & "*" & _
      '         Format(CSng(sngTotCharge), gstrAgcyCurrFormat)
      strRDLine(intI) = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*AIR TICKET " & Format((fConvertZero(txtFareInfo(4).Text)) + CDec(txtAComm) - sngDocFee - sngDisc, gstrAgcyCurrFormat) & _
               " TAX " & Format(fConvertZero(txtFareInfo(1).Text), gstrAgcyCurrFormat) & "*" & _
               Format((fConvertZero(txtFareInfo(4).Text)) + CDec(txtAComm) - sngDocFee - sngDisc + fConvertZero(txtFareInfo(1).Text), gstrAgcyCurrFormat)
   End If
'Else
'   strRDLine(intI) = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*AIR TICKET " & Format((CSng(txtFareInfo(0).Text)) + sngDocFee - sngDisc, gstrAgcyCurrFormat) & _
'            " TAXES " & Format(txtFareInfo(1).Text, gstrAgcyCurrFormat) & "*" & _
'            Format(CSng(sngTotCharge), gstrAgcyCurrFormat)
'End If
Else
   strRDLine(intI) = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*AIR TICKET " & Format(gdblAmtToCom - gdblTaxToCom, gstrAgcyCurrFormat) & _
            " TAX " & Format(gdblTaxToCom, gstrAgcyCurrFormat) & "*" & _
            Format(CSng(sngTotCharge), gstrAgcyCurrFormat)
End If
intI = intI + 1



'gobjHost.TerminalEntry(strRDLine, True)


'Added on 24/08/04: Paid line for Air Fares
'FMR
'If gbolFMR = False Then
   'Remove on 26/7/2005: consolidate CC Paid lines
   'If cmbFOP(2).Text = "CC" Then
   '    If Not booIsTMPCard Then
   '        ReDim Preserve strRDLine(intI)
   '        strRDLine(intI) = "RP.T/" & Format(dtmNewDate, "DDMMM") & "*" & cmbFOP(3).Text & "XXXXXXXXXXX" & Right(txtTktMod(2).Text, 4) & "*" & _
   '                    Format(CSng(sngTotCharge), gstrAgcyCurrFormat)
   '        intI = intI + 1
   '        'gobjHost.TerminalEntry strRDLine
   '    End If
   ' End If
'Else
'   If cmbFOP(2).Text = "CC" Then
'       If Not booIsTMPCard Then
'           ReDim Preserve strRDLine(intI)
'           strRDLine(intI) = "RP.T/" & Format(dtmNewDate, "DDMMM") & "*" & cmbFOP(3).Text & "XXXXXXXXXXX" & Right(txtTktMod(2).Text, 4) & "*" & _
'                       Format(CSng(sngTotCharge), gstrAgcyCurrFormat)
'           intI = intI + 1
'           'gobjHost.TerminalEntry strRDLine
'       End If
'    End If
'End If

If .TransactionFee > 0 Then
    'Clement
    ReDim Preserve strRDLine(intI)
    strRDLine(intI) = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*TRANSACTION FEE*" & Format(sngTmp, gstrAgcyCurrFormat)
    intI = intI + 1
    'gobjHost.TerminalEntry strRDLine
    'Added on 24/08/04: Paid line for Transaction Fee
    'Remove on 26/7/2005: consolidate CC Paid lines
    'If cmbFOP(2).Text = "CC" Then
    '   If Not booIsTMPCard Then
    '       ReDim Preserve strRDLine(intI)
    '       strRDLine(intI) = "RP.T/" & Format(dtmNewDate, "DDMMM") & "*" & cmbFOP(3).Text & "XXXXXXXXXXX" & Right(txtTktMod(2).Text, 4) & "*" & Format(sngTmp, gstrAgcyCurrFormat)
    '       intI = intI + 1
    '       'gobjHost.TerminalEntry strRDLine
    '   End If
    'End If
End If

'Added on 26/7/2005: Consolidate CC Paid lines
If cmbFOP(2).Text = "CC" Then
       If Not booIsTMPCard Then
           ReDim Preserve strRDLine(intI)
           strRDLine(intI) = "RP.T/" & Format(dtmNewDate, "DDMMM") & "*" & cmbFOP(3).Text & "XXXXXXXXXXX" & Right(txtTktMod(2).Text, 4) & "*" & Format(sngTotCharge + sngTmp, gstrAgcyCurrFormat)
           intI = intI + 1
           'gobjHost.TerminalEntry strRDLine
       End If
End If

If cmbPx.listindex = cmbPx.ListCount - 1 Then
 If UBound(strRDLine) <> 0 Then
    For intI = 1 To UBound(strRDLine)
 
        gobjHost.terminalEntry strRDLine(intI)
 
    Next intI
 End If
End If

End With
'strRDLine = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
'            "*0 PERCENT GST ON NON TAXABLE CHARGE*" & Format(0, gstrAgcyCurrFormat)
'gobjHost.TerminalEntry strRDLine
'strRDLine = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
'            "*INVOICE TOTAL DUE*" & Format(sngTotalInvDue, gstrAgcyCurrFormat)
'gobjHost.TerminalEntry strRDLine
'strRDLine = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
'            "*TTL AMOUNT CHARGE TO CREDIT CARD*" & Format(sngTotalAmtToCC, gstrAgcyCurrFormat)
'gobjHost.TerminalEntry strRDLine


'gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).TotAmount

'230108: Add Filefare agent for productivity tracking for HK
Dim strFFAgtSignon As String
Dim OSLineNo As Integer
    OSLineNo = OSLineNum
    If gobjHost.AgentGACode <> "" Then
    'DI.5@FT-OS/1540
        strFFAgtSignon = IIf(OSLineNo > 0, "DI." & OSLineNo & "@FT-OS/", "DI.FT-OS/") & gobjHost.AgentGACode
       
    Else
    
        strFFAgtSignon = IIf(OSLineNo > 0, "DI." & OSLineNo & "@FT-OS/", "DI.FT-OS/") & gobjHost.AgentSine
    End If
    
    gobjHost.terminalEntry strFFAgtSignon

End Sub
Private Function GetVL() As Boolean
   Dim strRes As String
   Dim strVL As String
   Dim strText As String
   Dim strCarrier As String
   Dim rs As ADODB.Recordset
   
   strText = cmbValCarrier.Text & " does not exist in Vendor Locator list" & _
             vbCrLf & "Please fill in Vendor Locator for " & cmbValCarrier.Text
   strCarrier = Trim(cmbValCarrier.Text)
   Set rs = gdbConn.Execute("Select VLCarrier from tblVLCarrier Where Carrier='" & strCarrier & "'")
   If rs.EOF = False Then
      strCarrier = rs!VLCarrier
   End If
   
   strRes = gobjHost.terminalEntry("*VL", True)
   If InStr(1, UCase(strRes), UCase(strCarrier)) = 0 Then
      strVL = InputBox(strText, "Missing Vendor Locator")
      If Trim(strVL) = "" Then
         GetVL = False
      Else
         gobjHost.terminalEntry "RL." & strCarrier & "*" & UCase(strVL)
         GetVL = True
      End If
   Else
      GetVL = True
   End If
      
   
End Function


Private Function ConvertNPText(NPText As String) As String
   
      Do Until InStr(NPText, "@") = 0
        NPText = Left(NPText, InStr(NPText, "@") - 1) & "$" & Mid(NPText, InStr(NPText, "@") + 1)
      Loop
   
  ConvertNPText = NPText
   
End Function

Private Sub ClearVar()
    mbytNumSegs = 0
    mbolFormLoaded = False
    mbolORPriceFBC = False
    mbolORTktFBC = False
    mstrCommType = ""
    mstrDiscType = ""
    mstrIT_BT = ""
    mstrFOPType = ""
    mstrFOPCCInfo = ""
    mstrFOPCode = ""
    mbolFareStored = False
    mstrFBUFields = ""
    'mbytFFNum = 0
    mbolValidData = False
    mstrBaseCurr = ""
    msngFareDiff = 0
    mstrTotalCurr = ""
    mstrFareCalc = ""
    mbolFBUMode = False
    mbolSysEndo = False
End Sub

Private Sub clearControls()

Dim lngC As Integer

    For lngC = 0 To gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareSegCount - 1
        chkConnection(lngC).value = False
        'txtFlightInfo(lngC).Text = ""
        txtFBC(lngC).Text = ""
        txtTktDesig(lngC).Text = ""
        txtPriceFBC(lngC).Text = ""
        txtNVB(lngC).Text = ""
        txtNVA(lngC).Text = ""
        txtBag(lngC).Text = ""
    Next lngC
    
    For lngC = 0 To 2
        txtFareInfo(lngC).Text = "PENDING"
        txtFareInfo(lngC).Locked = True
        txtFareInfo(lngC).Enabled = False
    Next lngC
    
    For lngC = 3 To txtFareInfo.Count - 1
        txtFareInfo(lngC).Text = ""
    Next lngC
    
    txtFareInfo(6).Locked = True
    txtFareInfo(6).Enabled = False
    
    optCommissionType(0).value = True
    chkShowASF.value = 1
    optDiscType(0).value = True
    chkDocFee.value = False
    cmbFareOnTkt.listindex = 0
    cmbFareType.listindex = -1
    chkNRCC.value = False
    
    txtTktMod(1).Text = ""
    txtTktMod(2).Text = ""
    txtTktMod(0).Text = ""
    txtTktMod(9).Text = ""
    
    cmbFOP(0).listindex = 0
    cmbFOP(2).listindex = -1
    cmbFOP(1).Visible = False
    cmbFOP(3).Visible = False
    txtTktMod(1).Visible = False
    txtTktMod(2).Visible = False
    
    
    With dtpCCExpDate(0)
        '.MinDate = DateAdd("M", -1, Date)
        '.MaxDate = DateAdd("yyyy", 10, Date)
        .value = DateAdd("m", 6, Date)
        .Visible = False
    End With
    With dtpCCExpDate(1)
        '.MinDate = DateAdd("M", -1, Date)
        '.MaxDate = DateAdd("yyyy", 5, Date)
        '.Value = DateAdd("m", 6, Date)
        .Visible = False
    End With
    
    
    txtMI(0).Text = ""
    txtMI(1).Text = ""
    txtMI(3).Text = ""
    txtMI(6).Text = ""
    
    'CS - Remove FF26 (Trip Type)
    'cboTripType.listindex = -1
    'CS - Add International or Domestic
    cboTrip.listindex = 0
    'CS Add Booking Method
    'cboBookingMethod.listindex = 0
    'CS Add Booking Action
    cboBookingAction.listindex = 0
    cboClassServ.listindex = 0

End Sub

Private Sub delFQRec()

Dim lngI As Long
Dim strSQL As String

    For lngI = 1 To cmbPx.ListCount
    
        strSQL = "Delete from tblFareQuote where [RecLoc] = '" & gobjPNR.RecLoc & "' and [SegID]=" & gFQSegID & ""
        strSQL = strSQL & " and [PxID]= " & gobjFareQuotes(lngI).FQ(1).PxNum & ""
        
       gdbConn.Execute (strSQL)
            
    Next lngI
End Sub

'FMR
Private Function FMR() As Boolean
   Dim strRes As String
   Dim strCmd As String
   Dim strTemp As String
   Dim strSearch As String
   Dim dblTotAmt As Double
   Dim strFirstLine As String
   Dim dblTotTax As Double
   Dim b As String
   Dim intCount As Integer
   Dim strMsg As String
   
   intCount = 0
   FMR = True
StartAgain:
Set gobjPNR = New CWT_GalileoPNR3.PNR
   gobjPNR.loadPNR
   
   'gbolFMR = True
   'dblTotTax = gobjPNR.FiledFare(1).PX(1).TaxTotal
   If gobjPNR.FiledFareCount < mbytFFNum Then
        FMR = False
        'MsgBox "No File Fare Found/File Fare Does Not Match."
        strMsg = "No File Fare Found/File Fare Does Not Match."
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        Exit Function
   End If
   dblTotTax = CDec(gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1).TaxTotal)
    
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    
        If gIntModuleType = gModuleType.SYEX Then
                ' ZhiSam - V1.2.18 20121015 - CR 196 - Smart Point SyEx version : enable ER if the user select "FMR" at cmbFOP(0)
                ' otherwise do not run ER for SyEx flow
                Select Case cmbFOP(0).Text
                    Case "FMR"
                        strRes = gobjHost.terminalEntry("R.FMR+ER")
                        strRes = gobjHost.terminalEntry("ER")
                        strRes = gobjHost.terminalEntry("ER")
                    
                End Select
                
        Else
                strRes = gobjHost.terminalEntry("R.FMR+ER")
                strRes = gobjHost.terminalEntry("ER")
                strRes = gobjHost.terminalEntry("ER")
        End If

   strSearch = "*MR TOTAL AMOUNT RECEIVABLE"
   strRes = gobjHost.terminalEntry("TMU" & mbytFFNum & "F@")
   strRes = gobjHost.terminalEntry("TMU" & mbytFFNum & "FMR")
   
   If InStr(1, strRes, "ERROR 8528") Then
      If intCount >= 4 Then GoTo UnableToFMR
        ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
        If gIntModuleType = gModuleType.SYEX Then
                ' do not run ER for SyEx flow
        Else
                strRes = gobjHost.terminalEntry("R.FMR+ER")
                strRes = gobjHost.terminalEntry("ER")
        End If
      

      intCount = intCount + 1
      GoTo StartAgain
   ElseIf InStr(1, strRes, "ERROR ") Then
UnableToFMR:
      FMR = False
      'MsgBox strRes, vbCritical, "TPro FMR Error"
      strMsg = strRes
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      Exit Function
   End If
   If InStr(1, strRes, strSearch) = 0 Then GoTo UnableToFMR
   strTemp = Trim(Mid(strRes, InStr(1, strRes, strSearch) + Len(strSearch)))
   strTemp = Mid(strTemp, 1, InStr(1, strTemp, " ") - 1)
   dblTotAmt = strTemp
   gdblTotAmt = dblTotAmt
   gdblTax = dblTotTax
   
   frmFMR.Show
   Do
      DoEvents
   Loop Until isLoaded("frmFMR") = False
   
   If gbolFMR = False Then
      FMR = False
      Exit Function
   End If
   
   
   strFirstLine = Mid(strRes, 1, InStr(1, strRes, "/F·") - 2)
   strCmd = strFirstLine
   strCmd = strCmd & "/F·" & gstrFOP(0) & Dot(59 - Len(gstrFOP(0)))
   If gstrFOP(1) <> "" Then
      strCmd = strCmd & "/F·" & gstrFOP(1) & IIf(gstrFOP(2) <> "", Dot(59 - Len(gstrFOP(1))), "")
   End If
   If gstrFOP(2) <> "" Then
      strCmd = strCmd & "/F·" & gstrFOP(2)
   End If
   
'DC36440336287004*D0505$200.00..............................
'   B = "*MR TOTAL AMOUNT RECEIVABLE    377.00 SGD" & _
' "/F·DC36440336287004*D0505$200.00.............................." & _
' "/F·INVAGT$177.00"
' a = gobjHost.TerminalEntry(B)
   mstrFMRCmd = strCmd
   strRes = gobjHost.terminalEntry(mstrFMRCmd)
   If InStr(1, strRes, "DBI AIRPLUS INTERNATIONAL DESCRIPTIVE BILLING") > 0 Then
      FMR = False
      'MsgBox "Unable to add FMR." & vbCrLf & vbCrLf & "Please inform operation manager to disable the DBI screen in Galileo"
      strMsg = "Unable to add FMR." & vbCrLf & vbCrLf & "Please inform operation manager to disable the DBI screen in Galileo"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      Exit Function
   End If
   If InStr(1, strRes, "TICKET MODIFIERS UPDATED") = 0 Then
      'If strRes = "IGNORE AND RE-RETRIEVE B/FILE" Then
      '   strRes = gobjHost.TerminalEntry("ER")
      '   strRes = gobjHost.TerminalEntry("ER")
      '   strRes = gobjHost.TerminalEntry("TMU1F@")
      '   strRes = gobjHost.TerminalEntry("TMU1FMR")

       '  strRes = gobjHost.TerminalEntry(mstrFMRCmd)
         
      'Else
         FMR = False
         'MsgBox "Unable to add FMR." & vbCrLf & vbCrLf & strRes
         strMsg = "Unable to add FMR."
         modMsgBox.OKMsg = "OK"
         modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      'End If
   Else
      FMR = True
      frmFareDiff.Show
      Do
        DoEvents
      Loop Until isLoaded("frmFareDiff") = False
   End If
   
   'strRes = gobjHost.TerminalEntry(mstrFMRCmd)
   
End Function
Private Function AddASF() As Boolean
Dim strCmd As String
Dim strRes As String
Dim sngDocFee As Single
Dim strMsg As String

AddASF = True
 'Added on 16/2/2005: handle ASF
     'If gstrFOPToCom = "CC" Then
     If cmbFOP(2) = "CC" Then
            If (cmbFareType.Text <> "SQN - SQ/MI PUBLISHED NETT FARE" And _
                cmbFareType.Text <> "SQC - CORPORATE FARE (A)") Then
                'Added on 15/10/04
                If chkShowASF.value = 1 Then
                       
                       strCmd = "TMU" & mbytFFNum & "/ASF"
                       'If fConvertZero(txtFareInfo(8).Text) > 0 And chkDocFee.Value = 0 Then
                       '     sngDocFee = fConvertZero(txtFareInfo(8).Text)
                       ' Else
                       '     sngDocFee = 0
                       ' End If
                       
                       If fConvertZero(txtFareInfo(8).Text) > 0 And chkDocFee.value = 1 Then
                            sngDocFee = fConvertZero(txtFareInfo(8).Text)
                        Else
                            sngDocFee = 0
                        End If
                       
                       
                            If cmbFareOnTkt.listindex <> 2 And cmbFareOnTkt.listindex <> 4 Then
                                 'If UCase(gstrCCVendor) = "DC" And Left(UCase(gstrCCNum), 7) = "3644033" Then
                                 'If UCase(cmbFOP(3)) = "DC" And Left(txtTktMod(2), 7) = "3644033" Then
                                  'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
                                 If IsTMPCard(Left(UCase(cmbFOP(3).Text), 2), UCase(txtTktMod(2).Text)) Then
                                     sngASF = fConvertZero(txtFareInfo(4).Text)
                                     strCmd = strCmd & Format(sngASF, gstrAgcyCurrFormat)
                                     
                                 Else
                                     'sngASF = fConvertZero(txtFareInfo(5).Text) + sngDocFee
                                     sngASF = fConvertZero(txtFareInfo(5).Text) - sngDocFee
                                     strCmd = strCmd & Format(sngASF, gstrAgcyCurrFormat)
                                     'strCmd = strCmd & Format(fConvertZero(txtFareInfo(5).Text) + sngDocFee, gstrAgcyCurrFormat)
                                 
                                 End If
                            End If
                     
                        strRes = gobjHost.terminalEntry(strCmd)

                        If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
                            AddASF = True
                        ElseIf InStr(strRes, "ERROR 8516 - INVALID FORMAT/DATA - MODIFIER ALREADY EXISTS") > 0 Then
                            gobjHost.terminalEntry "TMU" & mbytFFNum & "/ASF@"
                            strRes = gobjHost.terminalEntry(strCmd)
                            If InStr(strRes, "TICKET MODIFIERS UPDATED") > 0 Then
                                AddASF = True
                            Else
                                AddASF = False
                                'MsgBox "Unable to add ASF in TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
                                strMsg = "Unable to add ASF in TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
                                modMsgBox.OKMsg = "OK"
                                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                            End If
                        Else
                            AddASF = False
                            'MsgBox "Unable to add ASF in TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
                            strMsg = "Unable to add ASF in TMU!" & Chr(13) & Chr(13) & "Response from GDS is: " & Chr(13) & strRes
                            modMsgBox.OKMsg = "OK"
                            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                        End If
                                               
                       
                End If
            End If
     End If
End Function

Private Function Dot(Num As Integer) As String
   Dim i As Integer
   
   For i = 1 To Num
      Dot = Dot & "."
   Next
End Function



Private Sub loadClientMI()
Dim strMsg As String

If mintDisplayNo > 0 Then

    If isLoaded("frmClientMI") Then
        frmClientMI.Show
    Else
        cmdClientMI.Enabled = False
        Load frmClientMI
        frmClientMI.intLocation = 4
        frmClientMI.strPdtType = ""
        frmClientMI.intProdCode = 0
        If blnLoaded = False Then
            frmClientMI.intFileFareNum = mintDisplayNo + 1
        Else
            frmClientMI.intFileFareNum = mintDisplayNo
        End If
        'frmClientMI.intProdCode = frmOthSvcs.dbcProducts.BoundText
        frmClientMI.cmbMICat.Enabled = False
        frmClientMI.pGetClientMI (gobjPNR.CN)
        frmClientMI.Show
        cmdClientMI.Enabled = True
    End If
Else
    'MsgBox "Unable to retrieve file fare/display number!", vbCritical
    strMsg = "Unable to retrieve file fare/display number!"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End If
End Sub

Private Function getMSLineforMI() As String
    Dim lngC As Long
    Dim strMIData As String
    Dim strFF() As String
    'freefields format:{FF number1}-{FF value1}/{FF number2}-{FF value2}...
    Dim freefields As String
    Dim strTemp As String
    
    On Error GoTo err_getMSlineForMI
    
    'General MI data for all clients
    'If txtMI(3).Text <> "" Then
    '    freefields = freefields & IIf(freefields <> "", "/", "") & "7-" & txtMI(3).Text
    'End If
    If txtMI(6).Text <> "" Then
        freefields = freefields & IIf(freefields <> "", "/", "") & "81-" & txtMI(6).Text
    End If
    'If cboClassServ.Text <> "" Then
    '    freefields = freefields & IIf(freefields <> "", "/", "") & "8-" & Trim(cboClassServ.Text)
    'End If
    'CS - Remove FF26 (Trip Type)
    'If cboTripType.Text <> "" Then
    '    Select Case UCase(cboTripType.Text)
    '        Case "ROUND"
    '            freefields = freefields & IIf(freefields <> "", "/", "") & "26-R"
    '        Case "ONE WAY"
    '            freefields = freefields & IIf(freefields <> "", "/", "") & "26-O"
    '    End Select
    'End If
   
    'CS Add Booking Action
    If cboBookingAction.Text <> "" Then
       Select Case cboBookingAction.Text
          Case "Agent Booked"
              freefields = freefields & IIf(freefields <> "", "/", "") & "34-AB"
          Case "Self Booked"
              freefields = freefields & IIf(freefields <> "", "/", "") & "34-EB"
          Case "Air Modified"
              freefields = freefields & IIf(freefields <> "", "/", "") & "34-AA"
       End Select
    End If
    
    'CS Add Booking Tool
    If mstrBookingTool <> "" Then
       freefields = freefields & IIf(freefields <> "", "/", "") & "35-" & mstrBookingTool
    Else
       freefields = freefields & IIf(freefields <> "", "/", "") & "35-GAL"
    End If
    
    'CS Add Booking Method
    'If cboBookingMethod.Text <> "" Then
    '   freefields = freefields & IIf(freefields <> "", "/", "") & "36-" & Left(cboBookingMethod.Text, 1)
    'End If
    If mstrBookingTool <> "" Then
       freefields = freefields & IIf(freefields <> "", "/", "") & "36-S"
    Else
       freefields = freefields & IIf(freefields <> "", "/", "") & "36-G"
    End If
    
    freefields = freefields & IIf(freefields <> "", "/", "") & "47-CWT"
    'CS - Add International or Domestic
    'If cboTrip.Text <> "" Then
    '   freefields = freefields & IIf(freefields <> "", "/", "") & "41-" & Left(cboTrip.Text, 1)
    'End If
    
    'Client specific MI data
    If isLoaded("frmClientMI") Then
        freefields = freefields & "/" & frmClientMI.getMSXFreeFields()
    End If
    
    strMIData = ""
    If freefields <> "" Then
        strFF = Split(freefields, "/")
        For lngC = LBound(strFF) To UBound(strFF)
            If strFF(lngC) <> "" Then
               'strMIData = strMIData & "/FF" & strFF(lngC)
               strTemp = Mid(strFF(lngC), 1, InStr(1, strFF(lngC), "-") - 1)
               strMIData = strMIData & IIf(IsNumeric(strTemp), "/FF", "/") & strFF(lngC)
            End If
        Next
    End If
    getMSLineforMI = strMIData
    Exit Function
err_getMSlineForMI:
    getMSLineforMI = ""
End Function


Private Function getSegmentSelected() As String
Dim intD As Integer
Dim strSQL As String
Dim rsRec2 As New ADODB.Recordset

intD = 0
    strSQL = "Select distinct(SegSeq),DepDate,COS,Vendor,ArrCityCode,FlightNum,DepCityCode,AdviceAct,SegmentSelectText from tblFareSeg S,tblFareQuote Q "
    strSQL = strSQL & "where Q.RecLoc = '" & gobjPNR.RecLoc & "' and Q.SegID=" & gFQSegID & " and Q.RecLoc=S.RecLoc and Q.SegID= S.SegID order by SegSeq"
    'Set rsRec2 = gdbFQ.OpenRecordset(strSQL)
    
    Set rsRec2 = gdbConn.Execute(strSQL)
    While Not rsRec2.EOF
        For intD = 1 To gobjPNR.AirSegCount
                If gobjPNR.AirSeg(intD).ArriveAirport = rsRec2!ArrCityCode And gobjPNR.AirSeg(intD).DepartAirport = rsRec2!DepCityCode And gobjPNR.AirSeg(intD).Vendor = rsRec2!Vendor And gobjPNR.AirSeg(intD).FlightNumber = rsRec2!FlightNum And gobjPNR.AirSeg(intD).DepartDateTime = rsRec2!DepDate Then
                    getSegmentSelected = getSegmentSelected & IIf(getSegmentSelected = "", gobjPNR.AirSeg(intD).segnumber, "." & gobjPNR.AirSeg(intD).segnumber)
                    Exit For
                End If

        Next intD

     
    rsRec2.MoveNext
    Wend
    rsRec2.Close

End Function
Private Sub flagNRCC(FFNum As Integer)
    Dim strResp As String
    'If chkNRCC.value = vbChecked Then
    If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NRCC Then
        strResp = gobjHost.terminalEntry("NP.TT*FF" & Format(FFNum, "00") & "-NRCC")
    End If
End Sub
'added on 1/4/2005: if comm>tf then fop-cc,comm<tf then fop-cx
Private Function getTFFOP(strTemp As String) As String
getTFFOP = strTemp
    With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
        If gbolFMR Then
            If .DiscountAmt < .TransactionFee Then
            '    getTFFOP = "CC"
            'Else
                getTFFOP = strTemp
            End If
        ElseIf .NRCC Then
            If .MerchAmt > 0 Then
                getTFFOP = "CC"
            Else
                getTFFOP = strTemp
            End If
        Else
            getTFFOP = strTemp
        End If
    End With
    
    
End Function



Private Sub SendDIEntry(entry As String)

Dim strDI() As String
Dim intI As Integer
Dim strEntry As String
'entry = "DI.FT-SF/*2/1233.00+DI.FT-FOP/*2/CC/AX376220243091006/1670.00+DI.FT-FF8/*2/YN+DI.FT-FF41/*2/I+DI.FT-FF34/*2/AB+DI.FT-FF35/*2/GAL+DI.FT-FF36/*2/G+DI.FT-FF31/*2/Y+DI.FT-FF32/*2/64+DI.FT-FF38/*2/E+DI.FT-FF19/*2/TG+DI.FT-FF30/*2/MC+DI.FT-FF7/*2/BKK+DI.FT-LF/*2/2335+DI.FT-RF/*2/2666+DI.FT-EC/*2/C+DI.FT-FF10/*1/2133232+DI.FT-MS/PC35/V032001/TKFF02/PX1+DI.FT-MSX/S64/SF64/C64+DI.FT-MSX/FCC/CCNAX376220243091006/D64.00+DI.FT-MSX/FF19-TG/FF34-AB/FF35-GAL/FF36-G+DI.FT-MSX/FF47-#CWT/FF10-2133232+DI.FT-MSX/FF TRANSACTION FEE+DI.FT-MS/PC35/V032001/TKFF02/PX1+DI.FT-MSX/S50/SF50/C50+DI.FT-MSX/FCC/CCNAX376220243091006/D50.00+DI.FT-MSX/FF19-TG/FF34-AB/FF35-GAL/FF36-G+DI.FT-MSX/FF47-#CWT/FF10-2133232+DI.FT-MSX/FF47-#CWT/FF10-2133232+DI.FT-MSX/FF47-#CWT/FF10-2133232"
strDI = Split(entry, "+")
'strEntry = ""
If UBound(strDI) < 0 Then Exit Sub
If UBound(strDI) < 29 Then
    gobjHost.terminalEntry entry
Else
    For intI = LBound(strDI) To UBound(strDI)
        If (intI Mod 28) = 0 And intI > 0 Then
            gobjHost.terminalEntry strEntry
            strEntry = ""
            strEntry = strDI(intI)
        Else
            strEntry = strEntry & IIf(strEntry = "", "", "+") & strDI(intI)
        End If
    Next intI
End If
    gobjHost.terminalEntry strEntry
End Sub

Private Function validMI() As Boolean
Dim sngDocFee As Single
Dim sngTransFee As Single
Dim strMsg As String

 validMI = True

If gstrAgcyCountryCode = "SG" Then
    If gbolFMR = False Then
       If IsNumeric(txtTransFee) Then
          If CSng(txtTransFee) > 0 And chkNRCC.value = 1 Then
             sngTransFee = Format(txtTransFee, gstrAgcyCurrFormat)
          End If
       End If

       If fConvertZero(txtFareInfo(8).Text) > 0 And chkDocFee.value = 1 Then
          sngDocFee = fConvertZero(txtFareInfo(8).Text)
       Else
          sngDocFee = 0
       End If
       
       If fConvertZero(txtAComm) = 0 Then
          strMsg = validMIMsg((fConvertZero(txtFareInfo(5).Text)) - sngDocFee - sngTransFee)
       Else
          strMsg = validMIMsg((fConvertZero(txtFareInfo(4).Text)) - sngDocFee)
       End If
    Else
       strMsg = validMIMsg(gdblAmtToCom - gdblTaxToCom + fConvertZero(txtFareInfo(7)))
    End If
Else
    If gbolFMR = False Then
       strMsg = validMIMsg(fConvertZero(txtFareInfo(5).Text))
    Else
       strMsg = validMIMsg(gdblAmtToCom - gdblTaxToCom + gdblRebate)
    End If
End If

If strMsg <> "" Then
   'MsgBox strMsg
   modMsgBox.OKMsg = "OK"
   modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
   validMI = False
End If
End Function

Private Function validMIMsg(ByVal sngSF As Single) As String
Dim strMsg As String
Dim sngTax As Single
Dim intI As Integer
Dim intJ As Integer
Dim strTemp() As String
Dim bolFound As Boolean
Dim sngDisc As Single

'Deduct discount amount from SF
If gstrAgcyCountryCode = "SG" Then
   If fConvertZero(txtFareInfo(7).Text) > 0 Then
      sngDisc = fConvertZero(txtFareInfo(7).Text)
   End If
Else
   sngDisc = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).DiscountAmt
End If

sngSF = sngSF - sngDisc

'Calculate total taxes from file fare and added taxes
For intI = 0 To lstTax.ListCount - 1
    strTemp = Split(lstTax.List(intI), " ")
    If UBound(strTemp) > 0 Then
       sngTax = sngTax + strTemp(0)
    End If
Next

With gobjPNR.FiledFare(mbytFFNum).PX(cmbPx.listindex + 1)
    For intI = 1 To .TaxCount
        bolFound = False
        For intJ = 0 To lstTax.ListCount - 1
            strTemp = Split(lstTax.List(intJ), " ")
            If UBound(strTemp) > 0 And UCase(strTemp(1)) = UCase(.Tax(intI).TaxCode) Then
               bolFound = True
               Exit For
            End If
        Next
        If bolFound = False Then
           sngTax = sngTax + .Tax(intI).Amount
        End If
    Next
End With

strMsg = ""
'RSA If RF > SF+TAX then XX cannot be selected. If RF = SF+TAX , then XX must be selected
If (txtMI(0) - (sngSF + sngTax)) > 0 Then
   If Trim(txtRS.Text) = "XX" Then
      strMsg = strMsg & "XX code in Realized Saving Code cannot be selected..." & Chr(13)
   End If
ElseIf (txtMI(0) - (sngSF + sngTax)) = 0 Then
   If Trim(txtRS.Text) <> "XX" Then
      strMsg = strMsg & "XX code in Realized Saving Code must be selected..." & Chr(13)
   End If
End If
'MSA If LF < SF+TAX, L cannot be selected. If LF = SF+TAX, then L must be selected
If (txtMI(1) - (sngSF + sngTax)) < 0 Then
   If Trim(txtMS.Text) = "L" Then
      strMsg = strMsg & "L code in Missing Saving Code cannot be selected..." & Chr(13)
   End If
ElseIf (txtMI(1) - (sngSF + sngTax)) = 0 Then
   If Trim(txtMS.Text) <> "L" Then
      strMsg = strMsg & "L code in Missing Saving Code must be selected..." & Chr(13)
   End If
End If
validMIMsg = strMsg
End Function
Private Sub pAddFareTypeNP()
Dim intI As Integer
For intI = 1 To gobjPNR.GeneralRemarkCount
    If gobjPNR.GeneralRemark(intI).Qualifier = "TK" Then
        'NP.TK*FARETYPE-1:SQC
        If InStr(1, gobjPNR.GeneralRemark(intI).RemarkText, "FARETYPE-" & mbytFFNum) > 0 Then
            gobjHost.terminalEntry "NP." & gobjPNR.GeneralRemark(intI).ItemNum & "@"
            Exit For
        End If
  
    End If
    
Next

gobjHost.terminalEntry "NP.TK*FARETYPE-" & mbytFFNum & ":" & Left(cmbFareType.Text, InStr(1, cmbFareType.Text, "-") - 2) & IIf(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).International = False, ":DOM", "")

End Sub

Private Function valError(ByVal strResp As String) As Boolean
    
    Dim rsRec As ADODB.Recordset
    Dim strSQL As String
    
    valError = True
    strResp = Replace(strResp, vbCrLf, "")
    strSQL = "Select ErrMsg from tblFFError"
    Set rsRec = gdbConn.Execute(strSQL)
    Do While Not rsRec.EOF
        If InStr(1, UCase(strResp), UCase(rsRec!ErrMsg)) > 0 Then
           valError = False
           Exit Do
        End If
        rsRec.MoveNext
    Loop
    rsRec.Close
    
End Function

Function WriteToLog()

    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFileFare, _
    "frmPricingWiz1", gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFileFare, _
    "frmPricingWiz1", gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFileFare, _
    "frmPricingWiz1", gconProcessing, gstrProcessGrpID, , datTouchEnd

End Function


