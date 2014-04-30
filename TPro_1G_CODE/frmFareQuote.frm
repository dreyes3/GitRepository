VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmFareQuote 
   Caption         =   "CWT TravelPro - Fare Quote"
   ClientHeight    =   7800
   ClientLeft      =   105
   ClientTop       =   2190
   ClientWidth     =   9285
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFareQuote.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   9285
   Begin VB.CommandButton cmdPre 
      Caption         =   "<"
      Height          =   375
      Left            =   5400
      TabIndex        =   115
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   6000
      TabIndex        =   114
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPx 
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   113
      Text            =   "cmbPx"
      Top             =   480
      Width           =   5055
   End
   Begin VB.CheckBox chkAddRI 
      Caption         =   "Add to RI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   112
      Top             =   480
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   11139
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   4
      TabHeight       =   626
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "New Fare Quote"
      TabPicture(0)   =   "frmFareQuote.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMerchFeeAmount"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblMsg"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblClientType(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblHiddenComm"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTotalQuoteText"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTotalQuote"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDiscountAmtText"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDiscountAmt"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblCommissionAmtText"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCommissionAmt"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblNetBaseFareText"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblNetBaseFare"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblTax"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblGrossBaseFare"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblTransFee"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDiscount"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblCommission"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblClientType(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblFuelSurcharge"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblTktType"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblAddComm"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkTFOverride"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtTrxFee"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkItinType"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkFareType"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame3"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtHiddenComm"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtMerchantFee"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtDiscount"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtCommission"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "chkMerchantFee"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdRecalculate"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtTax"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtBaseFare"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cboClientType"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "chkNRCC(0)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtFuelSurcharge"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "Restrictions/Notes"
      TabPicture(1)   =   "frmFareQuote.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRmkAdd(0)"
      Tab(1).Control(1)=   "cmdRmkAdd(1)"
      Tab(1).Control(2)=   "cmdRmkAdd(2)"
      Tab(1).Control(3)=   "cmdRmkAdd(3)"
      Tab(1).Control(4)=   "cmdRmkAdd(4)"
      Tab(1).Control(5)=   "cmdRmkAdd(5)"
      Tab(1).Control(6)=   "cmdRmkAdd(6)"
      Tab(1).Control(7)=   "cmdRmkAdd(8)"
      Tab(1).Control(8)=   "cboOthRmks"
      Tab(1).Control(9)=   "lstRmks"
      Tab(1).Control(10)=   "txtFreeText"
      Tab(1).Control(11)=   "cmdRmkAdd(9)"
      Tab(1).Control(12)=   "cmdRmkAdd(7)"
      Tab(1).Control(13)=   "lblRmkText(0)"
      Tab(1).Control(14)=   "lblRmkText(1)"
      Tab(1).Control(15)=   "lblRmkText(2)"
      Tab(1).Control(16)=   "lblRmkText(3)"
      Tab(1).Control(17)=   "lblRmkText(4)"
      Tab(1).Control(18)=   "lblRmkText(5)"
      Tab(1).Control(19)=   "lblRmkText(6)"
      Tab(1).Control(20)=   "lblRmkText(7)"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "FQ Rules Text"
      TabPicture(2)   =   "frmFareQuote.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtRules"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Remove RI"
      TabPicture(3)   =   "frmFareQuote.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lswRI"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Best Fare Option"
      TabPicture(4)   =   "frmFareQuote.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblLFLabel(0)"
      Tab(4).Control(1)=   "lblLFLabel(1)"
      Tab(4).Control(2)=   "lblLFTotalQuoteText"
      Tab(4).Control(3)=   "lblLFLabel(11)"
      Tab(4).Control(4)=   "lblLFDiscountAmtText"
      Tab(4).Control(5)=   "lblLFLabel(9)"
      Tab(4).Control(6)=   "lblLFCommissionAmtText"
      Tab(4).Control(7)=   "lblLFLabel(8)"
      Tab(4).Control(8)=   "lblLFNetBaseFareText"
      Tab(4).Control(9)=   "lblLFLabel(7)"
      Tab(4).Control(10)=   "lblLFLabel(6)"
      Tab(4).Control(11)=   "lblLFLabel(2)"
      Tab(4).Control(12)=   "lblLFLabel(5)"
      Tab(4).Control(13)=   "lblLFLabel(4)"
      Tab(4).Control(14)=   "lblLFLabel(3)"
      Tab(4).Control(15)=   "lblLFlMerchFeeAmount"
      Tab(4).Control(16)=   "lblLFLabel(10)"
      Tab(4).Control(17)=   "lblLFFuelSurcharge"
      Tab(4).Control(18)=   "chkLFTFOverride"
      Tab(4).Control(19)=   "txtLFAirline"
      Tab(4).Control(20)=   "chkLFNoLowerFare"
      Tab(4).Control(21)=   "fraLFRouting"
      Tab(4).Control(22)=   "txtLFCWTNet"
      Tab(4).Control(23)=   "txtLFMerchantFee"
      Tab(4).Control(24)=   "txtLFDiscount"
      Tab(4).Control(25)=   "txtLFCommission"
      Tab(4).Control(26)=   "cmdLFRecalculate"
      Tab(4).Control(27)=   "txtLFTax"
      Tab(4).Control(28)=   "txtLFBaseFare"
      Tab(4).Control(29)=   "chkLFItinType"
      Tab(4).Control(30)=   "chkLFFareType"
      Tab(4).Control(31)=   "chkLFMerchantFee"
      Tab(4).Control(32)=   "txtLFTrxFee"
      Tab(4).Control(33)=   "chkNRCC(1)"
      Tab(4).Control(34)=   "txtLFFuelSurcharge"
      Tab(4).ControlCount=   35
      TabCaption(5)   =   "Best Fare Restrictions"
      TabPicture(5)   =   "frmFareQuote.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdLFRmkAdd(0)"
      Tab(5).Control(1)=   "cmdLFRmkAdd(1)"
      Tab(5).Control(2)=   "cmdLFRmkAdd(2)"
      Tab(5).Control(3)=   "cmdLFRmkAdd(3)"
      Tab(5).Control(4)=   "cmdLFRmkAdd(4)"
      Tab(5).Control(5)=   "cmdLFRmkAdd(5)"
      Tab(5).Control(6)=   "cmdLFRmkAdd(6)"
      Tab(5).Control(7)=   "cmdLFRmkAdd(8)"
      Tab(5).Control(8)=   "cboLFOthRmks"
      Tab(5).Control(9)=   "lstLFRmks"
      Tab(5).Control(10)=   "txtLFFreeText"
      Tab(5).Control(11)=   "cmdLFRmkAdd(9)"
      Tab(5).Control(12)=   "cmdLFRmkAdd(7)"
      Tab(5).Control(13)=   "lblLFRmkText(0)"
      Tab(5).Control(14)=   "lblLFRmkText(1)"
      Tab(5).Control(15)=   "lblLFRmkText(2)"
      Tab(5).Control(16)=   "lblLFRmkText(3)"
      Tab(5).Control(17)=   "lblLFRmkText(4)"
      Tab(5).Control(18)=   "lblLFRmkText(5)"
      Tab(5).Control(19)=   "lblLFRmkText(6)"
      Tab(5).Control(20)=   "lblLFRmkText(7)"
      Tab(5).ControlCount=   21
      Begin VB.TextBox txtLFFuelSurcharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -72480
         MaxLength       =   4
         TabIndex        =   121
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtFuelSurcharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2460
         MaxLength       =   4
         TabIndex        =   119
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CheckBox chkNRCC 
         Caption         =   "NRCC"
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
         Left            =   -66960
         TabIndex        =   117
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chkNRCC 
         Caption         =   "NRCC"
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
         Left            =   6120
         TabIndex        =   116
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cboClientType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmFareQuote.frx":0972
         Left            =   2520
         List            =   "frmFareQuote.frx":0974
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtBaseFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2460
         TabIndex        =   58
         Top             =   2580
         Width           =   1095
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2460
         TabIndex        =   57
         Top             =   4020
         Width           =   1095
      End
      Begin VB.CommandButton cmdRecalculate 
         Caption         =   "&Calculate Fare"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   6480
         Picture         =   "frmFareQuote.frx":0976
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   3360
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.CheckBox chkMerchantFee 
         Caption         =   "Include Merchant Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   4380
         Width           =   1935
      End
      Begin VB.TextBox txtCommission 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   54
         Top             =   2940
         Width           =   492
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   53
         Top             =   3300
         Width           =   492
      End
      Begin VB.TextBox txtMerchantFee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   52
         Top             =   4380
         Width           =   495
      End
      Begin VB.TextBox txtHiddenComm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2460
         TabIndex        =   51
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Routing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   49
         Top             =   5520
         Width           =   8595
         Begin VB.TextBox txtRouting 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            MaxLength       =   63
            TabIndex        =   50
            Top             =   300
            Width           =   8115
         End
      End
      Begin VB.CheckBox chkFareType 
         Caption         =   "Nett"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1200
         Width           =   1275
      End
      Begin VB.CheckBox chkItinType 
         Caption         =   "International"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1680
         Width           =   1275
      End
      Begin VB.TextBox txtTrxFee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   2460
         TabIndex        =   46
         Top             =   3660
         Width           =   1095
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Left            =   -74520
         TabIndex        =   45
         Top             =   1380
         Width           =   1155
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Left            =   -74520
         TabIndex        =   44
         Top             =   1680
         Width           =   1155
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Index           =   2
         Left            =   -74520
         TabIndex        =   43
         Top             =   1980
         Width           =   1155
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Index           =   3
         Left            =   -74520
         TabIndex        =   42
         Top             =   2280
         Width           =   1155
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Index           =   4
         Left            =   -74520
         TabIndex        =   41
         Top             =   2580
         Width           =   1155
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Index           =   5
         Left            =   -74520
         TabIndex        =   40
         Top             =   2880
         Width           =   1155
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Index           =   6
         Left            =   -74520
         TabIndex        =   39
         Top             =   3180
         Width           =   1155
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Index           =   8
         Left            =   -74520
         TabIndex        =   38
         Top             =   3780
         Width           =   1155
      End
      Begin VB.ComboBox cboOthRmks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73320
         TabIndex        =   37
         Top             =   3780
         Width           =   7095
      End
      Begin VB.ListBox lstRmks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   -74640
         TabIndex        =   36
         Top             =   4680
         Width           =   8295
      End
      Begin VB.TextBox txtFreeText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   35
         Top             =   4260
         Width           =   6795
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Index           =   9
         Left            =   -74520
         TabIndex        =   34
         Top             =   4260
         Width           =   1155
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Left            =   -74520
         TabIndex        =   33
         Top             =   1380
         Width           =   1155
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Left            =   -74520
         TabIndex        =   32
         Top             =   1680
         Width           =   1155
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Index           =   2
         Left            =   -74520
         TabIndex        =   31
         Top             =   1980
         Width           =   1155
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Index           =   3
         Left            =   -74520
         TabIndex        =   30
         Top             =   2280
         Width           =   1155
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Index           =   4
         Left            =   -74520
         TabIndex        =   29
         Top             =   2580
         Width           =   1155
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Index           =   5
         Left            =   -74520
         TabIndex        =   28
         Top             =   2880
         Width           =   1155
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Index           =   6
         Left            =   -74520
         TabIndex        =   27
         Top             =   3180
         Width           =   1155
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Index           =   8
         Left            =   -74520
         TabIndex        =   26
         Top             =   3780
         Width           =   1155
      End
      Begin VB.ComboBox cboLFOthRmks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73320
         TabIndex        =   25
         Top             =   3780
         Width           =   7095
      End
      Begin VB.ListBox lstLFRmks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   -74640
         TabIndex        =   24
         Top             =   4620
         Width           =   8295
      End
      Begin VB.TextBox txtLFFreeText 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   23
         Top             =   4260
         Width           =   6795
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Index           =   9
         Left            =   -74520
         TabIndex        =   22
         Top             =   4260
         Width           =   1155
      End
      Begin VB.TextBox txtLFTrxFee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   -72480
         TabIndex        =   21
         Top             =   3660
         Width           =   1095
      End
      Begin VB.CheckBox chkLFMerchantFee 
         Caption         =   "No Merchant Fee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4380
         Width           =   1935
      End
      Begin VB.CheckBox chkLFFareType 
         Caption         =   "Nett"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1560
         Width           =   1275
      End
      Begin VB.CheckBox chkLFItinType 
         Caption         =   "International"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2040
         Width           =   1275
      End
      Begin VB.TextBox txtLFBaseFare 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -72480
         TabIndex        =   17
         Top             =   2580
         Width           =   1095
      End
      Begin VB.TextBox txtLFTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -72480
         TabIndex        =   16
         Top             =   4020
         Width           =   1095
      End
      Begin VB.CommandButton cmdLFRecalculate 
         Caption         =   "&Calculate Fare"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   -68580
         Picture         =   "frmFareQuote.frx":0C80
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtLFCommission 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -71880
         MaxLength       =   4
         TabIndex        =   14
         Top             =   2940
         Width           =   492
      End
      Begin VB.TextBox txtLFDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -71880
         MaxLength       =   4
         TabIndex        =   13
         Top             =   3300
         Width           =   492
      End
      Begin VB.TextBox txtLFMerchantFee 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -71880
         MaxLength       =   4
         TabIndex        =   12
         Top             =   4380
         Width           =   495
      End
      Begin VB.TextBox txtLFCWTNet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -72480
         TabIndex        =   11
         Top             =   2220
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame fraLFRouting 
         Caption         =   "Routing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74820
         TabIndex        =   9
         Top             =   5520
         Width           =   8595
         Begin VB.TextBox txtLFRouting 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            MaxLength       =   63
            TabIndex        =   10
            Top             =   300
            Width           =   8115
         End
      End
      Begin VB.CheckBox chkLFNoLowerFare 
         Caption         =   "Tick here if current quote is 'Best Fare' (no lower fare available)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   8
         Top             =   1080
         Width           =   6375
      End
      Begin VB.TextBox txtLFAirline 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -72480
         TabIndex        =   7
         Top             =   1800
         Width           =   3075
      End
      Begin VB.TextBox txtRules 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74220
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   960
         Width           =   7515
      End
      Begin VB.CheckBox chkTFOverride 
         Caption         =   "Override TF"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   5
         Top             =   3660
         Width           =   1395
      End
      Begin VB.CheckBox chkLFTFOverride 
         Caption         =   "Override TF"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   4
         Top             =   3660
         Width           =   1395
      End
      Begin VB.CommandButton cmdLFRmkAdd 
         Caption         =   "Add"
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
         Index           =   7
         Left            =   -74520
         TabIndex        =   3
         Top             =   3480
         Width           =   1155
      End
      Begin VB.CommandButton cmdRmkAdd 
         Caption         =   "Add"
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
         Index           =   7
         Left            =   -74520
         TabIndex        =   2
         Top             =   3480
         Width           =   1155
      End
      Begin MSComctlLib.ListView lswRI 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   118
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8916
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LN"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Text"
            Object.Width           =   13229
         EndProperty
      End
      Begin VB.Label lblAddComm 
         Caption         =   "lblAddComm"
         Height          =   315
         Left            =   3780
         TabIndex        =   129
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ticket Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   124
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label lblTktType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4905
         TabIndex        =   123
         Top             =   2220
         Width           =   2415
      End
      Begin VB.Label lblLFFuelSurcharge 
         Alignment       =   1  'Right Justify
         Caption         =   "Fuel Surcharge:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74040
         TabIndex        =   122
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label lblFuelSurcharge 
         Alignment       =   1  'Right Justify
         Caption         =   "Fuel Surcharge:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   120
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label lblClientType 
         Alignment       =   1  'Right Justify
         Caption         =   "Pricing Schema:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   111
         Top             =   1140
         Width           =   1815
      End
      Begin VB.Label lblCommission 
         Alignment       =   1  'Right Justify
         Caption         =   "Commission/Mark-up:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   110
         Top             =   2940
         Width           =   1935
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount/Rebate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   109
         Top             =   3300
         Width           =   1935
      End
      Begin VB.Label lblTransFee 
         Alignment       =   1  'Right Justify
         Caption         =   "Transaction Fee:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   108
         Top             =   3660
         Width           =   1335
      End
      Begin VB.Label lblGrossBaseFare 
         Alignment       =   1  'Right Justify
         Caption         =   "Base Fare:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   107
         Top             =   2580
         Width           =   1335
      End
      Begin VB.Label lblTax 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax(es):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   106
         Top             =   4020
         Width           =   1335
      End
      Begin VB.Label lblNetBaseFare 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   105
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label lblNetBaseFareText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "net base fare"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4800
         TabIndex        =   104
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label lblCommissionAmt 
         Alignment       =   1  'Right Justify
         Caption         =   "%      Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   103
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label lblCommissionAmtText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "comm amt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4800
         TabIndex        =   102
         Top             =   2940
         Width           =   1215
      End
      Begin VB.Label lblDiscountAmt 
         Alignment       =   1  'Right Justify
         Caption         =   "%      Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   101
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Label lblDiscountAmtText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "disc amt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4800
         TabIndex        =   100
         Top             =   3300
         Width           =   1215
      End
      Begin VB.Label lblTotalQuote 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Fare Quote:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   99
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Label lblTotalQuoteText 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "total quote"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4440
         TabIndex        =   98
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label lblHiddenComm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hidden Commission:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   930
         TabIndex        =   97
         Top             =   2220
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "%    (Amount):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   96
         Top             =   4380
         Width           =   1095
      End
      Begin VB.Label lblClientType 
         Alignment       =   1  'Right Justify
         Caption         =   "(Client Type)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   95
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THIS IS A SPECIAL FARE * RESTRICTIONS APPLY"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -73320
         TabIndex        =   94
         Top             =   1380
         Width           =   6855
      End
      Begin VB.Label lblRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FARE VALID ONLY ON [Airline]"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -73320
         TabIndex        =   93
         Top             =   1680
         Width           =   6855
      End
      Begin VB.Label lblRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIN STAY: [MIN STAY] * MAX STAY: [MAX STAY]"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -73320
         TabIndex        =   92
         Top             =   1980
         Width           =   6855
      End
      Begin VB.Label lblRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VALID ONLY ON FLIGHTS/DATES SHOWN"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -73320
         TabIndex        =   91
         Top             =   2280
         Width           =   6855
      End
      Begin VB.Label lblRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PENALTY WILL APPLY FOR ANY CHANGES MADE TO TICKETED FLIGHTS"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -73320
         TabIndex        =   90
         Top             =   2580
         Width           =   6855
      End
      Begin VB.Label lblRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "REFUNDABLE THROUGH ISSUING OFFICE ONLY"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -73320
         TabIndex        =   89
         Top             =   2880
         Width           =   6855
      End
      Begin VB.Label lblRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NON REFUNDABLE * NO REFUND FOR UNUSED TICKET"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -73320
         TabIndex        =   88
         Top             =   3180
         Width           =   6855
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         Caption         =   "MSG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   900
         TabIndex        =   87
         Top             =   1860
         Width           =   2340
      End
      Begin VB.Label lblMerchFeeAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "disc amt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4800
         TabIndex        =   86
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblLFRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THIS IS A SPECIAL FARE * RESTRICTIONS APPLY"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -73320
         TabIndex        =   85
         Top             =   1380
         Width           =   6855
      End
      Begin VB.Label lblLFRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FARE VALID ONLY ON [Airline]"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -73320
         TabIndex        =   84
         Top             =   1680
         Width           =   6855
      End
      Begin VB.Label lblLFRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MIN STAY: [MIN STAY] * MAX STAY: [MAX STAY]"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -73320
         TabIndex        =   83
         Top             =   1980
         Width           =   6855
      End
      Begin VB.Label lblLFRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VALID ONLY ON FLIGHTS/DATES SHOWN"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -73320
         TabIndex        =   82
         Top             =   2280
         Width           =   6855
      End
      Begin VB.Label lblLFRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PENALTY WILL APPLY FOR ANY CHANGES MADE TO TICKETED FLIGHTS"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -73320
         TabIndex        =   81
         Top             =   2580
         Width           =   6855
      End
      Begin VB.Label lblLFRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "REFUNDABLE THROUGH ISSUING OFFICE ONLY"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -73320
         TabIndex        =   80
         Top             =   2880
         Width           =   6855
      End
      Begin VB.Label lblLFRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NON REFUNDABLE * NO REFUND FOR UNUSED TICKET"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -73320
         TabIndex        =   79
         Top             =   3180
         Width           =   6855
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "%    (Amount):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -71280
         TabIndex        =   78
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label lblLFlMerchFeeAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "disc amt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -70140
         TabIndex        =   77
         Top             =   4380
         Width           =   1215
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Commission/Mark-up:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -73620
         TabIndex        =   76
         Top             =   2940
         Width           =   1635
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount/Rebate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -73620
         TabIndex        =   75
         Top             =   3300
         Width           =   1635
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Transaction Fee:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74160
         TabIndex        =   74
         Top             =   3660
         Width           =   1455
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Base Fare:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74160
         TabIndex        =   73
         Top             =   2580
         Width           =   1455
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax(es):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -74160
         TabIndex        =   72
         Top             =   4020
         Width           =   1455
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -71340
         TabIndex        =   71
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label lblLFNetBaseFareText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "net base fare"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -70140
         TabIndex        =   70
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "%      Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -71340
         TabIndex        =   69
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label lblLFCommissionAmtText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "comm amt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -70140
         TabIndex        =   68
         Top             =   2940
         Width           =   1215
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "%      Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -71340
         TabIndex        =   67
         Top             =   3300
         Width           =   1095
      End
      Begin VB.Label lblLFDiscountAmtText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "disc amt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -70140
         TabIndex        =   66
         Top             =   3300
         Width           =   1215
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Best Fare Quote:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -73800
         TabIndex        =   65
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Label lblLFTotalQuoteText 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "total quote"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -70500
         TabIndex        =   64
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "CWT Net:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74160
         TabIndex        =   63
         Top             =   2220
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblLFLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Airline(s) for lower fare:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74460
         TabIndex        =   62
         Top             =   1800
         Width           =   1875
      End
      Begin VB.Label lblLFRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NON ENDORSABLE NON REROUTABLE"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -73320
         TabIndex        =   61
         Top             =   3480
         Width           =   6855
      End
      Begin VB.Label lblRmkText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NON ENDORSABLE NON REROUTABLE"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -73320
         TabIndex        =   60
         Top             =   3480
         Width           =   6855
      End
   End
   Begin MyCommandButton.MyButton cmdPrevious 
      Height          =   360
      Left            =   1800
      TabIndex        =   125
      Top             =   7320
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
   Begin MyCommandButton.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   6480
      TabIndex        =   127
      Top             =   7320
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
   Begin MyCommandButton.MyButton cmdSave 
      Height          =   360
      Left            =   3480
      TabIndex        =   128
      Top             =   7320
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
   Begin MyCommandButton.MyButton cmdSaveAdd 
      Height          =   360
      Left            =   4680
      TabIndex        =   126
      Top             =   7320
      Width           =   1605
      _ExtentX        =   2831
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
      Caption         =   "Finish && Add &Another"
      Depth           =   1
      GradientType    =   2
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmFareQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msngMarkup As Single
Dim msngIntlDiscount As Single
Dim msngDomDiscount As Single
Dim msngIntlComm As Single
Dim msngDOMComm As Single
Dim msngMerchFee As Single
Dim msngMerchFeeAmt As Single
Dim msngNetBaseFare As Single
Dim msngCommAmt As Single
Dim msngDiscAmt As Single
Dim msngTotalCharge As Single
Dim msngSellFare As Single
Dim msngCommPct As Single
Dim msngDiscPct As Single
Dim msngBaseFare As Single
Dim msngTransFee As Single
Dim msngTax As Single
Dim msngLFBaseFare As Single
Dim msngLFMerchFee As Single
Dim msngLFCommPct As Single
Dim msngLFDiscPct As Single
Dim msngLFTransFee As Single
Dim msngLFTax As Single
Dim msngLFCommAmt As Single
Dim msngLFNetBaseFare As Single
Dim msngLFSellFare As Single
Dim msngLFDiscAmt As Single
Dim msngLFTotalCharge As Single
Dim msngLFMerchFeeAmt As Single
Dim mblnLoaded As Boolean
Dim mblnLFLoaded As Boolean
Dim mblnFQExist As Boolean
Dim mblnLFExist As Boolean
Dim msgMFTotalCharge As Single
Dim msgLFMFTotalCharge As Single
Dim mblnAmend As Boolean
Dim msngHiddenComm As Single

'Timer
Dim startTime As Date
'230108
Dim msngFuelSurcharge As Single
Dim msngLFFuelSurcharge As Single

Dim datFormLoadEnd As Date
Dim datFormLoadStart As Date
Dim datTouchEnd As Date

Private Sub ClearVar()
  msngMarkup = 0
  msngIntlDiscount = 0
  msngDomDiscount = 0
  msngIntlComm = 0
  msngDOMComm = 0
  msngMerchFee = 0
  msngMerchFeeAmt = 0
  msngNetBaseFare = 0
  msngCommAmt = 0
  msngDiscAmt = 0
  msngTotalCharge = 0
  msngSellFare = 0
  msngCommPct = 0
  msngDiscPct = 0
  msngBaseFare = 0
  msngTransFee = 0
  msngTax = 0
  msngLFBaseFare = 0
  msngLFMerchFee = 0
  msngLFCommPct = 0
  msngLFDiscPct = 0
  msngLFTransFee = 0
  msngLFTax = 0
  msngLFCommAmt = 0
  msngLFNetBaseFare = 0
  msngLFSellFare = 0
  msngLFDiscAmt = 0
  msngLFTotalCharge = 0
  msngLFMerchFeeAmt = 0
  
    '230108
  msngFuelSurcharge = 0
  msngLFFuelSurcharge = 0
  'mblnLoaded As Boolean
  'mblnLFLoaded As Boolean
  'mblnFQExist As Boolean
  'mblnLFExist As Boolean

    '111108
  msngHiddenComm = 0
  
End Sub


Private Sub cboClientType_Click()
gobjLog.ProcedureName = "cboClientType_Click"
On Error GoTo ProcError

If cboClientType = "TP" Or cboClientType = "TF" Then
    txtTrxFee.Enabled = True
    chkTFOverride.Enabled = True
    txtLFTrxFee.Enabled = True
    chkLFTFOverride.Enabled = True
Else
    txtTrxFee.Enabled = False
    chkTFOverride.Enabled = False
    txtLFTrxFee.Enabled = False
    chkLFTFOverride.Enabled = False
End If
If cboClientType = "MN" Then
   txtCommission = 0
   txtDiscount = 0
End If

If mblnLoaded Then Call NumbersChanged
If mblnLFLoaded Then Call LFNumbersChanged
Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub





Private Sub chkLFFareType_Click()
gobjLog.ProcedureName = "chkLFFareType_Click"
On Error GoTo ProcError

If chkLFFareType.value = vbChecked Then
    chkLFFareType.Caption = "Published"
    txtLFCommission = msngIntlComm
    lblLFLabel(2).Caption = "Published Base:"
    txtLFCWTNet.Visible = False
    lblLFLabel(1).Visible = False
    lblLFLabel(3).Caption = "Commission (%):"
    lblLFLabel(8).Caption = "Commission Amount:"
    txtLFCommission = msngIntlComm
    chkNRCC(1).Visible = False
    chkNRCC(1).value = 0
Else
    chkLFFareType.Caption = "Nett"
    'gbolNetFare = True
    txtLFCommission = msngMarkup
    lblLFLabel(2).Caption = "Market Net:"
    txtLFCWTNet.Visible = False
    lblLFLabel(1).Visible = False
    lblLFLabel(3).Caption = "Mark Up (%):"
    lblLFLabel(8).Caption = "Mark Up Amount:"
    
    'modified on 27/12: This checking should be done separately for Low Fare and Normal Fare
    If (cboClientType = "MN" Or cboClientType = "TF" Or cboClientType = "TP") Then
        msngMarkup = 0
        txtLFCommission = "0"
    End If
    
    'added on 14/3/2005: add NRCC checkbox for HK
    'If UCase(gstrAgcyCountryCode) = "HK" Then
        chkNRCC(1).Visible = True
    'Else
    '    chkNRCC(1).Visible = False
    'End If
    
End If

If mblnLoaded Then Call LFNumbersChanged

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub chkLFNoLowerFare_Click()
gobjLog.ProcedureName = "chkLFNoLowerFare_Click"
On Error GoTo ProcError

Dim lngC As Long
Dim blnValue As Boolean

blnValue = chkLFNoLowerFare.value

txtLFAirline.Enabled = Not blnValue
txtLFCWTNet.Enabled = Not blnValue
txtLFBaseFare.Enabled = Not blnValue
txtLFCommission.Enabled = Not blnValue
txtLFDiscount.Enabled = Not blnValue
txtLFTrxFee.Enabled = Not blnValue
txtLFTax.Enabled = Not blnValue
chkLFMerchantFee.Enabled = Not blnValue
chkLFFareType.Enabled = Not blnValue
chkLFItinType.Enabled = Not blnValue
chkNRCC(1).Enabled = Not blnValue
cmdLFRecalculate.Visible = Not blnValue
lblLFNetBaseFareText.Enabled = Not blnValue
lblLFCommissionAmtText.Enabled = Not blnValue
lblLFDiscountAmtText.Enabled = Not blnValue
lblLFlMerchFeeAmount.Enabled = Not blnValue
lblLFTotalQuoteText.Enabled = Not blnValue
lblLFFuelSurcharge.Enabled = Not blnValue
txtLFFuelSurcharge.Enabled = Not blnValue

For lngC = 0 To 11
    lblLFLabel(lngC).Enabled = Not blnValue
Next

If mblnFQExist Then
    cmdSave.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdSave.Default = True
    'cmdSave.SetFocus
End If

mblnLFExist = True

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub chkLFTFOverride_Click()
If mblnLFLoaded Then Call LFNumbersChanged
End Sub
Private Sub chkNRCC_Click(Index As Integer)
    Select Case Index
        Case 0:
            If mblnLoaded Then Call NumbersChanged
        Case 1:
            If mblnLoaded Then Call LFNumbersChanged
    End Select
             
    'Added on 10 Nov 2008 by Jeremy.
    'If un-checked then add HiddenComm to get new fare, if checked then subtract HiddenComm and rollback to original fare
    If chkNRCC(0).value = vbChecked Then
        txtHiddenComm.Text = 0
    Else
        txtHiddenComm.Text = msngHiddenComm
    End If
 
End Sub

Private Sub chkTFOverride_Click()
If mblnLoaded Then Call NumbersChanged
End Sub

Private Sub cmbPx_Click()

If mblnAmend = False Then GetALComm
PopulateControls
End Sub



Private Sub cmdLFRmkAdd_Click(Index As Integer)
gobjLog.ProcedureName = "cmdAdd_Click"
On Error GoTo ProcError

Dim strTemp As String

    Select Case Index
        Case 0 To 7
            strTemp = lblLFRmkText(Index).Caption
            If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
                Load frmFareRmkFill
                With frmFareRmkFill
                    .lblRmkText = strTemp
                    .Show
                    .FormatRemark
                    Do
                      DoEvents
                    Loop Until frmFareRmkFill.Visible = False
                    strTemp = .lblRmkText.Caption
                    If strTemp = "" Then
                        'CANCEL
                    Else
                        lstLFRmks.AddItem strTemp
                    End If
                    Unload frmFareRmkFill
                End With
            Else
                lstLFRmks.AddItem strTemp
            End If
            Set frmFareRmkFill = Nothing
            
        Case 8 'remarks from list
            strTemp = cboLFOthRmks.Text
            If InStr(1, strTemp, "[") And InStr(1, strTemp, "]") Then
                Load frmFareRmkFill
                With frmFareRmkFill
                    .lblRmkText = strTemp
                    .Show
                    .FormatRemark
                    Do
                      DoEvents
                    Loop Until frmFareRmkFill.Visible = False
                    strTemp = .lblRmkText.Caption
                    If strTemp = "" Then
                        'CANCEL
                    Else
                        lstLFRmks.AddItem strTemp
                    End If
                    Unload frmFareRmkFill
                End With
            Else
                lstLFRmks.AddItem strTemp
            End If
            Set frmFareRmkFill = Nothing
                        
            
        Case 9 'free text
            If txtLFFreeText.Text <> "" Then lstLFRmks.AddItem txtLFFreeText.Text
            txtLFFreeText.Text = ""
    End Select
Exit Sub
ProcError:
    Call pErrorReport(True)
    
End Sub

Private Sub cmdNext_Click()
   If cmbPx.listindex < cmbPx.ListCount - 1 Then
      cmbPx.listindex = cmbPx.listindex + 1
   End If
End Sub

Private Sub cmdPre_Click()
   If cmbPx.listindex > 0 Then
      cmbPx.listindex = cmbPx.listindex - 1
   End If
End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload frmFareQuoteRequest
    Set frmFareQuoteRequest = Nothing
    'Dim frm As Form
    'For Each frm In Forms
    '    MsgBox frm.Name
    'Next
    Unload Me
End Sub

Private Sub cmdRmkAdd_Click(Index As Integer)
gobjLog.ProcedureName = "cmdAdd_Click"
On Error GoTo ProcError

Dim strTemp As String

    Select Case Index
        Case 0, 1, 2, 3, 4, 5, 6, 7
            If Index <> 9 Then
               strTemp = lblRmkText(Index).Caption
            Else
               strTemp = txtFreeText
            End If
            'strTemp = IIf(Index <> 9, lblRmkText(Index).Caption, txtFreeText)
            If Trim(strTemp) = "" Then Exit Sub
            If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
                Load frmFareRmkFill
                With frmFareRmkFill
                    .lblRmkText = strTemp
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
            
        Case 8 'remarks from list
            strTemp = cboOthRmks.Text
            If InStr(1, strTemp, "[") > 0 And InStr(1, strTemp, "]") > 0 Then
                Load frmFareRmkFill
                With frmFareRmkFill
                    .lblRmkText = strTemp
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

        Case 9 'free text
            If txtFreeText.Text <> "" Then lstRmks.AddItem txtFreeText.Text
            txtFreeText.Text = ""
            

    End Select
Exit Sub
ProcError:
    Call pErrorReport(True)
    

End Sub

Private Sub cmdCancel_Click()
gobjLog.ProcedureName = "cmdCancel_Click"
On Error GoTo ProcError

On Error GoTo ProcError

If fWantToQuit Then
    gbolCancelProcess = True
    Unload Me
    'Call pRedisplayMenu
End If
Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub cmdLFRecalculate_Click()
gobjLog.ProcedureName = "cmdRecalculate_Click"
On Error GoTo ProcError

Dim strMsg As String

If cboClientType.Text = "TF" And (Trim(txtLFTrxFee.Text) = "" Or txtLFTrxFee.Text = "0") And chkLFTFOverride.value = vbChecked Then strMsg = "Need Transaction Fee amount..." & vbCrLf
If chkLFMerchantFee.value = vbChecked And Trim(txtLFMerchantFee) = "" Then strMsg = strMsg & "Need Merchant Fee percentage..." & vbCrLf
'If chkLFMerchantFee.Value = vbChecked And chkLFFareType.Value = vbChecked Then
'    chkLFMerchantFee.Value = vbUnchecked
'    txtLFMerchantFee.Text = ""
'End If
    
If strMsg <> "" Then
    'MsgBox strMsg, vbApplicationModal + vbCritical + vbOKOnly, "MISSING OR INVALID DATA!"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    Exit Sub
End If

mblnLoaded = True
cmdRecalculate.Caption = "&Recalculate Fare"
Call CalculateLowFare

cmdLFRecalculate.Visible = False

If mblnFQExist Then
    cmdSave.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdSave.Default = True
    cmdSave.SetFocus
End If

mblnLFExist = True

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub cmdSaveAdd_Click()
' ########## need to add code to loop back to frmFareQuoteRequest
'Added on 28/10/04
Dim strRes As String
Dim intI As Integer
Dim strMsg As String
gobjLog.ProcedureName = "cmdSaveAdd_Click"
On Error GoTo ProcError
datTouchEnd = Now

 If cmdRecalculate.Visible Then
    SSTab1.Tab = 0
    'MsgBox "Need to recalculate Fare Quote"
    strMsg = "Need to recalculate Fare Quote"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Exit Sub
 End If
 
 If cmdLFRecalculate.Visible Then
    SSTab1.Tab = 4
    'MsgBox "Need to recalculate Low Fare Quote"
    strMsg = "Need to recalculate Low Fare Quote"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Exit Sub
 End If

 
If lstRmks.ListCount = 0 Then
    strMsg = "Do you want to add remarks for the fare quote?"
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop - Add Remarks") = vbYes Then
    'If MsgBox("Do you want to add remarks for the fare quote?", vbApplicationModal + vbQuestion + vbYesNo, "Fare quote") = vbYes Then
        SSTab1.Tab = 1
        Exit Sub
    End If
 End If
 
 If lstLFRmks.ListCount = 0 And chkLFNoLowerFare = vbUnchecked Then
    strMsg = "Do you want to add remarks for the low fare option?"
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop - Add Remarks") = vbYes Then
    'If MsgBox("Do you want to add remarks for the low fare option?", vbApplicationModal + vbQuestion + vbYesNo, "Fare quote") = vbYes Then
        SSTab1.Tab = 4
        Exit Sub
    End If
 End If
 

Me.Hide
frmWait.Show
Call WriteQuoteToGDS
'Timer
Call pAddToVBILog(gobjPNR.RecLoc, "Fare Quote", startTime)

With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
    'Added on 19/08/04: reassign in case the user adjust the base amount
    .BaseAmount = CSng(txtBaseFare.Text)
    '
    .MerchAmt = msngMerchFeeAmt
    .SellAmount = msngSellFare
    .NetAmount = IIf(gbolNetFare = True, msngBaseFare, 0)
    .TransactionFee = msngTransFee
     '230108
    .FuelSurcharge = msngFuelSurcharge
    '111108
    .HiddenComm = CSng(txtHiddenComm.Text)
    'If Me.chkLFNoLowerFare = vbUnchecked Then
    '    .LowFare = msngLFSellFare + .TaxTotal
    'Else
    '    .LowFare = .BaseAmount + .TaxTotal
    'End If
    'Modified on 26/07/04 to include discount amount
    .DiscountAmt = msngDiscAmt
    .TaxTotal = msngTax
    
    'Added on 15/03/05 to include NRCC Checkbox
    'If UCase(gstrAgcyCountryCode) = "HK" Then
        .NRCC = chkNRCC(0).value
    'End If
    If Me.chkLFNoLowerFare = vbUnchecked Then
        .LowFare = msngLFSellFare + msngLFTax - msngLFDiscAmt
    Else
        .LowFare = msngSellFare + msngTax - msngDiscAmt
    End If
    gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(1).CommissionOnTicket = txtCommission
 
    'added on 14/10/04
    'Call addObjToTables
    If cmbPx.listindex <> cmbPx.ListCount - 1 And Not .StoreFare Then
       clearControls
       ClearVar
       chkAddRI.value = vbChecked
       cmbPx.listindex = cmbPx.listindex + 1
       mblnLFLoaded = False
       mblnLFExist = False
       SSTab1.Tab = 0
       Me.Show
       frmWait.Hide
       Exit Sub
    End If
    
    'added on 6/1/05
    If cmbPx.listindex = cmbPx.ListCount - 1 And Not .StoreFare Then
        Call addObjToTables
    End If
    
    'Write to SQL Log
    WriteToLog
    
    If .StoreFare Then
        frmPricingWiz1.cmbPx.Clear
        For intI = 0 To cmbPx.ListCount - 1
           frmPricingWiz1.cmbPx.AddItem cmbPx.List(intI)
        Next
        frmPricingWiz1.cmbPx.listindex = 0
        
        Load frmPricingWiz1
        Unload Me
        Unload frmWait
        frmPricingWiz1.Show
    Else
        Load frmFareQuoteRequest
        Unload Me
        Unload frmWait
        frmFareQuoteRequest.Show
    End If
    
End With


Exit Sub
ProcError:
    'Write to SQL Log
    WriteToLog
    Call pErrorReport(True)


'
End Sub

Private Sub chkItinType_Click()
gobjLog.ProcedureName = "chkItinType_Click"


If chkItinType.value = vbChecked Then
    chkItinType.Caption = "International"
   If cboClientType.Text = "DU" Or cboClientType.Text = "MG" Then txtDiscount = msngIntlDiscount
Else
    chkItinType.Caption = "Domestic"
    If cboClientType.Text = "DU" Or cboClientType.Text = "MG" Then txtDiscount = msngDomDiscount
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFareQuote = Nothing
End Sub



Private Sub lstLFRmks_DblClick()
txtLFFreeText.Text = lstLFRmks.Text
lstLFRmks.RemoveItem lstLFRmks.listindex

End Sub

Private Sub lstRmks_DblClick()
txtFreeText.Text = lstRmks.Text
lstRmks.RemoveItem lstRmks.listindex
End Sub





Private Sub lswRI_Click()
Dim i As Integer
Dim j As Integer
For i = 1 To lswRI.ListItems.Count
    If lswRI.ListItems(i).Selected = True Then
        If InStr(lswRI.ListItems(i).SubItems(1), "FARE QUOTE") > 0 Then
            For j = i + 1 To lswRI.ListItems.Count
            If InStr(lswRI.ListItems(j).SubItems(1), "FARE QUOTE") > 0 Then
            Exit For
                
            Else
            lswRI.ListItems(j).Selected = True
                
            End If
            Next j
            
            Exit For
        End If
    End If
Next

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
    Case 0
        If cmdRecalculate.Visible Then cmdRecalculate.Default = True
    Case 4
        If Not mblnLFLoaded Then
            txtLFBaseFare.Text = ""
            txtLFTax.Text = txtTax.Text
            txtLFCommission.Text = txtCommission.Text
            txtLFDiscount.Text = txtDiscount.Text
            chkLFMerchantFee.value = chkMerchantFee.value
            txtLFMerchantFee.Text = txtMerchantFee.Text
            cmdLFRecalculate.Visible = True
            mblnLFLoaded = True
            lblLFlMerchFeeAmount.Caption = " "
            lblLFCommissionAmtText = " "
            lblLFDiscountAmtText = " "
            lblLFNetBaseFareText = " "
            lblLFTotalQuoteText = " "
            txtLFRouting.Text = txtRouting.Text
            
        End If
        If cmdLFRecalculate.Visible Then cmdLFRecalculate.Default = True
        
End Select
End Sub

Private Sub txtBaseFare_GotFocus()
pSetSelected
End Sub

Private Sub txtCommission_Change()
If mblnLoaded Then Call NumbersChanged
End Sub

Private Sub txtCommission_GotFocus()
pSetSelected
End Sub



Private Sub txtDiscount_Change()
If mblnLoaded Then Call NumbersChanged
End Sub

Private Sub chkMerchantFee_Click()

With chkMerchantFee
    If .value = vbChecked Then
        .Caption = "Include Merchant Fee"
        txtMerchantFee.Enabled = True
    Else
        .Caption = "No Merchant Fee"
        txtMerchantFee = ""
        txtMerchantFee.Enabled = False
    End If
End With

If mblnLoaded Then Call NumbersChanged


End Sub

Private Sub cmdRecalculate_Click()
gobjLog.ProcedureName = "cmdRecalculate_Click"
On Error GoTo ProcError

Dim strMsg As String

If cboClientType.Text = "TF" And (Trim(txtTrxFee.Text) = "") And chkTFOverride.value = vbChecked Then strMsg = "Need Transaction Fee amount..." & vbCrLf
If chkMerchantFee.value = vbChecked And Trim(txtMerchantFee) = "" Then strMsg = strMsg & "Need Merchant Fee percentage..." & vbCrLf
'If chkMerchantFee.Value = vbChecked And chkFareType.Value = vbChecked Then
'    chkMerchantFee.Value = vbUnchecked
 '   txtMerchantFee.Text = ""
'End If

'Added by JiYong, Hidden commission is not allowed if this is the markup up net fare
'Preethi - V1.2.1 20101011 - CR21 - Nett Fare Mark Up
'To allow validation for SG also
'If gstrAgcyCountryCode = "HK" Then
   If chkFareType.value = vbUnchecked Then
      If Trim(txtHiddenComm) <> "" Then
         If CSng(txtHiddenComm) > 0 Then
            With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
                 If Trim(txtBaseFare.Text) <> "" Then
                    If CSng(Trim(txtBaseFare)) > .ActualNetAmount And .ActualNetAmount > 0 Then
                       strMsg = strMsg & "Hidden commission is not allowed..." & vbCrLf
                    End If
                 End If
            End With
         End If
      End If
   End If
'End If

If strMsg <> "" Then
    'MsgBox strMsg, vbApplicationModal + vbCritical + vbOKOnly, "MISSING OR INVALID DATA!"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    Exit Sub
End If

mblnLoaded = True
cmdRecalculate.Caption = "&Recalculate Fare"
Call CalculateFare

cmdRecalculate.Visible = False

If mblnLFExist Then
    cmdSave.Enabled = True
    cmdSaveAdd.Enabled = True
    cmdSave.Default = True
    cmdSave.SetFocus
End If

mblnFQExist = True

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub cmdSave_Click()
Dim intI As Integer
Dim strRes As String
Dim SysStart As Date
Dim strNum As String
Dim strMsg As String
SysStart = Now
datTouchEnd = Now

gobjLog.ProcedureName = "cmdSave_Click"
On Error GoTo ProcError
 
 If cmdRecalculate.Visible Then
    SSTab1.Tab = 0
    strMsg = "Need to recalculate Fare Quote"
    'MsgBox "Need to recalculate Fare Quote"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Exit Sub
 End If
 
 If cmdLFRecalculate.Visible Then
    SSTab1.Tab = 4
    strMsg = "Need to recalculate Low Fare Quote"
    'MsgBox "Need to recalculate Low Fare Quote"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Exit Sub
 End If

If lstRmks.ListCount = 0 Then
    strMsg = "Do you want to add remarks for the fare quote?"
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop - Reminder") = vbYes Then
    'If MsgBox("Do you want to add remarks for the fare quote?", vbApplicationModal + vbQuestion + vbYesNo, "Fare quote") = vbYes Then
        SSTab1.Tab = 1
        Exit Sub
    End If
 End If
 
 If lstLFRmks.ListCount = 0 And chkLFNoLowerFare = vbUnchecked Then
    strMsg = "Do you want to add remarks for the low fare option?"
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop - Reminder") = vbYes Then
    'If MsgBox("Do you want to add remarks for the low fare option?", vbApplicationModal + vbQuestion + vbYesNo, "Fare quote") = vbYes Then
        SSTab1.Tab = 4
        Exit Sub
    End If
 End If
 
Me.Hide
frmWait.Show
If chkAddRI.value = 1 Then
   Call WriteQuoteToGDS
End If

'Added on 26/05/2005
If mblnAmend Then
    DeleteRI
End If



With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
    'Added on 19/08/04: reassign in case the user adjust the base amount
    .BaseAmount = CSng(txtBaseFare.Text)
    .MerchAmt = msngMerchFeeAmt
    .SellAmount = msngSellFare
    .NetAmount = IIf(gbolNetFare = True, msngBaseFare, 0)
    .TransactionFee = msngTransFee
    .HiddenComm = CSng(txtHiddenComm.Text)
    .DiscountPct = fConvertZero(txtDiscount)
    .OverrideTF = IIf(chkTFOverride = 1, True, False)
    .International = IIf(chkItinType = 1, True, False)
    .ClientType = cboClientType.Text
    '.NettFare = IIf(chkFareType = 0, True, False)
    .MerchPct = fConvertZero(txtMerchantFee)
     '230108
    .FuelSurcharge = msngFuelSurcharge
    .TaxTotal = msngTax
    
    'If Me.chkLFNoLowerFare = vbUnchecked Then
    '    .LowFare = msngLFSellFare + .TaxTotal
    'Else
    '    .LowFare = .BaseAmount + .TaxTotal
    'End If
    
    'Modified on 26/07/04 to include discount amount
    .DiscountAmt = msngDiscAmt
    
    'Added on 15/03/05 to include NRCC Checkbox
    'If UCase(gstrAgcyCountryCode) = "HK" Then
    .NRCC = chkNRCC(0).value
    'End If
    gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(1).CommissionOnTicket = txtCommission
    
    If Me.chkLFNoLowerFare = vbUnchecked Then
        .LowFare = msngLFSellFare + msngLFTax - msngLFDiscAmt
    Else
        .LowFare = msngSellFare + msngTax - msngDiscAmt
    End If
  
    If cmbPx.listindex <> cmbPx.ListCount - 1 Then
       clearControls
       ClearVar
       If mblnAmend = True Then
       Set gobjPNR = New CWT_GalileoPNR3.PNR
       gobjPNR.loadPNR
       GetRI
       End If
       cmbPx.listindex = cmbPx.listindex + 1
       chkAddRI.value = vbChecked
       mblnLFLoaded = False
       mblnLFExist = False
       SSTab1.Tab = 0
       Me.Show
       frmWait.Hide
       Exit Sub
    End If
    
    'added on 6/1/05
    If mblnAmend = True Then
        Call updateObjToTables
    Else
        If cmbPx.listindex = cmbPx.ListCount - 1 And Not .StoreFare Then
            Call addObjToTables
        End If
    End If
    
    
    'Timer
    Call pAddToVBILog(gobjPNR.RecLoc, "Fare Quote", gStartFareQuoteTime, gGetfareStart, "Get Fares", gGetfareEnd, startTime)
    Call pAddToVBILog(gobjPNR.RecLoc, "Fare Quote", startTime, SysStart, "Write GDS", , startTime)
    
    'Write to SQL Log
    WriteToLog
    
    If .StoreFare Then
        
        Load frmPricingWiz1
        Unload Me
        Unload frmWait
        frmPricingWiz1.Show
    Else
        Unload Me
        Unload frmWait
        'Call pRedisplayMenu
        'If CheckPreTrip = True Then
        '    frmPreTrip.Show
        'End If
    End If
End With


Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub
Private Sub clearControls()


Dim crtl As Control
    
    For Each crtl In Controls
        If TypeOf crtl Is TextBox Then crtl.Text = ""
        If TypeOf crtl Is ListBox Then crtl.Clear
        If TypeOf crtl Is CheckBox Then crtl.value = vbUnchecked
    Next crtl

lblLFNetBaseFareText.Caption = ""
lblLFCommissionAmtText.Caption = ""
lblLFDiscountAmtText.Caption = ""
lblLFlMerchFeeAmount.Caption = ""


End Sub
'Modified on 31/03/2005
Private Sub addObjToTables()
Dim strSQL As String
Dim intC As Integer
Dim intD As Integer
Dim intI As Integer
Dim intJ As Integer
Dim intPxNo As Integer
Dim strRecLoc As String
Dim intNewSegID As Integer
Dim strSelSeg() As String
Dim strAgtName As String
Dim rsRec As ADODB.Recordset

strRecLoc = gobjPNR.RecLoc
strAgtName = gobjHost.AgentName



strSQL = "Select Max(SegID) as MaxNum from tblFareQuote where [RecLoc] = '" & strRecLoc & "'"
Set rsRec = New ADODB.Recordset
rsRec.Open strSQL, gdbConn, 1, 2
If Not rsRec.EOF Then
    If IsNull(rsRec![maxNum]) Then
    intNewSegID = 1
    Else
    intNewSegID = rsRec![maxNum] + 1
    End If
End If

rsRec.Close

For intI = 1 To cmbPx.ListCount

intPxNo = gobjFareQuotes(intI).FQ(1).PxNum


With gobjFareQuotes(intI).FQ(1)



'230108: FUEL SURCHARGE

strSQL = "INSERT into tblFareQuote(RecLoc,SegId,PxID,BaseAmount,BaseCurrency, " & _
         "Commission,DiscountAmt,EquivAmount,EquivCurrency,FareConstructText, " & _
         "HighFareComponent,HPFApplies,ISO,ITNum,JourneyNum,JourneyType,LowFare, " & _
         "MerchAmt,NetAmount,PFType,LastTktDate,PFAccountCode,PIC,PrivateFare, " & _
         "QuoteType,ROE,SegmentSelected,SegmentSelectText,SellAmount,StoreFare, " & _
         "TCNum,TotalCurrency,TotAmount,TransactionFee,FuelSurcharge,UnableToQuote,ModifiedBy, " & _
         "ModifiedDate,NRCC,ComPct,MerchPct,DiscountPct,OverrideTF,International,ClientType,OverrideConx,FQPCC,PlatCarrier,Cat35,Cat35CommType,Ptkt,RuleID, HiddenComm, ActualNetAmount) " & _
         "VALUES('" & strRecLoc & "','" & intNewSegID & "'," & intPxNo & ", " & _
         "" & .BaseAmount & ",'" & .BaseCurrency & "'," & .Commission & "," & .DiscountAmt & ", " & _
         "" & .EquivAmount & ",'" & .EquivCurrency & "','" & .FareConstructText & "', " & _
         "" & .HighFareComponent & "," & IIf(.HPFApplies = True, 1, 0) & ",'" & .ISO & "','" & .ITNum & "', " & _
         "" & .JourneyNum & ",'" & .JourneyType & "'," & .LowFare & "," & .MerchAmt & ", " & _
         "" & .NetAmount & ",'" & .PFFareType & "', "

        If .LastTktDate <> CdatDefaultDate Then
            strSQL = strSQL & "'" & .LastTktDate & "', "
        Else
            strSQL = strSQL & "NULL, "
        End If
        
strSQL = strSQL & "'" & .PFAccountCode & "','" & .PIC & "'," & IIf(.PrivateFare = True, 1, 0) & ", " & _
        "'" & .QuoteType & "'," & .ROE & "," & IIf(.SegmentSelected = True, 1, 0) & ",'" & .SegmentSelectString & "', " & _
        "" & .SellAmount & "," & IIf(.StoreFare = True, 1, 0) & ",'" & .TCNum & "','" & .TotalCurrency & "', " & _
        "" & .TotAmount & "," & .TransactionFee & "," & .FuelSurcharge & "," & IIf(.UnableToQuote = True, 1, 0) & ",'" & strAgtName & "', " & _
        " getDate()," & IIf(.NRCC = True, 1, 0) & "," & .ComPct & "," & .MerchPct & "," & .DiscountPct & ", " & _
        "'" & IIf(.OverrideTF = True, 1, 0) & "', '" & IIf(.International = True, 1, 0) & "', '" & .ClientType & "', " & _
        "'" & IIf(.OverrideConx = True, 1, 0) & "', '" & .FQPCC & "', '" & .PlatCarrier & "', '" & IIf(.Cat35 = True, 1, 0) & "', '" & IIf(.CommType = True, 1, 0) & "', '" & IIf(.PTkt = True, 1, 0) & "', " & IIf(.RuleID <> "", "'" & .RuleID & "'", "Null") & ", '" & .HiddenComm & "'," & .ActualNetAmount & ")"

End With

gdbConn.Execute strSQL


For intC = 1 To gobjFareQuotes(intI).FQ(1).FareComponentCount
    With gobjFareQuotes(intI).FQ(1).FareComponent(intC)
    
strSQL = "INSERT into tblFareComponent (RecLoc,SegID,PxID,FCSeq,Amount, " & _
         "CommissionOnTicket,CurrencyCode,Destination,DestinationCountry, " & _
         "DestinationRegion,DirectionalInd,ETRequired,FareOnTicket, " & _
         "FBC,FOPCode,HPFApplies,ManualProcRequired,MPM, " & _
         "Net,NRCC,OpenSegAllow,Origin,OriginCountry,OriginRegion,OTWCarrier, " & _
         "PaperTktSurcharge,PriceFBC,RuleNum,RuleText,TicketFBC, " & _
         "TktDesignator,TktInfoApplies,TPM,ValueCode,Vendor, " & _
         "WLAllow,ModifiedBy,ModifiedDate) VALUES('" & strRecLoc & "', " & _
         "" & intNewSegID & "," & intPxNo & "," & intC & "," & .Amount & ", " & _
         "" & .CommissionOnTicket & ",'" & .CurrencyCode & "','" & .Destinantion & "', " & _
         "'" & .DestinationCountry & "','" & .DestinationRegion & "', " & _
         "'" & .DirectionalInd & "'," & IIf(.ETRequired = True, 1, 0) & ",'" & .FareOnTicket & "', " & _
         "'" & .FBC & "','" & .FOPCode & "'," & IIf(.HPFApplies = True, 1, 0) & "," & IIf(.ManualProcRequired = True, 1, 0) & ", " & _
         "" & .MPM & "," & IIf(.Net = True, 1, 0) & "," & IIf(.NRCC = True, 1, 0) & "," & IIf(.OpenSegAllow, 1, 0) & ",'" & .Origin & "', " & _
         "'" & .OriginCountry & "','" & .OriginRegion & "','" & .OTWCarrier & "', " & _
         "" & .PaperTktSurcharge & ",'" & .PriceFBC & "','" & .RuleNum & "','" & .RuleText(1) & "', " & _
         "'" & .TicketFBC & "','" & .TktDesignator & "'," & IIf(.TktInfoApplies = True, 1, 0) & ", " & _
         "" & .TPM & ",'" & .ValueCode & "','" & .Vendor & "'," & IIf(.WLAllow = True, 1, 0) & ", " & _
         "'" & strAgtName & "',getDate())"
            
    gdbConn.Execute strSQL

    
  'Adding tblFCEndorse

    For intD = 1 To .EndorsementCount
        strSQL = "INSERT into tblFCEndorse (RecLoc,SegID,PxID,FCSeq,EndorseSeq, " & _
        "EndorseText,ModifiedBy,ModifiedDate) " & _
        "VALUES('" & strRecLoc & "'," & intNewSegID & "," & intPxNo & "," & intC & ", " & _
        "" & intD & ",'" & .Endorsement(intD) & "','" & strAgtName & "', " & _
        " getDate())"
   
        gdbConn.Execute strSQL
    Next

    End With
Next



For intC = 1 To gobjFareQuotes(intI).FQ(1).FareSegCount
    With gobjFareQuotes(intI).FQ(1).FareSeg(intC)
    
    strSQL = "INSERT into tblFareSeg (RecLoc,SegID,PxID,SegSeq,ArrCityCode, " & _
             "BagInfo,DepCityCode,FBC,OverridePFBC,InfoText,NVA,NVB, " & _
             "Stopover,TD,Vendor,ModifiedBy,ModifiedDate,AdviceAct,DepDate,Cos,FlightNum) " & _
             "VALUES('" & strRecLoc & "'," & intNewSegID & "," & intPxNo & ", " & _
             "" & intC & ",'" & .ArrCityCode & "','" & .BagInfo & "', " & _
             "'" & .DepCityCode & "','" & .FBC & "','" & .OverridePFBC & "', " & _
             "'" & .InfoText & "', "

    
    If .NVA <> CdatDefaultDate Then
        strSQL = strSQL & "'" & Format(.NVA, CstrdateFormat) & "', "
    Else
        strSQL = strSQL & "NULL, "
    End If
    
    If .NVB <> CdatDefaultDate Then
        strSQL = strSQL & "'" & Format(.NVB, CstrdateFormat) & "', "
    Else
        strSQL = strSQL & "NULL, "
    End If
     
    strSQL = strSQL & "" & IIf(.Stopover = True, 1, 0) & ",'" & .TD & "','" & .Vendor & "', " & _
           "'" & strAgtName & "',getDate(), "
        
    If gobjFareQuotes(intI).FQ(1).SegmentSelectString = "" Then
          strSQL = strSQL & " '" & gobjPNR.AirSeg(intC).Status & "', " & _
                            " '" & Format(gobjPNR.AirSeg(intC).DepartDateTime, CstrdateFormat) & "', " & _
                            " '" & gobjPNR.AirSeg(intC).Class & "', '" & gobjPNR.AirSeg(intC).FlightNumber & "')"
        Else
           strSelSeg = Split(gobjFareQuotes(intI).FQ(1).SegmentSelectString, ".")
                For intD = 1 To gobjPNR.AirSegCount
                    If gobjPNR.AirSeg(intD).segnumber = strSelSeg(intC - 1) Then
                        strSQL = strSQL & "'" & gobjPNR.AirSeg(intD).Status & "', " & _
                        "'" & Format(gobjPNR.AirSeg(intD).DepartDateTime, CstrdateFormat) & "', " & _
                        "'" & gobjPNR.AirSeg(intD).Class & "','" & gobjPNR.AirSeg(intD).FlightNumber & "')"
                        Exit For
                    End If
                Next intD
    End If

    gdbConn.Execute strSQL
    End With
Next


For intC = 1 To gobjFareQuotes(intI).FQ(1).SurchargeCount
    With gobjFareQuotes(intI).FQ(1).Surcharge(intC)
    strSQL = "INSERT into tblFareSurcharge (RecLoc,SegID,PxID,SurchNum, " & _
             "Amount,CurrencyCode,RelatedSegment,ModifiedBy,ModifiedDate) " & _
             "VALUES('" & strRecLoc & "'," & intNewSegID & "," & intPxNo & ", " & _
             "" & intC & "," & .Amount & ",'" & .CurrencyCode & "', " & _
             "'" & .RelatedSegment & "','" & strAgtName & "', getDate())"
     
    End With
    
    gdbConn.Execute strSQL
Next

For intC = 1 To gobjFareQuotes(intI).FQ(1).TaxCount
    With gobjFareQuotes(intI).FQ(1).Tax(intC)
 
    strSQL = "INSERT into tblFareTax (RecLoc,SegID,PxID,TaxNum,Amount,TaxCode, " & _
             "ModifiedBy,ModifiedDate) VALUES('" & strRecLoc & "'," & intNewSegID & ", " & _
             "" & intPxNo & "," & intC & "," & .Amount & ",'" & .TaxCode & "', " & _
             "'" & strAgtName & "',getDate())"

        gdbConn.Execute strSQL

    End With
Next

Next intI


End Sub


Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim objFQ As CWT_Galileo3.FareQuote
Dim oldParent As Long
Dim strPax() As String
gobjLog.ModuleName = Me.Name
gobjLog.ProcedureName = "Form_Load"
On Error GoTo ProcError
    
datFormLoadStart = Now
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
Me.Move 0, 0
Me.Move frmSideBar.Width, 0
Screen.MousePointer = vbDefault
cmbPx.Clear

'Timer
startTime = Now
'mblnAmend = frmFareQuoteRequest.optFQAmend
mblnAmend = gblnAmend

If mblnAmend Then

        'For intX = 1 To frmFareQuoteRequest.lsvFQPx.ListItems.Count
        For intX = LBound(gstrFQPax) To UBound(gstrFQPax)
            strPax = Split(gstrFQPax(intX), ";")
        '    frmFareQuote.cmbPx.AddItem frmFareQuoteRequest.lsvFQPx.ListItems(intX).SubItems(1)
            frmFareQuote.cmbPx.AddItem strPax(1)
            'If frmFareQuoteRequest.lsvFQPx.ListItems(intX).SubItems(2) = "AD" Then
            If strPax(2) = "AD" Then
               frmFareQuote.cmbPx.ItemData(frmFareQuote.cmbPx.NewIndex) = 1
            Else
               frmFareQuote.cmbPx.ItemData(frmFareQuote.cmbPx.NewIndex) = 0
            End If
        Next intX

Else
        For intX = 1 To gobjFareQuotes.PxCount
            strPax = Split(gstrFQPax(intX - 1), ";")
            frmFareQuote.cmbPx.AddItem strPax(1)
            'frmFareQuote.cmbPx.AddItem frmFareQuoteRequest.lsvPx.ListItems(intX).SubItems(1)
            '29122004
            'If frmFareQuoteRequest.lsvPx.ListItems(intX).SubItems(2) = "AD" Then
            If strPax(2) = "AD" Then
               frmFareQuote.cmbPx.ItemData(frmFareQuote.cmbPx.NewIndex) = 1
            Else
               frmFareQuote.cmbPx.ItemData(frmFareQuote.cmbPx.NewIndex) = 0
            End If
        Next intX
End If
cmbPx.listindex = 0

'29122004
If gbolSkipAdult Then
   For intX = 0 To cmbPx.ListCount - 1
       If cmbPx.ItemData(intX) = 0 Then
          cmbPx.listindex = intX
          'mintStartPxNum = intX
          Exit For
       End If
   Next
End If


Call GetALComm

Call PopulateControls

Call PopulateRmks

Call GetHiddenComm

'Added on 25/05/2005: Get RI list for Farequote amendment
If mblnAmend Then
    Call GetRI
    SSTab1.TabEnabled(3) = True
    SSTab1.TabEnabled(2) = False
    cmdSaveAdd.Visible = False
    lblTitle.Caption = "Fare Quote Amendment"
Else
   SSTab1.TabEnabled(3) = False
   SSTab1.TabEnabled(2) = True
   cmdSaveAdd.Visible = True
   lblTitle.Caption = "Fare Quotation"
End If

'remove on 16/05/2005: Allow class override with the class user specified in FareQuote Screen
'If gbolOverrideFare = True Then
'   txtBaseFare = ""
'   txtTrxFee = ""
'End If

SSTab1.Tab = 0

cmdRecalculate.Default = True

mblnLoaded = True


    If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
    Else
      cmdPrevious.Visible = False
    End If

'pSendToFP ("*NPZ")
pDisplayToFP ("*NPZ")

datFormLoadEnd = Now
If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

Exit Sub

ProcError:
    Call pErrorReport(True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    If Not fWantToQuit Then
        Cancel = 1
    Else
        'Unload frmFake
    End If
Else
    'Unload frmFake
End If

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub
Private Sub PopulateRmks()

'Preethi - V1.2.2 20110223 - CR42 - New Fare Remarks For Fare Quote Screen
Dim rs As ADODB.Recordset
Dim strSQL As String

'cboOthRmks.AddItem "THIS MUST BE ISSUED AS AN ELECTRONIC TICKET"
'cboOthRmks.AddItem "[AMOUNT] SURCHARGE APPLIES FOR PAPER TICKET"
'Added on 31/12/04: Additional IR
'cboOthRmks.AddItem "THIS IS A BUSINESS TRIP AND WILL BE CHARGED TO THE COMPANY"
'cboOthRmks.AddItem "THIS IS A PERSONAL TRIP AND WILL REQUIRE CASH ON DELIVERY"
'cboOthRmks.AddItem "IN CASE OF NO SHOW, TICKET IS NONREFUNDABLE AND NOT VALID FOR TRAVEL"
'cboOthRmks.AddItem "REBOOKING MUST BE DONE ON/BEFORE ORIGINAL DEPARTURE DATE"
'cboOthRmks.AddItem "NO PARTIAL REFUND"
'cboOthRmks.AddItem "NO MILEAGE ACCURAL"
'cboOthRmks.AddItem "NO SHOW PENALTY APPLIES"
'cboOthRmks.AddItem "TICKET MUST BE ISSUED WITHIN 1 DAY OF RESERVATION MADE"
'cboOthRmks.AddItem "THIS IS A NORMAL AIRFARE"

'cboLFOthRmks.AddItem "THIS MUST BE ISSUED AS AN ELECTRONIC TICKET"
'cboLFOthRmks.AddItem "[AMOUNT] SURCHARGE APPLIES FOR PAPER TICKET"
'cboLFOthRmks.AddItem "THIS IS A BUSINESS TRIP AND WILL BE CHARGED TO THE COMPANY"
'cboLFOthRmks.AddItem "THIS IS A PERSONAL TRIP AND WILL REQUIRE CASH ON DELIVERY"
'cboLFOthRmks.AddItem "IN CASE OF NO SHOW, TICKET IS NONREFUNDABLE AND NOT VALID FOR TRAVEL"
'cboLFOthRmks.AddItem "REBOOKING MUST BE DONE ON/BEFORE ORIGINAL DEPARTURE DATE"
'cboLFOthRmks.AddItem "NO PARTIAL REFUND"
'cboLFOthRmks.AddItem "NO MILEAGE ACCURAL"
'cboLFOthRmks.AddItem "NO SHOW PENALTY APPLIES"
'cboLFOthRmks.AddItem "TICKET MUST BE ISSUED WITHIN 1 DAY OF RESERVATION MADE"
'cboLFOthRmks.AddItem "THIS IS A NORMAL AIRFARE"

strSQL = "Select Remarks from tblFareQuoteRemarks"
RunSQLCommand SQLType.Select_, strSQL, gdbConn, rs
While Not rs.EOF
 cboOthRmks.AddItem (rs!Remarks)
 cboLFOthRmks.AddItem (rs!Remarks)
 rs.MoveNext
Wend
rs.Close

cboOthRmks.listindex = 0
cboLFOthRmks.listindex = 0

End Sub
Private Sub PopulateControls()
gobjLog.ProcedureName = "PopulateControls"
On Error GoTo ProcError

Dim lngC As Long
Dim lngNS As Long
Dim sngHPF As Single
Dim strTemp As String
Dim intI As Integer
Dim intSegCount As Integer
Dim strPreSegArr As String
Dim strSeparator As String
Dim lngY As Integer

'added on 27/12
chkItinType.value = vbUnchecked

'modified on 13/1/2005
If cboClientType.ListCount = 0 Then
    For lngC = 0 To 4
        cboClientType.AddItem Mid("DUTFMGMNDB", ((lngC * 2) + 1), 2)
    Next lngC
End If



    If gobjPNR.FOPType = "CC" Then
        chkMerchantFee.value = vbChecked
        Call chkMerchantFee_Click
    
        txtMerchantFee.Text = gobjPNR.CompInfo.MerchFeePct
     
    End If

With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
    If .PrivateFare = True Then
        chkFareType.value = vbUnchecked
        'modified on 13/1/2005
        chkLFFareType.value = vbUnchecked
    End If
    Call chkFareType_Click
    
    If mblnAmend = True Then
        txtBaseFare.Text = .BaseAmount
    Else
        txtBaseFare.Text = IIf(.EquivAmount > 0, .EquivAmount, .BaseAmount)
    End If
    txtTax.Text = .TaxTotal
    If .PrivateFare = False Then
        chkFareType.value = vbChecked
        'modified on 13/1/2005
        chkLFFareType.value = vbChecked
    End If
    
    If .HPFApplies = True Then
        For lngC = 1 To .FareComponentCount
            If .FareComponent(lngC).Amount > sngHPF Then _
                sngHPF = (.FareComponent(lngC).Amount * .ROE) * .FareComponentCount
        Next
        txtBaseFare.Text = sngHPF
        With lblMsg
            .BackColor = &HFFFF&
            .Width = 2340
            .Left = 900
            .Caption = "Higher Point Fare Applies"
        End With
    End If
    'Added on 28/10/04
    If .StoreFare = False Then
        cmdSaveAdd.Visible = True
    Else
        cmdSaveAdd.Visible = False
    End If
    If .PTkt = True Then
        lblTktType.Caption = "Paper Ticket"
    Else
       lblTktType.Caption = "E-Ticket"
    End If

End With

If mblnAmend = True Then
With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
            If .ClientType = "" Then
                cboClientType = IIf(gobjPNR.CompInfo.ClientType = "", "DU", gobjPNR.CompInfo.ClientType)
            Else
                cboClientType = .ClientType
            End If
            

End With
Else
With gobjPNR.CompInfo
    cboClientType = IIf(.ClientType = "", "DU", .ClientType)
    msngMarkup = .MarkUp
    msngDomDiscount = .DiscountDomestic
    msngIntlDiscount = .DiscountInternational
    
    'Removed on 031104: use discount% as it is
    'If msngIntlDiscount <= msngDomDiscount Then
    '    txtDiscount = msngIntlDiscount
    'Else
    '    txtDiscount = msngDomDiscount
    'End If
    txtDiscount = IIf(chkItinType.value = vbChecked, msngIntlDiscount, msngDomDiscount)
End With
End If
    If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PrivateFare = True Then
        txtCommission = msngMarkup
    Else
        txtCommission = msngIntlComm
    End If

If cboClientType = "TP" Or cboClientType = "TF" Then
    txtTrxFee.Enabled = True
    chkTFOverride.Enabled = True
    txtLFTrxFee.Enabled = True
    chkLFTFOverride.Enabled = True
    If mblnAmend = True Then
    txtTrxFee.Text = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TransactionFee
    
    Else
       If gobjPNR.CompInfo.TFCalcBy = "COUPON" Then
          txtTrxFee.Text = fTrxFeeByCoupon(gobjPNR.CompInfo.TransactionFeeGroup)
       'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
       ElseIf gobjPNR.CompInfo.TFCalcBy = "PAX" Then
          txtTrxFee.Text = fTrxFeeByPax(gobjPNR.CompInfo.ClientID, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PxNum)
       Else
          txtTrxFee.Text = fTrxFee(gobjPNR.CompInfo.TransactionFeeGroup, CSng(txtBaseFare.Text))
       End If
    End If
End If

chkItinType.value = vbChecked
Call chkItinType_Click

'Added on 121108: Retrieve Hidden Comm


'Added on 8 Dec: Routing Separator '//' for surface segment


strTemp = ""
strPreSegArr = ""

If mblnAmend Then
    For lngY = 1 To gobjFareQuotes(1).FQ(1).FareSegCount
            For lngC = 1 To gobjPNR.AirSegCount
                With gobjFareQuotes(1).FQ(1).FareSeg(lngY)
                    If gobjPNR.AirSeg(lngC).ArriveAirport = .ArrCityCode And gobjPNR.AirSeg(lngC).DepartAirport = .DepCityCode And gobjPNR.AirSeg(lngC).Vendor = .Vendor And gobjPNR.AirSeg(lngC).FlightNumber = .FlightNum And gobjPNR.AirSeg(lngC).DepartDateTime = .DepDate And gobjPNR.AirSeg(lngC).Class = .Cos Then
                            'gobjPNR.AirSeg(lngC).SelectedForPricing = True
                            If strPreSegArr = gobjPNR.AirSeg(lngC).DepartCityName Or strTemp = "" Then
                                strSeparator = "/"
                            Else
                                strSeparator = "//"
                            End If
                            
                            strTemp = IIf(strTemp <> "", strTemp, gobjPNR.AirSeg(lngC).DepartCityName) & strSeparator & IIf(strSeparator = "//", gobjPNR.AirSeg(lngC).DepartCityName & "/" & gobjPNR.AirSeg(lngC).ArriveCityName, gobjPNR.AirSeg(lngC).ArriveCityName)
                            strPreSegArr = gobjPNR.AirSeg(lngC).ArriveCityName
  

                    End If
                End With
            Next lngC
    Next lngY
Else
    For lngC = 1 To gobjPNR.AirSegCount
        If gobjPNR.AirSeg(lngC).SelectedForPricing = True Then
        
        If strPreSegArr = gobjPNR.AirSeg(lngC).DepartCityName Or strTemp = "" Then
            strSeparator = "/"
        Else
            strSeparator = "//"
        End If
        
        strTemp = IIf(strTemp <> "", strTemp, gobjPNR.AirSeg(lngC).DepartCityName) & strSeparator & IIf(strSeparator = "//", gobjPNR.AirSeg(lngC).DepartCityName & "/" & gobjPNR.AirSeg(lngC).ArriveCityName, gobjPNR.AirSeg(lngC).ArriveCityName)
        strPreSegArr = gobjPNR.AirSeg(lngC).ArriveCityName
        
        End If
    Next
End If

txtRouting.Text = UCase(strTemp)

lblCommissionAmtText.Caption = " "
lblDiscountAmtText.Caption = " "
lblNetBaseFareText.Caption = " "
lblTotalQuoteText.Caption = " "
lblMerchFeeAmount.Caption = " "
If lblMsg.Caption = "MSG" Then lblMsg.Caption = " "


strTemp = ""

For lngC = 1 To gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponentCount
    With gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngC)
    strTemp = strTemp & IIf(strTemp = "", "", vbCrLf & vbCrLf) & "FARE COMPONENT " & lngC & ": " _
        & .Origin & .Destinantion & " (" & .Vendor & ")" & vbCrLf _
        & String(64, "-") & vbCrLf
        For lngNS = 1 To .RuleTextCount
            strTemp = strTemp & .RuleText(lngNS) & vbCrLf
        Next
    End With
    strTemp = strTemp & String(64, "-") & vbCrLf
Next

txtRules.Text = strTemp
Call chkFareType_Click
Call chkMerchantFee_Click

If mblnAmend = True Then
        With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
             chkTFOverride = IIf(.OverrideTF = True, 1, 0)
             chkItinType = IIf(.International = True, 1, 0)
                If .NetAmount > 0 Then
                    chkNRCC(0) = IIf(.NRCC = True, 1, 0)
                    chkFareType = 0
                Else
                    chkNRCC(0) = 0
                    chkFareType = 1
                End If
            txtMerchantFee.Text = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).MerchPct
                If CSng(txtMerchantFee) > 0 Then
                    chkMerchantFee = 1
                Else
                    chkMerchantFee = 0
                End If
            txtCommission = .ComPct
            txtDiscount = .DiscountPct
        End With
End If


'230108: Add Fuel Surcharge for HKG
Dim blnfuelcharge As Boolean
Dim sngfuelAmt As Single

pFuelSurcharge blnfuelcharge, sngfuelAmt

If blnfuelcharge = True Then
    txtFuelSurcharge.Visible = True
    lblFuelSurcharge.Visible = True
    txtLFFuelSurcharge.Visible = True
    lblLFFuelSurcharge.Visible = True
    txtFuelSurcharge.Text = sngfuelAmt
    txtLFFuelSurcharge.Text = sngfuelAmt
Else
    txtFuelSurcharge.Visible = False
    lblFuelSurcharge.Visible = False
    txtLFFuelSurcharge.Visible = False
    lblLFFuelSurcharge.Visible = False
    
End If
'Remove on 7/3/2005
'added on 10/11/04: HKG request - Disabled client type for corporate/group trip
'If gstrAgcyCountryCode = "HK" Then
'    If gTrxnType = "L" Then
'        cboClientType.Enabled = True
'    Else
'        cboClientType.Enabled = False
'    End If
'End If
Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub GetHiddenComm()

Dim strSQL As String
Dim rsHiddenComm As ADODB.Recordset

'Added by Jeremy on 11 Nov 2008. Only add Hidden Comm if FareType is Nett Fare
With gobjFareQuotes(cmbPx.listindex + 1).FQ(1)
       
    If mblnAmend = True Then
        
        msngHiddenComm = gobjFareQuotes(cmbPx.listindex + 1).FQ(1).HiddenComm
        
    Else
    
        strSQL = "Select Markup from tblMarkup where Tier = '" & gobjPNR.CompInfo.MarkUpGroup & "' And StartAmount <= " & .BaseAmount & " and EndAmount >= " & .BaseAmount & ""
        Set rsHiddenComm = gdbConn.Execute(strSQL)
    
        If Not rsHiddenComm.EOF Then
            msngHiddenComm = rsHiddenComm!MarkUp
        Else
            msngHiddenComm = 0
        End If
        rsHiddenComm.Close
        
    End If

    If chkFareType.value = vbUnchecked Then
        txtHiddenComm.Text = msngHiddenComm
    End If

End With

End Sub


Private Sub chkFareType_Click()
gobjLog.ProcedureName = "chkFareType_Click"
On Error GoTo ProcError

If chkFareType.value = vbChecked Then
    chkFareType.Caption = "Published"
    gbolNetFare = False
    txtCommission = msngIntlComm
    lblGrossBaseFare.Caption = "Published Base:"
    'Use below text box for hidden comm
    txtHiddenComm.Visible = True
    lblHiddenComm.Visible = True
    lblCommission.Caption = "Commission (%):"
    lblCommissionAmt.Caption = "Commission Amount:"
    txtCommission = msngIntlComm
    chkNRCC(0).Visible = False
    chkNRCC(0).value = 0
    
    'Added on 11 Nov 2008 by Jeremy.
    'If FareType change to Publish, set Hidden Comm to 0
    'If txtBaseFare.Text <> "" Then
        txtHiddenComm.Text = 0
    'End If

Else
    chkFareType.Caption = "Nett"
    gbolNetFare = True
    txtCommission = msngMarkup
    lblGrossBaseFare.Caption = "Market Net:"
    'Use below text box for hidden comm
    txtHiddenComm.Visible = True
    lblHiddenComm.Visible = True
    lblCommission.Caption = "Mark Up (%):"
    lblCommissionAmt.Caption = "Mark Up Amount:"
    
    'modified on 27/12: This checking should be done separately for Low Fare and Normal Fare
    If (cboClientType = "MN" Or cboClientType = "TF" Or cboClientType = "TP") Then
        msngMarkup = 0
        txtCommission = "0"
    End If
    
    'added on 14/3/2005: add NRCC checkbox for HK
    'If UCase(gstrAgcyCountryCode) = "HK" Then
        chkNRCC(0).Visible = True
        chkNRCC(1).Visible = True
    'Else
    '    chkNRCC(0).Visible = False
    '    chkNRCC(1).Visible = False
    'End If
    
    'Added on 11 Nov 2008 by Jeremy.
    'If FareType change to NettFare, add Hidden Comm
    'If txtBaseFare.Text <> "" Then
        txtHiddenComm.Text = msngHiddenComm
    'End If
    
End If

If mblnLoaded Then Call NumbersChanged

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub WriteQuoteToGDS()
gobjLog.ProcedureName = "WriteQuoteToGDS"
On Error GoTo ProcError

Dim lngC As Long
Dim strCmd As String
Dim strPIC As String
Dim strMsg As String
Call pClearWindow

'Preethi - V1.2.4 20110531 - CR 19 - Generate Standard Remarks For Visa and Fare Quotes
Dim rs As ADODB.Recordset
Dim strSQL As String

If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"

'Added on 28/10/04
If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PIC = "" Then
    strPIC = "ADULT "
ElseIf Left(gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PIC, 1) = "C" Then
    strPIC = "CHILD "
ElseIf gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PIC = "INF" Then
    strPIC = "INFANT "
Else
    strPIC = ""
End If
    
' 04MAY04 - Per Renee Mui in HK, eliminated the RI Discount Line for HK
'modified on 14/09/04: standardized discount line format: discount amount included in Fare Amount
'strCmd = "RI.******************** FARE QUOTE ********************" _
    & "+RI.********* SUBJECT TO CHANGE WITHOUT NOTICE *********" _
    & IIf(txtRouting = "", "", "+RI.FOR " & txtRouting) _
    & "+RI.FARE: " & gobjFareQuotes(1).TotalCurrency & " " & msngSellFare & " PLUS " & msngTax & " TAXES" _
    & IIf(msngDiscAmt <> 0 And gstrAgcyCountryCode <> "HK", "+RI.LESS DISCOUNT/REBATE AMOUNT: (" & msngDiscAmt & ")", "") _
    & IIf(msngTransFee <> 0, "+RI.PLUS TRANSACTION FEE: " & msngTransFee, "") _
    & "+RI.TOTAL QUOTE: " & gobjFareQuotes(1).TotalCurrency & " " & msngTotalCharge




    
 '230108
strCmd = "RI.***************** FARE QUOTE PER PAX ****************" _
    & "+RI.********* SUBJECT TO CHANGE WITHOUT NOTICE *********" _
    & IIf(txtRouting = "", "", "+RI.FOR " & txtRouting) _
    & "+RI." & strPIC & "FARE: " & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency & " " & (msngSellFare - msngDiscAmt) & " PLUS " & msngTax & " TAXES" _
    & IIf(msngTransFee <> 0, "+RI.PLUS TRANSACTION FEE: " & msngTransFee, "") _
    & IIf(msngFuelSurcharge <> 0, "+RI.PLUS FUEL CHARGE SERVICE FEE: " & msngFuelSurcharge, "") _
    & "+RI.TOTAL QUOTE: " & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency & " " & msngTotalCharge

gobjLog.ObjectName = "gobjHost"
If gobjHost Is Nothing Then Set gobjHost = New CWT_Galileo3.GalileoHost
gobjHost.terminalEntry strCmd

   strCmd = ""
   gobjLog.ObjectName = "With lstRmks"
   With lstRmks
        For lngC = 0 To .ListCount - 1
            .listindex = lngC
            strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI.  " & .Text
        Next
   End With
   
   strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI.****************************************************"
   
 gobjLog.ObjectName = "gobjHost"
 gobjHost.terminalEntry strCmd
   
   If chkLFNoLowerFare = vbUnchecked Then
        'modified on 14/09/04: standardized discount line format: discount amount included in Fare Amount
        'strCmd = "RI. **************** LOWER FARE OPTION *****************" _
            & IIf(txtLFRouting <> "", "+RI. FOR " & txtLFRouting, "") _
            & "+RI. FLYING ON " & txtLFAirline _
            & "+RI. FARE: " & gobjFareQuotes(1).TotalCurrency & " " & msngLFSellFare & " PLUS " & msngLFTax & " TAXES" _
            & IIf(msngLFDiscAmt <> 0 And gstrAgcyCountryCode <> "HK", "+RI. LESS DISCOUNT/REBATE AMOUNT: " & msngLFDiscAmt, "") _
            & IIf(msngLFTransFee <> 0, "+RI. PLUS TRANSACTION FEE: " & msngLFTransFee, "") _
            & "+RI. TOTAL LOWER FARE OPTION: " & gobjFareQuotes(1).TotalCurrency & " " & msngLFTotalCharge
        '230108
        strCmd = "RI. **************** LOWER FARE OPTION *****************" _
            & IIf(txtLFRouting <> "", "+RI. FOR " & txtLFRouting, "") _
            & "+RI. FLYING ON " & txtLFAirline _
            & "+RI. " & strPIC & "FARE: " & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency & " " & (msngLFSellFare - msngLFDiscAmt) & " PLUS " & msngLFTax & " TAXES" _
            & IIf(msngLFTransFee <> 0, "+RI. PLUS TRANSACTION FEE: " & msngLFTransFee, "") _
            & IIf(msngLFFuelSurcharge <> 0, "+RI. PLUS FUEL CHARGE SERVICE FEE: " & msngLFFuelSurcharge, "") _
            & "+RI. TOTAL LOWER FARE OPTION: " & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency & " " & msngLFTotalCharge

        gobjHost.terminalEntry strCmd
        strCmd = ""
            
            gobjLog.ObjectName = "With lstLFRmks"
        With lstLFRmks
            For lngC = 0 To .ListCount - 1
                .listindex = lngC
                strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI.    " & .Text
            Next
        End With
        
    gobjLog.ObjectName = "gobjHost"
    If strCmd <> "" Then
       gobjHost.terminalEntry strCmd
    End If
    strCmd = "RI. *****************************************************"
    gobjHost.terminalEntry strCmd
 Else
         strCmd = "RI. ************ NO LOWER FARE OPTION AVAILABLE ***************"
         gobjHost.terminalEntry strCmd
 End If
    'Preethi - V1.2.4 20110613 - CR 19 - Generate Standard Remarks For Visa and Fare Quotes
    strCmd = ""
    strSQL = "Select OptionValue  from tblModOptions where OptionCode = 'StandardFQRmk' order by OptionSecCode"
    Set rs = gdbConn.Execute(strSQL)
    Do Until rs.EOF
        strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & rs!optionvalue & ""
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    strSQL = ""
    gobjHost.terminalEntry strCmd
    gobjLog.ObjectName = ""
    
    'added on 08/10/04: Ticketing Date
    'strTktDate = InputBox("Please enter the ticketing date (ddMMM): ", "Fare Quote")
    'While Not (Len(strTktDate) = 0 Or IsDate(strTktDate))
    '    If Len(strTktDate) > 0 Then
    '        If Not IsDate(Format(strTktDate, "ddMMM")) Then
    '            MsgBox "Invalid Date Format: " & strTktDate, vbCritical
    '            strTktDate = InputBox("Please enter the ticketing date (ddMMM): ", "Fare Quote")
    '        End If
    '    End If
    'Wend
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Remove ER for SyEx flow
    If gIntModuleType = gModuleType.SYEX Then
        'do nothing as do not require ER
    Else
        If Not gobjHost.ENDPNR("TPRO FQ", True) Then
                strMsg = "Unable to end transaction. Please end transaction before proceeding."
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - End PNR"
                'MsgBox "Unable to end transaction. Please end transaction before proceeding."
        End If
    End If
    

    'pSendToFP "*RI"
     If mblnAmend <> True Then pDisplayToFP "*RI"

gobjLog.ProcedureName = ""

'Added on 14/10/04: add to VBI log table
''Timer
'Call pAddToVBILog(gobjPNR.RecLoc, "Fare Quote", StartTime)

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub
Private Sub CalculateFare()
On Error GoTo ProcError

Call ConvertData

msngCommAmt = fCommAmt(msngBaseFare, msngCommPct, chkFareType.value - 1)
If chkFareType.value = vbChecked Then
    msngNetBaseFare = msngBaseFare - msngCommAmt
Else
    msngNetBaseFare = CSng(txtBaseFare)
End If

If cboClientType.Text = "TF" And chkTFOverride.value = vbUnchecked Then
    If gobjPNR.CompInfo.TFCalcBy = "COUPON" Then
       msngTransFee = fTrxFeeByCoupon(gobjPNR.CompInfo.TransactionFeeGroup)
    'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
    ElseIf gobjPNR.CompInfo.TFCalcBy = "PAX" Then
       msngTransFee = fTrxFeeByPax(gobjPNR.CompInfo.ClientID, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PxNum)
    Else
       msngTransFee = fTrxFee(gobjPNR.CompInfo.TransactionFeeGroup, msngNetBaseFare)
    End If
ElseIf chkTFOverride.value = vbUnchecked Then
    msngTransFee = 0
End If

msngSellFare = msngNetBaseFare + msngCommAmt
msngDiscAmt = fDiscAmt(msngSellFare, msngDiscPct, msngCommAmt)
msngTotalCharge = (msngNetBaseFare + msngCommAmt + msngTransFee + msngTax) - msngDiscAmt
msgMFTotalCharge = fMFTotal(chkNRCC(0), msngTotalCharge, msngTransFee, msngNetBaseFare, msngTax)
msngMerchFeeAmt = fMerchantFee(msgMFTotalCharge, msngMerchFee, chkFareType.value - 1)
'111108
msngTotalCharge = msngTotalCharge + msngMerchFeeAmt + msngFuelSurcharge + txtHiddenComm.Text
'230108
'msngTotalCharge = msngTotalCharge + msngMerchFeeAmt + msngFuelSurcharge
'msngTotalCharge = msngTotalCharge + msngMerchFeeAmt
'131108: Sell Fare need to add Hidden Comm
msngSellFare = msngNetBaseFare + msngCommAmt + msngMerchFeeAmt + txtHiddenComm.Text
'msngSellFare = msngNetBaseFare + msngCommAmt + msngMerchFeeAmt

gobjLog.ProcedureName = "CalculateFare"
gobjLog.ObjectName = "gobjFareQuotes"

gobjFareQuotes(cmbPx.listindex + 1).FQ(1).NetAmount = msngNetBaseFare
gobjFareQuotes(cmbPx.listindex + 1).FQ(1).SellAmount = msngSellFare

'Added on 29/07/04 to include Commission Amount in FareQuotes
gobjFareQuotes(cmbPx.listindex + 1).FQ(1).Commission = msngCommAmt

'Added on 14/12/07 to include commision percentage
If msngCommAmt > 0 And gbolNetFare = False Then
   gobjFareQuotes(cmbPx.listindex + 1).FQ(1).CommissionPt = msngCommPct
End If
gobjLog.ObjectName = ""

lblDiscountAmtText.Caption = msngDiscAmt
lblNetBaseFareText.Caption = msngNetBaseFare
lblTotalQuoteText.Caption = msngTotalCharge
lblMerchFeeAmount.Caption = msngMerchFeeAmt
lblCommissionAmtText.Caption = msngCommAmt
txtTrxFee.Text = msngTransFee

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub CalculateLowFare()
On Error GoTo ProcError
Dim strMsg As String

gobjLog.ProcedureName = "CalculateLowFare"

If txtLFBaseFare.Text = "" Then
    strMsg = "Need LF Base Fare"
    'MsgBox "Need LF Base Fare", vbExclamation + vbOKOnly
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    Exit Sub
End If

Call ConvertLFData

msngLFCommAmt = fCommAmt(msngLFBaseFare, msngLFCommPct, chkLFFareType.value - 1)
If chkLFFareType.value = vbChecked Then
    msngLFNetBaseFare = msngLFBaseFare - msngLFCommAmt
Else
    msngLFNetBaseFare = CSng(txtLFBaseFare)
End If

If cboClientType.Text = "TF" And chkLFTFOverride.value = vbUnchecked Then
    If gobjPNR.CompInfo.TFCalcBy = "COUPON" Then
       msngLFTransFee = fTrxFeeByCoupon(gobjPNR.CompInfo.TransactionFeeGroup)
    'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
    ElseIf gobjPNR.CompInfo.TFCalcBy = "PAX" Then
       msngLFTransFee = fTrxFeeByPax(gobjPNR.CompInfo.ClientID, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PxNum)
    Else
       msngLFTransFee = fTrxFee(gobjPNR.CompInfo.TransactionFeeGroup, msngLFNetBaseFare)
    End If
ElseIf chkLFTFOverride.value = vbUnchecked Then
    msngLFTransFee = 0
End If

msngLFSellFare = msngLFNetBaseFare + msngLFCommAmt
msngLFDiscAmt = fDiscAmt(msngLFSellFare, msngLFDiscPct, msngLFCommAmt)
msngLFTotalCharge = (msngLFNetBaseFare + msngLFCommAmt + msngLFTransFee + msngLFTax) - msngLFDiscAmt
msgLFMFTotalCharge = fMFTotal(chkNRCC(1), msngLFTotalCharge, msngLFTransFee, msngLFNetBaseFare, msngLFTax)
'msngLFMerchFeeAmt = fMerchantFee(msngLFTotalCharge, msngLFMerchFee)
msngLFMerchFeeAmt = fMerchantFee(msgLFMFTotalCharge, msngLFMerchFee, chkLFFareType.value - 1)
'msngLFTotalCharge = msngLFTotalCharge + msngLFMerchFeeAmt
'230108
msngLFTotalCharge = msngLFTotalCharge + msngLFMerchFeeAmt + msngLFFuelSurcharge
msngLFSellFare = msngLFSellFare + msngLFMerchFeeAmt

gobjLog.ProcedureName = "CalculateLowFare"
'gobjLog.ObjectName = "gobjFareQuotes"

'gobjFareQuotes(1).NetAmount = msngLFNetBaseFare
'gobjFareQuotes(1).SellAmount = msngLFSellFare

'gobjLog.ObjectName = ""

lblLFDiscountAmtText.Caption = msngLFDiscAmt
lblLFNetBaseFareText.Caption = msngLFNetBaseFare
lblLFTotalQuoteText.Caption = msngLFTotalCharge
lblLFlMerchFeeAmount.Caption = msngLFMerchFeeAmt
lblLFCommissionAmtText.Caption = msngLFCommAmt
txtLFTrxFee.Text = msngLFTransFee

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub
Private Function fMFTotal(NRCC As Boolean, TotalCharge As Single, TransFee As Single, NetFare As Single, Tax As Single) As Single
gobjLog.ProcedureName = "fMFTotal"
On Error GoTo ProcError

If NRCC Then
    'If gstrAgcyCountryCode = "HK" Then
        If cboClientType = "TF" Then
           fMFTotal = TransFee
        Else
        
           fMFTotal = TotalCharge - NetFare - Tax
        
        End If
    'End If
    
Else
    If cboClientType = "TF" Then
        If gobjPNR.CompInfo.TFIncMF Then
            fMFTotal = TotalCharge
        Else
            fMFTotal = TotalCharge - TransFee
        End If
    Else
        fMFTotal = TotalCharge
    End If
End If



Exit Function
ProcError:
    Call pErrorReport(True)

End Function
Private Function fCommAmt(Fare As Single, CommPct As Single, NetFare As Boolean) As Single
gobjLog.ProcedureName = "fCommAmt"
On Error GoTo ProcError

Dim sngPct As Single
Dim sngAmt As Single
sngPct = CommPct * 0.01

If NetFare = True And (cboClientType = "MN" Or cboClientType = "TF" Or cboClientType = "TP") Then
        fCommAmt = 0
        'Modified on 1 Nov 07. if calculate commission for best fare, don't affect commission in new fare
        If SSTab1.Tab = 0 Then  'New Fare Quote tab
           txtCommission = "0"
        End If
Else
    If NetFare = True Then
        'Added on 24 Nov 08 by Jeremy to add Hidden Comm when calculating Markup
        sngAmt = ((Fare + txtHiddenComm.Text) / (1 - sngPct)) - (Fare + txtHiddenComm.Text)
        'sngAmt = (Fare / (1 - sngPct)) - Fare
        fCommAmt = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP") + IIf(sngAmt > 0, IIf(gstrAgcyCountryCode = "HK", IIf(cboClientType = "DU", 10, 0), 0), 0)
    Else
        sngAmt = Fare * sngPct
        fCommAmt = fCurrRound(sngAmt, gstrAgcyCurrCode, "DOWN")
    End If
End If

Exit Function
ProcError:
    Call pErrorReport(True)


End Function

Private Function fDiscAmt(SellFare As Single, DiscPct As Single, CommAmt As Single) As Single
gobjLog.ProcedureName = "fDiscAmt"
On Error GoTo ProcError

Dim sngPct As Single

sngPct = DiscPct * 0.01

Select Case cboClientType
Case "DU"
    'fDiscAmt = fCurrRound(SellFare * sngPct, gobjFareQuotes(cmbPx.ListIndex + 1).FQ(1).BaseCurrency, "DOWN")
    fDiscAmt = fCurrRound(SellFare * sngPct, gstrAgcyCurrCode, "DOWN")
Case "MN", "TF", "TP"
    fDiscAmt = CommAmt
Case Else
    fDiscAmt = 0
End Select

Exit Function
ProcError:
    Call pErrorReport(True)

End Function

Private Sub txtBaseFare_Change()
If mblnLoaded Then Call NumbersChanged
End Sub

Private Sub NumbersChanged()
cmdRecalculate.Visible = True
cmdRecalculate.Default = True
cmdSave.Enabled = False
cmdSaveAdd.Enabled = False
End Sub

Private Sub txtDiscount_GotFocus()
pSetSelected
End Sub

Private Sub txtFuelSurcharge_Change()
If mblnLoaded Then Call NumbersChanged
End Sub

Private Sub txtHiddenComm_Change()
If mblnLoaded Then Call NumbersChanged
End Sub

Private Sub txtLFCWTNet_Change()
If mblnLFLoaded Then Call LFNumbersChanged
End Sub

Private Sub txtLFFreeText_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 32, 35, 36, 38, 40, 41, 42, 46, 47, 64
            ' ALLOW
        Case Else
            KeyAscii = fAllowAlphaNumeric(KeyAscii)
    End Select
End Sub

Private Sub txtFreeText_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 32, 35, 36, 38, 40, 41, 42, 46, 47, 64
            ' ALLOW
        Case Else
            KeyAscii = fAllowAlphaNumeric(KeyAscii)
    End Select
End Sub


Private Sub txtLFFuelSurcharge_Change()
If mblnLFLoaded Then Call LFNumbersChanged
End Sub

Private Sub txtLFRouting_GotFocus()
pSetSelected
End Sub

Private Sub txtLFTrxFee_Change()
If mblnLFLoaded = True Then Call LFNumbersChanged
End Sub

Private Sub txtLFTrxFee_GotFocus()
pSetSelected
End Sub

Private Sub txtMerchantFee_Change()
If mblnLoaded = True And chkMerchantFee.value = 1 Then Call NumbersChanged
End Sub

Private Sub txtMerchantFee_GotFocus()
pSetSelected
End Sub

Private Sub txtRouting_GotFocus()
pSetSelected
End Sub

Private Sub txtTax_Change()
If mblnLoaded Then Call NumbersChanged
End Sub

Private Sub txtTrxFee_Click()
If mblnLoaded Then Call NumbersChanged
End Sub
Private Function fMerchantFee(TotalCharge As Single, MerchFeePct As Single, NetFare As Boolean) As Single
gobjLog.ProcedureName = "fMerchantFee"
On Error GoTo ProcError

Dim sngPct As Single
Dim sngAmt As Single

sngPct = MerchFeePct * 0.01

'''Added on 111108 By Jeremy
''If NetFare = True And (cboClientType = "MN" Or cboClientType = "MG" Or cboClientType = "DU") Then
    sngAmt = CDec((CDec(TotalCharge) + txtHiddenComm.Text) * sngPct)
''Else
''    'This is the original Merchant Fee
''    sngAmt = CDec(TotalCharge * sngPct)
''End If
'****************

'If chkMerchantFee.Value = vbChecked Then
    'If gstrAgcyCountryCode = "SG" Then
    '    sngAmt = (TotalCharge / (1 - sngPct)) - TotalCharge
    'Else
        ''''''''sngAmt = CDec(TotalCharge * sngPct)
    'End If
    fMerchantFee = fCurrRound(sngAmt, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).TotalCurrency, "UP")
    'fMerchantFee = fCurrRound(sngAmt, gobjFareQuotes(cmbPx.listindex + 1).FQ(1).BaseCurrency, "UP")
'Else
'    fMerchantFee = 0
'End If

Exit Function
ProcError:
    Call pErrorReport(True)

End Function

Private Sub ConvertData()
gobjLog.ProcedureName = "ConvertData"
On Error GoTo ProcError

Dim strErrSource As String

If Trim(txtBaseFare.Text) = "" Then
    msngBaseFare = 0
Else
    msngBaseFare = CSng(txtBaseFare)
End If

If Trim(txtMerchantFee) = "" Then
    msngMerchFee = 0
Else
    msngMerchFee = CSng(txtMerchantFee.Text)
End If

If Trim(txtCommission.Text) = "" Then
    msngCommPct = 0
Else
    msngCommPct = CSng(txtCommission.Text)
End If

If Trim(txtDiscount.Text) = "" Then
    msngDiscPct = 0
Else
    msngDiscPct = CSng(txtDiscount.Text)
End If

If Trim(txtTrxFee) = "" Or cboClientType.Text <> "TF" Then
    msngTransFee = 0
Else
    msngTransFee = CSng(txtTrxFee)
End If

If Trim(txtTax) = "" Then
    msngTax = 0
Else
    msngTax = CSng(txtTax)
End If
'230108
If Trim(txtFuelSurcharge) = "" Then
    msngFuelSurcharge = 0
Else
    msngFuelSurcharge = CSng(txtFuelSurcharge)
    
End If

Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub ConvertLFData()
gobjLog.ProcedureName = "ConvertLFData"
On Error GoTo ProcError

Dim strErrSource As String

If Trim(txtLFBaseFare.Text) = "" Then
    msngLFBaseFare = 0
Else
    msngLFBaseFare = CSng(txtLFBaseFare)
End If

If Trim(txtLFMerchantFee) = "" Then
    msngLFMerchFee = 0
Else
    msngLFMerchFee = CSng(txtLFMerchantFee.Text)
End If

If Trim(txtLFCommission.Text) = "" Then
    msngLFCommPct = 0
Else
    msngLFCommPct = CSng(txtLFCommission.Text)
End If

If Trim(txtLFDiscount.Text) = "" Then
    msngLFDiscPct = 0
Else
    msngLFDiscPct = CSng(txtLFDiscount.Text)
End If

If Trim(txtLFTrxFee) = "" Or cboClientType.Text <> "TF" Then
    msngLFTransFee = 0
Else
    msngLFTransFee = CSng(txtLFTrxFee)
End If

If Trim(txtLFTax) = "" Then
    msngLFTax = 0
Else
    msngLFTax = CSng(txtLFTax)
End If

'230108
If Trim(txtFuelSurcharge) = "" Then
    msngLFFuelSurcharge = 0
Else
    msngLFFuelSurcharge = CSng(txtLFFuelSurcharge)
End If
Exit Sub
ProcError:
    Call pErrorReport(True)

End Sub

Private Sub txtTax_GotFocus()
pSetSelected
End Sub

Private Sub GetALComm()
gobjLog.ProcedureName = "GetALComm"
On Error GoTo ProcError

Dim rsAL As New ADODB.Recordset
Dim strSQL As String
Dim lngC As Long
Dim strMsg As String

If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponentCount = 0 Then GoTo ProcError

'STRSQL = "SELECT * FROM tblAirlines WHERE [Code] = '" & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(1).Vendor & "'"
'If gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponentCount > 1 Then
'    For lngC = 2 To gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponentCount
'        STRSQL = STRSQL & " OR [Code] = '" & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).FareComponent(lngC).Vendor & "'"
'    Next
'End If
strSQL = "SELECT * FROM tblAirlines where [Code]='" & gobjFareQuotes(cmbPx.listindex + 1).FQ(1).PlatCarrier & "'"

gobjLog.ObjectName = "rsAL"

Set rsAL = gdbConn.Execute(strSQL)

If Not rsAL.EOF Then
    If Not IsNull(rsAL![IntlComm]) Then
        msngIntlComm = rsAL![IntlComm]
    Else
        msngIntlComm = 0
    End If
    
    If Not IsNull(rsAL![DomComm]) Then
        msngDOMComm = rsAL![DomComm]
    Else
        msngDOMComm = 0
    End If
End If

rsAL.Close
Set rsAL = Nothing

gobjLog.ObjectName = ""

Exit Sub
ProcError:
    'MsgBox "ERROR " & "Unable to FareQuote! FareQuote Component not Found" & Chr(13), vbApplicationModal + vbCritical + vbOKOnly, "FareQuote Error"
    strMsg = "Unable to FareQuote! FareQuote Component not Found"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Set frmFareQuote = Nothing
    Unload frmFareQuote
    Set gobjFareQuotes = Nothing
    
    
    'Screen.MousePointer = vbHourglass
    Exit Sub
    'frmMainMenu.WindowState = vbNormal
    
    'Screen.MousePointer = vbDefault


End Sub

Private Function fTrxFee(ByVal Group As String, Amount As Single) As Single
gobjLog.ProcedureName = "fTrxFee"
On Error GoTo ProcError

Dim rsTF As New ADODB.Recordset
Dim strSQL As String
Dim TempTF As Single
Dim sngTF() As Single
Dim strTFDest() As String
Dim lngFC As Long
Dim lngTFDest As Long
Dim intHigh As Integer
Dim intCurr As Integer
Dim strJrnyOrigCity As String
Dim strJrnyOrigCountry As String
Dim strFCOrig As String
Dim strFCDest As String

   
   
strSQL = "SELECT * FROM tblTransactionFees WHERE [TFGroup] ='" & Group & "'" _
& " AND ([StartAmount] <=  " & Amount _
& " AND ([EndAmount] >=" & Amount & " OR [EndAmount] = 0))" _
& " ORDER BY [CityCode] ASC,  Len([CityCode]) DESC"

gobjLog.ObjectName = "rsTF"

Set rsTF = gdbConn.Execute(strSQL)
With rsTF

'If Not .EOF Then
'    .MoveLast
'    .MoveFirst
'End If
'use 1st passenger FareComponent since travel on same destination
ReDim sngTF(gobjFareQuotes(1).FQ(1).FareComponentCount - 1)
    
For lngFC = 0 To gobjFareQuotes(1).FQ(1).FareComponentCount - 1
    Do Until .EOF
        strTFDest = Split(![CityCode], ";")
        'evaluating the first entry in each string to assign a hierarchal value
        If strTFDest(LBound(strTFDest)) = "X" Then
            intCurr = 1 'all destinations
            strFCDest = "X"
        ElseIf Len(strTFDest(LBound(strTFDest))) = 3 Then
            intCurr = 4 'city code
            strFCDest = gobjFareQuotes(1).FQ(1).FareComponent(lngFC + 1).Destinantion
        ElseIf Len(strTFDest(LBound(strTFDest))) = 2 Then
            If IsNumeric(strTFDest(LBound(strTFDest))) Then
                intCurr = 2 'region code
                strFCDest = gobjFareQuotes(1).FQ(1).FareComponent(lngFC + 1).DestinationRegion
            Else
                intCurr = 3 'country code
                strFCDest = gobjFareQuotes(1).FQ(1).FareComponent(lngFC + 1).DestinationCountry
            End If
        Else
        End If
        If intCurr > intHigh Then
           
        For lngTFDest = 0 To UBound(strTFDest)
            If strFCDest = strTFDest(lngTFDest) Then
                 intHigh = intCurr
                Select Case ![Type]
                    Case "F"
                        sngTF(lngFC) = ![TFAmount] + ![ExtraAmount]
                    Case "M"
                        sngTF(lngFC) = (Amount * (![TFAmount] * 0.01)) + ![ExtraAmount]
                    Case "D"
                        sngTF(lngFC) = (Amount / ((100 - ![TFAmount]) * 0.01) + ![ExtraAmount]) - Amount
                End Select
                If ![MaxAmount] <> 0 And sngTF(lngFC) > ![MaxAmount] Then sngTF(lngFC) = ![MaxAmount]
                If sngTF(lngFC) < ![MinAmount] Then sngTF(lngFC) = ![MinAmount]
                'sngTF(lngFC) = fCurrRound(sngTF(lngFC), gobjFareQuotes(1).BaseCurrency, "UP")
                'Preethi - V1.2.4 20110613 - CR 59 - Remove Round Up Function on Transaction Fee
                If UCase(gstrAgcyCountryCode) = "SG" Then
                  sngTF(lngFC) = Format(sngTF(lngFC), "#0.00")
                Else
                  sngTF(lngFC) = fCurrRound(sngTF(lngFC), gstrAgcyCurrCode, "UP")
                End If
                Exit For ' match found, don't need to go thru others
            End If
        Next
        End If
    .MoveNext
    Loop
    intHigh = 0
Next

End With
gobjLog.ObjectName = ""

 For lngFC = 0 To gobjFareQuotes(1).FQ(1).FareComponentCount - 1
    If sngTF(lngFC) > fTrxFee Then
        fTrxFee = sngTF(lngFC)
    End If
Next

Exit Function
ProcError:
    Call pErrorReport(True)

End Function

Private Sub chkLFItinType_Click()

If chkLFItinType.value = vbChecked Then
    chkLFItinType.Caption = "International"
   If cboClientType.Text = "DU" Or cboClientType.Text = "MG" Then txtLFDiscount = msngIntlDiscount
Else
    chkLFItinType.Caption = "Domestic"
    If cboClientType.Text = "DU" Or cboClientType.Text = "MG" Then txtLFDiscount = msngDomDiscount
End If
End Sub

Private Sub txtLFBaseFare_GotFocus()
pSetSelected
End Sub

Private Sub txtLFCommission_Change()
If mblnLFLoaded Then Call LFNumbersChanged
End Sub

Private Sub txtLFCommission_GotFocus()
pSetSelected
End Sub

Private Sub txtLFDiscount_Change()
If mblnLFLoaded Then Call LFNumbersChanged
End Sub

Private Sub chkLFMerchantFee_Click()
If mblnLFLoaded Then
    Call LFNumbersChanged
    If chkLFMerchantFee.value = vbChecked Then
        txtLFMerchantFee.Enabled = True
        chkLFMerchantFee.Caption = "Include Merchant Fee"
    Else
        txtLFMerchantFee = ""
        txtLFMerchantFee.Enabled = False
        chkLFMerchantFee.Caption = "No Merchant Fee"
    End If
End If

End Sub

Private Sub txtLFBaseFare_Change()
If mblnLFLoaded Then Call LFNumbersChanged
End Sub

Private Sub LFNumbersChanged()
cmdLFRecalculate.Visible = True
cmdLFRecalculate.Default = True
cmdSave.Enabled = False
cmdSaveAdd.Enabled = False
End Sub

Private Sub txtLFDiscount_GotFocus()
pSetSelected
End Sub

Private Sub txtLFMerchantFee_Change()
If mblnLFLoaded = True And chkMerchantFee.value = 1 Then Call LFNumbersChanged
End Sub

Private Sub txtLFMerchantFee_GotFocus()
pSetSelected
End Sub

Private Sub txtLFTax_Change()
If mblnLFLoaded Then Call LFNumbersChanged
End Sub

Private Sub txtLFTax_GotFocus()
pSetSelected
End Sub

Private Sub GetRI()

   Dim i As Integer
   Dim item As ListItem
   Dim strText As String
   
   
   If lswRI.ListItems.Count > 0 Then lswRI.ListItems.Clear
   For i = 1 To gobjPNR.ItinRemarkCount
      With gobjPNR.ItinRemark(i)
         Set item = lswRI.ListItems.Add(, , Format(.ItemNum, "000"))
         strText = .RemarkText
         item.SubItems(1) = strText
      End With
   Next
  
   For i = 1 To lswRI.ListItems.Count
      lswRI.ListItems(i).Selected = False
   Next

End Sub

Private Sub DeleteRI()
Dim intI As Integer
Dim strNum As String

        If lswRI.ListItems.Count = 0 Then Exit Sub
        
        For intI = 1 To lswRI.ListItems.Count
           If lswRI.ListItems(intI).Selected Then
              strNum = strNum & IIf(strNum <> "", ".", "") & Format(lswRI.ListItems(intI).Text, "0")
           End If
        Next
        If strNum = "" Then Exit Sub
        gobjHost.terminalEntry "RI." & strNum & "@"
        pDisplayToFP ("*RI")
        
End Sub

Private Sub updateObjToTables()
Dim strSQL As String
Dim intI As Integer
Dim intPxNo As Integer
Dim strRecLoc As String
Dim strAgtName As String
Dim rsRec As ADODB.Recordset

strRecLoc = gobjPNR.RecLoc
strAgtName = gobjHost.AgentName


For intI = 1 To cmbPx.ListCount

intPxNo = gobjFareQuotes(intI).FQ(1).PxNum


With gobjFareQuotes(intI).FQ(1)

'230108: FUEL SURCHARGE
'111108: HIDDEN COMM
strSQL = "update tblFareQuote set BaseAmount=" & .BaseAmount & ", " & _
         "MerchAmt=" & .MerchAmt & ", " & _
         "SellAmount=" & .SellAmount & ", " & _
         "NetAmount=" & .NetAmount & ", " & _
         "TransactionFee=" & .TransactionFee & ", " & _
         "HiddenComm=" & .HiddenComm & ", " & _
         "ComPct=" & .ComPct & ", " & _
         "MerchPct=" & .MerchPct & ", " & _
         "FuelSurcharge=" & .FuelSurcharge & ", " & _
         "DiscountPct=" & .DiscountPct & ", " & _
         "NRCC= '" & IIf(.NRCC = True, 1, 0) & "', " & _
         "OverrideTF= '" & IIf(.OverrideTF = True, 1, 0) & "', " & _
         "International= '" & IIf(.International = True, 1, 0) & "', " & _
         "LowFare=" & .LowFare & ", " & _
         "ClientType= '" & .ClientType & "', " & _
         "ModifiedBy='" & strAgtName & "', " & _
         "ModifiedDate= getDate()" & _
         " where RecLoc='" & strRecLoc & "' and SegID='" & gFQSegID & "' and PxID='" & intPxNo & "'  "
         
            

gdbConn.Execute strSQL


End With

Next intI


End Sub

'230108
Private Sub pFuelSurcharge(ByRef blnfuelcharge As Boolean, ByRef sngAmt As Single)
Dim strSQL As String
Dim rs As ADODB.Recordset
strSQL = "select * from tblModoptions where Optioncode='FuelCharge' or optioncode='DefaultFuelCharge'"

blnfuelcharge = False

Set rs = gdbConn.Execute(strSQL)
While Not rs.EOF
    
    If rs!optioncode = "FuelCharge" Then
        If UCase(rs!optionvalue) = "TRUE" Then
            blnfuelcharge = rs!optionvalue
        Else
            blnfuelcharge = False
        End If
    End If
    
    If rs!optioncode = "DefaultFuelCharge" Then
       
            sngAmt = rs!optionvalue
       
    End If
    
    rs.MoveNext
    
Wend



End Sub

Private Function fTrxFeeByCoupon(ByVal Group As String) As Single
    
    Dim i As Integer
    Dim j As Double
    Dim lngY As Integer
    Dim lngC As Integer
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    i = 0
    fTrxFeeByCoupon = 0
    
    'Get number of coupons
    If mblnAmend Then
       For lngY = 1 To gobjFareQuotes(1).FQ(1).FareSegCount
            For lngC = 1 To gobjPNR.AirSegCount
                With gobjFareQuotes(1).FQ(1).FareSeg(lngY)
                    If gobjPNR.AirSeg(lngC).ArriveAirport = .ArrCityCode And gobjPNR.AirSeg(lngC).DepartAirport = .DepCityCode And gobjPNR.AirSeg(lngC).Vendor = .Vendor And gobjPNR.AirSeg(lngC).FlightNumber = .FlightNum And gobjPNR.AirSeg(lngC).DepartDateTime = .DepDate And gobjPNR.AirSeg(lngC).Class = .Cos Then
                       i = i + 1
                    End If
                End With
            Next lngC
        Next lngY
    Else
        For lngC = 1 To gobjPNR.AirSegCount
            If gobjPNR.AirSeg(lngC).SelectedForPricing = True Then
               i = i + 1
            End If
        Next
    End If
    
    If i > 0 Then
       strSQL = "Select * from tblTransactionFeesbyCoupon Where TFGroup='" & gobjPNR.CompInfo.TransactionFeeGroup & "' order by Type desc, StartCoupon asc"
       Set rs = gdbConn.Execute(strSQL)
       Do Until rs.EOF
          If Trim(rs!Type) = "SOLO" Then
             If IsNumeric(rs!EndCoupon) Then
                If CInt(rs!StartCoupon) <= i And i <= CInt(rs!EndCoupon) Then
                   fTrxFeeByCoupon = rs!TFAmount
                   Exit Do
                End If
             ElseIf Trim(rs!EndCoupon) = "X" Then
                If CInt(rs!StartCoupon) <= i Then
                   fTrxFeeByCoupon = rs!TFAmount
                   Exit Do
                End If
             End If
          ElseIf Trim(rs!Type) = "GROUP" Then
             If IsNumeric(rs!EndCoupon) Then
                If CInt(rs!StartCoupon) <= i And i <= CInt(rs!EndCoupon) Then
                   fTrxFeeByCoupon = fTrxFeeByCoupon + rs!TFAmount
                ElseIf i > CInt(rs!EndCoupon) Then
                   fTrxFeeByCoupon = fTrxFeeByCoupon + rs!TFAmount
                End If
             ElseIf Trim(rs!EndCoupon) = "X" Then
                If CInt(rs!StartCoupon) <= i Then
                   If IsNull(rs!Threshold) = False Then
                       j = (i - Int(rs!StartCoupon) + 1) / Int(rs!Threshold)
                       If j < 1 Then
                          fTrxFeeByCoupon = fTrxFeeByCoupon + rs!TFAmount
                       Else
                          fTrxFeeByCoupon = fTrxFeeByCoupon + (Int(j) * rs!TFAmount)
                          If ((i - Int(rs!StartCoupon) + 1) Mod Int(rs!Threshold)) > 0 Then
                              fTrxFeeByCoupon = fTrxFeeByCoupon + rs!TFAmount
                          End If
                       End If
                   ElseIf IsNull(rs!Threshold) = True Then
                       fTrxFeeByCoupon = fTrxFeeByCoupon + rs!TFAmount
                   End If
                End If
             End If
          End If
          rs.MoveNext
       Loop
       rs.Close
       Set rs = Nothing
    End If
End Function

Function WriteToLog()

    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareQuote, _
    Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareQuote, _
    Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareQuote, _
    Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd

End Function

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Private Function fTrxFeeByPax(ByVal ClientID As Long, PaxNum As Integer) As Currency
gobjLog.ProcedureName = "fTrxFeeByPax"
On Error GoTo ProcError

Dim rsTF As New ADODB.Recordset
Dim strSQL As String
Dim curTFAmount As Currency
 
curTFAmount = 0

If SearchGASalesByPC_Pax("35", PaxNum) Then
    curTFAmount = 0
ElseIf NonVoidCouponFound = True Then
    curTFAmount = 0
Else
    strSQL = "SELECT TFAmount FROM tblTransactionFeesByPax "
    strSQL = strSQL & "WHERE ClientID = '" & ClientID & "'"
    
    gobjLog.ObjectName = "rsTF"
    
    Set rsTF = gdbConn.Execute(strSQL)
    With rsTF
        If .EOF = False Then
            If IsNumeric(rsTF!TFAmount & "") Then
                curTFAmount = rsTF!TFAmount
    '        Else
    '            curTFAmount = 0
            End If
        Else
            curTFAmount = 0
        End If
    End With
    gobjLog.ObjectName = ""
End If

fTrxFeeByPax = curTFAmount

Exit Function
ProcError:
    Call pErrorReport(True)

End Function


