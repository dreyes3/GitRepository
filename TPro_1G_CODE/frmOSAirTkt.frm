VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOSAirTkt 
   Appearance      =   0  'Flat
   ClientHeight    =   7890
   ClientLeft      =   1095
   ClientTop       =   1800
   ClientWidth     =   11145
   LinkTopic       =   "CWT Travel Pro - "
   ScaleHeight     =   7890
   ScaleWidth      =   11145
   Begin VB.CheckBox chkAmdTktOnly 
      Caption         =   "Amend Ticket Num Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   97
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   31
      Top             =   720
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ticket Info"
      TabPicture(0)   =   "frmOSAirTkt.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtOriTktNum"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkReissuedTkt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optNettFare(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "optPublishedFare(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkformula"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboClientType"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtContact"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fraHotel"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Picture1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblOriTktNum"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblTktLength"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblCT"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblContact"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Remarks"
      TabPicture(1)   =   "frmOSAirTkt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFreeRmk"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdFreeRmkToEO"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdFreeRmkToItin"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "MI"
      TabPicture(2)   =   "frmOSAirTkt.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtMS"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtRS"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmbDispNum"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdClientMI"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraMI"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lvwRealECodes"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lvwMissECodes"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lvwECodes"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblLabels(40)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lblLabels(41)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "lblLabels(0)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "MCO/MPD"
      TabPicture(3)   =   "frmOSAirTkt.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkRequestMCO"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fraRmk"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "fraTrav"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "fraMCO"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Vendor Info"
      TabPicture(4)   =   "frmOSAirTkt.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "fraVendorInfo"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtOriTktNum 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -66960
         MaxLength       =   10
         TabIndex        =   178
         Top             =   2770
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox chkReissuedTkt 
         Caption         =   "Reissued Ticket"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -69000
         TabIndex        =   177
         Top             =   2550
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton optNettFare 
         Caption         =   "Bill Marked Up Net Fare"
         Height          =   255
         Index           =   1
         Left            =   -72720
         TabIndex        =   176
         Top             =   800
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.OptionButton optPublishedFare 
         Caption         =   "Bill Published Fare"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   175
         Top             =   800
         Visible         =   0   'False
         Width           =   1695
      End
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
         Left            =   -67560
         MaxLength       =   3
         TabIndex        =   160
         Tag             =   "BY-"
         Top             =   3000
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
         Left            =   -67440
         MaxLength       =   3
         TabIndex        =   159
         Tag             =   "BY-"
         Top             =   720
         Width           =   645
      End
      Begin VB.CheckBox chkRequestMCO 
         Caption         =   "Request MCO"
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
         Left            =   -74760
         TabIndex        =   155
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame fraRmk 
         Caption         =   "Remark(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   -69240
         TabIndex        =   123
         Top             =   2760
         Width           =   5055
         Begin VB.CommandButton cmdAddMCORmk 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   127
            Top             =   720
            Width           =   615
         End
         Begin VB.ListBox lstRmks 
            Height          =   2205
            Left            =   120
            MultiSelect     =   1  'Simple
            TabIndex        =   126
            Top             =   1200
            Width           =   4875
         End
         Begin VB.TextBox txtFT 
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
            Left            =   1200
            MaxLength       =   70
            TabIndex        =   122
            Tag             =   "NN"
            Top             =   240
            Width           =   3645
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   124
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "Free Text:"
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
            Left            =   120
            TabIndex        =   125
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.Frame fraTrav 
         Caption         =   "Traveller(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -69240
         TabIndex        =   117
         Top             =   840
         Width           =   5055
         Begin VB.ComboBox cmbTraName 
            Height          =   315
            Left            =   840
            TabIndex        =   121
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton cmdAddTraPax 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   119
            Top             =   360
            Width           =   615
         End
         Begin MSComctlLib.ListView lsvTraveller 
            Height          =   855
            Left            =   120
            TabIndex        =   118
            Top             =   840
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1508
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Traveller Name"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label Label25 
            Caption         =   "Name:"
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
            TabIndex        =   120
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraVendorInfo 
         Height          =   5895
         Left            =   360
         TabIndex        =   99
         Top             =   480
         Width           =   9975
         Begin VB.TextBox txtReplyEmail 
            Height          =   375
            Left            =   240
            TabIndex        =   168
            Top             =   3960
            Width           =   8655
         End
         Begin VB.TextBox txtVendor 
            Height          =   375
            Left            =   960
            TabIndex        =   108
            Top             =   360
            Width           =   7935
         End
         Begin VB.TextBox txtAddress1 
            Height          =   375
            Left            =   960
            TabIndex        =   107
            Top             =   840
            Width           =   7935
         End
         Begin VB.TextBox txtAddress2 
            Height          =   375
            Left            =   960
            TabIndex        =   106
            Top             =   1320
            Width           =   7935
         End
         Begin VB.TextBox txtCity 
            Height          =   375
            Left            =   960
            TabIndex        =   105
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtCountry1 
            Height          =   375
            Left            =   3840
            TabIndex        =   104
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   960
            TabIndex        =   103
            Top             =   2280
            Width           =   7935
         End
         Begin VB.TextBox txtFaxNo 
            Height          =   375
            Left            =   960
            TabIndex        =   102
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox txtCreditTerms 
            Height          =   420
            Left            =   4200
            TabIndex        =   101
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtTel 
            Height          =   375
            Left            =   960
            TabIndex        =   100
            Top             =   3240
            Width           =   1935
         End
         Begin VB.Label Label13 
            Caption         =   "Reply Email in EO (Only 1 email address is allowed)"
            Height          =   375
            Left            =   240
            TabIndex        =   169
            Top             =   3720
            Width           =   4095
         End
         Begin VB.Label Label15 
            Caption         =   "Vendor "
            Height          =   375
            Left            =   240
            TabIndex        =   116
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Address"
            Height          =   375
            Left            =   240
            TabIndex        =   115
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "City"
            Height          =   375
            Left            =   240
            TabIndex        =   114
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Country"
            Height          =   375
            Left            =   3120
            TabIndex        =   113
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Email (;)"
            Height          =   375
            Left            =   240
            TabIndex        =   112
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Fax No (,)"
            Height          =   375
            Left            =   240
            TabIndex        =   111
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Credit Terms"
            Height          =   255
            Left            =   3000
            TabIndex        =   110
            Top             =   2760
            Width           =   1065
         End
         Begin VB.Label Label22 
            Caption         =   "Contact No."
            Height          =   375
            Left            =   240
            TabIndex        =   109
            Top             =   3240
            Width           =   615
         End
      End
      Begin VB.ComboBox cmbDispNum 
         Height          =   315
         Left            =   -72280
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Frame11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -67800
         TabIndex        =   94
         Top             =   4260
         Width           =   3555
         Begin VB.TextBox txtPassengerID 
            Height          =   315
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   95
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Passenger ID:"
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
            Left            =   360
            TabIndex        =   96
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.CheckBox chkformula 
         Caption         =   "Apply Formula?"
         Height          =   255
         Left            =   -70800
         TabIndex        =   93
         Top             =   600
         Width           =   1455
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
         Left            =   -73560
         TabIndex        =   88
         Top             =   5520
         Width           =   2055
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
         ItemData        =   "frmOSAirTkt.frx":008C
         Left            =   -73200
         List            =   "frmOSAirTkt.frx":008E
         TabIndex        =   79
         Text            =   "cboClientType"
         Top             =   390
         Width           =   735
      End
      Begin VB.Frame fraMI 
         Height          =   3840
         Left            =   -74520
         TabIndex        =   63
         Top             =   1320
         Width           =   4695
         Begin VB.ComboBox cboBookingAction 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   173
            Top             =   3000
            Width           =   2390
         End
         Begin VB.ComboBox cboTrip 
            Height          =   315
            ItemData        =   "frmOSAirTkt.frx":0090
            Left            =   3720
            List            =   "frmOSAirTkt.frx":0092
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   3720
            Visible         =   0   'False
            Width           =   1695
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
            Left            =   2220
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   2565
            Width           =   1095
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
            TabIndex        =   78
            Tag             =   "BY-"
            Top             =   2160
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
            Index           =   1
            Left            =   2220
            TabIndex        =   70
            Tag             =   "BY-"
            Top             =   600
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
            Index           =   0
            Left            =   2220
            TabIndex        =   69
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
            Index           =   3
            Left            =   2220
            MaxLength       =   3
            TabIndex        =   68
            Tag             =   "BY-"
            Top             =   1080
            Width           =   885
         End
         Begin VB.ComboBox cboClassServ 
            Height          =   315
            Left            =   360
            TabIndex        =   66
            Text            =   "cboClassServ"
            Top             =   1800
            Width           =   4215
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
            TabIndex        =   174
            Top             =   3000
            Width           =   1875
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
            Left            =   600
            TabIndex        =   91
            Top             =   2565
            Width           =   1515
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
            TabIndex        =   76
            Top             =   600
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
            TabIndex        =   75
            Top             =   180
            Width           =   1875
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
            TabIndex        =   74
            Top             =   1080
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
            Left            =   240
            TabIndex        =   73
            Top             =   1440
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
            Left            =   2760
            TabIndex        =   72
            Top             =   3720
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
            TabIndex        =   71
            Top             =   2160
            Width           =   1875
         End
      End
      Begin VB.TextBox txtContact 
         Height          =   285
         Left            =   -72720
         MaxLength       =   50
         TabIndex        =   61
         Top             =   1080
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   47
         Top             =   4920
         Width           =   5955
         Begin VB.TextBox txtFuelSurcharge 
            Height          =   315
            Left            =   2160
            TabIndex        =   170
            Text            =   "0"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkTFNRCC 
            Caption         =   "TF in NRCC(If applicable)"
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
            Left            =   3360
            TabIndex        =   165
            Top             =   240
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Calc&ulate"
            Height          =   255
            Left            =   3420
            TabIndex        =   18
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Calc&ulate"
            Height          =   255
            Left            =   4680
            TabIndex        =   16
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtTrxnFee 
            Height          =   315
            Left            =   2160
            TabIndex        =   17
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtDiscount 
            Height          =   315
            Left            =   2160
            TabIndex        =   15
            Text            =   "0"
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblFuelSurcharge 
            Alignment       =   1  'Right Justify
            Caption         =   "Fuel Surcharge:"
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
            Left            =   120
            TabIndex        =   171
            Top             =   600
            Width           =   1965
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Tranx/Service Fee:"
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
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   1965
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Ticket Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   -67800
         TabIndex        =   46
         Top             =   3310
         Width           =   3555
         Begin VB.TextBox txtConjTkt 
            Height          =   315
            Left            =   2760
            MaxLength       =   2
            TabIndex        =   167
            Top             =   480
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtTktNum 
            Height          =   315
            Left            =   840
            MaxLength       =   10
            TabIndex        =   26
            Top             =   480
            Width           =   1755
         End
         Begin VB.TextBox txtALCode 
            Height          =   315
            Left            =   120
            MaxLength       =   3
            TabIndex        =   25
            Top             =   480
            Width           =   675
         End
         Begin VB.Line Line1 
            Visible         =   0   'False
            X1              =   2620
            X2              =   2700
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblAirlineCodeRem 
            Caption         =   "Please enter airline numeric code"
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   240
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   -68400
         TabIndex        =   43
         Top             =   5040
         Width           =   4155
         Begin VB.TextBox txtEONum 
            Height          =   315
            Left            =   2100
            TabIndex        =   28
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton cmdEO 
            Caption         =   "E&xchange Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   240
            Picture         =   "frmOSAirTkt.frx":0094
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2100
            TabIndex        =   30
            Top             =   1260
            Width           =   1275
         End
         Begin VB.CommandButton cmdDone 
            Caption         =   "&Done"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   29
            Top             =   1260
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   37
         Top             =   1320
         Width           =   5355
         Begin VB.TextBox txtSellingPrice 
            Height          =   315
            Left            =   2340
            TabIndex        =   52
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdCalculate 
            Caption         =   "Calc&ulate"
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
            Left            =   3600
            TabIndex        =   13
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CheckBox chkAbsorb 
            Caption         =   "CWT Absorb"
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
            Left            =   3600
            TabIndex        =   57
            Top             =   2760
            Width           =   1455
         End
         Begin VB.CheckBox chkNRCC 
            Caption         =   "UATP"
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
            Left            =   3600
            TabIndex        =   89
            Top             =   3000
            Width           =   975
         End
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   3600
            TabIndex        =   83
            Top             =   2400
            Width           =   1095
            Begin VB.OptionButton optDiscount 
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
               Height          =   255
               Index           =   1
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   85
               Top             =   0
               Width           =   375
            End
            Begin VB.OptionButton optDiscount 
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
               Height          =   255
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   3600
            TabIndex        =   82
            Top             =   1920
            Width           =   975
            Begin VB.OptionButton optCommType 
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
               Height          =   255
               Index           =   1
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   87
               Top             =   120
               Width           =   375
            End
            Begin VB.OptionButton optCommType 
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
               Height          =   255
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   120
               Width           =   375
            End
         End
         Begin VB.TextBox txtDC 
            Height          =   315
            Left            =   2340
            TabIndex        =   65
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox txtGrossFare 
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtTaxCode 
            Height          =   315
            Index           =   1
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   60
            Top             =   1320
            Width           =   555
         End
         Begin VB.TextBox txtTax 
            Height          =   315
            Index           =   1
            Left            =   2340
            TabIndex        =   58
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtTaxCode 
            Height          =   315
            Index           =   0
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   56
            Top             =   960
            Width           =   555
         End
         Begin VB.TextBox txtTax 
            Height          =   315
            Index           =   0
            Left            =   2340
            TabIndex        =   55
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtMerchFee 
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtCommission 
            Height          =   315
            Left            =   2340
            TabIndex        =   64
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtSellPrice 
            Height          =   315
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox txtPubFare 
            Height          =   315
            Left            =   2340
            TabIndex        =   51
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtNettFare 
            Height          =   315
            Left            =   2340
            TabIndex        =   50
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Discount:"
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
            Left            =   120
            TabIndex        =   81
            Top             =   2400
            Width           =   2145
         End
         Begin VB.Label lblFare 
            Alignment       =   1  'Right Justify
            Caption         =   "Nett Fare:"
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
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   2145
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Tax/Tax Code:"
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
            Left            =   720
            TabIndex        =   45
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Tax/Tax Code:"
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
            Left            =   720
            TabIndex        =   44
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Merchant Fee:"
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
            Left            =   720
            TabIndex        =   41
            Top             =   2760
            Width           =   1545
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Commission:"
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
            Left            =   120
            TabIndex        =   40
            Top             =   2040
            Width           =   2145
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nett Cost in EO:"
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
            Left            =   360
            TabIndex        =   39
            Top             =   1680
            Width           =   1905
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Selling Fare:"
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
            Left            =   120
            TabIndex        =   38
            Top             =   3120
            Width           =   2145
         End
         Begin VB.Label lblSellingPrice 
            Alignment       =   1  'Right Justify
            Caption         =   "Selling Price:"
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
            Left            =   720
            TabIndex        =   156
            Top             =   600
            Width           =   1545
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Gross Fare:"
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
            Left            =   720
            TabIndex        =   54
            Top             =   600
            Width           =   1545
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Form of Payment (FOP)"
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
         Left            =   -74760
         TabIndex        =   36
         Top             =   5880
         Width           =   5895
         Begin VB.CheckBox chkWaiveMercFee 
            Caption         =   "Waive Merchant Fee"
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
            Left            =   2640
            TabIndex        =   23
            Top             =   720
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtCCNum 
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
            Left            =   2580
            MaxLength       =   18
            TabIndex        =   21
            Tag             =   "NN"
            Top             =   300
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.ComboBox cmbFOPType 
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
            ItemData        =   "frmOSAirTkt.frx":04D6
            Left            =   240
            List            =   "frmOSAirTkt.frx":04E0
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   300
            Width           =   1515
         End
         Begin VB.ComboBox cmbCCType 
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
            ItemData        =   "frmOSAirTkt.frx":04ED
            Left            =   1800
            List            =   "frmOSAirTkt.frx":0509
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   300
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpCCExp 
            Height          =   360
            Left            =   4680
            TabIndex        =   22
            Top             =   300
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
            Format          =   61669379
            CurrentDate     =   36526
            MaxDate         =   73050
            MinDate         =   36526
         End
      End
      Begin VB.Frame fraHotel 
         Caption         =   "Select Air Segments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   -69120
         TabIndex        =   35
         Top             =   480
         Width           =   4935
         Begin VB.CheckBox chkSelectAll 
            Caption         =   "Select All Segments"
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   1680
            Width           =   2655
         End
         Begin VB.ListBox lstFlights 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   24
            Top             =   300
            Width           =   4695
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   1455
         Left            =   -69120
         Picture         =   "frmOSAirTkt.frx":052D
         ScaleHeight     =   1395
         ScaleWidth      =   1215
         TabIndex        =   34
         Top             =   3500
         Width           =   1275
      End
      Begin VB.TextBox txtFreeRmk 
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
         Left            =   -72240
         MaxLength       =   70
         TabIndex        =   5
         Tag             =   "NN"
         Top             =   3330
         Width           =   6165
      End
      Begin VB.CommandButton cmdFreeRmkToEO 
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
         Left            =   -65940
         Picture         =   "frmOSAirTkt.frx":716F
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add Free Text to Exchange Order Remarks"
         Top             =   3210
         Width           =   495
      End
      Begin VB.CommandButton cmdFreeRmkToItin 
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
         Left            =   -65400
         Picture         =   "frmOSAirTkt.frx":75B1
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Add Free Text to Itinerary Remarks"
         Top             =   3210
         Width           =   495
      End
      Begin VB.Frame Frame3 
         Caption         =   "Associated Itinerary Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   -74880
         TabIndex        =   33
         Top             =   3840
         Width           =   10275
         Begin VB.ListBox lstItinRmks 
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   300
            Width           =   4755
         End
         Begin VB.ListBox lstItinRmks 
            Height          =   2205
            Index           =   1
            Left            =   5400
            MultiSelect     =   1  'Simple
            TabIndex        =   12
            Top             =   300
            Width           =   4755
         End
         Begin VB.CommandButton cmdItinRmksAddAll 
            Caption         =   "ALL"
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
            Left            =   4860
            Picture         =   "frmOSAirTkt.frx":79F3
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Add All Remarks"
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton cmdItinRmksAddOne 
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
            Left            =   4860
            Picture         =   "frmOSAirTkt.frx":7E35
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Add Selected Remark"
            Top             =   300
            Width           =   495
         End
         Begin VB.CommandButton cmdItinRmksRemove 
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
            Left            =   4860
            Picture         =   "frmOSAirTkt.frx":8277
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Remove Selected Remark"
            Top             =   1860
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Exchange Order Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   10275
         Begin VB.CommandButton cmdEORmksRemove 
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
            Left            =   4860
            Picture         =   "frmOSAirTkt.frx":86B9
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Remove Selected Remark"
            Top             =   1860
            Width           =   495
         End
         Begin VB.CommandButton cmdEORmksAddOne 
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
            Left            =   4860
            Picture         =   "frmOSAirTkt.frx":8AFB
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Add Selected Remark"
            Top             =   300
            Width           =   495
         End
         Begin VB.CommandButton cmdEORmksAddAll 
            Caption         =   "ALL"
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
            Left            =   4860
            Picture         =   "frmOSAirTkt.frx":8F3D
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Add All Remarks"
            Top             =   1080
            Width           =   495
         End
         Begin VB.ListBox lstEORmks 
            Height          =   2205
            Index           =   1
            Left            =   5400
            MultiSelect     =   1  'Simple
            TabIndex        =   4
            Top             =   240
            Width           =   4755
         End
         Begin VB.ListBox lstEORmks 
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   0
            Top             =   300
            Width           =   4755
         End
      End
      Begin VB.Frame fraMCO 
         Caption         =   "MCO Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   -74880
         TabIndex        =   128
         Top             =   840
         Width           =   5535
         Begin VB.TextBox txtOrginalPOI 
            Height          =   315
            Left            =   2400
            TabIndex        =   154
            Top             =   4680
            Width           =   1095
         End
         Begin VB.TextBox txtOrginalFOP 
            Height          =   315
            Left            =   2400
            TabIndex        =   153
            Top             =   4320
            Width           =   1935
         End
         Begin VB.TextBox txtConjunction 
            Height          =   315
            Left            =   2400
            TabIndex        =   152
            Top             =   3960
            Width           =   3015
         End
         Begin VB.TextBox txtExchangeFor 
            Height          =   315
            Left            =   2400
            TabIndex        =   151
            Top             =   3600
            Width           =   3015
         End
         Begin VB.TextBox txtMCOTaxes 
            Height          =   315
            Left            =   2400
            TabIndex        =   150
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox txtHeadlineCurrency 
            Height          =   315
            Left            =   2400
            TabIndex        =   149
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtROE 
            Height          =   315
            Left            =   2400
            TabIndex        =   148
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox txtEquiAmt 
            Height          =   315
            Left            =   2400
            TabIndex        =   147
            Top             =   2160
            Width           =   1095
         End
         Begin VB.TextBox txtFOP 
            Height          =   315
            Left            =   2400
            TabIndex        =   146
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtContactPerson 
            Height          =   315
            Left            =   2400
            TabIndex        =   145
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtLoc 
            Height          =   315
            Left            =   2400
            TabIndex        =   144
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtRecLoc 
            Height          =   285
            Left            =   2400
            TabIndex        =   130
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtTypeOfService 
            Height          =   315
            Left            =   2400
            TabIndex        =   129
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "Original Place of Issue:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   143
            Top             =   4680
            Width           =   2175
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Original FOP:"
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
            Left            =   720
            TabIndex        =   142
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "In Conjunction with:"
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
            TabIndex        =   141
            Top             =   3960
            Width           =   2055
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Issue in Exchange For:"
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
            Left            =   120
            TabIndex        =   140
            Top             =   3600
            Width           =   2175
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Taxes:"
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
            Left            =   1560
            TabIndex        =   139
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "FOP:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   138
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Equivalent Amt Paid:"
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
            TabIndex        =   137
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Rate of Exchange:"
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
            Left            =   360
            TabIndex        =   136
            Top             =   3240
            Width           =   1935
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Headline Currency:"
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
            Left            =   360
            TabIndex        =   135
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact:"
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
            Left            =   1440
            TabIndex        =   134
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Location of Issuance:"
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
            Left            =   120
            TabIndex        =   133
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Record Locator:"
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
            Left            =   600
            TabIndex        =   132
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Type of Service:"
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
            Left            =   720
            TabIndex        =   131
            Top             =   720
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView lvwRealECodes 
         Height          =   1815
         Left            =   -69720
         TabIndex        =   161
         Top             =   1080
         Width           =   5115
         _ExtentX        =   9022
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
         Left            =   -69720
         TabIndex        =   162
         Top             =   3360
         Width           =   5115
         _ExtentX        =   9022
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
      Begin MSComctlLib.ListView lvwECodes 
         Height          =   4275
         Left            =   -69840
         TabIndex        =   77
         Top             =   6000
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
      Begin VB.Label lblOriTktNum 
         Caption         =   "Issue in Exchange For:"
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
         Left            =   -69000
         TabIndex        =   180
         Top             =   2850
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblTktLength 
         Caption         =   "(min/max char 10 numeric)"
         Height          =   255
         Left            =   -69000
         TabIndex        =   179
         Top             =   3105
         Visible         =   0   'False
         Width           =   2175
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
         Left            =   -69600
         TabIndex        =   164
         Top             =   3000
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
         Left            =   -69720
         TabIndex        =   163
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Filed Fare Number:"
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
         Left            =   -74280
         TabIndex        =   157
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label lblCT 
         Caption         =   "Client Type"
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
         Left            =   -74400
         TabIndex        =   80
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label lblContact 
         Caption         =   "Vendor Contact Person"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   59
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Free Text:"
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
         Left            =   -73860
         TabIndex        =   42
         Top             =   3390
         Width           =   1545
      End
   End
   Begin MSAdodcLib.Adodc datVendors 
      Height          =   375
      Left            =   4440
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dbcVendors 
      Bindings        =   "frmOSAirTkt.frx":937F
      DataSource      =   "datVendors"
      Height          =   360
      Left            =   240
      TabIndex        =   90
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   "Description"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOSAirTkt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsRmks As New ADODB.Recordset
Dim mblnConsTkt As Boolean
Dim msngMarkup As Single
Dim msngIntlDiscount As Single
Dim msngMerchFeePct As Single
'Timer
'Dim StartTime As Date
Dim mstrPCAmend As String
Dim mstrTktNum As String
Dim mRF() As String
Dim mLF() As String
'CS Change EC
'Dim mEC() As String
Dim mRS() As String
Dim mMS() As String
Dim mFF7() As String
Dim mFF8() As String
'CS Remove FF26
'Dim mFF26() As String
'CS Add FF41
'Dim mFF41() As String
Dim mFF81() As String
Dim mFF38() As String
Dim mFF34() As String
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date
Dim mstrBookingTool As String
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Dim mobjWebFares As New WebFares
Dim mobjDeclinedWebFares As New WebFares
Dim mbolWebFareSelected As Boolean
Dim mintWebFareSelected As Integer
Dim mbolEORetrieving As Boolean

Private Sub cboClientType_Click()
If chkformula.value = 1 Then
    If cboClientType = "TF" Then
        txtTrxnFee.Enabled = True
        
    Else
        txtTrxnFee.Enabled = False
        
    End If
    If cboClientType = "MN" Then
       txtCommission = 0
       txtDC = 0
    End If

End If
End Sub



Private Sub chkAbsorb_Click()
   EnableCalculate
   If chkAbsorb.value = 1 Then
      
   Else
      txtMerchFee = 0
   End If
End Sub

'02062005
Private Sub chkAmdTktOnly_Click()
   If chkAmdTktOnly.value = 1 Then
      cmdEO.Enabled = False
      cmdDone.Enabled = True
      'chkTktAtInv.Enabled = False
   Else
      cmdEO.Enabled = True
      'chkTktAtInv.Enabled = True
   End If
End Sub

Private Sub chkformula_Click()
EnableCalculate
txtGrossFare = ""
txtCost = ""
txtMerchFee = ""
txtSellPrice = ""
If chkformula Then
    optCommType(1).value = True
    optDiscount(1).value = True
    optCommType(0).value = False
    optDiscount(0).value = False
    optCommType(1).Enabled = True
    optDiscount(1).Enabled = True
    txtCost.Locked = True
    'If UCase(gstrAgcyCountryCode) = "HK" Then txtGrossFare.Locked = True
    txtMerchFee.Locked = True
    chkAbsorb.Enabled = True
    If UCase(gstrAgcyCountryCode) = "HK" Then
        If cboClientType.Text = "TF" Then
            txtTrxnFee.Enabled = True
        Else
            txtTrxnFee.Enabled = False
        End If
    End If
Else
    optCommType(1).value = False
    optDiscount(1).value = False
    optCommType(0).value = True
    optDiscount(0).value = True
    optCommType(1).Enabled = False
    optDiscount(1).Enabled = False
    
    'If UCase(gstrAgcyCountryCode) = "HK" Then
        txtCost.Locked = False
        txtCost = ""
    'Else
    '    txtCost.Locked = True
    'End If
    'If UCase(gstrAgcyCountryCode) = "HK" Then txtGrossFare.Locked = False
    txtMerchFee.Locked = False
    chkAbsorb.Enabled = False
    txtTrxnFee.Enabled = True
End If
End Sub

Private Sub chkNRCC_Click()
EnableCalculate
If Not mblnConsTkt Then
   If chkNRCC.value = 1 Then
      cmbFOPType.Text = "CC"
      'Added on 130807: V46 To disable "cwt absorb" once UATP is selected
      chkAbsorb.Enabled = False
      chkAbsorb.value = 0
   Else
      cmbFOPType.Text = "CX"
      chkAbsorb.Enabled = True
   End If
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

Private Sub chkReissuedTkt_Click()
  If chkReissuedTkt.value = vbUnchecked Then
     txtOriTktNum.Text = ""
     txtOriTktNum.Enabled = False
  Else
     txtOriTktNum.Enabled = True
  End If
End Sub

Private Sub chkRequestMCO_Click()
If chkRequestMCO Then
    fraTrav.Enabled = True
    fraMCO.Enabled = True
    fraRmk.Enabled = True
Else
    fraTrav.Enabled = False
    fraMCO.Enabled = False
    fraRmk.Enabled = False
End If
End Sub

Private Sub chkSelectAll_Click()

Dim i As Integer
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Dim intWebFareMatch As Integer

If chkSelectAll.value = 1 Then

For i = 0 To lstFlights.ListCount - 1
    lstFlights.Selected(i) = True
Next
End If

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
If mblnConsTkt Then
   intWebFareMatch = checkWebFareSelected
   If intWebFareMatch > 0 Then
      populateWebFare (intWebFareMatch)
   Else
      txtALCode.Enabled = True
   End If
End If

End Sub

'Private Sub chkTktAtInv_Click()
'If chkTktAtInv.value = vbChecked Then
    'txtALCode.Enabled = False
'    txtTktNum.Enabled = False
    'txtALCode.Text = ""
'    txtTktNum.Text = ""

'Else
    'txtALCode.Enabled = True
'    txtTktNum.Enabled = True
'End If

'End Sub

Private Sub cmbDispNum_Click()
   Dim intI As Integer
   
   If cmbDispNum.Text = "" Then
      txtMI(0) = ""
      txtMI(1) = ""
      'CS Change EC
      'txtMI(2) = ""
      txtRS = ""
      txtMS = ""
      txtMI(3) = ""
      txtMI(6) = ""
      cboClassServ.listindex = 0
      'CS Remove FF26
      'cboTripType.listindex = 0
      'CS Add FF41
      'cboTrip.listindex = 0
      chkPaperTkt.value = 0
   Else
       txtMI(0) = mRF(cmbDispNum.listindex - 1)
       txtMI(1) = mLF(cmbDispNum.listindex - 1)
       'CS Change EC
       'txtMI(2) = mEC(cmbDispNum.listindex - 1)
       txtRS = mRS(cmbDispNum.listindex - 1)
       txtMS = mMS(cmbDispNum.listindex - 1)
       txtMI(3) = mFF7(cmbDispNum.listindex - 1)
       cboClassServ.listindex = 0
       cboClassServ = matchList(mFF8(cmbDispNum.listindex - 1))
       cboBookingAction = matchBookingTool(mFF34(cmbDispNum.listindex - 1))
       'For intI = 1 To cboMIFareType.ListCount - 1
       '    If (Mid(cboMIFareType.List(intI), 1, InStr(1, cboMIFareType.List(intI), "-") - 1)) = mFF8(cmbDispNum.listindex - 1) Then
       '       cboMIFareType.listindex = intI
       '    End If
       'Next
       txtMI(6) = mFF81(cmbDispNum.listindex - 1)
       'CS Remove FF26
       'cboTripType.listindex = 0
       'If UCase(mFF26(cmbDispNum.listindex - 1)) = "R" Then
       '   cboTripType.Text = "Round"
       'ElseIf UCase(mFF26(cmbDispNum.listindex - 1)) = "O" Then
       '   cboTripType.Text = "One Way"
       'End If
       
       'CS Add FF41
       'cboTrip.listindex = 0
       'If UCase(mFF41(cmbDispNum.listindex - 1)) = "I" Then
       '   cboTrip.Text = "INTERNATIONAL"
       'ElseIf UCase(mFF41(cmbDispNum.listindex - 1)) = "D" Then
       '   cboTrip.Text = "DOMESTIC"
       'End If
       If UCase(mFF38(cmbDispNum.listindex - 1)) = "P" Then
          chkPaperTkt.value = 1
       Else
          chkPaperTkt.value = 0
       End If
   End If
End Sub

Private Sub cmbFOPType_Click()
Dim blnCC As Boolean
    
EnableCalculate
blnCC = (cmbFOPType = "CX") Or (frmOSAirTkt.cmbFOPType = "CC")
cmbCCType.Visible = blnCC
txtCCNum.Visible = blnCC
dtpCCExp.Visible = blnCC
'chkWaiveMercFee.Visible = blnCC

End Sub



Private Sub cmdAddMCORmk_Click()
If txtFT <> "" Then
    lstRmks.AddItem txtFT
    txtFT.Text = ""
End If
End Sub

Private Sub cmdAddTraPax_Click()
Dim item As ListItem

   
   Set item = lsvTraveller.ListItems.Add(, , cmbTraName.Text)
   
End Sub

Private Sub cmdCalculate_Click()
Dim sngSF As Single
Dim sngComm As Single
Dim sngCommPct As Single
Dim sngDC As Single
Dim strCNType As String
Dim sngMFTotal As Single
Dim strMsg As String



'Modified on 5/1/2005: Getting company information by Profile Name instead on CN
'strCNType = fGetCNType(gobjPNR.CN)
strCNType = fGetCNType(gobjPNR.CompInfo.ProfileName)

If UCase(gstrAgcyCountryCode) = "HK" Then
    If txtNettFare.Text = "" Then
        'MsgBox "Need " & Replace(lblFare.Caption, ":", ""), vbApplicationModal + vbExclamation
         strMsg = "Need " & Replace(lblFare.Caption, ":", "")
         modMsgBox.OKMsg = "OK"
         modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        Exit Sub
    End If
ElseIf UCase(gstrAgcyCountryCode) = "SG" Then
    If txtPubFare.Text = "" Then
        'MsgBox "Need " & Replace(lblFare.Caption, ":", ""), vbApplicationModal + vbExclamation
        strMsg = "Need " & Replace(lblFare.Caption, ":", "")
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        Exit Sub
    End If
    If txtSellingPrice.Visible = True And txtSellingPrice = "" Then
        'MsgBox "Need Selling Price", vbApplicationModal + vbExclamation
        strMsg = "Need Selling Price"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        Exit Sub
    End If
End If



If UCase(gstrAgcyCountryCode) = "HK" Then
    HK_Calculate

ElseIf UCase(gstrAgcyCountryCode) = "SG" Then
    If chkformula.value = 0 Then
    'Modified on 22/06/05: remove commission in selling price

    'txtSellPrice = fCurrRound(IIf(mblnConsTkt = True, fConvertZero(txtPubFare), fConvertZero(txtSellingPrice)) - fConvertZero(txtDC) + fConvertZero(txtTax(0)) + fConvertZero(txtTax(1)) + fConvertZero(txtMerchFee) - IIf(mblnConsTkt = True, 0, 0), gstrAgcyCurrCode, "UP")
    'txtCost = fCurrRound(fConvertZero(txtPubFare) - fConvertZero(txtCommission), gstrAgcyCurrCode, "UP")
    'JY - 20100503 - Remove the logic to round up txtSellPrice and txtCost
    txtSellPrice = IIf(mblnConsTkt = True, fConvertZero(txtPubFare), fConvertZero(txtSellingPrice)) - fConvertZero(txtDC) + fConvertZero(txtTax(0)) + fConvertZero(txtTax(1)) + fConvertZero(txtMerchFee) - IIf(mblnConsTkt = True, 0, 0)
    txtCost = fConvertZero(txtPubFare) - fConvertZero(txtCommission)
    
    Else
    sngSF = fConvertZero(txtPubFare.Text)

    If optCommType(1).value = True Then
    sngCommPct = fConvertZero(txtCommission.Text)
    sngComm = fCommAmt(sngSF, sngCommPct, strCNType)
    optCommType(0).value = True
    txtCommission.Text = sngComm
    End If
    
    sngDC = fConvertZero(txtDC)
    
    If optDiscount(1).value Then
       sngDC = fDiscAmt(sngSF, sngDC, sngComm, strCNType)
       txtDC = sngDC
       optDiscount(0).value = True
    End If
    
    
    'If txtCommission.Text = "" Or fConvertZero(txtCommission) = 0 Or strCNType = "MN" Or strCNType = "TF" Or strCNType = "MG" Then
    'modified on 6/7/2005: open commission entry regardless of client type
    If txtCommission.Text = "" Or fConvertZero(txtCommission) = 0 Then
       txtCommission.Text = "0"
       optCommType(0).value = True
    End If
    'If txtDiscount.Text = "" Or strCNType = "MN" Or strCNType = "TF" Or strCNType = "MG" Then
    'modified on 21/9: Don't offset any Discount value
    If txtDiscount.Text = "" Then
    txtDC.Text = "0"
    optDiscount(0).value = True
    End If
    
    
    If txtCost.Text = "" Then txtCost.Text = "0"
    If txtTax(0).Text = "" Then txtTax(0).Text = "0"
    If txtTax(1).Text = "" Then txtTax(1).Text = "0"
    If txtTrxnFee = "" Then txtTrxnFee = "0"
    
    'txtCost = fCurrRound(sngSF - CSng(txtCommission) + txtTax(0) + txtTax(1), gstrAgcyCurrCode, "UP")
    'txtCost = fCurrRound(sngSF - fConvertZero(txtCommission), gstrAgcyCurrCode, "UP")
    'JY - 20100503 - Remove the logic to round up txtCost
    txtCost = sngSF - fConvertZero(txtCommission)
    'sngSF = CSng(txtCost) + CSng(txtCommission) - CSng(txtDC)
    sngSF = IIf(mblnConsTkt = True, fConvertZero(txtCost), fConvertZero(txtSellingPrice)) + IIf(mblnConsTkt = True, fConvertZero(txtCommission), 0) - fConvertZero(txtDC) + fConvertZero(txtTax(0)) + fConvertZero(txtTax(1))

    If chkAbsorb.value = 0 Then
       If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.value = vbUnchecked Then
          sngMFTotal = fMFTotal(False, sngSF, txtTrxnFee, fConvertZero(txtPubFare.Text), fConvertZero(txtTax(0)) + fConvertZero(txtTax(1)))
          Me.txtMerchFee = fMerchantFee(sngMFTotal, msngMerchFeePct)
       Else
          Me.txtMerchFee.Text = "0"
       End If
    Else
       Me.txtMerchFee.Text = "0"
    End If
    'txtSellPrice = fCurrRound(txtPubFare + txtTax(0) + txtTax(1) - txtCommission.Text + txtMerchFee + txtTrxnFee, gstrAgcyCurrCode, "UP")
    'txtSellPrice = fCurrRound(sngSF + txtMerchFee, gstrAgcyCurrCode, "UP")
    'JY - 20100503 - Remove the logic to round up txtCost
    txtSellPrice = sngSF + CSng(txtMerchFee)
    'sngSF = sngSF + CSng(txtCommission.Text)
    'If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.Value = vbUnchecked Then
    '    Me.txtMerchFee = fMerchantFee(sngSF, 2.5)
    'Else
    '    Me.txtMerchFee.Text = "0"
    'End If
    'sngSF = sngSF + CSng(txtMerchFee.Text)
    'txtSellPrice = sngSF
    End If
    
End If
cmdCalculate.Enabled = False

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
If mblnConsTkt Then populateLowFare

End Sub

Private Sub HK_Calculate()
Dim sngSF As Single
Dim sngGrossFare As Single
Dim strCNType As String
Dim sngComm As Single
Dim sngCommPct As Single
Dim sngDiscountPct As Single
Dim sngDC As Single
Dim sngMFTotal As Single

'If txtNettFare.Text = "" Then
'    MsgBox "Need Nett Fare", vbApplicationModal + vbExclamation
'    Exit Sub
'End If

If chkformula.value = 0 Then
    'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
    If mbolWebFareSelected = True Then
        txtDC = fCurrRound(fConvertZero(txtDC), gstrAgcyCurrCode, "DOWN")
        txtSellPrice = fConvertZero(txtNettFare) + fConvertZero(txtCommission) - fConvertZero(txtDC) + fConvertZero(txtTax(0)) + fConvertZero(txtTax(1)) + fConvertZero(txtMerchFee)
        txtCost = fConvertZero(txtNettFare)
    Else
        txtDC = fCurrRound(fConvertZero(txtDC), gstrAgcyCurrCode, "DOWN")
        txtSellPrice = fCurrRound(fConvertZero(txtNettFare) + fConvertZero(txtCommission) - fConvertZero(txtDC) + fConvertZero(txtTax(0)) + fConvertZero(txtTax(1)) + fConvertZero(txtMerchFee), gstrAgcyCurrCode, "UP")
        txtCost = fCurrRound(fConvertZero(txtNettFare), gstrAgcyCurrCode, "UP")
    End If
Else

strCNType = cboClientType.Text
sngSF = fConvertZero(txtNettFare.Text)

'COMPUTE GROSS
'Remove restriction for TF and MN by JiYong
'If UCase(strCNType) = "TF" Or UCase(strCNType) = "MN" Then
'   sngGrossFare = 0
'Else
    'sngGrossFare = fCurrRound((sngSF / 0.88) + 10, gstrAgcyCurrCode, "UP")
    If optCommType(1).value = True Then
        sngCommPct = fConvertZero(txtCommission.Text)
        sngComm = fCommAmt(sngSF, sngCommPct)
        optCommType(0).value = True
        'Added TF and MN by JiYong
        'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
        If mbolWebFareSelected = True Then
            If UCase(strCNType) = "MG" Or UCase(strCNType) = "DB" Or UCase(strCNType) = "TF" Or UCase(strCNType) = "MN" Then
                sngGrossFare = Format(sngSF / (1 - (sngCommPct * 0.01)), gstrHKWebFareDecimalPoint)
            Else
                sngGrossFare = Format(sngSF / (1 - (sngCommPct * 0.01)) + 10, gstrHKWebFareDecimalPoint)
            End If
        Else
            If UCase(strCNType) = "MG" Or UCase(strCNType) = "DB" Or UCase(strCNType) = "TF" Or UCase(strCNType) = "MN" Then
                sngGrossFare = fCurrRound(sngSF / (1 - (sngCommPct * 0.01)), gstrAgcyCurrCode, "UP")
            Else
                sngGrossFare = fCurrRound(sngSF / (1 - (sngCommPct * 0.01)) + 10, gstrAgcyCurrCode, "UP")
            End If
        End If
    Else
       sngComm = txtCommission.Text
       'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
       If mbolWebFareSelected = True Then
            sngGrossFare = Format(sngSF + fConvertZero(txtCommission.Text), gstrHKWebFareDecimalPoint)
       Else
            sngGrossFare = fCurrRound(sngSF + fConvertZero(txtCommission.Text), gstrAgcyCurrCode, "UP")
       End If
    End If


'End If
txtGrossFare = sngGrossFare
txtCommission.Text = Format(sngComm, "0.00")

sngDC = fConvertZero(txtDC.Text)

If optDiscount(1).value = True Then
    'sngDC = fDiscAmt(sngSF, sngDiscountPct, sngComm)
    'sngDiscountPct = fConvertZero(txtDC.Text)
    sngDC = fDiscAmt(sngSF + sngComm, sngDC, sngComm)
    optDiscount(0).value = True
End If

txtDC = fCurrRound(sngDC, gstrAgcyCurrCode, "DOWN")

'Remove commission = 0 for TF and MN by JiYong
If txtCommission.Text = "" Or fConvertZero(txtCommission) = 0 Then 'Or strCNType = "MN" Or strCNType = "TF"
       txtCommission.Text = "0"
       optCommType(0).value = True
End If

If txtDiscount.Text = "" Or strCNType = "MN" Or strCNType = "TF" Then
        txtDC.Text = "0"
        optDiscount(0).value = True
End If
If txtCost.Text = "" Then txtCost.Text = "0"
If txtTax(0).Text = "" Then txtTax(0).Text = "0"
If txtTax(1).Text = "" Then txtTax(1).Text = "0"
If txtTrxnFee = "" Then txtTrxnFee = "0"
If txtDC = "" Then txtDC = "0"

'txtCost = sngSF + CSng(txtTax(0)) + CSng(txtTax(1))
txtCost = sngSF

If UCase(strCNType) = "TF" Or UCase(strCNType) = "MN" Then
   'Remove commission = 0 for TF and MN by Ji Yong and add in txtCommission in the calculation
   'txtCommission = 0
   txtDC = 0
   'sngSF = txtCost
   'Removed by JiYong to use sngGrossFare instead of txtCost
   'sngSF = txtCost + fConvertZero(txtTax(0)) + fConvertZero(txtTax(1))
   sngSF = sngGrossFare + fConvertZero(txtTax(0)) + fConvertZero(txtTax(1))
   
Else
   sngSF = sngGrossFare + fConvertZero(txtTax(0)) + fConvertZero(txtTax(1)) - fConvertZero(txtDC)
End If
If chkAbsorb.value = 0 Then
   'If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.Value = vbUnchecked Then
   If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.value = vbUnchecked Then
       sngMFTotal = fMFTotal(chkNRCC, sngSF, fConvertZero(txtTrxnFee), fConvertZero(txtNettFare.Text), fConvertZero(txtTax(0)) + fConvertZero(txtTax(1)))
       Me.txtMerchFee = fMerchantFee(sngMFTotal, msngMerchFeePct)
   Else
       Me.txtMerchFee.Text = "0"
   End If
Else
   Me.txtMerchFee.Text = "0"
End If
'txtSellPrice = fCurrRound(txtPubFare + txtTax(0) + txtTax(1) - txtCommission.Text + txtMerchFee + txtTrxnFee, gstrAgcyCurrCode, "UP")

'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
If mbolWebFareSelected = True Then
    txtSellPrice = Format(sngSF + txtMerchFee, gstrHKWebFareDecimalPoint)
Else
    txtSellPrice = fCurrRound(sngSF + txtMerchFee, gstrAgcyCurrCode, "UP")
End If

End If
'sngSF = sngSF + CSng(txtCommission.Text)
'If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.Value = vbUnchecked Then
'    Me.txtMerchFee = fMerchantFee(sngSF, 2.5)
'Else
'    Me.txtMerchFee.Text = "0"
'End If
'sngSF = sngSF + CSng(txtMerchFee.Text)
'txtSellPrice = sngSF
End Sub
Private Function fMFTotal(NRCC As Boolean, TotalCharge As Single, TransFee As Single, NetFare As Single, Tax As Single) As Single
gobjLog.ProcedureName = "fMFTotal"
On Error GoTo ProcError

If NRCC Then
    If UCase(gstrAgcyCountryCode) = "HK" Then
        If cboClientType = "TF" Then
           fMFTotal = TransFee
        Else
        
           fMFTotal = TotalCharge - NetFare - Tax
        
        End If
    End If
    
Else
    If cboClientType = "TF" Then
        If gobjPNR.CompInfo.TFIncMF Then
            fMFTotal = TotalCharge + TransFee
        Else
            fMFTotal = TotalCharge
        End If
    Else
        fMFTotal = TotalCharge
    End If
End If



Exit Function
ProcError:
    Call pErrorReport(True)

End Function


Private Sub cmdCancel_Click()
    Unload frmClientMI
    Unload Me
End Sub

Private Sub cmdClientMI_Click()
    Call loadClientMI
End Sub
Private Sub loadClientMI()
    If isLoaded("frmClientMI") Then
        frmClientMI.Show 'vbModal
    Else
        Load frmClientMI
        frmClientMI.intLocation = 4
        frmClientMI.intProdCode = frmOthSvcs.dbcProducts.BoundText
        frmClientMI.strPdtType = frmOthSvcs.datProducts.Recordset![Type]
        frmClientMI.cmbMICat.Enabled = False
        frmClientMI.pGetClientMI (gobjPNR.CN)
        '230108
        
        
        frmClientMI.Show 'vbModal
    End If
End Sub
Private Sub cmdDone_Click()
    Dim freefields As String
    Dim strMsg As String
    
    '02062005
    datTouchEnd = Now
    If chkAmdTktOnly.Visible = True And chkAmdTktOnly.value = 1 Then
       'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
       'Don't apply this logic if web fare applied
       'If Len(txtALCode) <> 3 Or Len(txtTktNum) <> 10 Or IsNumeric(txtALCode) = False Or IsNumeric(txtTktNum) = False Then
       If mbolWebFareSelected = False And (Len(txtALCode) <> 3 Or Len(txtTktNum) <> 10 Or IsNumeric(txtALCode) = False Or IsNumeric(txtTktNum) = False) Then
          strMsg = strMsg & "Invalid ticket number..." & Chr(13)
          'MsgBox strMsg
          modMsgBox.OKMsg = "OK"
          modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
          Exit Sub
       'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
       ElseIf mbolWebFareSelected = True Then
         If Trim(txtTktNum.Text) = "" Then
            strMsg = "Need ticket number..." & Chr(13)
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
            Exit Sub
         End If
       End If
       'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
       If mbolWebFareSelected = True Then updateTktInRI
       UpdatePO
       Call pTktQueue
       Unload frmClientMI
       Me.MousePointer = 0
       Unload Me
       'Unload frmWait
       Set gobjEO = Nothing
       Exit Sub
    End If
    If Not validData Then Exit Sub
    cmdDone.Enabled = False
    
    gSysStartOthSvcsTime = Now
    If gobjEO Is Nothing Or txtEONum = "" Then Call SetEOObj(gobjEO)
    
    freefields = ""
    If gobjEO.FF7 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "7-" & gobjEO.FF7
    If gobjEO.FF8 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "8-" & gobjEO.FF8
    If gobjEO.FF81 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "81-" & gobjEO.FF81
    ''CS Change EC
    'If gobjEO.rs <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "30-" & gobjEO.rs
    'CS Remove FF26
    'If gobjEO.FF26 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "26-" & gobjEO.FF26
    If gobjEO.FF38 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "38-" & gobjEO.FF38
    ''CS Add FF41
    'If gobjEO.FF41 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "41-" & gobjEO.FF41
    
    If frmClientMI.MSXfreefields <> "" Then
        freefields = freefields & "/" & frmClientMI.MSXfreefields
    End If
    
    'Check for completion of client MI
    If isRequireClientMI(gobjPNR.CN, 4) And frmClientMI.MSXfreefields = "" Then
        cmdDone.Enabled = True
        'MsgBox "Client MI data is incomplete", vbCritical
        strMsg = "Client MI data is incomplete"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        loadClientMI
        Exit Sub
    Else
        'Timer
        'frmWait.Show
        Me.MousePointer = 11
        Call modOthSvcs.WriteOSToGDS(gobjEO, frmOthSvcs.datProducts.Recordset![Type], gStartOthSvcsTime, freefields)
        Log
        Call pTktQueue
          
        
        Unload frmClientMI
        Me.MousePointer = 0
        Unload Me
        'Unload frmWait
        Set gobjEO = Nothing
    End If
        
    
End Sub

Private Sub cmdEO_Click()
Dim lngC As Long
Dim strMsg As String
datTouchEnd = Now
If Not validData Then Exit Sub
gSysStartOthSvcsTime = Now
'chk clientMI
If isRequireClientMI(gobjPNR.CN, 4) And frmClientMI.MSXfreefields = "" Then
        'cmdDone.Enabled = True
        'MsgBox "Client MI data is incomplete", vbCritical
        strMsg = "Client MI data is incomplete"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        loadClientMI
        Exit Sub
End If

Call SetEOObj(gobjEO)

'Added on 8/3/2005: To end PNR to get RecLoc
'Call modOthSvcs.SetEONumber

Log

Load frmExchangeOrder
frmExchangeOrder.Show '1, Me
        Do
            DoEvents
        Loop Until isLoaded("frmExchangeOrder") = False

If gbolIgnoreEO Then Exit Sub

txtEONum.Text = gobjEO.EONumber
txtEONum.Locked = True
Unload frmExchangeOrder
Set frmExchangeOrder = Nothing

'Modified on 24/1/2005: Change of EO# format
'If gobjEO.TicketNumber = "0000" Then
'    'gobjEO.TicketNumber = frmOthSvcs.datProducts.Recordset![TktPrefix] & gobjEO.TicketNumber & txtEONum
'    gobjEO.TicketNumber = frmOthSvcs.datProducts.Recordset![TktPrefix] & gobjEO.TicketNumber & Right(gobjEO.EONumber, Len(gobjEO.EONumber) - Len(frmOthSvcs.datProducts.Recordset![TktPrefix] & Format(Now, "yymm")))
'End If

cmdEO.Enabled = False

'Added on 8/3/2005: To end PNR to get RecLoc
cmdDone.Enabled = False
Call pTktQueue
Set gobjEO = Nothing
Unload Me

'Call cmdDone_Click

End Sub

Private Sub SetEOObj(ByRef objEO As EO)
Dim lngC As Long
Dim strPaxName As String

'Set objEO = New EO
'txtEONum = ""
If objEO Is Nothing Then
   Set objEO = New EO
Else
   'txtEONum = objEO.EONumber
   'Set objEO = New EO
   'objEO.EONumber = txtEONum
   If gbolEOAmend Then
      objEO.EONumber = txtEONum
   Else
      txtEONum = objEO.EONumber
   End If
   Set objEO = New EO
   objEO.EONumber = txtEONum
End If


With objEO
    If gbolEOAmend Then
       .EONumber = txtEONum
       'frmOthSvcs.datProducts.DatabaseName = gstrTProDBSource
       'frmOthSvcs.datProducts.RecordSource = "SELECT * FROM tblProductCodes where ProductCode = '" & mstrPCAmend & "'"
       'frmOthSvcs.datProducts.Refresh
    End If
    .BillingDescription = ""
    .CN = gobjPNR.CN
    If mblnConsTkt = False Then

       .SF = fConvertZero(txtSellingPrice)
    
    End If
    .CommissionAmt = fConvertZero(txtCommission.Text) + fConvertZero(txtMerchFee)
    .Cost = fConvertZero(txtCost.Text)
    .CreatedBy = gobjHost.AgentSine
    '.CreatedByName = gobjHost.AgentName
    .CreatedByName = gobjHost.AgentProfile
    .CreatedByPCC = gobjHost.AgentPCC
    .CreateDtTm = Now()
    '.FOP = cmbFOPType.Text & IIf(cmbFOPType.Text = "CX" Or (frmOSAirTkt.cmbFOPType = "CC"), "/" & cmbCCType.Text & "/" & txtCCNum.Text & "/" & Format(dtpCCExp.Value, "ddmm"), "")
    .FOP = cmbFOPType.Text & IIf(cmbFOPType.Text = "CX" Or (frmOSAirTkt.cmbFOPType = "CC"), "/" & cmbCCType.Text & "/" & txtCCNum.Text & "/" & Format(dtpCCExp.value, "MMYY"), "") & "/" & chkNRCC.value & "/" & chkWaiveMercFee.value
    
    For lngC = 1 To gobjPNR.PassengerCount
     strPaxName = strPaxName & gobjPNR.PassengerName(lngC).LastName & "/" & gobjPNR.PassengerName(lngC).FirstName & IIf(lngC = gobjPNR.PassengerCount, "", vbCrLf)
    Next
    .PaxName = strPaxName
    
    .PNRRecLoc = gobjPNR.RecLoc
    '.ProductCode = frmOthSvcs.datProducts.Recordset![SortKey]
    .ProductCode = frmOthSvcs.datProducts.Recordset![ProductCode]
    .ProductSortKey = frmOthSvcs.datProducts.Recordset![SortKey]
    .SellPrice = fConvertZero(txtSellPrice.Text)
'    If UCase(gstrAgcyCountryCode) = "SG" Then
'    .ServiceDate = DateAdd("M", 6, Date)
'    Else
'    .ServiceDate = DateAdd("d", 90, Date) 'DateAdd("m", 3, gobjPNR.AirSeg(lstFlights.ListCount).ArriveDateTime)
'    End If

    'ZhiSam - V1.2.23 20130911 - CR-289 - E-Invoice: Due and Paid Lines Date
    If bfunctCheckRTLine = True Then
        .ServiceDate = dtfunctRTDate
    Else
        .ServiceDate = DateAdd("d", 90, Date)
    End If
    .TaxAdd IIf(txtTax(0).Text = "", 0, fConvertZero(txtTax(0).Text)), txtTaxCode(0).Text
    .TaxAdd IIf(txtTax(1).Text = "", 0, fConvertZero(txtTax(1).Text)), txtTaxCode(1).Text
    
    '02062005
    'If chkTktAtInv.value = 0 Then
    .TicketNumber = Format(txtALCode.Text, "000") & txtTktNum.Text
    If Trim(txtConjTkt.Text) <> "" Then
       .ConjunctTicket = Format(Trim(txtConjTkt.Text), "00")
    End If
    
    'Else
    '   .TicketNumber = Format(txtALCode.Text, "000")
    'End If
    
    '02062005
    'If UCase(gstrAgcyCountryCode) = "HK" Then
    '    .TicketNumber = IIf(Me.chkTktAtInv.Value = vbChecked, "**CT", Format(txtALCode.Text, "000") & txtTktNum.Text)
    'Else
    '    If gbolEOAmend Then
    '       .TicketNumber = mstrTktNum
    '    Else
    '       .TicketNumber = "0000"
    '    End If
    'End If
    
    'preethi – V1.2.6 20110905 – CR99 - Add Option for Fare Type in EO
  If mblnConsTkt = False Then
    If optPublishedFare(0).value = True Then
       .FareType = 1
    Else
      If optNettFare(1).value = True Then
         .FareType = 2
      Else
         .FareType = 0
      End If
    End If
    'preethi – V1.2.6 20110905 – CR98 - Reissue Ticket Box in EO
    If chkReissuedTkt.value = vbChecked Then
       .TktNumber = txtOriTktNum.Text
    Else
       .TktNumber = ""
    End If
  End If
    '.VendorCode = frmOthSvcs.dbcVendors.BoundText
    
    .DescriptionLineAdd frmOthSvcs.dbcProducts.Text
    For lngC = 0 To lstFlights.ListCount - 1
       If lstFlights.Selected(lngC) Then .DescriptionLineAdd lstFlights.List(lngC)
    Next
    '.DescriptionLineAdd ""
    .DescriptionLineAdd ""
    .DescriptionLineAdd ""
    '.VendorCode = frmOthSvcs.datSelectedVendor.Recordset!VendorNumber
    '.Email = frmOthSvcs.datSelectedVendor.Recordset![Email] & ""
    '.FaxNo = frmOthSvcs.datSelectedVendor.Recordset![FaxNumber] & ""
    '.VendorName = frmOthSvcs.datSelectedVendor.Recordset!VendorName
    .VendorCode = frmOthSvcs.datSelectedVendor.Recordset!VendorNumber
    .Misc = frmOthSvcs.datSelectedVendor.Recordset!Misc
    .Email = txtEmail 'frmOthSvcs.datSelectedVendor.Recordset![Email] & ""
    .FaxNo = txtFaxNo 'frmOthSvcs.datSelectedVendor.Recordset![FaxNumber] & ""
    .VendorName = txtVendor 'frmOthSvcs.datSelectedVendor.Recordset!VendorName & ""
    .Address1 = txtAddress1 'frmOthSvcs.datSelectedVendor.Recordset!Address1 & ""
    .Address2 = txtAddress2 'frmOthSvcs.datSelectedVendor.Recordset!Address2 & ""
    .City = txtCity 'frmOthSvcs.datSelectedVendor.Recordset!City & ""
    .Country = txtCountry1 'frmOthSvcs.datSelectedVendor.Recordset!Country & ""
    .ContactNum = txtTel
    
    .ContactPerson = txtContact
    
    
    '.Email = frmOthSvcs.datVendors.Recordset![Email] & ""
    '.FaxNo = frmOthSvcs.datVendors.Recordset![FaxNumber] & ""
    For lngC = 0 To lstEORmks(1).ListCount - 1
        .RemarkAdd lstEORmks(1).List(lngC)
    Next
    
    'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
    If mbolWebFareSelected = True Then
       .WebFareApplied = True
       'Store LCC Fare Quote information into RI
       With mobjWebFares.WebFare(mintWebFareSelected)
            objEO.PlatingCarrier = .PlatingCarrier
            objEO.RIRemarkAdd "*************LCC FARE QUOTE PER PAX**************"
            objEO.RIRemarkAdd "LCC CONFIRMATION NUMBER: " & txtTktNum.Text
            objEO.RIRemarkAdd "BOOKING DATE: " & Format(.BookingDate, "dd MMMM yyyy")
            objEO.RIRemarkAdd "*************************************************"
            objEO.RIRemarkAdd "FOR " & getRouting(.Routing)
            objEO.RIRemarkAdd "BASE FARE: " & .FareCurrency & " " & Format(fConvertZero(txtSellPrice.Text) - fConvertZero(txtTax(0).Text) - fConvertZero(txtTax(1).Text), "0.00") & " PLUS " & Format(fConvertZero(txtTax(0).Text), "0.00") & " TAXES"
            If fConvertZero(txtTax(1).Text) > 0 Then objEO.RIRemarkAdd "SPECIAL SERVICE: " & Format(fConvertZero(txtTax(1).Text), "0.00")
            If fConvertZero(txtTrxnFee.Text) > 0 Then objEO.RIRemarkAdd "PLUS TRANSACTION FEE: " & Format(fConvertZero(txtTrxnFee.Text), "0.00")
            'JY – V1.2.3 20110429 – CR65 - Add Fuel Surcharge into RI line
            'objEO.RIRemarkAdd "TOTAL QUOTE: " & .FareCurrency & " " & Format(fConvertZero(txtSellPrice.Text) + fConvertZero(txtTrxnFee.Text), "0.00")
            If fConvertZero(txtFuelSurcharge.Text) > 0 Then objEO.RIRemarkAdd "PLUS FUEL CHARGE SERVICE FEE: " & Format(fConvertZero(txtFuelSurcharge.Text), "0.00")
            objEO.RIRemarkAdd "TOTAL QUOTE: " & .FareCurrency & " " & Format(fConvertZero(txtSellPrice.Text) + fConvertZero(txtTrxnFee.Text) + fConvertZero(txtFuelSurcharge.Text), "0.00")
            objEO.RIRemarkAdd "          VALID ON " & .PlatingCarrier & " ONLY"
            objEO.RIRemarkAdd "          SUBJECT TO AIRLINE FARE RESTRICTIONS"
            objEO.RIRemarkAdd "          AND PENALTY FEE"
            objEO.RIRemarkAdd "*****************FLIGHT SCHEDULE*****************"
            For lngC = 1 To .AirSegCount
                With .AirSeg(CInt(lngC))
                    objEO.RIRemarkAdd " " & mobjWebFares.WebFare(mintWebFareSelected).PlatingCarrier & " " & LeftAlign(.FlightNum, 4) & " " & .Class & " " & .DepDate & " " & .DepCity & .ArrCity & " " & .DepTime & " " & .ArrTime
                End With
            Next
            objEO.RIRemarkAdd "**************************************************"
       End With
    End If
    
    For lngC = 0 To lstItinRmks(1).ListCount - 1
        .RIRemarkAdd lstItinRmks(1).List(lngC)
    Next
    If gobjPNR.CompInfo.MI = True Then
        .RF = txtMI(0)
        .LF = txtMI(1)
        'CS Change EC
        '.EC = txtMI(2)
        .rs = txtRS
        .MS = txtMS
        .FF7 = txtMI(3)
        If cboClassServ.Text <> "" Then .FF8 = Trim(Left(cboClassServ.Text, 2))
        'End If
        'CS Remove FF26
        '.FF26 = IIf(UCase(cboTripType.Text) = "ROUND", "R", IIf(UCase(cboTripType.Text) = "ONE WAY", "O", ""))
        'CS Add FF41
        '.FF41 = IIf(UCase(cboTrip.Text) = "INTERNATIONAL", "I", IIf(UCase(cboTrip.Text) = "DOMESTIC", "D", ""))
        .FF81 = txtMI(6)
        .FF38 = chkPaperTkt.Caption
        If Not mblnConsTkt Then
        .BookingAction = cboBookingAction.Text
        .BookingTool = mstrBookingTool
        End If
        'FF10,11 will be maintained by client related MI
        '.FF10 = txtMI(4)
        '.FF11 = txtMI(5)
    End If
    '02062005
    If cboClientType.listindex = -1 Then
       .ClientType = ""
    Else
       .ClientType = cboClientType.Text
    End If
    .NettFare = IIf(txtNettFare = "", 0, txtNettFare)
    .PublishedFare = IIf(txtPubFare = "", 0, txtPubFare)
    .GrossFare = IIf(txtGrossFare = "", 0, txtGrossFare)
    .Discount = IIf(txtDC = "", 0, txtDC)
    .MerchFee = IIf(txtMerchFee = "", 0, txtMerchFee)
    .CWTAbsorb = IIf(chkAbsorb.value = 1, True, False)
    .TranxFee = IIf(txtTrxnFee = "", 0, txtTrxnFee)
    .SegSelect = SegmentSelect
    '.TktNum = txtALCode & ";" & txtTktNum & ";" & chkTktAtInv.Value
    .ListBoxRem = ListBoxRemark
    .PassengerID = IIf(txtPassengerID = "", "1", txtPassengerID)
    .TFNRCC = IIf(chkTFNRCC.value = 1, True, False)
    .ReplyEmail = Trim(UCase(txtReplyEmail.Text))
End With

  
    
End Sub

Private Function SegmentSelect() As String
   Dim i As Integer
   
   SegmentSelect = ""
   For i = 0 To lstFlights.ListCount - 1
      If lstFlights.Selected(i) Then
         SegmentSelect = SegmentSelect & IIf(SegmentSelect <> "", vbCrLf, "") & lstFlights.List(i)
      End If
   Next
End Function

Private Function ListBoxRemark() As String
   Dim i As Integer
   
   ListBoxRemark = ""
   For i = 0 To lstEORmks(1).ListCount - 1
      ListBoxRemark = ListBoxRemark & IIf(ListBoxRemark <> "", vbCrLf, "") & lstEORmks(1).List(i)
   Next
End Function

Private Sub cmdEORmksAddAll_Click()
Dim lngC As Long
Dim strTemp As String

With lstEORmks(0)
    For lngC = 0 To .ListCount - 1
    
    strTemp = .List(lngC)
   If strTemp = "" Then Exit For
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
           
                    Else
                        lstEORmks(1).AddItem strTemp
                        'lstEORmks(0).RemoveItem lngC
                        'lngC = lngC - 1

                    End If
                    Unload frmFareRmkFill
                End With
              
            Set frmFareRmkFill = Nothing
            
     Else
        lstEORmks(1).AddItem .List(lngC)
        'If lngC = 0 Then
        '    lstEORmks(0).RemoveItem 0
        '    lngC = lngC - 1
        'End If
     End If

    Next lngC
End With

End Sub

Private Sub cmdEORmksAddOne_Click()
Dim strTemp As String
With lstEORmks(0)
If .SelCount > 0 Then
strTemp = .Text

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
                        lstEORmks(1).AddItem strTemp
                        'lstEORmks(0).RemoveItem lstEORmks(0).ListIndex
                    End If
                    Unload frmFareRmkFill
                End With

            Set frmFareRmkFill = Nothing
    Else
    lstEORmks(1).AddItem .Text
   ' .RemoveItem .ListIndex
    End If
End If
End With




End Sub

Private Sub cmdEORmksRemove_Click()

'With lstEORmks(1)
'If .SelCount > 0 Then
 '   .RemoveItem .ListIndex
'End If
'End With
Dim intC As Integer
With lstEORmks(1)
For intC = .ListCount - 1 To 0 Step -1

If .Selected(intC) = True Then
    .RemoveItem intC
End If
Next intC
End With
End Sub
Private Sub cmdFreeRmkToEO_Click()
If txtFreeRmk.Text <> "" Then
    lstEORmks(1).AddItem txtFreeRmk.Text
    txtFreeRmk.Text = ""
End If

End Sub

Private Sub cmdFreeRmkToItin_Click()
If txtFreeRmk.Text <> "" And ValidIR(txtFreeRmk.Text) = True Then
    lstItinRmks(1).AddItem txtFreeRmk.Text
    txtFreeRmk.Text = ""
End If
End Sub


Private Sub cmdItinRmksAddAll_Click()
Dim lngC As Long
Dim strTemp As String

With lstItinRmks(0)
    For lngC = 0 To .ListCount - 1
        strTemp = .List(lngC)
   If strTemp = "" Then Exit For
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
           
                    Else
                        If ValidIR(strTemp) = True Then
                        lstItinRmks(1).AddItem strTemp
                        'lstItinRmks(0).RemoveItem lngC
                        'lngC = lngC - 1
                        End If

                    End If
                    Unload frmFareRmkFill
                End With
              
            Set frmFareRmkFill = Nothing
            
     Else
        
        If ValidIR(strTemp) = True Then
        lstItinRmks(1).AddItem .List(lngC)
    
        'If lngC = 0 Then
        'lstItinRmks(0).RemoveItem 0
        'lngC = lngC - 1
 
        'Else
        'lstItinRmks(0).RemoveItem lngC
        'lngC = lngC - 1
        'End If
        
        
        End If
     
     End If

    Next lngC
End With

End Sub

Private Sub cmdItinRmksAddOne_Click()
Dim strTemp As String

With lstItinRmks(0)

If .SelCount > 0 And ValidIR(.Text) = True Then

strTemp = .Text

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
                        lstItinRmks(1).AddItem strTemp
                        'lstItinRmks(0).RemoveItem lstItinRmks(0).ListIndex
                    End If
                    Unload frmFareRmkFill
                End With

            Set frmFareRmkFill = Nothing
    Else
    lstItinRmks(1).AddItem .Text
    '.RemoveItem .ListIndex
    End If
End If
End With
End Sub


Private Sub cmdItinRmksRemove_Click()


'With Me.lstItinRmks(1)
'If .SelCount > 0 Then
'    .RemoveItem .ListIndex
'End If
'End With
Dim intC As Integer
With Me.lstItinRmks(1)
For intC = .ListCount - 1 To 0 Step -1

If .Selected(intC) = True Then
    .RemoveItem intC
End If
Next intC

End With
End Sub


'Private Sub dbcVendors_Click(Area As Integer)
'   frmOthSvcs.dbcVendors.Text = dbcVendors.Text
'   frmOthSvcs.datSelectedVendor.DatabaseName = gstrTProDBSource 'GetSetting("TPro", "Startup", "TProDBSource", "NOT FOUND")
'   frmOthSvcs.datSelectedVendor.RecordSource = "SELECT * FROM tblVendors WHERE [VendorNumber] =  '" & dbcVendors.BoundText & "'"
'   frmOthSvcs.datSelectedVendor.Refresh
'   GetVendorInfo
'End Sub
Private Sub dbcVendors_Change()
   frmOthSvcs.dbcVendors.Text = dbcVendors.Text
   With frmOthSvcs.datSelectedVendor
   'frmOthSvcs.datSelectedVendor.DatabaseName = gstrTProDBSource 'GetSetting("TPro", "Startup", "TProDBSource", "NOT FOUND")
   .ConnectionString = gstrConn
   .Mode = adModeRead
   .CommandType = adCmdText
   .RecordSource = "SELECT * FROM tblVendors WHERE [VendorNumber] =  '" & dbcVendors.BoundText & "'"
   .Refresh
   End With
   GetVendorInfo
End Sub

Private Sub LockedText(Locked As Boolean)
   txtVendor.Locked = Locked
   txtAddress1.Locked = Locked
   txtAddress2.Locked = Locked
   txtCity.Locked = Locked
   txtCountry1.Locked = Locked
   txtCreditTerms.Locked = Locked
   txtTel.Locked = Locked
   txtEmail.Locked = False
   txtFaxNo.Locked = False
End Sub

Private Sub GetVendorInfo()
   If gbolEOAmend Then
      'If dbcVendors.BoundText = "021222" Then
      If frmOthSvcs.datSelectedVendor.Recordset!Misc = True Then
         'fraVendorInfo.Enabled = True
         LockedText False
      Else
         'fraVendorInfo.Enabled = False
         LockedText True
      End If
   Else
      'If frmOthSvcs.dbcVendors.BoundText = "021222" Then
      If frmOthSvcs.datSelectedVendor.Recordset!Misc = True Then
         'fraVendorInfo.Enabled = True
         LockedText False
      Else
         'fraVendorInfo.Enabled = True
         LockedText True
      End If
   End If
   txtVendor = frmOthSvcs.datSelectedVendor.Recordset!VendorName & ""
   txtAddress1 = frmOthSvcs.datSelectedVendor.Recordset!Address1 & ""
   txtAddress2 = frmOthSvcs.datSelectedVendor.Recordset!Address2 & ""
   txtCity = frmOthSvcs.datSelectedVendor.Recordset!City & ""
   txtCountry1 = frmOthSvcs.datSelectedVendor.Recordset!Country & ""
   txtEmail = frmOthSvcs.datSelectedVendor.Recordset!Email & ""
   txtFaxNo = frmOthSvcs.datSelectedVendor.Recordset!FaxNumber & ""
   txtTel = frmOthSvcs.datSelectedVendor.Recordset!ContactNo & ""
End Sub

Private Sub Form_Load()
Dim lngC As Long
Dim rsRemarks As New ADODB.Recordset
Dim strSQL As String
Dim blnChkDate As Boolean
Dim blnChkVendor As Boolean
Dim intPaxCount As Integer
Dim strTemp As String

Dim oldParent As Long
    
    datFormLoadStart = Now
    gintY = 0
    gintX = 0
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
Set gobjEO = New EO
Set gobjPreEO = New EO
'FormCenter

frmOSAirTkt.Caption = "CWT TravelPro - " & frmOthSvcs.dbcProducts.Text
chkformula.value = 1
pSetInitialValues
'Timer
gStartOthSvcsTime = Now

For lngC = 0 To 4
    cboClientType.AddItem Mid("DUTFMGMNDB", ((lngC * 2) + 1), 2)
Next lngC

mblnConsTkt = (frmOthSvcs.datProducts.Recordset![Type] = "CT")

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
mbolWebFareSelected = False

'BSP
If Not mblnConsTkt Then
'added on 2/8/05:MCO
cmbFOPType.AddItem "CC"
SSTab1.TabEnabled(3) = True
SSTab1.TabEnabled(1) = False
'fraHotel.Enabled = False
chkNRCC.Enabled = True
If gobjPNR.RecLoc <> "" Then txtRecLoc = gobjPNR.RecLoc
 intPaxCount = gobjPNR.PassengerCount
        If intPaxCount > 0 Then
            For lngC = 1 To intPaxCount
                With gobjPNR.PassengerName(lngC)
                    strTemp = .LastName & "/" & .FirstName
                End With
               cmbTraName.AddItem strTemp
            Next
        End If
        
  If cmbTraName.ListCount > 0 Then cmbTraName.listindex = 0
  chkRequestMCO_Click
  cboBookingAction.Visible = True
  lblLabels(39).Visible = True
Else
    SSTab1.TabEnabled(3) = False
    SSTab1.TabEnabled(1) = True
    'fraHotel.Enabled = True
    chkNRCC.Enabled = False
    'Remove on 100807:  To default 10 zeros for ticket number, Tpro Conso - To remove "Enter Ticket Number later" label
    If Not gbolEOAmend Then txtTktNum.Text = "0000000000"
    lblAirlineCodeRem.Visible = True
    txtContact.Visible = True
    lblContact.Visible = True
    cboBookingAction.Visible = False
    lblLabels(39).Visible = False
End If

'Preethi-V1.2.2 20110113 - CR30 - EO FOP for LCC Transaction
If (frmOthSvcs.datProducts.Recordset![EnableCCFOP] & "") <> "" And frmOthSvcs.datProducts.Recordset![EnableCCFOP] = "True" Then
   cmbFOPType.AddItem "CC"
End If

'Remove on 100807:  To default 10 zeros for ticket number, Tpro Conso - To remove "Enter Ticket Number later" label
'chkTktAtInv.Visible =mblnConsTkt
'If mblnConsTkt Then
'    If Not gbolEOAmend Then txtTktNum.Text = "0000000000"
'    lblAirlineCodeRem.Visible = True
'End If

    'Modified on 15/03/05: NRCC checkbox for HK, SG for BSP only
    'If UCase(gstrAgcyCountryCode) = "SG" Then
    '    chkNRCC.Visible = Not mblnConsTkt
    'Else
    '    chkNRCC.Visible = True
    'End If
'If Not mblnConsTkt Then cmbFOPType.AddItem "CC"
'cmdEO.Enabled = mblnConsTkt


'strSQL = "SELECT * FROM tblProductRemarks " _
'    & "WHERE [ProductType] = '" & frmOthSvcs.datProducts.Recordset![Type] & "'"

If gbolEOAmend Then
   strSQL = "SELECT * FROM tblProductRemarks " _
    & "WHERE [ProductType] = '" & frmEOAmend.lsvEO.SelectedItem.SubItems(4) & "'"
Else
   strSQL = "SELECT * FROM tblProductRemarks " _
    & "WHERE [ProductType] = '" & frmOthSvcs.datProducts.Recordset![Type] & "'"
End If

Set rsRemarks = gdbConn.Execute(strSQL)

With rsRemarks
    Do Until .EOF
        If ![RmkType] & "" = "I" Then
            lstItinRmks(0).AddItem ![Text]
        Else
            lstEORmks(0).AddItem ![Text]
        End If
        .MoveNext
    Loop
End With

'Set gobjPNR = New CWT_GalileoPNR.PNR
'If gobjLog.LogOpen Then gobjPNR.OpenLog gobjLog
Set gobjPNR = New CWT_GalileoPNR3.PNR
With gobjPNR
    Call .loadPNR
    'If .FOPType = "CC" And .FOP_CCExpireDate <= Date Then
    '   MsgBox "Invalid credit card expiry date, please update!"
    '   'Unload Me
    '   gbolToMainMenu = True
    '   frmMainMenu.Show
    '   Exit Sub
    'End If
    
    For lngC = 1 To .AirSegCount
        lstFlights.AddItem .AirSeg(lngC).TextAirSeg
    Next
     'Modified on 27/01/05: add on credit card vendor validation together with CC ExpireDate Validation
    If .FOPType = "CC" Then
        blnChkDate = True
        blnChkVendor = True
        
        If .FOP_CCCode <> "" Then
            If validateCCVendor(cmbCCType) = True Then
                Me.cmbCCType.Text = .FOP_CCCode
            Else
                blnChkVendor = False
            End If
        End If
        Me.txtCCNum.Text = .FOP_CCNum
        
        If validateCCDate(.FOP_CCExpireDate) Then
            Me.dtpCCExp.value = .FOP_CCExpireDate
        Else
            blnChkDate = False
        End If
        'Preethi-V1.2.2 20110113 - CR30 - EO FOP for LCC Transaction
        'Me.cmbFOPType = IIf(mblnConsTkt = True, "CX", "CC")
        
        If mblnConsTkt = "True" And ((frmOthSvcs.datProducts.Recordset![EnableCCFOP] & "") = "" Or _
                           frmOthSvcs.datProducts.Recordset![EnableCCFOP] = "False") Then
           Me.cmbFOPType = "CX"
        Else
           Me.cmbFOPType = "CC"
        End If
        
        If blnChkVendor = False Or blnChkDate = False Then
           promptCCError blnChkVendor, blnChkDate
        End If
    Else
        Me.cmbFOPType = "INV"
    End If
    
 End With
 
 cmbDispNum.Clear
 cmbDispNum.AddItem ""
 If MaxMIDispNum > 0 Then
    For lngC = 1 To MaxMIDispNum
        cmbDispNum.AddItem "*" & lngC
    Next
 End If
 
 With gobjPNR.CompInfo
    'cboClientType = IIf(.ClientType = "", "DU", .ClientType)
    msngMarkup = .MarkUp
    msngIntlDiscount = .DiscountInternational
    Me.txtCommission.Text = msngMarkup
    Me.optCommType(1).value = vbChecked
    Me.txtDC.Text = msngIntlDiscount
    Me.optDiscount(1).value = vbChecked
    msngMerchFeePct = .MerchFeePct

    
 End With
    
 If UCase(gstrAgcyCountryCode) = "HK" Then
    'Label3.Visible = False
    lblFare.Caption = "Nett Fare:"
    txtPubFare.Visible = False
    'Label10.Visible = True
    txtNettFare.Visible = True
    Frame4.Visible = True
    lblCT.Visible = True
    cboClientType.Visible = True
    cboClientType.Text = fGetCNType(gobjPNR.CompInfo.ProfileName)
    
    If cboClientType = "TF" Then
       txtTrxnFee.Enabled = True
    Else
       txtTrxnFee.Enabled = False
    End If
    txtSellingPrice.Visible = False
    lblSellingPrice.Visible = False
 Else
    'Label3.Visible = True
    txtPubFare.Visible = True
    'Label10.Visible = False
    If mblnConsTkt Then
        lblFare.Caption = "Selling Fare to Client:"
    Else
        lblFare.Caption = "Publish Fare:"
    End If
    txtNettFare.Visible = False
    Frame4.Visible = True
    txtGrossFare.Visible = False
    Label11.Visible = False
    'Label3.Visible = True
    txtPubFare.Visible = True
    'Label10.Visible = False
    txtNettFare.Visible = False
    lblCT.Visible = False
    cboClientType.Visible = False
    If mblnConsTkt Then
       txtSellingPrice.Visible = False
       lblSellingPrice.Visible = False
    Else
       txtSellingPrice.Visible = True
       lblSellingPrice.Visible = True
    End If
 End If
 
If Not mblnConsTkt Then
   txtConjTkt.Visible = True
   Line1.Visible = True
Else
   txtConjTkt.Visible = False
   Line1.Visible = False
End If
  
'added on 10/11/04: HKG request - Disabled client type for corporate/group trip
'If gstrAgcyCountryCode = "HK" Then
'    If gTrxnType = "L" Then
'        cboClientType.Enabled = True
'    Else
'        cboClientType.Enabled = False
'    End If
'End If
 
 SSTab1.Tab = 0
 
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
If mblnConsTkt Then
    Set mobjWebFares = AddLCCWebFare
    Set mobjDeclinedWebFares = AddLCCFareOption
End If

 dbcVendors.Visible = False
 
 'preethi – V1.2.6 20110905 – CR99 - Add Option for Fare Type in EO
If mblnConsTkt = False Then
   optPublishedFare(0).Visible = True
   optNettFare(1).Visible = True
   'preethi – V1.2.6 20110905 – CR98 - Reissue Ticket Box in EO
   chkReissuedTkt.Visible = True
   chkReissuedTkt.Enabled = True
   txtOriTktNum.Visible = True
   lblOriTktNum.Visible = True
   lblTktLength.Visible = True
End If

  If gbolEOAmend Then
    '02062005
    If mblnConsTkt Then
       chkAmdTktOnly.Visible = True
    Else
       chkAmdTktOnly.Visible = False
    End If
    
    'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
    mbolEORetrieving = True
    RetrieveData
    mbolEORetrieving = False
    If mblnConsTkt = True And mbolWebFareSelected = True Then txtALCode.Enabled = False
 Else
    GetVendorInfo
 End If
 
 'Added on 24 Jan: Vender credit term for EO Transaction
If frmOthSvcs.datSelectedVendor.Recordset![CreditTerms] <> "" Then txtCreditTerms = frmOthSvcs.datSelectedVendor.Recordset![CreditTerms]
txtCreditTerms.Enabled = False
txtCreditTerms.Locked = True

'added on 14/06/2005: for SG, auto enable/disable EO button
If UCase(gstrAgcyCountryCode) = "SG" Then
    If IsNull(frmOthSvcs.datSelectedVendor.Recordset!RaiseType) Then
        cmdEO.Enabled = False
        cmdDone.Enabled = True
    Else
        cmdEO.Enabled = True
        cmdDone.Enabled = False
    End If
Else
    cmdEO.Enabled = mblnConsTkt
End If


'added on 30032006: TF NRCC
'Changed on 28092007 : Applied to both coutries
'If UCase(gstrAgcyCountryCode) = "SG" Then
    If Not mblnConsTkt Then chkTFNRCC.Visible = True
'End If
txtReplyEmail.Text = GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjPNR.Agent, gobjPNR.PCCOwner, False, True)


'230108: Add Fuel Surcharge for HKG
Dim blnfuelcharge As Boolean
Dim sngfuelAmt As Single

pFuelSurcharge blnfuelcharge, sngfuelAmt

If blnfuelcharge = True Then
    txtFuelSurcharge.Visible = True
    lblFuelSurcharge.Visible = True
    txtFuelSurcharge.Text = sngfuelAmt
   
Else
    txtFuelSurcharge.Visible = False
    lblFuelSurcharge.Visible = False
    
    
End If


If OSNoMF(frmOthSvcs.datProducts.Recordset![ProductCode], frmOthSvcs.datSelectedVendor.Recordset!VendorNumber) = True Then
    chkAbsorb.value = 1
    chkAbsorb.Enabled = False
Else
    chkAbsorb.value = 0
    chkAbsorb.Enabled = True
End If

datFormLoadEnd = Now
If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
End Sub

Private Sub RetrieveData()
   Dim strSQL As String
   Dim rsEO As New ADODB.Recordset
   Dim i As Integer
   Dim j As Integer
   Dim strFOP() As String
   Dim strSegSel() As String
   Dim strTktNum() As String
   Dim strRem() As String
   Dim strMI() As String
   Dim strTemp() As String
   
   dbcVendors.Visible = True
   cmdDone.Enabled = False
   mstrPCAmend = frmEOAmend.lsvEO.SelectedItem.SubItems(1)
    'datVendors.DatabaseName = gstrTProDBSource
   
    'datVendors.RecordSource = "SELECT * FROM tblVendors WHERE [ProductCodes] LIKE '*" & mstrPCAmend & "*' ORDER BY [VendorName]"
    'datVendors.Refresh

'With dbcVendors
'    .Text = ""
'    .ListField = "VendorName"
'    '.BoundColumn = "SortKey"
'    .BoundColumn = "VendorNumber"
'    .Refresh
'End With
   'modified on 23/03/2005
    datVendors.ConnectionString = gstrConn
    datVendors.Mode = adModeRead
    datVendors.CommandType = adCmdText
    datVendors.RecordSource = "SELECT * FROM tblVendors WHERE [ProductCodes] LIKE '%" & mstrPCAmend & "%' ORDER BY [VendorName]"
    datVendors.Refresh
    
    With dbcVendors
         Set .DataSource = datVendors
        .Text = ""
        .ListField = "VendorName"
        .BoundColumn = "VendorNumber"
        .Refresh
    End With
    
   dbcVendors.BoundText = frmEOAmend.lsvEO.SelectedItem.SubItems(5)
      
   txtEONum.Locked = True
   strSQL = "Select * from tblExchangeOrder where "
   strSQL = strSQL & "ExchangeID = '"
   strSQL = strSQL & frmEOAmend.lsvEO.SelectedItem.Text & "'"

   Set rsEO = New ADODB.Recordset
   rsEO.Open strSQL, gdbConn, adOpenKeyset, adLockReadOnly
   With rsEO
     txtEONum = !ExchangeID
     For i = 0 To cboClientType.ListCount - 1
         If cboClientType.List(i) = !ClientType Then
            cboClientType.listindex = i
            Exit For
         End If
     Next
     mstrTktNum = !TktNum & ""
     'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
     'Disable due to duplicate code in this procedure
     'If Len(mstrTktNum) >= 3 Then
     '   txtALCode = Left(mstrTktNum, 3)
     '   If InStr(1, mstrTktNum, "-") > 0 Then
     '      txtTktNum = Mid(mstrTktNum, 4, Len(mstrTktNum) - 6)
     '      txtConjTkt = Right(mstrTktNum, 2)
     '   Else
     '      txtTktNum = Mid(mstrTktNum, 4)
     '      txtConjTkt = ""
     '   End If
     'End If
     txtContact = !ContactPerson
     txtNettFare = !NettFare
     txtPubFare = !PubFare
     txtGrossFare = !GrossFare
     txtTax(0).Text = !Tax1
     txtTaxCode(0).Text = !TAXCODE1 & ""
     txtTax(1).Text = !Tax2
     txtTaxCode(1).Text = !TaxCode2 & ""
     txtCost = !Cost
     txtCommission = !Commission - !MerchantFee
     optCommType(0).value = True
     txtDC = !Discount
     optDiscount(0).value = True
     txtMerchFee = !MerchantFee
     chkAbsorb.value = IIf(!CWTAbsorb = True, 1, 0)
     'txtSellPrice = !SellPrice
     txtTrxnFee = !TransactionFee
     strFOP = Split(!FOP, "/")
     If strFOP(0) = "INV" Then
        
        If UBound(strFOP) = 2 Then
           chkNRCC.value = strFOP(1)
           chkWaiveMercFee.value = strFOP(2)
        End If
        cmbFOPType.Text = strFOP(0)
     Else
        If UBound(strFOP) = 5 Then
           chkNRCC.value = strFOP(4)
           chkWaiveMercFee.value = strFOP(5)
        End If
        cmbFOPType.Text = strFOP(0)
        cmbCCType.Text = strFOP(1)
        txtCCNum.Text = strFOP(2)
        dtpCCExp.value = "1/" & MMM(Left(strFOP(3), 2)) & "/" & Right(strFOP(3), 2)

     End If
     strSegSel = Split(!SegmentSelect, vbCrLf)
     For i = 0 To lstFlights.ListCount - 1
        For j = 0 To UBound(strSegSel)
           'If lstFlights.List(i) = strSegSel(J) Then
           If InStr(strSegSel(j), ".") > 0 Then
                If InStr(lstFlights.List(i), Mid(strSegSel(j), InStr(strSegSel(j), "."))) > 0 Then
                   lstFlights.Selected(i) = True
                   Exit For
                End If
           End If
        Next
     Next
     'strTktNum = Split(!TktNum, ";")
     'txtALCode = strTktNum(0)
     'txtTktNum = strTktNum(1)
     'chkTktAtInv.Value = strTktNum(2)
     
     If !TktNum & "" = "**CT" Then
        'chkTktAtInv.value = 1
        txtTktNum = "00000000000"
        txtALCode = ""
        txtTktNum = ""
     ElseIf Len(!TktNum & "") > 3 Then
        'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
        If mbolWebFareSelected = False Then
            txtALCode = Left(!TktNum, 3)
            'txtTktNum = Mid(!TktNum, 4)
            'chkTktAtInv.value = 0
            mstrTktNum = !TktNum & ""
            If InStr(1, mstrTktNum, "-") > 0 Then
               txtTktNum = Mid(mstrTktNum, 4, Len(mstrTktNum) - 6)
               txtConjTkt = Right(mstrTktNum, 2)
            Else
               txtTktNum = Mid(mstrTktNum, 4)
               txtConjTkt = ""
            End If
        Else
            txtTktNum = !TktNum & ""
        End If
     End If
     lstEORmks(1).Clear
     strRem = Split(!ListBoxRemark, vbCrLf)
     For i = 0 To UBound(strRem)
        lstEORmks(1).AddItem strRem(i)
     Next
     strMI = Split(!MIData, ";")
     If UBound(strMI) >= 6 Then
        txtMI(0) = strMI(0)
        txtMI(1) = strMI(1)
        'CS Change EC
        'txtMI(2) = strMI(2)
        txtMS = strMI(2)
        If UBound(strMI) >= 8 Then
           txtRS = strMI(8)
        End If
        txtMI(3) = strMI(3)
        txtMI(6) = strMI(6)
        'cboMIFareType.Text = strMI(4)
        If strMI(4) <> "" Then
           cboClassServ = matchList(strMI(4))
        End If
        'cboMIFareType.listindex = 0
        'For i = 1 To cboMIFareType.ListCount - 1
        '   If (Mid(cboMIFareType.List(i), 1, InStr(1, cboMIFareType.List(i), "-") - 1)) = strMI(4) Then
        '      cboMIFareType.listindex = i
        '   End If
        'Next
        'CS Remove FF26
        'If strMI(5) = "O" Then
        '   cboTripType.Text = "ONE WAY"
        'ElseIf strMI(5) = "R" Then
        '   cboTripType.Text = "ROUND"
        'End If
        'CS Add FF41
        'If strMI(5) = "I" Then
        '   cboTrip.Text = "INTERNATIONAL"
        'ElseIf strMI(5) = "O" Then
        '   cboTrip.Text = "DOMESTIC"
        'End If
     End If
     If UBound(strMI) >= 7 Then
        If strMI(7) = "ET" Then
            chkPaperTkt.value = vbUnchecked
        Else
            chkPaperTkt.value = vbUnchecked
        End If
     End If
     'CS Add FF41
     'If UBound(strMI) >= 9 Then
     '   If strMI(9) = "I" Then
     '      cboTrip.Text = "INTERNATIONAL"
     '   ElseIf strMI(9) = "D" Then
     '      cboTrip.Text = "DOMESTIC"
     '   End If
     'End If
     'If dbcVendors.BoundText = "021222" Then
     If frmOthSvcs.datSelectedVendor.Recordset!Misc = True Then
        strTemp = Split(!VendorInfo, vbCrLf)
        If UBound(strTemp) >= 6 Then
           txtVendor = strTemp(0)
           txtAddress1 = strTemp(1)
           txtAddress2 = strTemp(2)
           txtCity = strTemp(3)
           txtCountry1 = strTemp(4)
           txtEmail = strTemp(5)
           txtFaxNo = strTemp(6)
        End If
        If UBound(strTemp) >= 7 Then
           txtTel = strTemp(7)
        End If
     End If
     If IsNull(!VendorEmail) = False Then
        txtEmail = !VendorEmail & ""
        txtFaxNo = !VendorFax & ""
     End If
     'preethi – V1.2.6 20110905 – CR99 - Add Option for Fare Type in EO
     If mblnConsTkt = False Then
       If !FareType = 1 Then
          optPublishedFare(0).value = True
       Else
         If !FareType = 2 Then
            optNettFare(1).value = True
         Else
            optPublishedFare(0).value = False
            optNettFare(1).value = False
         End If
       End If
    
     'preethi – V1.2.6 20110905 – CR98 - Reissue Ticket Box in EO
     If !OriTktNum <> "NULL" Then
        chkReissuedTkt.Enabled = True
        chkReissuedTkt.value = vbChecked
        txtOriTktNum.Enabled = True
        txtOriTktNum.Text = !OriTktNum
     Else
        chkReissuedTkt.Enabled = True
        chkReissuedTkt.value = vbUnchecked
        txtOriTktNum.Enabled = False
        txtOriTktNum.Text = ""
     End If
   End If
  End With
     rsEO.Close
     SetEOObj gobjPreEO
     Unload frmEOAmend
End Sub

Private Function fCommAmt(Fare As Single, CommPct As Single, Optional client As String) As Single
gobjLog.ProcedureName = "fCommAmt"
On Error GoTo ProcError

Dim sngPct As Single
Dim sngAmt As Single
sngPct = CommPct * 0.01

If UCase(gstrAgcyCountryCode) = "SG" Then

    cboClientType = client
    sngAmt = Fare * sngPct
    If mblnConsTkt Then
       fCommAmt = fCurrRound(sngAmt, gstrAgcyCurrCode, "DO")
    Else
       fCommAmt = Format(sngAmt, gstrAgcyCurrFormat)
    End If
    
ElseIf UCase(gstrAgcyCountryCode) = "HK" Then
    'Removed commission = 0 for TF and MN by JiYong
    'If cboClientType = "MN" Or cboClientType = "TF" Or cboClientType = "TP" Then
    If cboClientType = "TP" Then
        fCommAmt = 0
        txtCommission = "0"
    Else
        sngAmt = (Fare / (1 - sngPct)) - Fare
        'fCommAmt = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP") + IIf(sngAmt > 0, IIf(gstrAgcyCountryCode = "HK", IIf(cboClientType = "DU", 10, 0), 0), 0)
       fCommAmt = fCurrRound(sngAmt + IIf(sngAmt > 0, IIf(gstrAgcyCountryCode = "HK", IIf(cboClientType = "DU", 10, 0), 0), 0), gstrAgcyCurrCode)
    End If
Else
    fCommAmt = 0
    txtCommission = "0"
End If



Exit Function
ProcError:
    Call pErrorReport(True)


End Function
Private Function fDiscAmt(SellFare As Single, DiscPct As Single, CommAmt As Single, Optional client As String) As Single

gobjLog.ProcedureName = "fDiscAmt"
'On Error GoTo ProcError

Dim sngPct As Single
'Dim sngAmt As Single
sngPct = DiscPct * 0.01


If UCase(gstrAgcyCountryCode) = "SG" Then
    cboClientType = client
End If


'Select Case cboClientType
'Case "DU", "DB"
    '    fDiscAmt = fCurrRound(SellFare * sngPct, gstrAgcyCurrCode, "DOWN")
'Case "MN", "TF", "TP"
    '    fDiscAmt = CommAmt
'Case Else
    '    fDiscAmt = 0
'End Select
If UCase(gstrAgcyCountryCode) = "SG" Then
    Select Case cboClientType
    Case "DU", "DB", "MN", "TF", "TP"
             fDiscAmt = fCurrRound(SellFare * sngPct, gstrAgcyCurrCode, "DOWN")
    'Case "MN", "TF", "TP"
    '    fDiscAmt = CommAmt
    Case Else
        fDiscAmt = 0
    End Select
Else
    Select Case cboClientType
    Case "DU", "DB"
       fDiscAmt = fCurrRound(SellFare * sngPct, gstrAgcyCurrCode, "DOWN")
    Case "MN", "TF", "TP"
       fDiscAmt = CommAmt
    Case Else
       fDiscAmt = 0
    End Select
End If



Exit Function
ProcError:
    Call pErrorReport(True)


End Function

Private Function fMerchantFee(TotalCharge As Single, MerchFeePct As Single) As Single
gobjLog.ProcedureName = "fMerchantFee"
On Error GoTo ProcError

Dim sngPct As Single
Dim sngAmt As Single

sngPct = MerchFeePct * 0.01
sngAmt = CDec(TotalCharge * sngPct)
fMerchantFee = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP")

Exit Function
ProcError:
    Call pErrorReport(True)

End Function

Private Function validData() As Boolean
Dim strMsg As String
Dim bolEC As Boolean
Dim intX As Integer
If UCase(gstrAgcyCountryCode) = "HK" Then
   If txtNettFare.Text = "" Then txtNettFare.Text = "0"
Else
   If txtPubFare.Text = "" Then txtPubFare.Text = "0"
End If
If txtCost.Text = "" Then txtCost = "0"

'02062005
'If frmOthSvcs.datProducts.Recordset![ProductCode] = "00" Then
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
If mbolWebFareSelected = False Then
    If txtTktNum <> "0000000000" Then
    'If chkTktAtInv.value = 0 Then
       If IsNumeric(txtTktNum) = False Then
          strMsg = strMsg & "Invalid ticket number..." & Chr(13)
       End If
    End If
Else
   If Trim(txtTktNum.Text) = "" Then strMsg = "Need ticket number..." & Chr(13)
End If
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
If mbolWebFareSelected = False Then
    If Len(txtALCode) <> 3 Or IsNumeric(txtALCode) = False Then
        strMsg = strMsg & "Invalid Airline Code..." & Chr(13)
    End If
End If

If Trim(txtConjTkt) <> "" Then
   If Len(txtConjTkt) <> 2 Or IsNumeric(txtConjTkt) = False Then
      strMsg = strMsg & "Invalid Conjunction Ticket Number..." & Chr(13)
   End If
End If
'02062005
'If UCase(gstrAgcyCountryCode) = "HK" And chkTktAtInv.Value = 0 Then
'   If Len(txtALCode) <> 3 Or Len(txtTktNum) <> 10 Or IsNumeric(txtALCode) = False Or IsNumeric(txtTktNum) = False Then
'      strMsg = strMsg & "Invalid ticket number..." & Chr(13)
'   End If'
'
'End If
If lstFlights.SelCount = 0 Then strMsg = strMsg & "Need to select air segment(s) for this transaction..." & Chr(13)
If UCase(gstrAgcyCountryCode) = "HK" Then
    'If txtNettFare.Text = "0" Then strMsg = strMsg & "Need " & Replace(lblFare.Caption, ":", "") & "..." & Chr(13)
ElseIf UCase(gstrAgcyCountryCode) = "SG" Then
    If mblnConsTkt Then
       'If txtPubFare.Text = "0" Then strMsg = strMsg & "Need " & Replace(lblFare.Caption, ":", "") & "..." & Chr(13)
       If txtPubFare.Text = "" Then strMsg = strMsg & "Need " & Replace(lblFare.Caption, ":", "") & "..." & Chr(13)
    End If
    For intX = 0 To 1
       If (txtTax(intX) <> "0" And txtTax(intX) <> "") And txtTaxCode(intX) = "" Then
          strMsg = strMsg & "Need Tax Code " & intX + 1 & "..." & Chr(13)
       End If
    Next

    'Add on 200106: TMP card must absorb MF
    'If (cmbFOPType <> "INV" And Left(UCase(cmbCCType.Text), 2) = "DC" And _
        Left(UCase(txtCCNum.Text), 7) = "3644033") And chkAbsorb.value <> 1 Then
     'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    If (cmbFOPType <> "INV" And _
        IsTMPCard(Left(UCase(cmbCCType.Text), 2), UCase(txtCCNum.Text))) And _
        (chkAbsorb.value <> 1) Then
        strMsg = strMsg & "Need to tick aborbed merchant fee for TMP card and recalculate selling price" & Chr(13)
    End If

End If
'If txtPubFare.Text <> "" And CSng(txtPubFare.Text) < CSng(txtCost) Then strMsg = strMsg & "Published Fare should be higher or equal to nett cost..." & Chr(13)
'If mblnConsTkt Then
'   If txtSellPrice.Text = "" Or txtSellPrice.Text = "0" Then strMsg = strMsg & "Need to calculate Selling Price..." & Chr(13)
'Else
   If txtSellPrice.Text = "" Then strMsg = strMsg & "Need to calculate Selling Price..." & Chr(13)
'End If
If cmbFOPType.Text = "" Then strMsg = strMsg & "Need form of payment..." & Chr(13)
'JY – V1.2.3 20110418 – CR63 - Add validation if CC is selected as FOP
'If cmbFOPType.Text = "CX" Then
If cmbFOPType.Text = "CX" Or cmbFOPType.Text = "CC" Then
    If cmbCCType.Text = "" Then strMsg = strMsg & "Need valid credit vendor code..." & Chr(13)
    If txtCCNum = "" Then strMsg = strMsg & "Need valid credit card number..." & Chr(13)
    If LastDate(dtpCCExp.value) < Date Then strMsg = strMsg & "Need valid expiration date..." & Chr(13)
    If (txtCCNum.Text <> "" And cmbCCType.Text <> "") Then If ValidCCNum(cmbCCType.Text, txtCCNum.Text) = False Then strMsg = strMsg & "Credit card number is invalid or wrong card vendor selected ..." & Chr(13)
End If
'300908 Detect MI by CN
If gobjPNR.CompInfo.MI = True Then
If txtMI(0).Text = "" Then strMsg = strMsg & "Need Reference Fare (MI)..." & Chr(13)
If txtMI(1).Text = "" Then strMsg = strMsg & "Need Low Fare (MI)..." & Chr(13)
'CS Change EC
'If txtMI(2).Text = "" Then strMsg = strMsg & "Need Exception Code (MI)..." & Chr(13)
If txtRS.Text = "" Then strMsg = strMsg & "Need Realised Saving Code (MI)..." & Chr(13)
If txtMS.Text = "" Then strMsg = strMsg & "Need Missed Code (MI)..." & Chr(13)
If txtMI(3).Text = "" Then strMsg = strMsg & "Need Final Desitnation (MI)..." & Chr(13)

'Added on 18/2/2005
If txtMI(0).Text <> "" And txtMI(1) <> "" Then
    If CDec(txtMI(0)) < CDec(txtMI(1)) Then
        strMsg = strMsg & "Reference Fare(MI) must be greater than Low Fare(MI)..." & Chr(13)
    End If
End If


'Preethi - V1.1.1 20100831 - CR15 - Reference Fare and Low Fare Validation
 If mblnConsTkt = False Then
    If txtSellPrice.Text <> "" And txtMI(0).Text <> "" Then
       If CDec(txtMI(0)) < CDec(txtSellPrice) Then
          strMsg = strMsg & "Reference Fare must be greater than or equal to Selling Fare" & Chr(13)
       End If
    End If
    If txtSellPrice.Text <> "" And txtMI(1).Text <> "" Then
       If CDec(txtMI(1)) > CDec(txtSellPrice) Then
          strMsg = strMsg & "Low Fare must be lower than or equal to Selling Fare" & Chr(13)
       End If
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
bolEC = False
For intX = 1 To lvwRealECodes.ListItems.Count
    If txtRS = lvwRealECodes.ListItems.item(intX) Then
        bolEC = True
        Exit For
    End If
Next intX
If bolEC = False Then strMsg = strMsg & "Invalid Realised Saving Code..." & Chr(13)
bolEC = False
For intX = 1 To lvwMissECodes.ListItems.Count
    If txtMS = lvwMissECodes.ListItems.item(intX) Then
        bolEC = True
        Exit For
    End If
Next intX
If bolEC = False Then strMsg = strMsg & "Invalid Missed Saving Code..." & Chr(13)

End If


If mblnConsTkt = False Then
    If chkRequestMCO Then
        If txtRecLoc.Text = "" Then strMsg = strMsg & "Need RecLoc(MCO)..." & Chr(13)
        If txtTypeOfService.Text = "" Then strMsg = strMsg & "Need Type of Service(MCO)..." & Chr(13)
        If txtLoc.Text = "" Then strMsg = strMsg & "Need Location of Issuance(MCO)..." & Chr(13)
        If txtContactPerson.Text = "" Then strMsg = strMsg & "Need Contact Person(MCO)..." & Chr(13)
        If txtFOP.Text = "" Then strMsg = strMsg & "Need FOP(MCO)..." & Chr(13)
        If txtEquiAmt.Text = "" Then strMsg = strMsg & "Need Equivalent Amt Paid(MCO)..." & Chr(13)
        If txtROE.Text = "" Then strMsg = strMsg & "Need Rate of Exchange(MCO)..." & Chr(13)
        If txtHeadlineCurrency.Text = "" Then strMsg = strMsg & "Need Headline Currency(MCO)..." & Chr(13)
        If txtMCOTaxes.Text = "" Then strMsg = strMsg & "Need Taxes(MCO)..." & Chr(13)
        If lsvTraveller.ListItems.Count = 0 Then strMsg = strMsg & "Need Traveller Name(MCO)..." & Chr(13)
    End If
End If

Dim strTmp1() As String
Dim strTmp2 As String
Dim intTmpI As Integer
txtFaxNo = Trim(txtFaxNo)
'If InStr(1, txtFaxNo, " ") > 0 Then
'   strMsg = strMsg & "Fax number cannot accept space..." & Chr(13)
'Else
'   strTmp1 = Split(txtFaxNo, ",")
'   For intTmpI = 0 To UBound(strTmp1)
'      If IsNumeric(strTmp1(intTmpI)) = False Then
'         strMsg = strMsg & "Invalid fax number cannot accept space..." & Chr(13)
'      End If
'   Next
'End If
If cmdCalculate.Enabled Then
   strMsg = strMsg & "Please click calculate button..." & Chr(13)
End If

'MI Validation
'If txtMI(0) <> "" And txtMI(1) <> "" Then
'   strMsg = strMsg & validMI
'End If

'preethi – V1.2.6 20110905 – CR99 - Add Option for Fare Type in EO
If mblnConsTkt = False Then
   If optPublishedFare(0).value = False And optNettFare(1).value = False Then
       strMsg = strMsg & "Please indicate if you are billing a ticket issued on Published Fare or Marked Up Net Fare..." & Chr(13)
   End If
   If optNettFare(1).value = True And CDec(txtCommission) > 0 Then
      'JiYong – V1.2.6 20111003 – CR99 - Add Option for Fare Type in EO
      'Additional logic for CR99 - allow commission field to be input if client type = MG (For HK only)
      If Not (UCase(gstrAgcyCountryCode) = "HK" And cboClientType.Text = "MG") Then
         strMsg = strMsg & "Commission is not applicable for Marked Up Net Fare billing..." & Chr(13)
      End If
   End If
    'preethi – V1.2.6 20110905 – CR98 - Reissue Ticket Box in EO
    If chkReissuedTkt.value = vbChecked Then
       If Len(txtOriTktNum.Text) = 0 Then
          strMsg = strMsg & "Please Enter the Original Ticket Number" & Chr(13)
       End If
       If Len(txtOriTktNum.Text) <> 0 And (Len(txtOriTktNum.Text) < 10 Or (Not IsNumeric(txtOriTktNum.Text))) Then
           strMsg = strMsg & "Invalid Original Ticket Number" & Chr(13)
       End If
    End If
End If

If strMsg = "" Then
    validData = True
Else
    'cmdDone.Enabled = True
    'If gbolEOAmend = False Then cmdDone.Enabled = True
    'MsgBox strMsg, vbApplicationModal + vbExclamation, "Travel Pro"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"

End If

End Function










Private Sub Form_Unload(Cancel As Integer)
'Preethi - V1.1.1 20100831 - IR2 - Client MI screen is populated with old data
Unload frmClientMI
End Sub

Private Sub lstFlights_Click()

Dim intWebFareMatch As Integer
'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
If mblnConsTkt Then
   intWebFareMatch = checkWebFareSelected
   If intWebFareMatch > 0 And mbolEORetrieving = False Then
      populateWebFare (intWebFareMatch)
   Else
      txtALCode.Enabled = True
   End If
End If

End Sub

Private Sub lstRmks_Click()
Dim intC As Integer
With lstRmks

For intC = .ListCount - 1 To 0 Step -1

If .Selected(intC) = True Then
    txtFT = .List(intC)
End If
Next intC
End With
End Sub

Private Sub lstRmks_DblClick()
Dim intC As Integer
With lstRmks

For intC = .ListCount - 1 To 0 Step -1

If .Selected(intC) = True Then
    .RemoveItem intC
End If
Next intC
End With
End Sub

Private Sub lsvTraveller_DblClick()
lsvTraveller.ListItems.Remove (lsvTraveller.SelectedItem.Index)
End Sub

'CS Change EC
'Private Sub lvwECodes_DblClick()
'   txtMI(2).Text = lvwECodes.SelectedItem
'End Sub

Private Sub lvwMissECodes_ItemClick(ByVal item As MSComctlLib.ListItem)
txtMS = lvwMissECodes.SelectedItem
End Sub

Private Sub lvwRealECodes_ItemClick(ByVal item As MSComctlLib.ListItem)
txtRS = lvwRealECodes.SelectedItem
End Sub

Private Sub optCommType_Click(Index As Integer)
EnableCalculate
End Sub

Private Sub optDiscount_Click(Index As Integer)
EnableCalculate
End Sub

Private Sub txtCommission_Change()
EnableCalculate
End Sub

Private Sub txtCommission_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub



Private Sub txtDC_Change()
EnableCalculate
End Sub

Private Sub txtDC_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtEquiAmt_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtFreeRmk_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii, "#$*()/.: ?@")
End Sub







Private Sub txtHeadlineCurrency_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii)
End Sub

Private Sub txtMCOTaxes_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

'Added on 18/2/2005
Private Sub txtMI_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0, 1
        KeyAscii = fAllowNumeric(KeyAscii, ".")
    Case 2
        KeyAscii = fAllowNumeric(KeyAscii)
    Case 3
        KeyAscii = fAllowAlpha(KeyAscii)
    Case 6
        KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Select

End Sub



Private Sub txtNettFare_Change()
   EnableCalculate
End Sub

Private Sub txtNettFare_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtPubFare_Change()
   EnableCalculate
End Sub


Private Sub txtSellingPrice_Change()
   EnableCalculate
End Sub

Private Sub txtSellingPrice_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtTax_Change(Index As Integer)
   EnableCalculate
End Sub

Private Sub txtTax_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtTaxCode_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = fAllowAlpha(KeyAscii)
End Sub

Private Sub txtTrxnFee_GotFocus()
Call pSetSelected
End Sub

Private Sub txtALCode_GotFocus()
Call pSetSelected
End Sub

Private Sub txtCCNum_GotFocus()
Call pSetSelected
End Sub

Private Sub txtCommission_GotFocus()
Call pSetSelected
End Sub

Private Sub txtCost_GotFocus()
Call pSetSelected
End Sub

Private Sub txtDiscount_GotFocus()
Call pSetSelected
End Sub

Private Sub txtMerchFee_GotFocus()
Call pSetSelected
End Sub

Private Sub EnableCalculate()
   cmdCalculate.Enabled = True
End Sub

Private Sub txtPubFare_GotFocus()
Call pSetSelected
End Sub

Private Sub txtSellPrice_GotFocus()
Call pSetSelected
End Sub

Private Sub txtTax_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtTaxCode_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtTktNum_GotFocus()
Call pSetSelected
End Sub

Private Sub txtFreeRmk_GotFocus()
Call pSetSelected
End Sub


Sub FormCenter()
    Top = (Screen.Height * 0.95) / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
End Sub

Private Sub pSetInitialValues()
Dim strSQL As String
Dim rsECodes As New ADODB.Recordset
Dim strContract As String
Dim lngC As Long
Dim lngY As Long
Dim item As ListItem
'Modified on 18/2/2005
'Modified on 2/2/2005: add on client specific EC
'CS Change EC
'strSql = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='AS' OR tblExceptionCodes.ExceptionCodeGroup='AC') ORDER BY CAST(tblClientEC.EC AS integer)"
'CS Change EC
'strSQL = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND tblExceptionCodes.ProdType='" & "AIR" & "' AND tblExceptionCodes.ECInd='S' ORDER BY tblClientEC.EC"
strSQL = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.ClientID & " AND tblExceptionCodes.ProdType='AIR' and TBLCLIENTEC.PRODTYPE='AIR' ORDER BY tblClientEC.EC"

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

'Set rsECodes = gdbTPro.OpenRecordset(strSQL)
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
'    rsECodes.Close
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


'    cmbFOP(0).AddItem "MULTI"
'Added on 21/07/04
'populate combo box values for Class of Services, Trip Type
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
'cboMIFareType.AddItem "C- Client Negotiated Fares"
'cboMIFareType.AddItem "W- CWT Negotiated Fares"


'CS Remove FF26
'cboTripType.AddItem ""
'cboTripType.AddItem "Round"
'cboTripType.AddItem "One Way"
'CS Add FF26
cboTrip.AddItem "INTERNATIONAL"
cboTrip.AddItem "DOMESTIC"
cboTrip.listindex = 0
        
mstrBookingTool = GetBookingTool

'If mstrBookingTool <> "" Then

'   cboBookingAction.AddItem "Agent Booked"
'   cboBookingAction.AddItem "Self Booked"
'   cboBookingAction.AddItem "Air Modified"
'   cboBookingAction.Text = "Self Booked"

'Else

'  cboBookingAction.AddItem "Agent Booked"
'  cboBookingAction.AddItem "Self Booked"
'  cboBookingAction.AddItem "Air Modified"
'  cboBookingAction.Text = "Agent Booked"
  
'End If
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
        
'300908 Detect MI by CN
If gobjPNR.CompInfo.MI = False Then
    fraMI.Enabled = False
    cmbDispNum.Enabled = False
    lblLabels(0).Enabled = False
    lblLabels(28).Enabled = False
    lblLabels(27).Enabled = False
    lblLabels(30).Enabled = False
    lblLabels(36).Enabled = False
    lblLabels(32).Enabled = False
    lblLabels(26).Enabled = False
    lblLabels(41).Enabled = False
    lblLabels(40).Enabled = False
    lblLabels(39).Enabled = False
    txtMI(0).Enabled = False
    txtMI(1).Enabled = False
    txtMI(3).Enabled = False
    txtMI(6).Enabled = False
    txtMI(0).Enabled = False
    chkPaperTkt.Enabled = False
    cboBookingAction.Enabled = False
    txtRS.Enabled = False
    txtMS.Enabled = False
    lvwRealECodes.Enabled = False
    lvwMissECodes.Enabled = False
    cboClassServ.Enabled = False
End If
        
End Sub

'02062005
Private Sub UpdatePO()
Dim i, j As Integer
Dim intTemp As Integer
Dim strTemp As String
Dim strSQL As String
Dim rsPC As ADODB.Recordset
Dim strPC As String
Dim rsDS As ADODB.Recordset
Dim strZero As String
Dim intLength As Integer
'modified on 15Jun
Set rsDS = gdbConn.Execute("Select Length from tbldocstruct where StructID='PO'")
While Not rsDS.EOF

intLength = intLength + rsDS![length]

rsDS.MoveNext
Wend


Set rsDS = Nothing

For i = 1 To intLength
strZero = strZero & "0"
Next

'With gobjPNR
'   For i = 1 To .GASaleRecordCount
'      If .AcctRemarkCount > 1 And .GASalesRecord(i).BegLine >= 1 And .GASalesRecord(i).ProductCode = "02" Then
'         For J = .GASalesRecord(i).BegLine To .GASalesRecord(i).EndLine
'             intTemp = InStr(1, .AcctRemark(J).RemarkText, "/PO" & IIf(mstrTktNum <> "", mstrTktNum, txtEONum))
'             If intTemp <> 0 Then
'                strTemp = Mid(.AcctRemark(J).RemarkText, 1, intTemp - 1) & "/PO" & Format(txtALCode, "000") & txtTktNum & Mid(.AcctRemark(J).RemarkText, intTemp + 16)
'                strTemp = "DI." & J & "@FT-" & Replace(strTemp, "@", ".")
'                gobjHost.TerminalEntry strTemp
'             Else
'                intTemp = InStr(1, .AcctRemark(J).RemarkText, "/TK" & IIf(mstrTktNum <> "", mstrTktNum, txtEONum))
'                If intTemp <> 0 Then
'                   strTemp = Mid(.AcctRemark(J).RemarkText, 1, intTemp - 1) & "/TK" & Format(txtALCode, "000") & Mid(.AcctRemark(J).RemarkText, intTemp + 3)
'                   strTemp = "DI." & J & "@FT-" & Replace(strTemp, "@", ".")
'                   gobjHost.TerminalEntry strTemp
'                End If
'             End If
'         Next
'      End If
'   Next
'End With



With gobjPNR
   For i = 1 To .GASaleRecordCount
    strPC = ""
    strSQL = "Select Type from tblProductCodes where SortKey='" & .GASalesRecord(i).ProductCode & "' "
    Set rsPC = gdbConn.Execute(strSQL)
    If Not rsPC.EOF Then
    strPC = rsPC!Type
    End If
    
      If .AcctRemarkCount > 1 And .GASalesRecord(i).BegLine >= 1 And (strPC = "CT") Then
        If InStr(1, .AcctRemark(.GASalesRecord(i).BegLine).RemarkText, "TK" & txtEONum) <> 0 Or _
           InStr(1, .AcctRemark(.GASalesRecord(i).BegLine).RemarkText, "TK" & Left(mstrTktNum, 3) & txtEONum) <> 0 Then
         For j = .GASalesRecord(i).BegLine To .GASalesRecord(i).EndLine

             intTemp = InStr(1, .AcctRemark(j).RemarkText, "/PO" & IIf(Len(mstrTktNum) > 3, Right(mstrTktNum, intLength), strZero))
             If intTemp <> 0 Then
                'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
                If mbolWebFareSelected = False Then
                    strTemp = Mid(.AcctRemark(j).RemarkText, 1, intTemp - 1) & "/PO" & Right(txtALCode & txtTktNum, intLength) & Mid(.AcctRemark(j).RemarkText, intTemp + intLength + 3)
                Else
                    strTemp = Mid(.AcctRemark(j).RemarkText, 1, intTemp - 1) & "/PO" & Right(txtALCode & txtTktNum, intLength) & Mid(.AcctRemark(j).RemarkText, intTemp + Len(mstrTktNum) + 3)
                End If
                'strTemp = Mid(.AcctRemark(J).RemarkText, 1, intTemp - 1) & "/PO" & Format(txtALCode, "000") & txtTktNum & Mid(.AcctRemark(J).RemarkText, intTemp + 16)
                strTemp = "DI." & j & "@FT-" & Replace(strTemp, "@", ".")
                gobjHost.terminalEntry strTemp
             Else
                'intTemp = InStr(1, .AcctRemark(J).RemarkText, "/TK" & IIf(mstrTktNum <> "", mstrTktNum, txtEONum))
                intTemp = InStr(1, .AcctRemark(j).RemarkText, "/TK" & IIf(mstrTktNum <> "", Left(mstrTktNum, 3) & txtEONum, txtEONum))
                If intTemp <> 0 Then
                   If InStr(1, .AcctRemark(.GASalesRecord(i).BegLine).RemarkText, "TK" & txtEONum) Then
                      strTemp = Mid(.AcctRemark(j).RemarkText, 1, intTemp - 1) & "/TK" & Format(txtALCode, "000") & Mid(.AcctRemark(j).RemarkText, intTemp + 3)
                   Else
                      strTemp = Mid(.AcctRemark(j).RemarkText, 1, intTemp - 1) & "/TK" & Format(txtALCode, "000") & Mid(.AcctRemark(j).RemarkText, intTemp + 6)
                   End If
                   strTemp = "DI." & j & "@FT-" & Replace(strTemp, "@", ".")
                   gobjHost.terminalEntry strTemp
                End If
             End If
         Next
       End If
      End If
      rsPC.Close
      Set rsPC = Nothing
   Next
End With

gobjHost.terminalEntry "R.XO+ER"
'gobjHost.TerminalEntry "ER"
gobjHost.terminalEntry "ER"
gobjHost.terminalEntry "ER"

strSQL = "Update tblExchangeOrder "
strSQL = strSQL & "Set tktNum = '" & Format(txtALCode, "000") & txtTktNum & "' "
strSQL = strSQL & "Where exchangeid = '" & txtEONum & "'"

gdbConn.Execute strSQL

End Sub


Private Function MaxMIDispNum() As Integer
    Dim i As Integer
    Dim strDIText() As String
    Dim strTmp As String
    Dim intMaxDisp As Integer
     
    With gobjPNR
       For i = 1 To .AcctRemarkCount
           strTmp = .AcctRemark(i).RemarkText
           'CS Remove FF26, Add FF41
           If Left(strTmp, 2) = "RF" Or Left(strTmp, 2) = "LF" Or _
              Left(strTmp, 2) = "EC" Or Left(strTmp, 3) = "FF7" Or _
              Left(strTmp, 3) = "FF8" Or Left(strTmp, 4) = "FF81" Or _
              Left(strTmp, 4) = "FF38" Or Left(strTmp, 4) = "FF34" Then
              strDIText = Split(strTmp, "/")
              If UBound(strDIText) = 2 Then
                 If Left(strDIText(1), 1) = "*" And IsNumeric(Mid(strDIText(1), 2)) Then
                    If intMaxDisp < Mid(strDIText(1), 2) Then
                       intMaxDisp = Mid(strDIText(1), 2)
                    End If
                 End If
              End If
           End If
       Next
       
       If intMaxDisp <= 0 Then
          MaxMIDispNum = 0
          Exit Function
       Else
          MaxMIDispNum = intMaxDisp
       End If
       
       ReDim mRF(intMaxDisp - 1)
       ReDim mLF(intMaxDisp - 1)
       'CS Change EC
       'ReDim mEC(intMaxDisp - 1)
       ReDim mRS(intMaxDisp - 1)
       ReDim mMS(intMaxDisp - 1)
       ReDim mFF7(intMaxDisp - 1)
       ReDim mFF8(intMaxDisp - 1)
       ReDim mFF81(intMaxDisp - 1)
       ReDim mFF34(intMaxDisp - 1)
       'CS Remove FF26
       'ReDim mFF26(intMaxDisp - 1)
       'CS Add FF41
       'ReDim mFF41(intMaxDisp - 1)
       ReDim mFF38(intMaxDisp - 1)
    
       For i = 1 To .AcctRemarkCount
           strTmp = .AcctRemark(i).RemarkText
           'CS Remove FF26, Add FF41, Add FF30
           If Left(strTmp, 2) = "RF" Or Left(strTmp, 2) = "LF" Or _
              Left(strTmp, 2) = "EC" Or Left(strTmp, 3) = "FF7" Or _
              Left(strTmp, 3) = "FF8" Or Left(strTmp, 4) = "FF81" Or _
              Left(strTmp, 4) = "FF38" Or Left(strTmp, 4) = "FF30" Or Left(strTmp, 4) = "FF34" Then
              strDIText = Split(strTmp, "/")
              If UBound(strDIText) = 2 Then
                 If Left(strDIText(1), 1) = "*" And IsNumeric(Mid(strDIText(1), 2)) Then
                    Select Case strDIText(0)
                       Case "RF"
                          mRF(Mid(strDIText(1), 2) - 1) = Replace(strDIText(2), "@", ".")
                       Case "LF"
                          mLF(Mid(strDIText(1), 2) - 1) = Replace(strDIText(2), "@", ".")
                       'CS Change EC
                       'Case "EC"
                       '   mEC(Mid(strDIText(1), 2) - 1) = Replace(strDIText(2), "@", ".")
                       Case "EC"
                          mMS(Mid(strDIText(1), 2) - 1) = Replace(strDIText(2), "@", ".")
                       Case "FF30"
                          mRS(Mid(strDIText(1), 2) - 1) = Replace(strDIText(2), "@", ".")
                       Case "FF7"
                          mFF7(Mid(strDIText(1), 2) - 1) = strDIText(2)
                       Case "FF8"
                          mFF8(Mid(strDIText(1), 2) - 1) = strDIText(2)
                       Case "FF81"
                          mFF81(Mid(strDIText(1), 2) - 1) = strDIText(2)
                       'CS Remove FF26
                       'Case "FF26"
                       '   mFF26(Mid(strDIText(1), 2) - 1) = strDIText(2)
                       'Add FF 41
                       'Case "FF41"
                       '   mFF41(Mid(strDIText(1), 2) - 1) = strDIText(2)
                       Case "FF38"
                          mFF38(Mid(strDIText(1), 2) - 1) = strDIText(2)
                       Case "FF34"
                          mFF34(Mid(strDIText(1), 2) - 1) = strDIText(2)
                    End Select
                 End If
              End If
           End If
       Next
    End With
End Function
Private Function matchBookingTool(ByVal prefix As String) As String
Dim i As Integer
matchBookingTool = ""

    Select Case (UCase(Trim(prefix)))
    Case "AB"
       matchBookingTool = "AB - Agent Booked"
    Case "EB"
       matchBookingTool = "EB - Self Booked"
    Case "AA"
       matchBookingTool = "AA - Air Modified"
    Case "AM"
       matchBookingTool = "AM - Multiple Modification"
    End Select
End Function
Private Function matchList(ByVal prefix As String) As String
Dim i As Integer
matchList = ""
For i = 0 To cboClassServ.ListCount - 1
    If Left(cboClassServ.List(i), 2) = UCase(Trim(prefix)) Then
       matchList = cboClassServ.List(i)
       Exit For
    End If
Next
End Function

Private Function validMI() As String
Dim sngSF As Single
Dim strProductType As String
Dim strProduct As String
Dim strMsg As String
Dim sngTax As Single

strMsg = ""
strProductType = frmOthSvcs.datProducts.Recordset![Type]
strProduct = frmOthSvcs.dbcProducts.BoundText

sngTax = IIf(txtTax(0).Text = "", 0, fConvertZero(txtTax(0).Text)) + IIf(txtTax(1).Text = "", 0, fConvertZero(txtTax(1).Text))
If strProductType = "CT" Then
   sngSF = fConvertZero(txtSellPrice.Text) - sngTax + IIf(txtDC = "", 0, txtDC)
ElseIf strProductType = "BT" And (strProduct = "00" Or strProduct = "01") Then
   If gstrAgcyCountryCode = "HK" Then
      sngSF = fConvertZero(txtSellPrice.Text) - sngTax + IIf(txtDC = "", 0, txtDC)
   Else
      sngSF = fConvertZero(txtSellingPrice) + IIf(txtMerchFee = "", 0, txtMerchFee)
   End If
End If

'Deduct discount amount from Selling Price
If txtDC <> "" Then
   sngSF = sngSF - txtDC
End If

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

validMI = strMsg

End Function
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
Private Sub Log()
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
End Sub

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Private Function checkWebFareSelected() As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim bolMatch As Boolean
    Dim objSegSelected As FareOptionSegments
        
    mbolWebFareSelected = False
    mintWebFareSelected = 0
    
    'Get flight selected
    Set objSegSelected = GetSelectedSegment
    
    If objSegSelected.FareOptionSegmentCount = 0 Then Exit Function
    
    'Compare flight selected with segments in Web Fares
    For i = 1 To mobjWebFares.WebFareCount
        If mobjWebFares.WebFare(i).AirSegCount = objSegSelected.FareOptionSegmentCount Then
           bolMatch = True
           For j = 1 To mobjWebFares.WebFare(i).AirSegCount
                With mobjWebFares.WebFare(i).AirSeg(j)
                     'Match Vendor, Flight Number, Class, Departure Date, Origin, Destination, Departure Time, Arrival Time
                     If UCase(mobjWebFares.WebFare(i).PlatingCarrier) = UCase(objSegSelected.FareOptionSegment(j).Carrier) And _
                        UCase(.FlightNum) = UCase(objSegSelected.FareOptionSegment(j).FlightNum) And _
                        UCase(.Class) = UCase(objSegSelected.FareOptionSegment(j).Class) And _
                        UCase(.DepDate) = UCase(objSegSelected.FareOptionSegment(j).DepDate) And _
                        UCase(.DepCity) = UCase(objSegSelected.FareOptionSegment(j).DepCity) And _
                        UCase(.ArrCity) = UCase(objSegSelected.FareOptionSegment(j).ArrCity) And _
                        UCase(.DepTime) = UCase(objSegSelected.FareOptionSegment(j).DepTime) And _
                        UCase(.ArrTime) = UCase(objSegSelected.FareOptionSegment(j).ArrTime) Then
                     Else
                        bolMatch = False
                        Exit For
                     End If
                End With
           Next
           If bolMatch = True Then
             checkWebFareSelected = i
             mintWebFareSelected = i
             mbolWebFareSelected = True
             Exit For
           End If
        End If
    Next
                
End Function

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Private Function GetSelectedSegment() As FareOptionSegments
    
    Dim i As Integer
    Dim objAirSeg As New FareOptionSegment
    Dim objAirSegs As New FareOptionSegments
    
    For i = 0 To lstFlights.ListCount - 1
        If lstFlights.Selected(i) = True Then
           'Populate into collection if flight is selected
           Set objAirSeg = New FareOptionSegment
           With gobjPNR.AirSeg(i + 1)
                objAirSeg.Carrier = .Vendor
                objAirSeg.FlightNum = .FlightNumber
                objAirSeg.Class = .Class
                objAirSeg.DepDate = Format(.DepartDateTime, "ddmmm")
                objAirSeg.DepCity = .DepartAirport
                objAirSeg.ArrCity = .ArriveAirport
                objAirSeg.DepTime = Format(.DepartDateTime, "hhnn")
                objAirSeg.ArrTime = Format(.ArriveDateTime, "hhnn")
                objAirSegs.AddFareOptionSegment objAirSeg
           End With
        End If
    Next
    Set GetSelectedSegment = objAirSegs
    
End Function

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Private Sub populateWebFare(intMatch As Integer)

    'Clear the value in textbox
    
    txtTax(0).Text = "0"
    txtTaxCode(0).Text = ""
    txtTax(1).Text = "0"
    txtTaxCode(1).Text = ""
    txtCost.Text = ""
    txtCommission.Text = "0"
    txtDC.Text = "0"
    txtMerchFee.Text = ""
    
    txtALCode.Text = ""
    txtALCode.Enabled = False
        
    If UCase(gstrAgcyCountryCode) = "SG" Then
       txtPubFare.Text = mobjWebFares.WebFare(intMatch).BaseFare
    ElseIf UCase(gstrAgcyCountryCode) = "HK" Then
       txtNettFare.Text = mobjWebFares.WebFare(intMatch).BaseFare
    End If
    
    txtTax(0).Text = mobjWebFares.WebFare(intMatch).Tax
    txtTaxCode(0).Text = mobjWebFares.WebFare(intMatch).TaxCode
    txtTktNum.Text = mobjWebFares.WebFare(intMatch).ConfirmationNum
    cmdCalculate_Click
    
End Sub

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration

Private Sub populateLowFare()

    Dim i As Integer
    Dim j As Integer
    Dim strLowFareCarrier As String
    Dim curLowFare As Currency
    Dim bolMatch As Boolean
    Dim objSegSelected As FareOptionSegments
    
    If gobjPNR.CompInfo.MI = False Then Exit Sub
    
    'Get flight selected
    Set objSegSelected = GetSelectedSegment
    
    curLowFare = CDec(txtSellPrice.Text)
    
    For i = 1 To objSegSelected.FareOptionSegmentCount
        With objSegSelected.FareOptionSegment(i)
             If i = 1 Then
                strLowFareCarrier = .Carrier
             Else
                If UCase(strLowFareCarrier) <> UCase(.Carrier) Then
                   'Populate low fare carrier to blank if different carriers selected
                    strLowFareCarrier = ""
                    Exit For
                End If
             End If
        End With
    Next
        
    If objSegSelected.FareOptionSegmentCount = 0 Then Exit Sub
    
    'Compare flight selected with segments in Declined Web Fares
    For i = 1 To mobjDeclinedWebFares.WebFareCount
        If mobjDeclinedWebFares.WebFare(i).AirSegCount = objSegSelected.FareOptionSegmentCount Then
           bolMatch = True
           For j = 1 To mobjDeclinedWebFares.WebFare(i).AirSegCount
                With mobjDeclinedWebFares.WebFare(i).AirSeg(j)
                     'Match Departure Date, Origin, Destination
                     If UCase(.DepDate) = UCase(objSegSelected.FareOptionSegment(j).DepDate) And _
                        UCase(.DepCity) = UCase(objSegSelected.FareOptionSegment(j).DepCity) And _
                        UCase(.ArrCity) = UCase(objSegSelected.FareOptionSegment(j).ArrCity) Then
                     Else
                        bolMatch = False
                        Exit For
                     End If
                End With
           Next
           If bolMatch = True Then
              'Get the lowest fare and and carrier attached
              With mobjDeclinedWebFares.WebFare(i)
                If CDec(curLowFare) > CDec(.BaseFare) + CDec(.Tax) Then
                   strLowFareCarrier = .PlatingCarrier
                   curLowFare = CDec(.BaseFare) + CDec(.Tax)
                End If
              End With
           End If
        End If
    Next
    
    'Populate low fare carrier and low fare
    txtMI(6).Text = strLowFareCarrier
    txtMI(1).Text = CDec(curLowFare)
    
End Sub

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Private Function getRouting(strRouting As String) As String
    
    Dim strAry() As String
    Dim strDelimiter As String
    Dim strTemp As String
    Dim i As Integer
    
    strDelimiter = "-"
        
    strAry = Split(strRouting, strDelimiter)
    
    For i = 0 To UBound(strAry)
        strTemp = fGetCityNameOnly(Trim(strAry(i)))
        If strTemp = "" Then
           strTemp = UCase(Trim(strAry(i)))
        End If
        If i = 0 Then
            getRouting = strTemp
        Else
            getRouting = getRouting & "/" & strTemp
        End If
    Next
    
End Function

'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
Private Sub updateTktInRI()

Dim i As Integer

For i = 1 To gobjPNR.ItinRemarkCount
    With gobjPNR.ItinRemark(i)
        If .RemarkText = "LCC CONFIRMATION NUMBER: " & mstrTktNum Then
            gobjHost.terminalEntry "RI." & .ItemNum & "@" & "LCC CONFIRMATION NUMBER: " & txtTktNum.Text
            Exit For
        End If
    End With
Next

End Sub
