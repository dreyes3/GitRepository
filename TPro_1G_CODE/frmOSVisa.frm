VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOSVisa 
   ClientHeight    =   7650
   ClientLeft      =   1800
   ClientTop       =   6645
   ClientWidth     =   11160
   LinkTopic       =   "CWT Travel Pro - Visa"
   ScaleHeight     =   7650
   ScaleWidth      =   11160
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   300
      TabIndex        =   29
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Visa Info"
      TabPicture(0)   =   "frmOSVisa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraHotel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtContact"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Remarks"
      TabPicture(1)   =   "frmOSVisa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFreeRmk"
      Tab(1).Control(1)=   "cmdFreeRmkToEO"
      Tab(1).Control(2)=   "cmdFreeRmkToItin"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(5)=   "Label5"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Vendor Info"
      TabPicture(2)   =   "frmOSVisa.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraVendorInfo"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "MI"
      TabPicture(3)   =   "frmOSVisa.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdClientMI"
      Tab(3).ControlCount=   1
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
         Left            =   -74760
         TabIndex        =   86
         Top             =   480
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Height          =   3795
         Left            =   240
         TabIndex        =   68
         Top             =   1155
         Width           =   4815
         Begin VB.TextBox txtCWTHandlingFee 
            Height          =   315
            Left            =   1980
            TabIndex        =   97
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtCWTHandlingMF 
            BackColor       =   &H8000000B&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3600
            TabIndex        =   95
            Top             =   1080
            Width           =   735
         End
         Begin VB.CheckBox chkCWTHandlingMF 
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
            Left            =   3360
            TabIndex        =   94
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtNettMF 
            BackColor       =   &H8000000B&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3600
            TabIndex        =   93
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkNettMerchant 
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
            Left            =   3360
            TabIndex        =   92
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtVendorHandling 
            Height          =   315
            Left            =   1980
            TabIndex        =   91
            Top             =   720
            Width           =   1215
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
            Left            =   3240
            TabIndex        =   77
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox txtGST 
            Height          =   315
            Left            =   1980
            TabIndex        =   76
            Top             =   2580
            Visible         =   0   'False
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
            Left            =   1980
            TabIndex        =   75
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtMerchFee 
            Height          =   315
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   2940
            Width           =   1215
         End
         Begin VB.TextBox txtCommission 
            Height          =   315
            Left            =   1980
            TabIndex        =   73
            Top             =   2235
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            Height          =   315
            Left            =   1980
            TabIndex        =   72
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtSellPrice 
            Height          =   315
            Left            =   1980
            TabIndex        =   71
            Top             =   3300
            Width           =   1215
         End
         Begin VB.TextBox txtGrossSale 
            Height          =   315
            Left            =   1980
            TabIndex        =   70
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkGSTAbsorb 
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
            Height          =   375
            Left            =   3240
            TabIndex        =   69
            Top             =   2520
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label_HKMerchantFee_1 
            Alignment       =   2  'Center
            Caption         =   "Merchant Fee"
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
            Left            =   3240
            TabIndex        =   96
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label_VenderHandlingFee 
            Alignment       =   1  'Right Justify
            Caption         =   "Vendor Handling :"
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
            TabIndex        =   90
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label Label_CWTHandlingFee 
            Alignment       =   1  'Right Justify
            Caption         =   "CWT Handling :"
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
            TabIndex        =   87
            Top             =   1080
            Width           =   1785
         End
         Begin VB.Label lblGST 
            Alignment       =   1  'Right Justify
            Caption         =   "GST:"
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
            TabIndex        =   83
            Top             =   2580
            Visible         =   0   'False
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
            Left            =   360
            TabIndex        =   82
            Top             =   2940
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
            Left            =   360
            TabIndex        =   81
            Top             =   2280
            Width           =   1545
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nett Cost:"
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
            TabIndex        =   80
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Selling Price in DI:"
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
            TabIndex        =   79
            Top             =   3300
            Width           =   1785
         End
         Begin VB.Label Label26 
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
            Left            =   360
            TabIndex        =   78
            Top             =   1440
            Width           =   1545
         End
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
         Height          =   600
         Left            =   6480
         TabIndex        =   65
         Top             =   4080
         Width           =   3795
         Begin VB.TextBox txtPassengerID 
            Height          =   315
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   66
            Top             =   200
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
            TabIndex        =   67
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.Frame fraVendorInfo 
         Height          =   5895
         Left            =   -74640
         TabIndex        =   46
         Top             =   675
         Width           =   9975
         Begin VB.TextBox txtReplyEmail 
            Height          =   375
            Left            =   240
            TabIndex        =   84
            Top             =   3960
            Width           =   8655
         End
         Begin VB.TextBox txtTel 
            Height          =   375
            Left            =   960
            TabIndex        =   63
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox txtCreditTerms 
            Height          =   420
            Left            =   4080
            TabIndex        =   60
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtVendor 
            Height          =   375
            Left            =   960
            TabIndex        =   53
            Top             =   360
            Width           =   7935
         End
         Begin VB.TextBox txtAddress1 
            Height          =   375
            Left            =   960
            TabIndex        =   52
            Top             =   840
            Width           =   7935
         End
         Begin VB.TextBox txtAddress2 
            Height          =   375
            Left            =   960
            TabIndex        =   51
            Top             =   1320
            Width           =   7935
         End
         Begin VB.TextBox txtCity 
            Height          =   375
            Left            =   960
            TabIndex        =   50
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtCountry1 
            Height          =   375
            Left            =   3840
            TabIndex        =   49
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   960
            TabIndex        =   48
            Top             =   2280
            Width           =   7935
         End
         Begin VB.TextBox txtFaxNo 
            Height          =   375
            Left            =   960
            TabIndex        =   47
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "Reply Email in EO (Only 1 email address is allowed)"
            Height          =   375
            Left            =   240
            TabIndex        =   85
            Top             =   3720
            Width           =   4095
         End
         Begin VB.Label Label14 
            Caption         =   "Contact No."
            Height          =   375
            Left            =   240
            TabIndex        =   62
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Credit Terms"
            Height          =   255
            Left            =   3000
            TabIndex        =   61
            Top             =   2760
            Width           =   945
         End
         Begin VB.Label Label15 
            Caption         =   "Vendor "
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Address"
            Height          =   375
            Left            =   240
            TabIndex        =   58
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "City"
            Height          =   375
            Left            =   240
            TabIndex        =   57
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Country"
            Height          =   375
            Left            =   3120
            TabIndex        =   56
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Email (;)"
            Height          =   375
            Left            =   240
            TabIndex        =   55
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Fax No (,)"
            Height          =   375
            Left            =   240
            TabIndex        =   54
            Top             =   2760
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Visa Number: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6720
         TabIndex        =   44
         Top             =   3120
         Width           =   3495
         Begin VB.TextBox txtVisaNum 
            Height          =   315
            Left            =   240
            MaxLength       =   10
            TabIndex        =   45
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.TextBox txtContact 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   43
         Top             =   840
         Width           =   2775
      End
      Begin VB.Frame Frame6 
         Height          =   1575
         Left            =   6480
         TabIndex        =   36
         Top             =   4680
         Width           =   3795
         Begin VB.TextBox txtEONum 
            Height          =   315
            Left            =   1920
            TabIndex        =   13
            Top             =   480
            Width           =   1815
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
            Left            =   120
            Picture         =   "frmOSVisa.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
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
            Left            =   1920
            TabIndex        =   15
            Top             =   960
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
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1695
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
         Left            =   240
         TabIndex        =   34
         Top             =   5115
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
            Left            =   2580
            TabIndex        =   4
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
            TabIndex        =   2
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
            ItemData        =   "frmOSVisa.frx":04B2
            Left            =   240
            List            =   "frmOSVisa.frx":04BC
            Style           =   2  'Dropdown List
            TabIndex        =   0
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
            ItemData        =   "frmOSVisa.frx":04C9
            Left            =   1800
            List            =   "frmOSVisa.frx":04E5
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   300
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtpCCExp 
            Height          =   360
            Left            =   4680
            TabIndex        =   3
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
            Format          =   16973827
            CurrentDate     =   36526
            MaxDate         =   73050
            MinDate         =   36526
         End
      End
      Begin VB.Frame fraHotel 
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
         Left            =   5160
         TabIndex        =   33
         Top             =   480
         Width           =   5115
         Begin VB.TextBox txtDay 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   1
            EndProperty
            Height          =   345
            Left            =   3840
            MaxLength       =   2
            TabIndex        =   88
            Top             =   1980
            Width           =   495
         End
         Begin VB.ComboBox cmbProcess 
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
            ItemData        =   "frmOSVisa.frx":0509
            Left            =   1680
            List            =   "frmOSVisa.frx":0513
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1980
            Width           =   2115
         End
         Begin VB.ComboBox cmbValidityPeriod 
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
            ItemData        =   "frmOSVisa.frx":0528
            Left            =   2460
            List            =   "frmOSVisa.frx":0535
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1560
            Width           =   1935
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   345
            Left            =   2070
            TabIndex        =   9
            Top             =   1560
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtValidity"
            BuddyDispid     =   196675
            OrigLeft        =   2880
            OrigTop         =   1560
            OrigRight       =   3120
            OrigBottom      =   1905
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtValidity 
            Alignment       =   1  'Right Justify
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
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   8
            Tag             =   "NN"
            Text            =   "0"
            Top             =   1560
            Width           =   375
         End
         Begin VB.ComboBox cmbEntries 
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
            ItemData        =   "frmOSVisa.frx":054E
            Left            =   1680
            List            =   "frmOSVisa.frx":055B
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1150
            Width           =   2715
         End
         Begin VB.ComboBox cmbVisaType 
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
            ItemData        =   "frmOSVisa.frx":0579
            Left            =   1680
            List            =   "frmOSVisa.frx":058C
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   2715
         End
         Begin VB.TextBox txtCountry 
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
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   5
            Tag             =   "NN"
            Top             =   300
            Width           =   3165
         End
         Begin VB.Label Label_Days 
            Caption         =   "Days"
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
            Left            =   4440
            TabIndex        =   89
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Processing:"
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
            TabIndex        =   41
            Top             =   1980
            Width           =   1305
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Validity:"
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
            TabIndex        =   40
            Top             =   1560
            Width           =   1305
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Entries:"
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
            TabIndex        =   39
            Top             =   1140
            Width           =   945
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Type:"
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
            TabIndex        =   38
            Top             =   720
            Width           =   1305
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Country:"
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
            TabIndex        =   37
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   1455
         Left            =   5160
         Picture         =   "frmOSVisa.frx":05CF
         ScaleHeight     =   1395
         ScaleWidth      =   1155
         TabIndex        =   32
         Top             =   3120
         Width           =   1215
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
         TabIndex        =   21
         Tag             =   "NN"
         Top             =   3495
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
         Picture         =   "frmOSVisa.frx":6911
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Add Free Text to Exchange Order Remarks"
         Top             =   3375
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
         Picture         =   "frmOSVisa.frx":6D53
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Add Free Text to Itinerary Remarks"
         Top             =   3375
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
         TabIndex        =   31
         Top             =   3975
         Width           =   10275
         Begin VB.ListBox lstItinRmks 
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   24
            Top             =   300
            Width           =   4755
         End
         Begin VB.ListBox lstItinRmks 
            Height          =   2205
            Index           =   1
            Left            =   5400
            MultiSelect     =   1  'Simple
            TabIndex        =   28
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
            Picture         =   "frmOSVisa.frx":7195
            Style           =   1  'Graphical
            TabIndex        =   26
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
            Picture         =   "frmOSVisa.frx":75D7
            Style           =   1  'Graphical
            TabIndex        =   25
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
            Picture         =   "frmOSVisa.frx":7A19
            Style           =   1  'Graphical
            TabIndex        =   27
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
         TabIndex        =   30
         Top             =   735
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
            Picture         =   "frmOSVisa.frx":7E5B
            Style           =   1  'Graphical
            TabIndex        =   19
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
            Picture         =   "frmOSVisa.frx":829D
            Style           =   1  'Graphical
            TabIndex        =   17
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
            Picture         =   "frmOSVisa.frx":86DF
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Add All Remarks"
            Top             =   1080
            Width           =   495
         End
         Begin VB.ListBox lstEORmks 
            Height          =   2205
            Index           =   1
            Left            =   5400
            MultiSelect     =   1  'Simple
            TabIndex        =   20
            Top             =   300
            Width           =   4755
         End
         Begin VB.ListBox lstEORmks 
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   300
            Width           =   4755
         End
      End
      Begin VB.Label Label13 
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
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   840
         Width           =   2175
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
         TabIndex        =   35
         Top             =   3555
         Width           =   1545
      End
   End
   Begin MSAdodcLib.Adodc datVendors 
      Height          =   375
      Left            =   4560
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
      Bindings        =   "frmOSVisa.frx":8B21
      DataSource      =   "datVendors"
      Height          =   360
      Left            =   360
      TabIndex        =   64
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
Attribute VB_Name = "frmOSVisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsRmks As ADODB.Recordset
Dim mobjEO As EO
'Timer
'Dim StartTime As Date
Dim mstrPCAmend As String
Dim mstrTktNum As String
Dim mbolEOButton As Boolean
Dim sngGST As Single
Dim sngNettCostGST As Single
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date

Private Sub chkAbsorb_Click()
   EnableCalculate
   If chkAbsorb.value = 1 Then
      
   Else
      txtMerchFee = 0
   End If
End Sub

Private Sub chkGSTAbsorb_Click()
EnableCalculate
End Sub
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop)
Private Sub chkNettMerchant_Click()
Dim sngMFPct As Single
Dim sngCost As Single
Dim sngAmt As Single

sngCost = fConvertZero(txtCost)
sngMFPct = gobjPNR.CompInfo.MerchFeePct


If chkNettMerchant.value = 1 Then
    sngAmt = sngCost * (sngMFPct * 0.01)
    txtNettMF.Text = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP")
Else
    txtNettMF.Text = ""
End If

End Sub

Private Sub chkVenderMF_Click()

End Sub
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop)
Private Sub chkCWTHandlingMF_Click()
Dim sngMFPct As Single
Dim sngCost As Single
Dim sngVendorHandling As Single
Dim sngAmt As Single

sngCost = fConvertZero(txtCWTHandlingFee)
sngVendorHandling = fConvertZero(txtVendorHandling)
sngMFPct = gobjPNR.CompInfo.MerchFeePct


If chkCWTHandlingMF.value = 1 Then
    sngAmt = (sngCost + sngVendorHandling) * (sngMFPct * 0.01)
    txtCWTHandlingMF.Text = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP")
Else
    txtCWTHandlingMF.Text = ""
End If
End Sub

Private Sub cmbFOPType_Click()
Dim blnCC As Boolean
    
EnableCalculate
blnCC = (cmbFOPType = "CX")
cmbCCType.Visible = blnCC
txtCCNum.Visible = blnCC
dtpCCExp.Visible = blnCC
'chkWaiveMercFee.Visible = blnCC
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop)
If UCase(gstrAgcyCountryCode) = "HK" And frmOthSvcs.datProducts.Recordset![ProductCode] = "06" Then
    If (cmbFOPType = "INV") Then
        chkNettMerchant.Enabled = False
        chkNettMerchant.value = 0
    
        chkCWTHandlingMF.Enabled = False
        chkCWTHandlingMF.value = 0
       
    Else
        chkNettMerchant.Enabled = True
        chkCWTHandlingMF.Enabled = True
        
    End If
End If



End Sub

Private Sub cmdCalculate_Click()
'Dim sngSF As Single
Dim sngCost As Single
Dim sngGSTPct As Single
Dim sngSellingPrice As Single
Dim sngMF As Single
Dim sngMFPct As Single
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop)
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost): Variable declaration regarding to handling fee calculation logic for HK only
Dim sngVendorHandling As Single
Dim sngCWTHandling As Single
Dim sngMF_NettCost As Single
Dim sngMF_CWTHandling As Single
Dim sngGrossSales As Single
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - End of Declaration*************************


'If txtCommission.Text = "" Then txtCommission.Text = "0"
'If txtCost.Text = "" Then txtCost.Text = "0"

'sngSF = fConvertZero(txtCost.Text)
'If optCommType(1).Value Then
'    txtCommission.Text = fCommAmt(sngSF, fConvertZero(txtCommission.Text))
'    optCommType(0).Value = True
'End If
'sngSF = sngSF + fConvertZero(txtCommission.Text)
 
'If chkAbsorb.Value = 0 Then
'   If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.Value = vbUnchecked Then
'       Me.txtMerchFee = fMerchantFee(sngSF, fGetMerchFee(gstrAgcyCountryCode))
'   Else
'      Me.txtMerchFee.Text = "0"
'   End If
'Else
'   Me.txtMerchFee.Text = "0"
'End If

'sngSF = sngSF + fConvertZero(txtMerchFee.Text)

    'txtGST.Text = fGST(txtCost)
    
    'sngSF = sngSF + CSng(txtGST.Text)
    'If txtGST <> 0 Then
    '   sngSF = Format(sngSF, "0.00")
    'Else
    '   sngSF = Format(fCurrRound(sngSF, gstrAgcyCurrCode, "UP"), "0.00")
    'End If
'sngSF = Format(fCurrRound(sngSF, gstrAgcyCurrCode, "UP"), "0.00")
'txtSellPrice = sngSF
'txtGST.Text = fGST(txtSellPrice)

sngCost = fConvertZero(txtCost)

'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost):  handling fee calculation logic for HK only
If UCase(gstrAgcyCountryCode) = "HK" And frmOthSvcs.datProducts.Recordset![ProductCode] = "06" Then
    
    sngVendorHandling = fConvertZero(txtVendorHandling)
    sngCWTHandling = fConvertZero(txtCWTHandlingFee)
    sngMF_NettCost = fConvertZero(txtNettMF)
    sngMF_CWTHandling = fConvertZero(txtCWTHandlingMF)
    
    sngGrossSales = sngCost + sngVendorHandling + sngCWTHandling + sngMF_NettCost + sngMF_CWTHandling
    txtGrossSale = Format(sngGrossSales, gstrAgcyCurrFormat)
End If
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - End of handling fee calculation logic*************************

'Added on 6/7/2005: change of calculation logic

sngSellingPrice = fConvertZero(txtGrossSale) 'refer to gross sales
txtCost = Format(sngCost, gstrAgcyCurrFormat)
txtGrossSale = Format(sngSellingPrice, gstrAgcyCurrFormat)


'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) Product code 06 (Visa Cost):  No MF for FOP (HK only)
If UCase(gstrAgcyCountryCode) = "HK" And frmOthSvcs.datProducts.Recordset![ProductCode] = "06" Then
    txtCommission = Format(IIf(sngSellingPrice < sngCost, 0, sngSellingPrice - sngCost - sngVendorHandling), gstrAgcyCurrFormat)
    
Else 'for SG

    If chkGSTAbsorb Then
            sngGST = 0
            sngNettCostGST = 0
    Else
            sngGST = fGST(sngSellingPrice)
            sngNettCostGST = fGST(sngCost)
    End If
    txtGST = Format(sngGST, gstrAgcyCurrFormat)
    sngGSTPct = frmOthSvcs.datProducts.Recordset![GST] * 0.01
    sngMFPct = gobjPNR.CompInfo.MerchFeePct

    If chkAbsorb.value = 0 Then
        If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.value = vbUnchecked Then
            'similar to sngSellingPrice
            sngMF = fMerchantFee(sngSellingPrice * (1 + sngGSTPct), sngMFPct)
        Else
            sngMF = "0"
        End If
    Else
        sngMF = "0"
    End If
    sngSellingPrice = (sngSellingPrice + sngGST + sngMF) / (1 + sngGSTPct)
    sngGST = fGST(sngSellingPrice)
    txtMerchFee = Format(sngMF, gstrAgcyCurrFormat)
    txtCommission = Format(IIf(sngSellingPrice < sngCost, 0, sngSellingPrice - sngCost), gstrAgcyCurrFormat)
End If


''ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost):  handling fee calculation logic for HK only
'If UCase(gstrAgcyCountryCode) = "HK" And frmOthSvcs.datProducts.Recordset![ProductCode] = "06" Then
'    txtCommission = Format(IIf(sngSellingPrice < sngCost, 0, sngSellingPrice - sngCost - sngVendorHandling), gstrAgcyCurrFormat)
'Else
'    txtCommission = Format(IIf(sngSellingPrice < sngCost, 0, sngSellingPrice - sngCost), gstrAgcyCurrFormat)
'End If
''ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - End of handling fee calculation logic*************************
txtSellPrice = Format(sngSellingPrice, gstrAgcyCurrFormat)
cmdCalculate.Enabled = False

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
Dim strMsg As String

datTouchEnd = Now

'modified on 24/02/05: if no EO num, check for Visa number is valid
If txtEONum = "" And gstrAgcyCountryCode <> "HK" Then
    If Not validData(True) Then Exit Sub
Else
    If Not validData Then Exit Sub
End If

'230108
If isRequireClientMI(gobjPNR.CN, 6) And frmClientMI.MSXfreefields = "" Then
        'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
        'If frmClientMI.bolCheck = False Then
        cmdDone.Enabled = True
        'MsgBox "Client MI data is incomplete", vbCritical
        strMsg = "Client MI data is incomplete"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        loadClientMI
        Exit Sub
        'End If
End If
Dim freefields As String

If frmClientMI.MSXfreefields <> "" Then
        freefields = freefields & "/" & frmClientMI.MSXfreefields
End If

cmdDone.Enabled = False
gSysStartOthSvcsTime = Now

If gobjEO Is Nothing Or txtEONum = "" Then Call SetEOObj(gobjEO)

'Added on 24/2/05: Capture visa number if no EO number
If gobjEO.TicketNumber = "0000" And gstrAgcyCountryCode <> "HK" Then Call SetVisaNo

    'Timer
    'frmWait.Show
    Me.MousePointer = 11
    Call modOthSvcs.WriteOSToGDS(gobjEO, frmOthSvcs.datProducts.Recordset![Type], gStartOthSvcsTime, freefields)
    'Remarks by JY ANG. Eliminate Queue Screen for non-air products
    'Call pTktQueue
    Unload frmClientMI
    Me.MousePointer = 0
    Log
    Unload Me
    'Unload frmWait
    Set gobjEO = Nothing
End Sub
Private Sub SetVisaNo()
Dim strEOtemp As String
Dim strZero As String
Dim lengthEO As Integer
Dim lngI As Long
    With gobjEO
              .EONumber = Trim(txtVisaNum.Text)
       
              lengthEO = Len(.EONumber)
              'tktlen
              For lngI = 1 To 7 - lengthEO
                  strZero = strZero & "0"
              Next lngI
              strEOtemp = strZero & .EONumber
              .TicketNumber = frmOthSvcs.datProducts.Recordset![TktPrefix] & strEOtemp
              .EONumber = .TicketNumber
              
     End With
 
End Sub

Private Sub cmdEO_Click()
Dim lngC As Long
Dim strMsg As String
datTouchEnd = Now
mbolEOButton = True


If Not validData Then Exit Sub

'230108
If isRequireClientMI(gobjPNR.CN, 6) And frmClientMI.MSXfreefields = "" Then
        'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
        'If frmClientMI.bolCheck = False Then
        cmdDone.Enabled = True
        'MsgBox "Client MI data is incomplete", vbCritical
        strMsg = "Client MI data is incomplete"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        loadClientMI
        Exit Sub
        'End If
End If

gSysStartOthSvcsTime = Now
'If Not ValidData Then Exit Sub
mbolEOButton = False
Call SetEOObj(gobjEO)

'Added on 8/3/2005: To end PNR to get RecLoc
'Call modOthSvcs.SetEONumber

Load frmExchangeOrder
frmExchangeOrder.Show '1, Me
        Do
            DoEvents
        Loop Until isLoaded("frmExchangeOrder") = False
If gbolIgnoreEO Then Exit Sub
'Modified on 24/2/05: capture EO Number when click on EO, if not use visa num enter
'added on 10/11/04 to capture Visa Number for SG
'If gstrAgcyCountryCode = "HK" Then
txtEONum.Text = gobjEO.EONumber
'Else
'txtEONum.Text = txtVisaNum.Text
'End If
txtEONum.Locked = True
Unload frmExchangeOrder
Set frmExchangeOrder = Nothing

'If gobjEO.TicketNumber = "0000" Then
'    gobjEO.TicketNumber = frmOthSvcs.datProducts.Recordset![TktPrefix] & gobjEO.TicketNumber & txtEONum
'End If

cmdEO.Enabled = False

'Added on 8/3/2005: To end PNR to get RecLoc
cmdDone.Enabled = False
'Remarks by JY ANG. Eliminate Queue Screen for non-air products
'Call pTktQueue
Set gobjEO = Nothing
Log
Unload Me

'Call cmdDone_Click
End Sub


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
            'lstEORmks(0).RemoveItem 0
            'lngC = lngC - 1
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
    '.RemoveItem .ListIndex
    End If
End If
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

'Private Sub dbcVendors_Click(Area As Integer)
'   frmOthSvcs.dbcVendors.Text = dbcVendors.Text
'   frmOthSvcs.datSelectedVendor.DatabaseName = gstrTProDBSource
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
         'fraVendorInfo.Enabled = False
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
gStartOthSvcsTime = Now
frmOSVisa.Caption = "CWT TravelPro - " & frmOthSvcs.dbcProducts.Text
'FormCenter
SSTab1.Tab = 0
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
        
        Me.cmbFOPType = "CX"
        
        If blnChkVendor = False Or blnChkDate = False Then
           promptCCError blnChkVendor, blnChkDate
        End If
    Else
        Me.cmbFOPType = "INV"
    End If
    
 End With
  
cmbVisaType.listindex = 0
cmbValidityPeriod.listindex = 0
cmbProcess.listindex = 0
cmbEntries.listindex = 0

dbcVendors.Visible = False


 If gbolEOAmend Then
    RetrieveData
 Else
    GetVendorInfo
 End If
 
 'Added on 24 Jan: Vender credit term for EO Transaction
If frmOthSvcs.datSelectedVendor.Recordset![CreditTerms] <> "" Then txtCreditTerms = frmOthSvcs.datSelectedVendor.Recordset![CreditTerms]
txtCreditTerms.Enabled = False
txtCreditTerms.Locked = True

'added on 14/06/2006 : for SG, auto enable/disable EO button


If UCase(gstrAgcyCountryCode) = "SG" Then


    If frmOthSvcs.datProducts.Recordset![GST] > 0 Then
        lblGST.Visible = True
        txtGST.Visible = True
        chkGSTAbsorb.Visible = True
    End If
    If IsNull(frmOthSvcs.datSelectedVendor.Recordset!RaiseType) Then
        cmdEO.Enabled = False
        cmdDone.Enabled = True
    Else
        cmdEO.Enabled = True
        cmdDone.Enabled = False
        Frame4.Visible = False
        txtVisaNum.Visible = False
    End If
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost): enable handling fee for HK, disable handling fee for SG
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost): enable Processing Days text box for HK and SG

    If frmOthSvcs.datProducts.Recordset![ProductCode] = "06" Then
        Label_Days.Visible = True
        txtDay.Visible = True
        SSTab1.TabVisible(3) = False
        
    Else
        Label_Days.Visible = False
        txtDay.Visible = False
    End If
    
    'Nett Cost merchant fees
    chkNettMerchant.Visible = False
    txtNettMF.Visible = False
    'visa handling fees
    Label_CWTHandlingFee.Visible = False
    txtCWTHandlingFee.Visible = False
    chkCWTHandlingMF.Visible = False
    txtCWTHandlingMF.Visible = False
    'Vendor Handling fees
    Label_VenderHandlingFee.Visible = False
    txtVendorHandling.Visible = False
    'Merchant Fees label
    Label_HKMerchantFee_1.Visible = False
    
    
'country HK
Else
    Frame4.Visible = False
    txtVisaNum.Visible = False
    lblGST.Visible = False
    txtGST.Visible = False
    chkGSTAbsorb.Visible = False
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost): enable handling fee for HK, disable handling fee for SG
'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost): enable Processing Days text box for HK and SG
    If frmOthSvcs.datProducts.Recordset![ProductCode] = "06" Then
        
        Label_Days.Visible = True
        txtDay.Visible = True
        SSTab1.TabVisible(3) = False
        txtMerchFee.Visible = False
        chkAbsorb.Visible = False
        Label2.Visible = False
        
        'Nett Cost merchant fees
        chkNettMerchant.Visible = True
        txtNettMF.Visible = True
        'CWT handling fees
        Label_CWTHandlingFee.Visible = True
        txtCWTHandlingFee.Visible = True
        chkCWTHandlingMF.Visible = True
        txtCWTHandlingMF.Visible = True
        'Vendor Handling fees
        Label_VenderHandlingFee.Visible = True
        txtVendorHandling.Visible = True
        'Merchant Fees label
        Label_HKMerchantFee_1.Visible = True
        
        ' Merchant Fees chkbox value
        If (cmbFOPType = "INV") Then
            chkNettMerchant.Enabled = False
            chkNettMerchant.value = 0
    
            chkCWTHandlingMF.Enabled = False
            chkCWTHandlingMF.value = 0
       
        Else
            chkNettMerchant.Enabled = True
            chkCWTHandlingMF.Enabled = True
        End If
                
    Else
        Label_Days.Visible = False
        txtDay.Visible = False
       'Nett Cost merchant fees
        chkNettMerchant.Visible = False
        txtNettMF.Visible = False
        'CWT handling fees
        Label_CWTHandlingFee.Visible = False
        txtCWTHandlingFee.Visible = False
        chkCWTHandlingMF.Visible = False
        txtCWTHandlingMF.Visible = False
        'Vendor Handling fees
        Label_VenderHandlingFee.Visible = False
        txtVendorHandling.Visible = False
        'Merchant Fees label
        Label_HKMerchantFee_1.Visible = False
        
        
    End If
    
End If
txtReplyEmail.Text = GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjPNR.Agent, gobjPNR.PCCOwner, False, True)

If OSNoMF(frmOthSvcs.datProducts.Recordset![ProductCode], frmOthSvcs.datSelectedVendor.Recordset!VendorNumber) = True Then
    chkAbsorb.value = 1
    chkAbsorb.Enabled = False
Else
    chkAbsorb.value = 0
    chkAbsorb.Enabled = True
End If

'Preethi-V1.2.6 20110906 - CR70 - Grey off Commission Box For 14 Product Codes
If frmOthSvcs.datProducts.Recordset![FullComm] <> "" And frmOthSvcs.datProducts.Recordset![FullComm] = True Then
   txtCost.Text = "0"
   txtCost.Enabled = False
   txtCommission.Enabled = False
End If
'230108
'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
'frmClientMI.bolCheck = False
datFormLoadEnd = Now
If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
End Sub

Private Sub RetrieveData()
   Dim strSQL As String
   Dim rsEO As ADODB.Recordset
   Dim i As Integer
   Dim j As Integer
   Dim strFOP() As String
   Dim strRem() As String
   Dim strTemp() As String
      
   cmdDone.Enabled = False
   dbcVendors.Visible = True
   mstrPCAmend = frmEOAmend.lsvEO.SelectedItem.SubItems(1)
   'datVendors.DatabaseName = gstrTProDBSource
    
   'datVendors.RecordSource = "SELECT * FROM tblVendors WHERE [ProductCodes] LIKE '*" & mstrPCAmend & "*' ORDER BY [VendorName]"
   'datVendors.Refresh

   'With dbcVendors
   '   .Text = ""
   '   .ListField = "VendorName"
      '.BoundColumn = "SortKey"
   '   .BoundColumn = "VendorNumber"
   '   .Refresh
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
     mstrTktNum = !TktNum & ""
     txtEONum = !ExchangeID
     txtContact = !ContactPerson
     txtCost = !Cost
     txtCommission = !Commission - !MerchantFee
     'optCommType(0).Value = True
     txtGST.Text = !Tax1
     txtMerchFee = !MerchantFee
     chkAbsorb.value = IIf(!CWTAbsorb = True, 1, 0)
     'txtSellPrice = !SellPrice
     strFOP = Split(!FOP, "/")
     If strFOP(0) = "INV" Then
        cmbFOPType.Text = strFOP(0)
        If UBound(strFOP) = 2 Then
           'chkNRCC.Value = strFOP(1)
           chkWaiveMercFee.value = strFOP(2)
        End If
     Else
        cmbFOPType.Text = strFOP(0)
        cmbCCType.Text = strFOP(1)
        txtCCNum.Text = strFOP(2)
        dtpCCExp.value = "1/" & MMM(Left(strFOP(3), 2)) & "/" & Right(strFOP(3), 2)
        If UBound(strFOP) = 5 Then
           'chkNRCC.Value = strFOP(4)
           chkWaiveMercFee.value = strFOP(5)
        End If
     End If
     
     lstEORmks(1).Clear
     strRem = Split(!ListBoxRemark, vbCrLf)
     For i = 0 To UBound(strRem)
        lstEORmks(1).AddItem strRem(i)
     Next
     
     strTemp = Split(!VisaInfo, vbCrLf)
     If UBound(strTemp) = 6 Then
        txtCountry = strTemp(0)
        For i = 0 To cmbVisaType.ListCount - 1
           If cmbVisaType.List(i) = strTemp(1) Then
              cmbVisaType.listindex = i
              Exit For
           End If
        Next
        For i = 0 To cmbEntries.ListCount - 1
           If cmbEntries.List(i) = strTemp(2) Then
              cmbEntries.listindex = i
              Exit For
           End If
        Next
        txtValidity = strTemp(3)
        For i = 0 To cmbValidityPeriod.ListCount - 1
           If cmbValidityPeriod.List(i) = strTemp(4) Then
              cmbValidityPeriod.listindex = i
              Exit For
           End If
        Next
        For i = 0 To cmbProcess.ListCount - 1
           If cmbProcess.List(i) = strTemp(5) Then
              cmbProcess.listindex = i
              Exit For
           End If
        Next
     End If
     txtVisaNum = strTemp(6)
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
  End With
  rsEO.Close
  SetEOObj gobjPreEO
  Unload frmEOAmend
End Sub

Private Function fCommAmt(Cost As Single, CommPct As Single) As Single
gobjLog.ProcedureName = "fCommAmt"
'On Error GoTo ProcError

Dim sngPct As Single
Dim sngAmt As Single
sngPct = CommPct * 0.01

        'sngAmt = (Cost / (1 - sngPct)) - Cost
        sngAmt = Cost * sngPct
        fCommAmt = fCurrRound(sngAmt, gstrAgcyCurrCode, "DO")

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

    'If gstrAgcyCountryCode = "HK" Then
        sngAmt = CDec(TotalCharge * sngPct)
    'Else
    '    sngAmt = (TotalCharge / (1 - sngPct)) - TotalCharge
    'End If
    fMerchantFee = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP")

Exit Function
ProcError:
    Call pErrorReport(True)

End Function

Private Function validData(Optional ValidateVISA As Boolean = False) As Boolean
Dim strMsg As String
Dim i As Integer
'Dim rs As adodb.Recordset

If gbolEOAmend = False Then
   If frmOthSvcs.dbcVendors.Text = "" Then strMsg = strMsg & "Select vendor for this transaction..." & Chr(13)
End If
If txtCountry = "" Then strMsg = strMsg & "Need country name..." & Chr(13)
If txtSellPrice.Text = "" Or txtSellPrice.Text = "0" Then strMsg = strMsg & "Need to calculate Selling Price..." & Chr(13)
If cmbFOPType.Text = "" Then strMsg = strMsg & "Need form of payment..." & Chr(13)

'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost): Processing Days text box - Only allow two Numeric Characters
If txtDay.Visible = True Then
    If txtDay.Text <> "" Then
       If Len(txtDay.Text) > 2 Then
            strMsg = strMsg & "Processing days must not exceed two numeric character ..." & Chr(13)
       End If
    End If
End If



'If gstrAgcyCountryCode <> "HK" And ValidateVISA = True Then
    'If Len(txtVisaNum) <> 8 Then
    '   strMsg = strMsg & "Invalid Visa Number ..." & Chr(13)
    'ElseIf UCase(Left(txtVisaNum, 1)) <> "V" Or IsNumeric(Mid(txtVisaNum, 2, 7)) = False Then
    '   strMsg = strMsg & "Invalid Visa Number ..." & Chr(13)
    'Else
    '   If mbolEOButton And gbolEOAmend = False Then
    '      Set rs = gdbConn.Execute("select tblExchangeOrder where ExchangeID='" & txtVisaNum & "'")
          'rs.Index = "ExchangeID"
          'rs.Seek "=", txtVisaNum

    '      If rs.EOF = False Then
    '         strMsg = strMsg & "Visa number exist..."
    '      End If
    '      rs.Close
    '      Set rs = Nothing
    'End If
    'End If
'End If

If cmbFOPType.Text = "CX" Then
    If cmbCCType.Text = "" Then strMsg = strMsg & "Need valid credit vendor code..." & Chr(13)
    If txtCCNum = "" Then strMsg = strMsg & "Need valid credit card number..." & Chr(13)
    If LastDate(dtpCCExp.value) < Date Then strMsg = strMsg & "Need valid expiration date..." & Chr(13)
    If (txtCCNum.Text <> "" And cmbCCType.Text <> "") Then If ValidCCNum(cmbCCType.Text, txtCCNum.Text) = False Then strMsg = strMsg & "Credit card number is invalid or wrong card vendor selected ..." & Chr(13)
End If
If UCase(gstrAgcyCountryCode) = "SG" Then
    'Add on 200106: TMP card must absorb MF
    'If (cmbFOPType <> "INV" And Left(UCase(cmbCCType.Text), 2) = "DC" And _
        Left(UCase(txtCCNum.Text), 7) = "3644033") And chkAbsorb.value <> 1 Then
     'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    If (cmbFOPType <> "INV" And IsTMPCard(Left(UCase(cmbCCType.Text), 2), UCase(txtCCNum.Text))) And _
        chkAbsorb.value <> 1 Then
        strMsg = strMsg & "Need to tick aborbed merchant fee for TMP card and recalculate selling price" & Chr(13)
    End If
End If

If InStr(1, UCase(txtCountry), "/PC") <> 0 Then
   strMsg = strMsg & "'/PC' exist in Country text box.." & Chr(13)
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

If strMsg = "" Then
    validData = True
Else
    'MsgBox strMsg, vbApplicationModal + vbExclamation, "Travel Pro"
     modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    'cmdDone.Enabled = True
    'If gbolEOAmend = False Then cmdDone.Enabled = True
End If

End Function


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






Private Sub Form_Unload(Cancel As Integer)
'Preethi - V1.1.1 20100831 - IR2 - Client MI screen is populated with old data
Unload frmClientMI
End Sub

Private Sub txtCCNum_GotFocus()
Call pSetSelected
End Sub

Private Sub txtCCNum_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii)
End Sub

Private Sub txtCommission_GotFocus()
Call pSetSelected
End Sub

Private Sub txtCommission_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtCost_Change()
EnableCalculate
End Sub

Private Sub txtCost_GotFocus()
Call pSetSelected
End Sub

Private Sub EnableCalculate()
   cmdCalculate.Enabled = True
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtCountry_GotFocus()
Call pSetSelected

End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
'Preethi - V1.2.14 20120914 -  CR187  EO Visa - Allow alphanumeric and limit characters in Country box
    KeyAscii = fAllowAlphaNumeric(KeyAscii, "-. ")
End Sub

Private Sub txtFreeRmk_GotFocus()
Call pSetSelected
End Sub

Private Sub txtFreeRmk_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowAlphaNumeric(KeyAscii, "#$*()/.: ?@")
End Sub



Private Sub txtGrossSale_Change()
EnableCalculate
End Sub

Private Sub txtCWTHandlingFee_Change()
EnableCalculate
End Sub

Private Sub txtMerchFee_GotFocus()
Call pSetSelected
End Sub

Private Sub txtMerchFee_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtNettMF_Change()
EnableCalculate
End Sub

Private Sub txtSellPrice_GotFocus()
Call pSetSelected
End Sub

Private Sub txtSellPrice_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtValidity_GotFocus()
Call pSetSelected

End Sub

Private Sub txtValidity_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii)

End Sub

Private Sub SetEOObj(ByRef objEO As EO)
Dim lngC As Long
Dim lengthEO As Integer
Dim strZero As String
Dim lngI As Long
Dim strEOtemp As String
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
    
    'Remove on 24/2/05: Capture visa number if no EO number
    
    'If UCase(gstrAgcyCountryCode) <> "HK" Then
    '    .EONumber = Trim(txtVisaNum.Text)
    'End If
    .BillingDescription = ""
    .CN = gobjPNR.CN
    .CommissionAmt = fConvertZero(txtCommission.Text) '+ fConvertZero(txtMerchFee)
    .Cost = fConvertZero(txtCost.Text)
    .CreatedBy = gobjHost.AgentSine
    '.CreatedByName = gobjHost.AgentName
    .CreatedByName = gobjHost.AgentProfile
    .CreatedByPCC = gobjHost.AgentPCC
    .CreateDtTm = Now()
    .VisaCountry = txtCountry
    .VisaEntries = cmbEntries.Text & ""
    .VisaProcess = cmbProcess.Text & ""
    .VisaType = cmbVisaType.Text & ""
    .DescriptionLineAdd frmOthSvcs.dbcProducts.Text
    .DescriptionLineAdd "FOR " & txtCountry.Text
    .DescriptionLineAdd UCase(cmbVisaType.Text) & " VISA - " & cmbEntries.Text & " ENTRIES - VALID " & txtValidity.Text & " " & cmbValidityPeriod.Text
    
    '.FOP = cmbFOPType.Text & IIf(cmbFOPType.Text = "CX", "/" & cmbCCType.Text & "/" & txtCCNum.Text & "/" & Format(dtpCCExp.Value, "ddmm"), "")
    .FOP = cmbFOPType.Text & IIf(cmbFOPType.Text = "CX", "/" & cmbCCType.Text & "/" & txtCCNum.Text & "/" & Format(dtpCCExp.value, "MMYY"), "") & "/" & "0" & "/" & chkWaiveMercFee.value
    
    '.PaxName = gobjPNR.PassengerName(1).LastName & "/" & gobjPNR.PassengerName(1).FirstName
    For lngC = 1 To gobjPNR.PassengerCount
     strPaxName = strPaxName & gobjPNR.PassengerName(lngC).LastName & "/" & gobjPNR.PassengerName(lngC).FirstName & IIf(lngC = gobjPNR.PassengerCount, "", vbCrLf)
    Next
    .PaxName = strPaxName
    
    .PNRRecLoc = gobjPNR.RecLoc
    '.ProductCode = frmOthSvcs.datProducts.Recordset![SortKey]
    .ProductCode = frmOthSvcs.datProducts.Recordset![ProductCode]
    .ProductSortKey = frmOthSvcs.datProducts.Recordset![SortKey]
    .SellPrice = fConvertZero(txtSellPrice.Text)
    '.ServiceDate = DateAdd("d", Date, 90)
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
    '.TaxAdd CSng(txtGST.Text), IIf(txtGST.Text = "0", "", "GST")
    '.TaxAdd 0, ""
    '.TaxAdd CSng(txtGST.Text), "GST"
    'If UCase(gstrAgcyCountryCode) = "SG" Then
    '    .TaxAdd fConvertZero(txtGST.Text), "GST"
    'Else
    '    .TaxAdd fConvertZero(txtGST.Text), ""
    'End If
    If UCase(gstrAgcyCountryCode) = "SG" Then
        .TaxAdd sngGST, "GST"
           .NettGST = sngNettCostGST
    Else
        .TaxAdd 0, ""
        .NettGST = 0
        'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost): Add Vendor Handling fees to HK-EO
        '.VendorHandling (fConvertZero(txtVendorHandling))
        .VendorHandling = (fConvertZero(txtVendorHandling))
    End If
    If gbolEOAmend Then
       .TicketNumber = mstrTktNum
    Else
       .TicketNumber = "0000"
    End If
    '.VendorCode = frmOthSvcs.dbcVendors.BoundText
    'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop) - Product code 06 (Visa Cost): input Processing Days to remarks
    '.RemarkAdd "PROCESSING - " & cmbProcess.Text
    If txtDay.Text <> "" And Trim(txtDay.Text) <> "0" Then
        .RemarkAdd "PROCESSING - " & cmbProcess.Text & " - " & txtDay.Text & "day(s)"
    Else
        .RemarkAdd "PROCESSING - " & cmbProcess.Text
    End If
    '.Email = frmOthSvcs.datVendors.Recordset![Email] & ""
    '.FaxNo = frmOthSvcs.datVendors.Recordset![FaxNumber] & ""
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
    
    For lngC = 0 To lstEORmks(1).ListCount - 1
        .RemarkAdd lstEORmks(1).List(lngC)
    Next
    For lngC = 0 To lstItinRmks(1).ListCount - 1
        .RIRemarkAdd lstItinRmks(1).List(lngC)
    Next
    
    .MerchFee = fConvertZero(txtMerchFee)
    .CWTAbsorb = IIf(chkAbsorb.value = 1, True, False)
    .VisaInfo = VisaInfo
    .ListBoxRem = ListBoxRemark
    .PassengerID = IIf(txtPassengerID = "", "1", txtPassengerID)
    .ReplyEmail = Trim(UCase(txtReplyEmail.Text))
End With

End Sub

Private Function VisaInfo() As String
   VisaInfo = txtCountry & vbCrLf
   VisaInfo = VisaInfo & cmbVisaType.Text & vbCrLf
   VisaInfo = VisaInfo & cmbEntries.Text & vbCrLf
   VisaInfo = VisaInfo & txtValidity.Text & vbCrLf
   VisaInfo = VisaInfo & cmbValidityPeriod.Text & vbCrLf
   VisaInfo = VisaInfo & cmbProcess.Text & vbCrLf
   VisaInfo = VisaInfo & txtVisaNum.Text
   
End Function

Private Function ListBoxRemark() As String
   Dim i As Integer
   
   ListBoxRemark = ""
   For i = 0 To lstEORmks(1).ListCount - 1
      ListBoxRemark = ListBoxRemark & IIf(ListBoxRemark <> "", vbCrLf, "") & lstEORmks(1).List(i)
   Next
End Function

Private Function fGST(TotalCharge As Single) As Single
Dim sngAmt As Single
Dim sngPct As Single

sngPct = frmOthSvcs.datProducts.Recordset![GST] * 0.01
sngAmt = sngPct * TotalCharge
'fGST = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP")
fGST = Format(sngAmt, "0.00")

End Function

Sub FormCenter()
    Top = (Screen.Height * 0.95) / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
End Sub

Private Sub txtVenderMF_Change()

End Sub

Private Sub txtVendorHandling_Change()
EnableCalculate
End Sub

Private Sub txtCWTHandlingMF_Change()
EnableCalculate
End Sub

Private Sub txtVisaNum_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii)
End Sub
Private Sub cmdClientMI_Click()
    Call loadClientMI
End Sub
Private Sub loadClientMI()
    If isLoaded("frmClientMI") Then
        frmClientMI.Show 'vbModal
    Else
        Load frmClientMI
        frmClientMI.intLocation = 6
        frmClientMI.intProdCode = frmOthSvcs.dbcProducts.BoundText
        frmClientMI.cmbMICat.Enabled = False
        frmClientMI.pGetClientMI (gobjPNR.CN)
        '230108
        frmClientMI.strPdtType = frmOthSvcs.datProducts.Recordset![Type]
        frmClientMI.Show 'vbModal
    End If
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



