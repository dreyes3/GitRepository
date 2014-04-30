VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOSCarTxfr 
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "CWT Travel Pro - Car Transfer"
   ScaleHeight     =   8280
   ScaleWidth      =   11160
   Begin TabDlg.SSTab SSTab1 
      Height          =   7635
      Left            =   120
      TabIndex        =   39
      Top             =   600
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   13467
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
      TabCaption(0)   =   "Service Info"
      TabPicture(0)   =   "frmOSCarTxfr.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Picture1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtContact"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Remarks"
      TabPicture(1)   =   "frmOSCarTxfr.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "cmdFreeRmkToItin"
      Tab(1).Control(4)=   "cmdFreeRmkToEO"
      Tab(1).Control(5)=   "txtFreeRmk"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Vendor Info"
      TabPicture(2)   =   "frmOSCarTxfr.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraVendorInfo"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "MI"
      TabPicture(3)   =   "frmOSCarTxfr.frx":0054
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
         Left            =   -74880
         TabIndex        =   112
         Top             =   600
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Height          =   2835
         Left            =   120
         TabIndex        =   90
         Top             =   1515
         Width           =   4755
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
            TabIndex        =   99
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtGST 
            Height          =   315
            Left            =   1980
            TabIndex        =   98
            Top             =   1620
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
            TabIndex        =   97
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtMerchFee 
            Height          =   315
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   1980
            Width           =   1215
         End
         Begin VB.TextBox txtCommission 
            Height          =   315
            Left            =   1980
            TabIndex        =   95
            Top             =   1275
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            Height          =   315
            Left            =   1980
            TabIndex        =   94
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtSellPrice 
            Height          =   315
            Left            =   1980
            TabIndex        =   93
            Top             =   2340
            Width           =   1215
         End
         Begin VB.TextBox txtGrossSale 
            Height          =   315
            Left            =   1980
            TabIndex        =   92
            Top             =   600
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
            TabIndex        =   91
            Top             =   1560
            Visible         =   0   'False
            Width           =   1455
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
            TabIndex        =   105
            Top             =   1620
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
            TabIndex        =   104
            Top             =   1980
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
            TabIndex        =   103
            Top             =   1320
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
            TabIndex        =   102
            Top             =   240
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
            TabIndex        =   101
            Top             =   2340
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
            TabIndex        =   100
            Top             =   600
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
         Height          =   675
         Left            =   6120
         TabIndex        =   87
         Top             =   5400
         Width           =   4395
         Begin VB.TextBox txtPassengerID 
            Height          =   315
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   88
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
            TabIndex        =   89
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.Frame fraVendorInfo 
         Height          =   5895
         Left            =   -74760
         TabIndex        =   56
         Top             =   1035
         Width           =   9975
         Begin VB.TextBox txtReplyEmail 
            Height          =   375
            Left            =   240
            TabIndex        =   110
            Top             =   3960
            Width           =   8655
         End
         Begin VB.TextBox txtTel 
            Height          =   375
            Left            =   960
            TabIndex        =   73
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox txtCreditTerms 
            Height          =   420
            Left            =   4200
            TabIndex        =   70
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtVendor 
            Height          =   375
            Left            =   960
            TabIndex        =   63
            Top             =   360
            Width           =   7935
         End
         Begin VB.TextBox txtAddress1 
            Height          =   375
            Left            =   960
            TabIndex        =   62
            Top             =   840
            Width           =   7935
         End
         Begin VB.TextBox txtAddress2 
            Height          =   375
            Left            =   960
            TabIndex        =   61
            Top             =   1320
            Width           =   7935
         End
         Begin VB.TextBox txtCity 
            Height          =   375
            Left            =   960
            TabIndex        =   60
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtCountry1 
            Height          =   375
            Left            =   3840
            TabIndex        =   59
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   960
            TabIndex        =   58
            Top             =   2280
            Width           =   7935
         End
         Begin VB.TextBox txtFaxNo 
            Height          =   375
            Left            =   960
            TabIndex        =   57
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label9 
            Caption         =   "Reply Email in EO (Only 1 email address is allowed)"
            Height          =   375
            Left            =   240
            TabIndex        =   111
            Top             =   3720
            Width           =   4095
         End
         Begin VB.Label Label11 
            Caption         =   "Contact No."
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Credit Terms"
            Height          =   255
            Left            =   3000
            TabIndex        =   71
            Top             =   2760
            Width           =   1065
         End
         Begin VB.Label Label15 
            Caption         =   "Vendor "
            Height          =   375
            Left            =   240
            TabIndex        =   69
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Address"
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "City"
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Country"
            Height          =   375
            Left            =   3120
            TabIndex        =   66
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Email (;)"
            Height          =   375
            Left            =   240
            TabIndex        =   65
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Fax No (,)"
            Height          =   375
            Left            =   240
            TabIndex        =   64
            Top             =   2760
            Width           =   735
         End
      End
      Begin VB.TextBox txtContact 
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   55
         Top             =   990
         Width           =   2655
      End
      Begin VB.PictureBox Picture1 
         Height          =   1395
         Left            =   1680
         Picture         =   "frmOSCarTxfr.frx":0070
         ScaleHeight     =   1335
         ScaleWidth      =   1395
         TabIndex        =   53
         Top             =   4470
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Height          =   1455
         Left            =   6120
         TabIndex        =   43
         Top             =   6000
         Width           =   4395
         Begin VB.TextBox txtEONum 
            Height          =   315
            Left            =   2100
            TabIndex        =   22
            Top             =   360
            Width           =   2055
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
            Picture         =   "frmOSCarTxfr.frx":63B2
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
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
            Left            =   2160
            TabIndex        =   24
            Top             =   840
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
            TabIndex        =   23
            Top             =   840
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
         Left            =   180
         TabIndex        =   42
         Top             =   6075
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
            ItemData        =   "frmOSCarTxfr.frx":67F4
            Left            =   240
            List            =   "frmOSCarTxfr.frx":67FE
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
            ItemData        =   "frmOSCarTxfr.frx":680B
            Left            =   1800
            List            =   "frmOSCarTxfr.frx":6827
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
            Format          =   61931523
            CurrentDate     =   36526
            MaxDate         =   73050
            MinDate         =   36526
         End
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
         TabIndex        =   31
         Tag             =   "NN"
         Top             =   3810
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
         Picture         =   "frmOSCarTxfr.frx":684B
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Add Free Text to Exchange Order Remarks"
         Top             =   3690
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
         Picture         =   "frmOSCarTxfr.frx":6C8D
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Add Free Text to Itinerary Remarks"
         Top             =   3690
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
         TabIndex        =   41
         Top             =   4290
         Width           =   10275
         Begin VB.ListBox lstItinRmks 
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   34
            Top             =   300
            Width           =   4755
         End
         Begin VB.ListBox lstItinRmks 
            Height          =   2205
            Index           =   1
            Left            =   5400
            MultiSelect     =   1  'Simple
            TabIndex        =   38
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
            Picture         =   "frmOSCarTxfr.frx":70CF
            Style           =   1  'Graphical
            TabIndex        =   36
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
            Picture         =   "frmOSCarTxfr.frx":7511
            Style           =   1  'Graphical
            TabIndex        =   35
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
            Picture         =   "frmOSCarTxfr.frx":7953
            Style           =   1  'Graphical
            TabIndex        =   37
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
         TabIndex        =   40
         Top             =   1050
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
            Picture         =   "frmOSCarTxfr.frx":7D95
            Style           =   1  'Graphical
            TabIndex        =   28
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
            Picture         =   "frmOSCarTxfr.frx":81D7
            Style           =   1  'Graphical
            TabIndex        =   26
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
            Picture         =   "frmOSCarTxfr.frx":8619
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Add All Remarks"
            Top             =   1080
            Width           =   495
         End
         Begin VB.ListBox lstEORmks 
            Height          =   2205
            Index           =   1
            Left            =   5400
            MultiSelect     =   1  'Simple
            TabIndex        =   29
            Top             =   300
            Width           =   4755
         End
         Begin VB.ListBox lstEORmks 
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   25
            Top             =   300
            Width           =   4755
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4935
         Left            =   4920
         TabIndex        =   44
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   8705
         _Version        =   393216
         TabOrientation  =   3
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Departure"
         TabPicture(0)   =   "frmOSCarTxfr.frx":8A5B
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(1)=   "Label7"
         Tab(0).Control(2)=   "cmbFlights(0)"
         Tab(0).Control(3)=   "dtpPUDateTime"
         Tab(0).Control(4)=   "Frame4"
         Tab(0).Control(5)=   "Frame7"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Arrival"
         TabPicture(1)   =   "frmOSCarTxfr.frx":8A77
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label6"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "dtpRtnDateTime"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmbFlights(1)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame8"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Frame9"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).ControlCount=   6
         Begin VB.Frame Frame9 
            Caption         =   "To:"
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
            Left            =   120
            TabIndex        =   52
            Top             =   2160
            Width           =   4815
            Begin VB.ComboBox cmbLocation 
               Height          =   315
               Index           =   3
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   85
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtLocation 
               Height          =   285
               Index           =   3
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   84
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtRtToAddress 
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
               Index           =   2
               Left            =   120
               TabIndex        =   18
               Tag             =   "NN"
               Top             =   1320
               Width           =   4485
            End
            Begin VB.TextBox txtRtToAddress 
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
               Index           =   1
               Left            =   120
               TabIndex        =   17
               Tag             =   "NN"
               Top             =   960
               Width           =   4485
            End
            Begin VB.TextBox txtRtToAddress 
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
               Index           =   0
               Left            =   120
               TabIndex        =   16
               Tag             =   "NN"
               Top             =   600
               Width           =   4485
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Location:"
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
               Index           =   3
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   945
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "From:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   4815
            Begin VB.TextBox txtClientTel 
               Height          =   285
               Index           =   1
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   107
               Top             =   1680
               Width           =   1575
            End
            Begin VB.ComboBox cmbLocation 
               Height          =   315
               Index           =   2
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   82
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtLocation 
               Height          =   285
               Index           =   2
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   81
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtRtFromAddress 
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
               Index           =   0
               Left            =   120
               TabIndex        =   13
               Tag             =   "NN"
               Top             =   600
               Width           =   4485
            End
            Begin VB.TextBox txtRtFromAddress 
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
               Index           =   1
               Left            =   120
               TabIndex        =   14
               Tag             =   "NN"
               Top             =   960
               Width           =   4485
            End
            Begin VB.TextBox txtRtFromAddress 
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
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Tag             =   "NN"
               Top             =   1320
               Width           =   4485
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Tel:"
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
               Index           =   4
               Left            =   480
               TabIndex        =   106
               Top             =   1680
               Width           =   585
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Location:"
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
               Index           =   2
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   945
            End
         End
         Begin VB.ComboBox cmbFlights 
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
            ItemData        =   "frmOSCarTxfr.frx":8A93
            Left            =   720
            List            =   "frmOSCarTxfr.frx":8A9A
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   4080
            Width           =   4515
         End
         Begin VB.Frame Frame7 
            Caption         =   "To:"
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
            Left            =   -74880
            TabIndex        =   48
            Top             =   2160
            Width           =   4815
            Begin VB.ComboBox cmbLocation 
               Height          =   315
               Index           =   1
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtLocation 
               Height          =   285
               Index           =   1
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   78
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtPUToAddress 
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
               Index           =   0
               Left            =   120
               TabIndex        =   8
               Tag             =   "NN"
               Top             =   600
               Width           =   4485
            End
            Begin VB.TextBox txtPUToAddress 
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
               Index           =   1
               Left            =   120
               TabIndex        =   9
               Tag             =   "NN"
               Top             =   960
               Width           =   4485
            End
            Begin VB.TextBox txtPUToAddress 
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
               Index           =   2
               Left            =   120
               TabIndex        =   10
               Tag             =   "NN"
               Top             =   1320
               Width           =   4485
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Location:"
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
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   945
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "From:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -74880
            TabIndex        =   47
            Top             =   120
            Width           =   4815
            Begin VB.TextBox txtClientTel 
               Height          =   285
               Index           =   0
               Left            =   1200
               MaxLength       =   10
               TabIndex        =   109
               Top             =   1680
               Width           =   1575
            End
            Begin VB.ComboBox cmbLocation 
               Height          =   315
               Index           =   0
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtLocation 
               Height          =   285
               Index           =   0
               Left            =   2880
               MaxLength       =   10
               TabIndex        =   75
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtPUFromAddress 
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
               Index           =   2
               Left            =   120
               TabIndex        =   7
               Tag             =   "NN"
               Top             =   1320
               Width           =   4485
            End
            Begin VB.TextBox txtPUFromAddress 
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
               Index           =   1
               Left            =   120
               TabIndex        =   6
               Tag             =   "NN"
               Top             =   960
               Width           =   4485
            End
            Begin VB.TextBox txtPUFromAddress 
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
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Tag             =   "NN"
               Top             =   600
               Width           =   4485
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Tel:"
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
               Index           =   5
               Left            =   480
               TabIndex        =   108
               Top             =   1680
               Width           =   585
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "Location:"
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
               Left            =   120
               TabIndex        =   77
               Top             =   240
               Width           =   945
            End
         End
         Begin MSComCtl2.DTPicker dtpPUDateTime 
            Height          =   315
            Left            =   -71840
            TabIndex        =   11
            Top             =   4500
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MMM/yyyy  HH:mm"
            Format          =   61931523
            CurrentDate     =   38413
         End
         Begin VB.ComboBox cmbFlights 
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
            ItemData        =   "frmOSCarTxfr.frx":8AA8
            Left            =   -74280
            List            =   "frmOSCarTxfr.frx":8AAF
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   4080
            Width           =   4515
         End
         Begin MSComCtl2.DTPicker dtpRtnDateTime 
            Height          =   315
            Left            =   3160
            TabIndex        =   19
            Top             =   4500
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/MMM/yyyy  HH:mm"
            Format          =   61931523
            CurrentDate     =   38413
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Date && Time:"
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
            Left            =   1320
            TabIndex        =   50
            Top             =   4500
            Width           =   1725
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Flight:"
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
            Top             =   4080
            Width           =   585
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Date && Time:"
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
            Left            =   -73680
            TabIndex        =   46
            Top             =   4500
            Width           =   1725
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Flight:"
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
            Left            =   -74880
            TabIndex        =   45
            Top             =   4080
            Width           =   585
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
         Left            =   120
         TabIndex        =   54
         Top             =   990
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
         TabIndex        =   30
         Top             =   3870
         Width           =   1545
      End
   End
   Begin MSAdodcLib.Adodc datVendors 
      Height          =   375
      Left            =   4560
      Top             =   120
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
      Bindings        =   "frmOSCarTxfr.frx":8ABD
      DataSource      =   "datVendors"
      Height          =   360
      Left            =   360
      TabIndex        =   74
      Top             =   120
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
Attribute VB_Name = "frmOSCarTxfr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsRmks As ADODB.Recordset
Dim mobjEO As EO
Dim mstrAddr() As String
Dim mblnReturn As Boolean
Dim mstrPUFromLocation As String
Dim mstrPUToLocation As String
Dim mstrDOFromLocation As String
Dim mstrDOToLocation As String
'Timer
'Dim StartTime As Date
Dim sngNettCostGST As Single
Dim mstrPCAmend As String
Dim mstrTktNum As String
Dim sngGST As Single
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

Private Sub cmbFlights_Click(Index As Integer)
   Dim strTime As String
   
   If Index = 0 Then
      'strTime = Mid(cmbFlights(Index).Text, 14, 2) & "/" & Mid(cmbFlights(Index).Text, 16, 3) & "/" & Mid(cmbFlights(Index).Text, 19, 2) & " " & Mid(cmbFlights(Index).Text, 33, 2) & ":" & Mid(cmbFlights(Index).Text, 35, 2)
       If cmbFlights(Index).listindex > 0 Then strTime = gobjPNR.AirSeg(cmbFlights(Index).listindex).DepartDateTime
   Else
      'strTime = Mid(cmbFlights(Index).Text, 14, 2) & "/" & Mid(cmbFlights(Index).Text, 16, 3) & "/" & Mid(cmbFlights(Index).Text, 19, 2) & " " & Mid(cmbFlights(Index).Text, 38, 2) & ":" & Mid(cmbFlights(Index).Text, 40, 2)
       If cmbFlights(Index).listindex > 0 Then strTime = gobjPNR.AirSeg(cmbFlights(Index).listindex).ArriveDateTime
   End If
   If IsDate(strTime) Then
      If Index = 0 Then
       
         dtpPUDateTime.value = DateAdd("h", -2, strTime)
      Else
         dtpRtnDateTime.value = strTime
      End If
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

End Sub

Private Sub cmbLocation_Click(Index As Integer)
Dim lngC As Long

If cmbLocation(Index).Text = "Other" Then
    txtLocation(Index).Visible = True
Else
    txtLocation(Index).Text = ""
    txtLocation(Index).Visible = False
End If


lngC = UBound(mstrAddr)
If lngC > 3 Then lngC = 3

If cmbLocation(Index).listindex = 1 Then
    If Left(mstrAddr(0), 4) = "ATTN" Then
       For lngC = 1 To lngC
           Select Case Index
           Case 0:
           txtPUFromAddress(lngC - 1).Text = mstrAddr(lngC) & ""
           Case 1:
           txtPUToAddress(lngC - 1).Text = mstrAddr(lngC) & ""
           Case 2:
           txtRtFromAddress(lngC - 1).Text = mstrAddr(lngC) & ""
           Case 3:
           txtRtToAddress(lngC - 1).Text = mstrAddr(lngC) & ""
           End Select
       Next
    Else
       For lngC = 0 To lngC
           If lngC > 2 Then Exit For
           'txtPUToAddress(lngC).Text = mstrAddr(lngC) & ""
           Select Case Index
           Case 0:
           txtPUFromAddress(lngC).Text = mstrAddr(lngC) & ""
           Case 1:
           txtPUToAddress(lngC).Text = mstrAddr(lngC) & ""
           Case 2:
           txtRtFromAddress(lngC).Text = mstrAddr(lngC) & ""
           Case 3:
           txtRtToAddress(lngC).Text = mstrAddr(lngC) & ""
           End Select
       Next
    End If
Else
    
    For lngC = 1 To 3
    Select Case Index
    Case 0:
        txtPUFromAddress(lngC - 1).Text = ""
    Case 1:
        txtPUToAddress(lngC - 1).Text = ""
    Case 2:
        txtRtFromAddress(lngC - 1).Text = ""
    Case 3:
        txtRtToAddress(lngC - 1).Text = ""
    End Select
    Next

End If

Select Case Index
Case 0:
    For lngC = Me.txtPUFromAddress.LBound To Me.txtPUFromAddress.UBound
        Me.txtPUFromAddress(lngC).Enabled = True
    Next
    mstrPUFromLocation = cmbLocation(Index).Text
Case 1:
    For lngC = Me.txtPUToAddress.LBound To Me.txtPUToAddress.UBound
        Me.txtPUToAddress(lngC).Enabled = True
    Next
    mstrPUToLocation = cmbLocation(Index).Text
Case 2:
    For lngC = Me.txtRtFromAddress.LBound To Me.txtRtFromAddress.UBound
        Me.txtRtFromAddress(lngC).Enabled = True
    Next
    mstrDOFromLocation = cmbLocation(Index).Text
Case 3:
    For lngC = Me.txtRtToAddress.LBound To Me.txtRtToAddress.UBound
        Me.txtRtToAddress(lngC).Enabled = True
    Next
    mstrDOToLocation = cmbLocation(Index).Text
End Select

End Sub


Private Sub cmdCalculate_Click()
'Dim sngSF As Single
Dim sngCost As Single
Dim sngGSTPct As Single
Dim sngSellingPrice As Single
Dim sngMF As Single
Dim sngMFPct As Single
'If txtCommission.Text = "" Then txtCommission.Text = "0"
'If txtCost.Text = "" Then txtCost.Text = "0"

'sngSF = fConvertZero(txtCost.Text)
'If optCommType(1).Value Then
'    txtCommission.Text = fCommAmt(sngSF, fConvertZero(txtCommission.Text))
'    optCommType(0).Value = True
'End If

'sngSF = sngSF + fConvertZero(txtCommission.Text)

'Hard Code
'If chkAbsorb.Value = 0 Then
'   If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.Value = vbUnchecked Then
'      Me.txtMerchFee = fMerchantFee(sngSF, fGetMerchFee(gstrAgcyCountryCode))
'   Else
'       Me.txtMerchFee.Text = "0"
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

'Added on 6/7/2005: change of calculation logic

sngCost = fConvertZero(txtCost)
sngSellingPrice = fConvertZero(txtGrossSale) 'refer to gross sales

txtCost = Format(sngCost, gstrAgcyCurrFormat)
txtGrossSale = Format(sngSellingPrice, gstrAgcyCurrFormat)

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
       'sngMF = fMerchantFee(sngSellingPrice * (1 + sngGSTPct), fGetMerchFee(gstrAgcyCountryCode))
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
txtSellPrice = Format(sngSellingPrice, gstrAgcyCurrFormat)

cmdCalculate.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
Dim strMsg As String
datTouchEnd = Now
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
Dim freefields As String

If frmClientMI.MSXfreefields <> "" Then
        freefields = freefields & "/" & frmClientMI.MSXfreefields
End If
cmdDone.Enabled = False


gSysStartOthSvcsTime = Now
If gobjEO Is Nothing Or txtEONum = "" Then Call SetEOObj(gobjEO)
    'Timer
    'frmWait.Show
    Me.MousePointer = 11
    
    'Amended on 251108 by Jeremy to add freefields variable
    Call modOthSvcs.WriteOSToGDS(gobjEO, frmOthSvcs.datProducts.Recordset![Type], gStartOthSvcsTime, freefields)
    'Call modOthSvcs.WriteOSToGDS(gobjEO, frmOthSvcs.datProducts.Recordset![Type], gStartOthSvcsTime)
    
    'Remarks by JY ANG. Eliminate Queue Screen for non-air products
    'Call pTktQueue
    Unload frmClientMI
    Me.MousePointer = 0
    
    Log
    
    Unload Me
    
    'Unload frmWait
    Set gobjEO = Nothing
    
End Sub

Private Sub cmdEO_Click()
Dim lngC As Long
Dim strMsg As String
datTouchEnd = Now
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
'If gbolEOAmend = True Then
'   datSelectedVendor.DatabaseName = GetSetting("TPro", "Startup", "TProDBSource", "NOT FOUND")
'   datSelectedVendor.RecordSource = "SELECT * FROM tblVendors WHERE [VendorNumber] =  '" & dbcVendors.BoundText & "'"
'   datSelectedVendor.Refresh
'End If
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
'Remarks by JY ANG. Eliminate Queue Screen for non-air products
'Call pTktQueue
Set gobjEO = Nothing
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
 
       ' Else
        'lstItinRmks(0).RemoveItem lngC
        'lngC = lngC - 1
       ' End If
        
        
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

'Private Sub dbcVendors_Click(Area As Integer)
'   frmOthSvcs.dbcVendors.Text = dbcVendors.Text
'   With frmOthSvcs.datSelectedVendor
   'frmOthSvcs.datSelectedVendor.DatabaseName = gstrTProDBSource 'GetSetting("TPro", "Startup", "TProDBSource", "NOT FOUND")
'   .ConnectionString = gstrConn
'   .Mode = adModeRead
'   .CommandType = adCmdText
'   .RecordSource = "SELECT * FROM tblVendors WHERE [VendorNumber] =  '" & dbcVendors.BoundText & "'"
'   .Refresh
'   End With
'   GetVendorInfo
   
'End Sub
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
       'frmOthSvcs.datSelectedVendor.Recordset!VendorNumber
      'If dbcVendors.BoundText = "021222" Then
      If frmOthSvcs.datSelectedVendor.Recordset!Misc = True Then
         'fraVendorInfo.Enabled = True
         LockedText False
      Else
         'fraVendorInfo.Enabled = False
         LockedText True
      End If
   Else
      'frmOthSvcs.datSelectedVendor.Recordset!VendorNumber
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
Dim rsRemarks As ADODB.Recordset
Dim strSQL As String
Dim blnChkDate As Boolean
Dim blnChkVendor As Boolean
'Timer
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
SSTab1.Tab = 0
frmOSCarTxfr.Caption = "CWT TravelPro - " & frmOthSvcs.dbcProducts.Text
mblnReturn = True

'FormCenter




dtpPUDateTime.value = Date
dtpRtnDateTime.value = Date
dtpPUDateTime.value = Null
dtpRtnDateTime.value = Null

SSTab2.Tab = 0
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
        ElseIf ![RmkType] & "" = "E" Then
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
    If InStr(1, .DeliveryAddress, "@") <> 0 Then
       mstrAddr = Split(.DeliveryAddress, "@")
    Else
       mstrAddr = Split(.BillingAddress, "@")
    End If
    For lngC = 1 To .AirSegCount
        cmbFlights(0).AddItem .AirSeg(lngC).TextAirSeg & " " & Format(.AirSeg(lngC).DepartDateTime, "hhmm") & "-" & Format(.AirSeg(lngC).ArriveDateTime, "hhmm")
        cmbFlights(1).AddItem .AirSeg(lngC).TextAirSeg & " " & Format(.AirSeg(lngC).DepartDateTime, "hhmm") & "-" & Format(.AirSeg(lngC).ArriveDateTime, "hhmm")
    
    
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
        Me.cmbFOPType = "CX"
        
        If blnChkVendor = False Or blnChkDate = False Then
           promptCCError blnChkVendor, blnChkDate
        End If
    Else
        Me.cmbFOPType = "INV"
    End If
    
 End With
  
 Me.cmbFlights(0).listindex = 0
 Me.cmbFlights(1).listindex = 0
 
'optRtFromLocation(0).Value = True
'optRtToLocation(0).Value = True
'optPUFromLocation(0).Value = True
'optPUToLocation(0).Value = True
 
'Added on 8/4/05: Change to combobox, to allow other location to be specified in text box
For lngC = 0 To cmbLocation.Count - 1
    cmbLocation(lngC).AddItem "Airport", 0
    'cmbLocation(lngC).Tag = "Airport"
    cmbLocation(lngC).AddItem "Office", 1
    'cmbLocation(lngC).Tag = "Office"
    cmbLocation(lngC).AddItem "Home", 2
    'cmbLocation(lngC).Tag = "Home"
    cmbLocation(lngC).AddItem "Other", 3
    'cmbLocation(lngC).Tag = "Other"
    cmbLocation(lngC).listindex = 0
    txtLocation(lngC).Visible = False
Next
 
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

'added on 14/06/2005: for SG, auto enable/disable EO button
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
    End If
Else
        lblGST.Visible = False
        txtGST.Visible = False
        chkGSTAbsorb.Visible = False

End If

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
txtReplyEmail.Text = GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjPNR.Agent, gobjPNR.PCCOwner, False, True)

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
   Dim k As Integer
   Dim strFOP() As String
   Dim strRem() As String
   Dim strTemp() As String
    Dim strDescTemp() As String
    Dim strPUTemp() As String
    Dim strLocation() As String
    Dim strLoc() As String
    Dim strPickupFrom() As String
    Dim strReturnFrom() As String
   'Dim strDB As String
   
   cmdDone.Enabled = False
   dbcVendors.Visible = True
   
mstrPCAmend = frmEOAmend.lsvEO.SelectedItem.SubItems(1)
'strDB = GetSetting("TPro", "Startup", "TProDBSource", "NOT FOUND")


'datVendors.DatabaseName = gstrTProDBSource 'strDB


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
     gstrPreEOType = !EOType & ""
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
     If InStr(!PickUpFrom, ";") > 0 Then
     strPickupFrom = Split(!PickUpFrom, ";")
        If UBound(strPickupFrom) = 1 Then
            txtClientTel(0) = strPickupFrom(1)
        End If
     End If
     If InStr(!ReturnFrom, ";") > 0 Then
     strReturnFrom = Split(!ReturnFrom, ";")
        If UBound(strReturnFrom) = 1 Then
            txtClientTel(1) = strReturnFrom(1)
        End If
     End If
     
     'PICK UP INFORMATION:;  FROM:    (Office);   AFFINITY EQUITY PARTNERS (S) PTE LTD;   9 TEMASEK BOULEVARD;    SUNTEC TOWER 2 #27-03;  TO: (Airport);RETURN INFORMATION:;  FROM:   (Airport);  TO: (Office);   AFFINITY EQUITY PARTNERS (S) PTE LTD;   9 TEMASEK BOULEVARD;    SUNTEC TOWER 2 #27-03; - -
     If InStr(!PickUpFrom, ";") > 0 Then
        strTemp = Split(strPickupFrom(0), vbCrLf)
     Else
        strTemp = Split(!PickUpFrom, vbCrLf)
     End If
     For i = 0 To UBound(strTemp)
         If i = 0 Then
            If IsNumeric(strTemp(i)) Then
               If strTemp(i) < 3 Then
                  'optPUFromLocation(strTemp(i)).Value = True
                  cmbLocation(0).listindex = strTemp(i)
               ElseIf strTemp(i) = 3 Then
                  cmbLocation(0).listindex = strTemp(i)
                  
                  
                  If InStr(!Remarks, "RETURN INFORMATION") > 0 Then
                    strDescTemp = Split(!Remarks, "RETURN INFORMATION:")
                  Else
                    ReDim strDescTemp(0)
                    strDescTemp(0) = !Remarks
                  End If
                  strPUTemp = Split(strDescTemp(0), ";")
                  For k = 0 To UBound(strPUTemp)
                    If InStr(strPUTemp(k), "FROM:") > 0 Then
                         If InStr(strPUTemp(k), "(") = 0 Then Exit For
                         strLocation = Split(strPUTemp(k), "TO:")
                         'strLocation = Split(strLoc(k), "(")
                         'txtLocation(0) = Replace(strLocation(UBound(strLocation)), ")", "")
                         'If InStr(strLocation(0), ")") > 0 Then txtLocation(0) = Left(strLocation(0), InStr(strLocation(0), ")") - 1)
                         strLoc = Split(strLocation(0), "(")
                         If InStr(strLoc(UBound(strLoc)), ")") > 0 Then txtLocation(0) = Left(strLoc(UBound(strLoc)), InStr(strLoc(UBound(strLoc)), ")") - 1)

                         Exit For
                    End If
                  Next
               End If
            End If
         ElseIf i = 4 Then
            Exit For
         Else
            txtPUFromAddress(i - 1) = strTemp(i)
         End If
     Next
     strTemp = Split(!PickUpTo, vbCrLf)
     For i = 0 To UBound(strTemp)
         If i = 0 Then
            If IsNumeric(strTemp(i)) Then
               If strTemp(i) < 3 Then
                  'optPUToLocation(strTemp(i)).Value = True
                  cmbLocation(1).listindex = strTemp(i)
               ElseIf strTemp(i) = 3 Then
               cmbLocation(1).listindex = strTemp(i)
                  If InStr(!Remarks, "RETURN INFORMATION") > 0 Then
                    strDescTemp = Split(!Remarks, "RETURN INFORMATION:")
                  Else
                    ReDim strDescTemp(0)
                    strDescTemp(0) = !Remarks
                  End If
                  strPUTemp = Split(strDescTemp(0), ";")
                  For k = 0 To UBound(strPUTemp)
                    If InStr(strPUTemp(k), "TO:") > 0 Then
                    If InStr(strPUTemp(k), "(") = 0 Then Exit For
                    
                         strLocation = Split(strPUTemp(k), "TO:")
                         strLoc = Split(strLocation(UBound(strLocation)), "(")
                         If InStr(strLoc(UBound(strLoc)), ")") > 0 Then txtLocation(1) = Left(strLoc(UBound(strLoc)), InStr(strLoc(UBound(strLoc)), ")") - 1)

                         'txtLocation(1) = Replace(strLocation(UBound(strLocation)), ")", "")
                         'If InStr(strLocation(UBound(strLocation)), ")") > 0 Then txtLocation(1) = Left(strLocation(UBound(strLocation)), InStr(strLocation(UBound(strLocation)), ")") - 1)
                         Exit For
                    End If
                  Next
                End If
                End If
         ElseIf i = 4 Then
            Exit For
         Else
            txtPUToAddress(i - 1) = strTemp(i)
         End If
     Next
     If InStr(!PickUpFlight, ".") > 0 Then
        For i = 0 To cmbFlights(0).ListCount - 1
         If InStr(cmbFlights(0).List(i), Mid(!PickUpFlight, InStr(!PickUpFlight, "."))) > 0 Then
          
              cmbFlights(0).listindex = i
              Exit For
           End If
        Next
     End If
     'modified on 2/3/2005
     If !PickUpTime <> CdatDefaultDate Then
        dtpPUDateTime.value = !PickUpTime
     Else
        dtpPUDateTime.value = Null
     End If
     'strTemp = Split(IIf(InStr(!ReturnFrom, ";") > 0, strReturnFrom(0), !ReturnFrom), vbCrLf)
     If InStr(!ReturnFrom, ";") > 0 Then
        strTemp = Split(strReturnFrom(0), vbCrLf)
     Else
        strTemp = Split(!ReturnFrom, vbCrLf)
     End If
     For i = 0 To UBound(strTemp)
         If i = 0 Then
            If IsNumeric(strTemp(i)) Then
               If strTemp(i) < 3 Then
                  'optRtFromLocation(strTemp(i)).Value = True
                  cmbLocation(2).listindex = strTemp(i)
              ElseIf strTemp(i) = 3 Then
                  cmbLocation(2).listindex = strTemp(i)
                    If InStr(!Remarks, "RETURN INFORMATION") > 0 And InStr(!Remarks, "PICKUP INFORMATION") >= 0 Then
                        strDescTemp = Split(!Remarks, "RETURN INFORMATION:")
                    Else
                    ReDim strDescTemp(1)
                        strDescTemp(1) = !Remarks
                    End If
                    
                     strPUTemp = Split(strDescTemp(1), ";")
                        For k = 0 To UBound(strPUTemp)
                          If InStr(strPUTemp(k), "FROM:") > 0 Then
                          If InStr(strPUTemp(k), "(") = 0 Then Exit For
                               strLocation = Split(strPUTemp(k), "TO:")
                               'txtLocation(2) = Replace(strLocation(UBound(strLocation)), ")", "")
                               strLoc = Split(strLocation(0), "(")
                               If InStr(strLoc(UBound(strLoc)), ")") > 0 Then txtLocation(2) = Left(strLoc(UBound(strLoc)), InStr(strLoc(UBound(strLoc)), ")") - 1)
                               Exit For
                          End If
                        Next
                End If
            
            End If
         ElseIf i = 4 Then
            Exit For
         Else
            txtRtFromAddress(i - 1) = strTemp(i)
         End If
     Next
     strTemp = Split(!ReturnTo, vbCrLf)
     For i = 0 To UBound(strTemp)
         If i = 0 Then
            If IsNumeric(strTemp(i)) Then
               If strTemp(i) < 3 Then
                  'optRtToLocation(strTemp(i)).Value = True
                  cmbLocation(3).listindex = strTemp(i)
               ElseIf strTemp(i) = 3 Then
               cmbLocation(3).listindex = strTemp(i)
               If InStr(!Remarks, "RETURN INFORMATION") > 0 And InStr(!Remarks, "PICKUP INFORMATION") >= 0 Then
                        strDescTemp = Split(!Remarks, "RETURN INFORMATION:")
                Else
                    ReDim strDescTemp(1)
                    strDescTemp(1) = !Remarks
                End If
                    
                     strPUTemp = Split(strDescTemp(1), ";")
                        For k = 0 To UBound(strPUTemp)
                          If InStr(strPUTemp(k), "TO:") > 0 Then
                          If InStr(strPUTemp(k), "(") = 0 Then Exit For
                               strLocation = Split(strPUTemp(k), "TO:")
                               'txtLocation(3) = Replace(strLocation(UBound(strLocation)), ")", "")
                               'If InStr(strLocation(UBound(strLocation)), ")") > 0 Then txtLocation(3) = Left(strLocation(UBound(strLocation)), InStr(strLocation(UBound(strLocation)), ")") - 1)
                                strLoc = Split(strLocation(UBound(strLocation)), "(")
                                If InStr(strLoc(UBound(strLoc)), ")") > 0 Then txtLocation(3) = Left(strLoc(UBound(strLoc)), InStr(strLoc(UBound(strLoc)), ")") - 1)

                               Exit For
                          End If
                        Next
            
               End If
            End If
         ElseIf i = 4 Then
            Exit For
         Else
            txtRtToAddress(i - 1) = strTemp(i)
         End If
     Next
     If InStr(!ReturnFlight, ".") > 0 Then
        For i = 0 To cmbFlights(1).ListCount - 1
        
           If InStr(cmbFlights(1).List(i), Mid(!ReturnFlight, InStr(!ReturnFlight, "."))) > 0 Then
              cmbFlights(1).listindex = i
              Exit For
           End If
        Next
     End If
     'modified on 2/3/2005
     If !ReturnTime <> CdatDefaultDate Then
        dtpRtnDateTime.value = !ReturnTime
     Else
        dtpRtnDateTime.value = Null
     End If
     
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

Private Function fGST(TotalCharge As Single) As Single
Dim sngAmt As Single
Dim sngPct As Single

sngPct = frmOthSvcs.datProducts.Recordset![GST] * 0.01
sngAmt = sngPct * TotalCharge
'fGST = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP")
fGST = Format(sngAmt, "0.00")


End Function

Private Function validData() As Boolean
Dim strMsg As String
Dim i As Integer

If gbolEOAmend = False Then
   If frmOthSvcs.dbcVendors.Text = "" Then strMsg = strMsg & "Select vendor for this transaction..." & Chr(13)
End If

If Me.cmbLocation(0).listindex <> 0 Then
    If Me.txtPUFromAddress(0).Text = "" Then
        strMsg = strMsg & "Need Pick Up From Address..." & Chr(13)
    End If
End If

If Me.cmbLocation(1).listindex <> 0 Then
    If Me.txtPUToAddress(0).Text = "" Then
        strMsg = strMsg & "Need Pick Up To Address..." & Chr(13)
    End If
End If

If IsNull(dtpPUDateTime.value) And IsNull(dtpRtnDateTime.value) Then
    strMsg = strMsg & "Need to enter Pick up or Return date time ..." & Chr(13)
End If

'Disble validation to enable back date for car transfer
'If IsNull(dtpPUDateTime.value) = False Then
'    If Me.dtpPUDateTime.value < Date Then strMsg = strMsg & "Pick up date cannot be past..." & Chr(13)
'End If
'If IsNull(dtpRtnDateTime.value) = False Then
'    If Me.dtpRtnDateTime.value < Date Then strMsg = strMsg & "Return date cannot be past..." & Chr(13)
'End If

If mblnReturn Then
    If Me.cmbLocation(2).listindex > 0 And Me.txtRtFromAddress(0).Text = "" Then strMsg = strMsg & "Need Return From Address..." & Chr(13)
    If Me.cmbLocation(3).listindex > 0 And Me.txtRtToAddress(0).Text = "" Then strMsg = strMsg & "Need Return To Address..." & Chr(13)
'    If Me.dtpRtnDateTime.Value < Date Then strMsg = strMsg & "Return date cannot be past..." & Chr(13)
End If

If txtSellPrice.Text = "" Or txtSellPrice.Text = "0" Then strMsg = strMsg & "Need to calculate Selling Price..." & Chr(13)
If cmbFOPType.Text = "" Then strMsg = strMsg & "Need form of payment..." & Chr(13)
If cmbFOPType.Text = "CX" Then
    If cmbCCType.Text = "" Then strMsg = strMsg & "Need valid credit vendor code..." & Chr(13)
    If txtCCNum = "" Then strMsg = strMsg & "Need valid credit card number..." & Chr(13)
    If LastDate(dtpCCExp.value) < Date Then strMsg = strMsg & "Need valid expiration date..." & Chr(13)
    If (txtCCNum.Text <> "" And cmbCCType.Text <> "") Then If ValidCCNum(cmbCCType.Text, txtCCNum.Text) = False Then strMsg = strMsg & "Credit card number is invalid or wrong card vendor selected ..." & Chr(13)
End If
If UCase(gstrAgcyCountryCode) = "SG" Then
    'Add on 200106: TMP card must absorb MF
    'If IsTMPCard(Left(UCase(cmbCCType.Text), 2), Left(UCase(txtCCNum.Text), 7)) Then
    'If (cmbFOPType <> "INV" And Left(UCase(cmbCCType.Text), 2) = "DC" And _
        Left(UCase(txtCCNum.Text), 7) = "3644033") And chkAbsorb.value <> 1 Then
     'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    If (cmbFOPType <> "INV" And _
      IsTMPCard(Left(UCase(cmbCCType.Text), 2), UCase(txtCCNum.Text))) And _
      chkAbsorb.value <> 1 Then
        strMsg = strMsg & "Need to tick aborbed merchant fee for TMP card and recalculate selling price" & Chr(13)
    End If
End If

If InStr(1, UCase(txtVendor), "/PC") <> 0 Then
   strMsg = strMsg & "'/PC' exist in Vendor Name.." & Chr(13)
End If

If InStr(1, UCase(txtTel), "/PC") <> 0 Then
   strMsg = strMsg & "'/PC' exist in Contact No..." & Chr(13)
End If

For i = 0 To 3
   If InStr(1, UCase(cmbLocation(i).Text), "/PC") <> 0 Then
      strMsg = strMsg & "'/PC' exist in Location drop down box..." & Chr(13)
      Exit For
   End If
Next

For i = 0 To 3
   If InStr(1, UCase(txtLocation(i).Text), "/PC") <> 0 Then
      strMsg = strMsg & "'/PC' exist in Location text box..." & Chr(13)
      Exit For
   End If
Next

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
    'If gbolEOAmend = False Then cmdDone.Enabled = True
End If

End Function




Private Sub Form_Unload(Cancel As Integer)
'Preethi - V1.1.1 20100831 - IR2 - Client MI screen is populated with old data
Unload frmClientMI
End Sub

'Private Sub optPUFromLocation_Click(Index As Integer)
'Dim lngC As Long

'lngC = UBound(mstrAddr)
'If lngC > 3 Then lngC = 3

'If Index = 1 Then
    
'    If Left(mstrAddr(0), 4) = "ATTN" Then
'       For lngC = 1 To lngC
'          txtPUFromAddress(lngC - 1).Text = mstrAddr(lngC) & ""
'       Next
'    Else
'       For lngC = 0 To lngC
'          If lngC > 2 Then Exit For
'          txtPUFromAddress(lngC).Text = mstrAddr(lngC) & ""
'       Next
'    End If
'Else
    
'    For lngC = 1 To 3
'        txtPUFromAddress(lngC - 1).Text = ""
'    Next

'End If

'For lngC = Me.txtPUFromAddress.LBound To Me.txtPUFromAddress.UBound
'    Me.txtPUFromAddress(lngC).Enabled = True
'Next
'mstrPUFromLocation = optPUFromLocation(Index).Tag

'End Sub

'Private Sub optPUToLocation_Click(Index As Integer)
'Dim lngC As Long

'lngC = UBound(mstrAddr)
'If lngC > 3 Then lngC = 3

'If Index = 1 Then
'    If Left(mstrAddr(0), 4) = "ATTN" Then
'       For lngC = 1 To lngC
'           txtPUToAddress(lngC - 1).Text = mstrAddr(lngC) & ""
'       Next
'    Else
'       For lngC = 0 To lngC
'           If lngC > 2 Then Exit For
'           txtPUToAddress(lngC).Text = mstrAddr(lngC) & ""
'       Next
'    End If
'Else
    
'    For lngC = 1 To 3
'        txtPUToAddress(lngC - 1).Text = ""
'    Next

'End If


'For lngC = Me.txtPUToAddress.LBound To Me.txtPUToAddress.UBound
'    Me.txtPUToAddress(lngC).Enabled = True
'Next
'mstrPUToLocation = optPUToLocation(Index).Tag

'End Sub

'Private Sub optRtFromLocation_Click(Index As Integer)
'Dim lngC As Long


'lngC = UBound(mstrAddr)

'If lngC > 3 Then lngC = 3

'If Index = 1 Then
'    If Left(mstrAddr(0), 4) = "ATTN" Then
'       For lngC = 1 To lngC
'           txtRtFromAddress(lngC - 1).Text = mstrAddr(lngC) & ""
'       Next
'    Else
'       For lngC = 0 To lngC
'           If lngC > 2 Then Exit For
'           txtRtFromAddress(lngC).Text = mstrAddr(lngC) & ""
'       Next
'    End If
'Else
    
'    For lngC = 1 To 3
'        txtRtFromAddress(lngC - 1).Text = ""
'    Next

'End If

'For lngC = Me.txtRtFromAddress.LBound To Me.txtRtFromAddress.UBound
'    Me.txtRtFromAddress(lngC).Enabled = True
'Next
'mstrDOFromLocation = optRtFromLocation(Index).Tag

'End Sub

'Private Sub optRtToLocation_Click(Index As Integer)
'Dim lngC As Long

'lngC = UBound(mstrAddr)
'If lngC > 3 Then lngC = 3

'If Index = 1 Then
'    If Left(mstrAddr(0), 4) = "ATTN" Then
'       For lngC = 1 To lngC
'           txtRtToAddress(lngC - 1).Text = mstrAddr(lngC) & ""
'       Next
'    Else
'       For lngC = 0 To lngC
'           If lngC > 2 Then Exit For
'           txtRtToAddress(lngC).Text = mstrAddr(lngC) & ""
'       Next
'    End If
'Else
    
'    For lngC = 1 To 3
'        txtRtToAddress(lngC - 1).Text = ""
'    Next

'End If

'For lngC = Me.txtRtToAddress.LBound To Me.txtRtToAddress.UBound
'    Me.txtRtToAddress(lngC).Enabled = True
'Next
'mstrDOToLocation = optRtToLocation(Index).Tag

'End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
'    If SSTab2.Tab > 0 Then mblnReturn = True
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

Private Sub txtCCNum_GotFocus()
Call pSetSelected
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

Private Sub txtCost_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtFreeRmk_KeyPress(KeyAscii As Integer)

KeyAscii = fAllowAlphaNumeric(KeyAscii, "#$*()/.: ?@")

End Sub

Private Sub txtGrossSale_Change()
   EnableCalculate
End Sub

Private Sub EnableCalculate()
   cmdCalculate.Enabled = True
End Sub

Private Sub txtMerchFee_GotFocus()
Call pSetSelected
End Sub

Private Sub txtPassengerID_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii)
End Sub

Private Sub txtPUFromAddress_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtPUToAddress_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtRtFromAddress_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtRtToAddress_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtSellPrice_GotFocus()
Call pSetSelected
End Sub

Private Sub txtFreeRmk_GotFocus()
Call pSetSelected
End Sub

Private Sub SetEOObj(ByRef objEO As EO)
Dim lngC As Long
Dim strPU As String
Dim strRT As String
Dim strPaxName As String

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
'Set objEO = New EO

With objEO
    
    If gbolEOAmend Then
       .EONumber = txtEONum
       'frmOthSvcs.datProducts.DatabaseName = gstrTProDBSource
       'frmOthSvcs.datProducts.RecordSource = "SELECT * FROM tblProductCodes where ProductCode = '" & mstrPCAmend & "'"
       'frmOthSvcs.datProducts.Refresh
    End If
    
    
    '   txtEONum = .EONumber
       
    
    .BillingDescription = ""
    .CN = gobjPNR.CN
    .CommissionAmt = fConvertZero(txtCommission.Text) '+ fConvertZero(txtMerchFee)
    .Cost = fConvertZero(txtCost.Text)
    .CreatedBy = gobjHost.AgentSine
    '.CreatedByName = gobjHost.AgentName
    .CreatedByName = gobjHost.AgentProfile
    .CreatedByPCC = gobjHost.AgentPCC
    .CreateDtTm = Now()
    .DescriptionLineAdd frmOthSvcs.dbcProducts.Text
      'modified on 31/03/2005: Change format on TUR line description
    If Not IsNull(dtpPUDateTime.value) Then
        'For lngC = 0 To cmbLocation(0).ListCount
        '    If optPUFromLocation(lngC).Value Then
        If cmbLocation(0).listindex = 3 Then
                strPU = IIf(txtLocation(0).Text <> "", Trim(txtLocation(0).Text), mstrPUFromLocation)
        Else
                strPU = cmbLocation(0).Text
        End If
        '    End If
        'Next
        'For lngC = 0 To optPUToLocation.Count - 1
        '    If optPUToLocation(lngC).Value Then
        If cmbLocation(1).listindex = 3 Then
                strPU = strPU & "/" & IIf(txtLocation(1).Text <> "", Trim(txtLocation(1).Text), mstrPUToLocation) & " PICKUP " & Format(Me.dtpPUDateTime.value, "hhnn") & " HRS" & " " & Format(Me.dtpPUDateTime.value, "ddmmm")
        Else
                strPU = strPU & "/" & cmbLocation(1).Text & " PICKUP " & Format(Me.dtpPUDateTime.value, "hhnn") & " HRS" & " " & Format(Me.dtpPUDateTime.value, "ddmmm")
        End If
        
        
        '    End If
        'Next
        '.DescriptionLineAdd Format(DateAdd("d", -1, dtpPUDateTime.Value), "ddmmm") & "-" & frmOthSvcs.dbcProducts.Text & " " & strPU
        '.DescriptionLineAdd Format(DateAdd("d", -1, dtpPUDateTime.Value), "ddmmm") & "-" & txtVendor & " " & "-PHONE" & " " & txtTel
        
        .DescriptionLineAdd Format(IIf(dtpPUDateTime.value >= Date, dtpPUDateTime.value, Date), "ddmmm") & "-" & frmOthSvcs.dbcProducts.Text & " " & strPU
        .DescriptionLineAdd Format(IIf(dtpPUDateTime.value >= Date, dtpPUDateTime.value, Date), "ddmmm") & "-" & txtVendor & " " & "-PHONE" & " " & txtTel

    End If
    If mblnReturn And Not IsNull(dtpRtnDateTime.value) Then
        'For lngC = 0 To optRtFromLocation.Count - 1
        '    If optRtFromLocation(lngC).Value Then
        '        strRT = optPUFromLocation(lngC).Caption
        '    End If
        'Next
        'For lngC = 0 To optRtToLocation.Count - 1
        '    If optRtToLocation(lngC).Value Then
        '        strRT = strRT & "/" & optRtToLocation(lngC).Caption & " PICKUP " & Format(Me.dtpRtnDateTime.Value, "hhnn") & " " & Format(Me.dtpRtnDateTime.Value, "ddmmm")
        '    End If
        'Next
        If cmbLocation(2).listindex = 3 Then
                strRT = IIf(txtLocation(2).Text <> "", Trim(txtLocation(2).Text), mstrDOToLocation)
        Else
                strRT = cmbLocation(2).Text
        End If
        
        If cmbLocation(3).listindex = 3 Then
                 strRT = strRT & "/" & IIf(txtLocation(3).Text <> "", Trim(txtLocation(3).Text), mstrDOToLocation) & " PICKUP " & Format(Me.dtpRtnDateTime.value, "hhnn") & " HRS" & " " & Format(Me.dtpRtnDateTime.value, "ddmmm")
        Else
                 strRT = strRT & "/" & cmbLocation(3).Text & " PICKUP " & Format(Me.dtpRtnDateTime.value, "hhnn") & " HRS" & " " & Format(Me.dtpRtnDateTime.value, "ddmmm")
        End If
        
        '.DescriptionLineAdd Format(DateAdd("d", 1, dtpRtnDateTime.Value), "ddmmm") & "-" & frmOthSvcs.dbcProducts.Text & " " & strRT
        .DescriptionLineAdd Format(IIf(dtpRtnDateTime.value >= Date, dtpRtnDateTime.value, Date), "ddmmm") & "-" & frmOthSvcs.dbcProducts.Text & " " & strRT
        If IsNull(dtpPUDateTime) Then
         '   .DescriptionLineAdd Format(DateAdd("d", 1, dtpRtnDateTime.Value), "ddmmm") & "-" & txtVendor & " " & "-PHONE" & " " & txtTel
        .DescriptionLineAdd Format(IIf(dtpRtnDateTime.value >= Date, dtpRtnDateTime.value, Date), "ddmmm") & "-" & txtVendor & " " & "-PHONE" & " " & txtTel
        End If
        
    End If
    .FOP = cmbFOPType.Text & IIf(cmbFOPType.Text = "CX", "/" & cmbCCType.Text & "/" & txtCCNum.Text & "/" & Format(dtpCCExp.value, "MMYY"), "") & "/" & 0 & "/" & chkWaiveMercFee.value
    
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
    
    'If UCase(gstrAgcyCountryCode) = "SG" Then
        '.ServiceDate = DateAdd("M", 6, Date)
    'Else
        '.ServiceDate = DateAdd("d", 90, Date) 'DateAdd("m", 3, gobjPNR.AirSeg(lstFlights.ListCount).ArriveDateTime)
    'End If
    'ZhiSam - V1.2.23 20130911 - CR-289 - E-Invoice: Due and Paid Lines Date
    If bfunctCheckRTLine = True Then
        .ServiceDate = dtfunctRTDate
    Else
        .ServiceDate = DateAdd("d", 90, Date)
    End If
    
    '.TaxAdd CSng(txtGST.Text), IIf(txtGST.Text = "0", "", "GST")
    '.TaxAdd 0, ""
    
    If UCase(gstrAgcyCountryCode) = "SG" Then
        '.TaxAdd fConvertZero(txtGST.Text), "GST"
        .TaxAdd sngGST, "GST"
        .NettGST = sngNettCostGST
    Else
        .TaxAdd 0, ""
        .NettGST = 0
    End If
    
    If gbolEOAmend Then
       .TicketNumber = mstrTktNum
    Else
       .TicketNumber = "0000"
    End If
    '.VendorCode = frmOthSvcs.dbcVendors.BoundText
    '.Email = frmOthSvcs.datVendors.Recordset![Email] & ""
    '.FaxNo = frmOthSvcs.datVendors.Recordset![FaxNumber] & ""
    'If gbolEOAmend Then
    '   .VendorCode = datSelectedVendor.Recordset![VendorNumber] 'frmEOAmend.lsvEO.SelectedItem.SubItems(5)
    '   .Email = datSelectedVendor.Recordset![Email] & ""
    '   .FaxNo = datSelectedVendor.Recordset![FaxNumber] & ""
    '   .VendorName = datSelectedVendor.Recordset!VendorName
    '   .Address1 = datSelectedVendor.Recordset!Address1
    '   .Address2 = datSelectedVendor.Recordset!Address2
    '   .City = datSelectedVendor.Recordset!City
    '   .Country = datSelectedVendor.Recordset!Country
    'Else
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
    'End If
    .ContactPerson = txtContact

    
    
    If Not IsNull(dtpPUDateTime.value) Then
    .RemarkAdd "PICK UP INFORMATION:"
    If cmbLocation(0).listindex = 3 Then
    .RemarkAdd "  FROM:" & Chr(9) & IIf(txtLocation(0).Text <> "", "(" & Trim(txtLocation(0).Text) & ")", mstrPUFromLocation)
    Else
    .RemarkAdd "  FROM:" & Chr(9) & IIf(mstrPUFromLocation <> "", "(" & mstrPUFromLocation & ")", "")
    End If
    For lngC = 0 To 2
        If txtPUFromAddress(lngC).Text = "" Then Exit For
        .RemarkAdd Chr(9) & txtPUFromAddress(lngC).Text
    Next
    If Trim(txtClientTel(0)) <> "" Then .RemarkAdd Chr(9) & "TEL:" & Trim(txtClientTel(0))
    
    If cmbLocation(1).listindex = 3 Then
   .RemarkAdd "  TO:" & Chr(9) & IIf(txtLocation(1).Text <> "", "(" & Trim(txtLocation(1).Text) & ")", mstrPUToLocation)
    Else
    .RemarkAdd "  TO:" & Chr(9) & IIf(mstrPUToLocation <> "", "(" & mstrPUToLocation & ")", "")
    End If
    For lngC = 0 To 2
        If txtPUToAddress(lngC).Text = "" Then Exit For
        .RemarkAdd Chr(9) & txtPUToAddress(lngC).Text
    Next
    If cmbFlights(0).listindex > 0 Then .RemarkAdd "   FLIGHT INFO -  " & Mid(cmbFlights(0).Text, 5, 23) & " " & IIf(cmbFlights(0).listindex > 0, Format(gobjPNR.AirSeg(cmbFlights(0).listindex).DepartDateTime, "hhmm"), "")
    End If
    If mblnReturn And Not IsNull(dtpRtnDateTime.value) Then
        .RemarkAdd "RETURN INFORMATION:"
        If cmbLocation(2).listindex = 3 Then
        .RemarkAdd "  FROM:" & Chr(9) & IIf(txtLocation(2).Text <> "", "(" & Trim(txtLocation(2).Text) & ")", mstrDOFromLocation)
        Else
         .RemarkAdd "  FROM:" & Chr(9) & IIf(mstrDOFromLocation <> "", "(" & mstrDOFromLocation & ")", "")
        End If
        For lngC = 0 To 2
            If txtRtFromAddress(lngC).Text = "" Then Exit For
            .RemarkAdd Chr(9) & txtRtFromAddress(lngC).Text
        Next
        If Trim(txtClientTel(1)) <> "" Then .RemarkAdd Chr(9) & "TEL:" & Trim(txtClientTel(1))
        If cmbLocation(3).listindex = 3 Then
            .RemarkAdd "  TO:" & Chr(9) & IIf(txtLocation(3).Text <> "", "(" & Trim(txtLocation(3).Text) & ")", mstrDOToLocation)
        Else
            .RemarkAdd "  TO:" & Chr(9) & IIf(mstrDOToLocation <> "", "(" & mstrDOToLocation & ")", "")
        End If
        For lngC = 0 To 2
            If txtRtToAddress(lngC).Text = "" Then Exit For
            .RemarkAdd Chr(9) & txtRtToAddress(lngC).Text
        Next
        If cmbFlights(1).listindex > 0 Then .RemarkAdd "   FLIGHT INFO -  " & Mid(cmbFlights(1).Text, 5, 23) & " " & IIf(cmbFlights(0).listindex > 0, Format(gobjPNR.AirSeg(cmbFlights(1).listindex).DepartDateTime, "hhmm"), "")
    End If
    
    .RemarkAdd " - - - - - - - - - - - - - - - - - - - - - - - - - "
    
    For lngC = 0 To lstEORmks(1).ListCount - 1
        .RemarkAdd lstEORmks(1).List(lngC)
    Next
    For lngC = 0 To lstItinRmks(1).ListCount - 1
        .RIRemarkAdd lstItinRmks(1).List(lngC)
    Next
    
    .MerchFee = fConvertZero(txtMerchFee)
    .CWTAbsorb = IIf(chkAbsorb.value = 1, True, False)
    
    .PickUpFrom = PickUpFrom & ";" & txtClientTel(0)
    .PickUpTo = PickUpTo
    
    If Not IsNull(dtpPUDateTime.value) Then .PickUpTime = dtpPUDateTime.value
    
    .PickUpFlight = cmbFlights(0).Text
    .ReturnFrom = ReturnFrom & ";" & txtClientTel(1)
    .ReturnTo = ReturnTo
    If Not IsNull(dtpRtnDateTime.value) Then .ReturnTime = dtpRtnDateTime.value
    .ReturnFlight = cmbFlights(1).Text
    .ListBoxRem = ListBoxRemark
    .PassengerID = IIf(txtPassengerID = "", "1", txtPassengerID)
    .ReplyEmail = Trim(UCase(txtReplyEmail.Text))
End With
End Sub

Private Function PickUpFrom() As String
   Dim i As Integer
   
   'For i = 0 To 3
   '   If optPUFromLocation(i).Value Then
        PickUpFrom = cmbLocation(0).listindex
   '      Exit For
   '   End If
   'Next
   For i = 0 To 2
      If Trim(txtPUFromAddress(i).Text) <> "" Then
         PickUpFrom = PickUpFrom & vbCrLf & Trim(txtPUFromAddress(i).Text)
      End If
   Next
End Function

Private Function PickUpTo() As String
   Dim i As Integer
   
   'For i = 0 To 3
   '   If optPUToLocation(i).Value Then
         PickUpTo = cmbLocation(1).listindex
   '      Exit For
   '   End If
   'Next
   For i = 0 To 2
      If Trim(txtPUToAddress(i).Text) <> "" Then
         PickUpTo = PickUpTo & vbCrLf & Trim(txtPUToAddress(i).Text)
      End If
   Next
End Function

Private Function ReturnFrom() As String
   Dim i As Integer
   
   'For i = 0 To 3
   '   If optRtFromLocation(i).Value Then
         ReturnFrom = cmbLocation(2).listindex
   '      Exit For
   '   End If
   'Next
   For i = 0 To 2
      If Trim(txtRtFromAddress(i).Text) <> "" Then
         ReturnFrom = ReturnFrom & vbCrLf & Trim(txtRtFromAddress(i).Text)
      End If
   Next
End Function

Private Function ReturnTo() As String
   Dim i As Integer
   
   'For i = 0 To 3
   '   If optRtToLocation(i).Value Then
         ReturnTo = cmbLocation(3).listindex
   '      Exit For
   '   End If
   'Next
   For i = 0 To 2
      If Trim(txtRtToAddress(i).Text) <> "" Then
         ReturnTo = ReturnTo & vbCrLf & Trim(txtRtToAddress(i).Text)
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

Sub FormCenter()
    Top = (Screen.Height * 0.95) / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
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


