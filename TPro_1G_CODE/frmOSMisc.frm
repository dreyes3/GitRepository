VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOSMisc 
   ClientHeight    =   7200
   ClientLeft      =   1170
   ClientTop       =   1365
   ClientWidth     =   11160
   LinkTopic       =   "CWT Travel Pro - "
   ScaleHeight     =   7200
   ScaleWidth      =   11160
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
      Bindings        =   "frmOSMisc.frx":0000
      DataSource      =   "datVendors"
      Height          =   360
      Left            =   360
      TabIndex        =   0
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   5
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
      TabCaption(0)   =   "Service Info"
      TabPicture(0)   =   "frmOSMisc.frx":0019
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTkt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtContact"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraHotel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraAir"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Remarks"
      TabPicture(1)   =   "frmOSMisc.frx":0035
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "txtFreeRmk"
      Tab(1).Control(2)=   "cmdFreeRmkToEO"
      Tab(1).Control(3)=   "cmdFreeRmkToItin"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "Frame5"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "MI"
      TabPicture(2)   =   "frmOSMisc.frx":0051
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLabels(40)"
      Tab(2).Control(1)=   "lblLabels(41)"
      Tab(2).Control(2)=   "lvwRealECodes"
      Tab(2).Control(3)=   "lvwMissECodes"
      Tab(2).Control(4)=   "lvwECodes"
      Tab(2).Control(5)=   "cmdClientMI"
      Tab(2).Control(6)=   "Frame4"
      Tab(2).Control(7)=   "txtMS"
      Tab(2).Control(8)=   "txtRS"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Insurance"
      TabPicture(3)   =   "frmOSMisc.frx":006D
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtInsAdd(2)"
      Tab(3).Control(1)=   "txtInsAdd(1)"
      Tab(3).Control(2)=   "Frame7"
      Tab(3).Control(3)=   "txtInsAdd(0)"
      Tab(3).Control(4)=   "cmbGeoArea"
      Tab(3).Control(5)=   "cmbInsPlan"
      Tab(3).Control(6)=   "txtInsDays"
      Tab(3).Control(7)=   "dtpInsFromDate"
      Tab(3).Control(8)=   "Label24"
      Tab(3).Control(9)=   "lblGeoArea"
      Tab(3).Control(10)=   "Label21"
      Tab(3).Control(11)=   "Label22"
      Tab(3).Control(12)=   "Label23"
      Tab(3).ControlCount=   13
      TabCaption(4)   =   "Vendor Info"
      TabPicture(4)   =   "frmOSMisc.frx":0089
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraVendorInfo"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraAir 
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
         Height          =   1440
         Left            =   120
         TabIndex        =   135
         Top             =   4920
         Visible         =   0   'False
         Width           =   5895
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
            Height          =   1080
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   136
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.TextBox txtRS 
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
         Height          =   340
         Left            =   -67440
         MaxLength       =   3
         TabIndex        =   124
         Tag             =   "BY-"
         Top             =   960
         Width           =   645
      End
      Begin VB.TextBox txtMS 
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
         Height          =   340
         Left            =   -67440
         MaxLength       =   3
         TabIndex        =   123
         Tag             =   "BY-"
         Top             =   3240
         Width           =   645
      End
      Begin VB.TextBox txtInsAdd 
         Height          =   285
         Index           =   2
         Left            =   -70560
         TabIndex        =   117
         Top             =   5880
         Width           =   3495
      End
      Begin VB.TextBox txtInsAdd 
         Height          =   285
         Index           =   1
         Left            =   -70560
         TabIndex        =   116
         Top             =   5520
         Width           =   3495
      End
      Begin VB.Frame Frame7 
         Caption         =   "Insured Person(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -74520
         TabIndex        =   107
         Top             =   2160
         Width           =   8175
         Begin MSComctlLib.ListView lsvInsPax 
            Height          =   1215
            Left            =   240
            TabIndex        =   115
            Top             =   1560
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   2143
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Insured Persons"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Relationship"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Premium Amount"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton cmdAddInsPax 
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
            Left            =   4440
            TabIndex        =   114
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtInsPremiumAmt 
            Height          =   315
            Left            =   2040
            TabIndex        =   113
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtInsRelation 
            Height          =   285
            Left            =   2040
            TabIndex        =   112
            Top             =   720
            Width           =   3135
         End
         Begin VB.ComboBox cmbInsName 
            Height          =   315
            Left            =   2040
            TabIndex        =   111
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label14 
            Caption         =   "Premium Amount:"
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
            TabIndex        =   110
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Relationship:"
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
            TabIndex        =   109
            Top             =   720
            Width           =   1335
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
            Left            =   1200
            TabIndex        =   108
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtInsAdd 
         Height          =   285
         Index           =   0
         Left            =   -70560
         TabIndex        =   106
         Top             =   5160
         Width           =   3495
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
         TabIndex        =   92
         Top             =   600
         Width           =   10275
         Begin VB.ListBox lstEORmks 
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   97
            Top             =   300
            Width           =   4755
         End
         Begin VB.ListBox lstEORmks 
            Height          =   2205
            Index           =   1
            Left            =   5400
            MultiSelect     =   1  'Simple
            TabIndex        =   96
            Top             =   300
            Width           =   4755
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
            Picture         =   "frmOSMisc.frx":00A5
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Add All Remarks"
            Top             =   1080
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
            Picture         =   "frmOSMisc.frx":04E7
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "Add Selected Remark"
            Top             =   300
            Width           =   495
         End
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
            Picture         =   "frmOSMisc.frx":0929
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Remove Selected Remark"
            Top             =   1860
            Width           =   495
         End
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
         TabIndex        =   86
         Top             =   3840
         Width           =   10275
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
            Picture         =   "frmOSMisc.frx":0D6B
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Remove Selected Remark"
            Top             =   1860
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
            Picture         =   "frmOSMisc.frx":11AD
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Add Selected Remark"
            Top             =   300
            Width           =   495
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
            Picture         =   "frmOSMisc.frx":15EF
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   "Add All Remarks"
            Top             =   1080
            Width           =   495
         End
         Begin VB.ListBox lstItinRmks 
            Height          =   2205
            Index           =   1
            Left            =   5400
            MultiSelect     =   1  'Simple
            TabIndex        =   88
            Top             =   300
            Width           =   4755
         End
         Begin VB.ListBox lstItinRmks 
            Height          =   2205
            Index           =   0
            Left            =   60
            TabIndex        =   87
            Top             =   300
            Width           =   4755
         End
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
         Picture         =   "frmOSMisc.frx":1A31
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Add Free Text to Itinerary Remarks"
         Top             =   3255
         Width           =   495
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
         Picture         =   "frmOSMisc.frx":1E73
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Add Free Text to Exchange Order Remarks"
         Top             =   3255
         Width           =   495
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
         TabIndex        =   83
         Tag             =   "NN"
         Top             =   3375
         Width           =   6165
      End
      Begin VB.Frame fraHotel 
         Caption         =   "Additional Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   5160
         TabIndex        =   77
         Top             =   360
         Width           =   5355
         Begin VB.TextBox txtBTADescription 
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
            Left            =   2160
            MaxLength       =   20
            TabIndex        =   131
            Tag             =   "NN"
            Top             =   960
            Width           =   3075
         End
         Begin VB.TextBox txtDescription 
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
            Left            =   2160
            TabIndex        =   79
            Tag             =   "NN"
            Top             =   600
            Width           =   3075
         End
         Begin VB.TextBox txtDescription 
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
            Left            =   2160
            TabIndex        =   78
            Tag             =   "NN"
            Top             =   240
            Width           =   3075
         End
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   315
            Left            =   2160
            TabIndex        =   80
            Top             =   1320
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyyy"
            Format          =   16973827
            CurrentDate     =   38153
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "BTA Description:"
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
            TabIndex        =   132
            Top             =   960
            Width           =   1905
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
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
            TabIndex        =   82
            Top             =   1320
            Width           =   1185
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Description:"
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
            TabIndex        =   81
            Top             =   240
            Width           =   1185
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
         Left            =   120
         TabIndex        =   71
         Top             =   3840
         Width           =   5895
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
            ItemData        =   "frmOSMisc.frx":22B5
            Left            =   1800
            List            =   "frmOSMisc.frx":22D1
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   300
            Visible         =   0   'False
            Width           =   735
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
            ItemData        =   "frmOSMisc.frx":22F5
            Left            =   240
            List            =   "frmOSMisc.frx":2302
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   300
            Width           =   1515
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
            TabIndex        =   73
            Tag             =   "NN"
            Top             =   300
            Visible         =   0   'False
            Width           =   2025
         End
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
            TabIndex        =   72
            Top             =   720
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker dtpCCExp 
            Height          =   360
            Left            =   4680
            TabIndex        =   76
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
      Begin VB.Frame Frame1 
         Height          =   3015
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   4875
         Begin VB.ComboBox cboPlan 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   120
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton cmdCalCost 
            Caption         =   "Calculate  Cost"
            Height          =   255
            Left            =   3240
            TabIndex        =   121
            Top             =   600
            Visible         =   0   'False
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
            TabIndex        =   120
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtGrossSale 
            Height          =   315
            Left            =   1980
            TabIndex        =   63
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtSellPrice 
            Height          =   315
            Left            =   1980
            TabIndex        =   65
            Top             =   2580
            Width           =   1215
         End
         Begin VB.TextBox txtCost 
            Height          =   315
            Left            =   1980
            TabIndex        =   62
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtCommission 
            Height          =   315
            Left            =   1980
            TabIndex        =   61
            Top             =   1515
            Width           =   1215
         End
         Begin VB.TextBox txtMerchFee 
            Height          =   315
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   2220
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
            TabIndex        =   64
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtGST 
            Height          =   315
            Left            =   1980
            TabIndex        =   59
            Top             =   1860
            Visible         =   0   'False
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
            TabIndex        =   58
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Plan: "
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
            TabIndex        =   129
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
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
            TabIndex        =   119
            Top             =   840
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
            TabIndex        =   70
            Top             =   2580
            Width           =   1785
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
            TabIndex        =   69
            Top             =   480
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
            TabIndex        =   68
            Top             =   1560
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
            TabIndex        =   67
            Top             =   2220
            Width           =   1545
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
            TabIndex        =   66
            Top             =   1860
            Visible         =   0   'False
            Width           =   1545
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   6600
         TabIndex        =   52
         Top             =   4440
         Width           =   3915
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
            Left            =   180
            TabIndex        =   56
            Top             =   1260
            Width           =   1695
         End
         Begin VB.CommandButton cmdCancel 
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
            TabIndex        =   55
            Top             =   1260
            Width           =   1275
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
            Left            =   180
            Picture         =   "frmOSMisc.frx":2313
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtEONum 
            Height          =   315
            Left            =   1920
            MaxLength       =   13
            TabIndex        =   53
            Top             =   540
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   1395
         Left            =   5160
         Picture         =   "frmOSMisc.frx":2755
         ScaleHeight     =   1335
         ScaleWidth      =   1215
         TabIndex        =   51
         Top             =   2400
         Width           =   1275
      End
      Begin VB.TextBox txtContact 
         Height          =   285
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   50
         Top             =   435
         Width           =   2655
      End
      Begin VB.Frame Frame4 
         Height          =   3480
         Left            =   -74160
         TabIndex        =   35
         Top             =   960
         Width           =   4095
         Begin VB.ComboBox cboTrip 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   3360
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
            TabIndex        =   42
            Tag             =   "BY-"
            Top             =   2160
            Width           =   885
         End
         Begin VB.ComboBox cboClassServ 
            Enabled         =   0   'False
            Height          =   315
            Left            =   420
            TabIndex        =   40
            Text            =   "cboClassServ"
            Top             =   1800
            Width           =   3495
         End
         Begin VB.ComboBox cboTripType 
            Height          =   315
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   3000
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtMI 
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
            Index           =   3
            Left            =   2220
            MaxLength       =   3
            TabIndex        =   39
            Tag             =   "BY-"
            Top             =   1080
            Width           =   885
         End
         Begin VB.TextBox txtMI 
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
            Left            =   2220
            TabIndex        =   38
            Tag             =   "BY-"
            Top             =   180
            Width           =   1665
         End
         Begin VB.TextBox txtMI 
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
            Left            =   2220
            TabIndex        =   37
            Tag             =   "BY-"
            Top             =   600
            Width           =   1665
         End
         Begin VB.CheckBox chkPaperTkt 
            Caption         =   "ET"
            Enabled         =   0   'False
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
            TabIndex        =   36
            Top             =   2595
            Width           =   1095
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
            TabIndex        =   49
            Top             =   2160
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
            Left            =   3000
            TabIndex        =   48
            Top             =   3000
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
            Left            =   240
            TabIndex        =   47
            Top             =   1440
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
            TabIndex        =   46
            Top             =   1080
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
            TabIndex        =   45
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
            TabIndex        =   44
            Top             =   600
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
            TabIndex        =   43
            Top             =   2595
            Width           =   1515
         End
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
         Left            =   -73200
         TabIndex        =   34
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Frame fraTkt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   6600
         TabIndex        =   24
         Top             =   2400
         Visible         =   0   'False
         Width           =   3915
         Begin VB.OptionButton optConsol 
            Caption         =   "Consolidator Tkt"
            Height          =   255
            Left            =   1680
            TabIndex        =   118
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtFFNo 
            Height          =   315
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   30
            Top             =   1080
            Width           =   555
         End
         Begin VB.TextBox txtTktNo 
            Height          =   315
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   29
            Top             =   1080
            Width           =   1515
         End
         Begin VB.TextBox txtALCode 
            Height          =   315
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   28
            Top             =   1080
            Width           =   555
         End
         Begin VB.OptionButton optNormal 
            Caption         =   "BSP (with File Fare)"
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton optBSPConsol 
            Caption         =   "BSP (without File Fare)"
            Height          =   255
            Left            =   1680
            TabIndex        =   26
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtPassengerID 
            Height          =   315
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   25
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label lblTktNo 
            Alignment       =   1  'Right Justify
            Caption         =   "Ticket No:"
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
            TabIndex        =   33
            Top             =   1080
            Width           =   1425
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Ticket Type: "
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
            TabIndex        =   32
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label lblPassengerID 
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
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   1425
         End
      End
      Begin VB.Frame fraVendorInfo 
         Height          =   5535
         Left            =   -74640
         TabIndex        =   6
         Top             =   600
         Width           =   9975
         Begin VB.TextBox txtReplyEmail 
            Height          =   375
            Left            =   240
            TabIndex        =   133
            Top             =   3960
            Width           =   8655
         End
         Begin VB.TextBox txtTel 
            Height          =   375
            Left            =   960
            TabIndex        =   15
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox txtCreditTerms 
            Height          =   420
            Left            =   4200
            TabIndex        =   14
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox txtVendor 
            Height          =   375
            Left            =   960
            TabIndex        =   13
            Top             =   360
            Width           =   7935
         End
         Begin VB.TextBox txtAddress1 
            Height          =   375
            Left            =   960
            TabIndex        =   12
            Top             =   840
            Width           =   7935
         End
         Begin VB.TextBox txtAddress2 
            Height          =   375
            Left            =   960
            TabIndex        =   11
            Top             =   1320
            Width           =   7935
         End
         Begin VB.TextBox txtCity 
            Height          =   375
            Left            =   960
            TabIndex        =   10
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtCountry1 
            Height          =   375
            Left            =   3840
            TabIndex        =   9
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   960
            TabIndex        =   8
            Top             =   2280
            Width           =   7935
         End
         Begin VB.TextBox txtFaxNo 
            Height          =   375
            Left            =   960
            TabIndex        =   7
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "Reply Email in EO (Only 1 email address is allowed)"
            Height          =   375
            Left            =   240
            TabIndex        =   134
            Top             =   3720
            Width           =   4095
         End
         Begin VB.Label Label11 
            Caption         =   "Contact No."
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Credit Terms"
            Height          =   255
            Left            =   3000
            TabIndex        =   22
            Top             =   2760
            Width           =   1065
         End
         Begin VB.Label Label15 
            Caption         =   "Vendor "
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "Address"
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "City"
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Country"
            Height          =   375
            Left            =   3120
            TabIndex        =   18
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label19 
            Caption         =   "Email (;)"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Fax No (,)"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   2760
            Width           =   735
         End
      End
      Begin VB.ComboBox cmbGeoArea 
         Height          =   315
         Left            =   -72600
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cmbInsPlan 
         Height          =   315
         Left            =   -72600
         TabIndex        =   4
         Text            =   "cmbInsPlan"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtInsDays 
         Height          =   315
         Left            =   -72600
         TabIndex        =   3
         Top             =   1320
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpInsFromDate 
         Height          =   315
         Left            =   -69480
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16973825
         CurrentDate     =   38506
      End
      Begin MSComctlLib.ListView lvwECodes 
         Height          =   4275
         Left            =   -69480
         TabIndex        =   98
         Top             =   5760
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
         Enabled         =   0   'False
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
      Begin MSComctlLib.ListView lvwMissECodes 
         Height          =   1875
         Left            =   -69720
         TabIndex        =   125
         Top             =   3600
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
         Enabled         =   0   'False
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
      Begin MSComctlLib.ListView lvwRealECodes 
         Height          =   1815
         Left            =   -69720
         TabIndex        =   128
         Top             =   1320
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
         Enabled         =   0   'False
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
         TabIndex        =   127
         Top             =   960
         Width           =   2235
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
         Left            =   -69480
         TabIndex        =   126
         Top             =   3240
         Width           =   1995
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Address of First Name Insured Person:"
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
         Left            =   -74640
         TabIndex        =   105
         Top             =   5160
         Width           =   3825
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
         TabIndex        =   104
         Top             =   3435
         Width           =   1545
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
         TabIndex        =   103
         Top             =   435
         Width           =   2055
      End
      Begin VB.Label lblGeoArea 
         Alignment       =   1  'Right Justify
         Caption         =   "Geographical Area:"
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
         Left            =   -74640
         TabIndex        =   102
         Top             =   840
         Width           =   1905
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of Days:"
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
         Left            =   -74640
         TabIndex        =   101
         Top             =   1320
         Width           =   1905
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "From Date:"
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
         Left            =   -70800
         TabIndex        =   100
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Plan Selected:"
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
         Left            =   -74520
         TabIndex        =   99
         Top             =   1800
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmOSMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsRmks As ADODB.Recordset
Dim mobjEO As EO
Dim sngGST As Single
'Dim StartTime As Date
Dim mstrPCAmend As String
Dim mstrTktNum As String
Dim blnMI As Boolean
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

Private Sub chkPaperTkt_Click()
With chkPaperTkt
    If .value = vbChecked Then
        .Caption = "PT"
    Else
        .Caption = "ET"
    End If
End With
End Sub

Private Sub cmbFOPType_Click()
Dim blnCC As Boolean
    
EnableCalculate
'blnCC = (cmbFOPType = "CX")
'cmbCCType.Visible = blnCC
'txtCCNum.Visible = blnCC
'dtpCCExp.Visible = blnCC
'chkWaiveMercFee.Visible = blnCC

'MODIFIED 30062006
If cmbFOPType = "CC" Or cmbFOPType = "CX" Then
    cmbCCType.Visible = True
    txtCCNum.Visible = True
    dtpCCExp.Visible = True
Else
    cmbCCType.Visible = False
    txtCCNum.Visible = False
    dtpCCExp.Visible = False
End If

End Sub



Private Sub cmbInsName_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowAlpha(KeyAscii, "/ ")
End Sub

Private Sub cmdAddInsPax_Click()
Dim strError As String
Dim item As ListItem

   If cmbInsName.Text = "" Then strError = "Need Insured Person Name"
   If txtInsPremiumAmt = "" Then strError = "Need Premium Amount"
   
   If strError <> "" Then
        'MsgBox strError
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strError, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        Exit Sub
   End If
   
   Set item = lsvInsPax.ListItems.Add(, , cmbInsName.Text)
   If txtInsRelation.Text <> "" Then item.SubItems(1) = txtInsRelation.Text
   item.SubItems(2) = Format(txtInsPremiumAmt.Text, "0.00")
   
End Sub

Private Sub cmdCalCost_Click()
Dim sngInsPct As Single
Dim strMsg As String

sngInsPct = 0
If txtGrossSale = "" Then
    'MsgBox "Please input Selling Price"
    strMsg = "Please input Selling Price"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    Exit Sub
End If
If cboPlan.ListCount > 0 Then
   sngInsPct = 1 - (cboPlan.ItemData(cboPlan.listindex) / 100)
End If
txtCost = Format((txtGrossSale * sngInsPct), gstrAgcyCurrFormat)

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
 
'If chkAbsorb.Value = 0 Then
'   If Me.cmbFOPType.Text = "CX" And Me.chkWaiveMercFee.Value = vbUnchecked Then
'       Me.txtMerchFee = fMerchantFee(sngSF, fGetMerchFee(gstrAgcyCountryCode))
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
        If frmOthSvcs.dbcProducts.BoundText = 35 Or _
        frmOthSvcs.dbcProducts.BoundText = 41 Or _
        frmOthSvcs.dbcProducts.BoundText = 70 Or _
        frmOthSvcs.dbcProducts.BoundText = 50 Then
            frmClientMI.intLocation = 4
        Else
            frmClientMI.intLocation = 6
        End If
        frmClientMI.intProdCode = frmOthSvcs.dbcProducts.BoundText
        frmClientMI.cmbMICat.Enabled = False
        frmClientMI.pGetClientMI (gobjPNR.CN)
        
 '230108
        frmClientMI.strPdtType = frmOthSvcs.datProducts.Recordset![Type]
        frmClientMI.Show 'vbModal
    End If
End Sub
Private Sub cmdDone_Click()
    Dim freefields As String
    Dim strMsg As String
    datTouchEnd = Now
    If Not validData Then Exit Sub
    cmdDone.Enabled = False
    
    gSysStartOthSvcsTime = Now
    If gobjEO Is Nothing Or txtEONum = "" Then Call SetEOObj(gobjEO)
    
    freefields = ""
    'If frmOthSvcs.dbcProducts.BoundText = "35" Or frmOthSvcs.dbcProducts.BoundText = "50" Then
    '230108
    'If blnMI = True Then
  
        If gobjEO.FF7 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "7-" & gobjEO.FF7
        If gobjEO.FF8 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "8-" & gobjEO.FF8
        If gobjEO.FF81 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "81-" & gobjEO.FF81
        ''CS Change EC
        'If gobjEO.rs <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "30-" & gobjEO.rs
        'CS Remove FF26
        'If gobjEO.FF26 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "26-" & gobjEO.FF26
        ''CS Add FF41
        'If gobjEO.FF41 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "41-" & gobjEO.FF41
        If gobjEO.FF38 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "38-" & gobjEO.FF38
   
    
    If frmClientMI.MSXfreefields <> "" Then
        freefields = freefields & "/" & frmClientMI.MSXfreefields
    End If
    
    'Check for completion of client MI
    'If Not isCompleteClientMI(gobjPNR.CN, 6, , frmOthSvcs.dbcProducts.BoundText) Then
  
    If frmClientMI.MSXfreefields = "" Then
        If frmOthSvcs.dbcProducts.BoundText = "35" Or _
            frmOthSvcs.dbcProducts.BoundText = "41" Or _
            frmOthSvcs.dbcProducts.BoundText = "50" Or _
            frmOthSvcs.dbcProducts.BoundText = "70" Then
            If isRequireClientMI(gobjPNR.CN, 4) Then
                cmdDone.Enabled = True
                'MsgBox "Client MI data is incomplete", vbCritical
                strMsg = "Client MI data is incomplete"
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
                loadClientMI
                Exit Sub
            End If
            
        Else
            If isRequireClientMI(gobjPNR.CN, 6) Then
                cmdDone.Enabled = True
                'MsgBox "Client MI data is incomplete", vbCritical
                strMsg = "Client MI data is incomplete"
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
                loadClientMI
                Exit Sub
            End If
        End If
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
    Else
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
    End If
        
End Sub

Private Sub cmdEO_Click()
Dim lngC As Long
Dim strMsg As String
datTouchEnd = Now
If Not validData Then Exit Sub
gSysStartOthSvcsTime = Now
If isRequireClientMI(gobjPNR.CN, 6) And frmClientMI.MSXfreefields = "" Then 'And blnMI = True Then
        'cmdDone.Enabled = True
        'MsgBox "Client MI data is incomplete", vbCritical
        strMsg = "Client MI data is incomplete"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        loadClientMI
        Exit Sub
End If

Call SetEOObj(gobjEO)
Log
'Added on 8/3/2005: To end PNR to get RecLoc
'Call modOthSvcs.SetEONumber

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
   ' .RemoveItem .ListIndex
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
'   frmOthSvcs.datSelectedVendor.DatabaseName = gstrTProDBSource
'   frmOthSvcs.datSelectedVendor.RecordSource = "SELECT * FROM tblVendors WHERE [VendorNumber] =  '" & dbcVendors.BoundText & "'"
'   frmOthSvcs.datSelectedVendor.Refresh
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
Dim rsRemarks As ADODB.Recordset
Dim rsECodes As ADODB.Recordset
Dim rsInsurance As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim item As ListItem
Dim strSQL As String
Dim blnChkDate As Boolean
Dim blnChkVendor As Boolean
Dim strTemp As String
Dim intPaxCount As Integer
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
SSTab1.Tab = 0
dtpDate.value = Date
'Timer
gStartOthSvcsTime = Now
frmOSMisc.Caption = "CWT TravelPro - " & frmOthSvcs.dbcProducts.Text
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

txtDescription(0).Text = frmOthSvcs.dbcProducts.Text
'Added on 20/08/2007: Auto populate BTA Description
txtBTADescription.Text = frmOthSvcs.dbcProducts.Text

'added on 30/5/2005: Due line descroption can only accept up to 35 chars of free text format
txtDescription(0).MaxLength = 35

With rsRemarks
    Do Until .EOF
        If ![RmkType] & "" = "I" Then
            lstItinRmks(0).AddItem ![Text]
        Else
            lstEORmks(0).AddItem ![Text]
        End If
        .MoveNext
    Loop
    .Close
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
    For lngC = 1 To .AirSegCount
        lstFlights.AddItem .AirSeg(lngC).TextAirSeg
    Next
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


'230108

'SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False

Set rs = gdbConn.Execute("select MI,TktNo from tblproductcodes where productcode=" & frmOthSvcs.dbcProducts.BoundText)

If Not rs.EOF Then
     '230108
    'If rs!MI = True Then
    'SSTab1.TabEnabled(2) = rs!MI
    'blnMI = rs!MI
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
        
        'CS Remove FF26
        'cboTripType.AddItem ""
        'cboTripType.AddItem "Round"
        'cboTripType.AddItem "One Way"
        
        'CS Add FF41
        'cboTrip.AddItem "INTERNATIONAL"
        'cboTrip.AddItem "DOMESTIC"
        'cboTrip.listindex = -1
        'cboTrip.Enabled = False
   
        'Modified on 18/2/2005
        'Modified on 2/2/2005: add on client specific EC
        'CS Change EC
        'strSql = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='AS' OR tblExceptionCodes.ExceptionCodeGroup='AC') ORDER BY CAST(tblClientEC.EC AS integer)"
        'CS Change EC
        'strSQL = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND tblExceptionCodes.ProdType='" & "AIR" & "' AND tblExceptionCodes.ECInd='S' ORDER BY tblClientEC.EC"
        strSQL = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.ClientID & " AND tblExceptionCodes.ProdType='" & "AIR" & "' ORDER BY tblClientEC.EC"

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
        
'        'strSql = "SELECT distinct(CAST(tblClientEC.EC AS integer)),description,exceptioncodegroup FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC and tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='AS' OR tblExceptionCodes.ExceptionCodeGroup='AC') ORDER BY CAST(tblClientEC.EC AS integer)"
'        'Set rsECodes = gdbTPro.OpenRecordset(strSQL)
'        Set rsECodes = gdbConn.Execute(strSql)
'        If Not rsECodes.EOF Then'
'
'             rsECodes.MoveFirst
'             Do While Not rsECodes.EOF
'                Set Item = lvwECodes.ListItems.Add(, , rsECodes!EC)
'                      If rsECodes!Remarks = "" Then
'                       Item.SubItems(1) = rsECodes!Description
'                      Else
'                       Item.SubItems(1) = rsECodes!Remarks
'                      End If
'                rsECodes.MoveNext
'              Loop
'           rsECodes.Close
'
'        Else
'
'            rsECodes.Close
'            strSql = "SELECT * FROM tblExceptionCodes where ExceptionCodeGroup='AC' order by CAST(ExceptionCode AS integer) "
'            'Set rsECodes = gdbTPro.OpenRecordset(strSQL)
'            Set rsECodes = gdbConn.Execute(strSql)
'               If Not rsECodes.EOF Then rsECodes.MoveFirst
'
'               Do While Not rsECodes.EOF
'                  Set Item = lvwECodes.ListItems.Add(, , rsECodes!exceptioncode)
'                  Item.SubItems(1) = rsECodes!Description
'                  rsECodes.MoveNext
'                Loop
'              rsECodes.Close
'
'    End If
'
'    Set rsECodes = Nothing
'End If
      
If rs!TktNo = True Then
        fraTkt.Visible = rs!TktNo
        lblTktNo.Visible = False
        txtFFNo.Visible = False
        txtALCode.Visible = False
        txtTktNo.Visible = False
        lblPassengerID.Visible = False
        txtPassengerID.Visible = False
    End If
End If
rs.Close
Set rs = Nothing

Select Case frmOthSvcs.dbcProducts.BoundText
    'Case "35", "50"
        'Enable MI Tab
    '    SSTab1.TabEnabled(2) = True
        'Added on 20/4/2005: Get tkt/ff information
    '    fraTkt.Visible = True
        
    '    lblTktNo.Visible = False
    '    txtFFNo.Visible = False
    '    txtALCode.Visible = False
    '    txtTktNo.Visible = False
    '    lblPassengerID.Visible = False
    '    txtPassengerID.Visible = False
        
        
        'Added on 12/01/05
        'populate combo box values for Class of Services, Trip Type
    '    cboClassServ.AddItem ""
    '    cboClassServ.AddItem "FF"   'First Class - Full Fare
    '    cboClassServ.AddItem "FD"   'First Class - Discounted Fare
    '    cboClassServ.AddItem "FN"   'First Class - Nett Fare
    '    cboClassServ.AddItem "CF"   'Business Class - Full Fare
    '    cboClassServ.AddItem "CD"
    '    cboClassServ.AddItem "CN"
    '    cboClassServ.AddItem "YF"   'Economy Class - Full Fare
    '    cboClassServ.AddItem "YD"
    '    cboClassServ.AddItem "YN"
        
    '    cboTripType.AddItem ""
    '    cboTripType.AddItem "Round"
    '    cboTripType.AddItem "One Way"
        
    '    If UCase(gstrAgcyCountryCode) = "SG" Then
    '       If frmOthSvcs.dbcProducts.BoundText = "50" Then
    '          cmbFOPType.Text = "INV"
    '          cmbFOPType.Enabled = False
    '       End If
    '    End If
'Modified on 18/2/2005
'Modified on 2/2/2005: add on client specific EC
'strSql = "SELECT * FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC where tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='AS' OR tblExceptionCodes.ExceptionCodeGroup='AC') ORDER BY CAST(tblClientEC.EC AS integer)"
'strSql = "SELECT distinct(CAST(tblClientEC.EC AS integer)),description,exceptioncodegroup FROM tblExceptionCodes INNER JOIN tblClientEC ON tblExceptionCodes.ExceptionCode = tblClientEC.EC and tblClientEC.ClientID=" & gobjPNR.CompInfo.clientId & " AND (tblExceptionCodes.ExceptionCodeGroup='AS' OR tblExceptionCodes.ExceptionCodeGroup='AC') ORDER BY CAST(tblClientEC.EC AS integer)"
'Set rsECodes = gdbTPro.OpenRecordset(strSQL)
'Set rsECodes = gdbConn.Execute(strSql)
'If Not rsECodes.EOF Then

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

'Else
   
'        rsECodes.Close
'        strSql = "SELECT * FROM tblExceptionCodes where ExceptionCodeGroup='AC' order by CAST(ExceptionCode AS integer) "
'        'Set rsECodes = gdbTPro.OpenRecordset(strSQL)
'        Set rsECodes = gdbConn.Execute(strSql)
'           If Not rsECodes.EOF Then rsECodes.MoveFirst
           
'           Do While Not rsECodes.EOF
'              Set Item = lvwECodes.ListItems.Add(, , rsECodes!exceptioncode)
'              Item.SubItems(1) = rsECodes!Description
'              rsECodes.MoveNext
'            Loop
'          rsECodes.Close
     
'End If

'Set rsECodes = Nothing

Case "09"
        'lblGeoArea.Visible = True
        'cmbGeoArea.Visible = True
        SSTab1.TabEnabled(3) = True
        
        cmbGeoArea.AddItem ""
        cmbGeoArea.AddItem "Asean"
        cmbGeoArea.AddItem "Asia/AU/NZ"
        cmbGeoArea.AddItem "Worldwide exc US/CAN"
        cmbGeoArea.AddItem "Worldwide inc US/CAN"
        
        cmbInsPlan.AddItem "Individual Plan - Single Trip"
        cmbInsPlan.AddItem "Individual Plan - Annual"
        'cmbInsPlan.AddItem "Individual Plan"
        cmbInsPlan.AddItem "Family Budget Plan"
        'cmbInsPlan.AddItem "Annual Policy Plan"
        cmbInsPlan.listindex = 0
        
        dtpInsFromDate.value = Date
        
        intPaxCount = gobjPNR.PassengerCount
        If intPaxCount > 0 Then
            For lngC = 1 To intPaxCount
                With gobjPNR.PassengerName(lngC)
                    strTemp = .LastName & "/" & .FirstName
                End With
               cmbInsName.AddItem strTemp
            Next
        End If
        
        If cmbInsName.ListCount > 0 Then cmbInsName.listindex = 0
        If UCase(gstrAgcyCountryCode) = "SG" Then
            Label9.Visible = True
            Set rsInsurance = gdbConn.Execute("Select * FROM tblInsurance")
            Do While Not rsInsurance.EOF
                cboPlan.AddItem rsInsurance!Type
                cboPlan.ItemData(cboPlan.NewIndex) = rsInsurance!Commission
                rsInsurance.MoveNext
            Loop
            rsInsurance.Close
            Set rsInsurance = Nothing
            If cboPlan.ListCount > 0 Then cboPlan.listindex = 0
            cboPlan.Visible = True
            cmdCalCost.Visible = True
        End If
Case "70"
        If UCase(gstrAgcyCountryCode) = "SG" Then
           cmbFOPType.Text = "INV"
           cmbFOPType.Enabled = False
           txtCost = 0
           txtCost.Enabled = False
        End If
Case "50"
        If UCase(gstrAgcyCountryCode) = "SG" Then
              cmbFOPType.Text = "INV"
              cmbFOPType.Enabled = False
        End If
        fraAir.Visible = True
Case "35"
        fraAir.Visible = True
End Select

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
txtReplyEmail.Text = GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjPNR.Agent, gobjPNR.PCCOwner, False, True)
'230108
'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
'frmClientMI.bolCheck = False

'300908 Detect MI by CN
If gobjPNR.CompInfo.MI = False Then
    
    lblLabels(28).Enabled = False
    lblLabels(27).Enabled = False
    lblLabels(30).Enabled = False
    lblLabels(36).Enabled = False
    lblLabels(32).Enabled = False
    lblLabels(26).Enabled = False
    lblLabels(41).Enabled = False
    lblLabels(40).Enabled = False
    
    txtMI(0).Enabled = False
    txtMI(1).Enabled = False
    txtMI(3).Enabled = False
    txtMI(6).Enabled = False
    txtMI(0).Enabled = False
    chkPaperTkt.Enabled = False
    
    txtRS.Enabled = False
    txtMS.Enabled = False
    lvwRealECodes.Enabled = False
    lvwMissECodes.Enabled = False
    
    cboClassServ.Enabled = False
 
    
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
   Dim strMI() As String
   Dim strPax() As String
   Dim strAdd() As String
   Dim strPaxInfo() As String
   Dim item As ListItem
   
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
     
     strTemp = Split(!AdditionalInfo, vbCrLf)
     For i = 0 To UBound(strTemp)
        If i = 10 Then Exit For
        Select Case i
           Case 2 'date
              If IsDate(strTemp(i)) Then
                 dtpDate.value = strTemp(i)
              End If
              'Exit For
           Case 3 'geographical area for insurance
              If strTemp(i) <> "" Then
              cmbGeoArea.Text = strTemp(i)
              End If
           Case 4
              If strTemp(i) <> "" Then
              txtInsDays = strTemp(i)
              End If
          Case 5
              If strTemp(i) <> "" Then
              dtpInsFromDate.value = strTemp(i)
              End If
           Case 6
              If strTemp(i) <> "" Then
              cmbInsPlan.Text = strTemp(i)
              End If
           Case 7
              If strTemp(i) <> "" Then
              strPax = Split(strTemp(i), ",")
              For j = 0 To UBound(strPax)
                strPaxInfo = Split(strPax(j), ";")
                Set item = lsvInsPax.ListItems.Add(, , strPaxInfo(0))
                If strPaxInfo(1) <> "" Then item.SubItems(1) = strPaxInfo(1)
                item.SubItems(2) = strPaxInfo(2)
              Next
              End If
            Case 8
              If strTemp(i) <> "" Then
                strAdd = Split(strTemp(i), ";")
                For j = 0 To UBound(strAdd)
                  If strAdd(j) <> "" Then txtInsAdd(j) = strAdd(j)
                Next
              End If
            Case 9
              If strTemp(i) <> "" Then
                txtBTADescription = strTemp(i)
              End If
           Case Else 'desc text
              txtDescription(i) = strTemp(i)
        End Select
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
        If strMI(4) <> "" Then
           cboClassServ.Text = matchList(strMI(4))
        End If
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
        'Else
        '   cboTrip.listindex = -1
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
      '  ElseIf strMI(9) = "D" Then
      '     cboTrip.Text = "DOMESTIC"
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
     
     'added on 20/4/2005: retrieve ticket no and default to BSP/Consol ticket
     If !TktNum <> "" And Len(!TktNum) = 13 Then
        optBSPConsol.value = True
        txtALCode = Left(!TktNum, 3)
        txtTktNo = Mid(!TktNum, 4)
     Else
        optNormal.value = True
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

Private Function validData() As Boolean
Dim strMsg As String
Dim i As Integer

If gbolEOAmend = False Then
   If frmOthSvcs.dbcVendors.Text = "" Then strMsg = strMsg & "Select vendor for this transaction..." & Chr(13)
End If
If txtDescription(0).Text = "" Then strMsg = strMsg & "Need Desctiption..." & Chr(13)
If Me.dtpDate < Date Then strMsg = strMsg & "Service date cannot be past..." & Chr(13)
If txtSellPrice.Text = "" Then strMsg = strMsg & "Need to calculate Selling Price..." & Chr(13)
If cmbFOPType.Text = "" Then strMsg = strMsg & "Need form of payment..." & Chr(13)
If cmbFOPType.Text = "CX" Or cmbFOPType.Text = "CC" Then
    If cmbCCType.Text = "" Then strMsg = strMsg & "Need valid credit vendor code..." & Chr(13)
    If txtCCNum = "" Then strMsg = strMsg & "Need valid credit card number..." & Chr(13)
    If dtpCCExp.value < Date Then strMsg = strMsg & "Need valid expiration date..." & Chr(13)
    If (txtCCNum.Text <> "" And cmbCCType.Text <> "") Then If ValidCCNum(cmbCCType.Text, txtCCNum.Text) = False Then strMsg = strMsg & "Credit card number is invalid or wrong card vendor selected ..." & Chr(13)
End If

If (frmOthSvcs.dbcProducts.BoundText = "50" Or frmOthSvcs.dbcProducts.BoundText = "35") And lstFlights.SelCount = 0 Then
    strMsg = strMsg & "Need to select air segment(s) for this transaction..."
End If

If Trim(Len(txtBTADescription)) = 0 And cmbCCType.Text = "AX" Then strMsg = strMsg & "Missing BTA Description..." & Chr(13)


    'If blnMI = True Then
    '    If txtMI(3).Text = "" Then strMsg = strMsg & "Need Final Destination (MI)..." & Chr(13)
    'End If
    
    If fraTkt.Visible = True Then
        If optConsol.value = False And optBSPConsol.value = False And optNormal.value = False Then
            strMsg = strMsg & "Need to select Ticket Type..." & Chr(13)
        End If
        
        If optConsol.value = True And ((Len(txtTktNo) <> 10 And Len(txtALCode) <> 3) Or IsNumeric(txtTktNo) = False) Then
           strMsg = strMsg & "Invalid EO number..." & Chr(13)
        End If
        If optBSPConsol.value = True Then
           If Len(txtALCode) <> 3 Or IsNumeric(txtALCode) = False Or Len(txtTktNo) <> 10 Or IsNumeric(txtTktNo) = False Then
              strMsg = strMsg & "Invalid Ticket number..." & Chr(13)
           End If
        End If
        If optNormal.value = True Then
            If txtFFNo = "" Then
              strMsg = strMsg & "Need File fare number..." & Chr(13)
           End If
        End If
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
Select Case frmOthSvcs.dbcProducts.BoundText
'    Case "35", "50"
        'If txtMI(0).Text = "" Then strMsg = strMsg & "Need Reference Fare (MI)..." & Chr(13)
        'If txtMI(1).Text = "" Then strMsg = strMsg & "Need Low Fare (MI)..." & Chr(13)
        'If txtMI(2).Text = "" Then strMsg = strMsg & "Need Exception Code (MI)..." & Chr(13)
        'If blnMI = True Then
        '    If txtMI(3).Text = "" Then strMsg = strMsg & "Need Final Desitnation (MI)..." & Chr(13)
        'End If
        
        'If fraTkt.Visible = True Then
        '    If optConsol.Value = True And ((Len(txtTktNo) <> 10 And Len(txtALCode) <> 3) Or IsNumeric(txtTktNo) = False) Then
        '       strMsg = strMsg & "Invalid EO number..." & Chr(13)
        '    End If
        '    If optBSPConsol.Value = True Then
        '       If Len(txtALCode) <> 3 Or IsNumeric(txtALCode) = False Or Len(txtTktNo) <> 10 Or IsNumeric(txtTktNo) = False Then
        '          strMsg = strMsg & "Invalid Ticket number..." & Chr(13)
        '       End If
        '    End If
        'End If
        
    'If optNormal.Value Then
        'If txtFFNo = "" Then strMsg = strMsg & "Need file fare no..." & Chr(13)
    'Else
        'If txtALCode = "" Then strMsg = strMsg & "Need airline code..." & Chr(13)
        'If txtTktNo = "" Then strMsg = strMsg & "Need Ticket No..." & Chr(13)
    'End If
    
    Case "09"
        If UCase(gstrAgcyCountryCode) = "SG" Then
            If cmbGeoArea.Text = "" Then strMsg = strMsg & "Need Geographical Area(Insurance)..." & Chr(13)
            If txtInsDays = "" Then strMsg = strMsg & "Need Insurance Period/Days(Insurance)..." & Chr(13)
            If dtpInsFromDate.value < Date Then strMsg = strMsg & "Insurance date cannot be past(Insurance)..." & Chr(13)
            If txtInsAdd(0) = "" Then strMsg = strMsg & "Need Insured Address(Insurance)..." & Chr(13)
            If lsvInsPax.ListItems.Count = 0 Then strMsg = strMsg & "Need Insured Person(Insurance)..." & Chr(13)
        End If
End Select

For i = 0 To 1
   If InStr(1, UCase(txtDescription(i).Text), "/PC") <> 0 Then
      strMsg = strMsg & "'/PC' exist in description text box.." & Chr(13)
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
    'cmdDone.Enabled = True
    'If gbolEOAmend = False Then cmdDone.Enabled = True
    'MsgBox strMsg, vbApplicationModal + vbExclamation, "Travel Pro"
     modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
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

Private Sub lsvInsPax_DblClick()
    lsvInsPax.ListItems.Remove (lsvInsPax.SelectedItem.Index)
End Sub

Private Sub lsvInsPax_ItemClick(ByVal item As MSComctlLib.ListItem)
cmbInsName.Text = item.Text
txtInsRelation.Text = item.ListSubItems(1)
txtInsPremiumAmt.Text = item.ListSubItems(2)
End Sub

'Private Sub lvwECodes_DblClick()
'   txtMI(2).Text = lvwECodes.SelectedItem
'End Sub

Private Sub lvwMissECodes_ItemClick(ByVal item As MSComctlLib.ListItem)
txtMS = lvwMissECodes.SelectedItem
End Sub

Private Sub lvwRealECodes_ItemClick(ByVal item As MSComctlLib.ListItem)
txtRS = lvwRealECodes.SelectedItem
End Sub

Private Sub optBSPConsol_Click()
        lblTktNo.Caption = "Ticket No:"
        lblTktNo.Visible = True
        txtFFNo.Visible = False
        txtALCode.Visible = True
        txtTktNo.Visible = True
        lblPassengerID.Visible = True
        txtPassengerID.Visible = True
        txtEONum.Locked = True
End Sub

Private Sub optConsol_Click()
   lblTktNo.Caption = "EO No:"
   lblTktNo.Visible = True
   txtFFNo.Visible = False
   txtALCode.Visible = True
   txtTktNo.Visible = True
   lblPassengerID.Visible = True
   txtPassengerID.Visible = True
   txtEONum.Locked = False
End Sub

Private Sub optNormal_Click()
        lblTktNo.Caption = "File Fare No:"
        lblTktNo.Visible = True
        txtFFNo.Visible = True
        txtALCode.Visible = False
        txtTktNo.Visible = False
        lblPassengerID.Visible = True
        txtPassengerID.Visible = True
        txtEONum.Locked = True
End Sub



Private Sub txtALCode_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii)
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

Private Sub EnableCalculate()
   cmdCalculate.Enabled = True
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub

Private Sub txtDescription_GotFocus(Index As Integer)
Call pSetSelected
End Sub

Private Sub txtDescription_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = fAllowAlphaNumeric(KeyAscii, "/.- ")
End Sub



Private Sub txtDisplayNo_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii)
End Sub



Private Sub txtFFNo_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii)
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

Private Sub txtInsDays_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii)
End Sub

Private Sub txtInsPremiumAmt_KeyPress(KeyAscii As Integer)
 KeyAscii = fAllowNumeric(KeyAscii, ".")
End Sub



Private Sub txtInsRelation_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii)
End Sub

Private Sub txtMerchFee_GotFocus()
Call pSetSelected
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

Private Sub txtPassengerID_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii)
End Sub

Private Sub txtSellPrice_GotFocus()
Call pSetSelected
End Sub

Private Sub SetEOObj(ByRef objEO As EO)
Dim lngC As Long
Dim strPaxName As String
Dim strInsPax As String
Dim strInsAdd As String

'Set objEO = New EO
'txtEONum = ""
If objEO Is Nothing Then
   Set objEO = New EO
Else
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
    ElseIf optConsol.value = True Then
       .EONumber = txtEONum
    End If

    .BillingDescription = ""
    .CN = gobjPNR.CN
    .CommissionAmt = fConvertZero(txtCommission.Text) '+ fConvertZero(txtMerchFee)
    .Cost = fConvertZero(txtCost.Text)
    .CreatedBy = gobjHost.AgentSine
    '.CreatedByName = gobjHost.AgentName
    .CreatedByName = gobjHost.AgentProfile
    .CreatedByPCC = gobjHost.AgentPCC
    .CreateDtTm = Now()
    .DescriptionLineAdd txtDescription(0).Text
    If txtDescription(1).Text <> "" Then .DescriptionLineAdd txtDescription(1).Text
    '.FOP = cmbFOPType.Text & IIf(cmbFOPType.Text = "CX", "/" & cmbCCType.Text & "/" & txtCCNum.Text & "/" & Format(dtpCCExp.Value, "ddmm"), "")
    .FOP = cmbFOPType.Text & IIf(cmbFOPType.Text = "CX" Or cmbFOPType.Text = "CC", "/" & cmbCCType.Text & "/" & txtCCNum.Text & "/" & Format(dtpCCExp.value, "MMYY"), "") & "/" & "0" & "/" & chkWaiveMercFee.value
    
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
    If UCase(gstrAgcyCountryCode) = "SG" Then
        '.TaxAdd fConvertZero(txtGST.Text), "GST"
        .TaxAdd sngGST, "GST"
        .NettGST = sngNettCostGST
    Else
        .TaxAdd 0, ""
        .NettGST = 0
        
    End If

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
    '.VendorCode = frmOthSvcs.dbcVendors.BoundText
    '.Email = frmOthSvcs.datVendors.Recordset![Email] & ""
    '.FaxNo = frmOthSvcs.datVendors.Recordset![FaxNumber] & ""
    
    For lngC = 0 To lstEORmks(1).ListCount - 1
        .RemarkAdd lstEORmks(1).List(lngC)
    Next
    For lngC = 0 To lstItinRmks(1).ListCount - 1
        .RIRemarkAdd lstItinRmks(1).List(lngC)
    Next
    .TktType = fraTkt.Visible
    .RF = txtMI(0)
    .LF = txtMI(1)
    'CS Change EC
    '.EC = txtMI(2)
    .rs = txtRS
    .MS = txtMS
    .FF7 = txtMI(3)
    .FF8 = Trim(Left(cboClassServ.Text, 2))
    'CS Remove FF26
    '.FF26 = IIf(UCase(cboTripType.Text) = "ROUND", "R", IIf(UCase(cboTripType.Text) = "ONE WAY", "O", ""))
    'CS Add FF41
    '.FF41 = IIf(UCase(cboTrip.Text) = "INTERNATIONAL", "I", IIf(UCase(cboTrip.Text) = "DOMESTIC", "D", ""))
    .FF81 = txtMI(6)
    If chkPaperTkt.Enabled Then .FF38 = chkPaperTkt.Caption
    'FF10,11 will be maintained by client related MI
    '.FF10 = txtMI(4)
    '.FF11 = txtMI(5)
    .MerchFee = fConvertZero(txtMerchFee)
    .CWTAbsorb = IIf(chkAbsorb.value = 1, True, False)
    strInsPax = ""
    strInsAdd = ""
    For lngC = 1 To lsvInsPax.ListItems.Count
        strInsPax = strInsPax & lsvInsPax.ListItems(lngC) & ";" & lsvInsPax.ListItems(lngC).SubItems(1) & ";" & lsvInsPax.ListItems(lngC).SubItems(2) & IIf(lngC = lsvInsPax.ListItems.Count, "", ",")
    Next
    
    For lngC = 0 To txtInsAdd.Count - 1
        strInsAdd = strInsAdd & txtInsAdd(lngC) & IIf(lngC = txtInsAdd.Count - 1, "", ";")
    Next
    
    .AdditionalInfo = txtDescription(0).Text & vbCrLf & txtDescription(1).Text & vbCrLf & Format(dtpDate.value, "Medium Date") & vbCrLf & cmbGeoArea.Text & vbCrLf & txtInsDays _
                      & vbCrLf & dtpInsFromDate.value & vbCrLf & cmbInsPlan.Text _
                      & vbCrLf & strInsPax & vbCrLf & strInsAdd & vbCrLf & txtBTADescription
    



    .ListBoxRem = ListBoxRemark
    .PassengerID = IIf(txtPassengerID = "", "1", txtPassengerID)
    
    If gbolEOAmend Then
       .TicketNumber = mstrTktNum
    Else
       .TicketNumber = "0000"
    End If
    'Added for Disc & TF
    If optNormal Then
        .TicketNumber = IIf(txtFFNo = "", "0000", txtFFNo)
    Else
        .TicketNumber = IIf(txtTktNo = "", "0000", Format(txtALCode, "000") & txtTktNo)
    End If
    .ReplyEmail = Trim(UCase(txtReplyEmail.Text))
End With
End Sub

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


Private Sub txtTktNo_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii)
End Sub

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


Private Sub txtBTADescription_KeyPress(KeyAscii As Integer)
 KeyAscii = fAllowAlphaNumeric(KeyAscii, ".-,: ")
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

