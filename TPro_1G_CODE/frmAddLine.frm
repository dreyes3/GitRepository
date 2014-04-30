VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAddLine 
   Caption         =   "CWT TravelPro - Remarks"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
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
      Left            =   6480
      TabIndex        =   66
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      TabIndex        =   0
      Top             =   7440
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6720
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11853
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Itinerary Remarks (RI)"
      TabPicture(0)   =   "frmAddLine.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lswRI"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Accounting (DI)"
      TabPicture(1)   =   "frmAddLine.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "lswDI"
      Tab(1).Control(2)=   "cmdRefDI"
      Tab(1).Control(3)=   "cmdDelDI"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Invoice Remarks (RD,RP,RT,TUR)"
      TabPicture(2)   =   "frmAddLine.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(1)=   "lswInv"
      Tab(2).Control(2)=   "cmdRefInv"
      Tab(2).Control(3)=   "cmdDelInv"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "General Remarks (NP)"
      TabPicture(3)   =   "frmAddLine.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label5"
      Tab(3).Control(1)=   "Frame1"
      Tab(3).Control(2)=   "lswNP"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Special Services (SSR)"
      TabPicture(4)   =   "frmAddLine.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label2"
      Tab(4).Control(1)=   "Frame2"
      Tab(4).Control(2)=   "lswSS"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Other Services Info (OSI)"
      TabPicture(5)   =   "frmAddLine.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lswOSI"
      Tab(5).Control(1)=   "Frame4"
      Tab(5).Control(2)=   "Label1"
      Tab(5).ControlCount=   3
      Begin VB.CommandButton cmdDelInv 
         Caption         =   "Delete"
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
         Left            =   -68880
         TabIndex        =   70
         Top             =   6120
         Width           =   1455
      End
      Begin VB.CommandButton cmdRefInv 
         Caption         =   "Refresh"
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
         Left            =   -67320
         TabIndex        =   69
         Top             =   6120
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelDI 
         Caption         =   "Delete"
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
         Left            =   -68880
         TabIndex        =   68
         Top             =   6120
         Width           =   1455
      End
      Begin VB.CommandButton cmdRefDI 
         Caption         =   "Refresh"
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
         Left            =   -67320
         TabIndex        =   67
         Top             =   6120
         Width           =   1455
      End
      Begin MSComctlLib.ListView lswInv 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   63
         Top             =   960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8864
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LN"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Text"
            Object.Width           =   14991
         EndProperty
      End
      Begin MSComctlLib.ListView lswDI 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   62
         Top             =   960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LN"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Text"
            Object.Width           =   13227
         EndProperty
      End
      Begin MSComctlLib.ListView lswRI 
         Height          =   2595
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LN"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Flight Info"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Text"
            Object.Width           =   10583
         EndProperty
      End
      Begin MSComctlLib.ListView lswOSI 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   60
         Top             =   960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "GFAX"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vendor"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Text"
            Object.Width           =   12876
         EndProperty
      End
      Begin MSComctlLib.ListView lswSS 
         Height          =   1620
         Left            =   -74880
         TabIndex        =   59
         Top             =   960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2858
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "GFAX"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Segment Number / Flight Info / Status / Text"
            Object.Width           =   7937
         EndProperty
      End
      Begin MSComctlLib.ListView lswNP 
         Height          =   3210
         Left            =   -74880
         TabIndex        =   58
         Top             =   1080
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5662
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LN"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Qualifier"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Text"
            Object.Width           =   13229
         EndProperty
      End
      Begin VB.Frame Frame5 
         Caption         =   "New Remark"
         Height          =   2775
         Left            =   120
         TabIndex        =   47
         Top             =   3840
         Width           =   9135
         Begin VB.ComboBox cmbRIFFText 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1440
            Width           =   7095
         End
         Begin VB.TextBox txtRIRmk 
            Height          =   285
            Left            =   1920
            TabIndex        =   54
            Top             =   1800
            Width           =   7095
         End
         Begin VB.CommandButton cmdAddRI 
            Caption         =   "Add"
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
            Left            =   2880
            TabIndex        =   52
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdUpdRI 
            Caption         =   "Modify"
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
            TabIndex        =   51
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdDelRI 
            Caption         =   "Delete"
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
            Left            =   6000
            TabIndex        =   50
            Top             =   2280
            Width           =   1455
         End
         Begin VB.CommandButton cmdRefRI 
            Caption         =   "Refresh"
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
            Left            =   7560
            TabIndex        =   49
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ListBox lstRISegment 
            Height          =   840
            Left            =   120
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   48
            Top             =   480
            Width           =   5295
         End
         Begin VB.Label Label19 
            Caption         =   "Select Free Form Text"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Remark"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Segment(s)"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "New Remark"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   31
         Top             =   4680
         Width           =   9135
         Begin VB.TextBox txtOSIRmk 
            Height          =   285
            Left            =   1920
            TabIndex        =   46
            Top             =   1080
            Width           =   7095
         End
         Begin VB.ComboBox cmbOSIVendor 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox cmbOSIFFText 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   720
            Width           =   7095
         End
         Begin VB.CommandButton cmdRefOSI 
            Caption         =   "Refresh"
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
            Left            =   7560
            TabIndex        =   35
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdDelOSI 
            Caption         =   "Delete"
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
            Left            =   6000
            TabIndex        =   34
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdUpdOSI 
            Caption         =   "Modify"
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
            TabIndex        =   33
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdAddOSI 
            Caption         =   "Add"
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
            Left            =   2880
            TabIndex        =   32
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Remark"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "Select Vendor"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Select Free Form Text"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "New Remark"
         Height          =   3855
         Left            =   -74880
         TabIndex        =   16
         Top             =   2760
         Width           =   9135
         Begin VB.Frame Frame3 
            Caption         =   "SSR Code"
            Height          =   1935
            Left            =   120
            TabIndex        =   25
            Top             =   1350
            Width           =   8895
            Begin VB.TextBox txtSSRmk 
               Height          =   285
               Left            =   2280
               TabIndex        =   44
               Top             =   1560
               Width           =   6495
            End
            Begin VB.ComboBox cmbSSFFText 
               Height          =   315
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   1200
               Width           =   6495
            End
            Begin VB.OptionButton optSSRCode 
               Caption         =   "Miscellaneous"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   30
               Top             =   840
               Width           =   1575
            End
            Begin VB.ComboBox cmbSSCode 
               Height          =   315
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   240
               Width           =   6495
            End
            Begin VB.OptionButton optSSRCode 
               Caption         =   "Wheel Chair"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   29
               Top             =   540
               Width           =   1575
            End
            Begin VB.OptionButton optSSRCode 
               Caption         =   "Meals"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label17 
               Caption         =   "Remark"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1560
               Width           =   1935
            End
            Begin VB.Label Label3 
               Caption         =   "Select Free Form Text"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   1200
               Width           =   1935
            End
         End
         Begin VB.ListBox lstSSSegment 
            Height          =   840
            Left            =   3720
            MultiSelect     =   2  'Extended
            TabIndex        =   24
            Top             =   480
            Width           =   5295
         End
         Begin VB.ListBox lstSSName 
            Height          =   840
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   22
            Top             =   480
            Width           =   3495
         End
         Begin VB.CommandButton cmdRefSS 
            Caption         =   "Refresh"
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
            Left            =   7560
            TabIndex        =   20
            Top             =   3360
            Width           =   1455
         End
         Begin VB.CommandButton cmdDelSS 
            Caption         =   "Delete"
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
            Left            =   6000
            TabIndex        =   19
            Top             =   3360
            Width           =   1455
         End
         Begin VB.CommandButton cmdUpdSS 
            Caption         =   "Modify"
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
            TabIndex        =   18
            Top             =   3360
            Width           =   1455
         End
         Begin VB.CommandButton cmdAddSS 
            Caption         =   "Add"
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
            Left            =   2880
            TabIndex        =   17
            Top             =   3360
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Segment(s)"
            Height          =   255
            Left            =   3720
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Name(s)"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "New Remark"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   9
         Top             =   4560
         Width           =   9135
         Begin VB.ComboBox cmbNPQualifier 
            Height          =   315
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbNPType 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtNPRmk 
            Height          =   285
            Left            =   1320
            TabIndex        =   41
            Top             =   1080
            Width           =   7695
         End
         Begin VB.CommandButton cmdAddNP 
            Caption         =   "Add"
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
            Left            =   2880
            TabIndex        =   15
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdUpdNP 
            Caption         =   "Modify"
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
            TabIndex        =   14
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdDelNP 
            Caption         =   "Delete"
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
            Left            =   6000
            TabIndex        =   13
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdRefNP 
            Caption         =   "Refresh"
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
            Left            =   7560
            TabIndex        =   12
            Top             =   1560
            Width           =   1455
         End
         Begin VB.ComboBox cmbNPRmk 
            Height          =   315
            ItemData        =   "frmAddLine.frx":00A8
            Left            =   1320
            List            =   "frmAddLine.frx":00AA
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   720
            Width           =   7695
         End
         Begin VB.Label Label8 
            Caption         =   "Remark Qualifier:"
            Height          =   375
            Left            =   3600
            TabIndex        =   72
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label27 
            Caption         =   "Remark Type:"
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Remark"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Select Remark"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   -74880
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   -74880
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   -74880
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Label lblRmkStatus 
      Caption         =   "Remarks and Service Information"
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
      Left            =   120
      TabIndex        =   8
      Top             =   110
      Width           =   5055
   End
End
Attribute VB_Name = "frmAddLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date

Private Sub cmbNPQualifier_Click()
If cmbNPQualifier.ListCount > 0 Then
   loadDefRemarks "NP", cmbNPRmk, "RelateTo=" & cmbNPQualifier.ItemData(cmbNPQualifier.listindex)
   buildNPRemark
End If
End Sub

Private Sub cmbNPRmk_Click()
buildNPRemark
End Sub

Private Sub cmbNPType_Click()
If cmbNPType.ListCount > 0 Then
   loadRmkSubType cmbNPQualifier, "subType2", "SubType1 = '" & cmbNPType.Text & "' Order By SubType2"
   loadDefRemarks "NP", cmbNPRmk, "RelateTo=" & cmbNPQualifier.ItemData(cmbNPQualifier.listindex)
End If
End Sub

Private Sub cmbOSIFFText_Click()
txtOSIRmk = cmbOSIVendor.Text & "*" & cmbOSIFFText.Text
End Sub

Private Sub cmbOSIVendor_Click()
txtOSIRmk = cmbOSIVendor.Text & "*" & cmbOSIFFText.Text
End Sub

Private Sub cmbRIFFText_Click()
buildRIRemark
End Sub

Private Sub cmbSSCode_Click()
Dim rsFFText As ADODB.Recordset
Dim strSql As String
strSql = "Select Additional from tblRemarksType Where Num = " & cmbSSCode.ItemData(cmbSSCode.listindex)
Set rsFFText = gdbConn.Execute(strSql)
If Not rsFFText.EOF Then
   ' 0-Must Not Have Additional Text
   ' 1-Must Have Additional Text
   ' 2-Optional Additional Text
   If rsFFText!additional = 1 Or rsFFText!additional = 2 Then
      cmbSSFFText.Enabled = True
      loadDefRemarks "SS", cmbSSFFText, "RelateTo=" & cmbSSCode.ItemData(cmbSSCode.listindex)
   Else
      cmbSSFFText.Enabled = False
   End If
Else
   cmbSSFFText.Clear
   cmbSSFFText.Enabled = True
End If
Set rsFFText = Nothing
buildSSRemark
End Sub

Private Sub cmbSSFFText_Click()
buildSSRemark
End Sub

Private Sub cmdAddNP_Click()
If validData(txtNPRmk.Text) = True Then
   terminalEntry "NP." & txtNPRmk.Text, "loadNPList"
End If
End Sub

Private Sub cmdAddOSI_Click()
If validData(txtOSIRmk.Text) = True Then
   terminalEntry "SI." & txtOSIRmk.Text, "loadOSIList"
End If
End Sub

Private Sub cmdAddRI_Click()
If validData(txtRIRmk.Text) = True Then
   terminalEntry "RI." & txtRIRmk.Text, "loadRIList"
End If
End Sub

Private Sub cmdAddSS_Click()
If validData(txtSSRmk.Text) = True Then
   terminalEntry "SI." & txtSSRmk.Text, "loadSSList"
End If
End Sub

Private Sub cmdCancel_Click()
gobjHost.terminalEntry "IR"
gobjHost.terminalEntry "IR"
gobjHost.terminalEntry "IR"
Unload Me
End Sub

Private Sub cmdDelDI_Click()
delRemark lswDI, "DI.", "@", "loadDIList"
End Sub

Private Sub cmdDelInv_Click()
delRemark lswInv, "X", "", "loadINVList"
End Sub

Private Sub cmdDelNP_Click()
delRemark lswNP, "NP.", "@", "loadNPList"
End Sub

Private Sub cmdDelOSI_Click()
delRemark lswOSI, "SI.", "@", "loadOSIList"
End Sub

Private Sub cmdDelRI_Click()
delRemark lswRI, "RI.", "@", "loadRIList"
End Sub

Private Sub cmdDelSS_Click()
Dim i As Integer
Dim j As Integer
Dim strTemp As String
For i = 1 To lswSS.ListItems.Count
    If lswSS.ListItems(i).Selected Then
       'Get Passenger Absolute Number
        strTemp = IIf(strTemp = "", "SI.", strTemp & "+SI.")
        For j = 0 To lstSSName.ListCount - 1
            If lstSSName.List(j) = lswSS.ListItems(i).SubItems(2) Then
               'Assign Passenger Absolute Number, segment number and SS code
                strTemp = strTemp & "P" & CStr(j + 1)
                strTemp = strTemp & "S" & Mid(lswSS.ListItems(i).SubItems(3), 1, _
                          InStr(1, lswSS.ListItems(i).SubItems(3), "/") - 2)
                strTemp = strTemp & "/" & lswSS.ListItems(i).SubItems(1)
                Exit For
            End If
        Next
        strTemp = strTemp & "@"
    End If
Next
If strTemp = "" Then Exit Sub
terminalEntry strTemp, "loadSSList"

End Sub

Private Sub cmdOK_Click()
datTouchEnd = Now
gobjHost.terminalEntry "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
gobjHost.terminalEntry "ER+ER+ER"
'gobjHost.terminalEntry "ER"
'gobjHost.terminalEntry "ER"

       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModRmk, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModRmk, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModRmk, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd


Unload Me
End Sub

Private Sub cmdRefDI_Click()
loadPNR
loadDIList
End Sub

Private Sub cmdRefInv_Click()
loadPNR
loadInvList
End Sub

Private Sub cmdRefNP_Click()
loadPNR
loadNPList
loadRmkSubType cmbNPType, "subtype1", "Type='NP' Order By SubType1 ", True
End Sub

Private Sub cmdRefOSI_Click()
loadPNR
loadVendor
loadOSIList
loadDefRemarks "OSI", cmbOSIFFText
End Sub

Private Sub cmdRefRI_Click()
loadPNR
loadRIList
loadDefRemarks "RI", cmbRIFFText
loadSegments lstRISegment, True, True, True
End Sub

Private Sub cmdRefSS_Click()
loadPNR
loadSSList
optSSRCode(0).value = True
loadPassengers
loadSegments lstSSSegment, True, False, False

End Sub

Private Sub cmdUpdNP_Click()
Dim strTemp As String
If validData(txtNPRmk.Text) Then
   strTemp = "NP." & Format(lswNP.SelectedItem, "0") & "@" & IIf(Trim(cmbNPQualifier) = "", "", Trim(cmbNPQualifier)) & txtNPRmk.Text
   terminalEntry strTemp, "loadNPList"
End If
End Sub

Private Sub cmdUpdOSI_Click()
Dim strTemp As String
If validData(txtOSIRmk.Text) Then
    strTemp = "SI." & Format(lswOSI.SelectedItem, "0") & "@" & txtOSIRmk.Text
    terminalEntry strTemp, "loadOSIList"
End If
End Sub

Private Sub cmdUpdRI_Click()
Dim strTemp As String
If validData(txtRIRmk.Text) Then
   strTemp = "RI." & Format(lswRI.SelectedItem, "0") & "@" & txtRIRmk.Text
   terminalEntry strTemp, "loadRIList"
End If
End Sub

Private Sub cmdUpdSS_Click()
Dim i As Integer
Dim strTemp As String
If validData(txtSSRmk.Text) = True Then
    'Get Passenger Absolute Number
    For i = 0 To lstSSName.ListCount - 1
        If lstSSName.List(i) = lswSS.SelectedItem.SubItems(2) Then
           'Assign Passenger Absolute Number, segmetn number and SS code
           strTemp = "P" & CStr(i + 1)
           strTemp = strTemp & "S" & Mid(lswSS.SelectedItem.SubItems(3), 1, _
                     InStr(1, lswSS.SelectedItem.SubItems(3), "/") - 2)
           strTemp = strTemp & "/" & lswSS.SelectedItem.SubItems(1)
           Exit For
        End If
    Next
    strTemp = "SI." & strTemp & "@" & Trim(txtSSRmk)
    terminalEntry strTemp, "loadSSList"
End If
End Sub

Private Sub Form_Load()
Dim oldParent As Long
datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
Me.Move 0, 0
Me.Move frmSideBar.Width, 0

RemoveMenus Me, False, False, _
        False, False, False, True, True
loadPNR
'load all remarks in PNR
loadNPList
loadSSList
loadOSIList
loadRIList
loadDIList
loadInvList

'load all frequently used remarks from database
loadRmkSubType cmbNPType, "subtype1", "Type='NP' Order By SubType1 ", True
optSSRCode(0).value = True
loadDefRemarks "OSI", cmbOSIFFText
loadDefRemarks "RI", cmbRIFFText

'load passengers and air segments
loadPassengers
loadSegments lstSSSegment, True, False, False
loadSegments lstRISegment, True, True, True

'load vendor
loadVendor
'Me.Top = (Screen.Height - Me.Height) / 2
'Me.Left = (Screen.Width - Me.Width) - 25
datFormLoadEnd = Now
If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

End Sub
Private Sub loadPNR()
Set gobjPNR = New CWT_GalileoPNR3.PNR
'If gobjLog.LogOpen Then gobjPNR.OpenLog gobjLog
gobjPNR.loadPNR
End Sub
Sub loadNPList()
Dim i As Integer
Dim item As ListItem
Dim strText As String

lswNP.ListItems.Clear
For i = 1 To gobjPNR.GeneralRemarkCount
   With gobjPNR.GeneralRemark(i)
      Set item = lswNP.ListItems.Add(, , Format(.ItemNum, "000"))
      strText = .Qualifier
      'Qualifier with 1 char will contain *. For instance, qualifier S will returned *S
      If Mid(strText, 1, 1) = "*" Then
         strText = Mid(strText, 2, Len(strText) - 1)
      End If
      item.SubItems(1) = strText
      item.SubItems(2) = .RemarkText
   End With
Next
For i = 1 To lswNP.ListItems.Count
   lswNP.ListItems(i).Selected = False
Next
cmdUpdNP.Enabled = False
cmdDelNP.Enabled = False
txtNPRmk = ""
End Sub

Private Sub lstRISegment_Click()
buildRIRemark
End Sub

Private Sub lstSSName_Click()
buildSSRemark
End Sub

Private Sub lstSSSegment_Click()
buildSSRemark
End Sub

Private Sub lswNP_Click()
txtNPRmk = IIf(lswNP.SelectedItem.SubItems(1) = "", "", _
           lswNP.SelectedItem.SubItems(1) & "*") & lswNP.SelectedItem.SubItems(2)

chkListView lswNP, cmdUpdNP, cmdDelNP
End Sub

Sub loadSSList()
Dim i As Integer
Dim intC As Integer
Dim intTemp As Integer
Dim strTemp As String
Dim item As ListItem

lswSS.ListItems.Clear
For i = 1 To gobjPNR.SSRCount
    With gobjPNR.SSR(i)
         Set item = lswSS.ListItems.Add(, , Format(.GFax, "000"))
         item.SubItems(1) = .SSCode
         intTemp = .PsgNum
         'Search Passenger Name
         For intC = 1 To gobjPNR.PassengerCount
             With gobjPNR.PassengerName(intC)
                  If .PassengerNum = intTemp Then
                     item.SubItems(2) = Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
                     Exit For
                  End If
             End With
         Next intC
         intTemp = .SegNum
         'Search Air Segment
         For intC = 1 To gobjPNR.AirSegCount
             With gobjPNR.AirSeg(intC)
                  If .segnumber = intTemp Then
                     strTemp = .segnumber & " / " & .Vendor & ":" & Format(.DepartDateTime, "DDMMM") & " " & .DepartCityCode & "-" & .ArriveCityCode
                     Exit For
                  End If
             End With
         Next intC
         strTemp = strTemp & " / " & .Status & " / " & .Text
         item.SubItems(3) = strTemp
         strTemp = ""
    End With
Next
For i = 1 To lswSS.ListItems.Count
   lswSS.ListItems(i).Selected = False
Next
cmdUpdSS.Enabled = False
cmdDelSS.Enabled = False
txtSSRmk = ""
End Sub
Sub loadOSIList()
Dim i As Integer
Dim item As ListItem

lswOSI.ListItems.Clear
For i = 1 To gobjPNR.OSICount
   With gobjPNR.OSI(i)
      Set item = lswOSI.ListItems.Add(, , Format(.GFax, "000"))
      item.SubItems(1) = .Vendor
      item.SubItems(2) = .Text
   End With
Next
For i = 1 To lswOSI.ListItems.Count
   lswOSI.ListItems(i).Selected = False
Next
cmdUpdOSI.Enabled = False
cmdDelOSI.Enabled = False
txtOSIRmk = ""
End Sub
Sub loadRIList()
Dim i As Integer
Dim intC As Integer
Dim intTemp As Integer
Dim item As ListItem

lswRI.ListItems.Clear
For i = 1 To gobjPNR.ItinRemarkCount
   With gobjPNR.ItinRemark(i)
      Set item = lswRI.ListItems.Add(, , Format(.ItemNum, "000"))
      intTemp = .SegNum
      'Search Air Segment
      For intC = 1 To gobjPNR.AirSegCount
          With gobjPNR.AirSeg(intC)
               If .segnumber = intTemp Then
                   item.SubItems(1) = .segnumber & " / " & .Vendor & ":" & Format(.DepartDateTime, "DDMMM") & " " & .DepartCityCode & "-" & .ArriveCityCode
                   Exit For
               End If
          End With
      Next intC
      
      'Search Hotel Segment
      For intC = 1 To gobjPNR.HotelSegCount
          With gobjPNR.HotelSeg(intC)
               If .SegNum = intTemp Then
                   item.SubItems(1) = .SegNum & " / " & .Vendor & ":" & Format(.CheckInDate, "DDMMM") & " " & .CityCode
                   Exit For
               End If
          End With
      Next intC
      
      'Search Car Segment
      For intC = 1 To gobjPNR.CarSegCount
          With gobjPNR.CarSeg(intC)
               If .SegNum = intTemp Then
                   item.SubItems(1) = .SegNum & " / " & .Vendor & ":" & Format(.StartDtTime, "DDMMM") & " " & .StartPt
                   Exit For
               End If
          End With
      Next intC
      
      
      item.SubItems(2) = .RemarkText
   End With
Next
For i = 1 To lswRI.ListItems.Count
   lswRI.ListItems(i).Selected = False
Next
cmdUpdRI.Enabled = False
cmdDelRI.Enabled = False
txtRIRmk = ""
End Sub
Sub loadInvList()
Dim i As Integer
Dim item As ListItem
Dim strText As String
Dim strSegType As String

lswInv.ListItems.Clear
For i = 1 To gobjPNR.TurSegCount
   With gobjPNR.TurSeg(i)
      Set item = lswInv.ListItems.Add(, , Format(.SegNum, "00"))
      strText = .SegType & " " & .Vnd & " " & .Status & " " & .NumPersons & " " & _
                .StartPt & " " & Format(.StartDt, "DDMMM") & "-" & .Text
      item.SubItems(1) = strText
   End With
Next
For i = 1 To gobjPNR.PaidDueCount
   With gobjPNR.PaidDue(i)
      
      If .SegType = "D" Then
         strSegType = "DUE "
      ElseIf .SegType = "P" Then
         strSegType = "PAID"
      ElseIf .SegType = "T" Then
         strSegType = "TEXT"
      End If
   
      Set item = lswInv.ListItems.Add(, , Format(.SegNum, "00"))
      strText = .productType & " " & "  **  " & strSegType & "  **  " & _
                Format(.SegDate, "DDMMM") & "-" & .FreeText
      If strSegType <> "TEXT" Then
         strText = strText & strSegType & _
                   " " & .CurrencyCode & " " & Format(.Amount, "0.00") & "**"
      End If
      item.SubItems(1) = strText
   End With
Next
For i = 1 To lswInv.ListItems.Count
   lswInv.ListItems(i).Selected = False
Next
End Sub
Sub loadDIList()
Dim i As Integer
Dim item As ListItem
Dim strText As String

lswDI.ListItems.Clear
For i = 1 To gobjPNR.AcctRemarkCount
   With gobjPNR.AcctRemark(i)
      Set item = lswDI.ListItems.Add(, , Format(.ItemNum, "000"))
      If .RemarkType = "FA" Then
         strText = "AGT INF-"
      ElseIf .RemarkType = "FT" Then
         strText = "FREE TEXT-"
      ElseIf .RemarkType = "AC" Then
         strText = "AC ACCT-"
      ElseIf .RemarkType = "FS" Then
         strText = "$ SAVE-"
      ElseIf .RemarkType = "TK" Then
         strText = "TKT NO-"
      Else
         strText = "-"
      End If
      strText = strText & .RemarkText
      If .RemarkType = "FA" Then
          item.SubItems(1) = "AR"
      Else
          item.SubItems(1) = .RemarkType
      End If
      item.SubItems(2) = strText
      strText = ""
   End With
Next
For i = 1 To lswDI.ListItems.Count
   lswDI.ListItems(i).Selected = False
Next
End Sub
Private Sub loadDefRemarks(ByVal strType As String, ByRef cmbRmk As ComboBox, Optional ByVal strCriteria2 As String)
Dim rsRemarks As ADODB.Recordset
Dim strSql As String
strSql = "Select * from tblremarks where Type='" & strType & "'" & IIf(strCriteria2 = "", "", " AND " & strCriteria2) & " order by Remark"
Set rsRemarks = gdbConn.Execute(strSql)
cmbRmk.Clear
While Not rsRemarks.EOF
      cmbRmk.AddItem rsRemarks!Remark
      rsRemarks.MoveNext
Wend
If cmbRmk.ListCount > 0 Then
   cmbRmk.listindex = 0
End If
Set rsRemarks = Nothing
End Sub

Private Sub loadRmkSubType(ByRef combo1 As ComboBox, ByVal strColumn As String, Optional ByVal strCriteria As String, Optional ByVal bolField As Boolean)
Dim rsSubType As ADODB.Recordset
Dim strSql As String
strSql = "Select DISTINCT " & IIf(bolField = True, strColumn, "*") & " from tblRemarksType" & IIf(strCriteria = "", "", " Where " & strCriteria)
Set rsSubType = gdbConn.Execute(strSql)
combo1.Clear
While Not rsSubType.EOF
    combo1.AddItem rsSubType.Fields(strColumn)
    If bolField = False Then
       combo1.ItemData(combo1.NewIndex) = rsSubType!Num
    End If
    rsSubType.MoveNext
Wend
Set rsSubType = Nothing
If combo1.ListCount > 0 Then
   combo1.listindex = 0
End If
End Sub

Private Sub lswOSI_click()
txtOSIRmk = lswOSI.SelectedItem.SubItems(1) & "*" & lswOSI.SelectedItem.SubItems(2)
chkListView lswOSI, cmdUpdOSI, cmdDelOSI
End Sub

Private Sub lswRI_Click()
txtRIRmk = lswRI.SelectedItem.SubItems(2)
chkListView lswRI, cmdUpdRI, cmdDelRI
End Sub

Private Sub lswSS_Click()
txtSSRmk = ""
chkListView lswSS, cmdUpdSS, cmdDelSS
End Sub

Private Sub optSSRCode_Click(Index As Integer)
'Index 0 = MEAL, 1 = WHEEL, 2 = MISC
If Index = 0 Then
   loadRmkSubType cmbSSCode, "subtype2", "Type='SS' AND " & "SubType1='MEAL' Order By SubType2"
   cmbSSCode.Top = optSSRCode(Index).Top
ElseIf Index = 1 Then
   loadRmkSubType cmbSSCode, "subtype2", "Type='SS' AND " & "SubType1='WHEEL' Order By SubType2"
   cmbSSCode.Top = optSSRCode(Index).Top
ElseIf Index = 2 Then
   loadRmkSubType cmbSSCode, "subtype2", "Type='SS' AND " & "SubType1='MISC' Order By SubType2"
   cmbSSCode.Top = optSSRCode(Index).Top
End If
End Sub

Private Sub loadPassengers()
Dim intC As Integer
Dim strTemp As String
lstSSName.Clear
For intC = 1 To gobjPNR.PassengerCount
    With gobjPNR.PassengerName(intC)
        strTemp = Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
    End With
    lstSSName.AddItem strTemp
Next
End Sub

Private Sub loadSegments(ByRef lstbox As ListBox, bolAirSeg As Boolean, bolHotelSeg As Boolean, bolCarSeg As Boolean)
Dim intC As Integer

lstbox.Clear
'load Air Segments
If bolAirSeg = True Then
    For intC = 1 To gobjPNR.AirSegCount
       With gobjPNR.AirSeg(intC)
            If .Flown = False Then
               lstbox.AddItem .TextAirSeg
               lstbox.ItemData(lstbox.NewIndex) = .segnumber
            End If
       End With
    Next
End If

'load Hotel Segments
If bolHotelSeg = True Then
    For intC = 1 To gobjPNR.HotelSegCount
        With gobjPNR.HotelSeg(intC)
             If .CheckInDate >= Date Then
                lstbox.AddItem .TextHtlSeg
                lstbox.ItemData(lstbox.NewIndex) = .SegNum
             End If
        End With
    Next
End If

'load Car Segments
If bolCarSeg = True Then
    For intC = 1 To gobjPNR.CarSegCount
        With gobjPNR.CarSeg(intC)
             If .StartDtTime >= Date Then
                lstbox.AddItem .TextCarSeg
                lstbox.ItemData(lstbox.NewIndex) = .SegNum
             End If
        End With
    Next
End If
End Sub

Private Sub loadVendor()
Dim intC As Integer
Dim strCarrierList As String
cmbOSIVendor.Clear
If gobjPNR.AirSegCount > 0 Then
    For intC = 1 To gobjPNR.AirSegCount
        With gobjPNR.AirSeg(intC)
            If InStr(1, strCarrierList, .Vendor) = 0 Then
                strCarrierList = strCarrierList & IIf(strCarrierList = "", "", "/") & .Vendor
                cmbOSIVendor.AddItem .Vendor
            End If
        End With
    Next
End If
If cmbOSIVendor.ListCount > 0 Then
   cmbOSIVendor.listindex = 0
   If cmbOSIVendor.ListCount > 1 Then cmbOSIVendor.AddItem "YY"
End If
End Sub

Private Function validData(ByVal strRmk As String) As Boolean
If Trim(strRmk) = "" Then
   MsgBox "Missing Remark...", vbApplicationModal + vbExclamation + vbOKOnly, "ERROR!"
   validData = False
Else
   validData = True
End If
End Function

Private Sub terminalEntry(ByVal strRmk As String, ByVal strFunction As String)
Dim strResponse As String
Dim strTemp, strType As String
Dim i, j, k As Integer
strResponse = gobjHost.terminalEntry(strRmk)
'Prompt msgbox if failed to add remark to PNR
If (Right(strResponse, 1) = "*") = False Then
   'Checking for invoice remark because Galileo will not return * upon success
   If SSTab1.Tab = 2 Then
      i = lswInv.ListItems.Count
      loadPNR
      loadInvList
      If lswInv.ListItems.Count <> i Then GoTo Success
   End If
   MsgBox strResponse, vbApplicationModal + vbExclamation + vbOKOnly, "ERROR!"
Else
Success:
   loadPNR
   CallByName Me, strFunction, VbMethod
End If
End Sub

Private Sub chkListView(ByRef lstView As ListView, ByRef cmbUpdate As CommandButton, ByRef cmbDelete As CommandButton)
Dim i As Integer
Dim intCountSelected As Integer
Dim bolDel As Boolean
Dim bolModify As Boolean
For i = 1 To lstView.ListItems.Count
   If lstView.ListItems(i).Selected Then
      bolDel = True
      bolModify = True
      intCountSelected = intCountSelected + 1
      If intCountSelected >= 2 Then
         bolModify = False
         Exit For
      End If
   End If
Next
cmbUpdate.Enabled = bolModify
cmbDelete.Enabled = bolDel

'Disable modify button if is preset remarks in Notepad
If bolModify = True And lstView.Name = "lswNP" Then
   cmbUpdate.Enabled = chkNPPreset(lstView.SelectedItem.SubItems(1), lstView.SelectedItem.SubItems(2))
   txtNPRmk.Enabled = cmbUpdate.Enabled
End If
End Sub

Private Sub delRemark(ByRef lstView As ListView, ByVal strSuffix As String, ByVal strPrefix As String, ByVal strFunction As String)
Dim i As Integer
Dim strNum As String
    
For i = 1 To lstView.ListItems.Count
    If lstView.ListItems(i).Selected Then
       strNum = strNum & IIf(strNum <> "", ".", "") & Format(lstView.ListItems(i).Text, "0")
    End If
Next
If strNum = "" Then Exit Sub
terminalEntry strSuffix & strNum & strPrefix, strFunction

End Sub

Private Sub buildSSRemark()
Dim i As Integer
Dim strTemp As String

'Build Passenger Selected
For i = 0 To lstSSName.ListCount - 1
    If lstSSName.Selected(i) = True Then
       strTemp = strTemp & IIf(strTemp = "", "P" & CStr(i + 1), "." & CStr(i + 1))
    End If
Next

'Build Segment Selected
For i = 0 To lstSSSegment.ListCount - 1
    If lstSSSegment.Selected(i) = True Then
       strTemp = strTemp & IIf(InStr(1, strTemp, "S") = 0, "S" & lstSSSegment.ItemData(i), "." & lstSSSegment.ItemData(i))
    End If
Next
If strTemp <> "" Then strTemp = strTemp & "/"
strTemp = strTemp & Mid(cmbSSCode.Text, 1, InStr(cmbSSCode.Text, "-") - 2)
If cmbSSFFText.Enabled = True And cmbSSFFText.ListCount > 0 Then
   strTemp = strTemp & "*" & cmbSSFFText.Text
End If
txtSSRmk = strTemp
End Sub

Private Sub txtNPRmk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtOSIRmk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRIRmk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSSRmk_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub buildRIRemark()
Dim strTemp As String
Dim i As Integer
'Build Segment Selected
For i = 0 To lstRISegment.ListCount - 1
    If lstRISegment.Selected(i) = True Then
       strTemp = strTemp & IIf(InStr(1, strTemp, "S") = 0, "S" & lstRISegment.ItemData(i), "." & lstRISegment.ItemData(i))
    End If
Next
txtRIRmk = strTemp & IIf(strTemp = "", "", "*") & cmbRIFFText.Text
End Sub

Private Sub buildNPRemark()
txtNPRmk.Enabled = chkNPPreset(cmbNPQualifier.Text, cmbNPRmk.Text)
txtNPRmk.Text = ""
txtNPRmk = IIf(cmbNPQualifier.Text = "", "", cmbNPQualifier.Text & "*") & cmbNPRmk.Text
End Sub

Private Function chkNPPreset(ByVal strQualifier As String, ByVal strRemark As String) As Boolean
Dim rsExist As ADODB.Recordset
Dim strSql As String

chkNPPreset = True
strSql = "Select a.Additional from tblRemarksType a inner join tblRemarks b on a.Type = 'NP' and a.Type=b.Type and a.Num = b.RelateTo and a.SubType2=" & _
        "'" & strQualifier & "' and b.Remark ='" & strRemark & "'"
Set rsExist = gdbConn.Execute(strSql)
If rsExist.EOF = False Then
   If rsExist!additional = 1 Then chkNPPreset = False
End If

Set rsExist = Nothing

End Function

Private Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    hMenu = GetSystemMenu(hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub


