VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDespDI 
   Caption         =   "Despatch / Visa Update"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   10545
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   1680
      TabIndex        =   20
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
   End
   Begin TabDlg.SSTab sstDespVisa 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Despatch"
      TabPicture(0)   =   "frmDespDI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAgentName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbAgentName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpStartDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "dtpEndDate"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmbAgentID"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraDesp"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkDesp"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdRefresh"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Visa"
      TabPicture(1)   =   "frmDespDI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "lblVisaAgentName"
      Tab(1).Control(4)=   "dtpStartSubDate"
      Tab(1).Control(5)=   "dtpEndSubDate"
      Tab(1).Control(6)=   "cmdVisaRefresh"
      Tab(1).Control(7)=   "cmbVisaAgentID"
      Tab(1).Control(8)=   "cmbVisaAgentName"
      Tab(1).Control(9)=   "fraVisa"
      Tab(1).Control(10)=   "chkVisa"
      Tab(1).Control(11)=   "Frame2"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Despatch Reset"
      TabPicture(2)   =   "frmDespDI.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdReset(0)"
      Tab(2).Control(1)=   "chkSelLine(0)"
      Tab(2).Control(2)=   "cmbEmpName(0)"
      Tab(2).Control(3)=   "cmbEmpID(0)"
      Tab(2).Control(4)=   "cmdRefreshList(0)"
      Tab(2).Control(5)=   "fraJobList(0)"
      Tab(2).Control(6)=   "dtpEDate(0)"
      Tab(2).Control(7)=   "dtpSDate(0)"
      Tab(2).Control(8)=   "lblEmpName(0)"
      Tab(2).Control(9)=   "Label9"
      Tab(2).Control(10)=   "Label8"
      Tab(2).Control(11)=   "Label7(0)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Visa Reset"
      TabPicture(3)   =   "frmDespDI.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdReset(1)"
      Tab(3).Control(1)=   "chkSelLine(1)"
      Tab(3).Control(2)=   "cmbEmpName(1)"
      Tab(3).Control(3)=   "cmbEmpID(1)"
      Tab(3).Control(4)=   "cmdRefreshList(1)"
      Tab(3).Control(5)=   "fraJobList(1)"
      Tab(3).Control(6)=   "dtpEDate(1)"
      Tab(3).Control(7)=   "dtpSDate(1)"
      Tab(3).Control(8)=   "Label12"
      Tab(3).Control(9)=   "Label11"
      Tab(3).Control(10)=   "lblEmpName(1)"
      Tab(3).Control(11)=   "Label7(1)"
      Tab(3).ControlCount=   12
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   60
         Top             =   5760
         Width           =   9855
         Begin VB.TextBox txtVisaCCNum 
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
            TabIndex        =   65
            Tag             =   "NN"
            Top             =   300
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.ComboBox cmbVisaFOPType 
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
            ItemData        =   "frmDespDI.frx":0070
            Left            =   240
            List            =   "frmDespDI.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   300
            Width           =   1515
         End
         Begin VB.ComboBox cmbVisaCCType 
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
            ItemData        =   "frmDespDI.frx":0087
            Left            =   1800
            List            =   "frmDespDI.frx":00A3
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   300
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chkVisaAbsorbMer 
            Caption         =   "Absorb Merchant Fee     (For Visa Handling Fee)"
            Height          =   375
            Left            =   6000
            TabIndex        =   62
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdVisaAmend 
            Caption         =   "Amend"
            Height          =   375
            Left            =   8280
            TabIndex        =   61
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpVisaCCExp 
            Height          =   360
            Left            =   4680
            TabIndex        =   66
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
            Format          =   62980099
            CurrentDate     =   36526
            MaxDate         =   73050
            MinDate         =   36526
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   495
         Left            =   5400
         TabIndex        =   6
         Top             =   5280
         Width           =   975
      End
      Begin VB.Frame Frame1 
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
         Height          =   735
         Left            =   240
         TabIndex        =   53
         Top             =   5760
         Width           =   8775
         Begin VB.CommandButton cmdAmend 
            Caption         =   "Amend"
            Height          =   375
            Left            =   7560
            TabIndex        =   59
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkAbsorbMer 
            Caption         =   "Absorb Merchant Fee"
            Height          =   375
            Left            =   5760
            TabIndex        =   58
            Top             =   240
            Width           =   1455
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
            ItemData        =   "frmDespDI.frx":00C7
            Left            =   1800
            List            =   "frmDespDI.frx":00E3
            Style           =   2  'Dropdown List
            TabIndex        =   56
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
            ItemData        =   "frmDespDI.frx":0107
            Left            =   240
            List            =   "frmDespDI.frx":0111
            Style           =   2  'Dropdown List
            TabIndex        =   55
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
            TabIndex        =   54
            Tag             =   "NN"
            Top             =   300
            Visible         =   0   'False
            Width           =   2025
         End
         Begin MSComCtl2.DTPicker dtpCCExp 
            Height          =   360
            Left            =   4680
            TabIndex        =   57
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
            Format          =   62980099
            CurrentDate     =   36526
            MaxDate         =   73050
            MinDate         =   36526
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   495
         Index           =   1
         Left            =   -66360
         TabIndex        =   52
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   495
         Index           =   0
         Left            =   -66360
         TabIndex        =   51
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CheckBox chkSelLine 
         Caption         =   "Select Line"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   39
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cmbEmpName 
         Height          =   315
         Index           =   1
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   4800
         Width           =   3855
      End
      Begin VB.ComboBox cmbEmpID 
         Height          =   315
         Index           =   1
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   4800
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmdRefreshList 
         Caption         =   "Refresh"
         Height          =   495
         Index           =   1
         Left            =   -69600
         TabIndex        =   41
         Top             =   5280
         Width           =   975
      End
      Begin VB.Frame fraJobList 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3735
         Index           =   1
         Left            =   -74880
         TabIndex        =   40
         Top             =   960
         Width           =   9975
         Begin MSComctlLib.ListView lvwJobList 
            Height          =   3495
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   16
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VisaID"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Sub. Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Embassy"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Pax Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Company Name"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "FOP"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "VendorNumber"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "VisaCost"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "VisaHandlingFee"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "AbsorbMer"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Tourist"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Business"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Transit"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Single"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Double"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "Multiple"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.CheckBox chkSelLine 
         Caption         =   "Select Line"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   27
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cmbEmpName 
         Height          =   315
         Index           =   0
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   4800
         Width           =   3855
      End
      Begin VB.ComboBox cmbEmpID 
         Height          =   315
         Index           =   0
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   4800
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmdRefreshList 
         Caption         =   "Refresh"
         Height          =   495
         Index           =   0
         Left            =   -69600
         TabIndex        =   30
         Top             =   5280
         Width           =   975
      End
      Begin VB.Frame fraJobList 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3735
         Index           =   0
         Left            =   -74880
         TabIndex        =   28
         Top             =   1080
         Width           =   9975
         Begin MSComctlLib.ListView lvwJobList 
            Height          =   3495
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "DespatchID"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "DeliveryDate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Attention"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Company"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "FOP"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "BillClient"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "OffScheduleAmt"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "AbsorbMer"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CheckBox chkVisa 
         Caption         =   "Select Line"
         Height          =   375
         Left            =   -74760
         TabIndex        =   22
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkDesp 
         Caption         =   "Select Line"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   2415
      End
      Begin VB.Frame fraVisa 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3735
         Left            =   -74880
         TabIndex        =   25
         Top             =   1080
         Width           =   9975
         Begin MSComctlLib.ListView lvwVisa 
            Height          =   3495
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "VisaID"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Sub. Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Embassy"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Pax Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Company Name"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "VendorNumber"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "VisaCost"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "VisaHandlingFee"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "AbsorbMer"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Tourist"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Business"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Transit"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Single"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Double"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Multiple"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame fraDesp 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3735
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   9975
         Begin MSComctlLib.ListView lvwDespatch 
            Height          =   3495
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6165
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "DespatchID"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "DeliveryDate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Attention"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Company"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "BillClient"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "OffScheduleAmt"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "AbsorbMer"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.ComboBox cmbVisaAgentName 
         Height          =   315
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4800
         Width           =   3855
      End
      Begin VB.ComboBox cmbVisaAgentID 
         Height          =   315
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4800
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmdVisaRefresh 
         Caption         =   "Refresh"
         Height          =   495
         Left            =   -69600
         TabIndex        =   11
         Top             =   5280
         Width           =   975
      End
      Begin VB.ComboBox cmbAgentID 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4800
         Visible         =   0   'False
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   38411
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   38411
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4800
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker dtpEndSubDate 
         Height          =   375
         Left            =   -71280
         TabIndex        =   13
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   38411
      End
      Begin MSComCtl2.DTPicker dtpStartSubDate 
         Height          =   375
         Left            =   -73560
         TabIndex        =   14
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   38411
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Index           =   0
         Left            =   -71280
         TabIndex        =   32
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   38411
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Index           =   0
         Left            =   -73560
         TabIndex        =   33
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   38411
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Index           =   1
         Left            =   -71280
         TabIndex        =   43
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   38411
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Index           =   1
         Left            =   -73560
         TabIndex        =   44
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62980097
         CurrentDate     =   38411
      End
      Begin VB.Label Label12 
         Caption         =   "Agent Name"
         Height          =   375
         Left            =   -74760
         TabIndex        =   49
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Delivery Date"
         Height          =   375
         Left            =   -74760
         TabIndex        =   48
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label lblEmpName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   47
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "to"
         Height          =   255
         Index           =   1
         Left            =   -71760
         TabIndex        =   46
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label lblEmpName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   38
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label9 
         Caption         =   "Agent Name"
         Height          =   375
         Left            =   -74760
         TabIndex        =   37
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Delivery Date"
         Height          =   375
         Left            =   -74760
         TabIndex        =   36
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "to"
         Height          =   255
         Index           =   0
         Left            =   -71760
         TabIndex        =   35
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label lblVisaAgentName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -74760
         TabIndex        =   19
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Agent Name"
         Height          =   375
         Left            =   -74760
         TabIndex        =   18
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Sub Date"
         Height          =   375
         Left            =   -74760
         TabIndex        =   17
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "to"
         Height          =   255
         Left            =   -71760
         TabIndex        =   16
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "to"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Delivery Date"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label lblAgentName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Agent Name"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   4800
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDespDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sngVisaCPct As Single
Dim sngVisaHPct As Single
Dim sngDespatchPct As Single
Dim strVisaCPre As String
Dim strVisaHPre As String
Dim strDespatchPre As String
Dim strVisaCDesc As String
Dim strVisaHDesc As String
Dim strDespatchDesc As String
Const strVisaHVendor = "021238"
Const mGSTPer = 0.05

Dim datTouchStart As Date
Dim datFormLoadStart As Date
Dim datTouchEnd As Date

Private Sub chkDesp_Click()
   If chkDesp.value = 0 Then
      fraDesp.Enabled = False
      lvwDespatch.HideSelection = True
   Else
      fraDesp.Enabled = True
      lvwDespatch.HideSelection = False
      lvwDespatch.SetFocus
   End If
End Sub

Private Sub chkSelLine_Click(Index As Integer)
   If chkSelLine(Index).value = 0 Then
      fraJobList(Index).Enabled = False
      lvwJobList(Index).HideSelection = True
   Else
      fraJobList(Index).Enabled = True
      lvwJobList(Index).HideSelection = False
      lvwJobList(Index).SetFocus
   End If
End Sub

Private Sub chkVisa_Click()
   If chkVisa.value = 0 Then
      fraVisa.Enabled = False
      lvwVisa.HideSelection = True
   Else
      fraVisa.Enabled = True
      lvwVisa.HideSelection = False
      lvwVisa.SetFocus
   End If
End Sub

Private Sub cmbAgentName_Click()
   cmbAgentID.listindex = cmbAgentName.listindex
End Sub

Private Sub cmbCCType_Click()
AbsorbMer
End Sub

Private Sub cmbEmpName_Click(Index As Integer)
   cmbEmpID(Index).listindex = cmbEmpName(Index).listindex
End Sub

Private Sub cmbFOPType_Click()
Dim blnCC As Boolean
    
blnCC = (cmbFOPType = "CX")
cmbCCType.Visible = blnCC
txtCCNum.Visible = blnCC
dtpCCExp.Visible = blnCC

AbsorbMer
End Sub

Private Sub cmbVisaAgentName_Click()
   cmbVisaAgentID.listindex = cmbVisaAgentName.listindex
End Sub

Private Sub cmbVisaCCType_Click()
VisaAbsorbMer
End Sub

Private Sub cmbVisaFOPType_Click()
Dim blnCC As Boolean
    
blnCC = (cmbVisaFOPType = "CX")
cmbVisaCCType.Visible = blnCC
txtVisaCCNum.Visible = blnCC
dtpVisaCCExp.Visible = blnCC

VisaAbsorbMer
End Sub
'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
'Private Sub cmdAmend_Click()
'   Dim strSql As String
'   Dim strFOP As String
'   Dim rsFOP As ADODB.Recordset
'   Dim intI As Integer
'   Dim DespatchID As String
'
'   If ValidCCData = False Then Exit Sub
'
'   DespatchID = lvwDespatch.SelectedItem.Text
'
'   If cmbFOPType.Text = "INV" Then
'      strFOP = "INV"
'   ElseIf cmbFOPType.Text = "CX" Then
'      strFOP = "CX/" & cmbCCType.Text & "/" & txtCCNum & "/" & Format(dtpCCExp, "mmyy")
'   End If
'
'   strSql = "Update tblDespatch "
'   strSql = strSql & "Set FOP = '" & strFOP & "' "
'   strSql = strSql & ", AbsorbMer = " & chkAbsorbMer.value
'   strSql = strSql & " Where DespatchID = " & lvwDespatch.SelectedItem.Text
'
'   gdbDespatch.Execute strSql
'
'    strSql = "Select * from tblDespFOPAmend "
'    strSql = strSql & "where DespatchID = -1"
'    Set rsFOP = New ADODB.Recordset
'    rsFOP.LockType = adLockBatchOptimistic
'    rsFOP.CursorType = adOpenKeyset
'    rsFOP.Source = strSql
'    Set rsFOP.ActiveConnection = gdbDespatch
'    rsFOP.Open
'    rsFOP.AddNew
'    rsFOP("DespatchId") = lvwDespatch.SelectedItem.Text
'    rsFOP("FPID") = gobjHost.AgentSine
'    rsFOP("FPPCC") = gobjHost.AgentPCC
'    rsFOP("AmendDate") = Now
'    rsFOP("OldFOP") = lvwDespatch.SelectedItem.SubItems(4)
'    rsFOP("NewFOP") = strFOP
'    rsFOP.UpdateBatch
'    rsFOP.Close
'    Set rsFOP = Nothing
'    cmdRefresh_Click
'
'    For intI = 1 To lvwDespatch.ListItems.Count
'        If lvwDespatch.ListItems(intI) = DespatchID Then
'            lvwDespatch.ListItems(intI).Selected = True
'            Exit For
'        End If
'    Next intI
'End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdRefresh_Click()
   GetDespatchList False
End Sub

Private Sub cmdRefreshList_Click(Index As Integer)
   Select Case Index
      Case 0
         GetDespatchList True
      Case 1
         GetVisaList True
   End Select
End Sub

Private Sub cmdReset_Click(Index As Integer)
   Dim i As Integer
   Dim intCount As Integer
   Dim strDespID As String
   Dim strVisaID As String
   Dim strSQL As String
   
   If chkSelLine(Index).value = 0 Then Exit Sub
   If lvwJobList(Index).ListItems.Count = 0 Then Exit Sub
   intCount = 0
   frmWait.Show
   If chkSelLine(Index).value = 1 Then
      For i = 1 To lvwJobList(Index).ListItems.Count
         With lvwJobList(Index).ListItems(i)
         If .Selected Then
            intCount = intCount + 1
            If Index = 0 Then
               strDespID = strDespID & IIf(strDespID = "", "", " ,") & .Text
            Else
               strVisaID = strVisaID & IIf(strVisaID = "", "'", " ,'") & .Text & "'"
            End If
         End If
         End With
      Next
   End If
   
   If strDespID = "" And strVisaID = "" Then
      Unload frmWait
      Exit Sub
   End If
   If strDespID <> "" Then
      strSQL = "Update TblDespatch Set UpdateDI = 0 "
      strSQL = strSQL & ", PNR = ''"
      strSQL = strSQL & "where DespatchId in (" & strDespID & ")"
      gdbDespatch.Execute strSQL
   End If
   If strVisaID <> "" Then
      strSQL = "Update TblRaiseReq Set UpdateDI = 0 "
      strSQL = strSQL & ", PNR = ''"
      strSQL = strSQL & "where JobID in (" & strVisaID & ")"
      gdbDespatch.Execute strSQL
   End If
   cmdRefreshList_Click (Index)
   Unload frmWait
End Sub

Private Sub cmdUpdate_Click()
   Dim i As Integer
   Dim curComm As Currency
   Dim sngGST As Currency
   Dim strTktNum As String
   Dim intCount As Integer
   Dim strDespID As String
   Dim strVisaID As String
   Dim strSQL As String
   Dim strTUR As String
   Dim strFF4041 As String
   Dim strMsg As String
   'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
   Dim strFOP As String
   Dim strVisaFOP As String
   Dim blnAbsorbMer As Boolean
   datTouchEnd = Now
   If chkDesp.value = 0 And chkVisa.value = 0 Then Exit Sub
   If lvwDespatch.ListItems.Count = 0 And lvwVisa.ListItems.Count = 0 Then Exit Sub
   intCount = 0
   
   frmWait.Show
   Set gobjPNR = New CWT_GalileoPNR3.PNR
   gobjPNR.loadPNR
   If gobjPNR.RecLoc = "" Then
      Unload frmWait
      'MsgBox "NO B.F. TO DISPLAY - RETRIEVE FIRST"
      strMsg = "NO B.F. TO DISPLAY - RETRIEVE FIRST"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - PNR Required"
      Exit Sub
   End If
   'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
   strFOP = ""
   strVisaFOP = ""
   blnAbsorbMer = False
   If cmbFOPType.Text = "INV" Then
      strFOP = "INV"
   Else
      strFOP = cmbFOPType.Text & "/" & cmbCCType.Text & "/" & txtCCNum.Text & "/" & Format(dtpCCExp.value, "MMYY")
   End If
   
   If chkAbsorbMer.Enabled = True And chkAbsorbMer.value = vbChecked Then
         blnAbsorbMer = True
   Else
         blnAbsorbMer = False
   End If
   
   If cmbVisaFOPType.Text = "INV" Then
      strVisaFOP = "INV"
   Else
      strVisaFOP = cmbVisaFOPType.Text & "/" & cmbVisaCCType.Text & "/" & txtVisaCCNum.Text & "/" & Format(dtpVisaCCExp.value, "MMYY")
   End If
   
   If chkDesp.value = 1 Then
      For i = 1 To lvwDespatch.ListItems.Count
         With lvwDespatch.ListItems(i)
         If .Selected Then
            intCount = intCount + 1
            curComm = .SubItems(4) - .SubItems(5)
            sngGST = fGST(.SubItems(4), sngDespatchPct)
            strTktNum = strDespatchPre & Format(.Text, "0000000")
            strDespID = strDespID & IIf(strDespID = "", "", " ,") & .Text
            strTUR = "DESPATCH-" & .Text & " " & Format(.SubItems(1), "ddmmm")
            strFF4041 = "/FF40-DESPATCH CHARGE/FF41-" & Format(.SubItems(1), "ddmmmyy")
            'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
            If .SubItems(4) > 0 Then
                UpdateGDS "08", "024000", .SubItems(4), curComm, sngDespatchPct, strFOP, _
                       strTktNum, strDespatchDesc, strTUR, blnAbsorbMer, "", strFF4041
            End If
         End If
         End With
      Next
   End If
   
   intCount = 0
   If chkVisa.value = 1 Then
      For i = 1 To lvwVisa.ListItems.Count
         With lvwVisa.ListItems(i)
         If .Selected Then
            intCount = intCount + 1
            curComm = 0
            sngGST = fGST(.SubItems(6), sngVisaCPct)
            strTktNum = strVisaCPre & Format(Mid(.Text, 2), "0000000")
            strVisaID = strVisaID & IIf(strVisaID = "", "'", " ,'") & .Text & "'"
            strTUR = "VISA-" & .Text & " " & .SubItems(2) & " " & Format(.SubItems(1), "ddmmm")
            
            strFF4041 = "/FF40-" & Left(Trim(.SubItems(2)), 20) & "/FF41-"
            If .SubItems(9) = True Then strFF4041 = strFF4041 & "TOURIST"
            If .SubItems(10) = True Then strFF4041 = strFF4041 & "BUSINESS"
            If .SubItems(11) = True Then strFF4041 = strFF4041 & "TRANSIT"
            If .SubItems(12) = True Then strFF4041 = strFF4041 & " " & "SGL" & " "
            If .SubItems(13) = True Then strFF4041 = strFF4041 & " " & "DBL" & " "
            If .SubItems(14) = True Then strFF4041 = strFF4041 & " " & "MUL" & " "
         
            
            If .SubItems(6) > 0 Then
            'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
            UpdateGDS "06", .SubItems(5), .SubItems(6), curComm, sngVisaCPct, strVisaFOP, _
                   strTktNum, strVisaCDesc, "", False, .SubItems(2), strFF4041
            End If
            If .SubItems(7) > 0 Then
               curComm = .SubItems(7)
               sngGST = fGST(.SubItems(7), sngVisaHPct)
               strTktNum = strVisaCPre & Format(Mid(.Text, 2), "0000000")
               strTUR = ""
               If .SubItems(7) > 0 Then
               'UpdateGDS "37", strVisaHVendor, .SubItems(8), curComm, sngVisaHPct, .SubItems(5), _
               '       strTktNum, strVisaHDesc, "", .SubItems(9), .SubItems(2), strFF4041
               'CC 20090706
               'MF always absorb for visa handling fee
               'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
               UpdateGDS "37", strVisaHVendor, .SubItems(7), curComm, sngVisaHPct, strVisaFOP, _
                      strTktNum, strVisaHDesc, "", True, .SubItems(2), strFF4041
               End If
            End If
         End If
         End With
      Next
   End If
   If strDespID = "" And strVisaID = "" Then
      Unload frmWait
      Exit Sub
   End If
   gobjHost.terminalEntry "R.Desp"
   gobjHost.terminalEntry "ER"
   gobjHost.terminalEntry "ER"
   gobjHost.terminalEntry "ER"
   If strDespID <> "" Then
      strSQL = "Update TblDespatch Set UpdateDI = 1 "
      strSQL = strSQL & ", PNR = '" & gobjPNR.RecLoc & "' "
      strSQL = strSQL & "where DespatchId in (" & strDespID & ")"
      gdbDespatch.Execute strSQL
   End If
   If strVisaID <> "" Then
      strSQL = "Update TblRaiseReq Set UpdateDI = 1 "
      strSQL = strSQL & ", PNR = '" & gobjPNR.RecLoc & "' "
      strSQL = strSQL & "where JobID in (" & strVisaID & ")"
      gdbDespatch.Execute strSQL
   End If
   cmdVisaRefresh_Click
   cmdRefresh_Click
   Unload frmWait
   EndLog
   datTouchStart = Now
End Sub

Private Sub ResetJobs()
   Dim i As Integer
   Dim curComm As Currency
   Dim sngGST As Currency
   Dim strTktNum As String
   Dim intCount As Integer
   Dim strDespID As String
   Dim strVisaID As String
   Dim strSQL As String
   Dim strTUR As String
   
   If chkSelLine(0).value = 0 Then Exit Sub
   If lvwJobList(0).ListItems.Count = 0 Then Exit Sub
   intCount = 0
   frmWait.Show
   If chkSelLine(0).value = 1 Then
      For i = 1 To lvwJobList(0).ListItems.Count
         With lvwJobList(0).ListItems(i)
         If .Selected Then
            intCount = intCount + 1
            strDespID = strDespID & IIf(strDespID = "", "", " ,") & .Text
            strVisaID = strVisaID & IIf(strVisaID = "", "'", " ,'") & .Text & "'"
         End If
         End With
      Next
   End If
   
   If strDespID = "" And strVisaID = "" Then
      Unload frmWait
      Exit Sub
   End If
   If strDespID <> "" Then
      strSQL = "Update TblDespatch Set UpdateDI = False where "
      strSQL = strSQL & "DespatchId in (" & strDespID & ")"
      gdbDespatch.Execute strSQL
   End If
   If strVisaID <> "" Then
      strSQL = "Update TblRaiseReq Set UpdateDI = False where "
      strSQL = strSQL & "JobID in (" & strVisaID & ")"
      gdbDespatch.Execute strSQL
   End If
   cmdRefreshList_Click (0)
   Unload frmWait
End Sub

'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
'Private Sub cmdVisaAmend_Click()
'   Dim strSql As String
'   Dim strFOP As String
'   Dim rsFOP As ADODB.Recordset
'   Dim intI As Integer
'   Dim JobID As String
'
'   If ValidVisaCCData = False Then Exit Sub
'
'   JobID = lvwVisa.SelectedItem.Text
'
'   If cmbVisaFOPType.Text = "INV" Then
'      strFOP = "INV"
'   ElseIf cmbVisaFOPType.Text = "CX" Then
'      strFOP = "CX/" & cmbVisaCCType.Text & "/" & txtVisaCCNum & "/" & Format(dtpVisaCCExp, "mmyy")
'   End If
'
'   strSql = "Update tblRaiseReq "
'   strSql = strSql & "Set FOP = '" & strFOP & "' "
'   strSql = strSql & ", AbsorbMer = " & chkVisaAbsorbMer.value
'   strSql = strSql & " Where JOBID = '" & lvwVisa.SelectedItem.Text & "'"
'
'   gdbDespatch.Execute strSql
'
'    strSql = "Select * from tblVisaFOPAmend "
'    strSql = strSql & "where VisaID = '-1'"
'    Set rsFOP = New ADODB.Recordset
'    rsFOP.LockType = adLockBatchOptimistic
'    rsFOP.CursorType = adOpenKeyset
'    rsFOP.Source = strSql
'    Set rsFOP.ActiveConnection = gdbDespatch
'    rsFOP.Open
'    rsFOP.AddNew
'    rsFOP("VisaId") = lvwVisa.SelectedItem.Text
'    rsFOP("FPID") = gobjHost.AgentSine
'    rsFOP("FPPCC") = gobjHost.AgentPCC
'    rsFOP("AmendDate") = Now
'    rsFOP("OldFOP") = lvwVisa.SelectedItem.SubItems(5)
'    rsFOP("NewFOP") = strFOP
'    rsFOP.UpdateBatch
'    rsFOP.Close
'    Set rsFOP = Nothing
'    cmdVisaRefresh_Click
'    For intI = 1 To lvwVisa.ListItems.Count
'        If lvwVisa.ListItems(intI) = JobID Then
'            lvwVisa.ListItems(intI).Selected = True
'            Exit For
'        End If
'    Next intI
'End Sub

Private Sub cmdVisaRefresh_Click()
   GetVisaList False
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim oldParent As Long
   datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   
   sstDespVisa.Tab = 0
   'Preethi - V1.2.2 20110223 - CR32 - Expand Date Range For Visa Jobs
   dtpStartDate.value = DateAdd("d", -30, Date)
   dtpEndDate.value = Date
   'Preethi - V1.2.2 20110223 - CR32 - Expand Date Range For Visa Jobs
   dtpStartSubDate.value = DateAdd("d", -30, Date)
   dtpEndSubDate.value = Date

   GetAgentName
   GetDespatchList False
   GetVisaList False
   GetGSTPct
   chkDesp.value = 0
   
   'CC 20090706
   chkVisaAbsorbMer.Visible = False
   chkVisaAbsorbMer.value = 1
   
   For i = 0 To 1
      'Preethi - V1.2.2 20110223 - CR32 - Expand Date Range For Visa Jobs
      dtpSDate(i).value = DateAdd("d", -30, Date)
      dtpEDate(i).value = Date
      cmdRefreshList_Click (i)
   Next
    datTouchStart = Now
    If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
    StartLog
   
End Sub

Private Sub GetDespatchList(ResetJob As Boolean)
   Dim rsDespatch As ADODB.Recordset
   Dim strSQL As String
   Dim item As ListItem
   Dim i As Integer
   
   If ResetJob = False Then
      lvwDespatch.ListItems.Clear
      lblAgentName = cmbAgentName.Text
   Else
      lvwJobList(0).ListItems.Clear
      lblEmpName(0).Caption = cmbEmpName(0).Text
   End If
   
   If ResetJob = False Then
     'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
      strSQL = "Select DespatchId, DeliveryDate, Attention, CompanyName " ', FOP "
      strSQL = strSQL & ", AbsorbMer "
      strSQL = strSQL & ", BillAmount, OffScheduleAmount "
      strSQL = strSQL & "from tblDespatch where UpdateDI = 0 "
      If cmbAgentName.listindex <> -1 Then
         strSQL = strSQL & " and EmployeeNo = '" & cmbAgentID.Text & "' "
      End If
      strSQL = strSQL & " and deliveryDate >= '" & Format(dtpStartDate.value, CstrdateFormat) & "' "
      strSQL = strSQL & " and deliveryDate <= '" & Format(dtpEndDate.value, CstrdateFormat) & "' "
      strSQL = strSQL & " order by deliverydate"
   Else
     'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
      strSQL = "Select DespatchId, DeliveryDate, Attention, CompanyName " ', FOP "
      strSQL = strSQL & ", AbsorbMer "
      strSQL = strSQL & ", BillAmount, OffScheduleAmount "
      strSQL = strSQL & "from tblDespatch where UpdateDI = 1 "
      If cmbEmpName(0).listindex <> -1 Then
         strSQL = strSQL & " and EmployeeNo = '" & cmbEmpID(0).Text & "' "
      End If
      strSQL = strSQL & " and deliveryDate >= '" & Format(dtpSDate(0).value, CstrdateFormat) & "' "
      strSQL = strSQL & " and deliveryDate <= '" & Format(dtpEDate(0).value, CstrdateFormat) & "' "
      strSQL = strSQL & " order by deliverydate"
   End If
   
   'Set rsDespatch = gdbDespatch.OpenRecordset(strSQL)
   Set rsDespatch = gdbDespatch.Execute(strSQL)
   With rsDespatch
      Do Until .EOF
         If ResetJob = False Then
            Set item = lvwDespatch.ListItems.Add(, , !DespatchID)
         Else
            Set item = lvwJobList(0).ListItems.Add(, , !DespatchID)
         End If
         item.SubItems(1) = !DeliveryDate
         item.SubItems(2) = !Attention
         item.SubItems(3) = !CompanyName
         'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
         'item.SubItems(4) = !FOP & ""
         item.SubItems(4) = !BillAmount
         item.SubItems(5) = !OffScheduleAmount
         item.SubItems(6) = !AbsorbMer
         .MoveNext
      Loop
   End With
   rsDespatch.Close
   Set rsDespatch = Nothing
   If ResetJob = False Then
      For i = 1 To lvwDespatch.ListItems.Count
          lvwDespatch.ListItems(i).Selected = False
      Next
   Else
      For i = 1 To lvwJobList(0).ListItems.Count
          lvwJobList(0).ListItems(i).Selected = False
      Next
   End If
   'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
    If gobjPNR.FOPType = "CC" Then
       cmbFOPType.Text = "CX"
       cmbCCType.Text = gobjPNR.FOP_CCCode
       txtCCNum.Text = gobjPNR.FOP_CCNum
       dtpCCExp.value = gobjPNR.FOP_CCExpireDate
    ElseIf gobjPNR.FOPType = "INV" Then
       cmbFOPType.Text = "INV"
    End If

End Sub

Private Sub GetVisaList(ResetJob As Boolean)
   Dim rsVisa As ADODB.Recordset
   Dim strSQL As String
   Dim item As ListItem
   Dim i As Integer
   
   If ResetJob = False Then
      lvwVisa.ListItems.Clear
      lblVisaAgentName = cmbVisaAgentName.Text
      'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
      strSQL = "Select JobID, LastAmendSubDate, Embassy, Passenger " ', FOP "
      strSQL = strSQL & ", AbsorbMer "
      strSQL = strSQL & ", Company, a.Fees, a.Fees2, VisaHandlingFee, VendorNumber "
      strSQL = strSQL & ", a.Tourist, a.Business, UKTransit, Single, [Double], Multiple "
      strSQL = strSQL & "from TblRaiseReq a, tblEmbassy b where "
      strSQL = strSQL & "a.Embassy = b.Country "
      strSQL = strSQL & "and UpdateDI = 0 "
      'strSql = strSql & "and (Status = 'Completed' "
      'strSql = strSql & "or (a.fees <> 0 or a.fees2 <> 0 ))"
      strSQL = strSQL & "and Status = 'Completed' "
      If cmbAgentName.listindex <> -1 Then
         strSQL = strSQL & " and ConsultantID = '" & cmbVisaAgentID.Text & "' "
      End If
      strSQL = strSQL & " and LastAmendSubDate >= '" & Format(dtpStartSubDate.value, CstrdateFormat) & "' "
      strSQL = strSQL & " and LastAmendSubDate <= '" & Format(dtpEndSubDate.value, CstrdateFormat) & "' "
      strSQL = strSQL & " order by Passenger"
   Else
      lvwJobList(1).ListItems.Clear
      lblEmpName(1) = cmbEmpName(1).Text
      'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
      strSQL = "Select JobID, LastAmendSubDate, Embassy, Passenger " ', FOP "
      strSQL = strSQL & ", AbsorbMer "
      strSQL = strSQL & ", Company, a.Fees, a.Fees2, VisaHandlingFee, VendorNumber "
      strSQL = strSQL & ", a.Tourist, a.Business, UKTransit, Single, [Double], Multiple "
      strSQL = strSQL & "from TblRaiseReq a, tblEmbassy b where "
      strSQL = strSQL & "a.Embassy = b.Country "
      strSQL = strSQL & "and UpdateDI = 1 "
      'strSQL = strSQL & "and Status = 'Completed' "
      If cmbAgentName.listindex <> -1 Then
         strSQL = strSQL & " and ConsultantID = '" & cmbEmpID(1).Text & "' "
      End If
      strSQL = strSQL & " and LastAmendSubDate >= '" & Format(dtpSDate(1).value, CstrdateFormat) & "' "
      strSQL = strSQL & " and LastAmendSubDate <= '" & Format(dtpEDate(1).value, CstrdateFormat) & "' "
      strSQL = strSQL & " order by Passenger"
   End If
   
   Set rsVisa = gdbDespatch.Execute(strSQL)
   With rsVisa
      Do Until .EOF
         If ResetJob = False Then
            Set item = lvwVisa.ListItems.Add(, , !JobID)
         Else
            Set item = lvwJobList(1).ListItems.Add(, , !JobID)
         End If
         item.SubItems(1) = !LastAmendSubDate
         item.SubItems(2) = !Embassy
         item.SubItems(3) = !Passenger
         item.SubItems(4) = !Company & ""
         'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
         'item.SubItems(5) = !FOP & ""
         item.SubItems(5) = !VendorNumber & ""
         item.SubItems(6) = !Fees + !Fees2
         item.SubItems(7) = IIf(!VisaHandlingFee & "" = "", 0, !VisaHandlingFee)
         'CC 20090706
         item.SubItems(8) = True    '!AbsorbMer
         item.SubItems(9) = !Tourist
         item.SubItems(10) = !Business
         item.SubItems(11) = !UKTransit
         item.SubItems(12) = !Single
         item.SubItems(13) = !Double
         item.SubItems(14) = !Multiple
         .MoveNext
      Loop
   End With
   rsVisa.Close
   Set rsVisa = Nothing
   If ResetJob = False Then
      For i = 1 To lvwVisa.ListItems.Count
          lvwVisa.ListItems(i).Selected = False
      Next
   Else
      For i = 1 To lvwJobList(1).ListItems.Count
          lvwJobList(1).ListItems(i).Selected = False
      Next
   End If
   
   'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
    If gobjPNR.FOPType = "CC" Then
       cmbVisaFOPType.Text = "CX"
       cmbVisaCCType.Text = gobjPNR.FOP_CCCode
       txtVisaCCNum.Text = gobjPNR.FOP_CCNum
       dtpVisaCCExp.value = gobjPNR.FOP_CCExpireDate
    ElseIf gobjPNR.FOPType = "INV" Then
       cmbVisaFOPType.Text = "INV"
    End If
    
End Sub

Private Sub GetAgentName()
   Dim rsDespatch As ADODB.Recordset
   Dim strSQL As String
   Dim i As Integer
   Dim j As Integer
   Dim intAgent As Integer
   
   strSQL = "Select EmployeeName, EmployeeNo, FPID, FPPCC from tblEmployee "
   strSQL = strSQL & " Order by EmployeeName"
   Set rsDespatch = gdbDespatch.Execute(strSQL)
   
   cmbAgentName.Clear
   cmbAgentID.Clear
   cmbVisaAgentName.Clear
   cmbVisaAgentID.Clear
   For j = 0 To 1
      cmbEmpName(j).Clear
      cmbEmpID(j).Clear
   Next
   i = -1
   With rsDespatch
      Do Until .EOF
         cmbAgentName.AddItem !EmployeeName
         cmbAgentID.AddItem !EmployeeNo
         cmbVisaAgentName.AddItem !EmployeeName
         cmbVisaAgentID.AddItem !EmployeeNo
         For j = 0 To 1
             cmbEmpName(j).AddItem !EmployeeName
             cmbEmpID(j).AddItem !EmployeeNo
         Next
         i = i + 1
         'If !FPID = gobjPNR.Agent And !FPPCC = gobjPNR.PCCOwner Then
         If !FPID = gobjHost.AgentSine And !FPPCC = gobjHost.AgentPCC Then
            intAgent = i
         End If
         .MoveNext
      Loop
      cmbAgentName.listindex = intAgent
      cmbVisaAgentName.listindex = intAgent
      For j = 0 To 1
         cmbEmpName(j).listindex = intAgent
         cmbEmpID(j).listindex = intAgent
      Next
   End With
   rsDespatch.Close
   Set rsDespatch = Nothing
End Sub

Private Sub lvwDespatch_ItemClick(ByVal item As MSComctlLib.ListItem)
   Dim strFOP() As String
   Dim strMsg As String
   
   If lvwDespatch.ListItems.Count = 0 Then Exit Sub
   'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
'   If lvwDespatch.SelectedItem.SubItems(4) = "" Then
'      lvwDespatch.ListItems(lvwDespatch.SelectedItem.Index).Selected = False
'      'MsgBox "Please update FOP."
'      strMsg = "Please update FOP."
'      modMsgBox.OKMsg = "OK"
'      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
'      Exit Sub
'   Else
'      strFOP = Split(lvwDespatch.SelectedItem.SubItems(4), "/")
'      If strFOP(0) = "INV" Or strFOP(0) = "CX" Then
'         cmbFOPType.Text = strFOP(0)
'         If UBound(strFOP) = 3 Then
'            cmbCCType.Text = strFOP(1)
'            txtCCNum = strFOP(2)
'            dtpCCExp.value = "01/" & MMM(Left(strFOP(3), 2)) & "/" & Right(strFOP(3), 2)
            If lvwDespatch.SelectedItem.SubItems(6) = True Then
               chkAbsorbMer.value = 1
            Else
                chkAbsorbMer.value = 0
             End If
            AbsorbMer
'         End If
'      Else
'         cmbFOPType.listindex = -1
'      End If
      
   'End If
End Sub

Private Sub lvwVisa_ItemClick(ByVal item As MSComctlLib.ListItem)
 Dim strFOP() As String
 Dim strMsg As String
 
 If lvwVisa.ListItems.Count = 0 Then Exit Sub
 'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
'   If lvwVisa.SelectedItem.SubItems(5) = "" Then
'      lvwVisa.ListItems(lvwVisa.SelectedItem.Index).Selected = False
'      'MsgBox "Please update FOP."
'      strMsg = "Please update FOP."
'      modMsgBox.OKMsg = "OK"
'      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
'      Exit Sub
'   Else
'      strFOP = Split(lvwVisa.SelectedItem.SubItems(5), "/")
'      If strFOP(0) = "INV" Or strFOP(0) = "CX" Then
'         cmbVisaFOPType.Text = strFOP(0)
'         If UBound(strFOP) = 3 Then
'            cmbVisaCCType.Text = strFOP(1)
'            txtVisaCCNum = strFOP(2)
'            dtpVisaCCExp.value = "01/" & MMM(Left(strFOP(3), 2)) & "/" & Right(strFOP(3), 2)
            If lvwVisa.SelectedItem.SubItems(8) = True Then
               chkVisaAbsorbMer.value = 1
            Else
                chkVisaAbsorbMer.value = 0
             End If
            VisaAbsorbMer
'         End If
'      Else
'         cmbVisaFOPType.listindex = -1
'      End If
   'End If
End Sub

Public Sub UpdateGDS(ProductCode As String, VendorCode As String, SF As Currency, _
    Commission As Currency, GSTPer As Single, FOP As String, TktNum As String, _
    Desc As String, TUR As String, AbsorbMer As Boolean, Optional Country As String, Optional strFF4041 As String)
Dim strTemp As String
Dim strTemp2 As String
Dim strMSLine As String
Dim strFOP() As String
Dim strFF() As String
Dim lngC As Long
Dim lngLen As Long
Dim strPNR As String
Dim mbytFFNum As Integer
Dim bolIsCC As Boolean
Dim strGST As String
Dim strMIData As String
Dim strSegNum As String
Dim ServiceDate As Date
Dim strCountry As String

Dim dblSF As Double
Dim dblGST As Double
Dim lngMer As Long

If UCase(FOP) = "INV" Then
   lngMer = 0
'ElseIf UCase(Left(FOP, 13)) = "CX/DC/3644033" Then
    'lngMer = 0
 'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
 ElseIf UCase(Left(FOP, 2)) = "CX" And _
    IsTMPCard(UCase(Mid(FOP, 4, 2)), UCase(Mid(FOP, 7))) Then
        lngMer = 0
Else
   If AbsorbMer = False Then
      lngMer = fCurrRound(SF * 0.02, gstrAgcyCurrCode, "UP")
   Else
      lngMer = 0
   End If
End If
'dblSF = SF + (SF * mGSTPer) + lngMer
dblSF = SF + (SF * GSTPer * 0.01) + lngMer
dblSF = Format(dblSF / (1 + (GSTPer * 0.01)), "0.00")
dblGST = Format(dblSF * GSTPer * 0.01, "0.00")
   


'If AbsorbMer = False Then
'   SF = fCurrRound(SF * 0.02, gstrAgcyCurrCode, "UP") + SF
'End If
If ProductCode = "06" Then
    Commission = lngMer
'JY – V1.2.6 20110909 – CR58 - Remove despatch cost
'ElseIf ProductCode = "37" Then
ElseIf ProductCode = "37" Or ProductCode = "08" Then
    Commission = dblSF
End If

'ServiceDate = DateAdd("m", 6, Date)
    'ZhiSam - V1.2.23 20130911 - CR-289 - E-Invoice: Due and Paid Lines Date
    If bfunctCheckRTLine = True Then
        ServiceDate = dtfunctRTDate
    Else
        ServiceDate = DateAdd("d", 90, Date)
    End If
    
strMSLine = "/PC" & ProductCode _
    & "/V" & VendorCode _
    & "/S" & Format(dblSF, gstrAgcyCurrFormat) _
    & "/SF" & Format(dblSF, gstrAgcyCurrFormat) _
    & "/C" & Format(Commission, gstrAgcyCurrFormat)

If GSTPer > 0 Then
   strMSLine = strMSLine & "/G" & Format(dblGST, gstrAgcyCurrFormat)
   strGST = Format(dblGST, gstrAgcyCurrFormat)
End If

 strFOP = Split(FOP, "/")
    'Added on 25/08/04: Check if CC payment
    bolIsCC = False
    If strFOP(0) = "CC" Or Left(strFOP(0), 2) = "CX" Then
        'If strFOP(1) = "DC" And Left(strFOP(2), 7) = "3644033" Then
         'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
        If IsTMPCard(strFOP(1), strFOP(2)) Then
            bolIsCC = False
        Else
            bolIsCC = True
        End If
    End If

    strTemp = ""
    
    If bolIsCC = True Then
       
        Select Case strFOP(0)
        Case "CX"
            Select Case strFOP(1)
                Case "AX"
                    strFOP(0) = strFOP(0) & "2"
                Case "DC"
                    strFOP(0) = strFOP(0) & "3"
                Case "VI", "CA"
                    strFOP(0) = strFOP(0) & "4"
                Case "TP"
                    strFOP(0) = strFOP(0) & "5"
            End Select
            strTemp = "/F" & strFOP(0) & "/CCN" & strFOP(1) & strFOP(2) & "EXP" & strFOP(3) _
                    & "/D" & Format((dblSF + dblGST), gstrAgcyCurrFormat)
        Case "CC"
            strTemp = "/FCC/CCN" & strFOP(2) & strFOP(3) _
                    & "/D" & Format((dblSF + dblGST), gstrAgcyCurrFormat)
        End Select
    Else
      
        strTemp = strTemp & "/FS"
    End If

    strMSLine = strMSLine & strTemp
    
    'strMIData = ""
    'If freefields <> "" Then
    '    strFF = Split(freefields, "/")
    '    For lngC = LBound(strFF) To UBound(strFF)
    '        strMIData = strMIData & "/FF" & strFF(lngC)
    '    Next
    'End If
    ''tktlen
     'Preethi - V1.1.1 20100906 - CR17 - FF47 for Despatch and Visa
     
    strMSLine = strMSLine & strMIData & IIf(TktNum = "", "", "/TK" & Format(TktNum, "0000000000")) & strFF4041 & "/FF47-CWT"
    
    
'This routine will make sure that entry does not exceed max char and will add the least amount of lines to PNR based on Max length of 45 char after FT-
    
strTemp = "DI./0"
lngC = 0
Do Until Len(strMSLine) = 0
    If Len(strMSLine) <= 42 Then
        lngLen = Len(strMSLine)
    Else
        lngLen = InstrLast(Left(strMSLine, 42), "/") - 1
    End If
    
    strTemp = strTemp & IIf(lngC = 0, "+DI.FT-MS", "+DI.FT-MSX") & Left(strMSLine, lngLen)
    strMSLine = Mid(strMSLine, lngLen + 1)
    lngC = lngC + 1
Loop

strTemp = strTemp & "+DI.FT-MSX/FF " & IIf(Country = "", "", Country & " ") & Desc

gobjHost.terminalEntry strTemp

If TUR <> "" Then
   strTemp = "0TURZZBK1" & gstrAgcyCityCode & Format(ServiceDate, "ddmmm") & "-"
   strTemp = strTemp & TUR
   gobjHost.terminalEntry strTemp
End If
'        strTemp = "0TURZZBK1" & gstrAgcyCityCode & Format(.ServiceDate, "ddmmm") & "-"      '& .DescriptionLine(1)
'        For lngC = 1 To .DescriptionLinesCount
'            strTemp2 = strTemp2 & "*" & .DescriptionLine(lngC)
'        Next
'        'added on 10/12: split TUR line for invoice display
'
'        lngC = 0
'        Do Until Len(strTemp2) = 0
'           If Len(strTemp2) <= 42 Then
'              lngLen = Len(strTemp2)
'           Else
'              lngLen = InstrLast(Left(strTemp2, 42), " ") - 1
'              If lngLen <= 0 Then lngLen = 42
'           End If
'
'           gobjHost.TerminalEntry UCase(strTemp & Left(strTemp2, lngLen))
'           strTemp2 = Mid(strTemp2, lngLen + 1)
'           lngC = lngC + 1
'        Loop

'Added on 4/7/2005: remove &,(,),' char
'If InStr(Country, "(") > 0 Then
'    strCountry = strCountry & Replace(Country, "(", "")
'ElseIf InStr(strCountry, ")") > 0 Then
'    strCountry = strCountry & Replace(strCountry, ")", "")
'ElseIf InStr(strCountry, "&") > 0 Then
'    strCountry = strCountry & Replace(strCountry, "&", "")
'ElseIf InStr(strCountry, "'") > 0 Then
'    strCountry = strCountry & Replace(strCountry, "'", "")
'End If

Country = Replace(Country, "(", "")
Country = Replace(Country, ")", "")
Country = Replace(Country, "&", "")
Country = Replace(Country, "'", "")
strCountry = Country

    strTemp = "RD.T/" & Format(ServiceDate, "ddmmm") & "*" & IIf(Country = "", "", strCountry & "-") & Desc & "*" & Format(dblSF, gstrAgcyCurrFormat)

gobjHost.terminalEntry strTemp

If GSTPer <> 0 Then
   'strtemp = "RD.T/" & Format(ServiceDate, "ddmmm") & "*5 PERCENT GST*" & strGST
   strTemp = "RD.T/" & Format(ServiceDate, "ddmmm") & "*" & CStr(GSTPer) & " PERCENT GST*" & strGST
   gobjHost.terminalEntry strTemp
End If

If bolIsCC Then
    strTemp = "RP.T/" & Format(ServiceDate, "ddmmm") & "*" & strFOP(1) & "XXXXXXXXXXX" & Right(strFOP(2), 4) & "*" & Format(dblSF, gstrAgcyCurrFormat)
    gobjHost.terminalEntry strTemp
    
    If GSTPer <> 0 Then
       strTemp = "RP.T/" & Format(ServiceDate, "ddmmm") & "*" & strFOP(1) & "XXXXXXXXXXX" & Right(strFOP(2), 4) & "*" & strGST
       gobjHost.terminalEntry strTemp
    End If
    
End If

'gobjHost.TerminalEntry "NP.SS*VBIXO"

'gobjHost.TerminalEntry "R.XO"
'gobjHost.TerminalEntry "ER"
'gobjHost.TerminalEntry "ER"
'strPNR = gobjPNR.RecLoc
''gobjHost.TerminalEntry "TKPDID"
'gobjHost.TerminalEntry "*" & strPNR
'End With
 
'Added on 14/10/04: add to VBI log table
'Timer
'Call pAddToVBILog(gobjPNR.RecLoc, "Other Services", StartTime)
 
End Sub

Private Function fGST(TotalCharge As Single, sngPct As Single) As Single
Dim sngAmt As Single

sngAmt = sngPct * TotalCharge * 0.01
'fGST = fCurrRound(sngAmt, gstrAgcyCurrCode, "UP")
fGST = Format(sngAmt, "0.00")

End Function

Private Sub GetGSTPct()
   Dim rsPC As ADODB.Recordset
   
   'Set rsPC = gdbTPro.OpenRecordset("tblProductCodes")
   'Set rsPC = gdbConn.Execute("tblProductCodes")
   'Set rsPC = New ADODB.Recordset
   'rsPC.Open "select * from tblProductCodes", gdbConn, adOpenDynamic, adLockBatchOptimistic
   
   'rsPC.Index = "ProductCode"
   'rsPC.Seek "=", "06"
   Set rsPC = gdbConn.Execute("select * from tblProductCodes where ProductCode='06'")
   If rsPC.EOF = False Then
      sngVisaCPct = rsPC!GST
      strVisaCPre = rsPC!TktPrefix
      strVisaCDesc = rsPC!Description
   End If
   
   'rsPC.Seek "=", "08"
   Set rsPC = gdbConn.Execute("select * from tblProductCodes where ProductCode='08'")
   If rsPC.EOF = False Then
      sngDespatchPct = rsPC!GST
      strDespatchPre = rsPC!TktPrefix
      strDespatchDesc = rsPC!Description
   End If
   'rsPC.Seek "=", "37"
   Set rsPC = gdbConn.Execute("select * from tblProductCodes where ProductCode='37'")
   If rsPC.EOF = False Then
      sngVisaHPct = rsPC!GST
      strVisaHPre = rsPC!TktPrefix
      strVisaHDesc = rsPC!Description
   End If
   rsPC.Close
   Set rsPC = Nothing
End Sub



Private Sub sstDespVisa_Click(PreviousTab As Integer)
   If sstDespVisa.Tab = 2 Or sstDespVisa.Tab = 3 Then
      cmdUpdate.Enabled = False
   Else
      cmdUpdate.Enabled = True
   End If
End Sub

Private Sub txtCCNum_LostFocus()
AbsorbMer
End Sub

Private Sub AbsorbMer()
If cmbFOPType.Text = "INV" Then
   chkAbsorbMer.value = 0
   chkAbsorbMer.Enabled = False
'ElseIf cmbCCType.Text = "DC" And Left(txtCCNum, 7) = "3644033" Then
 'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
ElseIf IsTMPCard(cmbCCType.Text, txtCCNum) Then
   chkAbsorbMer.value = 0
   chkAbsorbMer.Enabled = False
Else
   chkAbsorbMer.Enabled = True
End If
End Sub

Private Sub VisaAbsorbMer()
'CC 20090706
'If cmbVisaFOPType.Text = "INV" Then
'   chkVisaAbsorbMer.value = 0
'   chkVisaAbsorbMer.Enabled = False
'ElseIf cmbVisaCCType.Text = "DC" And Left(txtVisaCCNum, 7) = "3644033" Then
'   chkVisaAbsorbMer.value = 0
'   chkVisaAbsorbMer.Enabled = False
'Else
'   chkVisaAbsorbMer.Enabled = True
'End If
End Sub

Private Sub txtVisaCCNum_LostFocus()
VisaAbsorbMer
End Sub
'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
'Private Function ValidCCData() As Boolean
'Dim strMsg As String
'
'If cmbFOPType.Text = "" Then strMsg = strMsg & "Need form of payment..." & Chr(13)
'If cmbFOPType.Text = "CX" Then
'    If cmbCCType.Text = "" Then strMsg = strMsg & "Need valid credit vendor code..." & Chr(13)
'    If txtCCNum = "" Then strMsg = strMsg & "Need valid credit card number..." & Chr(13)
'    If LastDate(dtpCCExp.value) < Date Then strMsg = strMsg & "Need valid expiration date..." & Chr(13)
'    If (txtCCNum.Text <> "" And cmbCCType.Text <> "") Then If ValidCCNum(cmbCCType.Text, txtCCNum.Text) = False Then strMsg = strMsg & "Credit card number is invalid or wrong card vendor selected ..." & Chr(13)
'End If
'
'If strMsg = "" Then
'   ValidCCData = True
'Else
'   'MsgBox strMsg
'   modMsgBox.OKMsg = "OK"
'   modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
'   ValidCCData = False
'End If
'End Function
'Preethi - V1.2.10 20120320 - CR 142 - Despatch Visa Job - Remove FOP Box
'Private Function ValidVisaCCData() As Boolean
'Dim strMsg As String
'
'If cmbVisaFOPType.Text = "" Then strMsg = strMsg & "Need form of payment..." & Chr(13)
'If cmbVisaFOPType.Text = "CX" Then
'    If cmbVisaCCType.Text = "" Then strMsg = strMsg & "Need valid credit vendor code..." & Chr(13)
'    If txtVisaCCNum = "" Then strMsg = strMsg & "Need valid credit card number..." & Chr(13)
'    If LastDate(dtpVisaCCExp.value) < Date Then strMsg = strMsg & "Need valid expiration date..." & Chr(13)
'    If (txtVisaCCNum.Text <> "" And cmbVisaCCType.Text <> "") Then If ValidCCNum(cmbVisaCCType.Text, txtVisaCCNum.Text) = False Then strMsg = strMsg & "Credit card number is invalid or wrong card vendor selected ..." & Chr(13)
'End If
'
'If strMsg = "" Then
'   ValidVisaCCData = True
'Else
'   'MsgBox strMsg
'   modMsgBox.OKMsg = "OK"
'   modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
'   ValidVisaCCData = False
'End If
'End Function
Private Sub EndLog()
             
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModDesUpdate, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datTouchStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModDesUpdate, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
End Sub
Private Sub StartLog()

       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModDesUpdate, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datTouchStart, datFormLoadStart

End Sub
