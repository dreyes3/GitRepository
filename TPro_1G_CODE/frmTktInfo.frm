VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTktInfo 
   Caption         =   "CWT TravelPro - Ticketing"
   ClientHeight    =   8955
   ClientLeft      =   165
   ClientTop       =   570
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTktInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   11400
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
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
      Left            =   360
      TabIndex        =   25
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton cmdTicketAqua 
      Caption         =   "Ticket/Itin/ Inv(Aqua)"
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
      Left            =   5040
      TabIndex        =   23
      Top             =   10200
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdTicket 
      Caption         =   "Ticket/Itin"
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
      Left            =   3000
      TabIndex        =   22
      Top             =   10080
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdInvoice 
      Caption         =   "Invoice Only"
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
      Left            =   5040
      TabIndex        =   19
      Top             =   9840
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   2400
      TabIndex        =   16
      Top             =   7800
      Width           =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10920
      Top             =   10080
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Ticket/Itin/Inv"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   9840
      Visible         =   0   'False
      Width           =   1800
   End
   Begin TabDlg.SSTab sstTicketing 
      Height          =   8415
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Accounting/ Printer/ Formats/ Segments"
      TabPicture(0)   =   "frmTktInfo.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblGASRTtotal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFFTotal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDueTotal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblPaidTotal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCXTotal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label19"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label22"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblTaxTotal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDITaxTotal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lvwFiledFares"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lvwGASalesRec"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lvwPaid"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lvwDue"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdAddCXPaid"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdDelete"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      Begin VB.Frame Frame5 
         Caption         =   "Ticket Selection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4440
         TabIndex        =   55
         Top             =   7080
         Width           =   6495
         Begin VB.CheckBox chkAll 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   61
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkTkt 
            Caption         =   "Ticket"
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
            Left            =   1560
            TabIndex        =   60
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkItinerary 
            Caption         =   "Itinerary"
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
            Left            =   2640
            TabIndex        =   59
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkInvoice 
            Caption         =   "Invoice"
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
            Left            =   4080
            TabIndex        =   58
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkMir 
            Caption         =   "Mir"
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
            Left            =   5400
            TabIndex        =   57
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkAqua 
            Caption         =   "Aqua Invoice and Mir"
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
            Left            =   3240
            TabIndex        =   56
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fare File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   4440
         TabIndex        =   53
         Top             =   5280
         Width           =   6405
         Begin MSComctlLib.TreeView tvStoredFare 
            Height          =   1320
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   6120
            _ExtentX        =   10795
            _ExtentY        =   2328
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Segments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4440
         TabIndex        =   51
         Top             =   3960
         Width           =   6405
         Begin VB.ListBox lstSegments 
            Height          =   690
            Left            =   120
            MultiSelect     =   1  'Simple
            TabIndex        =   52
            Top             =   360
            Width           =   5925
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Passengers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   4440
         TabIndex        =   49
         Top             =   2880
         Width           =   6405
         Begin VB.ListBox lstPax 
            Height          =   480
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   6135
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Devices"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   4095
         Begin VB.ComboBox cmbSTP 
            Height          =   330
            Left            =   1560
            TabIndex        =   40
            Text            =   "cmbSTP"
            Top             =   2160
            Width           =   1275
         End
         Begin VB.ComboBox cmbSTPLoc 
            Height          =   330
            Left            =   1560
            TabIndex        =   39
            Text            =   "cmbSTPLoc"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox chkInvType 
            Caption         =   "Check1"
            Height          =   255
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtIATA 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2040
            TabIndex        =   37
            Top             =   5520
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.ComboBox cmbMIR 
            Height          =   330
            Left            =   1560
            TabIndex        =   36
            Text            =   "cmbMIR"
            Top             =   1800
            Width           =   1275
         End
         Begin VB.ComboBox cmbITIN 
            Height          =   330
            Left            =   1560
            TabIndex        =   35
            Text            =   "cmbITIN"
            Top             =   720
            Width           =   1275
         End
         Begin VB.ComboBox cmbTKT 
            Height          =   330
            Left            =   1560
            TabIndex        =   34
            Text            =   "cmbTKT"
            Top             =   1080
            Width           =   1275
         End
         Begin VB.ComboBox cmbINV 
            Height          =   330
            Left            =   1560
            TabIndex        =   33
            Text            =   "cmbINV"
            Top             =   1440
            Width           =   1275
         End
         Begin VB.CheckBox chkTktType 
            Caption         =   "Check1"
            Height          =   255
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1200
            Width           =   735
         End
         Begin VB.ComboBox cmbITINDYO 
            Height          =   330
            Left            =   1440
            TabIndex        =   31
            Text            =   "cmbITINDYO"
            Top             =   5520
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.ComboBox cmbINVDYO 
            Height          =   330
            Left            =   960
            TabIndex        =   30
            Text            =   "cmbINVDYO"
            Top             =   5640
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.TextBox txtCRInv 
            Height          =   285
            Left            =   0
            TabIndex        =   29
            Top             =   5520
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtCRItin 
            Height          =   285
            Left            =   480
            TabIndex        =   28
            Top             =   5520
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox chkItinType 
            Height          =   255
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Location:"
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
            Left            =   630
            TabIndex        =   48
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblSTPlabel 
            Alignment       =   1  'Right Justify
            Caption         =   "STP:"
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
            Left            =   630
            TabIndex        =   47
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "MIR:"
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
            Left            =   630
            TabIndex        =   46
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Itinerary:"
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
            Left            =   630
            TabIndex        =   45
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Ticket:"
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
            Left            =   630
            TabIndex        =   44
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Invoice:"
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
            Left            =   630
            TabIndex        =   43
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Processing Status:"
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
            Left            =   240
            TabIndex        =   42
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label lblStatus 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   240
            TabIndex        =   41
            Top             =   3000
            Width           =   3615
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         Left            =   7200
         TabIndex        =   17
         Top             =   1440
         Width           =   1515
      End
      Begin VB.CommandButton cmdAddCXPaid 
         Caption         =   "Add C&X Paid Line"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSComctlLib.ListView lvwDue 
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   450
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwPaid 
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   450
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwGASalesRec 
         Height          =   1335
         Left            =   180
         TabIndex        =   5
         Top             =   660
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Line"
            Object.Width           =   1552
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PC"
            Object.Width           =   2002
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Vendor"
            Object.Width           =   2002
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cost"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Sell"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwFiledFares 
         Height          =   255
         Left            =   5640
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   450
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FF"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "FOP"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Base"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Tax"
            Object.Width           =   2275
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblDITaxTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   6000
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblTaxTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   5760
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10980
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label22 
         Caption         =   "Select GA Sales Record and/or Due Lines and"
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
         Left            =   5760
         TabIndex        =   18
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label Label19 
         Caption         =   "CC / CX Total:"
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
         Left            =   1800
         TabIndex        =   14
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label lblCXTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label lblPaidTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   9840
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblDueTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   9960
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblFFTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   9960
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblGASRTtotal 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label14 
         Caption         =   "Total:"
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
         Left            =   2640
         TabIndex        =   8
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "GA Sales Records (DI/MS lines)"
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
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   3255
      End
   End
   Begin MSComctlLib.ListView lsvCPG 
      Height          =   1215
      Left            =   240
      TabIndex        =   24
      Top             =   9000
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CCVendor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CCNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ExpiryDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ProcessInd"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "StatusCode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "TransID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Error"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ReceiptNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "QSICode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "QSIDesc"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblPNR 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmTktInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobjTE As HostAccess.TerminalEmulation
Dim mIsEINVClient As Boolean
Dim strNow As String
Dim bolTktItinInv As Boolean
Dim bolInvOnly As Boolean
Dim bolTktItin As Boolean
'Timer
Dim startTime As Date
Dim SysStart As Date
Dim invStart As Date
Dim bolFirst As Boolean
Dim bolSent As Boolean
Dim strDoc As String

Dim mstrACPrinter As String
Dim hInternetSession As Long
Dim hInternetConnect As Long
Dim hHttpOpenRequest As Long
Dim mstrTransError As String
Dim intDefaultItin As Integer
Dim strITINDefault As String
Private Const OK As String = 1

Private Const cSELFORECOLOR As Long = &HFFFFFF
Private Const cSELBACKCOLOR As Long = &HFF0000

Private Enum eCommand
    [Select] = 0
    '[Select All] = 1
    [Clear] = 1
    '[Clear All] = 3
    '[Transfer] = 4
    '[Toggle] = 5
    '[Toggle Selection] = 6
End Enum

'
'## TreeView nodes can have different fore & back colors plus a Bold state.
'   For Muli-Node selection to work, we need to capture and store this
'   information for each node. We can't use a collection of nodes due to
'   only pointers to objects are stored and not seperate new objects.
'   Therefore a specialised collection is required.
'
'   I haven't used a type'd array due to the overhead of management. Therefore
'   a collection class of variants has been used. I've chosen variants over
'   explicit properties (i.e. NodeKey, ForeColor, BackColor, Bold) to make
'   the class more generic for future projects - any type of
'   variable/object and any number of elements per collection item can stored.
'
'## Tag Element IDs used
Private Enum eTags                      'Private Type tTags
    [Node Key] = 0                      '    lNodeKey   as Long
    [ForeColor] = 1                     '    lForeColor as OLE_COLOR
    [BackColor] = 2                     '    lBackColor as OLE_COLOR
    [Bold] = 3                          '    bBold      as Boolean
End Enum                                'End Type

Private moTags As TreeViewTag     'Private moTags() as tTags
Private moSelNode As MSComctlLib.node

Dim strPax As String
Dim strFF As String
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date



'Private Sub chkSTP_Click()
'   If chkSTP.Value = 1 Then
'      cmbSTP.Enabled = True
'      txtIATA.Enabled = True
      'txtIATA = gstrIATA
      'If Len(cmbSTP.Text) > 0 Then
      '  txtIATA = lsvSTP.ListItems(cmbSTP.ListIndex + 1).SubItems(1)
      'Else
      '  txtIATA = ""
      'End If
'      GetSTPDevice (cmbSTP.Text)
'   Else
'      GetSTPDevice ("HQ")
'      lblSTP = ""
'      cmbSTP.Enabled = False
'      txtIATA.Enabled = False
'      txtIATA = ""
'   End If
'End Sub

Private Sub addPaidLines()
Dim strTemp As String
Dim item As ListItem
Dim dtmNewDate As Date
Dim bolIsCC As Boolean
Dim intX As Integer
'Added on 17/08/04
Dim curTotAirFares As Currency
'Dim colFOPGroup As Collection  -- for sorting
Dim objOthFOP As AirFaresFOP
Dim lngCreditTerms As Long

'Clement
'dtmNewDate = DateAdd("M", 6, Date)

'ZhiSam - V1.2.23 20130911 - CR-289 - E-Invoice: Due and Paid Lines Date
 If bfunctCheckRTLine = True Then
    dtmNewDate = dtfunctRTDate
 Else
    dtmNewDate = DateAdd("d", 90, Date)
 End If

'Modified on 17/08/04: use FOP in DI lines to determine CC payment
curTotAirFares = 0
'Set colFOPGroup = New Collection
Set objOthFOP = New AirFaresFOP

With gobjPNR
For intX = 1 To .AirFaresFOPCount
    'MsgBox "FF#" & CStr(.AirFaresFOP(IntX).FFNumber) & " Type:" & .AirFaresFOP(IntX).FOPType & " Code:" & .AirFaresFOP(IntX).FOP_CCCode & " Number:" & .AirFaresFOP(IntX).FOP_CCNum & " Amt:" & CStr(.AirFaresFOP(IntX).FOPAmount)
    If Left(UCase(.AirFaresFOP(intX).FOPType), 2) = "CC" Or Left(UCase(.AirFaresFOP(intX).FOPType), 2) = "CX" Then
        'If Not (Left(UCase(.AirFaresFOP(intX).FOP_CCCode), 2) = "DC" And _
                Left(UCase(.AirFaresFOP(intX).FOP_CCNum), 7) = "3644033") Then
         'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
        If Not IsTMPCard(Left(UCase(.AirFaresFOP(intX).FOP_CCCode), 2), UCase(.AirFaresFOP(intX).FOP_CCNum)) Then
            curTotAirFares = curTotAirFares + .AirFaresFOP(intX).FOPAmount
            
            'Organize FOP grouping for Air Fares
            'colFOPGroup.Add .AirFaresFOP(intX)
        End If
    End If
Next
'Organize FOP grouping for Other Fares
For intX = 1 To .GASaleRecordCount
    With .GASalesRecord(intX)
    If Left(.FOP, 2) = "CX" Then
        objOthFOP.FOP_CCCode = UCase(Left(.CCNumber, 2))
        objOthFOP.FOP_CCNum = UCase(Mid(.CCNumber, 3))
        objOthFOP.FOPAmount = .SellAmount
        'colFOPGroup.Add objOthFOP
    End If
    End With
Next
'
End With

'For Sorting
'Set objOthFOP = Nothing
'For Each objOthFOP In colFOPGroup
'    MsgBox objOthFOP.FOP_CCCode & "-" & objOthFOP.FOP_CCNum & "-" & CStr(objOthFOP.FOPAmount)
'Next


If CCur(lblCXTotal.Caption) > 0 Or curTotAirFares > 0 Then
   bolIsCC = True
Else
   bolIsCC = False
End If

'bolIsCC = False
'If UCase(gstrAgcyCountryCode) = "HK" Then
'   If UCase(gobjPNR.FOPType) = "CC" Then
'      bolIsCC = True
'   End If
'Else
'   If UCase(gobjPNR.FOPType) = "CC" Then
'      If UCase(gobjPNR.FOP_CCCode) = "DC" And Left(gobjPNR.FOP_CCNum, 7) = "3644033" Then
'         bolIsCC = False
'      Else
'         bolIsCC = True
'      End If
'   End If
'End If
    
    '''strTemp = "RP.T/" & Format(gobjPNR.PaidDue(gobjPNR.PaidDueCount).SegDate, "ddmmm") _
    '''    & "*CREDIT CARD PAYMENT*" & lblCXTotal.Caption
    'Clement
    'If bolIsCC Then
'          strTemp = "RP.T/" & Format(dtmNewDate, "ddmmm") _
'              & "*" & gobjPNR.FOP_CCCode & "XXXXXXXXXXX" & Right(gobjPNR.FOP_CCNum, 4) & "*" & lblCXTotal.Caption
        'Modified on 17/08/04
        'Override on 24/08/04: Request from HKG users: split the paid lines by item
        'strTemp = "RP.T/" & Format(dtmNewDate, "ddmmm") _
        '    & "*" & gobjPNR.AirFaresFOP(1).FOP_CCCode & "XXXXXXXXXXX" & Right(gobjPNR.AirFaresFOP(1).FOP_CCNum, 4) _
        '    & "*" & CStr(CCur(lblCXTotal.Caption) + curTotAirFares)
        'MakeEntry strTemp
    
    
        'strTemp = gobjPNR.PaidDue(gobjPNR.PaidDueCount).SegNum + 1
        'Set item = lvwPaid.ListItems.Add(, , strTemp)
        '                item.SubItems(1) = "*CREDIT CARD PAYMENT-" & gobjPNR.FOP_CCCode & "XXXXXXXXXXX" & Right(gobjPNR.FOP_CCNum, 4) & "*"
        '                item.SubItems(2) = lblDueTotal.Caption
        'lblPaidTotal.Caption = CSng(lblPaidTotal.Caption) + CSng(lblCXTotal.Caption)
    'End If
    
    If UCase(gstrAgcyCountryCode) = "SG" Then
       strTemp = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
                 "*0 PERCENT GST ON NON TAXABLE CHARGE*" & Format(0, gstrAgcyCurrFormat)
       MakeEntry strTemp
    End If
'remove on 14/7/2005: use due-paid to get the due amount
    'If bolIsCC = False Then
    'added checking on 31/03/2005
        If CSng(lblPaidTotal) > CSng(lblDueTotal) Then
            strTemp = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
                     "*INVOICE TOTAL DUE*" & Format(0, gstrAgcyCurrFormat)
        Else
           strTemp = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
                     "*INVOICE TOTAL DUE*" & Format(CSng(lblDueTotal) - CSng(lblPaidTotal), gstrAgcyCurrFormat)
        End If
    'Else
    '   strTemp = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
    '             "*INVOICE TOTAL DUE*" & Format(0, gstrAgcyCurrFormat)
    'End If
    
    MakeEntry strTemp
    
            
            
   ' If UCase(gstrAgcyCountryCode) = "SG" Then
        If (CSng(lblDueTotal) - CSng(lblPaidTotal)) > 0 Then
            lngCreditTerms = gobjPNR.CompInfo.CreditTerms
            If lngCreditTerms > 0 Then
                strTemp = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
                        "*PLEASE MAKE PAYMENT WITHIN " & lngCreditTerms & " DAYS"
                MakeEntry strTemp
                strTemp = "RD.T/" & Format(dtmNewDate, "DDMMM") & "*FROM INVOICE DATE"
                MakeEntry strTemp
            End If
        End If
        'Modified on 30/1/2007
        'If bolIsCC Then
        '   strTemp = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
        '            "*TTL AMOUNT CHARGE TO CREDIT CARD*" & Format(CSng(lblPaidTotal), gstrAgcyCurrFormat)
        '   MakeEntry strTemp
        'End If
        
        If bolIsCC Then
          strTemp = "RD.T/" & Format(dtmNewDate, "DDMMM") & _
                    "*TTL AMOUNT CHARGE TO CREDIT CARD*" & Format(CSng(lblPaidTotal) - DiscountPaidAmt, gstrAgcyCurrFormat)
          MakeEntry strTemp
       End If
 '   End If

'gobjFareQuotes(1).TotAmount

End Sub




Private Sub chkAll_Click()
    If chkAll.value = vbChecked Then
        'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
        If gobjPNR.CompInfo.AquaItin = False Then
            chkItinerary.Enabled = False
            chkItinerary.value = vbChecked
        End If
        chkInvoice.Enabled = False
        chkInvoice.value = vbChecked
        chkMir.Enabled = False
        chkMir.value = vbChecked
        chkTkt.Enabled = False
        chkTkt.value = vbChecked
    Else
        'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
        If gobjPNR.CompInfo.AquaItin = False Then
            chkItinerary.Enabled = True
            chkItinerary.value = vbUnchecked
        End If
        chkInvoice.Enabled = True
        chkInvoice.value = vbUnchecked
        chkMir.Enabled = True
        chkMir.value = vbUnchecked
        chkTkt.Enabled = True
        chkTkt.value = vbUnchecked
    End If
End Sub

Private Sub chkAqua_Click()
    If chkAqua.value = vbChecked Then
        chkInvoice.value = vbUnchecked
        chkInvoice.Enabled = False
        chkMir.value = vbUnchecked
        chkMir.Enabled = False
    Else
        chkInvoice.value = vbUnchecked
        chkInvoice.Enabled = True
        chkMir.value = vbUnchecked
        chkMir.Enabled = True
    End If
End Sub

Private Sub chkInvoice_Click()
    If chkInvoice.value = vbChecked Then
        chkAqua.value = vbUnchecked
        chkAqua.Enabled = False
        chkMir.value = vbChecked
    Else
        chkAqua.value = vbUnchecked
        chkAqua.Enabled = True
        chkMir.value = vbUnchecked
    End If
End Sub

Private Sub chkItinType_Click()

If chkItinType.value = 1 Then
      chkItinType.Caption = "E-Itin"
      cmbITIN.Enabled = False
      cmbITIN.Text = ""
Else
      chkItinType.Caption = "P-Itin"
      cmbITIN.Enabled = True
      cmbITIN.listindex = intDefaultItin
End If

End Sub

Private Sub chkMir_Click()
    If chkMir.value = vbChecked Then
        chkAqua.value = vbUnchecked
        chkAqua.Enabled = False
    Else
        chkAqua.value = vbUnchecked
        chkAqua.Enabled = True
    End If
End Sub

Private Sub chkTktType_Click()
Dim i As Integer
Dim bolFound As Boolean
'If cmbTKT.ListIndex > -1 Then
   If chkTktType.value = 1 Then
      chkTktType.Caption = "E-Tkt"
      For i = 0 To cmbTKT.ListCount - 1
          If cmbTKT.ItemData(i) = 0 Then
             cmbTKT.listindex = i
             cmbSTP.listindex = -1
             bolFound = True
             Exit For
          End If
      Next
   If bolFound <> True Then cmbTKT.listindex = -1
   Else
      chkTktType.Caption = "P-Tkt"
      For i = 0 To cmbTKT.ListCount - 1
          If cmbTKT.ItemData(i) = 1 Then
             cmbTKT.listindex = i
             
             bolFound = True
             Exit For
             
          End If
      Next
    If cmbSTP.ListCount > 0 Then cmbSTP.listindex = 0
    If bolFound <> True Then cmbTKT.listindex = -1
   End If
'End If
End Sub

Private Sub cmbINV_Click()
 Dim intI As Integer
   If cmbINV.listindex > -1 Then
   If cmbINV.ItemData(cmbINV.listindex) = 0 Then 'E Ticke
      For intI = 0 To cmbINVDYO.ListCount - 1
          If cmbINVDYO.ItemData(intI) = 0 Then
             cmbINVDYO.listindex = intI
             Exit For
          End If
      Next
      chkInvType.value = 1
      chkInvType.Caption = "E-Inv"
   Else
      For intI = 0 To cmbINVDYO.ListCount - 1
          If cmbINVDYO.ItemData(intI) = 1 Then
             cmbINVDYO.listindex = intI
             Exit For
          End If
      Next
      chkInvType.value = 0
      chkInvType.Caption = "P-Inv"
   End If
   End If
End Sub

Private Sub cmbSTP_Click()

   Dim i As Integer
'If cmbSTP.Text <> "" Then
If cmbSTP.listindex > -1 Then
   If cmbSTP.ItemData(cmbSTP.listindex) = 0 Then 'E Ticket
      For i = 0 To cmbITINDYO.ListCount - 1
          If cmbITINDYO.ItemData(i) = 0 Then
             cmbITINDYO.listindex = i
             txtCRItin = ""
             Exit For
          End If
      Next
   Else
      For i = 0 To cmbITINDYO.ListCount - 1
          If cmbITINDYO.ItemData(i) = 1 Then
             cmbITINDYO.listindex = i
             If UCase(gstrAgcyCountryCode) = "SG" Then txtCRItin = "1-12"
             Exit For
          End If
      Next
   End If
End If
End Sub

Private Sub cmbSTPLoc_Click()
    'lblSTP = lsvSTP.ListItems(cmbSTP.ListIndex + 1).Text
    'txtIATA = lsvSTP.ListItems(cmbSTP.ListIndex + 1).SubItems(1)
   'lblSTP = lstSTP.List(cmbSTP.ListIndex)
   Call GetAddress(gobjPNR.PCCOwner, cmbSTPLoc.Text)
   If cmbTKT.listindex > -1 Then
    If cmbTKT.ItemData(cmbTKT.listindex) = 0 Then
       chkTktType.value = 1
       chkTktType.Caption = "E-Tkt"
    Else
       chkTktType.value = 0
       chkTktType.Caption = "P-Tkt"
    End If
   Else
      chkTktType.value = 0
      chkTktType.Caption = "P-Tkt"
   End If
   
   SetPE
End Sub

Private Sub cmbTKT_Click()
   Dim i As Integer
   If cmbTKT.listindex > -1 Then
   If cmbTKT.ItemData(cmbTKT.listindex) = 0 Then 'E Ticke
      For i = 0 To cmbITINDYO.ListCount - 1
          If cmbITINDYO.ItemData(i) = 0 Then
             cmbITINDYO.listindex = i
             txtCRItin = ""
             Exit For
          End If
      Next
      chkTktType.value = 1
      chkTktType.Caption = "E-Tkt"
   Else
      For i = 0 To cmbITINDYO.ListCount - 1
          If cmbITINDYO.ItemData(i) = 1 Then
             cmbITINDYO.listindex = i
             If UCase(gstrAgcyCountryCode) = "SG" Then txtCRItin = "1-12"
             Exit For
          End If
      Next
      chkTktType.value = 0
      chkTktType.Caption = "P-Tkt"
   End If
   End If
End Sub

Private Sub cmdCancel_Click()
If fWantToQuit Then
    Unload Me
End If

End Sub

Private Sub cmdConfirm_Click()
Dim intRetries As Integer
Dim strAQQueue As String
Dim strMLQueue As String
Dim strKey As String
Dim strRes As String
Dim strMsg As String
'TKT
'Timer1.Enabled = False
'Timer1.Enabled = True
bolFirst = True
SysStart = Now
datTouchEnd = Now
intRetries = 0
getSelected

If chkAll = vbUnchecked And chkTkt = vbUnchecked And chkItinerary = vbUnchecked And chkInvoice = vbUnchecked And chkMir = vbUnchecked And chkAqua = vbUnchecked Then
    'MsgBox "Please select to issue at least one document!"
    strMsg = "Please select to issue at least one document!"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    Exit Sub
End If

'CC - V1.2.4 20110629 - CR52 HBT Implementation - Remove CompareRDDI logic (check if total amount in Due line is equal to the amount in DI line)
'If UCase(gstrAgcyCountryCode) = "SG" Then
'    If CompareRDDI = False Then Exit Sub
'End If

cmdConfirm.Enabled = False



If chkAll.value = vbChecked Then
     strDoc = "ALL"
     lblStatus = "Ready to issue Ticket/Itinerary"
     If IssueTktItinNew(intRetries, "ALL", chkInvoice.value) = True Then bolTktItinInv = True
     lblStatus = "Ready to issue Invoice/Mir"
     IssueInvoice "ALL"
Else
    If chkTkt.value = vbChecked And chkItinerary.value = vbChecked Then
        strDoc = "ALL"
        lblStatus = "Ready to issue Ticket/Itinerary"
        IssueTktItinNew intRetries, "ALL", chkInvoice.value
    ElseIf chkTkt.value = vbChecked Then
        strDoc = "TKT"
        lblStatus = "Ready to issue Ticket"
        IssueTktItinNew intRetries, "TKT", chkInvoice.value
        
    ElseIf chkItinerary.value = vbChecked Then
        strDoc = "ITIN"
        lblStatus = "Ready to issue Itinerary"
        IssueTktItinNew intRetries, "ITIN", chkInvoice.value
    End If
    
    
    If chkInvoice.value = vbChecked And chkMir.value = vbChecked Then
        strDoc = "ALL"
        lblStatus = "Ready to issue Invoice/Mir"
        IssueInvoice "ALL"
    Else
        If chkInvoice.value = vbChecked Then
            strDoc = "INV"
            lblStatus = "Ready to issue Invoice"
            IssueInvoice "INV"
            bolInvOnly = True
        End If
        
        If chkMir.value = vbChecked Then
            strDoc = "MIR"
            lblStatus = "Ready to issue Mir"
            IssueInvoice "MIR"
        End If
        
        If chkAqua.value = vbChecked Then
            If chkTkt.value = vbChecked Then
                strRes = MakeEntry("*" & gobjPNR.RecLoc)
                If InStr(1, strRes, "FINISH OR IGNORE") <> 0 Or _
                    InStr(1, strRes, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
                    Sleep (1000)
                    MakeEntry "I"
                    Sleep (500)
                    strRes = MakeEntry("*" & gobjPNR.RecLoc)
                End If
                Sleep (1000)
                MakeEntry "IR"
                MakeEntry "IR"
            End If
            lblStatus = "Ready to Queue PNR to Aqua"
            GetAquaQueue strAQQueue, strMLQueue
            strKey = pAddToAQQueueLog
            If strKey <> "" Then
                AddQKeytoNP strKey
            End If
            MakeEntry "NP.TT*INVONLY"
            MakeEntry "NP.TQ*" & strMLQueue
            MakeEntry "R.TPRO TKT+ER"
            MakeEntry "ER"
            strRes = MakeEntry("ER")
            If InStr(1, strRes, "SIMULTANEOUS CHANGES TO BOOKING FILE") <> 0 Then
                MakeEntry "IR"
                MakeEntry "IR"
                MakeEntry "NP.TT*INVONLY"
                MakeEntry "NP.TQ*" & strMLQueue
                MakeEntry "R.TPRO TKT+ER"
                MakeEntry "ER"
            End If
            MakeEntry "QEB/" & strAQQueue
            lblStatus = "PNR Completed queuing to Aqua"
        End If
    
    End If
    
End If
Dim strPath As String
Dim intFile As Integer
strPath = App.Path
strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
intFile = FreeFile()
Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "Selection: " & chkAll.Name & ":" & chkAll.value & "," _
& chkTkt.Name & ":" & chkTkt.value & "," _
& chkItinerary.Name & ":" & chkItinerary.value & "," _
& chkMir.Name & ":" & chkMir.value & "," _
& chkInvoice.Name & ":" & chkInvoice.value & "," _
& chkAqua.Name & ":" & chkAqua.value
Close #intFile

If cmbITIN.Text <> "" Then
    MakeEntry "HMLM" & cmbITIN.Text & "DI"
Else
    MakeEntry "HMLM" & strITINDefault & "DI"
End If
lblStatus = "Process completed"
cmdConfirm.Enabled = True

 'If chkInvoice.value = False And chkMir.value = False Then
 strMsg = "Do you want to redisplay PNR " & lblPNR.Caption & " in focalpoint?"
 modMsgBox.YESMsg = "Yes"
 modMsgBox.NOMsg = "No"
 If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop") = vbNo Then
 'If MsgBox("Do you want to redisplay PNR " & lblPNR.Caption & " in focalpoint?" _
        , vbApplicationModal + vbQuestion + vbYesNo) = vbNo Then
    'exit Sub
    GoTo EndLog
 Else
    pDisplayToFP "*" & gobjPNR.RecLoc
 End If
 
 'End If
EndLog:
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModIssueDoc, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModIssueDoc, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModIssueDoc, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
    
End Sub

Private Sub cmdDelete_Click()
Dim strMsg As String
With lvwGASalesRec
   strMsg = "Are you sure you want to delete the selected GA Sales Record?" & Chr(13) _
        & "(DI lines " & .SelectedItem & ")"
   modMsgBox.YESMsg = "Yes"
   modMsgBox.NOMsg = "No"
   If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop") = vbNo Then Exit Sub
    
    'If MsgBox("Are you sure you want to delete the selected GA Sales Record?" & Chr(13) _
    '    & "(DI lines " & .SelectedItem & ")" _
    '    , vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        
    MakeEntry "DI." & .SelectedItem & "@"
End With

With lvwDue
    strMsg = "Are you sure you want to delete the selected Due Line " _
        & "( " & .SelectedItem & ")?"
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop - Reminder") = vbYes Then
    'If MsgBox("Are you sure you want to delete the selected Due Line " _
        & "( " & .SelectedItem & ")?" _
        , vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        
        MakeEntry "X" & .SelectedItem
    End If
End With
Set gobjPNR = New CWT_GalileoPNR3.PNR
Call gobjPNR.loadPNR
Call FillInGSR
Call FillInPaidDue

End Sub

'Private Sub cmdInvoice_Click()
'    If UCase(gstrAgcyCountryCode) = "SG" Then
'    If CompareRDDI = False Then Exit Sub
'    End If
'    SysStart = Now
'    bolInvOnly = True
'    cmdInvoice.Enabled = False
'    Call IssueInvoice
'End Sub
Private Sub cmdNext_Click()
Dim intRetries As Integer
'Dim strTemp As String
'Dim strResp As String
'Dim strCT As String
'Dim strCI As String
'Dim bolTkt As Boolean

'Dim strPath As String
'Dim intFile As Integer
bolTktItinInv = True
cmdNext.Enabled = False
SysStart = Now
'Call addPaidLines

intRetries = 0
'CC - V1.2.3 20110629 - CR52 HBT Implementation - Remove CompareRDDI logic (check if total amount in Due line is equal to the amount in DI line)
'If UCase(gstrAgcyCountryCode) = "SG" Then
'    If CompareRDDI = False Then Exit Sub
'End If

'If IssueTktItin(intRetries) = False Then cmdNext.Enabled = True


'strCT = UCase(cmbTKT & "D  D")
'strCI = UCase(cmbITIN & "D  D")
'strTemp = "HMLM" & IIf(chkSTP.Value = 1, cmbSTP.Text, cmbTKT.Text) & "DT/" & cmbITIN.Text & "DI"

'strResp = MakeEntry(strTemp)
'If InStr(1, strResp, strCT) Then
'    If MsgBox("Ticket Printer is DOWN!" & vbCrLf & "Is it ready to print?" _
'        , vbExclamation + vbDefaultButton1 + vbOKCancel + vbApplicationModal) = vbOK Then
'            strTemp = "HMOM" & cmbTKT & "-U"
'            MakeEntry strTemp
'    Else
'        Exit Sub
'    End If
'End If

'If InStr(1, strResp, strCI) Then
'    If MsgBox("Itinerary Printer is DOWN!" & vbCrLf & "Is it ready to print?" _
'        , vbExclamation + vbDefaultButton1 + vbOKCancel + vbApplicationModal) = vbOK Then
'            strTemp = "HMOM" & cmbITIN & "-U"
'            MakeEntry strTemp
'    Else
'        Exit Sub
'    End If
'End If
'MakeEntry "HMOM" & IIf(chkSTP.Value = 1, cmbSTP.Text, cmbTKT.Text) & "-U"
'MakeEntry "HMOM" & cmbITIN.Text & "-U"
'MakeEntry "R.TPRO TKT"
'MakeEntry "IMU@"
'strTemp = "IMUDYO" & cmbITINDYO & IIf(txtCRItin.Text <> "", "/CR" & txtCRItin, "")
'MakeEntry strTemp
'MakeEntry "*" & gobjPNR.RecLoc
'If InStr(1, MakeEntry("TKPDTDID"), " GENERATED ") Then
'    bolTkt = True
'ElseIf InStr(1, MakeEntry("TKPDTDID"), " GENERATED ") Then
'    bolTkt = True
'Else
'    MsgBox "Unable to issue ticket/itinerary!"
'    Exit Sub
'End If
''Added on 30/07/04 - additional input for HKG


'strPath = App.Path
'strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
'intFile = FreeFile()
'    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
'    Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "start frmHKTktInput.Show vbModal"
'    Close #intFile
    
    
'If UCase(gstrAgcyCountryCode) = "HK" Then
'    MakeEntry "*" & gobjPNR.RecLoc
'    Sleep (1000)
'    Load frmHKTktInput
'    frmHKTktInput.Show vbModal
'End If

'    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
'    Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "End frmHKTktInput.Show vbModal"
'    Close #intFile

'
'cmdNext.Caption = "Waiting for Tkt info"

'    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
'    Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "Start Invoice"
'    Close #intFile

'Timer1.Enabled = True


End Sub
'Private Function IssueTktItin(Retries As Integer, Optional Invoice As Boolean = True, Optional AquaInvoice As Boolean = False) As Boolean
'Dim strTemp As String
'Dim strResp As String
'Dim strCT As String
'Dim strCI As String
'Dim bolTkt As Boolean
'Dim bolItin As Boolean
'Dim bolETkt As Boolean
'Dim strPath As String
'Dim intFile As Integer
'Dim intStage As Integer
'Dim bolSplitCmd As Boolean
'Dim strItin As String
'Dim itinResponse As String

'Const CONFIGPRNT As Integer = 1
'Const STATUSPRNT As Integer = 2
'Const PNRFORPRNT As Integer = 3
'Const TICKETPRNT As Integer = 4

'Dim TktResponse As String
'Dim strTKP As String

'On Error GoTo ErrIssueTktItin

'check whether eticket

'bolSplitCmd = SplitCmd

'bolETkt = False
'If cmbTKT.listindex > -1 Then

'    If cmbTKT.ItemData(cmbTKT.listindex) = 0 Then
'        bolETkt = True
'    Else
'        bolETkt = False
'    End If

'ElseIf cmbSTP.listindex > -1 Then
    
'    If cmbSTP.ItemData(cmbSTP.listindex) = 0 Then
'        bolETkt = True
'    Else
'        bolETkt = False
'    End If


'End If


'strCT = UCase(cmbTKT & "D  D")
'strCI = UCase(cmbITIN & "D  D")
'intStage = CONFIGPRNT


'If bolETkt = True Then
'   strTemp = "HMLM" & IIf(cmbSTP.Text <> "", cmbSTP.Text, cmbTKT.Text) & "DI"
'   strResp = MakeEntry(strTemp)

'Else                                            'normal
'   strTemp = "HMLM" & IIf(cmbSTP.Text <> "", cmbSTP.Text & "DS", cmbTKT.Text & "DT")
'   strResp = MakeEntry(strTemp)
'   If UCase(gstrAgcyCountryCode) = "SG" Then                        'ITIN -not stp
'      If cmbSTP.Text = "" Then  'HQ
'      strTemp = "HMLM" & cmbITIN.Text & "DI"
'      strResp = MakeEntry(strTemp)
'      End If
'   Else
'      strTemp = "HMLM" & cmbITIN.Text & "DI"
'      strResp = MakeEntry(strTemp)
'   End If
'End If
'intStage = STATUSPRNT

'If InStr(1, strResp, strCT) Then
'    If MsgBox("Ticket Printer is DOWN!" & vbCrLf & "Is it ready to print?" _
'        , vbExclamation + vbDefaultButton1 + vbOKCancel + vbApplicationModal) = vbOK Then
'            strTemp = "HMOM" & IIf(cmbSTP.Text <> "", cmbSTP.Text, cmbTKT.Text) & "-U"
'            MakeEntry strTemp
'    Else
        
'        Exit Function
'    End If
'End If

'If InStr(1, strResp, strCI) Then
'    If MsgBox("Itinerary Printer is DOWN!" & vbCrLf & "Is it ready to print?" _
'        , vbExclamation + vbDefaultButton1 + vbOKCancel + vbApplicationModal) = vbOK Then
'            strTemp = "HMOM" & cmbITIN & "-U"
'            MakeEntry strTemp
'    Else
        
'        Exit Function
'    End If
'End If

'MakeEntry "HMOM" & IIf(cmbSTP.Text <> "", cmbSTP.Text, cmbTKT.Text) & "-U" 'bring status up

'If UCase(gstrAgcyCountryCode) = "SG" Then
'   If cmbSTP.Text <> "" Then
'      MakeEntry "HMLM" & mstrACPrinter & "DT"
'      MakeEntry "HMOM" & mstrACPrinter & "-U"
'   End If
'End If

'If bolETkt = False Then
'   MakeEntry "HMOM" & cmbITIN.Text & "-U" 'then bring itin up
'   MakeEntry "HMOM" & cmbITIN.Text & "-ITN"
'End If

'MakeEntry "R.TPRO TKT"
'MakeEntry "IMU@"
'strTemp = "IMUDYO" & cmbITINDYO
'MakeEntry strTemp
'If txtCRItin.Text <> "" Then
'   strTemp = "IMUCR" & txtCRItin
'   MakeEntry strTemp
'End If
'intStage = PNRFORPRNT

'Modified on 07/09/04: end transaction after DYO changes
'MakeEntry "R.TPRO TKT"
'MakeEntry "ER"
'MakeEntry "ER"
'MakeEntry "ER"


'intStage = TICKETPRNT

'If optAll.Value Then
'   strTKP = "TKPDTD"
'ElseIf optSelection.Value Then
'   strTKP = "TKP" & txtFF & "P" & txtPax & "/DTD"
'End If
'If bolSplitCmd = True Then
'    If cmbSTP.Text <> "" Then
'      TktResponse = MakeEntry(strTKP & "/STP" & txtIATA)
'    Else
'      TktResponse = MakeEntry(strTKP) 'TKPDTDID
'    End If
'Else
'    If cmbSTP.Text <> "" Then
'       If UCase(gstrAgcyCountryCode) = "SG" Then 'sg
'          TktResponse = MakeEntry(strTKP & "/STP" & txtIATA)
'       Else
'          If optAll.Value Then 'hk
'             TktResponse = MakeEntry(strTKP & "ID/STP" & txtIATA) 'TKPDTDID/STP13305164
'          Else
'             TktResponse = MakeEntry(strTKP & IIf(chkItinerary.Value = 1, "ID", "") & "/STP" & txtIATA)
'          End If 'TKP1P1DTDID/STP13305164
'       End If
'    Else 'HQ
'       If optAll.Value Then
'          TktResponse = MakeEntry(strTKP & IIf(cmbITIN <> "", "ID", "")) 'TKPDTDID
'       Else
'          TktResponse = MakeEntry(strTKP & IIf(chkItinerary.Value = 1, "ID", "")) 'TKP1P1DTDID
'       End If
'    End If
'End If



'If InStr(1, TktResponse, "TKT GENERATED") > 0 Then
'    MsgBox TktResponse
'    MakeEntry "*" & gobjPNR.RecLoc
'    MakeEntry "NP.SS*VBITKT"
'    IssueTktItin = True
'    If Invoice = True Then
'        If AquaInvoice = True Then
'            Exit Function
'        End If
'        MakeEntry "R.TPRO TKT+ER"
'        MakeEntry "ER"
'        MakeEntry "ER"
'    Else
'        MakeEntry "R.TPRO TKT+ER"
'        MakeEntry "ER"
'        MakeEntry "ER"
'        Exit Function
'    End If
'Else
'    If Invoice = True Then
'        If AquaInvoice = True Then
'            MsgBox TktResponse & vbCrLf & "Invoice will not be queue to Aqua for invoice issuance."
'        Else
'           MsgBox TktResponse & vbCrLf & "Invoice will not be issued."
'        End If
'    Else
'        MsgBox TktResponse
'    End If
'        IssueTktItin = False
'        strPath = App.Path
'        strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
'        intFile = FreeFile()
'        Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
'        Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "Exit function"
'        Close #intFile
'        Exit Function
'End If

            '29122004
            'If (InStr(1, TktResponse, "ELECTRONIC TICKETING FAILED") <> 0 Or _
            '   InStr(1, TktResponse, "INSUFFICIENT FUNDS") <> 0 Or _
            '   InStr(1, TktResponse, "PART CANCELLED FARE TICKET BY PASSENGER") <> 0 Or _
            '   InStr(1, TktResponse, "SIMULTANEOUS CHANGES") <> 0)  Then
            '   strPath = App.Path
            '   strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
            '   intFile = FreeFile()
            '    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
            '    Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "back to main menu"
            '    Close #intFile
            '   Call pRedisplayMenu
            '   frmMainMenu.WindowState = 1
            '   Exit Function
            'End If
            'If chkSTP.Value Then
            '   If UCase(gstrAgcyCountryCode) = "SG" Then
            '      TktResponse = MakeEntry("TKPDTD/STP" & txtIATA)
            '   Else
            '      TktResponse = MakeEntry("TKPDTDID/STP" & txtIATA)
            '   End If
            'Else
            '   TktResponse = MakeEntry("TKPDTDID")
            'End If


'If bolSplitCmd = True Then

'strResp = MakeEntry("*" & gobjPNR.RecLoc)
'If InStr(1, strResp, "FINISH OR IGNORE") <> 0 Or _
'   InStr(1, strResp, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
'   Sleep (1000)
'   MakeEntry "I"
'   Sleep (500)
'   strResp = MakeEntry("*" & gobjPNR.RecLoc)
'End If
'Sleep (1000)
'MakeEntry "IR"
'MakeEntry "IR"

'print bring itinerary printer up

'    strTemp = "HMLM" & cmbITIN.Text & "DI"
'    strResp = MakeEntry(strTemp)


'If InStr(1, strResp, strCI) Then
'    If MsgBox("Itinerary Printer is DOWN!" & vbCrLf & "Is it ready to print?" _
'        , vbExclamation + vbDefaultButton1 + vbOKCancel + vbApplicationModal) = vbOK Then
'           strTemp = "HMOM" & cmbITIN & "-U"
'            MakeEntry strTemp
'    Else
'        Exit Function
'    End If
'End If



'MakeEntry "HMOM" & cmbITIN.Text & "-U"
'MakeEntry "HMOM" & cmbITIN.Text & "-ITN"

'MakeEntry "R.TPRO TKT"
'MakeEntry "IMU@"
'strTemp = "IMUDYO" & cmbITINDYO
'MakeEntry strTemp
'If txtCRItin.Text <> "" Then
'   strTemp = "IMUCR" & txtCRItin
'   MakeEntry strTemp
'End If

'Modified on 07/09/04: end transaction after DYO changes
'MakeEntry "R.TPRO TKT"
'MakeEntry "ER"
'MakeEntry "ER"
'MakeEntry "ER"

'If optAll.Value Then
'   strItin = "TKPDID"
'ElseIf optSelection.Value Then
'   strItin = "TKP" & txtFF & "P" & txtPax & "/DID"
'End If

'If optAll.Value Then
'   If cmbITIN <> "" Then
'     itinResponse = MakeEntry(strItin) 'TKPDTDID
'   End If
'Else
'   If chkItinerary.Value = 1 Then
'     itinResponse = MakeEntry(strItin) 'TKP1P1DTDID
'   End If
'End If


'MsgBox itinResponse

'End If



        'MsgBox "Ticket Response = " & TktResponse
        'bolTkt = True
        
        'If InStr(1, MakeEntry("TKPDTDID"), " GENERATED ") Then
        '    bolTkt = True
        'ElseIf InStr(1, MakeEntry("TKPDTDID"), " GENERATED ") Then
        '    bolTkt = True
        'Else
        '    MsgBox "Unable to issue ticket/itinerary!"
        '    Exit Sub
        'End If
'Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", startTime, SysStart, "Ticket/Itin", , startTime)



        '29122004
        'If optSelection.Value And chkInv.Value = 0 Then Exit Sub
'If optSelection.Value And chkInv.Value = 0 Then
'   bolTktItin = True
'   strResp = MakeEntry("*" & gobjPNR.RecLoc)
'   If InStr(1, strResp, "FINISH OR IGNORE") <> 0 Or _
'      InStr(1, strResp, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
'      Sleep (1000)
'      MakeEntry "I"
'      Sleep (500)
'      strResp = MakeEntry("*" & gobjPNR.RecLoc)
'   End If
'   Sleep (1000)
   '29122004
'   MakeEntry "IR"
'   MakeEntry "IR"
    
            ' Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", StartTime)
             'Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", startTime, SysStart, "Ticket/Itin", , startTime)
'   Exit Function
'End If

            'Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", startTime, SysStart, "Ticket/Itin", , startTime)


'If Invoice = True Then
'    cmdNext.Caption = "Waiting for Tkt info"
'    strPath = App.Path
'    strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
'    intFile = FreeFile()
    
'        Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
'        Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "Start Invoice"
'        Close #intFile
'    Timer1.Enabled = True
'End If

'Exit Function

'ErrIssueTktItin:
'Select Case Err.Number
'    Case -2147467259
'        If intStage < 4 Then
'            Retries = Retries + 1
'            If Retries < 3 Then
'                Call IssueTktItin(Retries, Invoice)
'            Else
'                MsgBox CONNECTION_FAIL & " : Please re-run the ticketing process", vbCritical
'                Exit Function
'            End If
'        Else
'            Call IssueInvoice
'        End If
'    Case Else
'        MsgBox "ERROR " & Err.Number & vbCrLf _
'            & Err.Description, "RUN TIME ERROR"
'        Resume Next
'End Select

'End Function



Private Sub cmdTicket_Click()
Dim intRetries As Integer
Dim strRes As String
bolTktItinInv = False

cmdTicket.Enabled = False
SysStart = Now

intRetries = 0
'CC - V1.2.3 20110629 - CR52 HBT Implementation - Remove CompareRDDI logic (check if total amount in Due line is equal to the amount in DI line)
'If CompareRDDI = False Then Exit Sub
'Call IssueTktItin(intRetries, False)

If UCase(gstrAgcyCountryCode) = "SG" Then
   If gobjPNR.CN = "39001001" Then
   
        strRes = MakeEntry("*" & gobjPNR.RecLoc)
        If InStr(1, strRes, "FINISH OR IGNORE") <> 0 Or _
           InStr(1, strRes, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
           Sleep (1000)
           MakeEntry "I"
           Sleep (500)
           strRes = MakeEntry("*" & gobjPNR.RecLoc)
        End If
        'MakeEntry "QEB/781Q/26"
   End If
End If

cmdTicket.Enabled = True

End Sub

'Private Sub cmdTicketAqua_Click()
'Dim intRetries As Integer

'bolTktItinInv = False

'cmdTicketAqua.Enabled = False
'SysStart = Now

'intRetries = 0

'If CompareRDDI = False Then Exit Sub
'If IssueTktItin(intRetries, True, True) = True Then
'        MakeEntry "NP.TT*INVONLY"
'        MakeEntry "NP.TQ*61"
'        MakeEntry "R.TPRO TKT+ER"
'        MakeEntry "ER"
'        MakeEntry "ER"
'        MakeEntry "QEB/5E4P/81"
'End If
'cmdTicketAqua.Enabled = True
'Call pRedisplayMenu
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmTktInfo = Nothing
End Sub

Private Sub chkInvType_Click()
   Dim intI As Integer

   If chkInvType.value = 1 Then
      chkInvType.Caption = "E-Inv"
      'cmbEINV.Enabled = True
      'cmbINV.Enabled = False
      For intI = 0 To cmbINV.ListCount - 1
          If cmbINV.ItemData(intI) = 0 Then
             cmbINV.listindex = intI
             Exit For
          End If
      Next
      
      For intI = 0 To cmbINVDYO.ListCount - 1
          If cmbINVDYO.ItemData(intI) = 0 Then
             cmbINVDYO.listindex = intI
             Exit For
          End If
      Next
   Else
      chkInvType.Caption = "P-Inv"
      'cmbEINV.Enabled = False
      'cmbINV.Enabled = True
     For intI = 0 To cmbINV.ListCount - 1
          If cmbINV.ItemData(intI) = 1 Then
             cmbINV.listindex = intI
             Exit For
          End If
      Next
      
      For intI = 0 To cmbINVDYO.ListCount - 1
          If cmbINVDYO.ItemData(intI) = 1 Then
             cmbINVDYO.listindex = intI
             Exit For
          End If
      Next
   End If
End Sub







'Private Sub optAll_Click()
'   txtFF.Enabled = False
'   txtPax.Enabled = False
'   chkItinerary.Enabled = False
'   chkInv.Enabled = False
'End Sub

'Private Sub optAllNew_Click()
'If optAllNew.Value = False Then
'   optAllNew.Value = True
'   chkItinerary.Enabled = False
'   chkItinerary.Value = vbChecked
'   chkInvoice.Enabled = False
'   chkInvoice.Value = vbChecked
'   chkMir.Enabled = False
'   chkMir.Value = vbChecked
'   chkTkt.Enabled = False
'   chkTkt.Value = vbChecked
'Else
'   optAllNew.Value = False
'   chkItinerary.Enabled = True
'   chkItinerary.Value = vbUnchecked
'   chkInvoice.Enabled = True
'   chkInvoice.Value = vbUnchecked
'   chkMir.Enabled = True
'   chkMir.Value = vbUnchecked
'   chkTkt.Enabled = True
'   chkTkt.Value = vbUnchecked
'End If
   
   
'End Sub

'Private Sub optSelection_Click()
'   txtFF.Enabled = True
'   txtPax.Enabled = True
'   chkItinerary.Enabled = True
'   chkInv.Enabled = True
'End Sub

Private Sub timer1_Timer()
'If bolFirst <> True Then
bolFirst = False
    'invStart = Now
'End If
Call IssueInvoice(strDoc)
End Sub
Private Sub IssueInvoice(doctype As String)
Dim intNT As Integer
Dim strTemp As String
Dim lngPos As Long
Dim intFFNum As Integer
Dim i As Integer
Dim strRmks() As String
Dim InvResponse As String
Dim strRes As String
Dim strPath As String
Dim intFile As Integer
Dim intPx As Integer
Dim strCSTktNum As String
Dim bolInvoiceIssued As Integer
Dim strMsg As String
Dim strFFSelected() As String

Timer1.Enabled = False
invStart = Now

strPath = App.Path
strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
intFile = FreeFile()
Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "Start issue " & doctype & " in IssueInvoice"
Close #intFile

strRes = MakeEntry("*" & gobjPNR.RecLoc)
If InStr(1, strRes, "FINISH OR IGNORE") <> 0 Or _
   InStr(1, strRes, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
   Sleep (1000)
   MakeEntry "I"
   Sleep (500)
   strRes = MakeEntry("*" & gobjPNR.RecLoc)
End If
Sleep (1000)
'29122004
MakeEntry "IR"
MakeEntry "IR"

ReStartInv:
Set gobjPNR = New CWT_GalileoPNR3.PNR
With gobjPNR
    .loadPNR
End With
If doctype <> "MIR" Then
Call addPaidLines
ChangeTktNum
'29122004
If ChangeTktNum = False Then Exit Sub
With gobjPNR
For intNT = 1 To .FiledFareCount
  For intPx = 1 To .FiledFare(intNT).PxCount
    If .FiledFare(intNT).PX(intPx).TicketNumber <> "" Then
        For i = 1 To .GASaleRecordCount
        If chkReplace(.GASalesRecord(i).ProductCode) = True Then
            If .AcctRemarkCount > 1 And .GASalesRecord(i).BegLine > 1 Then
                strRmks = Split(.AcctRemark(.GASalesRecord(i).BegLine - 1).RemarkText, "/")
                If UBound(strRmks) > 0 Then
                    If Left(.GASalesRecord(i).TicketNumber, 10) <> "0000000000" Then
                       GoTo NextGASaleRecord
                    End If
                    If i = 1 And (Left(strRmks(1), 1) <> "*" Or IsNumeric(Mid(strRmks(1), 2)) = False) Then
                       MakeEntry "IR"
                       'MsgBox "Please move MS/MSX DI line for " & .GASalesRecord(i).ProductCode & " below the File Fare DI line in order to replace the ticket number."
                       strMsg = "Please move MS/MSX DI line for " & .GASalesRecord(i).ProductCode & " below the File Fare DI line in order to replace the ticket number."
                       modMsgBox.OKMsg = "OK"
                       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                       Exit Sub
                    End If
                    If IsNumeric(Mid(strRmks(1), 2)) And Left(strRmks(1), 1) = "*" Then
                       intFFNum = Int(Mid(strRmks(1), 2))
                       'If Int(Mid(strRmks(1), 2)) > intNT Then
                       '   GoTo NextFF
                       'End If
                    End If
                    'If (strRmks(1) = ("*" & intNT) Or strRmks(1) = ("*0" & intNT) Or strRmks(0) = "MSX") Then 'And Left(.GASalesRecord(i).TicketNumber, 10) = "0000000000" Then
                    If intFFNum = intNT And Left(.GASalesRecord(i).TicketNumber, 10) = "0000000000" Then
                    'If .GASaleRecordCount <= intNT And Left(.GASalesRecord(intNT).TicketNumber, 10) = "0000000000" Then
                        lngPos = .GASalesRecord(i).BegLine
                        'Modified on 27/1/06: CS Changes
                        
                        strCSTktNum = CSTktNumFormat(.FiledFare(intNT).PX(intPx).TicketNumber)
                        'strTemp = "DI." & .GASalesRecord(i).BegLine & "@FT-" & Left(.AcctRemark(lngPos).RemarkText, InStr(.AcctRemark(lngPos).RemarkText, "/TK") + 2) _
                        '    & fConvertTkTNo(.FiledFare(intNT).PX(intPx).TicketNumber)
                        'strTemp = "DI." & .GASalesRecord(i).BegLine & "@FT-" & Left(.AcctRemark(lngPos).RemarkText, InStr(.AcctRemark(lngPos).RemarkText, "/TK") + 2) _
                        '    & .FiledFare(intNT).PX(intPx).TicketNumber & "/PX" & intPx
                        strTemp = "DI." & .GASalesRecord(i).BegLine & "@FT-" & Left(.AcctRemark(lngPos).RemarkText, InStr(.AcctRemark(lngPos).RemarkText, "/TK") + 2) _
                            & strCSTktNum & "/PX" & intPx
                        MakeEntry strTemp
                        'Exit For
                    End If
                End If
            End If
            
            End If
NextGASaleRecord:
            
        Next i
    '29122004
    ElseIf .FiledFare(intNT).PX(.FiledFare(intNT).PxCount).TicketNumber = "" Then
    'Else
        strMsg = "Still need to wait for ticket numbers"
        modMsgBox.OKMsg = "Ok"
        modMsgBox.CANCELMsg = "Cancel"
        If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbOKCancel + vbDefaultButton2, "CWT Desktop") = vbOK Then
        'If MsgBox("Still need to wait for ticket numbers", vbApplicationModal + vbOKCancel) = vbOK Then
            strRes = MakeEntry("IR")
            Timer1.Enabled = True
            Exit Sub
        Else
            Exit Sub
        End If
    End If
NextFF:
Next intPx
Next
End With

'Added on 30/07/04 - additional input for HKG
If UCase(gstrAgcyCountryCode) = "HK" Then
    strPath = App.Path
    strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
    intFile = FreeFile()
    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
    Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "start frmHKTktInput.Show vbModal"
    Close #intFile
    
    
    MakeEntry "*" & gobjPNR.RecLoc
    Sleep (1000)
    Load frmHKTktInput
    frmHKTktInput.Show
    Do
      DoEvents
    Loop Until isLoaded("frmHKTktInput") = False

    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
    Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "End frmHKTktInput.Show vbModal"
    Close #intFile
End If

intFFNum = gobjPNR.FiledFareCount + 1
For i = 1 To intFFNum
  MakeEntry "TMU" & i & "/IT@"
  'REQUEST BY HKG 200206
  MakeEntry "TMU" & i & "/BT@"
Next

End If

If doctype = "INV" Or doctype = "ALL" Then
    MakeEntry "HMLM" & cmbINV.Text & "DI"
    MakeEntry "HMOM" & cmbINV.Text & "-U"
    MakeEntry "R.TPRO TKT"
    MakeEntry "IMU@"
    strTemp = "IMUDYO" & cmbINVDYO & IIf(txtCRInv.Text <> "", "/CR" & txtCRInv, "")
    MakeEntry strTemp
    
End If

If doctype = "MIR" Or doctype = "ALL" Then
    MakeEntry "HMLM" & cmbMIR.Text & "DA"
    MakeEntry "HMOM" & cmbMIR.Text & "-U"
End If

'Preethi - V1.2.1 20101011 - CR21 - Nett Fare Mark Up
'If UCase(gstrAgcyCountryCode) = "HK" Then
   'Remove NF from TMU if DI.FT-NF/*n/xxxx exists. n denotes file fare number
   If strFF = "" Then
        For intNT = 1 To gobjPNR.FiledFareCount
             For i = 1 To gobjPNR.AcctRemarkCount
                 With gobjPNR.AcctRemark(i)
                      If .RemarkType = "FT" And InStr(.RemarkText, "NF/*" + CStr(intNT)) > 0 Then
                          MakeEntry "TMU" + CStr(intNT) + "NF@"
                          Exit For
                      End If
                 End With
             Next
        Next
   Else
        strFFSelected = Split(strFF, ".")
        For intNT = LBound(strFFSelected) To UBound(strFFSelected)
             For i = 1 To gobjPNR.AcctRemarkCount
                 With gobjPNR.AcctRemark(i)
                      If .RemarkType = "FT" And InStr(.RemarkText, "NF/*" + strFFSelected(intNT)) > 0 Then
                          MakeEntry "TMU" + strFFSelected(intNT) + "NF@"
                          Exit For
                      End If
                 End With
             Next
        Next
   End If
'End If

'Modified on 07/09/04: end transaction after DYO changes
MakeEntry "R.TPRO TKT"
strTemp = MakeEntry("ER")

If InStr(1, strTemp, "SIMULTANEOUS CHANGES TO BOOKING FILE") <> 0 Then
   MakeEntry "IR"
   GoTo ReStartInv
End If
MakeEntry "ER"
MakeEntry "ER"

Dim strMir As String
Dim strInv As String

        If strFF = "" Then
            strMir = "TKP"
            strInv = "TKP"
        Else
            strMir = "TKP" & strFF
            strInv = "TKP" & strFF
        End If
        
        If strPax = "" Then
                strMir = strMir & "DAD"
                strInv = strInv & "DID"
        Else
                strMir = strMir & "P" & strPax & "/DAD"
                strInv = strInv & "P" & strPax & "/DID"
        End If



If doctype = "ALL" Or doctype = "MIR" Then
    InvResponse = MakeEntry(strMir & IIf(doctype = "ALL", "ID", ""))
Else
    InvResponse = MakeEntry(strInv)
End If
'InvResponse = MakeEntry("TKPDADID")
'29122004

If InStr(1, InvResponse, "GENERATED") = 0 Then
   'MsgBox "Unable to issue Invoice!" & vbCrLf & "Response:" & vbCrLf & InvResponse
   strMsg = "Unable to issue Invoice!" & vbCrLf & "Response:" & vbCrLf & InvResponse
   modMsgBox.OKMsg = "OK"
   modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
   bolInvoiceIssued = False
Else
   'MsgBox "GDSResponse: " & InvResponse
   strMsg = "GDSResponse: " & InvResponse
   modMsgBox.OKMsg = "OK"
   modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop"
   bolInvoiceIssued = True
End If
'If InStr(1, InvResponse, "ITINERARY/INVOICE GENERATED") = 0 Then MsgBox InvResponse

'If InStr(1, MakeEntry("TKPDADID"), " GENERATED ") Then
'invoice issued
'ElseIf InStr(1, MakeEntry("TKPDADID"), " GENERATED ") Then
    'invoice issued on 2nd attempt
'Else
'    MsgBox "Unable to issue Invoice!"
'    Exit Sub
'End If
   
'Timer
If doctype = "ALL" Then
    strTemp = "Inv/Mir"
Else
    strTemp = doctype
End If
If bolTktItinInv Then
    Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", invStart, invStart, strTemp, , startTime)
Else
    Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", startTime, SysStart, strTemp, , startTime)
End If

'Added on 280306: Do Credit Card Sale for SG


'If UCase(gstrAgcyCountryCode) = "SG" Then
'    If bolInvoiceIssued = True And gobjPNR.CompInfo.CPG Then
'        DoAutoApproval
'    End If
'End If
'Added on 280306: Do Credit Card Sale for SG


If UCase(gstrAgcyCountryCode) = "SG" Then
    'If bolInvoiceIssued = True And gobjPNR.CompInfo.CPG Then
    If gobjPNR.CompInfo.CPG Then
        DoAutoApproval
    Else
        If CheckCPGClient(gobjPNR.CN) Then
            DoAutoApproval
        End If
    End If
End If

  'Call pRedisplayMenu
  
If UCase(gstrAgcyCountryCode) = "SG" Then
   If gobjPNR.CN = "39001001" Then
   
        strRes = MakeEntry("*" & gobjPNR.RecLoc)
        If InStr(1, strRes, "FINISH OR IGNORE") <> 0 Or _
           InStr(1, strRes, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
           Sleep (1000)
           MakeEntry "I"
           Sleep (500)
           strRes = MakeEntry("*" & gobjPNR.RecLoc)
        End If
        'MakeEntry "QEB/781Q/26"
   End If
End If

'MakeEntry "HMLMC2557ADI"
'MakeEntry "HMLM" & cmbITIN.Text & "DI"


'Remarks by JY ANG 31 Oct 2007. Eliminate duplicate msgbox when issue invoice and MIR
'If MsgBox("Do you want to redisplay PNR " & lblPNR & " in focalpoint?" _
'        , vbApplicationModal + vbQuestion + vbYesNo) = vbNo Then
'    Exit Sub
'Else
'    pDisplayToFP "*" & gobjPNR.RecLoc
'End If
 
'Call pRedisplayMenu

'End If

End Sub
Private Function chkReplace(SortKey As String) As Boolean

Dim strSQL As String
Dim rs As ADODB.Recordset

strSQL = "SELECT TktNo from tblProductcodes where SortKey='" & SortKey & "'"
Set rs = gdbConn.Execute(strSQL)

If Not rs.EOF Then
    chkReplace = rs!TktNo
Else
    chkReplace = False
End If
rs.Close
Set rs = Nothing

End Function
Private Sub Form_Load()
Dim item As ListItem
Dim intI As Integer
Dim sngTotal As Single
Dim sngPaid As Single
Dim sngCX As Single
Dim oldParent As Long
datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)


Me.Move 0, 0
Me.Move frmSideBar.Width, 0

'Clement
strNow = Format(Now, "DDMMYYhhmmss")

'Timer
startTime = Now

'optAll.Value = 1
If UCase(gstrAgcyCountryCode) = "SG" Then
txtCRItin = "1-12"
End If
'chkSTP_Click
Set mobjTE = New HostAccess.TerminalEmulation
Set gobjPNR = New CWT_GalileoPNR3.PNR
   gobjPNR.loadPNR
   'Timer
   'Call pAddToVBILog(gobjPNR.RecLoc, "Ticketing")

Call GetAddress(gobjPNR.PCCOwner, "HQ")
FillLocation

   lblPNR = "Record: " & gobjPNR.RecLoc
   
        For intI = 1 To gobjPNR.AirSegCount
          lstSegments.AddItem gobjPNR.AirSeg(intI).TextAirSeg
        Next
      
       For intI = 1 To gobjPNR.HotelSegCount
          lstSegments.AddItem gobjPNR.HotelSeg(intI).TextHtlSeg
      Next
      
    Call FillInGSR
    Call FillInPaidDue
    Call FillInFiledFares
    
    If IsNumeric(lblGASRTtotal) And IsNumeric(lblTaxTotal) Then
       lblDITaxTotal = Format(CDec(lblGASRTtotal) + CDec(lblTaxTotal), gstrAgcyCurrFormat)
    Else
       lblDITaxTotal = ""
    End If
   'chkSTP.Value = vbUnchecked
   'GetAddress (gobjPNR.PCCOwner)
   
   'If IsEINV(gobjPNR.CN) Then
   'added on 10/5/2005
   'If cmbTKT.ListIndex > -1 Then
   ' If cmbTKT.ItemData(cmbTKT.ListIndex) = 0 Then
   '    chkTktType.Value = 1
   '    chkTktType.Caption = "E-Tkt"
   ' Else
   '    chkTktType.Value = 0
   '    chkTktType.Caption = "P-Tkt"
   ' End If
   'Else
   '   chkTktType.Value = 0
   '   chkTktType.Caption = "P-Tkt"
   'End If
   
  
   
 sstTicketing.Tab = 0
 
 If UCase(gstrAgcyCountryCode) = "SG" Then
    cmdTicket.Visible = True
    cmdTicketAqua.Visible = True
    If gobjPNR.FiledFareCount > 1 Or gobjPNR.PassengerCount > 1 Then
        cmdTicketAqua.Enabled = False
    End If
 Else
    cmdTicket.Visible = False
    cmdTicketAqua.Visible = False
 End If
 
 PopulatePax
 PopulateFilefare
 SetPE
 Set moTags = New TreeViewTag
 setDefault
  datFormLoadEnd = Now
  If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
 

 
 
End Sub
Private Sub SetPE()
Dim intI As Integer
 If IsEINV(gobjPNR.CompInfo.ProfileName) Then
      'cmbEINV.Enabled = True
      'cmbINV.Enabled = False
      For intI = 0 To cmbINV.ListCount - 1
             If cmbINV.ItemData(intI) = 0 Then
                cmbINV.listindex = intI
             Exit For
             End If
      Next
      chkInvType.value = 1
      'chkInvType.Visible = False
      chkInvType.Caption = "E-Inv"
      For intI = 0 To cmbINVDYO.ListCount - 1
          If cmbINVDYO.ItemData(intI) = 0 Then
             cmbINVDYO.listindex = intI
             Exit For
          End If
      Next
   Else
      'cmbEINV.Enabled = False
      cmbINV.Enabled = True
      chkInvType.Visible = True
      chkInvType.value = 0
      chkInvType.Caption = "P-Inv"
   End If
 If IsEItin(gobjPNR.CompInfo.ProfileName) Then
    chkItinType.value = vbChecked
 Else
    chkItinType.value = vbUnchecked
 End If
End Sub
Private Sub setDefault()
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim strCheck() As String
Dim intI As Integer

strSQL = "select * from tblMODOptions where Optioncode='TktDefault'"

Set rs = gdbConn.Execute(strSQL)

If Not rs.EOF Then
    strCheck = Split(rs!optionvalue, ";")
End If

For intI = LBound(strCheck) To UBound(strCheck)
    If UCase(strCheck(intI)) = UCase("All") Then
        chkAll.value = vbChecked
    ElseIf UCase(strCheck(intI)) = UCase("Ticket") Then
        chkTkt.value = vbChecked
    ElseIf UCase(strCheck(intI)) = UCase("Itinerary") Then
        chkItinerary.value = vbChecked
    ElseIf UCase(strCheck(intI)) = UCase("Invoice") Then
        chkInvoice.value = vbChecked
    ElseIf UCase(strCheck(intI)) = UCase("Aqua Invoice and Mir") Then
        chkAqua = vbChecked
    End If
Next intI


rs.Close
Set rs = Nothing

'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
If gobjPNR.CompInfo.AquaItin Then
    chkItinerary.Enabled = False
    chkItinerary.value = 0
End If

End Sub
Private Sub PopulatePax()
Dim intI As Integer

'lstPax.AddItem "ALL"

For intI = 1 To gobjPNR.PassengerCount
    lstPax.AddItem gobjPNR.PassengerName(intI).FirstName & "/" & gobjPNR.PassengerName(intI).LastName
Next intI

'lstPax.Selected(1) = True
End Sub
Private Sub PopulateFilefare()
Dim intI As Integer
Dim intJ As Integer
Dim nodx As node
For intI = 1 To gobjPNR.FiledFareCount
    Set nodx = tvStoredFare.Nodes.Add(, , "SF" & intI, "STORED FARE - " & intI)
    
    For intJ = 1 To gobjPNR.FiledFare(intI).PxCount
    With gobjPNR.FiledFare(intI).PX(intJ)
    
    'treeview1.Nodes.Add "MyKey", tvwChild, , "Child Node #1"
        Set nodx = tvStoredFare.Nodes.Add("SF" & intI, tvwChild, "Child1FF" & intI & "PX" & intJ, "FareType: " & .FareGuarCode & "   Grand Total: " & .TotAmount)
        Set nodx = tvStoredFare.Nodes.Add("SF" & intI, tvwChild, "Child2FF" & intI & "PX" & intJ, "Quoted On: " & .CreatedDate)
        If .BaseCurrency = .TotalCurrency Then
            Set nodx = tvStoredFare.Nodes.Add("SF" & intI, tvwChild, "Child3FF" & intI & "PX" & intJ, "Based Fare: " & .BaseAmount & "    Total Taxes: " & .TaxTotal)
        Else
            Set nodx = tvStoredFare.Nodes.Add("SF" & intI, tvwChild, "Child4FF" & intI & "PX" & intJ, "Based Fare: " & .EquivAmount & "    Total Taxes: " & .TaxTotal)
        End If
        Set nodx = tvStoredFare.Nodes.Add("SF" & intI, tvwChild, "Child5FF" & intI & "PX" & intJ, "Modifiers")
        Set nodx = tvStoredFare.Nodes.Add("Child5FF" & intI & "PX" & intJ, tvwChild, "Child5AFF" & intI & "PX" & intJ, "Plating Carrier: " & .ValidatingCarrier)
        Set nodx = tvStoredFare.Nodes.Add("Child5FF" & intI & "PX" & intJ, tvwChild, "Child5BFF" & intI & "PX" & intJ, "Commission: " & .Commission)
        If .FOP_CCCode <> "" Then
        Set nodx = tvStoredFare.Nodes.Add("Child5FF" & intI & "PX" & intJ, tvwChild, "Child5CFF" & intI & "PX" & intJ, "FOP: " & .FOPType & .FOP_CCCode & .FOP_CCNum & "EXP" & Format(.FOP_CCExpireDate, "MMYY"))
        Else
        Set nodx = tvStoredFare.Nodes.Add("Child5FF" & intI & "PX" & intJ, tvwChild, "Child5CFF" & intI & "PX" & intJ, "FOP: " & .FOPType)
        End If
        Set nodx = tvStoredFare.Nodes.Add("Child5FF" & intI & "PX" & intJ, tvwChild, "Child5DFF" & intI & "PX" & intJ, "Tour Code: " & .TourCode)
        Set nodx = tvStoredFare.Nodes.Add("Child5FF" & intI & "PX" & intJ, tvwChild, "Child5EFF" & intI & "PX" & intJ, "NF: " & .NetAmount)
        Set nodx = tvStoredFare.Nodes.Add("Child5FF" & intI & "PX" & intJ, tvwChild, "Child5FFF" & intI & "PX" & intJ, "PT/ET: " & IIf(.ETktIndicator, "ET", "PT"))
        Set nodx = tvStoredFare.Nodes.Add("Child5FF" & intI & "PX" & intJ, tvwChild, "Child5GFF" & intI & "PX" & intJ, "ASF: " & .ASF)

    End With
    Next intJ

Next intI

 'Dim oNode As MSComctlLib.Node
 '   For Each oNode In tvStoredFare.Nodes
 '       'pColorize oNode
 '       oNode.Expanded = True
 '   Next

End Sub
Private Sub GetAddress(PCC As String, location As String)
   Dim rsTkt As New ADODB.Recordset
   Dim strSQL As String
   Dim item As ListItem
   Dim lngC As Long
   'strSQL = "Select * from tblDeviceAddr " & _
   '         "where PCC = '" & PCC & "' " & _
   '         "Order by Default"
   'strSQL = "Select * from tblDeviceAddr " & _
   '         "Order by DeviceDefault desc"
   'strSql = "SELECT DISTINCT Address, Type From tblDeviceAddr"
   strSQL = "select * from tblDeviceAddr where Loc= '" & location & "' order by devicedefault desc"
   
   Set rsTkt = gdbConn.Execute(strSQL)
   
   'cmbSTPLoc.Clear
   cmbITIN.Clear
   cmbTKT.Clear
   cmbINV.Clear
   'cmbEINV.Clear
   cmbMIR.Clear
   cmbINVDYO.Clear
   cmbITINDYO.Clear
   cmbSTP.Clear
   txtIATA = ""
   With rsTkt
     Do Until .EOF
       Select Case !Type
          Case "STP"
             lblSTPlabel.Visible = True
             cmbSTP.Visible = True
             txtIATA = !IATA
             cmbSTP.AddItem !Address

                If !EFormat = True Then
                   cmbSTP.ItemData(cmbSTP.NewIndex) = 0
                Else
                   cmbSTP.ItemData(cmbSTP.NewIndex) = 1
                End If
             
          Case "ITIN"
             cmbITIN.AddItem !Address
             If !devicedefault = True Then
                cmbITIN.Text = !Address
                strITINDefault = !Address
                intDefaultItin = cmbITIN.ListCount - 1
             End If
          Case "TKT"
             cmbTKT.AddItem !Address
             
             If !devicedefault = True Then
                mstrACPrinter = !Address
             End If
             
             If !EFormat = True Then
                cmbTKT.ItemData(cmbTKT.NewIndex) = 0
             Else
                cmbTKT.ItemData(cmbTKT.NewIndex) = 1
             End If
             

          Case "INV", "EINV"
             cmbINV.AddItem !Address
             
             If UCase(!Type) = "EINV" Then
                cmbINV.ItemData(cmbINV.NewIndex) = 0
             Else
                cmbINV.ItemData(cmbINV.NewIndex) = 1
             End If
          'Case "EINV"
          '   cmbEINV.AddItem !Address
          '   If !DeviceDefault = True Then
          '      cmbEINV.Text = !Address
          '   End If
          Case "MIR"
             cmbMIR.AddItem !Address
             If !devicedefault = True Then
                cmbMIR.Text = !Address
             End If
          Case "INVDYO"
             cmbINVDYO.AddItem !Address
             
                If !EFormat = True Then
                   cmbINVDYO.ItemData(cmbINVDYO.NewIndex) = 0
                Else
                   cmbINVDYO.ItemData(cmbINVDYO.NewIndex) = 1
                End If
 
            
          Case "ITINDYO"
             cmbITINDYO.AddItem !Address
            
                If !EFormat = True Then
                   cmbITINDYO.ItemData(cmbITINDYO.NewIndex) = 0
                Else
                   cmbITINDYO.ItemData(cmbITINDYO.NewIndex) = 1
                End If

             
       End Select
       .MoveNext
     Loop
    .Close
   End With
   
   
strSQL = "select * from tblDeviceAddr where Loc= '" & location & "' and DeviceDefault=1"
Set rsTkt = gdbConn.Execute(strSQL)
With rsTkt
    While Not .EOF
           Select Case !Type
          Case "STP"
             For lngC = 0 To cmbSTP.ListCount
                If cmbSTP.List(lngC) = !Address Then
                    cmbSTP.listindex = lngC
                    Exit For
                End If
             Next
          Case "ITIN"
             For lngC = 0 To cmbITIN.ListCount
                If cmbITIN.List(lngC) = !Address Then
                   cmbITIN.listindex = lngC
                   Exit For
                End If
             Next
          Case "TKT"

            For lngC = 0 To cmbTKT.ListCount
                If cmbTKT.List(lngC) = !Address Then
                   cmbTKT.listindex = lngC
                   Exit For
                End If
             Next
             If !EFormat = True Then
                    chkTktType = 1
                    chkTktType.Caption = "E-Tkt"
                Else
                    chkTktType = 0
                    chkTktType.Caption = "P-Tkt"
             End If
          Case "INV", "EINV"
      
            For lngC = 0 To cmbINV.ListCount
                If cmbINV.List(lngC) = !Address Then
                   cmbINV.listindex = lngC
                   Exit For
                End If
             Next
             
            
                
                If UCase(!Type) = "EINV" Then
                    chkInvType = 1
                    chkInvType.Caption = "E-Inv"
                Else
                    chkInvType = 0
                    chkInvType.Caption = "P-Inv"
                End If
        
          Case "MIR"
             
            For lngC = 0 To cmbMIR.ListCount
                If cmbMIR.List(lngC) = !Address Then
                   cmbMIR.listindex = lngC
                   Exit For
                End If
             Next
          Case "INVDYO"
        
            For lngC = 0 To cmbINVDYO.ListCount
                If cmbINVDYO.List(lngC) = !Address Then
                   cmbINVDYO.listindex = lngC
                   Exit For
                End If
             Next
            
          Case "ITINDYO"
             cmbITINDYO.Text = !Address
            For lngC = 0 To cmbITINDYO.ListCount
                If cmbITINDYO.List(lngC) = !Address Then
                   cmbITINDYO.listindex = lngC
                   Exit For
                End If
             Next
         
       End Select
    .MoveNext
Wend
End With

If cmbSTP.listindex < 0 Then
lblSTPlabel.Visible = False
cmbSTP.Visible = False
End If

   
   'GetDevice ("HQ")

   'If cmbITIN.ListCount > 0 Then
   '   cmbITIN.ListIndex = 0
   'End If
   'If cmbTKT.ListCount > 0 Then
   '   cmbTKT.ListIndex = 0
   'End If
   'If cmbINV.ListCount > 0 Then
   '   cmbINV.ListIndex = 0
   'End If
   'If cmbEINV.ListCount > 0 Then
   '   cmbEINV.ListIndex = 0
   'End If
   'If cmbMIR.ListCount > 0 Then
   '   cmbMIR.ListIndex = 0
   'End If
   'If cmbINVDYO.ListCount > 0 Then
   '   cmbINVDYO.ListIndex = 0
   'End If
   'If cmbITINDYO.ListCount > 0 Then
   '   cmbITINDYO.ListIndex = 0
   'End If
   
rsTkt.Close
Set rsTkt = Nothing


End Sub

'Modified on 6/1/2005: Getting company information by Profile Name instead on CN bec CN is not unique
Private Function IsEINV(ProName As String) As Boolean
   
   Dim strSQL As String
   Dim rsClients As ADODB.Recordset
   
   If gTrxnType = "L" Then
   
    strSQL = "select * from tblClients where CN = '" & gobjPNR.CN & "'"
   
   Else
   
    strSQL = "select * from tblClients where ProName = '" & ProName & "'"
   
   End If
   
   Set rsClients = gdbConn.Execute(strSQL)
   
   If rsClients.EOF Then
        IsEINV = False
        mIsEINVClient = False
   Else
        IsEINV = rsClients!EInv
        mIsEINVClient = rsClients!EInv
   End If
   
   'removed on 01/02/05
    'With grsClients
    '    .FindFirst strSQL
    '    If .NoMatch = False Then
    '        IsEINV = grsClients!EINV
    '        mIsEINVClient = grsClients!EINV
    '    Else
    '        IsEINV = False
    '        mIsEINVClient = False
    '    End If
    'End With
    

End Function

Private Function IsEItin(ProName As String) As Boolean
   
   Dim strSQL As String
   Dim rsClients As ADODB.Recordset
   
   
   If gTrxnType = "L" Then
   
     strSQL = "select * from tblClients where CN = '" & gobjPNR.CN & "'"
   
   Else
     strSQL = "select * from tblClients where ProName = '" & ProName & "'"
   End If
   Set rsClients = gdbConn.Execute(strSQL)
   
   If rsClients.EOF Then
        IsEItin = False
        'mIsEINVClient = False
   Else
        IsEItin = rsClients!EItin
       ' mIsEINVClient = rsClients!EINV
   End If
   
   'removed on 01/02/05
    'With grsClients
    '    .FindFirst strSQL
    '    If .NoMatch = False Then
    '        IsEINV = grsClients!EINV
    '        mIsEINVClient = grsClients!EINV
    '    Else
    '        IsEINV = False
    '        mIsEINVClient = False
    '    End If
    'End With
    

End Function
Private Sub txtCRInv_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 45, 46, 48 To 57
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select

End Sub

Private Sub txtCRItin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 45, 46, 48 To 57
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select

End Sub

Private Sub txtFF_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 48 To 57
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select
End Sub

Private Sub txtIATA_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 48 To 57
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select
End Sub

Private Function MakeEntry(entry As String) As String
'Modified on 10/5/2005: move previous makeentry code to SendFP, do retry for make entry error
Dim intTry As Integer
Dim strMsg As String

intTry = intTry + 1
MakeEntry = SendFP(entry)

lblStatus = "Passing Command " & entry
While bolSent = False And intTry < 4
    intTry = intTry + 1
    MakeEntry = SendFP(entry)
Wend

If bolSent <> True Then
    'MsgBox "Unable to process command: " & entry & Chr(13) & "Please toggle to focalpoint to enter/paste(Crtl+V) the entry before continue.", vbCritical, "Ticket/Invoice Make Entry"
    strMsg = "Unable to process command: " & entry & Chr(13) & "Please toggle to focalpoint to enter/paste(Crtl+V) the entry before continue."
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Clipboard.Clear
    Clipboard.setText (entry)
End If
'bolSent = False
End Function
Private Function SendFP(entry As String) As String
Dim strTemp As String
Dim lngP As Long ' used for recommended pause
Dim strPath As String
Dim intFile As Integer
Dim intRetry As Integer
Dim strRes As String

strPath = App.Path
strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
intFile = FreeFile()
Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & entry
Close #intFile

On Error GoTo Err_MakeEntry

    strRes = gobjHost.terminalEntry(entry)
'Use TerminalEntry to replace DDE
'With frmDDEOwner.txtDDE
'    pClearWindow
'    .LinkItem = "Transmit"
'    .Text = entry
'    .LinkPoke
    
'    .LinkItem = "CaptureAll"
'     .LinkRequest
'     strRes = .Text
'     If InStr(1, strRes, Chr(gintSOM)) <> 0 Then
'        strRes = Mid(strRes, 1, InStr(1, strRes, Chr(gintSOM) & Space(5)))
'     End If
'     If strRes = "" Then strRes = .Text
'    SendFP = strRes
'End With
    SendFP = strRes
    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
    Print #intFile, "Response"
    Print #intFile, strRes
    Print #intFile, ""
    Close #intFile
    For lngP = 0 To 8000000
    'recommended pause
    Next lngP
    bolSent = True
    
Exit Function

Err_MakeEntry:
    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
    Print #intFile, "MakeEntry Error: " & Err.Number & " " & Err.Description
    Print #intFile, "MakeEntry Command: " & entry
    Print #intFile, ""
    Close #intFile
    bolSent = False
    'MsgBox "Unable to process command: " & Entry, vbCritical, "Ticket/Invoice Make Entry"
    Exit Function
End Function

Private Function OldMakeEntry(ByVal entry As String) As String
Dim strTemp As String
Dim lngP As Long ' used for recommended pause
Dim strPath As String
Dim intFile As Integer
Dim intRetry As Integer
Dim strRes As String

'''
'On Error GoTo ErrMakeEntry
'''
strPath = App.Path
strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
intFile = FreeFile()
    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
    Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & entry
    Close #intFile
    entry = "<FORMAT>" & entry & "</FORMAT>"
    mobjTE.MakeEntry (entry)
    'modified on 07/09/04: wait until response received
    intRetry = 0
    
    While mobjTE.NumResponseLines = 0 And intRetry < 5
        intRetry = intRetry + 1
        Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
        Print #intFile, "Retrying... " & intRetry & " times to wait for GDS response"
        Print #intFile, ""
        Close #intFile
        Sleep (1000)
    Wend
    OldMakeEntry = ""
    For lngP = 0 To mobjTE.NumResponseLines - 1
        
        strRes = mobjTE.ResponseLine(lngP)
        If InStr(1, strRes, "<CARRIAGE_RETURN/>") > 0 Then
           OldMakeEntry = OldMakeEntry & Left(mobjTE.ResponseLine(lngP), InStr(1, strRes, "<CARRIAGE_RETURN/>") - 1) & vbCrLf
        Else
           OldMakeEntry = OldMakeEntry & mobjTE.ResponseLine(lngP) & vbCrLf
        End If
    Next
    'OldMakeEntry = mobjTE.ResponseXML
    
    Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
    Print #intFile, "Response"
    Print #intFile, OldMakeEntry
    Print #intFile, ""
    Close #intFile
    For lngP = 0 To 8000000
    'recommended pause
    Next lngP

'''
'Exit Function
'
'ErrMakeEntry:
'    If Err.Number = -2147467259 Then
'        OldMakeEntry = CONNECTION_FAIL
'    Else
'        Err.Raise Err.Number, "MakeEntry", Err.Description
'    End If
'''
    
    End Function

Private Function validData() As Boolean
Dim strMsg As String

strMsg = ""
If Len(cmbTKT.Text) <> 6 Then strMsg = "Need ticket printer address..." & vbCrLf
If Len(cmbITIN.Text) <> 6 Then strMsg = "Need itinerary printer address..." & vbCrLf
If Len(cmbINV.Text) <> 6 Then strMsg = "Need invoice printer address..." & vbCrLf
If Len(cmbMIR.Text) <> 6 Then strMsg = "Need accounting interface address..." & vbCrLf
If Len(cmbITINDYO.Text) <> 6 Then strMsg = "Need itinerary DYO number..." & vbCrLf
If Len(cmbINVDYO.Text) <> 6 Then strMsg = "Need invoice DYO number..." & vbCrLf

If strMsg <> "" Then
    'MsgBox strMsg, vbApplicationModal + vbCritical + vbOKOnly
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    validData = False
Else
    validData = True
End If

End Function

Private Sub FillInGSR()
Dim sngTotal As Single
Dim sngCX As Single
Dim lngC As Long
Dim item As ListItem
Dim strSF() As String
Dim sngSFTotal As Single
Dim strSFFOP() As String
Dim sngSFCXTotal As Single

lvwGASalesRec.ListItems.Clear
      
      
      sngTotal = 0
      sngCX = 0
      
      'Added on 27/7/05: Add to show DI SF
      For lngC = 1 To gobjPNR.AcctRemarkCount
      With gobjPNR.AcctRemark(lngC)
        If .RemarkType = "FT" And InStr(.RemarkText, "SF/*") > 0 Then
          Set item = lvwGASalesRec.ListItems.Add(, , .ItemNum)
              strSF = Split(.RemarkText, "/")
                  If UBound(strSF) = 2 Then
                    item.SubItems(1) = strSF(0)
                    item.SubItems(2) = "CWT"
                        If IsNumeric(Replace(strSF(2), "@", ".")) Then
                            sngSFTotal = sngSFTotal + Replace(strSF(2), "@", ".")
                            item.SubItems(4) = Replace(strSF(2), "@", ".")
                        End If
                  End If
        End If
        
        If .RemarkType = "FT" And InStr(.RemarkText, "FOP/*") > 0 Then
              strSFFOP = Split(.RemarkText, "/")
              If UBound(strSFFOP) > 1 Then
                  If Trim(strSFFOP(2)) <> "CASH" Then
                      If UBound(strSFFOP) = 4 Then
                          If IsNumeric(Replace(strSFFOP(4), "@", ".")) Then
                              sngSFCXTotal = sngSFCXTotal + Replace(strSFFOP(4), "@", ".")
                          End If
                      End If
                  End If
              End If
        End If
      End With
      
      Next
      
      
      For lngC = 1 To gobjPNR.GASaleRecordCount
        With gobjPNR.GASalesRecord(lngC)
            Set item = lvwGASalesRec.ListItems.Add(, , .BegLine & "-" & .EndLine)
            item.SubItems(1) = .ProductCode
            item.SubItems(2) = .VendorCode
            item.SubItems(3) = .BaseAmount
            'modified on 16Jun: request from sharon, add back GST, GARecord do not include GST to telly with due line
            item.SubItems(4) = .SellAmount + .GSTAmount + .Tax
            sngTotal = sngTotal + .SellAmount + .GSTAmount + .Tax
            'If (Left(.FOP, 2) = "CX" Or Left(.FOP, 2) = "CC") And Not (Left(.CCNumber, 2) = "DC" And Mid(.CCNumber, 3, 7) = "3644033") Then
             'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
            If (Left(.FOP, 2) = "CX" Or Left(.FOP, 2) = "CC") And Not _
             IsTMPCard(Left(.CCNumber, 2), Mid(.CCNumber, 3)) Then
                sngCX = sngCX + .SellAmount + .GSTAmount + .Tax
            End If
        End With
    Next
    lblGASRTtotal.Caption = CStr(Format(sngTotal + sngSFTotal, gstrAgcyCurrFormat))
    lblCXTotal.Caption = CStr(Format(sngCX + sngSFCXTotal, gstrAgcyCurrFormat))
End Sub

Private Sub FillInPaidDue()
Dim sngTotal As Single
Dim sngPaid As Single
Dim lngC As Long
Dim item As ListItem

lvwPaid.ListItems.Clear
lvwDue.ListItems.Clear

sngTotal = 0
sngPaid = 0
    For lngC = 1 To gobjPNR.PaidDueCount
        With gobjPNR.PaidDue(lngC)
            Select Case .SegType
                Case "P"
                    Set item = lvwPaid.ListItems.Add(, , .SegNum)
                    item.SubItems(1) = .FreeText
                    item.SubItems(2) = .Amount
                    sngPaid = sngPaid + .Amount
                Case "D"
                    Set item = lvwDue.ListItems.Add(, , .SegNum)
                    item.SubItems(1) = .FreeText
                    item.SubItems(2) = .Amount
                    If .FreeText <> "**INVOICE TOTAL DUE**" And .FreeText <> "**TOTAL AMT CHARGE TO CREDIT CARD**" Then
                       sngTotal = sngTotal + .Amount
                    End If
            End Select
        End With
    Next
    lblPaidTotal.Caption = Format(sngPaid, gstrAgcyCurrFormat)
    lblDueTotal.Caption = Format(sngTotal, gstrAgcyCurrFormat)
End Sub

Private Sub FillInFiledFares()
Dim sngTotal As Single
Dim lngC As Long
Dim item As ListItem
Dim intPx As Integer
Dim sngTaxTotal As Single
      sngTotal = 0
      
      For lngC = 1 To gobjPNR.FiledFareCount
        For intPx = 1 To gobjPNR.FiledFare(lngC).PxCount
        With gobjPNR.FiledFare(lngC).PX(intPx)
            Set item = lvwFiledFares.ListItems.Add(, , lngC) 'filed fare number
            item.SubItems(1) = IIf(.FOPType = "CC", .FOP_CCCode, .FOPType)
            
            If .BaseCurrency = .TotalCurrency Then
                item.SubItems(2) = .BaseAmount
            Else
                item.SubItems(2) = .EquivAmount
            End If

            'Item.SubItems(2) = .BaseAmount
            item.SubItems(3) = .TaxTotal
            item.SubItems(4) = IIf(.SellAmount = 0, .TotAmount, .SellAmount)
            sngTotal = sngTotal + IIf(.SellAmount = 0, .TotAmount, .SellAmount)
            sngTaxTotal = sngTaxTotal + .TaxTotal
        End With
        Next intPx
    Next
    lblFFTotal.Caption = Format(sngTotal, gstrAgcyCurrFormat)
    lblTaxTotal.Caption = Format(sngTaxTotal, gstrAgcyCurrFormat)
End Sub

'29122004
Private Function ChangeTktNum() As Boolean
 Dim intNT As Integer
 Dim intPx As Integer
 Dim i As Integer
 Dim strRmks() As String
 Dim intFFNum As Integer
 Dim lngPos As Long
 Dim strTemp As String
 Dim strRes As String
 Dim strDIFF As String
 Dim strDIPx As String
 Dim strCSTktNum As String
 Dim strMsg As String
 

With gobjPNR
ChangeTktNum = True
For intNT = 1 To .FiledFareCount
  For intPx = 1 To .FiledFare(intNT).PxCount
    If .FiledFare(intNT).PX(intPx).TicketNumber <> "" Then
        For i = 1 To .GASaleRecordCount ' ALL MS LINE
            strDIFF = .GASalesRecord(i).TicketNumber 'FF01
            If .GASalesRecord(i).BegLine > 0 Then strDIPx = .AcctRemark(.GASalesRecord(i).BegLine).RemarkText 'PX1
            strDIPx = Mid(strDIPx, InStr(1, strDIPx, "/PX") + 1)
            If .AcctRemarkCount > 1 And .GASalesRecord(i).BegLine > 0 Then
               If "FF" & Format(intNT, "00") = strDIFF And "PX" & .FiledFare(intNT).PX(intPx).PassengerNum = strDIPx Then
                  lngPos = .GASalesRecord(i).BegLine
                  'modified on 27/7/05
                  
                  strCSTktNum = CSTktNumFormat(.FiledFare(intNT).PX(intPx).TicketNumber)
                  'strTemp = "DI." & .GASalesRecord(i).BegLine & "@FT-" & Left(.AcctRemark(lngPos).RemarkText, InStr(.AcctRemark(lngPos).RemarkText, "/TK") + 2) & fConvertTkTNo(.FiledFare(intNT).PX(intPx).TicketNumber)
                  'strTemp = "DI." & .GASalesRecord(i).BegLine & "@FT-" & Left(.AcctRemark(lngPos).RemarkText, InStr(.AcctRemark(lngPos).RemarkText, "/TK") + 2) & .FiledFare(intNT).PX(intPx).TicketNumber & "/" & strDIPx
                  strTemp = "DI." & .GASalesRecord(i).BegLine & "@FT-" & Left(.AcctRemark(lngPos).RemarkText, InStr(.AcctRemark(lngPos).RemarkText, "/TK") + 2) & strCSTktNum & "/" & strDIPx
                  MakeEntry strTemp
               End If
            End If
        Next i
    ElseIf .FiledFare(intNT).PX(.FiledFare(intNT).PxCount).TicketNumber = "" Then
    'Else
        ChangeTktNum = False
        strMsg = "Still need to wait for ticket numbers"
        modMsgBox.OKMsg = "OK"
        modMsgBox.CANCELMsg = "Cancel"
        If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbOKCancel + vbDefaultButton2, "CWT Desktop") = vbOK Then
        'If MsgBox("Still need to wait for ticket numbers", vbApplicationModal + vbOKCancel) = vbOK Then
            strRes = MakeEntry("IR")
            Timer1.Enabled = True
            Exit Function
        Else
            Exit Function
        End If
    End If
'NextFF:
Next intPx
Next
End With
ChangeTktNum = True
End Function
Private Function DisplayNum() As Integer
 Dim i As Integer
 Dim strRemark() As String
 Dim intTmp As Integer
 
With gobjPNR

For i = 1 To .AcctRemarkCount
   strRemark = Split(.AcctRemark(i).RemarkText, "/")
   If UBound(strRemark) >= 1 Then
      If Left(strRemark(1), 1) = "*" And IsNumeric(Mid(strRemark(1), 2)) Then
         If Mid(strRemark(1), 2) > intTmp Then
            intTmp = Mid(strRemark(1), 2)
         End If
      End If
   End If
Next

End With
DisplayNum = intTmp
End Function

Private Sub txtPax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 48 To 57
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select
End Sub


Private Sub FillLocation()
Dim strSQL As String
Dim rsTkt As ADODB.Recordset
   'Added on 22/4/2005: implant selection to populate devices & DYO
   strSQL = "Select distinct(Loc) from tblDeviceAddr"
 
   Set rsTkt = gdbConn.Execute(strSQL)
   While Not rsTkt.EOF
              cmbSTPLoc.AddItem rsTkt!Loc
              'Set item = lsvSTP.ListItems.Add(, , rsTKT!Address)
              'item.SubItems(1) = rsTKT!IATA & ""
   rsTkt.MoveNext
   Wend
   
   'If cmbSTPLoc.ListCount > 0 Then
   '   cmbSTPLoc.ListIndex = 0
   'End If
   cmbSTPLoc.Text = "HQ"
   rsTkt.Close
   Set rsTkt = Nothing
End Sub

Private Function SplitCmd() As Boolean
Dim strSQL As String
Dim rsTkt As ADODB.Recordset

strSQL = "Select * from tblDeviceAddr where Loc='" & cmbSTPLoc & "' AND Address='" & cmbTKT & "'"
Set rsTkt = gdbConn.Execute(strSQL)
If Not rsTkt.EOF Then
   If rsTkt!SplitTktItin = True Then
    SplitCmd = True
   Else
    SplitCmd = False
   End If
Else
   SplitCmd = False
End If
rsTkt.Close
Set rsTkt = Nothing

End Function
Private Function CompareRDDI() As Boolean
Dim intresponse As Integer
Dim sngTotDI As Single
Dim sngTotRD As Single
Dim strMsg As String

sngTotDI = fConvertZero(lblGASRTtotal) + fConvertZero(lblTaxTotal)
sngTotRD = fConvertZero(lblDueTotal)
    If sngTotDI <> sngTotRD Then
        strMsg = "Total RD Amount(" & sngTotDI & ") is not telly with Total DI Amount(" & sngTotRD & ")" & Chr(13) & _
                  "Do you want to continue?" & Chr(13)
        modMsgBox.YESMsg = "Yes"
        modMsgBox.NOMsg = "No"
        intresponse = modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton1, "CWT Desktop")
        'intresponse = MsgBox("Total RD Amount(" & sngTotDI & ") is not telly with Total DI Amount(" & sngTotRD & ")" & Chr(13) & _
                  "Do you want to continue?" & Chr(13), vbYesNo + vbExclamation + vbDefaultButton1, "CWT TravelPro - Ticketing")
        If intresponse = vbYes Then
                CompareRDDI = True
        Else
                CompareRDDI = False
        End If
    Else
            CompareRDDI = True
    End If
End Function

Private Function CSTktNumFormat(GalTktNum As String) As String
   Dim strTmp1 As String
   Dim strTmp2 As String
   
   If InStr(1, GalTktNum, "-") <> 0 Then
      strTmp1 = Mid(GalTktNum, 1, InStr(1, GalTktNum, "-"))
      strTmp2 = Mid(GalTktNum, InStr(1, GalTktNum, "-") + 1)
      If Len(strTmp2) > 2 Then
         strTmp2 = Right(strTmp2, 2)
      End If
      CSTktNumFormat = strTmp1 & strTmp2
   Else
      CSTktNumFormat = GalTktNum
   End If
End Function
'Added on 290306: CPG
Private Sub DoAutoApproval()
Dim strMsg As String
Dim intresponse As Integer
Dim lngInvNum As Long
Dim i As Integer
Dim strRes As String
Dim CCAppStart As Date
Dim intFile As Integer
CCAppStart = Now

        sumCXFOP
        If lsvCPG.ListItems.Count > 0 Then
            strMsg = "Do you want to continue to send Credit Payment for invoice amount: " & vbCrLf
           
            'For i = 1 To lsvCPG.ListItems.count
                strMsg = strMsg & "1. " & lsvCPG.ListItems(1).Text & lsvCPG.ListItems(1).SubItems(1) & " - " & gstrAgcyCurrCode & " " & lsvCPG.ListItems(1).SubItems(2)
            'Next
            modMsgBox.YESMsg = "Yes"
            modMsgBox.NOMsg = "No"
            intresponse = modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbYesNo + vbDefaultButton2, "CWT Desktop")
            ' intresponse = MsgBox(strMsg, vbYesNo, "CPG - Credit Card Auto Approval")
            If intresponse = vbYes Then
                     intFile = FreeFile()
                     Open App.Path & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
                     Print #intFile, Now & Space(5) & "Start of CitiBank Approval Process"
                     Close #intFile
            
                 'have to retrieve pnr again
                    strRes = MakeEntry("*" & gobjPNR.RecLoc)
                    If InStr(1, strRes, "FINISH OR IGNORE") <> 0 Or _
                        InStr(1, strRes, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
                        Sleep (1000)
                        MakeEntry "I"
                        Sleep (500)
                        strRes = MakeEntry("*" & gobjPNR.RecLoc)
                    End If
                    
                 lngInvNum = InvoiceNum
                 If lngInvNum = 0 Then
                    'MsgBox "Unable to Proceed. System cannot capture invoice number."
                    strMsg = "Unable to Proceed. System cannot capture invoice number."
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                    Open App.Path & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
                    Print #intFile, Now & Space(5) & "Error: " & "Unable to Proceed. System cannot capture invoice number."
                    Close #intFile
                    Exit Sub
                 End If
                
                'For i = 1 To lsvCPG.ListItems.count
                    
                    'doCCSale 1, lngInvNum, lsvCPG.ListItems(1).Text, lsvCPG.ListItems(1).SubItems(1), lsvCPG.ListItems(1).SubItems(3), lsvCPG.ListItems(1).SubItems(2)
                    doNewCCSale 1, lngInvNum, lsvCPG.ListItems(1).Text, lsvCPG.ListItems(1).SubItems(1), lsvCPG.ListItems(1).SubItems(3), lsvCPG.ListItems(1).SubItems(2)
                'Next i
                pCloseHandles
                strMsg = ""
                For i = 1 To lsvCPG.ListItems.Count
                    strMsg = strMsg & "Transaction for " & lsvCPG.ListItems(i).Text & lsvCPG.ListItems(i).SubItems(1) & " "
                    strMsg = strMsg & "Amount: " & gstrAgcyCurrCode & " " & lsvCPG.ListItems(i).SubItems(2) & vbCrLf
                    strMsg = strMsg & "_____________________________________________" & vbCrLf
                    If lsvCPG.ListItems(i).SubItems(8) <> "" Then
                    strMsg = strMsg & "Process Indicator:" & lsvCPG.ListItems(i).SubItems(4) & vbCrLf
                    strMsg = strMsg & "Transaction ID:" & lsvCPG.ListItems(i).SubItems(7) & vbCrLf
                    strMsg = strMsg & "Status Code from Citibank:" & lsvCPG.ListItems(i).SubItems(5) & vbCrLf
                    strMsg = strMsg & "Status from Citibank:" & lsvCPG.ListItems(i).SubItems(6) & vbCrLf
                    'strMsg = strMsg & "ReceiptNo:" & lsvCPG.ListItems(i).SubItems(10) & vbCrLf
                    strMsg = strMsg & "Status Code from PaymentServer:" & lsvCPG.ListItems(i).SubItems(10) & vbCrLf
                    strMsg = strMsg & "Status from PaymentServer:" & lsvCPG.ListItems(i).SubItems(11) & vbCrLf

                    Else
                    strMsg = strMsg & "Process Indicator:" & lsvCPG.ListItems(i).SubItems(4) & vbCrLf
                    'strMsg = strMsg & "Error:" & lsvCPG.ListItems(i).SubItems(9) & vbCrLf
                    strMsg = strMsg & "Status Code from PaymentServer:" & lsvCPG.ListItems(i).SubItems(10) & vbCrLf
                    strMsg = strMsg & "Status from PaymentServer:" & lsvCPG.ListItems(i).SubItems(11) & vbCrLf
                    strMsg = strMsg & "Status Code from Citibank:" & lsvCPG.ListItems(i).SubItems(5) & vbCrLf
                    strMsg = strMsg & "Status from Citibank:" & lsvCPG.ListItems(i).SubItems(6) & vbCrLf

                    End If
                Next
                
                'MsgBox strMsg, vbOKOnly, "CPG - Credit Card Auto Approval"
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop"
                Call pAddToVMSUTrans(CStr(lngInvNum))
                Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", CCAppStart, CCAppStart, "CreditApproval", , startTime)
                

            Else
                Exit Sub
            End If
        End If
    
End Sub
Private Sub pAddToVMSUTrans(inv As String)
Dim strSQL As String
With lsvCPG.ListItems(1)

    strSQL = "Insert into tblVMSUTRAN (RECLOC,INVNO,INVFROM,OPSTYPE,CARDVENDOR,CARDNO,CARDEXP,TRANID,TRANAMT,STATUS,CREATEDDATE,RECEIPTNO,ERROR,QSICODE,GDS)"
    strSQL = strSQL & "VALUES('" & gobjPNR.RecLoc & "','" & inv & "','Ticketing','SALE',"
    strSQL = strSQL & "'" & .Text & "','" & .SubItems(1) & "','" & .SubItems(3) & "',"
    strSQL = strSQL & "'" & .SubItems(7) & "','" & .SubItems(2) & "','" & .SubItems(5) & "',"
    strSQL = strSQL & "'" & Now & "','" & .SubItems(9) & "', '" & .SubItems(8) & "','" & .SubItems(10) & "','GAL')"
    gdbConn.Execute strSQL
End With
End Sub

'Private Sub doCCSale(no As Integer, invno As Long, ccvendor As String, ccno As String, ccexp As String, ccamt As Single)
'    Dim iRetVal     As Integer
'    Dim sBuffer     As String * 1024
'    Dim vDllVersion As tWinInetDLLVersion
'    Dim sStatus     As String
'    Dim sOptionBuffer   As String
'    Dim lOptionBufferLen As Long
'    Dim SecFlag As Long
'    Dim dwSecFlag As Long
'    Dim dwPort As Long
'    Dim intFile As Integer
'
'    Dim sOpType As String
    
'    Dim strURL As String
'    Dim strSuccURL As String
'    Dim strFailURL As String
'    Dim strTransCurr As String
'    Dim strSKeyfile As String
'    Dim strSessionNum As String

    
    'frmWait.Show , Me
    
    'strURL = "cardmerchantuat.citibank.com.sg/servlet/MSLProcessor"
'    strURL = pGetValue("URL", "URL")
'    strSuccURL = "success.html"
'    strFailURL = "failure.html"
'    strTransCurr = "CWTSIN_SGD"
'    sOpType = "SALE"
'    intFile = FreeFile()
'    Open App.Path & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile

    
'    pCloseHandles
    
    
'    hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

'    hInternetConnect = InternetConnect(hInternetSession, fGetServerURL(strURL), dwPort, _
'                       vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)

'    Screen.MousePointer = vbHourglass
    
'    If CBool(hInternetSession) Then
'        InternetQueryOption hInternetSession, INTERNET_OPTION_VERSION, vDllVersion, Len(vDllVersion)
        
        'Debug.Print "Establishing secure connection" & " "
'        dwPort = INTERNET_DEFAULT_HTTPS_PORT
        'Debug.Print "Setting security flags" & " "
'        SecFlag = INTERNET_FLAG_SECURE Or _
                  INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
                  INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
        
        
'        hInternetConnect = InternetConnect(hInternetSession, fGetServerURL(strURL), dwPort, _
                               vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
        
'        If hInternetConnect > 0 Then
           
'            sOptionBuffer = vbNullString
'            lOptionBufferLen = 0

'            strSKeyfile = App.Path & "\" & pGetValue("SessionKey", "Ticketing")
'            strSessionNum = fGetSession(strSKeyfile)
            
'            Print #intFile, Now & Space(5) & "SessionID: " & strSessionNum
            
'            If strSessionNum = "" Then

'                Screen.MousePointer = vbDefault
'                Print #intFile, Now & Space(5) & "Error: " & "Unable to capture SessionID"

'                Exit Sub
'            End If
            
            'If LCase(sOpType) = "rsale" Then
            'sOptionBuffer = "sid=" & lblSessID.Caption & "&orderno=" & Trim(txtOrderNo.Text) & _
            '                "&op=" & sOpType & "&successURL=" & gstrSuccURL & _
            '                "&failureURL=" & gstrFailURL & _
            '                "&tid=" & Trim(txtTID.Text) & "&dba=" & gstrTransCurr & "&amount=" & Trim(txtAmt)
            'Else
 '           sOptionBuffer = "sid=" & strSessionNum & "&orderno=" & invno & _
                            "&op=" & sOpType & "&successURL=" & strSuccURL & _
                            "&failureURL=" & strFailURL & "&ccn=" & ccno & _
                            "&brand=" & IIf(ccvendor = "VI", "VISA", "MASTERCARD") & "&exp=" & ccexp & _
                            "&amount=" & ccamt & "&dba=" & strTransCurr
 '           Print #intFile, Now & Space(5) & "sOptionBuffer: " & sOptionBuffer

            'End If
            'If LCase(sOpType) = "rsale" Or LCase(sOpType) = "rauth" Then
            '    If txtTID.Text <> "" Then
            '        sOptionBuffer = sOptionBuffer & "&tid=" & txtTID.Text
            '    Else
            '        MsgBox "Trans ID must be entered.", vbCritical
            '        Screen.MousePointer = vbDefault
            '        Unload frmWait
            '        Exit Sub
            '    End If
            'End If
 '           lOptionBufferLen = Len(sOptionBuffer)
            
 '           hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", fGetURLPath(strURL) & "?" & sOptionBuffer, "HTTP/1.1", vbNullString, 0, _
 '               INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION Or SecFlag, 0)
           
 '           If CBool(hHttpOpenRequest) Then
                'Debug.Print sOptionBuffer
 '               Dim sHeader As String
                
                'sHeader = "Accept-Encoding: deflate, gzip" & vbCrLf
                'iRetVal = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
                'Debug.Print iRetVal & " " & Len(sHeader)
                                
 '               Dim dwTimeOut As Long
 '               dwTimeOut = 180000 ' time out is set to 3 minutes
 '               iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_CONNECT_TIMEOUT, _
 '           dwTimeOut, 4)
                'Debug.Print iRetVal & " " & Err.LastDllError & " " & "INTERNET_OPTION_CONNECT_TIMEOUT"
 '               iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, _
 '           dwTimeOut, 4)
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_RECEIVE_TIMEOUT"
 '               iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SEND_TIMEOUT, _
 '           dwTimeOut, 4)
                'Debug.Print iRetVal & " " & "INTERNET_OPTION_SEND_TIMEOUT"
                
'Resend:
'                iRetVal = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, vbNullString, 0)
                
'                Print #intFile, Now & Space(5) & "iRetVal: " & iRetVal
'                 If (iRetVal <> 1) And (Err.LastDllError = 12045) Then
'                    MsgBox "Invalid CA"
'                    'Certificate Authority is invalid.
'                    'Debug.Print "Invalid Cert Auth, resending" & " "
'                    dwSecFlag = SECURITY_FLAG_IGNORE_UNKNOWN_CA
'                    iRetVal = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SECURITY_FLAGS, dwSecFlag, 4)
'                    'Debug.Print iRetVal & " " & Err.LastDllError & " " & "INTERNET_OPTION_SECURITY_FLAGS"
'                    Print #intFile, Now & Space(5) & "Error: " & Err.LastDllError
'
'                    GoTo Resend
'                End If
'
'                If iRetVal Then
'                    Dim dwStatus As Long, dwStatusSize As Long
'                    dwStatusSize = Len(dwStatus)
'                    'HttpQueryInfo hInternetURLSession, HTTP_QUERY_FLAG_NUMBER Or HTTP_QUERY_STATUS_CODE, dwStatus, dwStatusSize, 0
'                    HttpQueryInfo hHttpOpenRequest, HTTP_QUERY_FLAG_NUMBER Or HTTP_QUERY_STATUS_CODE, dwStatus, dwStatusSize, 0
'                    Select Case dwStatus
'                    Case HTTP_STATUS_PROXY_AUTH_REQ
'                        iRetVal = InternetSetOptionStr(hHttpOpenRequest, INTERNET_OPTION_PROXY_USERNAME, _
'                            "IUSR_WEIHUA1", Len("IUSR_WEIHUA1") + 1)
'                        iRetVal = InternetSetOptionStr(hHttpOpenRequest, INTERNET_OPTION_PROXY_PASSWORD, _
'                            "IUSR_WEIHUA1", Len("IUSR_WEIHUA1") + 1)
'                        GoTo Resend
'                    Case HTTP_STATUS_DENIED
'                        iRetVal = InternetSetOptionStr(hHttpOpenRequest, INTERNET_OPTION_USERNAME, _
'                            "IUSR_WEIHUA1", Len("IUSR_WEIHUA1") + 1)
'                        iRetVal = InternetSetOptionStr(hHttpOpenRequest, INTERNET_OPTION_PASSWORD, _
'                            "IUSR_WEIHUA1", Len("IUSR_WEIHUA1") + 1)
'                        GoTo Resend
'                    End Select
'                'Dim httpResponse1 As Object
'                    'response headers
'                    GetQueryInfo hHttpOpenRequest, no, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_RAW_HEADERS_CRLF
'                    'lsvCPG.ListItems(no).SubItems(6) = httpResponse1
'                    'sStatus = "Ready"
'                Else
'                    ' HttpSendRequest failed
'                    sStatus = "HttpSendRequest call failed; Error code: " & Err.LastDllError & "."
'                End If
'            Else
'                ' HttpOpenRequest failed
'                sStatus = "HttpOpenRequest call failed; Error code: " & Err.LastDllError & "."
'            End If
'        Else
'            ' InternetConnect failed
'            sStatus = "InternetConnect call failed; Error code: " & Err.LastDllError & "."
'        End If
'    Else
        ' hInternetSession handle not allocated
'        sStatus = "InternetOpen call failed: Error code: " & Err.LastDllError & "."
'    End If
'    If sStatus <> "" Then
'        lsvCPG.ListItems(no).SubItems(8) = sStatus
'        lsvCPG.ListItems(no).SubItems(4) = "ERROR"
'        Print #intFile, Now & Space(5) & "Error: " & sStatus
'    Else
'    Print #intFile, Now & Space(5) & "Host Response: " & lsvCPG.ListItems(no).SubItems(6)
'    lsvCPG.ListItems(no).SubItems(5) = fGetStatusCode(lsvCPG.ListItems(no).SubItems(6))
'    lsvCPG.ListItems(no).SubItems(4) = IIf(fGetSuccessInd(lsvCPG.ListItems(no).SubItems(6)), "PASS", "FAIL")
'    lsvCPG.ListItems(no).SubItems(7) = fGetTID(lsvCPG.ListItems(no).SubItems(6))
'    lsvCPG.ListItems(no).SubItems(6) = getAuthDesc(lsvCPG.ListItems(no).SubItems(5))
'    Print #intFile, Now & Space(5) & "Response Description: " & lsvCPG.ListItems(no).SubItems(6)
    
'    End If
    
'    Screen.MousePointer = vbDefault
                        
'    Close #intFile
'End Sub

Private Sub pCloseHandles()
    InternetCloseHandle (hHttpOpenRequest)
    InternetCloseHandle (hInternetSession)
    InternetCloseHandle (hInternetConnect)

End Sub
'Private Function fGetSession(file As String) As String
   
'    Dim sid As String
'    Dim sidGen As SESSIONKEYLib.Generator
'    Set sidGen = New SESSIONKEYLib.Generator
    
'    sid = sidGen.generate(file)
'    sid = Mid(sid, 1, 32)
'    fGetSession = ""
'    If isHex(sid) Then
'        fGetSession = sid
'    Else
'        MsgBox "Unable to get SessionID. Response from host: " & sid
'    End If
'    Set sidGen = Nothing
    
    
'End Function

Private Function getAuthDesc(auth As String, responsefrom As String) As String
    Dim myRst As ADODB.Recordset
    On Error GoTo errDB
 
     Set myRst = gdbConn.Execute("select authcode, authdesc from tblVMSULookup where authcode='" & auth & "' AND ResponseFrom='" & responsefrom & "' ")
    getAuthDesc = ""
    If Not myRst.EOF Then
        getAuthDesc = myRst!AuthDesc
    End If
    Exit Function
errDB:
    getAuthDesc = ""

End Function
Private Function pGetValue(setupfor As String, soption As String) As String

Dim rs As ADODB.Recordset
Dim strMsg As String

On Error GoTo errDB
    
    Set rs = gdbConn.Execute("select svalue from tblVMSUSetup where soption='" & soption & "' and setup='" & setupfor & "'")
    
    If Not rs.EOF Then
        pGetValue = rs!svalue
    End If
    Exit Function
    
errDB:
    'MsgBox Err.Number & Err.Description
   strMsg = Err.Number & Err.Description
   modMsgBox.OKMsg = "OK"
   modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End Function
Private Function fGetServerURL(URL As String) As String

    If InStr(URL, "/") <> 0 Then
        fGetServerURL = Left(URL, InStr(URL, "/") - 1)
    Else
        fGetServerURL = URL
    End If
    
End Function
Private Function fGetURLPath(URL As String) As String
    
    If InStr(URL, "/") <> 0 Then
    fGetURLPath = Right(URL, Len(URL) - InStr(URL, "/") + 1)
    Else
    fGetURLPath = ""
    End If

End Function

Private Function isHex(hexval As String) As Boolean
    Dim i As Integer
    Dim myChar As String
    For i = 1 To Len(hexval)
        myChar = LCase(Mid(hexval, i, 1))
        If Asc(myChar) <> 10 And Asc(myChar) <> 13 Then
            Select Case myChar
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"
                Case Else
                    isHex = False
                    Exit Function
            End Select
        End If
    Next
    isHex = True
End Function
Private Function fGetTID(resp As String) As String
    Dim sSplitVal() As String
    Dim i As Integer
    
    On Error GoTo errTID
    fGetTID = ""
    If fGetSuccessInd(resp) Then
        sSplitVal = Split(resp, "&")
        For i = 0 To UBound(sSplitVal)
            If InStr(1, sSplitVal(i), "tid=") > 0 Then
                fGetTID = Mid(sSplitVal(i), InStr(1, sSplitVal(i), "tid=") + 4, InStr(1, sSplitVal(i), " ") - 4)
                Exit For
            End If
        Next
    End If
    Exit Function
errTID:
    fGetTID = ""
End Function
Private Function fGetStatusCode(resp As String) As String
    Dim sSplitVal() As String
    Dim i As Integer
    
    On Error GoTo errStatusCode
    fGetStatusCode = ""
    sSplitVal = Split(resp, "&")
    For i = 0 To UBound(sSplitVal)
        If InStr(1, sSplitVal(i), "authcode=") > 0 Then
            fGetStatusCode = Mid(sSplitVal(i), InStr(1, sSplitVal(i), "=") + 1)
            Exit For
        End If
    Next
    
    Exit Function
errStatusCode:
    fGetStatusCode = ""
    'pCloseHandles
End Function
Private Function fGetSuccessInd(resp As String) As Boolean
    Dim sSplitVal() As String
    On Error GoTo errSuccessInd
    fGetSuccessInd = False
    sSplitVal = Split(resp, "?")
    fGetSuccessInd = IIf(Mid(sSplitVal(0), InStr(1, sSplitVal(0), "servlet/") + 8) = "success.html", True, False)
    Exit Function
errSuccessInd:
    fGetSuccessInd = False
    'pCloseHandles
End Function

Private Function InvoiceNum() As Long
Dim strResponse As String
Dim strTemp As String
Dim intI As Integer
Dim strInv As String
Dim intLength As Integer
Dim intFile As Integer

Dim mxmldomINV As MSXML2.DOMDocument
Dim xmlnlINV As MSXML2.IXMLDOMNodeList


strTemp = ""
strTemp = strTemp & "<DocProdFareManipulation_4_0>"
strTemp = strTemp & "<TicketNumbersMods/>"
strTemp = strTemp & "</DocProdFareManipulation_4_0>"



Set mxmldomINV = New MSXML2.DOMDocument
Set mxmldomINV = CreateObject("microsoft.xmldom")
mxmldomINV.async = False
If mxmldomINV.loadXML(gobjHost.SendQuery(strTemp, "DocProdFareManipulation_4_2", "frmTktInfo", "InvoiceNum")) = False Then Exit Function
    

    
Set xmlnlINV = mxmldomINV.selectNodes("//TicketNumberData/ETicketNum/ItinInvNum")
intLength = xmlnlINV.length
If intLength > 0 Then
    InvoiceNum = xmlnlINV.item(intLength - 1).Text
Else
    InvoiceNum = 0
End If

    intFile = FreeFile()
    Open App.Path & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
    Print #intFile, Now & Space(5) & "DocProdFareManipulation_4_0 Response:"
    Print #intFile, Now & Space(5) & mxmldomINV.xml
    Print #intFile, Now & Space(5) & "Length: " & intLength
    Print #intFile, Now & Space(5) & "Invoice Number: " & InvoiceNum
Close #intFile

End Function

Private Sub sumCXFOP()
Dim intI As Integer
Dim intJ As Integer
Dim intK As Integer
Dim lngCXAmt As Single
Dim item As ListItem
lngCXAmt = 0
'For intI = 1 To gobjPNR.AirFaresFOPCount
'With gobjPNR.AirFaresFOP(intI)
'     If InStr(.FOPType, "CX") > 0 _
'     And (.FOP_CCCode = "VI" Or .FOP_CCCode = "CA") _
'     And CCExist(.FOP_CCCode, .FOP_CCNum) = False _
'     And IsFOPCC(.FOP_CCCode, .FOP_CCNum) Then
      
'      For intJ = 1 To gobjPNR.AirFaresFOPCount
'        If .FOP_CCNum = gobjPNR.AirFaresFOP(intJ).FOP_CCNum _
'        And .FOP_CCCode = gobjPNR.AirFaresFOP(intJ).FOP_CCCode _
'        And Left(.FOPType, 2) = Left(gobjPNR.AirFaresFOP(intJ).FOPType, 2) Then _
'            lngCXAmt = lngCXAmt + .FOPAmount
'      Next
      
'      For intK = 1 To gobjPNR.GASaleRecordCount
'        With gobjPNR.GASalesRecord(intK)
'            If gobjPNR.AirFaresFOP(intI).FOP_CCNum = Mid(.CCNumber, 3) _
'            And gobjPNR.AirFaresFOP(intI).FOP_CCCode = Left(.CCNumber, 2) _
'            And Left(gobjPNR.AirFaresFOP(intI).FOPType, 2) = Left(.FOP, 2) Then
'                lngCXAmt = lngCXAmt + .SellAmount + .GSTAmount + .Tax
'            End If
'        End With
'      Next
'
'        Set Item = lsvCPG.ListItems.Add(, , .FOP_CCCode)
'        Item.SubItems(1) = .FOP_CCNum
'        Item.SubItems(2) = lngCXAmt
'        Item.SubItems(3) = Format(gobjPNR.FOP_CCExpireDate, "yyyymm")
'    End If
'End With
'Next

'in cases where other pdt FOP is diff. from air FOP
'lngCXAmt = 0

'For intI = 1 To gobjPNR.AirFaresFOPCount
'With gobjPNR.AirFaresFOP(intI)

     'If InStr(.FOPType, "CX") > 0 _
     'And (.FOP_CCCode = "VI" Or .FOP_CCCode = "CA") _
     'And IsFOPCC(.FOP_CCCode, .FOP_CCNum) Then
      'If (gobjPNR.FOP_CCCode = "VI" Or gobjPNR.FOP_CCCode = "CA") Then
      'And IsFOPCC(.FOP_CCCode, .FOP_CCNum) Then
      
      
      For intJ = 1 To gobjPNR.AirFaresFOPCount
        'If gobjPNR.FOP_CCNum = gobjPNR.AirFaresFOP(intJ).FOP_CCNum _
        'And gobjPNR.FOP_CCCode = gobjPNR.AirFaresFOP(intJ).FOP_CCCode
        If IsFOPCC(gobjPNR.AirFaresFOP(intJ).FOP_CCCode, gobjPNR.AirFaresFOP(intJ).FOP_CCNum) _
        And (gobjPNR.AirFaresFOP(intJ).FOP_CCCode = "VI" Or gobjPNR.AirFaresFOP(intJ).FOP_CCCode = "CA") _
        And InStr(UCase(gobjPNR.AirFaresFOP(intJ).FOPType), "CX") = 1 Then
            lngCXAmt = lngCXAmt + gobjPNR.AirFaresFOP(intJ).FOPAmount
        End If
        
      Next
    
     'End If

 'End With
'Next


For intK = 1 To gobjPNR.GASaleRecordCount

     With gobjPNR
            If Left(.GASalesRecord(intK).FOP, 2) = "CX" _
            And (Left(.GASalesRecord(intK).CCNumber, 2) = "VI" Or Left(.GASalesRecord(intK).CCNumber, 2) = "CA") _
            And IsFOPCC(Left(.GASalesRecord(intK).CCNumber, 2), Mid(.GASalesRecord(intK).CCNumber, 3)) Then

            'And CCExist(Left(.GASalesRecord(intK).CCNumber, 2), Mid(.GASalesRecord(intK).CCNumber, 3)) = False
                                'For intJ = 1 To gobjPNR.GASaleRecordCount
                    'If Mid(.GASalesRecord(intK).CCNumber, 3) = Mid(.GASalesRecord(intJ).CCNumber, 3) _
                    '    And Left(.GASalesRecord(intK).CCNumber, 2) = Left(.GASalesRecord(intJ).CCNumber, 2) _
                    '    And Left(.GASalesRecord(intK).FOP, 2) = Left(.GASalesRecord(intJ).FOP, 2) Then
                            'modified on 110806
                            'lngCXAmt = lngCXAmt + .GASalesRecord(intK).SellAmount + .GASalesRecord(intK).GSTAmount + .GASalesRecord(intK).Tax
                            lngCXAmt = lngCXAmt + .GASalesRecord(intK).CollectedAmount
                    'End If
                    'Next
            End If
    End With
Next


            Set item = lsvCPG.ListItems.Add(, , gobjPNR.FOP_CCCode)
                item.SubItems(1) = gobjPNR.FOP_CCNum
                item.SubItems(2) = Format(lngCXAmt, gstrAgcyCurrFormat)
                item.SubItems(3) = Format(gobjPNR.FOP_CCExpireDate, "yymm")
            

End Sub


Private Function CCExist(ccvendor As String, CCNum As String) As Boolean
Dim intI As Integer
CCExist = False
If lsvCPG.ListItems.Count = 0 Then Exit Function
For intI = 1 To lsvCPG.ListItems.Count
    With lsvCPG.ListItems.item(intI)
        If .Text = ccvendor And .SubItems(1) = CCNum Then
            CCExist = True
            Exit Function
        End If
    End With
Next
End Function
Private Function IsFOPCC(ccvendor As String, CCNum As String) As Boolean
If InStr(CCNum, "EXP") > 0 Then
    CCNum = Mid(CCNum, 1, InStr(CCNum, "EXP") - 1)
End If
If ccvendor = gobjPNR.FOP_CCCode And CCNum = gobjPNR.FOP_CCNum Then
    If gobjPNR.FOP_CCExpireDate > Date Then
        IsFOPCC = True
    Else
        IsFOPCC = False
    End If
Else
    IsFOPCC = False
End If
End Function

Private Function GetQueryInfo(ByVal hHttpRequest As Long, i As Integer, ByVal iInfoLevel As Long) As Boolean
    Dim sBuffer         As String * 1024
    Dim lBufferLength   As Long
    lBufferLength = Len(sBuffer)
    GetQueryInfo = CBool(HttpQueryInfo(hHttpRequest, iInfoLevel, ByVal sBuffer, lBufferLength, 0))
    lsvCPG.ListItems(i).SubItems(6) = sBuffer
End Function

Private Function SendTrans(ByRef PC As PaymentClient, cardnum As String, cardexp As String, invno As String, transamt As Long, merchantid As String, imerchantid As String, imerchantpw As String, terminalID As String, ByRef QSICode) As Boolean


Dim strResult As String
Dim strEchoResult As String


strEchoResult = PC.echo("Test")


If strEchoResult <> "echo:Test" Then
        SendTrans = False
        mstrTransError = "Payment Client was not initialised correctly - echo test failed - should be: 'echo:Test', but received: '" & strEchoResult & "'"
        Set PC = Nothing
        Exit Function
End If


PC.addDigitalOrderField "CardNum", cardnum
PC.addDigitalOrderField "CardExp", cardexp
PC.addDigitalOrderField "IMerchantUserID", imerchantid
PC.addDigitalOrderField "IMerchantPassword", imerchantpw
PC.addDigitalOrderField "TerminalID", terminalID
' Create and send the Digital Order
If PC.sendMOTODigitalOrder(invno, merchantid, transamt, "en", "") <> OK Then
       SendTrans = False
       strResult = PC.getResultField("PaymentClient.Error")
       mstrTransError = "Digital Order has not created correctly - sendMOTODigitalOrder(" _
                     & invno & ", " & merchantid & ", " & transamt & ", en,) failed" _
                     & vbCrLf & "Error Msg: " & strResult
        Set PC = Nothing
        Exit Function
End If

' Use the "nextResult" command to check if the DR contains a valid result

If PC.nextResult <> OK Then
        
       ' Retrieve the Payment Client Error (There may be none to retrieve)
        strResult = PC.getResultField("PaymentClient.Error")
        ' Display an Error Page as the Digital Receipt doesn't contain a result
        mstrTransError = "No Results for Digital Receipt" & vbCrLf _
                         & "Error Msg: " & strResult
        SendTrans = False
        Set PC = Nothing
        Exit Function
End If

' Get the Financial Transaction receipt data from the Digital Receipt
' Get the QSI Response Code for the transaction

    QSICode = PC.getResultField("DigitalReceipt.QSIResponseCode")
    If Len(QSICode) = 0 Then
        ' Display an Error Page as the QSIResponseCode could not be retrieved
        mstrTransError = "No result for this field: 'DigitalReceipt.QSIResponseCode'"
        SendTrans = False
        Set PC = Nothing
        Exit Function
    End If



    ' Check if the result contains an error message
    If QSICode <> "null" And QSICode <> "0" Then
        ' Get the error returned from the Payment Client
        strResult = PC.getResultField("DigitalReceipt.ERROR")
        ' check if result contains a value
        If Len(strResult) <> 0 Then
            ' The response is an error message so generate an Error Page
            mstrTransError = "Error returned from Payment Server - QSIResponseCode=" & QSICode _
                             & vbCrLf & "Error Msg: " & strResult
            SendTrans = False
            Set PC = Nothing
            Exit Function
        End If
    End If

    SendTrans = True

End Function

Private Sub doNewCCSale(no As Integer, invno As Long, ccvendor As String, ccno As String, ccexp As String, ccamt As Single)

    Dim strErr As String
    Dim strMsg As String
    Dim intresponse As Integer
    Dim merchantid As String
    Dim imerchantid As String
    Dim imerchantpw As String
    Dim PCHost As String
    Dim PCPort As String
    Dim PCTimeout As String
    Dim bolTrans As Boolean
    Dim lngTranAmt As Long
    Dim strAcqCode As String
    Dim strTransID As String
    Dim objpayclient As PaymentClient
    Dim objpaysocket As PaymentClientSockets
    Dim intFile As Integer
    Dim strReceiptNo As String
    Dim terminalID As String
    Dim strQSICode As String
    intFile = FreeFile()
    Open App.Path & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile

    mstrTransError = ""
    
    Set objpayclient = New PaymentClient
    Set objpaysocket = New PaymentClientSockets
    
    'objpaysocket.HostAddress = "10.180.2.190"
    'objpaysocket.PortNumber = 9050
    'objpaysocket.TimeOut = 200000 '200seconds
    
    merchantid = pGetValue("AcctDetails", "MerchantID")
    imerchantid = pGetValue("AcctDetails", "IMerchantID")
    imerchantpw = pGetValue("AcctDetails", "IMerchantPW")
    PCHost = pGetValue("PaymentClientInfo", "Host")
    PCPort = pGetValue("PaymentClientInfo", "PortNumber")
    PCTimeout = pGetValue("PaymentClientInfo", "Timeout")
    terminalID = pGetValue("PaymentClientInfo", "TerminalID")
    
    Screen.MousePointer = vbHourglass
    
    lngTranAmt = ccamt * 100
    
    bolTrans = SendTrans(objpayclient, ccno, ccexp, CStr(invno), lngTranAmt, merchantid, imerchantid, imerchantpw, terminalID, strQSICode)

    
    If bolTrans = False Then
            
            strMsg = "Transaction Error! PNR: " & gobjPNR.RecLoc & " INV: " & invno & vbCrLf & _
                     mstrTransError
                     
            'MsgBox strMsg
            'lsvCPG.ListItems(no).SubItems(9) = mstrTransError
            Print #intFile, Now & Space(5) & "Error: " & strMsg
            
            
            lsvCPG.ListItems(no).SubItems(8) = mstrTransError
            lsvCPG.ListItems(no).SubItems(4) = "FAIL"
            lsvCPG.ListItems(no).SubItems(10) = strQSICode
            lsvCPG.ListItems(no).SubItems(11) = getAuthDesc(strQSICode, "PaymentServer")

            
     Else
            
            strTransID = objpayclient.getResultField("DigitalReceipt.TransactionNo")
            strAcqCode = objpayclient.getResultField("Rescode")
            strReceiptNo = objpayclient.getResultField("DigitalReceipt.ReceiptNo")
            lsvCPG.ListItems(no).SubItems(5) = strAcqCode
            lsvCPG.ListItems(no).SubItems(4) = IIf(strAcqCode = "00", "PASS", "FAIL")
            lsvCPG.ListItems(no).SubItems(7) = strTransID
            lsvCPG.ListItems(no).SubItems(6) = getAuthDesc(strAcqCode, "IMerchant")
            lsvCPG.ListItems(no).SubItems(9) = strReceiptNo
            lsvCPG.ListItems(no).SubItems(10) = strQSICode
            lsvCPG.ListItems(no).SubItems(11) = getAuthDesc(strQSICode, "PaymentServer")
            Print #intFile, Now & Space(5) & "Response Description: " & strAcqCode & "-" & lsvCPG.ListItems(no).SubItems(5) & vbTab & strTransID & vbTab & strReceiptNo

            
     End If
   
    
     Screen.MousePointer = vbDefault
                        
     Close #intFile
     Set objpayclient = Nothing
     Set objpaysocket = Nothing
End Sub
Private Function DiscountPaidAmt() As Double
    Dim i As Integer
    DiscountPaidAmt = 0
    For i = 1 To lvwPaid.ListItems.Count
        If InStr(1, lvwPaid.ListItems(i).SubItems(1), "CWT FARE DISCOUNT") > 0 Or _
           InStr(1, lvwPaid.ListItems(i).SubItems(1), "CLIENT DISCOUNT") > 0 Then
           DiscountPaidAmt = DiscountPaidAmt + lvwPaid.ListItems(i).SubItems(2)
        End If
    Next
End Function


Private Function IssueTktItinNew(Retries As Integer, doctype As String, Optional Invoice As Boolean = False, Optional AquaInvoice As Boolean = False) As Boolean
Dim strTemp As String
Dim strResp As String
Dim strCT As String
Dim strCI As String
Dim bolTkt As Boolean
Dim bolItin As Boolean
Dim bolETkt As Boolean
Dim strPath As String
Dim intFile As Integer
Dim intStage As Integer
Dim bolSplitCmd As Boolean
Dim strItin As String
Dim itinResponse As String
Dim lngC As Long

'Const CONFIGPRNT As Integer = 1
'Const STATUSPRNT As Integer = 2
'Const PNRFORPRNT As Integer = 3
'Const TICKETPRNT As Integer = 4

Dim TktResponse As String
Dim strTKP As String
Dim strMsg As String

On Error GoTo ErrIssueTktItin

'check whether eticket

bolSplitCmd = SplitCmd

lngC = 0

bolETkt = False

'--------------------TKT Setup----------------
If doctype = "ALL" Or doctype = "TKT" Then

If cmbTKT.listindex > -1 Then

    If cmbTKT.ItemData(cmbTKT.listindex) = 0 Then
        bolETkt = True
    Else
        bolETkt = False
    End If

ElseIf cmbSTP.listindex > -1 Then
    
    If cmbSTP.ItemData(cmbSTP.listindex) = 0 Then
        bolETkt = True
    Else
        bolETkt = False
    End If


End If

'hk: print E-tkt receipt from E-Tkt printer
'If gstrAgcyCountryCode = "HK" Then
'    If bolETkt = False Then
'        If doctype = "ALL" And bolETkt = True And chkItinType.Caption = "P-Itin" Then bolSplitCmd = True
'    End If
'End If



strCT = UCase(cmbTKT & "D  D")
strCI = UCase(cmbITIN & "D  D")
'intStage = CONFIGPRNT


If bolETkt = True Then
   strTemp = "HMLM" & IIf(cmbSTP.Text <> "", cmbSTP.Text, cmbTKT.Text) & "DI"
   strResp = MakeEntry(strTemp)

Else                                            'normal
   strTemp = "HMLM" & IIf(cmbSTP.Text <> "", cmbSTP.Text & "DS", cmbTKT.Text & "DT")
   strResp = MakeEntry(strTemp)
   
End If
'intStage = STATUSPRNT

If InStr(1, strResp, strCT) Then
    strMsg = "Ticket Printer is DOWN!" & vbCrLf & "Is it ready to print?"
    modMsgBox.OKMsg = "OK"
    modMsgBox.CANCELMsg = "Cancel"
    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbOKCancel + vbDefaultButton2, "CWT Desktop - Error") = vbOK Then
    'If MsgBox("Ticket Printer is DOWN!" & vbCrLf & "Is it ready to print?" _
        , vbExclamation + vbDefaultButton1 + vbOKCancel + vbApplicationModal) = vbOK Then
            strTemp = "HMOM" & IIf(cmbSTP.Text <> "", cmbSTP.Text, cmbTKT.Text) & "-U"
            MakeEntry strTemp
    Else
        
        Exit Function
    End If
End If

MakeEntry "HMOM" & IIf(cmbSTP.Text <> "", cmbSTP.Text, cmbTKT.Text) & "-U" 'bring status up

If UCase(gstrAgcyCountryCode) = "SG" Then
   If cmbSTP.Text <> "" Then
      MakeEntry "HMLM" & mstrACPrinter & "DT"
      MakeEntry "HMOM" & mstrACPrinter & "-U"
   End If
End If


End If


'--------------------ITIN Setup----------------

'Paper itin
If ((doctype = "ALL" And bolETkt = False) Or doctype = "ITIN") And chkItinType.Caption = "P-Itin" Then
    If UCase(gstrAgcyCountryCode) = "SG" Then                        'ITIN -not stp
          If cmbSTP.Text = "" Then  'HQ
            strTemp = "HMLM" & cmbITIN.Text & "DI"
            strResp = MakeEntry(strTemp)
          End If
    Else
          strTemp = "HMLM" & cmbITIN.Text & "DI"
          strResp = MakeEntry(strTemp)
    End If



    If InStr(1, strResp, strCI) Then
        strMsg = "Itinerary Printer is DOWN!" & vbCrLf & "Is it ready to print?"
        modMsgBox.OKMsg = "OK"
        modMsgBox.CANCELMsg = "Cancel"
        If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbOKCancel + vbDefaultButton1, "CWT Desktop - Error") = vbOK Then
        'If MsgBox("Itinerary Printer is DOWN!" & vbCrLf & "Is it ready to print?" _
            , vbExclamation + vbDefaultButton1 + vbOKCancel + vbApplicationModal) = vbOK Then
                strTemp = "HMOM" & cmbITIN & "-U"
                MakeEntry strTemp
        Else
            
            Exit Function
        End If
    End If

'If bolETkt = False Then
   MakeEntry "HMOM" & cmbITIN.Text & "-U" 'then bring itin up
   MakeEntry "HMOM" & cmbITIN.Text & "-ITN"
'End If
   

    
    'intStage = PNRFORPRNT
    

End If



    MakeEntry "IMU@"
    strTemp = "IMUDYO" & cmbITINDYO
    MakeEntry strTemp
    If txtCRItin.Text <> "" Then
       strTemp = "IMUCR" & txtCRItin
       MakeEntry strTemp
    End If


'MakeEntry "R.TPRO TKT"


'Modified on 07/09/04: end transaction after DYO changes
If strTemp <> "" Then
    MakeEntry "R.TPRO TKT"
    MakeEntry "ER"
    MakeEntry "ER"
    MakeEntry "ER"
End If
'MakeEntry "*" & gobjPNR.RecLoc

'intStage = TICKETPRNT


'--------------------Issue Commands----------------



'If optAll.Value Then
'   strTKP = "TKPDTD"
'ElseIf optSelection.Value Then
'   strTKP = "TKP" & txtFF & "P" & txtPax & "/DTD"
'End If


        If strFF = "" Then
            strTKP = "TKP"
            strItin = "TKP"
        Else
            strTKP = "TKP" & strFF
            strItin = "TKP" & strFF
        End If
        
        If strPax = "" Then
            'If strFF = "" Then
                strTKP = strItin & "DTD"
                strItin = strItin & "DID"
            'Else
            '    strTKP = strTKP & "P1-" & lstPax.ListCount & "/DTD"
            '    strItin = strTKP & "P1-" & lstPax.ListCount & "/DID"
            'End If
        Else
            'If strFF = "" Then
            '    strTKP = strTKP & "1-" & gobjPNR.FiledFareCount & "P" & strPax & "/DTD"
            '    strItin = strTKP & "1-" & gobjPNR.FiledFareCount & "P" & strPax & "/DID"
            'Else
                strTKP = strItin & "P" & strPax & "/DTD"
                strItin = strItin & "P" & strPax & "/DID"
            'End If
        End If
If bolSplitCmd = True Then
    If doctype = "ALL" Or doctype = "TKT" Then
        If bolETkt Or (doctype = "ALL" And chkItinType.Caption = "P-Itin") Then
            strTKP = strTKP & "ID"
        End If
        If cmbSTP.Text <> "" Then
          TktResponse = MakeEntry(strTKP & "/STP" & txtIATA)
        Else
          TktResponse = MakeEntry(strTKP) 'TKPDTDID
        End If
    End If
Else
    If doctype = "ALL" Or doctype = "TKT" Then
        If bolETkt Or (doctype = "ALL" And chkItinType.Caption = "P-Itin") Then
            strTKP = strTKP & "ID"
        End If
        If cmbSTP.Text <> "" Then
           'If UCase(gstrAgcyCountryCode) = "SG" Then 'sg
           '   TktResponse = MakeEntry(strTKP & "/STP" & txtIATA)
           'Else
              'If optAll.Value Then 'hk
                 
                 TktResponse = MakeEntry(strTKP & "/STP" & txtIATA) 'TKPDTDID/STP13305164
              'Else
              '   TktResponse = MakeEntry(strTKP & IIf(chkItinerary.Value = 1, "ID", "") & "/STP" & txtIATA)
              'End If 'TKP1P1DTDID/STP13305164
           'End If
        Else 'HQ
           'If optAll.Value Then
              TktResponse = MakeEntry(strTKP)  'TKPDTDID
           'Else
           '   TktResponse = MakeEntry(strTKP & IIf(chkItinerary.Value = 1, "ID", "")) 'TKP1P1DTDID
           'End If
        End If
  
    ElseIf doctype = "ITIN" And chkItinType.Caption = "P-Itin" Then
                 TktResponse = MakeEntry(strItin & IIf(cmbSTP.Text <> "", "/STP" & txtIATA, ""))

    End If
    
End If

'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
'Do not queue to EItin Queue for Aqua Itin Client's PNR
If gobjPNR.CompInfo.AquaItin = False Then
    If (doctype = "ITIN" Or doctype = "ALL") And chkItinType.Caption = "E-Itin" Then
        Sleep (3000)
        strResp = MakeEntry("*" & gobjPNR.RecLoc)
        If InStr(1, strResp, "FINISH OR IGNORE") <> 0 Or _
           InStr(1, strResp, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
           Sleep (1000)
           MakeEntry "I"
           Sleep (500)
           strResp = MakeEntry("*" & gobjPNR.RecLoc)
        End If
        If pAddToQueueLog(gobjPNR.RecLoc, "EItin") = False Then
            strMsg = "Cannot send E-itinerary." & vbCrLf & "Cannot add queue key to database."
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
            'MsgBox "Cannot send E-itinerary." & vbCrLf & "Cannot add queue key to database."
            Exit Function
        End If
        MakeEntry "IR"
        MakeEntry "IR"
        
        'Added by Jeremy on 25 Sept 2008 to update EITINQUEUETIME
        For lngC = 1 To gobjPNR.GeneralRemarkCount
            If InStr(gobjPNR.GeneralRemark(lngC).RemarkText, "EITINQUEUETIME:") And gobjPNR.GeneralRemark(lngC).Qualifier = "*I" Then
                gobjHost.terminalEntry "NP." & gobjPNR.GeneralRemark(lngC).ItemNum & "@"
            End If
        Next
        gobjHost.terminalEntry "NP.I*EITINQUEUETIME:" & Format(Now, "ddmmmhh:ss")
        'preethi - V1.1.1 20100915 - IR5 - Add Received From Before Queuing to EItin Queue
        MakeEntry "R.TPRO TKT"
        MakeEntry "ER"
        MakeEntry "ER"
        MakeEntry "ER"
        strResp = MakeEntry("QEB/5E4P/10")
        If Not (InStr(strResp, "ON QUEUE") > 0 Or InStr(strResp, gobjPNR.RecLoc) = 0) Then
            strMsg = "Cannot send E-itinerary. Response from GDS: " & strResp & vbCrLf & "Please resend using the E-Itinerary module."
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        End If
    End If
End If

If TktResponse <> "" Then
If InStr(1, TktResponse, "GENERATED") > 0 Then
    'MsgBox TktResponse
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, TktResponse, vbOKOnly + vbDefaultButton1, "CWT Desktop"
    MakeEntry "*" & gobjPNR.RecLoc
    MakeEntry "NP.SS*VBITKT"
    IssueTktItinNew = True
    If Invoice = True Then
        If chkAqua.value = True Then
            Exit Function
        End If
        MakeEntry "R.TPRO TKT+ER"
        MakeEntry "ER"
        MakeEntry "ER"
    Else
        MakeEntry "R.TPRO TKT+ER"
        MakeEntry "ER"
        MakeEntry "ER"
        Exit Function
    End If
Else
    If Invoice = True Then
        If chkAqua.value = True Then
            'MsgBox TktResponse & vbCrLf & "Invoice will not be queue to Aqua for invoice issuance."
            strMsg = TktResponse & vbCrLf & "Invoice will not be queue to Aqua for invoice issuance."
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop"
        Else
            'MsgBox TktResponse & vbCrLf & "Invoice will not be issued."
            strMsg = TktResponse & vbCrLf & "Invoice will not be issued."
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop"
        End If
    Else
        'MsgBox "Unable to issue! Response from GDS:" & vbCrLf & TktResponse
        strMsg = "Unable to issue! Response from GDS:" & vbCrLf & TktResponse
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End If
        IssueTktItinNew = False
        strPath = App.Path
        strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
        intFile = FreeFile()
        Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
        Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "Exit function"
        Close #intFile
        Exit Function
End If
End If


If bolSplitCmd = True And doctype = "ALL" And chkItinType.Caption = "P-Itin" Then

strResp = MakeEntry("*" & gobjPNR.RecLoc)
If InStr(1, strResp, "FINISH OR IGNORE") <> 0 Or _
   InStr(1, strResp, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
   Sleep (1000)
   MakeEntry "I"
   Sleep (500)
   strResp = MakeEntry("*" & gobjPNR.RecLoc)
End If
Sleep (1000)
MakeEntry "IR"
MakeEntry "IR"

'print bring itinerary printer up

    strTemp = "HMLM" & cmbITIN.Text & "DI"
    strResp = MakeEntry(strTemp)


If InStr(1, strResp, strCI) Then
    strMsg = "Itinerary Printer is DOWN!" & vbCrLf & "Is it ready to print?"
    modMsgBox.OKMsg = "OK"
    modMsgBox.CANCELMsg = "Cancel"
    If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbOKCancel + vbDefaultButton2, "CWT Desktop - Error") = vbOK Then
    'If MsgBox("Itinerary Printer is DOWN!" & vbCrLf & "Is it ready to print?" _
        , vbExclamation + vbDefaultButton1 + vbOKCancel + vbApplicationModal) = vbOK Then
            strTemp = "HMOM" & cmbITIN & "-U"
            MakeEntry strTemp
    Else
        Exit Function
    End If
End If

'If cmbTKT.ItemData(cmbTKT.ListIndex) = 1 Then 'Paper Ticket

MakeEntry "HMOM" & cmbITIN.Text & "-U"
MakeEntry "HMOM" & cmbITIN.Text & "-ITN"

MakeEntry "R.TPRO TKT"
MakeEntry "IMU@"
strTemp = "IMUDYO" & cmbITINDYO
MakeEntry strTemp
If txtCRItin.Text <> "" Then
   strTemp = "IMUCR" & txtCRItin
   MakeEntry strTemp
End If

'Modified on 07/09/04: end transaction after DYO changes
MakeEntry "R.TPRO TKT"
MakeEntry "ER"
MakeEntry "ER"
MakeEntry "ER"

'If optAll.Value Then
'   strItin = "TKPDID"
'ElseIf optSelection.Value Then
'   strItin = "TKP" & txtFF & "P" & txtPax & "/DID"
'End If
itinResponse = MakeEntry(strItin & IIf(cmbSTP.Text <> "", "/STP" & txtIATA, ""))

'If optAll.Value Then
'   If cmbITIN <> "" Then
'     itinResponse = MakeEntry(strItin) 'TKPDTDID
'   End If
'Else
'   If chkItinerary.Value = 1 Then
'     itinResponse = MakeEntry(strItin) 'TKP1P1DTDID
'   End If
'End If


'MsgBox itinResponse
modMsgBox.OKMsg = "OK"
modMsgBox.sMsgBox gVPMDIHwnd, itinResponse, vbOKOnly + vbDefaultButton1, "CWT Desktop"
End If



'MsgBox "Ticket Response = " & TktResponse
'bolTkt = True

'If InStr(1, MakeEntry("TKPDTDID"), " GENERATED ") Then
'    bolTkt = True
'ElseIf InStr(1, MakeEntry("TKPDTDID"), " GENERATED ") Then
'    bolTkt = True
'Else
'    MsgBox "Unable to issue ticket/itinerary!"
'    Exit Sub
'End If

If doctype = "ALL" Then
    strTemp = "Ticket/Itin"
Else
    strTemp = doctype
End If
Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", startTime, SysStart, strTemp, , startTime)




'If optSelection.Value And chkInv.Value = 0 Then
'   bolTktItin = True
'   strResp = MakeEntry("*" & gobjPNR.RecLoc)
'   If InStr(1, strResp, "FINISH OR IGNORE") <> 0 Or _
'      InStr(1, strResp, "RETRIEVAL INHIBITED WHILE IN QUEUE") Then
'      Sleep (1000)
'      MakeEntry "I"
'      Sleep (500)
'      strResp = MakeEntry("*" & gobjPNR.RecLoc)
'   End If
'   Sleep (1000)
   '29122004
'   MakeEntry "IR"
'   MakeEntry "IR"
    
   ' Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", StartTime)
    'Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", startTime, SysStart, "Ticket/Itin", , startTime)
'   Exit Function
'End If

'Call pAddToVBILog(gobjPNR.RecLoc, "Ticket", startTime, SysStart, "Ticket/Itin", , startTime)

Sleep (1000)

'If Invoice = True Then
'    cmdConfirm.Caption = "Waiting for Tkt info"
    strPath = App.Path
    strPath = strPath & IIf(Right(strPath, 1) = "\", "", "\")
    intFile = FreeFile()
    
        Open strPath & "\Log\TKT_" & gobjPNR.RecLoc & "_" & strNow & ".txt" For Append As #intFile
        Print #intFile, gobjPNR.RecLoc & Space(5) & Now & Space(5) & "Completed " & strTemp
        Close #intFile
'    Timer1.Enabled = True
'End If

Exit Function

ErrIssueTktItin:
'Select Case Err.Number
'    Case -2147467259
'        If intStage < 4 Then
'            Retries = Retries + 1
 '           If Retries < 3 Then
 ''               Call IssueTktItin(Retries, Invoice)
 '           Else
  ''              MsgBox CONNECTION_FAIL & " : Please re-run the ticketing process", vbCritical
  '              Exit Function
  '          End If
  '      Else
  '          Call IssueInvoice
  '      End If
  '  Case Else
        'MsgBox "ERROR " & Err.Number & vbCrLf _
            & Err.Description, "RUN TIME ERROR"
        strMsg = "ERROR " & Err.Number & vbCrLf _
            & Err.Description
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        Resume Next
'End Select

End Function
'===========================================================================
' Form Control Events
'
Private Sub pTreeViewSelect()

    Dim oNode    As MSComctlLib.node
    Dim lPtr     As Long
    Dim lSelItem As Variant
    Dim childnode As node
    On Error GoTo ErrorHandler
    lPtr = ObjPtr(moSelNode)
    'Select Case selectchoice
    '    Case [Select]
    If Not moTags.Exist(lPtr) Then
            With moSelNode
                moTags.Add lPtr, .key, .ForeColor, .BackColor, .Bold
                .ForeColor = cSELFORECOLOR
                .BackColor = cSELBACKCOLOR
                '.Bold = True
            End With
            
         For Each childnode In tvStoredFare.Nodes
         If childnode <> moSelNode Then
                If childnode.Root = moSelNode Then
                lPtr = ObjPtr(childnode)
                With childnode
                        moTags.Add lPtr, .key, .ForeColor, .BackColor, .Bold
                        .ForeColor = cSELFORECOLOR
                        .BackColor = cSELBACKCOLOR
                End With
               End If
          End If
         Next

    Else
    '    Case [Clear]
            With moSelNode

                    .ForeColor = moTags.Element(lPtr, [ForeColor])
                    .BackColor = moTags.Element(lPtr, [BackColor])

                    moTags.Remove lPtr

            End With
            
            For Each childnode In tvStoredFare.Nodes
            If childnode <> moSelNode Then
                If childnode.Root = moSelNode Then
                lPtr = ObjPtr(childnode)
                With childnode
                    .ForeColor = moTags.Element(lPtr, [ForeColor])
                    .BackColor = moTags.Element(lPtr, [BackColor])
                     moTags.Remove lPtr
                End With
               End If
           End If
          Next
    End If

     
    
         
    
    
    
    
    
    
    
    'End Select
    tvStoredFare.SetFocus
Exit Sub
ErrorHandler:
    tvStoredFare.Visible = True

End Sub
Private Sub tvStoredFare_NodeClick(ByVal node As MSComctlLib.node)
    If InStr(node.Text, "STORED FARE") Then
        Set moSelNode = node
        pTreeViewSelect
    End If
    node.Selected = False
End Sub
Private Sub getSelected()

Dim intI As Integer
Dim intJ As Integer
Dim lPtr     As Long
Dim oNode    As MSComctlLib.node
strPax = ""
strFF = ""
intJ = 0
For intI = 0 To lstPax.SelCount - 1
    If lstPax.Selected(intI) = True Then
        strPax = strPax & IIf(strPax = "", "", ".") & intI + 1
        intJ = intJ + 1
    End If
Next
If intJ = lstPax.SelCount Then strPax = ""

intJ = 0

For Each oNode In tvStoredFare.Nodes
    If InStr(oNode, "STORED FARE") Then
    lPtr = ObjPtr(oNode)
    
    If moTags.Exist(lPtr) Then
        strFF = strFF & IIf(strFF = "", "", ".") & Trim(Mid(oNode.Text, InStr(oNode.Text, "-") + 2))
        intJ = intJ + 1
    End If
    End If
Next

If intJ = gobjPNR.FiledFareCount Then strFF = ""
End Sub
Private Sub GetAquaQueue(ByRef strAQQueue As String, ByRef strMLQueue As String)
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim strCheck() As String
Dim intI As Integer

'strSQL = "select * from tblMODOptions where (Optioncode='AqTktQueue' and OptionSecCode='" & cmbSTPLoc.Text & "') or Optioncode='ManualTktQueue'"
'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    strSQL = "select * from tblMODOptions where (Optioncode='AqTktQueue' and OptionSecCode='" & cmbSTPLoc.Text & "') or Optioncode='ManualTktQueue_HKSG' and OptionSecCode='" & gobjPNR.CompInfo.AgencyName & "'"
'--

Set rs = gdbConn.Execute(strSQL)

While Not rs.EOF
    If rs!optioncode = "AqTktQueue" Then
        strAQQueue = rs!optionvalue
    Else
        strMLQueue = rs!optionvalue
    End If
    rs.MoveNext
Wend




End Sub

Private Function CheckCPGClient(CN As String) As Boolean
Dim strSQL As String
Dim rs As ADODB.Recordset

 
    strSQL = "select * from tblclients where cn='" & CN & "' and cpgclient='1'"
    Set rs = gdbConn.Execute(strSQL)
    
    If Not rs.EOF Then
        CheckCPGClient = True
    Else
        CheckCPGClient = False
    End If
    
    


End Function
