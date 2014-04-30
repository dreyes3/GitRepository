VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComctl.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Begin VB.Form frmFareQuoteRequest 
   Caption         =   "CWT TravelPro - Fare Quote"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   Icon            =   "frmFareQuoteRequest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   6630
   Begin VB.CheckBox chkLowestFare 
      Caption         =   "Apply lowest fare"
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
      Left            =   120
      TabIndex        =   34
      Top             =   3960
      Value           =   1  'Checked
      Width           =   2415
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3600
      Width           =   615
   End
   Begin VB.ComboBox cmbPCC 
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstActualSegNo 
      Height          =   450
      Left            =   6000
      TabIndex        =   30
      Top             =   6960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   5160
      TabIndex        =   28
      Top             =   3960
      Width           =   735
   End
   Begin VB.Frame frmAction 
      Caption         =   "Select desired action:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   6375
      Begin VB.OptionButton optFQAmend 
         Caption         =   "Fare quote Amendment"
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
         Left            =   2760
         TabIndex        =   27
         Top             =   600
         Width           =   3315
      End
      Begin VB.OptionButton optFQStore 
         Caption         =   "Fare quote && file fare"
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
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   2715
      End
      Begin VB.OptionButton optFileFare 
         Caption         =   "File fare only"
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
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   1995
      End
      Begin VB.OptionButton optFQOnly 
         Caption         =   "Fare quote only"
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
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   2235
      End
   End
   Begin VB.TextBox txtDataModified 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   22
      Top             =   4920
      Width           =   735
   End
   Begin VB.CheckBox chkSkipAdult 
      Caption         =   "Skip Adult"
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
      Left            =   2640
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox cmbPlatCarrier 
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
      Left            =   1560
      TabIndex        =   7
      Text            =   "cmbPlatCarrier"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame fraAirSegs 
      Caption         =   "Select Air Segment(s) if needed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   6375
      Begin MSComctlLib.ListView lsvAirSegs 
         Height          =   2295
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4048
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "O/X"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Segment Details"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Override FBC"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Passenger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   6375
      Begin VB.TextBox txtAge 
         Height          =   350
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
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
         Left            =   4920
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.ComboBox cboPTC 
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
         ItemData        =   "frmFareQuoteRequest.frx":08CA
         Left            =   3120
         List            =   "frmFareQuoteRequest.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cboPassenger 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2835
      End
      Begin MSComctlLib.ListView lsvPx 
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1931
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Num"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Passenger Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblAge 
         Alignment       =   1  'Right Justify
         Caption         =   "Age:"
         Height          =   315
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   315
         Index           =   0
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraFQAirSeg 
      Caption         =   "Fare Quote Air Segment(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   6375
      Begin VB.ListBox lstFQAirSegs 
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
         Height          =   2310
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   20
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame fraFQ 
      Caption         =   "Select Fare Quote:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   6375
      Begin VB.ComboBox cboFQ 
         Height          =   315
         ItemData        =   "frmFareQuoteRequest.frx":08CE
         Left            =   1080
         List            =   "frmFareQuoteRequest.frx":08D0
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   3615
      End
      Begin MSComctlLib.ListView lsvFQPx 
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2143
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Num"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Passenger Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Fare Quote Passenger(s)"
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
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fare Quote:"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin MyCommandButton.MyButton cmdPrevious 
      Height          =   360
      Left            =   1080
      TabIndex        =   35
      Top             =   6960
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
      Left            =   3960
      TabIndex        =   36
      Top             =   6960
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
   Begin MyCommandButton.MyButton cmdDone 
      Height          =   360
      Left            =   2760
      TabIndex        =   37
      Top             =   6960
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
   Begin VB.Label lblPCC 
      Alignment       =   1  'Right Justify
      Caption         =   "PCC:"
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
      Left            =   4800
      TabIndex        =   31
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblExcTax 
      Caption         =   "Exclude Tax:"
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
      Left            =   3960
      TabIndex        =   29
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblPlatCarrier 
      Caption         =   "Platting Carrier:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Air Fare Quotation"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmFareQuoteRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strAirSegNum(18) As String
Dim flgNetFare As Boolean
Dim strCmd As String
Dim bolOverride As Boolean
Dim bolSegSel As Boolean
Dim strSelSegs As String
Dim strSelSegNum As String
Dim strStopOver As String
'The following objects are for multi-line list view entry
'
Private Type LSTVIEWITEM
  mask As Long
  lngItem As Long
  lngSubItem As Long
  state As Long
  stateMask As Long
  pszText As String
  cchTextMax As Long
  lngImage As Long
  lngParam As Long
  lngIndent As Long
End Type

Private dataSetup As Boolean
Private dataModified As Boolean
Private itmClicked As ListItem
Private dwLastSubitemEdited As Long
Private intColumnClickID As Integer
Private Const LVM_FIRST = &H1000
Private Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Private Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON = &H2
Private Const LVHT_ONITEMLABEL = &H4
Private Const LVHT_ONITEMSTATEICON = &H8
Private Const LVHT_ONITEM = (LVHT_ONITEMICON Or _
                           LVHT_ONITEMLABEL Or _
                           LVHT_ONITEMSTATEICON)
Private Const LVIR_LABEL = 2

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type LVHITTESTINFO
  pt As POINTAPI
  flags As Long
  lngItem As Long
  lngSubItem  As Long
End Type

Private Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function SendMessage Lib "user32" _
  Alias "SendMessageA" _
(ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  lngParam As Any) As Long
  
Dim datFormLoadEnd As Date
Dim datFormLoadStart As Date
Dim datTouchEnd As Date

'
' end declaration for multi-line list view entry
'

Private Sub LoadAirSegs()
Dim intC As Integer
Dim intX As Integer
Dim itmx As ListItem
Dim intCurCount As Integer

intC = gobjPNR.AirSegCount

If intC > 0 Then
    
With lsvAirSegs
    .SortKey = 0
    .Sorted = False
    .View = lvwReport
    .GridLines = False
    .FullRowSelect = True
    .LabelEdit = lvwManual
End With
    intX = 1
    For intC = 1 To intC
       'lstAirSegs.AddItem gobjPNR.AirSeg(intC).TextAirSeg
       'added on 18/08/05 to filter flown segments
       If gobjPNR.AirSeg(intC).Flown = False Then
            Set itmx = lsvAirSegs.ListItems.Add(intX, , "")
            itmx.SubItems(1) = gobjPNR.AirSeg(intC).TextAirSeg
         
            lstActualSegNo.AddItem gobjPNR.AirSeg(intC).segnumber
            intCurCount = lstActualSegNo.ListCount
            lstActualSegNo.ItemData(intCurCount - 1) = intC
            
            intX = intX + 1
       End If
    Next
    For intC = 1 To lsvAirSegs.ListItems.Count
    lsvAirSegs.ListItems(intC).Selected = False
    Next
Else
    With lsvAirSegs
    '.FontSize = 14.25
    '.FontBold = True
    .ListItems.Add.SubItems(1) = "NO AIR SEGMENT FOUND!"
    .Enabled = False
    End With
    cmdDone.Enabled = False
End If

End Sub

Private Sub loadPassengers()
Dim intC As Integer
Dim strTemp As String

intC = gobjPNR.PassengerCount
If intC > 0 Then
    For intC = 1 To intC
        With gobjPNR.PassengerName(intC)
            strTemp = Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
        End With
       cboPassenger.AddItem strTemp
    Next
Else
    cboPassenger.AddItem " 1.1  UNKNOWN"
    cboPassenger.Enabled = False
End If

cboPassenger.listindex = 0

lblAge.Visible = False
txtAge.Visible = False

End Sub
'Modified on 21/03/2005

Private Sub cboFQ_Click()
Dim strSQL As String
Dim strSeg As String
Dim rsRec1 As New ADODB.Recordset
Dim rsRec2 As New ADODB.Recordset
Dim intSegID As String
Dim itmx As ListItem
Dim intC As Integer
Dim strSegNo() As String
Dim intSegNum As Integer
Dim lngC As Long

For intC = lsvFQPx.ListItems.Count To 1 Step -1
    lsvFQPx.ListItems.Remove intC
Next intC


chkSkipAdult.value = vbUnchecked
lstFQAirSegs.Clear


    intSegID = cboFQ.ItemData(cboFQ.listindex)
    gFQSegID = intSegID
    strSQL = "Select PxID,PIC,SellAmount,FQPCC from tblFareQuote where [RecLoc] = '" & gobjPNR.RecLoc & "' and SegID=" & intSegID & ""
    'Set rsRec1 = gdbFQ.OpenRecordset(strSQL)
    Set rsRec1 = gdbConn.Execute(strSQL)
    While Not rsRec1.EOF
    Set itmx = lsvFQPx.ListItems.Add(, , rsRec1!PxID)
        For intC = 1 To gobjPNR.PassengerCount
        With gobjPNR.PassengerName(intC)
            If .PassengerNum = rsRec1!PxID Then
               itmx.SubItems(1) = Format(.GDSNum, "@@@@ ") & .LastName & "/" & .FirstName
               Exit For
            End If
        End With
        Next intC
        itmx.SubItems(2) = IIf(rsRec1!PIC <> "", rsRec1!PIC, "AD")
    
    If Not chkSkipAdult.value And rsRec1!SellAmount = 0 Then
        chkSkipAdult.value = vbChecked
    End If
    
    
    If rsRec1!FQPCC <> "" Then
        cmbPCC = rsRec1!FQPCC
    End If
    
    rsRec1.MoveNext
    Wend
    
    lngC = 0
    
    strSQL = "Select distinct(SegSeq),DepDate,COS,Vendor,ArrCityCode,FlightNum,DepCityCode,AdviceAct,SegmentSelectText from tblFareSeg S,tblFareQuote Q "
    strSQL = strSQL & "where Q.RecLoc = '" & gobjPNR.RecLoc & "' and Q.SegID=" & intSegID & " and Q.RecLoc=S.RecLoc and Q.SegID= S.SegID order by SegSeq"
    'Set rsRec2 = gdbFQ.OpenRecordset(strSQL)
    Set rsRec2 = gdbConn.Execute(strSQL)
    While Not rsRec2.EOF
        For lngC = 1 To gobjPNR.AirSegCount

                If gobjPNR.AirSeg(lngC).ArriveAirport = rsRec2!ArrCityCode And gobjPNR.AirSeg(lngC).DepartAirport = rsRec2!DepCityCode And gobjPNR.AirSeg(lngC).Vendor = rsRec2!Vendor And gobjPNR.AirSeg(lngC).FlightNumber = rsRec2!FlightNum And gobjPNR.AirSeg(lngC).DepartDateTime = rsRec2!DepDate And gobjPNR.AirSeg(lngC).Class = rsRec2!Cos Then
                      strSeg = UCase(Format(CStr(gobjPNR.AirSeg(lngC).segnumber), "@@. ") & rsRec2!Vendor & Format(rsRec2!FlightNum, " @@@@") _
                      & rsRec2!Cos & Format(rsRec2!DepDate, " ddmmmyy ") & rsRec2!DepCityCode & rsRec2!ArrCityCode) & " " & rsRec2!AdviceAct
                      lstFQAirSegs.AddItem strSeg
                      Exit For
                End If

        Next lngC
        strSeg = ""

    rsRec2.MoveNext
    Wend
    
    'cmbPlatCarrier.Clear
    'strSql = "Select distinct(Vendor) from tblFareSeg where RecLoc = '" & gobjPNR.RecLoc & "' and SegID=" & intSegID & ""
    
    'Set rsRec2 = gdbConn.Execute(strSql)
    'If Not rsRec2.EOF Then
    '    While Not rsRec2.EOF
    '        If rsRec2!Vendor <> "" Then
    '            cmbPlatCarrier.AddItem rsRec2!Vendor
    '        End If
    '    rsRec2.MoveNext
    '    Wend
    'cmbPlatCarrier.listindex = 0
    'End If
rsRec1.Close
rsRec2.Close
Set rsRec1 = Nothing
Set rsRec2 = Nothing
End Sub



Private Sub cboPTC_Click()
If cboPTC.Text = "CH" Or cboPTC.Text = "VNN" Then
txtAge.Visible = True
lblAge.Visible = True
txtAge.Enabled = True
lblAge.Enabled = True
Else
txtAge.Visible = False
lblAge.Visible = False
txtAge.Enabled = False
lblAge.Enabled = False
End If


End Sub

Private Sub cmbPlatCarrier_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAdd_Click()
   Dim item As ListItem
   Dim strPxType As String
  
   
   'checking
   If validatePxData = False Then Exit Sub
   
   If cboPTC.Text = "AD" Then
      strPxType = "AD"
   ElseIf cboPTC.Text = "CH" Then
      strPxType = "C" & Format(txtAge, "00")
   ElseIf cboPTC.Text = "IN" Then
      strPxType = "INF"  'INFANT
   ElseIf cboPTC.Text = "VAC" Then
      strPxType = "VAC"
      lblExcTax.Visible = True
      cmbCountry.Visible = True
   ElseIf cboPTC.Text = "VNN" Then
      strPxType = "V" & Format(txtAge, "00")
      lblExcTax.Visible = True
      cmbCountry.Visible = True
   End If

      
   Set item = lsvPx.ListItems.Add(, , cboPassenger.listindex + 1)
   item.SubItems(1) = cboPassenger.Text
   item.SubItems(2) = strPxType
End Sub
Private Function validatePxData() As Boolean

 Dim intAge As Integer
 Dim strError As String
 Dim intI As Integer
 
 
 strError = ""
 
 If cboPTC.Text = "CH" Or cboPTC.Text = "VNN" Then
   If txtAge.Text <> "" Then
   intAge = txtAge.Text
  
   End If
   
   
   If intAge > 12 Or intAge < 2 Then
   strError = strError & "Invalid Children Age" & vbCrLf
   End If
End If

For intI = 1 To lsvPx.ListItems.Count
'Debug.Print lsvPx.ListItems(intI).ListSubItems(1)

If cboPassenger.Text = lsvPx.ListItems(intI).ListSubItems(1) Then
strError = strError & "Duplicate Passenger" & vbCrLf
End If
Next intI

If Len(strError) <> 0 Then
        validatePxData = False
        'MsgBox strError, vbApplicationModal + vbExclamation + vbOK, "ERROR!"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strError, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
Else
        validatePxData = True

End If

End Function


Private Sub cmdCancel_Click()
If fWantToQuit Then
    'Unload frmFake
    gbolCancelProcess = True
    Unload Me
    
End If


End Sub
Private Function validatedata(PxNum As String) As Boolean
Dim strError As String
Dim i As Integer
Dim j As Integer
Dim intLB As Integer
Dim intUB As Integer
Dim intTemp As Integer
Dim blnSortFlag As Boolean
Dim strPxNumAry() As String
Dim intEnterCount As Integer
Dim intSelCount As Integer


    validatedata = True
    intEnterCount = 0
    intSelCount = 0
    
    
    
    For i = 1 To lsvAirSegs.ListItems.Count
        If bolSegSel = True Then
            If lsvAirSegs.ListItems(i).Selected Then
                If lsvAirSegs.ListItems(i) <> "" Then
                    intEnterCount = intEnterCount + 1
                End If
                intSelCount = intSelCount + 1
            End If
        Else
                If lsvAirSegs.ListItems(i) <> "" Then
                    intEnterCount = intEnterCount + 1
                End If
                intSelCount = intSelCount + 1
        End If
    Next

    If lsvPx.ListItems.Count = 0 Then
        validatedata = False
        strError = strError & "No Passenger is selected" & vbCrLf
    
    ElseIf lsvPx.ListItems.Count > 1 Then
    'check order
    
    strPxNumAry = Split(PxNum, ";")
    
    intLB = LBound(strPxNumAry)
    intUB = UBound(strPxNumAry)
    
    For i = intLB To intUB - 1
        strPxNumAry(i) = CInt(strPxNumAry(i))
    Next i
    

        For j = intLB To intUB - 2
        
            If strPxNumAry(j) > strPxNumAry(j + 1) Then
               validatedata = False
               strError = strError & "Selected passenger's number must be in ascending sequence" & vbCrLf
                Exit For

            End If
        Next j
        
    ElseIf intEnterCount > 0 And intEnterCount <> intSelCount Then
        strError = strError & "Missing Connection(X)/Stopover(0) indicator" & vbCrLf
        validatedata = False
    End If
    
    If cmbCountry.ListCount > 1 And cmbCountry.Text <> "" Then
        If Len(cmbCountry.Text) <> 2 Then
            strError = strError & "Invalid Country Code" & vbCrLf
            validatedata = False
        End If
    End If
    
    If strError <> "" Then
        'MsgBox strError, vbOKOnly, "Error"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strError, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    End If
    
    
End Function
Private Function checkQuoted() As Boolean
Dim rsRec As ADODB.Recordset
Dim strSQL As String
Dim intI As Integer
Dim intJ As Integer
Dim strError As String

    intJ = 0
    
    For intI = 1 To lsvPx.ListItems.Count
    
    strSQL = "Select * from tblFareQuote where [RecLoc] = '" & gobjPNR.RecLoc & "'"
    strSQL = strSQL & " and [PxID]= " & lsvPx.ListItems(intI).Text & ""
    
    Set rsRec = gdbConn.Execute(strSQL)
    
    If rsRec.EOF Then
        If intJ = 0 Then
            strError = "No FareQuote was found for: " & vbCrLf
            strError = strError & lsvPx.ListItems(intI).SubItems(1) & vbCrLf
            intJ = intJ + 1
        Else
            strError = strError & lsvPx.ListItems(intI).SubItems(1) & vbCrLf
        End If
    End If
    Next
    
    
    
    If strError <> "" Then
        strError = strError & "You have do Fare Quote first before File Fare."
        'MsgBox strError, vbOKOnly, "Fare Quote Checking Error"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strError, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        checkQuoted = False
    Else
        checkQuoted = True
    End If
End Function


Private Sub cmdDone_Click()
Dim strCmd As String
Dim intX As Integer
Dim strCarriers As String
Dim strTemp As String
Dim strPx As String
Dim intPxID() As Integer
Dim lngI As Long
'Fareoverride '12012005
Dim strSegNum As String
Dim strClass As String
'Added on 17/12/2005: Passenger checking
Dim strPxNum As String
'added 17/5/2005: class override checking
Dim intSelSeg As Integer

Dim bolClassExist As Boolean
Dim intActualSegment As Integer
Dim intActualSegOrder As Integer
Dim bolMultipleFares As Boolean
Dim strMsg As String

'Added on 01 Marh 2013 for CR203
Dim strResponse As String
Dim dtLastTravelDate As Date
Dim dtDefaultDate As Date

' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
' Declare boolean variable to capture OBT functRemoveFileFare response
Dim bolRemoveFileFareResponse As Boolean
gGetfareStart = Now
datTouchEnd = Now
dtDefaultDate = "1/1/1900"
 ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
' ZhiSam - V1.3.0 20120203 - Replace P&C with Smart Point V2.1
'If it is not OBT booking, remove Filed Fare that do not have FOP (which was created by Smart Point)
'If it is OBT booking, just ignore it
    If gIntModuleType <> gModuleType.PC Then
        bolRemoveFileFareResponse = functRemoveFileFare()
         ' ZhiSam - V1.2.19 20120311 - CR-203 - Desktop to Create Retention Line and Update TAW to TAU (SyEx with Tpro)
        'check whether Retention Line is present in PNR
         If bfunctCheckRTLine() Then
            'there is retention line exist
            'do nothing
            'strResponse = "Found"
         Else
            'no retention line exist
            dtLastTravelDate = dtFunctLastTravelDate()
            'create retention line based on last travel date plus 90 days
            If dtLastTravelDate > dtDefaultDate Then
                'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
                If gobjPNR.DSInfo.DSPhaseNum >= 1 Then
                    gobjHost.terminalEntry ("RT.T/" & Format(DateAdd("D", 90, dtLastTravelDate), "ddmmm") & "*")
                Else
                    gobjHost.terminalEntry ("RT.T/" & Format(DateAdd("D", 90, dtLastTravelDate), "ddmmm") & "*RETENTION LINE")
                End If
            End If
         End If
   
    End If


Start:
bolMultipleFares = False
gblnAmend = optFQAmend.value

    If optFQAmend.value = True Or optFQOnly.value = True Then
        gbolPerformFF = False
    Else
        gbolPerformFF = True
    End If

If optFQOnly Or optFQStore Then
    For intX = 1 To lsvPx.ListItems.Count
        strPxNum = strPxNum & lsvPx.ListItems(intX).Text & ";"
        ReDim Preserve gstrFQPax(intX - 1)
        gstrFQPax(intX - 1) = lsvPx.ListItems(intX).Text & ";" & lsvPx.ListItems(intX).SubItems(1) & ";" & lsvPx.ListItems(intX).SubItems(2)
    Next
    
    intSelSeg = 0
    For intX = 1 To lsvAirSegs.ListItems.Count
    If lsvAirSegs.ListItems(intX).Selected Then
    intSelSeg = intSelSeg + 1
    End If
    Next
    If intSelSeg > 0 And intSelSeg <> lsvAirSegs.ListItems.Count Then
        bolSegSel = True
    Else
        bolSegSel = False
    End If
    If validatedata(strPxNum) = False Then Exit Sub
End If

frmWait.Show
'29122004
If chkSkipAdult.value = 1 Then
   gbolSkipAdult = True
Else
   gbolSkipAdult = False
End If

If Not gobjFareQuotes Is Nothing Then Set gobjFareQuotes = Nothing
Set gobjFareQuotes = New CWT_Galileo3.FareQuotes

strPx = ""
For intX = 1 To lsvPx.ListItems.Count
    strPx = strPx & lsvPx.ListItems(intX).Text & "*" & lsvPx.ListItems(intX).SubItems(2) & ";"

Next
If Right(strPx, 1) = ";" Then strPx = Left(strPx, Len(strPx) - 1)


If optFQOnly Or optFQStore Then
    
    strSelSegs = ""
    'If lstAirSegs.SelCount > 0 And lstAirSegs.SelCount <> lstAirSegs.ListCount Then
     'If intSelSeg > 0 And intSelSeg <> lsvAirSegs.ListItems.Count Then
    If bolSegSel = True Then
        For intX = 1 To lsvAirSegs.ListItems.Count
        intActualSegment = lstActualSegNo.List(intX - 1)
        intActualSegOrder = lstActualSegNo.ItemData(intX - 1)
            If lsvAirSegs.ListItems(intX).Selected Then
                
                'strSelSegs = strSelSegs & IIf(Len(strSelSegs) <> 0, ".", "") & gobjPNR.AirSeg(intX).SegNumber
                strSelSegs = strSelSegs & IIf(Len(strSelSegs) <> 0, ".", "") & intActualSegment
                gobjPNR.AirSeg(intActualSegOrder).SelectedForPricing = True
                strSelSegNum = strSelSegNum & IIf(Len(strSelSegNum) <> 0, ".", "") & intX
                strTemp = gobjPNR.AirSeg(intX).Vendor
                If InStr(strCarriers, strTemp) = 0 Then strCarriers = strCarriers & IIf(strCarriers <> "", "/", "") & strTemp
            Else
                gobjPNR.AirSeg(intActualSegOrder).SelectedForPricing = False
            End If
        Next intX
    Else
        'bolSegSel = False
        For intX = 1 To lsvAirSegs.ListItems.Count
        intActualSegOrder = lstActualSegNo.ItemData(intX - 1)
            gobjPNR.AirSeg(intActualSegOrder).SelectedForPricing = True
            strSelSegs = strSelSegs & IIf(Len(strSelSegs) <> 0, ".", "") & lstActualSegNo.List(intX - 1)
            strTemp = gobjPNR.AirSeg(intX).Vendor
            If InStr(strCarriers, strTemp) = 0 Then strCarriers = strCarriers & IIf(strCarriers <> "", "/", "") & strTemp
        Next intX
    End If


'Added on 30/06/05: Add stopover option
strStopOver = ""
If bolSegSel = True Then
      For intX = 1 To lsvAirSegs.ListItems.Count
          If lsvAirSegs.ListItems(intX).Selected = True Then
            If lsvAirSegs.ListItems(intX) <> "" Then
              strStopOver = strStopOver & IIf(Len(strStopOver) <> 0, ";", "") & IIf(lsvAirSegs.ListItems(intX) = "O", "N", "Y")
             End If
          End If
      Next
Else
     For intX = 1 To lsvAirSegs.ListItems.Count
          
            If lsvAirSegs.ListItems(intX) <> "" Then
               strStopOver = strStopOver & IIf(Len(strStopOver) <> 0, ";", "") & IIf(lsvAirSegs.ListItems(intX) = "O", "N", "Y")
            End If
       
      Next
End If


End If

'Added on 10/1/05: Split fare quote and file fare option, assign obj back from table
If optFileFare.value = True Or optFQAmend Then

        ReDim intPxID(frmFareQuoteRequest.lsvFQPx.ListItems.Count)
        
        For lngI = 1 To lsvFQPx.ListItems.Count
            intPxID(lngI) = lsvFQPx.ListItems(lngI).Text
            ReDim Preserve gstrFQPax(lngI - 1)
            gstrFQPax(lngI - 1) = lsvFQPx.ListItems(lngI).Text & ";" & lsvFQPx.ListItems(lngI).SubItems(1) & ";" & lsvFQPx.ListItems(lngI).SubItems(2)

        Next lngI
          
        
          
        If Not gobjFareQuotes Is Nothing Then Set gobjFareQuotes = Nothing
        Set gobjFareQuotes = New CWT_Galileo3.FareQuotes
        gobjFareQuotes.getFQFromTbl lsvFQPx.ListItems.Count, intPxID, gobjPNR.RecLoc, cboFQ.ItemData(frmFareQuoteRequest.cboFQ.listindex)
        
        If optFileFare.value Then
            'Write to SQL Log
            WriteToLog
            Load frmPricingWiz1
            Unload Me
            frmPricingWiz1.Show
        Else
            'Write to SQL Log
            WriteToLog
            Load frmFareQuote
            Unload Me
            frmFareQuote.Show
        End If
        Exit Sub

End If


On Error GoTo ProcError

'check if override class is fill up
If bolSegSel = True Then
      For intX = 1 To lsvAirSegs.ListItems.Count
          If lsvAirSegs.ListItems(intX).Selected = True Then
            If lsvAirSegs.ListItems(intX).SubItems(2) <> "" Then
              bolClassExist = True
            Else
             
             bolClassExist = False
              
            End If
          End If
      Next
Else
     For intX = 1 To lsvAirSegs.ListItems.Count
          
            If lsvAirSegs.ListItems(intX).SubItems(2) <> "" Then
              bolClassExist = True
            Else
             
             bolClassExist = False
              
            End If
       
      Next
End If
If bolClassExist = True Then
    bolOverride = True
End If


If bolOverride = True Then GoTo Fareoverride
'If txtClass <> "" Then
   'gobjFareQuotes.GetFareQuote bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC, txtClass, txtStartSeg, txtEndSeg
'Else
   '29122004
   gobjFareQuotes.GetFareQuote strPx, bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC, , , gbolSkipAdult, strStopOver, cmbCountry, cmbPCC, chkPaperTkt.Tag
   'gobjFareQuotes.GetFareQuote strPx, bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC
'End If
gbolOverrideFare = False
'29122004
'Dim aaa As CWT_Galileo3.FareQuote
'Set aaa = New CWT_Galileo3.FareQuote
'Set aaa = gobjFareQuotes(1).FQ(1)


If gobjFareQuotes(1).FQ(1).UnableToQuote And gobjFareQuotes(1).FQ(1).NoFare Then
Fareoverride:
''Set gobjPNR = New CWT_GalileoPNR3.PNR
'If gobjLog.LogOpen Then gobjPNR.OpenLog gobjLog
'gobjPNR.LoadPNR
'   Set gobjFareQuotes = New CWT_Galileo3.FareQuotes
   gbolOverrideFare = True
   Set gobjFareQuotes = New CWT_Galileo3.FareQuotes

   If bolSegSel Then
      '29122004
       bolClassExist = True
       strClass = ""
      For intX = 1 To lsvAirSegs.ListItems.Count
          If lsvAirSegs.ListItems(intX).Selected = True Then
            If lsvAirSegs.ListItems(intX).SubItems(2) = "" Then
              bolClassExist = False
            Else
              'gstrClass(intX) = lsvAirSegs.ListItems(intX).SubItems(2)
              strClass = strClass & IIf(Len(strClass) <> 0, ";", "") & lsvAirSegs.ListItems(intX).SubItems(2)
              
            End If
          End If
      Next
      
      If bolClassExist Then
        gobjFareQuotes.GetFareQuote strPx, bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC, strClass, strSelSegs, gbolSkipAdult, strStopOver, cmbCountry, cmbPCC, chkPaperTkt.Tag
        'gobjFareQuotes.GetFareQuote strPx, bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC, "Y", strSelSegs, gbolSkipAdult
        'gobjFareQuotes.GetFareQuote strPx, bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC, "Y", strSelSegs
        bolOverride = False
      Else
        'MsgBox "Please input the Overriding Price FBC"
        strMsg = "Please input the Overriding Price FBC"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        bolOverride = True
        Unload frmWait
        Me.Show
          For intX = 1 To lsvAirSegs.ListItems.Count
              If lsvAirSegs.ListItems(intX).Selected = True Then
                SetListItemFocus intX - 1, 2
                Exit For
              End If
          Next intX
          
          
        Exit Sub
      End If
   Else
      strTemp = ""
      bolClassExist = True
      strClass = ""
      For intX = 1 To lsvAirSegs.ListItems.Count
            If lsvAirSegs.ListItems(intX).SubItems(2) = "" Then
              bolClassExist = False
            Else
              strClass = strClass & IIf(Len(strClass) <> 0, ";", "") & lsvAirSegs.ListItems(intX).SubItems(2)
              'gstrClass(intX) = lsvAirSegs.ListItems(intX).SubItems(2)
            End If
      Next
      
      For intX = 1 To lsvAirSegs.ListItems.Count
          'FareOverride '12012005
          'strTemp = IIf(strTemp = "", intX, strTemp & "." & intX)
          strSegNum = Trim(Mid(lsvAirSegs.ListItems.item(intX).SubItems(1), 1, InStr(1, lsvAirSegs.ListItems.item(intX).SubItems(1), ".") - 1))
          strTemp = IIf(strTemp = "", strSegNum, strTemp & "." & strSegNum)
      Next
      '29122004
      If bolClassExist Then
      gobjFareQuotes.GetFareQuote strPx, bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC, strClass, strTemp, gbolSkipAdult, strStopOver, cmbCountry, cmbPCC, chkPaperTkt.Tag
      'gobjFareQuotes.GetFareQuote strPx, bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC, "Y", strTemp, gbolSkipAdult
      'gobjFareQuotes.GetFareQuote strPx, bolSegSel, strSelSegs, gobjPNR.CN, strCarriers, cmbPlatCarrier.Text, cboPTC, "Y", strTemp
      bolOverride = False
      Else
        'MsgBox "Please input the Overriding Price FBC"
        strMsg = "Please input the Overriding Price FBC"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        bolOverride = True
        Unload frmWait
        Me.Show
        
        SetListItemFocus 0, 2

        Exit Sub
      End If
   End If
End If

'Call pAddToVBILog(gobjPNR.RecLoc, "Fare Quote", gStartFareQuoteTime, SysStart, "Get Fares")

'If optFQStore.value Or optFQOnly.value Then
'    Do While gobjPNR.CheckPNRStatus < cwtPNREnded
'        strMsg = "You must have completed PNR present to file the fare." & Chr(13) & Chr(13) _
                            & "You may toggle to Focalpoint and finish and end the current PNR, then click 'Retry' below to continue." & Chr(13) _
                            & "If you would like to continue with a fare quote only, without filing the fare, click on 'Cancel' below to continue"
'        modMsgBox.RETRYMsg = "Retry"
'        modMsgBox.CANCELMsg = "Cancel"
 '       If modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbRetryCancel + vbDefaultButton1, "CWT Desktop - Error") = vbCancel Then
 '                      optFQOnly.value = True
 '
 '               gobjHost.TerminalEntry ("R.TPRO FAREQUOTE+ER")
 '               gobjHost.TerminalEntry ("ER")
 '               gobjHost.TerminalEntry ("ER")
 '               Set gobjPNR = New CWT_GalileoPNR3.PNR
 '               gobjPNR.LoadPNR
 '               GoTo Start
 '               Exit Do
 '       End If
        
 '   Loop
'End If


''If gobjFareQuotes(1).QuoteType = "UNABLE TO FARE QUOTE" Then
'If gobjFareQuotes(1).UnableToQuote Then
'   Call pRedisplayMenu
'   Exit Sub
'End If
For intX = 1 To lsvPx.ListItems.Count
    gobjFareQuotes(intX).FQ(1).StoreFare = optFQStore.value
Next





'Added on 7/3/2005: Check again if able to override farequote with Y
If gobjFareQuotes(1).FQ(1).UnableToQuote Then
    'Write to SQL Log
    WriteToLog
    Unload Me
    Exit Sub
Else
    If chkLowestFare.value = 0 Then
       'Added on 27/3/2008. List all the fares if chkLowestfare is unchecked
        intX = 0
        With gobjFareQuotes(1)
             For lngI = 1 To .FQCount
                 If .FQ(lngI).BaseAmount > 0 Then
                    intX = intX + 1
                 End If
             Next
             If intX > 1 Then bolMultipleFares = True
        End With
        If bolMultipleFares = True Then
           gbolSelectFare = True
           Load frmFareDisplay
           frmFareDisplay.Show
           Do
             DoEvents
           Loop Until isLoaded("frmFareDisplay") = False
           If gbolSelectFare = False Then
              Unload frmWait
              frmFareQuoteRequest.SetFocus
              Exit Sub
           End If
           
        End If
    End If
    'Write to SQL Log
    WriteToLog
    Load frmFareQuote
    Unload frmWait
    Unload Me

    frmFareQuote.Show
    Exit Sub
End If

ProcError:
    Unload frmWait
    If Err.Description = "UNABLE TO FARE QUOTE - FAILED RULES VALIDATION" Then
        strMsg = "UNABLE TO FARE QUOTE - FAILED RULES VALIDATION" & Chr(13) & Chr(13) _
            & "Please book a priceable itinerary and try again!"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        'MsgBox "UNABLE TO FARE QUOTE - FAILED RULES VALIDATION" & Chr(13) & Chr(13) _
            & "Please book a priceable itinerary and try again!"
            Call pErrorReport(False, False)
            'Write to SQL Log
            WriteToLog
            Unload Me
    Else
        'Write to SQL Log
        WriteToLog
        Call pErrorReport
    End If
    
End Sub

Private Sub chkPaperTkt_Click()
    With chkPaperTkt
        If .value = vbChecked Then
            .Caption = "PT"
            .Tag = "P"
        Else
            .Caption = "ET"
            .Tag = "E"
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload Me
End Sub

Private Sub Form_Load()

'Timer
Dim oldParent As Long
   
datFormLoadStart = Now
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
Me.Move 0, 0
Me.Move frmSideBar.Width, 0
pClearWindow

'Added on 11/1/2005: allow multiple FQ
fraFQ.Visible = False
fraFQAirSeg.Visible = False

Call LoadAirSegs
Call loadPassengers
'If gobjPNR.PassengerCount = 1 Then
'cboPTC.listindex = 0
'cmdAdd_Click
'End If
'modified on 12/11/2005
Call populatePlatingCarrier

Screen.MousePointer = vbDefault

'Added on 26/7/2005: Add Passenger Type
cboPTC.AddItem "AD"
cboPTC.AddItem "CH"
cboPTC.AddItem "IN"
cboPTC.AddItem "VAC"
cboPTC.AddItem "VNN"
cboPTC.listindex = 0

If gobjPNR.PassengerCount = 1 Then
cboPTC.listindex = 0
cmdAdd_Click
End If

cmbCountry.AddItem ""
cmbCountry.AddItem "US"
cmbCountry.listindex = 0

lblExcTax.Visible = False
cmbCountry.Visible = False
txtDataModified.Visible = False

'added on 240806: Allow FQ in other PCC if PCC is setup in tblclients
If gobjPNR.CompInfo.AltFQPCC = "" Then
    lblPCC.Visible = False
    cmbPCC.Visible = False
    cmbPCC.AddItem ""
Else
    lblPCC.Visible = True
    cmbPCC.Visible = True
    cmbPCC.AddItem gobjPNR.CompInfo.AltFQPCC
    cmbPCC.AddItem gstrHQPCC
    cmbPCC.listindex = 0
End If
chkPaperTkt_Click
lsvAirSegs.ToolTipText = "If none are selected, all segments will be priced. To Select Multiple Segments, Left Click the Segment & Hold Ctrl Key."

    If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
    Else
      cmdPrevious.Visible = False
    End If

datFormLoadEnd = Now
If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    If Not fWantToQuit Then
        Cancel = 1
    Else
        Me.Show
       
        'Call pRedisplayMenu
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFareQuoteRequest = Nothing
End Sub



Private Sub lsvPx_DblClick()
   lsvPx.ListItems.Remove (lsvPx.SelectedItem.Index)
End Sub


Private Sub optFileFare_Click()
    fraAirSegs.Visible = False
    Frame1.Visible = False
    fraFQ.Visible = True
    fraFQAirSeg.Visible = True
    populateFQControls
    lblPlatCarrier.Visible = False
    cmbPlatCarrier.Visible = False
    cmbPCC.Locked = True
    chkPaperTkt.Visible = False
    chkLowestFare.Visible = False
End Sub
'Added on 11/1/2005: allow multiple FQ
'modified on 21/03/2005
Private Sub populateFQControls()
Dim strSQL As String
Dim rsRec As New ADODB.Recordset
Dim strSeg As String
Dim strPx As String
Dim intNextPx As Integer
Dim intPrevPx As Integer
Dim intC As Integer
Dim intD As Integer
Dim blnNextFlag As Boolean
Dim strPrevSegID As String
Dim strPrevSegSeq As String
Dim strPrevPx As String
Dim blnNextPx As Boolean

'reset variable
strPx = ""
strSeg = ""
blnNextPx = False
intC = 1
intNextPx = 0
intPrevPx = 0
cboFQ.Clear

    strSQL = "SELECT DISTINCT Q.RecLoc, Q.SegID, Q.PxID, S.SegSeq,* " & _
             "FROM tblFareQuote Q INNER JOIN tblFareSeg S " & _
             "ON Q.SegID = S.SegID AND Q.RecLoc = S.RecLoc AND Q.PxID = S.PxID " & _
             "AND Q.RecLoc = '" & gobjPNR.RecLoc & "' order by Q.SegID,Q.PxID"
                     
    Set rsRec = gdbConn.Execute(strSQL)
    If rsRec.EOF Then
       cmdDone.Enabled = False
       cboFQ.AddItem "--No Fare Quote Found--"
       cboFQ.Enabled = False
    Else
       
        While Not rsRec.EOF
        If rsRec!SegID <> strPrevSegID Then
            If strSeg <> "" And strPx <> "" Then
                    cboFQ.AddItem strPrevSegID & "." & "   " & strPx & "   S" & strSeg
                    cboFQ.ItemData(cboFQ.NewIndex) = strPrevSegID
                    strPx = ""
                    strSeg = ""
                    blnNextPx = False
                    intC = 1
                    intNextPx = 0
                    intPrevPx = 0
            End If
        End If
        
        If rsRec!SegID = strPrevSegID Then

            If rsRec!PxID <> strPrevPx Then

            If rsRec!PxID <> intNextPx And blnNextFlag = True Then

                strPx = strPx & "," & rsRec!PxID
                intNextPx = rsRec!PxID + 1
                intPrevPx = rsRec!PxID
                blnNextFlag = False
                
            ElseIf rsRec!PxID <> intNextPx And blnNextFlag = False Then
                strPx = strPx & "," & rsRec!PxID
                intNextPx = rsRec!PxID + 1
                intPrevPx = rsRec!PxID
                blnNextFlag = False
            Else
                intNextPx = rsRec!PxID + 1
                intPrevPx = rsRec!PxID
                blnNextFlag = True
                
               
                If intC = 1 Then
                    strPx = strPx & "-" & rsRec!PxID
                    intC = intC + 1
                Else
                    strPx = Left(strPx, InStrRev(strPx, "-")) & rsRec!PxID
                End If
           
            End If
            
            blnNextPx = True
            
            End If
            
        Else
                strPx = "P" & rsRec!PxID
                intNextPx = rsRec!PxID + 1
                
        End If
         
        
    If blnNextPx = False Then
        If rsRec!SegID = strPrevSegID Then
            If strPrevPx = rsRec!PxID Then
                 If rsRec!SegSeq <> strPrevSegSeq Then
                    For intD = 1 To gobjPNR.AirSegCount
                            If gobjPNR.AirSeg(intD).ArriveAirport = rsRec!ArrCityCode And gobjPNR.AirSeg(intD).DepartAirport = rsRec!DepCityCode And gobjPNR.AirSeg(intD).Vendor = rsRec!Vendor And gobjPNR.AirSeg(intD).FlightNumber = rsRec!FlightNum And gobjPNR.AirSeg(intD).DepartDateTime = rsRec!DepDate Then
                                strSeg = strSeg & IIf(strSeg = "", gobjPNR.AirSeg(intD).segnumber, "." & gobjPNR.AirSeg(intD).segnumber)
                                Exit For
                            End If
            
                    Next intD
                 End If
            End If
        ElseIf strSeg = "" Then
            For intD = 1 To gobjPNR.AirSegCount
                            If gobjPNR.AirSeg(intD).ArriveAirport = rsRec!ArrCityCode And gobjPNR.AirSeg(intD).DepartAirport = rsRec!DepCityCode And gobjPNR.AirSeg(intD).Vendor = rsRec!Vendor And gobjPNR.AirSeg(intD).FlightNumber = rsRec!FlightNum And gobjPNR.AirSeg(intD).DepartDateTime = rsRec!DepDate Then
                                strSeg = strSeg & IIf(strSeg = "", gobjPNR.AirSeg(intD).segnumber, "." & gobjPNR.AirSeg(intD).segnumber)
                                Exit For
                            End If
            
                    Next intD
        End If
    End If

        strPrevSegID = rsRec!SegID
        strPrevSegSeq = rsRec!SegSeq
        strPrevPx = rsRec!PxID
        rsRec.MoveNext
        Wend
        
        If strSeg <> "" And strPx <> "" Then
                    cboFQ.AddItem strPrevSegID & "." & "   " & strPx & "   S" & strSeg
                    cboFQ.ItemData(cboFQ.NewIndex) = strPrevSegID
        End If
        
    End If
    
    If cboFQ.ListCount > 0 Then
        cboFQ.listindex = 0
    Else
       cmdDone.Enabled = False
       cboFQ.AddItem "--No Fare Quote Found--"
       cboFQ.Enabled = False
       cboFQ.listindex = 0
    End If
    rsRec.Close
    Set rsRec = Nothing
End Sub

Private Sub optFQAmend_Click()
    fraAirSegs.Visible = False
    Frame1.Visible = False
    fraFQ.Visible = True
    fraFQAirSeg.Visible = True
    lblPlatCarrier.Visible = False
    cmbPlatCarrier.Visible = False
    populateFQControls
    cmbPCC.Locked = False
    chkPaperTkt.Visible = False
    chkLowestFare.Visible = False
End Sub

Private Sub optFQOnly_Click()
    fraFQ.Visible = False
    fraFQAirSeg.Visible = False
    fraAirSegs.Visible = True
    Frame1.Visible = True
    If lsvAirSegs.Enabled = True Then
        cmdDone.Enabled = True
    End If
    chkSkipAdult.value = vbUnchecked
    populatePlatingCarrier
    lblPlatCarrier.Visible = True
    cmbPlatCarrier.Visible = True
    cmbPCC.Locked = False
    chkPaperTkt.Visible = True
    chkLowestFare.Visible = True
End Sub

Private Sub optFQStore_Click()
    fraFQ.Visible = False
    fraFQAirSeg.Visible = False
    fraAirSegs.Visible = True
    Frame1.Visible = True
    If lsvAirSegs.Enabled = True Then
        cmdDone.Enabled = True
    End If
    chkSkipAdult.value = vbUnchecked
    populatePlatingCarrier
    lblPlatCarrier.Visible = True
    cmbPlatCarrier.Visible = True
    cmbPCC.Locked = False
    chkPaperTkt.Visible = True
    chkLowestFare.Visible = True
End Sub
Private Sub populatePlatingCarrier()

Dim intC As Integer
Dim strCarrierList As String

cmbPlatCarrier.Clear
cmbPlatCarrier.AddItem ""

If gobjPNR.AirSegCount > 0 Then
    For intC = 1 To gobjPNR.AirSegCount
    With gobjPNR.AirSeg(intC)
        If InStr(1, strCarrierList, .Vendor) = 0 Then
            strCarrierList = strCarrierList & IIf(strCarrierList = "", "", "/") & .Vendor
            cmbPlatCarrier.AddItem .Vendor
        End If
    End With
    Next
    
    If Len(gobjPNR.AirSeg(1).Vendor) > 0 Then
        cmbPlatCarrier.Text = gobjPNR.AirSeg(1).Vendor
    End If
End If


End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
 KeyAscii = fAllowNumeric(KeyAscii)
End Sub

Private Sub txtClass_KeyPress(KeyAscii As Integer)
 KeyAscii = fAllowAlpha(KeyAscii)
End Sub
Private Sub lsvAirSegs_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

'this routine:
'1. sets the last change if the dataModified flag is set
'2. sets a flag to prevent setting the dataModified flag
'3. determines the item or subitem clicked
'4. calc's the position for the text box
'5. moves and shows the text box
'6. clears the dataModified flag
'7. clears the DoingSetup flag

  Dim hti As LVHITTESTINFO
  Dim fpx As Single
  Dim fpy As Single
  Dim fpw As Single
  Dim fph As Single
  Dim rc As RECT
  Dim topindex As Long
  Dim selSeg() As String
  Dim i As Integer
'prevent the textbox change event from
'registering as dataModified when the text is
'assigned to the textbox
  dataSetup = True

'if a pending dataModified flag is set, update the
'last edited item before moving on
  If dataModified And dwLastSubitemEdited > -1 Then
     If dwLastSubitemEdited = 0 Then
     itmClicked = txtDataModified.Text
     Else
     itmClicked.SubItems(dwLastSubitemEdited) = txtDataModified.Text
     End If
  End If

'hide the textbox
  txtDataModified.Visible = False

'get the position of the click
  With hti
     .pt.X = (X / Screen.TwipsPerPixelX)
     .pt.Y = (Y / Screen.TwipsPerPixelY)
     .flags = LVHT_ONITEM
  End With

'find out which subitem was clicked
  Call SendMessage(lsvAirSegs.hwnd, _
                   LVM_SUBITEMHITTEST, _
                   0, hti)

'if on an item (HTI.lngItem <> -1) and
'the click occurred on the subitem
'column of interest (HTI.lngSubItem = 2 -
'which is column 3 (0-based)) move and
'show the textbox
intColumnClickID = hti.lngSubItem

If hti.lngSubItem = 2 Or hti.lngSubItem = 0 Then
'select row

   If hti.lngItem <> -1 Then 'And hti.lngSubItem > 0 Then

    'prevent the listview label editing
    'from occurring if the control has
    'full row select set
     lsvAirSegs.LabelEdit = lvwManual

    'determine the bounding rectangle
    'of the subitem column
     rc.Left = LVIR_LABEL
     rc.Top = hti.lngSubItem
     Call SendMessage(lsvAirSegs.hwnd, _
                      LVM_GETSUBITEMRECT, _
                      hti.lngItem, _
                      rc)

    'we need to keep track of which
    'item was clicked so the item can
    'be updated later
    'position the text box
     Set itmClicked = lsvAirSegs.ListItems(hti.lngItem + 1)
     itmClicked.Selected = True

    'get the current top index
     topindex = SendMessage(lsvAirSegs.hwnd, _
                            LVM_GETTOPINDEX, _
                            0&, _
                            ByVal 0&)

    'establish the bounding rect for
    'the subitem in VB terms (the x
    'and y coordinates, and the height
    'and width of the item
     fpx = lsvAirSegs.Left + _
             (rc.Left * Screen.TwipsPerPixelX) + 180

     fpy = lsvAirSegs.Top + _
             (hti.lngItem + 1 - topindex) + _
             (rc.Top * Screen.TwipsPerPixelY) + 4250

    'a hard-coded height for the text box
     fph = 30

    'get the column width for the subitem
     fpw = SendMessage(lsvAirSegs.hwnd, _
                       LVM_GETCOLUMNWIDTH, _
                       hti.lngSubItem, _
                       ByVal 0&)

    'calc the required width of
    'the textbox to fit in the column
     fpw = (fpw * Screen.TwipsPerPixelX) - 20

    'assign the current subitem
    'value to the textbox
     With txtDataModified
        If hti.lngSubItem = 0 Then
        .Text = itmClicked
        '.Text = lsvAirSegs.ListItems(hti.lngItem + 1)
        Else
        .Text = itmClicked.SubItems(hti.lngSubItem)
        End If
        dwLastSubitemEdited = hti.lngSubItem

       'position it over the subitem, make
       'visible and assure the text box
       'appears overtop the listview
        .Move fpx, fpy, fpw, fph
        .Visible = True
        .ZOrder 0
        .SetFocus

     End With

    'clear the setup flag to allow the
    'textbox change event to set the
    '"dataModified" flag, and clear that flag
    'in preparation for editing
     dataSetup = False
     dataModified = False

  End If

End If
End Sub

Private Sub lsvAirSegs_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
Dim i As Integer
Dim a As Integer
Dim hti As LVHITTESTINFO
Dim selSeg() As String

'if showing the text box, set
'focus to it and select any
'text in the control
  If txtDataModified.Visible = True Then

     With txtDataModified
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
     End With

  End If
  
  If bolOverride = True Then
  
 If bolSegSel = False Then
 
For i = 1 To lsvAirSegs.ListItems.Count

    lsvAirSegs.ListItems(i).Selected = False
Next
Else
selSeg = Split(strSelSegNum, ".")
For i = 0 To UBound(selSeg)
    'For a = 1 To lsvAirSegs.ListItems.Count
     '   If CInt(selSeg(i)) = a Then
       
            lsvAirSegs.ListItems(CInt(selSeg(i))).Selected = True
            'Exit For
        'Else
         '   If lsvAirSegs.ListItems(a).Selected <> True Then
          '      lsvAirSegs.ListItems(a).Selected = False
           ' End If
            
            
            'Exit For
  '      End If
        
    Next
'Next
'get the position of the click
  With hti
     .pt.X = (X / Screen.TwipsPerPixelX)
     .pt.Y = (Y / Screen.TwipsPerPixelY)
     .flags = LVHT_ONITEM
  End With
 Call SendMessage(lsvAirSegs.hwnd, _
                   LVM_SUBITEMHITTEST, _
                   0, hti)
                   
If hti.lngItem <> -1 Then

For i = 0 To UBound(selSeg)
    If selSeg(i) = (hti.lngItem + 1) Then
        lsvAirSegs.ListItems(hti.lngItem + 1).Selected = True
        Exit For
    Else
        lsvAirSegs.ListItems(hti.lngItem + 1).Selected = False
    End If
Next
End If

 'If lsvAirSegs.ListItems(hti.lngItem + 1).Selected <> True Then
  '  lsvAirSegs.ListItems(hti.lngItem + 1).Selected = False
 'End If
 




'For i = 0 To UBound(selSeg)
'    For a = 1 To lsvAirSegs.ListItems.Count
'        If (CInt(selSeg(i)) <> a) And (lsvAirSegs.ListItems(a).Selected = True) Then
'        If lsvAirSegs.ListItems(a).Selected <> True Then
'            lsvAirSegs.ListItems(a).Selected = False
 '       End If
            'Exit For
        'Else
         '   If lsvAirSegs.ListItems(a).Selected <> True Then
          '      lsvAirSegs.ListItems(a).Selected = False
           ' End If
            
            
            'Exit For
'        End If
        
'    Next
'Next



End If
End If
End Sub


Private Sub txtDataModified_Change()
If Not dataSetup Then
   dataModified = True
   If intColumnClickID = 0 Then
        txtDataModified.MaxLength = 1
   Else
        txtDataModified.MaxLength = 10
   End If
End If
End Sub
Private Sub txtDataModified_KeyPress(KeyAscii As Integer)
    With lsvAirSegs
    If .ListItems.Count > 0 Then
    If intColumnClickID = 0 Then
            If IsNumeric(.SelectedItem.Index) Then
                If .SelectedItem.Index > 0 Then
                    KeyAscii = fAllowConx(KeyAscii)
                End If
        End If
    ElseIf intColumnClickID = 2 Then
        If IsNumeric(.SelectedItem.Index) Then
            If .SelectedItem.Index > 0 Then
                KeyAscii = fAllowAlphaNumeric(KeyAscii)
            End If
        End If
    End If
    End If
    End With
End Sub

Private Sub txtDataModified_LostFocus()

If dataModified And dwLastSubitemEdited > 0 Then
   itmClicked.SubItems(dwLastSubitemEdited) = txtDataModified.Text
   dataModified = False
End If


End Sub

Private Sub SetListItemFocus(ByVal intItem As Integer, intSubItem As Integer)
'this routine:
'1. sets the last change if the dataModified flag is set
'2. sets a flag to prevent setting the dataModified flag
'3. determines the item or subitem clicked
'4. calc's the position for the text box
'5. moves and shows the text box
'6. clears the dataModified flag
'7. clears the DoingSetup flag

  Dim hti As LVHITTESTINFO
  Dim fpx As Single
  Dim fpy As Single
  Dim fpw As Single
  Dim fph As Single
  Dim rc As RECT
  Dim topindex As Long
Dim X As Single
Dim Y As Single

'prevent the textbox change event from
'registering as dataModified when the text is
'assigned to the textbox
  dataSetup = True

'if a pending dataModified flag is set, update the
'last edited item before moving on
  If dataModified And dwLastSubitemEdited > 0 Then
     itmClicked.SubItems(dwLastSubitemEdited) = txtDataModified.Text
  End If

'hide the textbox
  txtDataModified.Visible = False

'get the position of the click
  With hti
     .pt.X = (X / Screen.TwipsPerPixelX)
     .pt.Y = (Y / Screen.TwipsPerPixelY)
     .flags = LVHT_ONITEM
  End With

'find out which subitem was clicked
'Call SendMessage(lsvAirSegs.hwnd, _
                   LVM_SUBITEMHITTEST, _
                   0, hti)

'if on an item (HTI.lngItem <> -1) and
'the click occurred on the subitem
'column of interest (HTI.lngSubItem = 2 -
'which is column 3 (0-based)) move and
'show the textbox
intColumnClickID = intSubItem
hti.lngSubItem = intSubItem
hti.lngItem = intItem
If hti.lngSubItem = 2 Then

   If hti.lngItem <> -1 And hti.lngSubItem > 0 Then
    
     Set itmClicked = lsvAirSegs.ListItems(hti.lngItem + 1)
    
     itmClicked.Selected = True
     itmClicked.EnsureVisible
    
    'prevent the listview label editing
    'from occurring if the control has
    'full row select set
     lsvAirSegs.LabelEdit = lvwManual

    'determine the bounding rectangle
    'of the subitem column
     rc.Left = LVIR_LABEL
     rc.Top = hti.lngSubItem
     
     Call SendMessage(lsvAirSegs.hwnd, _
                      LVM_GETSUBITEMRECT, _
                      hti.lngItem, _
                      rc)

    'we need to keep track of which
    'item was clicked so the item can
    'be updated later
    'position the text box
     'Set itmClicked = lsvAirSegs.ListItems(hti.lngItem + 1)
    
     'itmClicked.Selected = True
     'itmClicked.EnsureVisible
    'get the current top index
     topindex = SendMessage(lsvAirSegs.hwnd, _
                            LVM_GETTOPINDEX, _
                            0&, _
                            ByVal 0&)

    'establish the bounding rect for
    'the subitem in VB terms (the x
    'and y coordinates, and the height
    'and width of the item
     fpx = lsvAirSegs.Left + _
             (rc.Left * Screen.TwipsPerPixelX) + 180

     'fpy = lsvAirSegs.Top + _
     '        (hti.lngItem + 1 - topindex) + _
     '        (rc.Top * Screen.TwipsPerPixelY) + 4250


         fpy = (rc.Top * Screen.TwipsPerPixelY) + 4250 _
                + (hti.lngItem + 1 - topindex) + lsvAirSegs.Top
             
    'a hard-coded height for the text box
      fph = 30

    'get the column width for the subitem
     fpw = SendMessage(lsvAirSegs.hwnd, _
                       LVM_GETCOLUMNWIDTH, _
                       hti.lngSubItem, _
                       ByVal 0&)

    'calc the required width of
    'the textbox to fit in the column
     fpw = (fpw * Screen.TwipsPerPixelX) - 20

    'assign the current subitem
    'value to the textbox
     With txtDataModified

        .Text = itmClicked.SubItems(hti.lngSubItem)

        dwLastSubitemEdited = hti.lngSubItem

       'position it over the subitem, make
       'visible and assure the text box
       'appears overtop the listview
        .Move fpx, fpy, fpw, fph
        .Visible = True
        .ZOrder 0
        .SetFocus

     End With

    'clear the setup flag to allow the
    'textbox change event to set the
    '"dataModified" flag, and clear that flag
    'in preparation for editing
     dataSetup = False
     dataModified = False

  End If
End If
End Sub


Private Function fAllowConx(ByRef AsciiCode As Integer) As Integer
Dim lngC As Long

    Select Case AsciiCode
           Case 8, 79, 88, 111, 120 'allow only space, O & X
               fAllowConx = Asc(UCase(Chr(AsciiCode))) ' change valid characters to uppercase
           
    End Select

End Function

Function WriteToLog()

    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareQuoteRequest, _
    Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareQuoteRequest, _
    Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
    
    pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
    gconModAir, frmSideBar.cmbSelectType.Text, gconSModFareQuoteRequest, _
    Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd

End Function

