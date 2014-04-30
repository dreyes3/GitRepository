VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFareDisplay 
   Caption         =   "CWT TravelPro - Fare Display"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   8355
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Picture         =   "frmFareDisplay.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Picture         =   "frmFareDisplay.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
   Begin MSComctlLib.ListView lswFares 
      Height          =   2595
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Quote"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Currency"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Fare"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Fare Category"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Account Code"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Cat 35"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Fare Type"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Tour Code"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Please select one of the fares listed below"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmFareDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    gbolSelectFare = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim j As Integer
    Dim bolSelect As Boolean
    Dim bolStoreFare As Boolean
    Dim strMsg As String
    
    bolSelect = False
    With lswFares
          For i = 1 To .ListItems.Count
              If .ListItems(i).Selected = True Then
                  bolSelect = True
                  j = .ListItems(i).Text
                  Exit For
              End If
          Next
    End With
    If bolSelect = False Then
       'MsgBox "Please select a fare ..."
       strMsg = "Please select a fare ..."
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Sub
    End If
    
    If j > 1 Then
        For i = 1 To gobjFareQuotes.PxCount
            bolStoreFare = gobjFareQuotes(i).FQ(1).StoreFare
            gobjFareQuotes(i).FQAdd gobjFareQuotes(i).FQ(j), 1
            gobjFareQuotes(i).FQ(1).StoreFare = bolStoreFare
        Next
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim item As ListItem
    Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
    
    With gobjFareQuotes(1)
        For i = 1 To .FQCount
            With .FQ(i)
                If .BaseAmount > 0 Then
                   Set item = lswFares.ListItems.Add(, , i)
                   item.SubItems(1) = .BaseCurrency
                   item.SubItems(2) = .BaseAmount
                   If .PrivateFare = True And .PFAccountCode <> "" Then
                      item.SubItems(3) = "CORPORATE"
                   ElseIf .PrivateFare = True And .PFAccountCode = "" Then
                      item.SubItems(3) = "MARKET"
                   ElseIf .PrivateFare = False Then
                      item.SubItems(3) = "PUBLISHED"
                   End If
                   item.SubItems(4) = .PFAccountCode
                   item.SubItems(5) = IIf(.Cat35 = True, "YES", "NO")
                   item.SubItems(6) = .PFFareType
                   item.SubItems(7) = .ITNum
                End If
            End With
        Next
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then gbolSelectFare = False
End Sub

Private Sub lswFares_DblClick()
    cmdOK_Click
End Sub
