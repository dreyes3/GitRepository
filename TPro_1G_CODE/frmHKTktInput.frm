VERSION 5.00
Begin VB.Form frmHKTktInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticket/Invoice Input Form"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Issuing For"
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
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   4215
      Begin VB.TextBox txtNewNF 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optClient 
         Caption         =   "Others"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optClient 
         Caption         =   "Philips"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optClient 
         Caption         =   "NRCC"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblNewNF 
         Caption         =   "New Nett Fare:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fare Type"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton optFT 
         Caption         =   "Special"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optFT 
         Caption         =   "Normal"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmHKTktInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
If validEntry Then
    Call procInvForClients
End If
Unload Me
End Sub
Private Function validEntry()
    Dim strMsg As String
    
    validEntry = True
    If Not optFT(0).value Then
        If optClient(1).value Then
            If txtNewNF.Text = "" Then
                txtNewNF.Text = "0"
            ElseIf Not IsNumeric(txtNewNF.Text) Then
                'MsgBox "This is not a valid amount!"
                strMsg = "This is not a valid amount!"
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
                validEntry = False
            End If
        End If
    End If
    
End Function

Private Sub Form_Load()
   Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
End Sub

Private Sub optClient_Click(Index As Integer)
If optClient(1).value = True Then
    txtNewNF.Enabled = True
Else
    txtNewNF.Enabled = False
End If
End Sub

Private Sub optFT_Click(Index As Integer)
If optFT(0).value Then
    Frame2.Enabled = False
    clearFrame2
Else
    Frame2.Enabled = True
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub clearFrame2()
    optClient(0).value = False
    optClient(1).value = False
    optClient(2).value = False
    txtNewNF.Text = ""
End Sub

Private Sub procInvForClients()
If Not optFT(0).value Then
    Call removeEB
    If optClient(0).value Then
        Call adjustForNRCC
    ElseIf optClient(1).value Then
        Call adjustForPhilips
    End If
End If
End Sub

Private Sub removeEB()
Dim intF As Integer
Dim strEntry As String

For intF = 1 To gobjPNR.FiledFareCount
    With gobjPNR.FiledFare(intF)
        strEntry = "TMU" & intF & "EB@"
        gobjHost.terminalEntry strEntry
    End With
Next
gobjHost.terminalEntry "R.TKT"
gobjHost.terminalEntry "ER"
gobjHost.terminalEntry "ER"
gobjHost.terminalEntry "ER"
End Sub

Private Sub adjustForNRCC()
Dim intF As Integer
Dim strEntry As String

For intF = 1 To gobjPNR.FiledFareCount
    With gobjPNR.FiledFare(intF)
        strEntry = "TMU" & intF & "ASF@"
        gobjHost.terminalEntry strEntry
        strEntry = "TMU" & intF & "F@MSCARD"
        gobjHost.terminalEntry strEntry
    End With
Next
gobjHost.terminalEntry "R.TKT"
gobjHost.terminalEntry "ER"
gobjHost.terminalEntry "ER"
gobjHost.terminalEntry "ER"
End Sub

Private Sub adjustForPhilips()
Dim intF As Integer
Dim strEntry As String

If IsNumeric(txtNewNF.Text) Then
    For intF = 1 To gobjPNR.FiledFareCount
        With gobjPNR.FiledFare(intF)
            strEntry = "TMU" & intF & "NF@" & gstrAgcyCurrCode & Format(txtNewNF.Text, gstrAgcyCurrFormat)
            gobjHost.terminalEntry strEntry
        End With
    Next
End If
gobjHost.terminalEntry "R.TKT"
gobjHost.terminalEntry "ER"
gobjHost.terminalEntry "ER"
gobjHost.terminalEntry "ER"
End Sub

