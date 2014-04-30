VERSION 5.00
Begin VB.Form frmEndorseList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CWT TravelPro - Endorsements"
   ClientHeight    =   2910
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEndorse 
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Preloaded Endorsements in FC:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmEndorseList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    Dim i As Integer
    Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
    
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub
