VERSION 5.00
Begin VB.Form frmWait 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2775
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmWait.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3450
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   2610
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3180
      Begin VB.PictureBox Picture2 
         Height          =   525
         Left            =   2520
         Picture         =   "frmWait.frx":0442
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   4
         Top             =   240
         Width           =   525
      End
      Begin VB.PictureBox Picture1 
         Height          =   675
         Left            =   240
         Picture         =   "frmWait.frx":0D0C
         ScaleHeight     =   615
         ScaleWidth      =   2475
         TabIndex        =   3
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Please wait."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   360
         Left            =   900
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Working...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   420
         Left            =   180
         TabIndex        =   2
         Top             =   420
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
'pSetFormPosition Me.hWnd, vbTopMost

End Sub

Private Sub Form_Load()
Dim oldParent As Long
    
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)
    
Me.Move 0, 0
Me.Move frmSideBar.Width, 0
End Sub
