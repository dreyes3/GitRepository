VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalendar 
   Caption         =   "TAU Date"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin MSComCtl2.MonthView MonthViewTAU 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   63635458
      CurrentDate     =   41334
   End
   Begin VB.Label Label_TAU2 
      Caption         =   " Please select TAU date"
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
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label_TAU1 
      Caption         =   " Ticketing Field is in TAW.  "
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
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dSelectedDate As Date
Dim bPickedDate As Boolean



Private Sub cmdOK_Click()
Dim dDate As Date
   dDate = Me.MonthViewTAU.value
   dSelectedDate = Me.MonthViewTAU.value
   bPickedDate = True
   gbolUpdateTAUDate = True
   Me.Hide
   
End Sub

Private Sub Form_Load()
    ' ZhiSam - V1.2.19 20130311 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    Dim strTemp As String
    Dim strLeft As String
    Dim strRight As String
    Dim dteTemp As Date
    
    
    SwitchWinSetting (Me.hwnd)


    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
    bPickedDate = False
    
    gobjPNR.loadPNR
   
    strTemp = gobjPNR.TAWDate
    strLeft = Left(strTemp, 2)
    strRight = Right(strTemp, 3)
    strTemp = strLeft & "-" & strRight

    dteTemp = CDate(strTemp)
    
    'change the display to next year if TAW date is smaller than tody date
    If dteTemp < Date Then
        dteTemp = DateAdd("yyyy", 1, dteTemp)
    End If
    
    MonthViewTAU.value = dteTemp
    
    
    
End Sub

Private Sub MonthViewTAU_DateClick(ByVal DateClicked As Date)

Dim dDate As Date
   
   dSelectedDate = Me.MonthViewTAU.value
   bPickedDate = True
   CmdOk.SetFocus
   
End Sub

Function dGetSelectedDate() As Date

    dfunctSelectedDate = dSelectedDate

End Function

Function bPickDate() As Boolean
    bfunctPickDate = bPickedDate
End Function
