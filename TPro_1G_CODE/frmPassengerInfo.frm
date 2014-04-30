VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPassengerInfo 
   BackColor       =   &H00F7DFD6&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Traveller Information"
   ClientHeight    =   3555
   ClientLeft      =   4050
   ClientTop       =   315
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraFOP 
      BackColor       =   &H00F7DFD6&
      Caption         =   "Form of Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4200
      TabIndex        =   57
      Top             =   360
      Width           =   6000
      Begin VB.TextBox txtPerCCNum 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   2640
         MaxLength       =   16
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cboPerCCCode 
         BackColor       =   &H00C0FFFF&
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
         ItemData        =   "frmPassengerInfo.frx":0000
         Left            =   1680
         List            =   "frmPassengerInfo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtCCNum 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   2640
         MaxLength       =   16
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboCCCode 
         BackColor       =   &H00C0FFFF&
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
         ItemData        =   "frmPassengerInfo.frx":0004
         Left            =   1680
         List            =   "frmPassengerInfo.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpCCExpire 
         Height          =   315
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   1120
         _ExtentX        =   1984
         _ExtentY        =   556
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
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   12648447
         CustomFormat    =   "M/yyyy"
         Format          =   64028675
         CurrentDate     =   37987
         MaxDate         =   73050
         MinDate         =   37987
      End
      Begin MSComCtl2.DTPicker dtpPerCCExpire 
         Height          =   315
         Left            =   4800
         TabIndex        =   8
         Top             =   720
         Width           =   1120
         _ExtentX        =   1984
         _ExtentY        =   556
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
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   12648447
         CustomFormat    =   "M/yyyy"
         Format          =   64028675
         CurrentDate     =   37987
         MaxDate         =   73050
         MinDate         =   37987
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Credit Card:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expires:"
         Height          =   195
         Left            =   4800
         TabIndex        =   59
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Air Credit Card:"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   420
         Width           =   1050
      End
   End
   Begin VB.Frame fraBillAddress 
      BackColor       =   &H00F7DFD6&
      Caption         =   "Billing Address (If Required)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   -11215
      TabIndex        =   56
      Top             =   -3835
      Width           =   5655
      Begin VB.TextBox txtBillAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   27
         TabIndex        =   20
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox txtBillAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   1
         Left            =   120
         MaxLength       =   27
         TabIndex        =   21
         Top             =   600
         Width           =   5295
      End
      Begin VB.TextBox txtBillAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   2
         Left            =   120
         MaxLength       =   27
         TabIndex        =   22
         Top             =   960
         Width           =   5295
      End
      Begin VB.TextBox txtBillCity 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         MaxLength       =   27
         TabIndex        =   23
         Top             =   1320
         Width           =   1680
      End
      Begin VB.TextBox txtBillState 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1920
         MaxLength       =   27
         TabIndex        =   24
         Top             =   1320
         Width           =   1680
      End
      Begin VB.TextBox txtBillCountry 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3720
         MaxLength       =   27
         TabIndex        =   25
         Top             =   1320
         Width           =   1680
      End
      Begin VB.TextBox txtBillPostalCode 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   120
         MaxLength       =   27
         TabIndex        =   26
         Top             =   1680
         Width           =   1680
      End
   End
   Begin VB.Frame fraDelivAddress 
      BackColor       =   &H00F7DFD6&
      Caption         =   "Delivery Address (If Different From Company)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   -5455
      TabIndex        =   54
      Top             =   -3835
      Width           =   4215
      Begin VB.TextBox txtDelivAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   5
         Left            =   120
         MaxLength       =   27
         TabIndex        =   19
         Top             =   2040
         Width           =   3900
      End
      Begin VB.TextBox txtDelivAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   4
         Left            =   120
         MaxLength       =   27
         TabIndex        =   18
         Top             =   1680
         Width           =   3900
      End
      Begin VB.TextBox txtDelivAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   3
         Left            =   120
         MaxLength       =   27
         TabIndex        =   17
         Top             =   1320
         Width           =   3900
      End
      Begin VB.TextBox txtDelivAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   2
         Left            =   120
         MaxLength       =   27
         TabIndex        =   16
         Top             =   960
         Width           =   3900
      End
      Begin VB.TextBox txtDelivAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   1
         Left            =   120
         MaxLength       =   27
         TabIndex        =   15
         Top             =   600
         Width           =   3900
      End
      Begin VB.TextBox txtDelivAdd 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   0
         Left            =   1020
         MaxLength       =   27
         TabIndex        =   14
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attendant:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame fraPassportInfo 
      BackColor       =   &H00F7DFD6&
      Caption         =   "Passport Information"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   -8815
      TabIndex        =   46
      Top             =   -3295
      Width           =   7575
      Begin VB.ComboBox cmbNationality 
         BackColor       =   &H00C0FFFF&
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
         ItemData        =   "frmPassengerInfo.frx":0008
         Left            =   1680
         List            =   "frmPassengerInfo.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtNumber 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   1680
         MaxLength       =   16
         TabIndex        =   28
         Top             =   1000
         Width           =   2055
      End
      Begin VB.ComboBox cmbResidence 
         BackColor       =   &H00C0FFFF&
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
         ItemData        =   "frmPassengerInfo.frx":000C
         Left            =   1680
         List            =   "frmPassengerInfo.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1440
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpExpDate 
         Height          =   315
         Left            =   5640
         TabIndex        =   33
         Top             =   1485
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   12648447
         Format          =   64028673
         CurrentDate     =   37987
         MaxDate         =   109574
         MinDate         =   36892
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   315
         Left            =   5640
         TabIndex        =   32
         Top             =   1005
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   12648447
         Format          =   64028673
         CurrentDate     =   25569
         MaxDate         =   109574
      End
      Begin VB.CheckBox chkNoInfo 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Tick here if you do not have passport information"
         Height          =   255
         Left            =   120
         MaskColor       =   &H8000000F&
         TabIndex        =   27
         Top             =   260
         Width           =   4215
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00F7DFD6&
         Height          =   435
         Left            =   4800
         TabIndex        =   47
         Top             =   480
         Width           =   2535
         Begin VB.OptionButton optFemale 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Female"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   31
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optMale 
            BackColor       =   &H00F7DFD6&
            Caption         =   "Male"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Natinality:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passport Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   1100
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration:"
         Height          =   195
         Left            =   4800
         TabIndex        =   51
         Top             =   1575
         Width           =   735
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
         Height          =   195
         Left            =   4560
         TabIndex        =   50
         Top             =   1095
         Width           =   930
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Residence:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F7DFD6&
      Caption         =   "Traveler Name (From Passport)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Left            =   120
      TabIndex        =   43
      Top             =   360
      Width           =   3975
      Begin VB.TextBox txtSurname 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surname:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   780
         Width           =   795
      End
   End
   Begin VB.Frame fraContacts 
      BackColor       =   &H00F7DFD6&
      Caption         =   " Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      TabIndex        =   37
      Top             =   1560
      Width           =   10095
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   0
         Left            =   960
         MaxLength       =   20
         TabIndex        =   9
         Tag             =   "OFFICE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   1
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "OFFICE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   2
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   10
         Tag             =   "OFFICE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtPhone 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Index           =   3
         Left            =   960
         MaxLength       =   20
         TabIndex        =   12
         Tag             =   "OFFICE"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtEmail 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   3600
         TabIndex        =   13
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Business:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Home:"
         Height          =   195
         Left            =   5760
         TabIndex        =   41
         Top             =   420
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
         Height          =   195
         Left            =   3000
         TabIndex        =   40
         Top             =   420
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   780
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   195
         Left            =   3000
         TabIndex        =   38
         Top             =   780
         Width           =   420
      End
   End
   Begin MyCommandButton.MyButton cmdFinish 
      Height          =   360
      Left            =   7680
      TabIndex        =   35
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
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
      Picture         =   "frmPassengerInfo.frx":0010
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   16244694
      Caption         =   "Fnished"
      Depth           =   1
      PictureDisabled =   "frmPassengerInfo.frx":035D
      GradientType    =   2
   End
   Begin MyCommandButton.MyButton cmdCancel 
      Height          =   360
      Left            =   9120
      TabIndex        =   36
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
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
      Picture         =   "frmPassengerInfo.frx":0723
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   16244694
      Caption         =   "Cancel"
      Depth           =   1
      PictureDisabled =   "frmPassengerInfo.frx":0A4B
      GradientType    =   2
   End
   Begin MyCommandButton.MyButton cmdBack 
      Height          =   360
      Left            =   6240
      TabIndex        =   34
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
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
      Picture         =   "frmPassengerInfo.frx":0D9D
      AppearanceThemes=   1
      BackColorDown   =   3968251
      BackColorOver   =   6805503
      BackColorFocus  =   16765357
      BackColorDisabled=   16765357
      TransparentColor=   16244694
      Caption         =   "Back "
      Depth           =   1
      PictureDisabled =   "frmPassengerInfo.frx":10BC
      GradientType    =   2
   End
End
Attribute VB_Name = "frmPassengerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim datFormStart As Date

Private Sub chkNoInfo_Click()
   If chkNoInfo.Value = vbChecked Then
      fraPassportInfo.Enabled = False
   Else
      fraPassportInfo.Enabled = True
   End If
End Sub

Private Sub cmdBack_Click()
    gbolCancelMove = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If fWantToQuit = True Then
        moveBar
        Unload Me
    End If
End Sub

Private Sub cmdFinish_Click()
    Dim SysStart As Date
    SysStart = Now
    If validData Then
        If moveBar = True Then
           writeToGDS
           Unload Me
        Else
           Unload Me
        End If
    End If
End Sub

Private Function validData() As Boolean
    
    Dim strMsg As String

    If Len(txtSurname.Text) < 2 Then strMsg = strMsg & "Invalid or missing surname..." & Chr(13)
    If Len(txtFirstName.Text) < 1 Then strMsg = strMsg & "Missing first name..." & Chr(13)
    If txtPhone(1).Text = "" And txtPhone(2).Text = "" Then strMsg = strMsg & "Home or mobile phone contact required..." & Chr(13)
    If fraBillAddress.Enabled = True And txtBillAdd(0).Text = "" _
        And txtBillCity.Text = "" And txtBillPostalCode.Text = "" Then strMsg = strMsg & "Missiong or incomplete billing address..." & Chr(13)
    
    If cboCCCode.Visible = True Then
        If cboCCCode.Text <> "" Then
            
            If txtCCNum.Text = "" Then
               strMsg = strMsg & "Missing CC Number..." & Chr(13)
            ElseIf ValidCCNum(cboCCCode.Text, txtCCNum.Text) = False Then
               strMsg = strMsg & "Invalid or incomplete CC number..." & Chr(13)
            End If
            
            If dtpCCExpire.Value < Now Then strMsg = strMsg & "Invalid CC expire date or CC has expired" & Chr(13)
            
        ElseIf txtCCNum.Text <> "" Then
            If cboCCCode.Text = "" Then strMsg = strMsg & "Missing CC Code..." & Chr(13)
            If dtpCCExpire.Value < Now Then strMsg = strMsg & "Invalid CC expire date or CC has expired" & Chr(13)
        End If
    End If

    If txtPerCCNum <> "" Then
        If cboPerCCCode = "" Then strMsg = strMsg & "Missing personal CC Code..." & Chr(13)
        If dtpPerCCExpire.Value < Now Then strMsg = strMsg & "Invalid personal CC epire date or CC has expired" & Chr(13)
        If ValidCCNum(cboPerCCCode.Text, txtPerCCNum.Text) = False Then strMsg = strMsg & "Invalid or imcomplete personal CC number..." & Chr(13)
    End If

    If chkNoInfo.Value = vbUnchecked Then
        If cmbNationality.Text = "" Then strMsg = strMsg & "Missing or invalid passport nationality..." & Chr(13)
        If Len(txtNumber.Text) > 0 Then
            If dtpExpDate.Value < Now Then
                strMsg = strMsg & "Invalid passport expiration date..." & Chr(13)
            End If
        End If
    End If
    
    If strMsg = "" Then
        validData = True
    Else
        validData = False
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox VPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "Traveller Information"
    End If
    
End Function

Private Sub writeToGDS()
    Dim strCmd As String
    Dim lngC As Long
    Dim strMsg As String
    Dim strTemp As String
      
    strCmd = "N." & txtSurname & "/" & txtFirstName
    
    strTemp = "BHCF"  'for the phone qualifier (see below)
    For lngC = 0 To 3
        If txtPhone(lngC).Text <> "" Then
            strCmd = strCmd & "+P." & gstrAgcyCityCode & Mid(strTemp, lngC + 1, 1) & "*" & txtPhone(lngC) & "-" & txtPhone(lngC).Tag
        End If
    Next
    
    'Added on 3/1/05: Generate SI Information for Phone Field
    If txtPhone(1) <> "" Then
        strCmd = strCmd & "+SI.YY*CTCH " & gstrAgcyCityCode & " " & txtPhone(1)
    ElseIf txtPhone(2) <> "" Then
        strCmd = strCmd & "+SI.YY*CTCP " & gstrAgcyCityCode & " " & txtPhone(2) & "-MOBILE"
    End If
    
    If txtEmail.Text <> "" Then
        strTemp = txtEmail.Text
        
        Do While InStr(strTemp, "@") > 0
            strTemp = Left(strTemp, InStr(strTemp, "@") - 1) & "//" & Mid(strTemp, InStr(strTemp, "@") + 1)
        Loop
    
        Do While InStr(strTemp, "_") > 0
            strTemp = Left(strTemp, InStr(strTemp, "_") - 1) & "--" & Mid(strTemp, InStr(strTemp, "_") + 1)
        Loop
        
        strCmd = strCmd & "+P." & gstrAgcyCityCode & "E*" & strTemp
        strTemp = txtEmail.Text
        strTemp = convertText(strTemp)
        strCmd = strCmd & "+RI.ITI." & strTemp
        strCmd = strCmd & "+RI.INV." & strTemp
        strCmd = strCmd & "+RI.TKT." & strTemp
    End If
    
    If txtDelivAdd(1).Text <> "" Then
        If Len(gobjPNR.DeliveryAddress) > 0 Then gobjHost.terminalEntry "D.@"
        For lngC = 1 To 4
            strCmd = strCmd & IIf(txtDelivAdd(lngC).Text <> "", IIf(lngC = 1, "+D." & txtDelivAdd(lngC).Text, "*" & txtDelivAdd(lngC).Text), "")
        Next
    End If
    
    If fraBillAddress.Enabled = True Then
        If Len(gobjPNR.BillingAddress) > 0 Then gobjHost.terminalEntry "W.@"
        strCmd = strCmd & "+W." & txtBillAdd(0).Text & IIf(txtBillAdd(1).Text = "", "", "*" & txtBillAdd(1).Text) _
                 & IIf(txtBillAdd(2).Text = "", "", "*" & txtBillAdd(2).Text) & "*" & txtBillCity _
                 & IIf(txtBillState.Text = "", "", " ." & txtBillState.Text) & " P/" & txtBillPostalCode _
                 & IIf(txtBillCountry.Text = "", "", "*" & txtBillCountry.Text)
    End If
    
    'Accept INVAGT if no CC is provided
    If cboCCCode.Visible = True Then
        If gobjPNR.FOPType <> "" Then gobjHost.terminalEntry "F.@"
        If cboCCCode.Text <> "" And cboCCCode.Text <> "" Then
            strCmd = strCmd & "+F." & cboCCCode.Text & txtCCNum & "/D" & Format(dtpCCExpire, "mmyy")
        Else
            strCmd = strCmd & "+F.INVAGT"
        End If
    End If
    
    If txtPerCCNum.Text <> "" Then
        strCmd = strCmd & "+NP.F*PERS CC: " & cboPerCCCode.Text & txtPerCCNum & "/D" & Format(dtpPerCCExpire, "mmyy")
    End If
    
    strMsg = gobjHost.terminalEntry(strCmd)
    If Trim(strMsg) <> "*" Then
        strMsg = "Data not entered into GDS!" & Chr(13) & Chr(13) _
                 & "GDS RESPONSE WAS: " & strMsg
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox VPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "Traveller Information"
    Else
        If chkNoInfo = vbUnchecked Then
    
           strCmd = "NP.G*BDAY: " & Format(dtpDOB, "ddmmmyy") _
                    & IIf(Len(txtNumber.Text) > 0, "+NP.P*PASSPORT NO: " & txtNumber & "-" & GetCountryCode(cmbNationality.Text) & "-" & "EXP " & Format(dtpExpDate, "ddmmmyy"), "") _
                    & IIf(Me.cmbResidence.Text = "", "", "+NP.P*RESIDENCE: " & GetCountryCode(cmbResidence)) _
                    & "+NP.P*GENDER: " & IIf(optFemale.Value = True, "F", "M") _
         
           strMsg = gobjHost.terminalEntry(strCmd)
           If Trim(strMsg) <> "*" Then
                strMsg = "Data not entered into GDS!" & Chr(13) & Chr(13) _
                         & "GDS RESPONSE WAS: " & strMsg
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox VPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "Traveller Information"
            End If
        End If
    End If
End Sub

Private Sub Form_Load()

   Dim oldParent As Long
    
   oldParent = SetParent(Me.hWnd, VPMDIHwnd)
   GetPerCCCode
   GetFOP
   chkNoInfo_Click
   optMale.Value = True
   ssTravellerTab.TabIndex = 0
   dtpExpDate.Value = Date
   dtpDOB.Value = Date
   dtpCCExpire.Value = Date
   fraPassportInfo.Enabled = True
   txtDelivAdd(0).Locked = True
   GetCountry
   GetCN
   
End Sub

Private Sub txtBillAdd_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 32, 35, 45, 47, 48 To 57, 65 To 90, 97 To 122
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select
End Sub

Private Sub txtBillCity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 48 To 57, 65 To 90, 97 To 122
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select
End Sub

Private Sub txtBillCountry_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 48 To 57, 65 To 90, 97 To 122
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select
End Sub

Private Sub txtBillPostalCode_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtBillState_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtCCNum_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii)
End Sub

Private Sub txtDelivAdd_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
           Case 8, 32, 35, 45, 47, 48 To 57, 65 To 90, 97 To 122
               KeyAscii = Asc(UCase(Chr(KeyAscii))) ' change valid characters to uppercase
           Case Else
               KeyAscii = 0
    End Select
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlphaNumeric(KeyAscii, "@_-.")
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii, " ")
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
   KeyAscii = fAllowAlphaNumeric(KeyAscii)
End Sub

Private Sub txtPhone_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = fAllowNumeric(KeyAscii, "- ")
End Sub

Private Function GetCountryCode(CountryName As String) As String
    Dim rsCountry As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select CountryCode from tblCountryCodes " & _
             "Where [CountryName] = '" & CountryName & "'"
    
    Set rsCountry = gdbConn.Execute(strSql)
    
    If rsCountry.EOF = False Then
       GetCountryCode = rsCountry![CountryCode]
    Else
       GetCountryCode = ""
    End If
    
    rsCountry.Close
    Set rsCountry = Nothing
        
End Function

Private Sub GetCountry()
   Dim rsCountry As ADODB.Recordset
   Dim strSql As String
   
   cmbNationality.Clear
   cmbResidence.Clear
   
   strSql = "Select CountryName from tblCountryCodes " & _
            "order by CountryName"

   Set rsCountry = gdbConn.Execute(strSql)
   
   Do Until rsCountry.EOF
      cmbNationality.AddItem rsCountry!CountryName
      cmbResidence.AddItem rsCountry!CountryName
      rsCountry.MoveNext
   Loop
   
    rsCountry.Close
    Set rsCountry = Nothing
End Sub


Private Sub GetPerCCCode()
   cboPerCCCode.Clear
   cboPerCCCode.AddItem ""
   cboPerCCCode.AddItem "AX"
   cboPerCCCode.AddItem "CA"
   cboPerCCCode.AddItem "VI"
   cboPerCCCode.AddItem "DC"
   cboPerCCCode.AddItem "JP"
   cboPerCCCode.AddItem "TP"
   cboPerCCCode.ListIndex = 0
End Sub

Private Sub GetFOP()
   cboCCCode.Clear
   cboCCCode.AddItem ""
   cboCCCode.AddItem "AX"
   cboCCCode.AddItem "CA"
   cboCCCode.AddItem "VI"
   cboCCCode.AddItem "DC"
   cboCCCode.AddItem "JP"
   cboCCCode.AddItem "TP"
   cboCCCode.ListIndex = 0
End Sub

Private Sub GetCN()

    Dim rsClients As ADODB.Recordset
    Set rsClients = gdbConn.Execute("select * from tblClients where ProName = '" & frmSideBar.cmbBar.Text & "'")
    
    If Not rsClients.EOF Then
        cboCCCode.Visible = rsClients![FOPInTravPro]
        txtCCNum.Visible = rsClients![FOPInTravPro]
        dtpCCExpire.Visible = rsClients![FOPInTravPro]
        txtBillAdd(0).Visible = rsClients![BillAddInTravPro]
        txtBillAdd(1).Visible = rsClients![BillAddInTravPro]
        txtBillAdd(2).Visible = rsClients![BillAddInTravPro]
        txtBillCity.Visible = rsClients![BillAddInTravPro]
        txtBillState.Visible = rsClients![BillAddInTravPro]
        txtBillPostalCode.Visible = rsClients![BillAddInTravPro]
        txtBillCountry.Visible = rsClients![BillAddInTravPro]
        fraBillAddress.Enabled = rsClients![BillAddInTravPro]
        fraBillAddress.Visible = rsClients![BillAddInTravPro]
    End If
End Sub

Private Sub txtSurname_KeyPress(KeyAscii As Integer)
    KeyAscii = fAllowAlpha(KeyAscii)
End Sub


Public Function ValidCCNum(Vendor As String, CCNum As String) As Boolean

    Dim strCompare As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intZ As Integer
    Dim intCD As Integer
    
    CCNum = Trim(CCNum)
    
    Select Case Vendor
        Case "AX"
            If Len(CCNum) <> 15 Or (Left(CCNum, 2) <> "34" And Left(CCNum, 2) <> "37") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "TP"
           If Len(CCNum) <> 15 Or (Left(CCNum, 4) <> "1920" And Left(CCNum, 4) <> "1220") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "VI", "BA"
            If (Len(CCNum) <> 16 And Len(CCNum) <> 13) _
            Or (Left(CCNum, 1) <> "4") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "MC", "CA", "IB"
            If (Len(CCNum) <> 16) _
            Or (Left(CCNum, 2) <> "51" And Left(CCNum, 2) <> "52" And Left(CCNum, 2) <> "53" And Left(CCNum, 2) <> "54" And Left(CCNum, 2) <> "55") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "DS"
            If (Len(CCNum) <> 16) _
            Or (Left(CCNum, 4) <> "6011") Then
                ValidCCNum = False
                Exit Function
            End If
        Case "DC"
            If (Len(CCNum) <> 14) _
            Or (Left(CCNum, 2) <> "30" And Left(CCNum, 2) <> "36" And Left(CCNum, 2) <> "38") Then
                ValidCCNum = False
                Exit Function
            End If
        Case Else
            Err.Raise -1004, "CompanyProfile.ValidCCNum", "Unknown Credit Card Vendor"
    End Select
    strCompare = Format(CCNum, "00000000000000000000")
    
    For intX = 20 To 2 Step -2
    intY = CInt(Mid(strCompare, intX - 1, 1)) * 2
    intZ = CInt(Mid(strCompare, intX, 1))
    
    intCD = intCD + (intZ + IIf(intY < 10, intY, 1 + (intY - 10)))
    Next
    If (intCD / 10) - Int(intCD / 10) = 0 Then
       ValidCCNum = True
    Else
        ValidCCNum = False
    End If

End Function

Private Function moveBar() As Boolean

    If gobjHost Is Nothing Then gobjHost = New CWT_Galileo.GalileoHost

    'Back function in Adhoc traveler screen
    If gobjHost.moveProfile(gstrPCC, UCase(frmSideBar.cmbBar.Text)) <> "PRO" Then
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox VPMDIHwnd, "No Company Profile Found", vbOKOnly + vbDefaultButton1, "Traveller Information"
       gbolCancelMove = True
    Else
       'If Not gobjPNR Is Nothing Then Set gobjPNR = Nothing
       'Set gobjPNR = New CWT_GalileoPNR.PNR
       If gobjLog.LogOpen Then gobjPNR.OpenLog gobjLog
       gobjPNR.LoadPNR
       moveBar = True
    End If
End Function

