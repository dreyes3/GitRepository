VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient.ocx"
Begin VB.Form frmStartUp 
   Caption         =   "CWT Desktop - Startup"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   6570
   Begin vbskpro.Skinner Skinner1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButton     =   1
      MaxButton       =   0
      MinButton       =   0
      OldForeColor    =   0
      ChangeSkinButton=   0   'False
      SysDisableSkinCaption=   "&Disable Skin"
      LcK1            =   "..02*-0..*/305*.-2-/"
      LcK2            =   $"frmStartUp.frx":038A
      AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
   End
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   6588
      Color           =   16447215
      FinColor        =   16048334
      Caption         =   "ARGradient1"
      ForeColor       =   -2147483630
      GradientSteps   =   65
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   360
         Top             =   3120
      End
      Begin MyCommandButton.MyButton cmdSelect 
         Height          =   360
         Left            =   5400
         TabIndex        =   0
         Top             =   3120
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
         MouseIcon       =   "frmStartUp.frx":03A9
         MousePointer    =   99
         Picture         =   "frmStartUp.frx":06C3
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   14215660
         Caption         =   "&Select"
         Depth           =   1
         PictureOffsetX  =   2
         GradientType    =   2
      End
      Begin MyFramePanel.MyFrame MyFrame2 
         Height          =   2865
         Left            =   240
         Top             =   120
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   5054
         BackColor       =   14342838
         ForeColor       =   15979465
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   13557
         BackgroundMaskColor=   13421772
         BackgroundAlignment=   4
         Caption         =   ""
         CornerTopLeft   =   -1  'True
         CornerTopRight  =   -1  'True
         CornerBottomLeft=   -1  'True
         CornerBottomRight=   -1  'True
         OutSideColor    =   14215660
         HeaderGradientAlign=   5
         HeaderGradientSizeH=   "50%"
         HeaderColorTopLeft=   6973442
         HeaderColorTopRight=   6973442
         HeaderColorBottomLeft=   6973442
         HeaderColorBottomRight=   6973442
         HeaderShow      =   0   'False
         PictureOffsetX  =   5
         Begin MSComctlLib.ListView lvwDatabase 
            Height          =   2205
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   5880
            _ExtentX        =   10372
            _ExtentY        =   3889
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   10231
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Country"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label1 
            BackColor       =   &H00DADAB6&
            Caption         =   "Please select one of the databases below:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   10500
         End
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JY - V1.2.6 20110916 - CR109 - Startup form for users to select database to be connected (AU ESC)

Private Sub cmdSelect_Click()
    
    Dim i As Integer
    Dim strCountry As String
    Dim bolSelect As Boolean
    Dim strMsg As String
    
    bolSelect = False
    With lvwDatabase
          For i = 1 To .ListItems.Count
              If .ListItems(i).Selected = True Then
                  bolSelect = True
                  strCountry = .ListItems(i).SubItems(1)
                  Exit For
              End If
          Next
    End With
    
    If bolSelect = False Then
       strMsg = "Please select a database to be connected to ..."
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
       Exit Sub
    End If
    
    gstrDatabaseConnected = strCountry
    Unload Me
    
End Sub

Private Sub Form_Load()
   
    Dim oldParent As Long
    Dim i As Integer

    oldParent = SetParent(Me.hwnd, gVPMDIHwnd)
    Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
    Skinner1.CloseButton = skNo
    Me.Move 0, 0
    Me.Move (gFPWidth / 2) - (Me.Width / 2), 0
    

    Set item = lvwDatabase.ListItems.Add(, , "Singapore Database")
    item.SubItems(1) = "SG"
    Set item = lvwDatabase.ListItems.Add(, , "Hong Kong Database")
    item.SubItems(1) = "HK"
    
    For i = 1 To lvwDatabase.ListItems.Count
        lvwDatabase.ListItems(i).Selected = False
    Next
   
End Sub

Private Sub lvwDatabase_DblClick()
    cmdSelect_Click
End Sub

Private Sub timer1_Timer()
    Dim lCurHwnd As Long
    Dim buf As String * 256
    Dim title As String
    Dim length As Long
    Dim loadedForm As Form
    
    'Get the window text of Galileo Desktop
    length = GetWindowText(glngTargetHwnd, buf, Len(buf))
    title = Left$(buf, length)
    If InStr(1, UCase(title), UCase("Galileo Desktop - [")) > 0 Then
       'Hide all loaded forms if is in viewpoint mode
        If gMode <> "V" Then
            For Each loadedForm In Forms
                If loadedForm.Visible = True And UCase(loadedForm.Name) <> UCase("frmDDEOwner") Then
                   loadedForm.Visible = False
                   Set loadedForm = Nothing
                End If
            Next
            gMode = "V"
        End If
    Else
        'Show all loaded forms if is in focalpoint mode
        If gMode <> "F" Then
            For Each loadedForm In Forms
                If loadedForm.Visible = False And UCase(loadedForm.Name) <> UCase("frmDDEOwner") Then
                   loadedForm.Visible = True
                   loadedForm.WindowState = 0
                   loadedForm.SetFocus
                   Set loadedForm = Nothing
                End If
            Next
            gMode = "F"
        End If
    End If

End Sub
