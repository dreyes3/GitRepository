VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmSimiliarNames 
   BackColor       =   &H00CCCCCC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CWT Desktop - Similiar Names List "
   ClientHeight    =   3420
   ClientLeft      =   2880
   ClientTop       =   2715
   ClientWidth     =   11370
   Icon            =   "frmSimiliarNames.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   11370
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
      LcK2            =   $"frmSimiliarNames.frx":038A
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
      Begin MyCommandButton.MyButton cmdSelect 
         Height          =   360
         Left            =   8880
         TabIndex        =   2
         Top             =   3045
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
         MouseIcon       =   "frmSimiliarNames.frx":03A9
         MousePointer    =   99
         Picture         =   "frmSimiliarNames.frx":06C3
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   13421772
         Caption         =   "&Select"
         Depth           =   1
         PictureOffsetX  =   2
         GradientType    =   2
      End
      Begin MyFramePanel.MyFrame MyFrame2 
         Height          =   2865
         Left            =   240
         Top             =   120
         Width           =   11100
         _ExtentX        =   19579
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
         HeaderGradientAlign=   5
         HeaderGradientSizeH=   "50%"
         HeaderColorTopLeft=   6973442
         HeaderColorTopRight=   6973442
         HeaderColorBottomLeft=   6973442
         HeaderColorBottomRight=   6973442
         HeaderShow      =   0   'False
         PictureOffsetX  =   5
         Begin MSComctlLib.ListView lvwNameList 
            Height          =   2200
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   10800
            _ExtentX        =   19050
            _ExtentY        =   3889
            View            =   3
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label1 
            BackColor       =   &H00DADAB6&
            Caption         =   $"frmSimiliarNames.frx":09CC
            Height          =   495
            Left            =   120
            TabIndex        =   1
            Top             =   200
            Width           =   10500
         End
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   360
         Left            =   10080
         TabIndex        =   3
         Top             =   3045
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
         TransparentColor=   13421772
         Caption         =   "&Cancel"
         Depth           =   1
         GradientType    =   2
      End
   End
End
Attribute VB_Name = "frmSimiliarNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bolProfile As Boolean
Private bolFirstSel As Boolean

Private Sub cmdCancel_Click()
    If bolProfile = True Then
       gbolCancelMove = True
    End If
    Unload Me
    'Preethi - V1.2.4 20110621  - ER01 - Tracking of Touches to CWT Booking
'    If gobjPNR.RecLoc <> "" Then
'      If gobjPNR.CompInfo.PNRTrackingTouches = True Then
'          Load frmPNRTouchTracking
'          frmPNRTouchTracking.Show
'       End If
'    End If
End Sub

Private Sub cmdSelect_Click()
    Dim intC As Integer
    Dim strTemp() As String
    Dim bolSelect As Boolean
    
    For intC = 1 To lvwNameList.ListItems.Count
        If lvwNameList.ListItems(intC).Selected = True Then
           bolSelect = True
           If bolProfile = False Then
              Set gobjPNR = New CWT_GalileoPNR3.PNR
              gobjPNR.loadPNR (lvwNameList.ListItems(intC).SubItems(3))
              displayPNRinBar
              If bolFirstSel = False Then frmBars.searchPanel
              bolFirstSel = True
           Else
              frmSideBar.txtLastName(1).Text = lvwNameList.SelectedItem
              frmSideBar.txtFirstName(1).Text = ""
              'frmSideBar.searchProfile
              Unload Me
           End If
           Exit For
        End If
    Next
    If bolSelect = False Then
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, "Must select a record to proceed ...", vbOKOnly + vbDefaultButton1, "Data Required"
    End If
End Sub

Private Sub Form_Load()
    Dim oldParent As Long
    
    Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)


    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then ' X button is clicked
       If bolProfile = True Then
          gbolCancelMove = True
       End If
    End If
    'Preethi - V1.2.4 20110621  - ER01 - Tracking of Touches to CWT Booking
    'CC - V1.2.4 20110711  - ER01 - Added LoadPNR script (Fix the bug)
    Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.loadPNR
    If gobjPNR.RecLoc <> "" Then
      If gobjPNR.CompInfo.PNRTrackingTouches = True Then
          Load frmPNRTouchTracking
          frmPNRTouchTracking.Show
       End If
    End If
End Sub

Private Sub lvwNameList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim oListItem As ListItem
    Dim strTemp As String
   
    If ColumnHeader.Text = "Segment Date" Then
        For Each oListItem In lvwNameList.ListItems
            oListItem.SubItems(ColumnHeader.Index - 1) = oListItem.Tag
        Next oListItem
    End If
    lvwNameList.SortKey = ColumnHeader.Index - 1
    lvwNameList.Sorted = True
    If lvwNameList.SortOrder = lvwAscending Then
        lvwNameList.SortOrder = lvwDescending
    Else
        lvwNameList.SortOrder = lvwAscending
    End If
    If ColumnHeader.Text = "Segment Date" Then
        For Each oListItem In lvwNameList.ListItems
            If oListItem.SubItems(ColumnHeader.Index - 1) <> "" Then
               strTemp = Format(CDate(oListItem.SubItems(ColumnHeader.Index - 1)), "DDMMMYY")
               oListItem.SubItems(ColumnHeader.Index - 1) = strTemp
            End If
        Next oListItem
    End If
End Sub

Private Sub lvwNameList_DblClick()
    Dim bolMove As Boolean
    Dim strTemp() As String
    
    If bolProfile = False Then
       'Retrieve PNR
       Set gobjPNR = New CWT_GalileoPNR3.PNR
       gobjPNR.loadPNR (lvwNameList.SelectedItem.SubItems(3))
       displayPNRinBar
       If bolFirstSel = False Then frmBars.searchPanel
       bolFirstSel = True
    Else
        frmSideBar.txtLastName(1).Text = lvwNameList.SelectedItem
        frmSideBar.txtFirstName(1).Text = ""
        'frmSideBar.searchProfile
        Unload Me
    End If
End Sub

Private Sub lvwNameList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       lvwNameList_DblClick
    End If
End Sub

Private Sub lvwNameList_KeyUp(KeyCode As Integer, Shift As Integer)
'If bolProfile = False Then
'Select Case KeyCode
'    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
'        lvwNameList_DblClick
'   End Select
'End If
End Sub
