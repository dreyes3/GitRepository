VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmSeat 
   Caption         =   "CWT Desktop - Seat"
   ClientHeight    =   3390
   ClientLeft      =   1095
   ClientTop       =   3990
   ClientWidth     =   11340
   Icon            =   "frmSeat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   11340
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   5025
      Left            =   0
      Top             =   -120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8864
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
      Begin MyCommandButton.MyButton cmdPrevious 
         Height          =   360
         Left            =   7560
         TabIndex        =   17
         Top             =   3240
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
      Begin MyCommandButton.MyButton cmdFinishAll 
         Height          =   360
         Left            =   9120
         TabIndex        =   0
         Top             =   3250
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
         Caption         =   "&Finish"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   360
         Left            =   10200
         TabIndex        =   1
         Top             =   3250
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
      Begin MyCommandButton.MyButton cmdQuickFinish 
         Height          =   345
         Left            =   9120
         TabIndex        =   2
         Top             =   3465
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   609
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
         Caption         =   "&Finish"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyFramePanel.MyFrame topFrame 
         Height          =   3075
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   5424
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
         Begin VB.Frame seatFrm 
            BackColor       =   &H00DADAB6&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1935
            Left            =   4320
            TabIndex        =   18
            Top             =   360
            Width           =   6255
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexSeatMap 
               Height          =   1830
               Index           =   0
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Visible         =   0   'False
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   3228
               _Version        =   393216
               BackColor       =   16777215
               Rows            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               FocusRect       =   0
               HighLight       =   0
               GridLinesFixed  =   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   0
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexSeatMap 
               Height          =   1605
               Index           =   1
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Visible         =   0   'False
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   2831
               _Version        =   393216
               Rows            =   0
               FixedRows       =   0
               FixedCols       =   0
               GridLinesFixed  =   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   0
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexSeatMap 
               Height          =   1830
               Index           =   2
               Left            =   0
               TabIndex        =   23
               Top             =   0
               Visible         =   0   'False
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   3228
               _Version        =   393216
               BackColor       =   16777215
               Rows            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               FocusRect       =   0
               HighLight       =   0
               GridLinesFixed  =   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   0
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexSeatMap 
               Height          =   1830
               Index           =   3
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Visible         =   0   'False
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   3228
               _Version        =   393216
               BackColor       =   16777215
               Rows            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               FocusRect       =   0
               HighLight       =   0
               GridLinesFixed  =   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   0
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexSeatMap 
               Height          =   1830
               Index           =   4
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Visible         =   0   'False
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   3228
               _Version        =   393216
               BackColor       =   16777215
               Rows            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               FocusRect       =   0
               HighLight       =   0
               GridLinesFixed  =   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   0
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexSeatMap 
               Height          =   1830
               Index           =   5
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Visible         =   0   'False
               Width           =   6150
               _ExtentX        =   10848
               _ExtentY        =   3228
               _Version        =   393216
               BackColor       =   16777215
               Rows            =   0
               FixedRows       =   0
               FixedCols       =   0
               BackColorBkg    =   16777215
               FocusRect       =   0
               HighLight       =   0
               GridLinesFixed  =   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   0
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin VB.Label lblNoMap 
               BackColor       =   &H00DADAB6&
               Caption         =   "No Seat Map Available"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   0
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.Label lblCannotAssign 
               BackColor       =   &H00DADAB6&
               Caption         =   "Unable to assign seat for unconfirm air segment. Only Generic seat assigment is allowed."
               Height          =   615
               Left            =   120
               TabIndex        =   21
               Top             =   0
               Visible         =   0   'False
               Width           =   4575
            End
         End
         Begin MyFramePanel.MyFrame fraPreference 
            Height          =   675
            Left            =   4320
            Top             =   2280
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   1191
            BackColor       =   14342838
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            AppearanceThemes=   3
            BackgroundAlignment=   4
            Caption         =   "Select Seat Preferences:"
            CaptionAlignment=   0
            CornerRadius    =   10
            CornerTopLeft   =   -1  'True
            CornerTopRight  =   -1  'True
            OutSideColor    =   14215660
            HeaderHeight    =   20
            HeaderColorTopLeft=   0
            HeaderColorTopRight=   0
            HeaderColorBottomLeft=   0
            HeaderColorBottomRight=   0
            Begin VB.CheckBox chkSkipAssign 
               BackColor       =   &H00DADAB6&
               Caption         =   "Skip Assigned Seat"
               Height          =   255
               Left            =   3000
               TabIndex        =   15
               Top             =   600
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.CheckBox chkAll 
               BackColor       =   &H00DADAB6&
               Caption         =   "All Segment(s) and Passenger(s)"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   600
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.ComboBox cmbSeat 
               Height          =   315
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   240
               Width           =   2175
            End
            Begin VB.ComboBox cmbSmoking 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   240
               Width           =   2295
            End
            Begin MyCommandButton.MyButton cmdSet 
               Height          =   345
               Left            =   4800
               TabIndex        =   13
               Top             =   720
               Visible         =   0   'False
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   609
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
               AppearanceThemes=   2
               BackColorDown   =   3968251
               BackColorOver   =   6805503
               BackColorFocus  =   16765357
               BackColorDisabled=   12648447
               TransparentColor=   14215660
               Caption         =   "&Set"
               Depth           =   1
               GradientType    =   2
            End
         End
         Begin VB.ComboBox cmbDeck 
            Height          =   315
            Left            =   8880
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   50
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.PictureBox picMapIcon 
            Appearance      =   0  'Flat
            BackColor       =   &H00DADAB6&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2655
            Left            =   10440
            Picture         =   "frmSeat.frx":038A
            ScaleHeight     =   2655
            ScaleWidth      =   855
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSComctlLib.TreeView tvSeat 
            Height          =   2535
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   4471
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   441
            LabelEdit       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ComctlLib.ImageList ImageList 
            Left            =   10680
            Top             =   2280
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   30
            ImageHeight     =   26
            MaskColor       =   12632256
            UseMaskColor    =   0   'False
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   6
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "frmSeat.frx":1208
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "frmSeat.frx":199A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "frmSeat.frx":212C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "frmSeat.frx":28BE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "frmSeat.frx":3050
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "frmSeat.frx":3282
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label2 
            BackColor       =   &H00DADAB6&
            BackStyle       =   0  'Transparent
            Caption         =   "Select Seat"
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
            Left            =   4320
            TabIndex        =   8
            Top             =   45
            Width           =   3135
         End
         Begin VB.Label Label1 
            BackColor       =   &H00DADAB6&
            BackStyle       =   0  'Transparent
            Caption         =   "Select Segment and Pax"
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
            TabIndex        =   7
            Top             =   50
            Width           =   3135
         End
      End
      Begin vbskpro.Skinner Skinner1 
         Left            =   0
         Top             =   480
         _ExtentX        =   1270
         _ExtentY        =   1270
         CloseButton     =   1
         MaxButton       =   0
         MinButton       =   0
         OldForeColor    =   0
         ChangeSkinButton=   0   'False
         SysDisableSkinCaption=   "&Disable Skin"
         LcK1            =   "..02*-0..*/305*.-2-/"
         LcK2            =   $"frmSeat.frx":3A14
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
      Begin MyCommandButton.MyButton cmdReset 
         Height          =   345
         Left            =   8040
         TabIndex        =   16
         Top             =   3480
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
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
         AppearanceThemes=   2
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   14215660
         Caption         =   "&Refresh"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdAssignFinish 
         Height          =   360
         Left            =   9000
         TabIndex        =   3
         Top             =   3465
         Visible         =   0   'False
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
         Caption         =   "&Finish"
         Depth           =   1
         GradientType    =   2
      End
      Begin VB.Label lblSeatRmks 
         BackColor       =   &H00F4E0CE&
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   3210
         Width           =   6855
      End
      Begin VB.Label lblSeatRmk 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Seat Remarks: "
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
         TabIndex        =   9
         Top             =   3210
         Width           =   1575
      End
   End
   Begin VB.Menu mnuPopUpFlex 
      Caption         =   "Pop Up Flex"
      Visible         =   0   'False
      Begin VB.Menu subMenuAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu subMenuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "frmSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrFlex As String
Dim mstrSmoke As String
Dim mstrSeat As String
Dim strCurrentSegMap As String
Dim blnAutoPop As Boolean
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date



Private Sub chkAll_Click()

    If chkAll.value = 1 Then
        
        cmbSeat.Enabled = True
        cmbSmoking.Enabled = True
        cmdSet.Visible = False
        tvSeat.Enabled = False
        flexSeatMap(0).Enabled = False
        flexSeatMap(1).Enabled = False
        chkSkipAssign.Visible = True
        chkSkipAssign.value = 0
        
    Else
        
        If lblNoMap.Visible = False Then
            cmbSeat.Enabled = False
            cmbSmoking.Enabled = False
            flexSeatMap(0).Enabled = True
            flexSeatMap(1).Enabled = True
        End If
        
        tvSeat.Enabled = True
        chkSkipAssign.Visible = False
        chkSkipAssign.value = 0
        cmdSet.Enabled = True
        cmdSet.Visible = True
    
    End If

End Sub

Private Sub cmbDeck_Click()
Dim intIndex As Integer
Dim intI As Integer

    intIndex = cmbDeck.listindex
    
    For intI = 0 To flexSeatMap.Count - 1
        If intI = intIndex Then
            flexSeatMap(intI).Visible = True
        Else
            flexSeatMap(intI).Visible = False
        End If
    Next

End Sub
Private Function AssignFinish() As String
Dim intI As Integer
Dim strSegNo As String
Dim strPaxNo As String
Dim strTemp As String
Dim node As MSComctlLib.node
Dim childnode As MSComctlLib.node
Dim strSeat As String
Dim intJ As Integer
Dim strCmd As String
Dim strResp As String
Dim bolSame As Boolean
Dim strSC As String
Dim strSeatChr As String




    For Each node In tvSeat.Nodes
            
            strSegNo = Trim(Mid(node.Text, 1, InStr(node.Text, ".") - 1))
            
            If InStr(node.key, "S") > 0 Then
                If node.children > 0 Then
                    strTemp = node.Child.Text
                    
                    For intI = 1 To node.children
                        strPaxNo = Mid(strTemp, 2, InStr(strTemp, ".") - 2)
                        strSeat = ""
                        
                        If InStr(strTemp, "Seat:") > 0 Then
                       
                            strSeat = Mid(strTemp, InStr(strTemp, "Seat:") + 5)
                            If InStr(strSeat, "Status:") > 0 Then strSeat = Trim(Mid(strSeat, 1, InStr(strSeat, "Status:") - 2))
                            If InStr(strSeat, "Type:") > 0 Then strSeat = Trim(Mid(strSeat, 1, InStr(strSeat, "Type:") - 3))
                    
                        End If
                    
                        strSC = ""
                        
                        If InStr(strTemp, "Type:") > 0 Then
                            strSC = Mid(strTemp, InStr(strTemp, "Type:"))
                            strSC = Mid(strTemp, InStr(strTemp, "Type:") + 5)
                        End If
                    
                        strCmd = ""
                        
                        For intJ = 1 To gobjPNR.SeatDataCount
                            
                            If gobjPNR.SeatData(intJ).PaxNo = strPaxNo And gobjPNR.SeatData(intJ).SegNum = strSegNo Then
                                
                                If strSeat = "" And strSC = "" Then ' Delete
                                    
                                    strCmd = "S.P" & strPaxNo & "S" & strSegNo & "@"
                                    strResp = gobjHost.terminalEntry(strCmd)
                                    
                                    If InStr(strResp, "CANCELLED SEAT") Then
                                
                                    Else
                                    
                                        strError = strError & "Unable to cancel seat(Cmd: " & strCmd & "), response from GDS is:" & vbCrLf & strResp & vbCrLf
                                
                                    End If
                                    
                                End If
                                
                                If strSeat <> "" And gobjPNR.SeatData(intJ).SegNum > 0 And convertSeat(gobjPNR.SeatData(intJ).SeatLocation) <> Trim(strSeat) Then  'change
                                    
                                    strCmd = "S.P" & strPaxNo & "S" & strSegNo & "@" & strSeat
                                    strResp = gobjHost.terminalEntry(strCmd)
                                    
                                    If InStr(strResp, "CHANGED SEAT") Then
                                
                                    Else
                                        
                                        strError = strError & "Unable to change seat(Cmd: " & strCmd & "), response from GDS is:" & vbCrLf & strResp & vbCrLf
                                
                                    End If
                                
                                End If
                                
                                strSeatChr = ""
                                strSeatChr = gobjPNR.SeatData(intJ).SeatAttribute1
                                
                                If gobjPNR.SeatData(intJ).SeatAttribute2 <> "" Then
                                    strSeatChr = strSeatChr & gobjPNR.SeatData(intI).SeatAttribute2
                                End If
                    
                                If strSC <> "" And strSeat = "" And gobjPNR.SeatData(intJ).SegNum > 0 And strSeatChr <> strSC Then 'change
                                    
                                    strCmd = "S.P" & strPaxNo & "S" & strSegNo & "@" & strSC
                                    strResp = gobjHost.terminalEntry(strCmd)
                                    
                                    If InStr(strResp, "CHANGED SEAT") Then
                                
                                    Else
                                        
                                        strError = strError & "Unable to change seat(Cmd: " & strCmd & "), response from GDS is:" & vbCrLf & strResp & vbCrLf
                                
                                    End If
                                
                                End If
                                
                                Exit For
                                
                            End If
                        
                        Next intJ
                   
                    
                    If strCmd = "" Then ' add new
                    
                        bolSame = False
                    
                        If strSeat <> "" Then
                            
                            For intJ = 1 To gobjPNR.SeatDataCount
                                
                                If convertSeat(gobjPNR.SeatData(intJ).SeatLocation) = Trim(strSeat) Then
                                        
                                        bolSame = True
                                End If
                                
                            Next intJ
                            
                            If bolSame = False Then
                                strCmd = "S.P" & strPaxNo & "S" & strSegNo & "/" & strSeat
                                strResp = gobjHost.terminalEntry(strCmd)
                                
                                If InStr(strResp, "RESERVED SEAT") Then
                            
                                Else
                                
                                    strError = strError & "Unable to reserve seat(Cmd: " & strCmd & "), response from GDS is:" & vbCrLf & strResp & vbCrLf
                                   
                                End If
                                
                            End If
                            
                        End If
                        
                        If strSC <> "" And strSeat = "" Then
                            
                            bolSame = False
                            strSeatChr = ""
                            
                            For intJ = 1 To gobjPNR.SeatDataCount
                                strSeatChr = gobjPNR.SeatData(intJ).SeatAttribute1
                                
                                If gobjPNR.SeatData(intJ).SeatAttribute2 <> "" Then
                                    
                                    strSeatChr = strSeatChr & gobjPNR.SeatData(intJ).SeatAttribute2
                                
                                End If
                                
                                If strSeatChr = strSC And gobjPNR.SeatData(intJ).PaxNo = strPaxNo And gobjPNR.SeatData(intJ).SegNum = strSegNo Then
                                    
                                    bolSame = True
                                
                                End If
                                
                            Next intJ
                            
                            If bolSame = False Then
                            
                            strCmd = "S.P" & strPaxNo & "S" & strSegNo & "/" & strSC
                            strResp = gobjHost.terminalEntry(strCmd)
                                
                                If InStr(strResp, "RESERVED SEAT") Then
                            
                                Else
                                
                                    strError = strError & "Unable to reserve seat(Cmd: " & strCmd & "), response from GDS is:" & vbCrLf & strResp & vbCrLf
                                   
                                End If
                                
                            End If
                            
                        End If
                        
                    End If
            
                    If node.children > 1 Then strTemp = node.Child.Next
                    
                Next intI
                    
            End If
            
        End If
            
    Next node


If strError <> "" Then
    'MsgBox strError, , "CWT Agent Desktop - Seat"
    AssignFinish = strError
    
Else

    AssignFinish = ""

End If



End Function
Private Function convertSeat(seat As String) As String

Dim intI As Integer
Dim strTemp As String
    convertSeat = ""
    
    For intI = 1 To Len(seat)
        
        strTemp = Mid(seat, intI, 1)
        
            If IsNumeric(strTemp) = False Then
                
                    convertSeat = convertSeat & " " & strTemp
                
            Else
            
                If intI = 1 And strTemp = CStr(0) Then
                
                Else
                    
                    convertSeat = convertSeat & strTemp
                
                End If
                
            End If
        
    Next intI

End Function
Private Sub cmbSeat_Click()

        UpdateSCToList

End Sub

Private Sub cmbSmoking_Click()
        
        UpdateSCToList

End Sub
Private Sub UpdateSCToList()

Dim strTemp As String
        
    If blnAutoPop = True Then Exit Sub
    
    If cmbSmoking.Text <> "" Then
    
        strTemp = Trim(Left(cmbSmoking.Text, 1))
    
    End If
    
    If cmbSeat.Text <> "" Then
        
        strTemp = strTemp & Trim(Left(cmbSeat.Text, 1))
    
    End If
    
    If Len(strTemp) = 1 Then Exit Sub
    
    SeatPax
    clearfocusSeat
    clearOldSeat

    If strTemp = "" Then
        
        AssignSeatChr ""
    
    Else
        
        AssignSeatChr strTemp
        
    End If
    
End Sub
Private Sub cmdCancel_Click()
    gbolCancelProcess = True
    Unload Me

End Sub

Private Sub cmdNext_Click()

    If SftTabs.Tabs.Current + 1 < SftTabs.Tabs.Count Then
    
       SftTabs.Tabs.Current = SftTabs.Tabs.Current + 1
    
    End If

End Sub

Private Sub cmdFinishAll_Click()

Dim strErr As String
'If validateDuplicate = True Then Exit Sub

'QuickFinish
'If chkAll.Value = 1 Then
'    QuickFinish
'Else
    datTouchEnd = Now
    If CheckPNRStatus = cwtPNRSimulChange Then
    
        gobjHost.terminalEntry "IR"
    
    End If
        
    strErr = AssignFinish
    
    If strErr = "" Then
        
        gobjHost.terminalEntry "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
        gobjHost.terminalEntry "ER"
        gobjHost.terminalEntry "ER"
        pDisplayToFP ("*SD")
        
        'Log formload
        'Backup on 26 Sept - Jeremy
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModSeat), _
'       IIf(gbolCreatPNR = True, gconSModSeat, ""), Me.Name, gconFormLoad, gstrProcessGrpID, _
'       datFormLoadEnd, datFormLoadStart
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModSeat), _
'       IIf(gbolCreatPNR = True, gconSModSeat, ""), Me.Name, gconTouch, gstrProcessGrpID, _
'       datTouchEnd, datFormLoadEnd
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModSeat), _
'       IIf(gbolCreatPNR = True, gconSModSeat, ""), Me.Name, gconProcessing, gstrProcessGrpID, _
'        , datTouchEnd
        
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModSeat, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModSeat, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModSeat, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
        
        Unload Me
    
    Else
        
        gobjHost.terminalEntry "IR"
        'MsgBox strErr, , "CWT Agent Desktop - Seat"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strErr, vbOKOnly + vbDefaultButton1, "CWT Desktop - Seat"
    End If





End Sub

Private Function validateDuplicate() As Boolean
Dim intRow As Integer
Dim intI As Integer
Dim node As MSComctlLib.node
Dim strValidError As String
Dim strPax As String
Dim strSegment As String
Dim intNextRow As Integer
Dim strTemp As String
Dim strTemp2 As String


    validateDuplicate = False
    strValidError = ""
    
    For intRow = 1 To msFlexSeat.rows - 1
        
        If InStr(msFlexSeat.TextMatrix(intRow, 3), "ALL") > 1 Then
            'strTemp = msFlexSeat.TextMatrix(intRow, 2)
            For Each node In tvSeat.Nodes
                
                If InStr(node.Text, "Seat:") Then
                
                    strPax = Mid(node.Text, 2, InStr(node.Text, ".") - 2)
                    strSegment = Trim(Mid(node.Parent, 1, InStr(node.Parent, ".") - 1))
                    strValidError = strValidError & "Duplicate seat assignment for Pax" & strPax & " Segment" & strSegment & " in Seat Map and Quick Seat Assignment" & vbCrLf
                    Exit For
                
                End If
                
            Next node
            
        End If
                
    Next
    
    strPax = ""
    strSegment = ""
    
    For intRow = 1 To msFlexSeat.rows - 1
    
        For Each node In tvSeat.Nodes
             
             
            If InStr(node.Text, "Seat:") Then
            
                        strPax = Mid(node.Text, 2, InStr(node.Text, ".") - 2)
                        strSegment = Trim(Mid(node.Parent, 1, InStr(node.Parent, ".") - 1))
                                                
            End If
            
            If msFlexSeat.TextMatrix(intRow, 3) <> " ALL" Then
                    
                    strTemp2 = Trim(Mid(msFlexSeat.TextMatrix(intRow, 3), 1, InStr(msFlexSeat.TextMatrix(intRow, 3), ".") - 1))
            
            End If
            
            If Mid(msFlexSeat.TextMatrix(intRow, 2), 1, InStr(msFlexSeat.TextMatrix(intRow, 2), ".") - 1) = strPax And _
            strTemp2 = strSegment Then
            
                strValidError = strValidError & "Duplicate seat assignment for Pax" & strPax & " Segment " & strSegment & " in Seat Map and Quick Seat Assignment" & vbCrLf
                Exit For
                
            End If
             
         Next node
                   
    Next
    
    For intRow = 1 To msFlexSeat.rows - 1
    
            If msFlexSeat.TextMatrix(intRow, 0) = "" And msFlexSeat.TextMatrix(intRow, 1) = "" Then
                
                strValidError = strValidError & "Missing preference for Pax" & Left(msFlexSeat.TextMatrix(intRow, 2), InStr(msFlexSeat.TextMatrix(intRow, 2), ".") - 1) & " Segment " & Trim(Left(msFlexSeat.TextMatrix(intRow, 3), InStr(1, msFlexSeat.TextMatrix(intRow, 3), ".") - 1)) & vbCrLf
                               
            End If
                    
    Next
    
    For intRow = 1 To msFlexSeat.rows - 2
    
            For intNextRow = intRow + 1 To msFlexSeat.rows - 1
                
                If msFlexSeat.TextMatrix(intRow, 2) = msFlexSeat.TextMatrix(intNextRow, 2) And _
                msFlexSeat.TextMatrix(intRow, 3) = msFlexSeat.TextMatrix(intNextRow, 3) Then
                
                    strValidError = strValidError & "Duplicate seat assignment for Pax" & Left(msFlexSeat.TextMatrix(intRow, 2), InStr(msFlexSeat.TextMatrix(intRow, 2), ".") - 1) & " Segment " & Trim(Left(msFlexSeat.TextMatrix(intRow, 3), InStr(1, msFlexSeat.TextMatrix(intRow, 3), ".") - 1)) & vbCrLf
                
                End If
                
            Next
                   
    Next
    
    strTemp = ""
    
    For intRow = 1 To msFlexSeat.rows - 1
        
        If InStr(msFlexSeat.TextMatrix(intRow, 3), "ALL") > 1 Then
        
            strTemp = msFlexSeat.TextMatrix(intRow, 2)
            
            For intNextRow = 2 To msFlexSeat.rows - 1
                
                If msFlexSeat.TextMatrix(intNextRow, 3) <> "ALL" Then
                    
                    If strTemp = msFlexSeat.TextMatrix(intNextRow, 2) Then
                        
                        strValidError = strValidError & "Duplicate seat assignment for Pax" & Left(msFlexSeat.TextMatrix(intNextRow, 2), InStr(msFlexSeat.TextMatrix(intNextRow, 2), ".") - 1) & vbCrLf
                        Exit For
                    
                    End If
                
                End If
            
            Next
        
        End If
    
    Next
    
    If strValidError <> "" Then
    
        'MsgBox strValidError, , "CWT Agent Desktop - Seat"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strValidError, vbOKOnly + vbDefaultButton1, "CWT Desktop - Seat"
        validateDuplicate = True
        Exit Function
    
    End If


End Function


Private Sub QuickFinish()
Dim strCmd As String
Dim strPax As String
Dim intI As Integer
Dim intJ As Integer
Dim intK As Integer
Dim bolFound As Boolean

Dim strAirSeg As String
Dim strResp As String

Dim strSmoking As String
Dim strSeat As String
Dim strError As String
Dim strValidError As String
Dim intNextRow As Integer
Dim strTemp As String


    If chkSkipAssign.value = 0 Then
    
        strCmd = "S.@"
        strResp = gobjHost.terminalEntry(strCmd)
        
        If InStr(strResp, "CANCELLED SEAT") Then
            
        Else
            
            strError = strError & "Unable to cancel seat(Cmd: " & strCmd & "), response from GDS is:" & vbCrLf & strResp & vbCrLf
            
        End If

        If cmbSmoking <> "" Or cmbSeat <> "" Then
            
            If cmbSmoking <> "" Then strSmoking = Trim(Left(cmbSmoking, InStr(1, cmbSmoking, "-") - 1))
            If strSeat <> "" Then strSeat = Trim(Left(cmbSeat, InStr(1, cmbSeat, "-") - 1))
            strCmd = "S." & strSmoking & strSeat
            strResp = gobjHost.terminalEntry(strCmd)
            
            If InStr(strResp, "RESERVED SEAT") Then
        
            Else
                
                strError = strError & "Unable to reserve seat(Cmd: " & strCmd & "), response from GDS is:" & vbCrLf & strResp & vbCrLf
               
            End If
        
        End If
                    
    Else
    
        For intI = 1 To gobjPNR.AirSegCount
        
            For intJ = 1 To gobjPNR.PassengerCount
                
                For intK = 1 To gobjPNR.SeatDataCount
                
                    If gobjPNR.SeatData(intK).SegNum = gobjPNR.AirSeg(intI).segnumber And gobjPNR.SeatData(intK).PaxNo = intJ Then
                        
                        bolFound = True
                        Exit For
                    
                    End If
                
                Next intK
            
                
                If bolFound = False Then
                    
                    If cmbSmoking <> "" Or cmbSeat <> "" Then
                         
                        If cmbSmoking <> "" Then strSmoking = Trim(Left(cmbSmoking, InStr(1, cmbSmoking, "-") - 1))
                        If strSeat <> "" Then strSeat = Trim(Left(cmbSeat, InStr(1, cmbSeat, "-") - 1))

                        strCmd = "S." & "S" & gobjPNR.AirSeg(intI).segnumber & "P" & intJ & "/" & cmbSmoking & strSeat
                        strResp = gobjHost.terminalEntry(strCmd)
                        
                        If InStr(strResp, "RESERVED SEAT") Then
                    
                        Else
                            
                            strError = strError & "Unable to reserve seat(Cmd: " & strCmd & "), response from GDS is:" & vbCrLf & strResp & vbCrLf
                           
                        End If
                        
                    End If
                    
                End If
                        
            Next intJ
            
        Next intI
          
    End If

    
    If strError <> "" Then
        'MsgBox strError, , "CWT Agent Desktop - Seat"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strError, vbOKOnly + vbDefaultButton1, "CWT Desktop - Seat"
    End If

End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload Me
End Sub

Private Sub cmdReset_Click()
    Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.loadPNR
    PopulateControls

End Sub

Private Sub cmdSet_Click()

Dim strTemp As String
    
    If cmbSmoking.Text <> "" Then
        strTemp = Trim(Left(cmbSmoking.Text, 1))
    End If
    
    If cmbSeat.Text <> "" Then
        strTemp = strTemp & Trim(Left(cmbSeat.Text, 1))
    End If
    
    SeatPax
     
    If strTemp = "" Then
        AssignSeatChr ""
    Else
        AssignSeatChr strTemp
    End If
       
End Sub
Private Sub AssignSeatChr(seatChr As String)
Dim intI As Integer

    With tvSeat
    
        For intI = 1 To .Nodes.Count
        
           If .Nodes.item(intI).Selected Then
           'Or chkAll.Value = 1 Then
               If InStr(.Nodes.item(intI).key, "S") = 0 Then
                   
                   If InStr(.Nodes.item(intI).Text, "Type:") Then
                   
                        .Nodes.item(intI).Text = Mid(.Nodes.item(intI).Text, 1, InStr(.Nodes.item(intI).Text, "Type:") - 1)
                        
                   End If
                   
                   If InStr(.Nodes.item(intI).Text, "Seat:") Then
                   
                        .Nodes.item(intI).Text = Mid(.Nodes.item(intI).Text, 1, InStr(.Nodes.item(intI).Text, "Seat:") - 1)
                        
                   End If
                   
                   If seatChr <> "" Then
                   
                        .Nodes.item(intI).Text = Trim(.Nodes.item(intI).Text) & "  Type:" & seatChr
                        
                   End If
               
               End If
               
           End If
           
        Next
        
    End With
    
End Sub
Private Sub flexSeatMap_DblClick(Index As Integer)

Dim strTemp As String
Dim strOldSeat As String
        
        
      strTemp = Trim(flexSeatMap(Index).TextMatrix(flexSeatMap(Index).row, 0)) & " " & Trim(flexSeatMap(Index).ColHeaderCaption(0, flexSeatMap(Index).col))
      SeatPax
       
      If IsNumeric(Trim(strTemp)) Then
      
          Exit Sub
      
      End If
       
      If otherSelected(strTemp) = True Then
      
          Exit Sub
      
      End If
      
      If flexSeatMap(Index).Text <> "" Then
      
          If sameSeat(Index) = True Then
          
             Set flexSeatMap(Index).CellPicture = ImageList.ListImages(1).Picture
             'LoadPicture (App.Path & "\Icons\seat\seat_savail.gif")
             'clearfocusSeat
             
          Else
              
              flexSeatMap(Index).FocusRect = flexFocusHeavy
              Set flexSeatMap(Index).CellPicture = ImageList.ListImages(3).Picture
              'LoadPicture(App.Path & "\Icons\seat\seat_sbook.gif")
              'strTemp = Trim(flexSeatMap(Index).TextMatrix(flexSeatMap(Index).Row, 0)) & " " & Trim(flexSeatMap(Index).Text)
              clearfocusSeat
              clearOldSeat
              focusSeat strTemp
              AssignSeattoTree strTemp
              
          End If
    
      Else
          'clearfocusSeat
          clearOldSeat
          AssignSeattoTree ""
      End If
      
      ClearSeatPreference
        
        
End Sub
Private Sub ClearSeatPreference()

    blnAutoPop = True
    cmbSeat.listindex = 0
    cmbSmoking.listindex = 0
    blnAutoPop = False

End Sub

Private Function getOldSeat() As String
Dim intI As Integer
Dim strTemp As String

    With tvSeat
    
        For intI = 1 To .Nodes.Count
        
           If .Nodes.item(intI).Selected Then
           
               If InStr(.Nodes.item(intI).key, "S") = 0 Then
                   
                   If InStr(.Nodes.item(intI).Text, "Seat:") Then
                   
                        strTemp = Mid(.Nodes.item(intI).Text, InStr(.Nodes.item(intI).Text, "Seat:") + 5)
                        
                   End If
                   
                   If InStr(strTemp, "Status:") Then
                   
                        strTemp = Trim(Mid(strTemp, 1, InStr(strTemp, "Status:") - 1))
                        
                   End If
                   
                   If InStr(strTemp, "Type:") Then
                   
                        strTemp = Trim(Mid(strTemp, 1, InStr(strTemp, "Type:") - 1))
                        
                   End If
                   
               End If
               
           End If
           
        Next
        
    End With
    
    getOldSeat = strTemp

End Function
Private Function otherSelected(seat As String) As Boolean

Dim intI As Integer
Dim strTemp As String
Dim strChildSeat As String
Dim intJ As Integer
Dim strActual As String
Dim strPax As String
    With tvSeat
    
        
            
        For intI = 1 To .Nodes.Count
             
            If .Nodes.item(intI).Selected Then
                
                strActual = Mid(.Nodes.item(intI).Text, 2, InStr(.Nodes.item(intI).Text, ".") - 2)
                strTemp = .Nodes.item(intI).Parent
                Exit For
                
            End If
            
        Next
        
        For intI = 1 To .Nodes.Count
                        
            If .Nodes.item(intI).Text = strTemp Then
                
                strChildSeat = ""
               
                If InStr(.Nodes.item(intI).Child.Text, ":") > 0 Then
                    
                    strChildSeat = Mid(.Nodes.item(intI).Child.Text, InStr(.Nodes.item(intI).Child.Text, ":") + 1)
                    strPax = Mid(.Nodes.item(intI).Child.Text, 2, InStr(.Nodes.item(intI).Child.Text, ".") - 2)
                
                End If
               
               If .Nodes.item(intI).children > 1 Then
                    
                    For intJ = 1 To .Nodes.item(intI).children
                
                       If strChildSeat = "" Then
                        
                             strChildSeat = ""
                             strChildSeat = Mid(.Nodes.item(intI).Child.Next, InStr(.Nodes.item(intI).Child.Next, ":") + 1)
                             strPax = Mid(.Nodes.item(intI).Child.Next, 2, InStr(.Nodes.item(intI).Child.Next, ".") - 2)
                        
                        Else
                       
                
                             If strChildSeat = seat And strPax <> strActual Then
                                
                                otherSelected = True
                                Exit For
                             
                             Else
                                
                                strChildSeat = ""
                                
                                If .Nodes.item(intI).children > 1 Then
                                    
                                    If InStr(.Nodes.item(intI).Child.Next, ":") > 0 Then
                                        strChildSeat = Mid(.Nodes.item(intI).Child.Next, InStr(.Nodes.item(intI).Child.Next, ":") + 1)
                                        strPax = Mid(.Nodes.item(intI).Child.Next, 2, InStr(.Nodes.item(intI).Child.Next, ".") - 2)
                                    Else
                                    
                                        strChildSeat = ""
                                        strPax = Mid(.Nodes.item(intI).Child.Next, 2, InStr(.Nodes.item(intI).Child.Next, ".") - 2)
                                    End If
                                
                                End If
                                
                             End If
        
                        End If
               
                    Next intJ
               
               End If
               
            End If
            
            If otherSelected = True Then Exit For
            
        Next
        
    End With

End Function

Private Function sameSeat(mapindex) As Boolean

Dim intI As Integer
Dim strOldSeat As String
Dim strTemp() As String
Dim strOldRow As String
Dim strOldCol As String
With tvSeat

    
        
         For intI = 1 To .Nodes.Count
         
          If .Nodes.item(intI).Selected Then
            If InStr(tvSeat.Nodes.item(intI).Text, "Seat:") Then
                 strOldSeat = Mid(.Nodes.item(intI).Text, InStr(.Nodes.item(intI).Text, ":") + 1)
                 strTemp = Split(strOldSeat, " ")
                 
                 strOldRow = strTemp(0)
                 strOldCol = strTemp(1)
                 
                 If strOldRow = Trim(flexSeatMap(mapindex).TextMatrix(flexSeatMap(mapindex).row, 0)) _
                 And strOldCol = Trim(flexSeatMap(mapindex).Text) Then
                    .Nodes.item(intI).Text = Mid(.Nodes.item(intI).Text, 1, InStr(.Nodes.item(intI).Text, "Seat:") - 1)
                    sameSeat = True
                    Exit For
                 Else
                    sameSeat = False
                 End If
                 
            End If
        End If
            Next
End With
End Function
Private Sub clearOldSeat()

Dim intI As Integer
Dim intJ As Integer
Dim intK As Integer
Dim intL As Integer

Dim strOldSeat As String
Dim strOldRow As String
Dim strOldCol As String
Dim strTemp() As String
Dim bolFound As Boolean


bolFound = False
With tvSeat

    
        
         For intI = 1 To .Nodes.Count
         
         If .Nodes.item(intI).Selected Then
            If InStr(.Nodes.item(intI).Text, "Seat:") Then
                 strOldSeat = Mid(.Nodes.item(intI).Text, InStr(.Nodes.item(intI).Text, "Seat:") + 5)
                 If strOldSeat = "" Then Exit Sub
                 strTemp = Split(strOldSeat, " ")
                 strOldRow = strTemp(0)
                 strOldCol = strTemp(1)
                 
                 For intL = 0 To flexSeatMap.Count - 1
                    For intJ = 0 To flexSeatMap(intL).rows - 1
                           If strOldRow = flexSeatMap(intL).TextMatrix(intJ, 0) Then
                           
                           
                            For intK = 1 To flexSeatMap(intL).Cols - 1
                               If strOldCol = flexSeatMap(intL).ColHeaderCaption(0, intK) Then
                                   flexSeatMap(intL).col = intK
                                   flexSeatMap(intL).row = intJ
                                   Set flexSeatMap(intL).CellPicture = ImageList.ListImages(1).Picture
                                   
                                   'LoadPicture (App.Path & "\Icons\seat\seat_savail.gif")
                                   flexSeatMap(intL).Text = strOldCol
                                   flexSeatMap(intL).CellAlignment = flexAlignCenterCenter
                                   bolFound = True
                                   Exit For
                               End If
                            Next intK
                            If bolFound = True Then Exit For
                            
                            End If
                            
                            
                    Next intJ
                 Next intL
                 
                 If bolFound = True Then Exit For
            End If
            
          End If
   Next


End With

End Sub

Private Sub SeatPax()
Dim intI As Integer
With tvSeat

    
        
         For intI = 1 To .Nodes.Count
         
          If .Nodes.item(intI).Selected Then
            If Left(.Nodes.item(intI).Text, 1) <> "P" Then
                  .Nodes.item(intI + 1).Selected = True
            End If
          End If
         Next


End With

End Sub
Private Sub AssignSeattoTree(seatNum As String)
Dim intI As Integer
With tvSeat

    
        
         For intI = 1 To .Nodes.Count
         
          If .Nodes.item(intI).Selected Then
            If InStr(.Nodes.item(intI).Text, "Seat:") Then
                 .Nodes.item(intI).Text = Mid(.Nodes.item(intI).Text, 1, InStr(.Nodes.item(intI).Text, "Seat:") - 1)
                 
            End If
            
            If InStr(.Nodes.item(intI).Text, "Type:") Then
                 .Nodes.item(intI).Text = Mid(.Nodes.item(intI).Text, 1, InStr(.Nodes.item(intI).Text, "Type:") - 1)
                 
            End If
            
            
            If seatNum <> "" Then
            .Nodes.item(intI).Text = Trim(.Nodes.item(intI).Text) & "    Seat:" & seatNum
            End If
          End If
         Next


End With

End Sub
Private Sub flexSeatMap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'If Button = vbRightButton Then
    
        
       ' If Not IsNumeric(flexSeatMap(Index).TextMatrix(flexSeatMap(Index).row, flexSeatMap(Index).col)) Then
        
       '     PopupMenu mnuSeatProp
      ' End If
    'End If
End Sub
Private Sub flexSeatMap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strTemp As String
    'Debug.Print flexSeatMap(Index).ToolTipText
    'flexSeatMap(Index).ToolTipText = ""
new_txt = "(" & _
        Format$(flexSeatMap(Index).MouseRow) & _
        ", " & _
        Format$(flexSeatMap(Index).MouseCol) & _
        ")"
    If txt <> new_txt Then
        strTemp = seatAttribute(flexSeatMap(Index).MouseRow, flexSeatMap(Index).MouseCol)
        'Debug.Print flexSeatMap(Index).MouseCol
        'Debug.Print flexSeatMap(Index).MouseRow & " " & flexSeatMap(Index).MouseCol
        If strTemp <> "" And strTemp <> flexSeatMap(Index).ToolTipText Then
           flexSeatMap(Index).ToolTipText = strTemp
        ElseIf strTemp = "" Then
           flexSeatMap(Index).ToolTipText = ""
        End If
        
        'Debug.Print strTemp
        'flexSeatMap(Index).ToolTipText = new_txt
        'txt = new_txt
    End If

End Sub
Private Function seatAttribute(row As String, col As String) As String
    Dim intI As Integer
    Dim intJ As Integer
    Dim intK As Integer
    
    'Debug.Print row & "-" & col
    
    If cmbDeck.ListCount = 0 Or gobjSeatMaps.SeatMapCount = 0 Or gobjPNR.AirSegCount = 0 Then Exit Function
    
    For intI = 1 To gobjSeatMaps.SeatMap(cmbDeck.listindex + 1).SeatCountRow
    
    With gobjSeatMaps.SeatMap(cmbDeck.listindex + 1).SeatRow(intI)
    
        If .RowNumber = flexSeatMap(cmbDeck.listindex).TextMatrix(row, 0) Then
            For intJ = 1 To .SeatCount
                If .seat(intJ).ColumnID = flexSeatMap(cmbDeck.listindex).ColHeaderCaption(0, col) And .seat(intJ).Status <> "N" Then
            
                    For intK = 1 To .seat(intJ).SeatAttributeCount
                        seatAttribute = seatAttribute & IIf(seatAttribute = "", "", ", ") & DecodeAttributes(.seat(intJ).seatAttribute(intK).seatAttribute)
                    
                    Next intK
                    If seatAttribute <> "" Then
                        seatAttribute = .RowNumber & .seat(intJ).ColumnID & ": " & seatAttribute
                    Else
                        seatAttribute = ""
                    End If
                    Exit For
                End If
            Next intJ
            
        End If
    
    End With
    
    Next intI
    
    
End Function

Private Function DecodeAttributes(Code As String) As String

Dim strSql As String
Dim rs As ADODB.Recordset

strSql = "SELECT * FROM TBLSEATCODES WHERE CODE='" & Code & "'"

Set rs = gdbConn.Execute(strSql)

If Not rs.EOF Then

DecodeAttributes = rs!seatdesc

End If

rs.Close

End Function


Private Sub Form_Load()
   Dim oldParent As Long
    
    datFormLoadStart = Now

    gintY = 0
    gintX = 0
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
    pDisplayToFP "*SD"

    PopulateControls
    If gbolMoveProfile = True Then
      cmdPrevious.Visible = True
    Else
      cmdPrevious.Visible = False
    End If
    
   
    datFormLoadEnd = Now

  
   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
End Sub

Private Sub PopulateControls()
    getSeatCharacteristic
    getSeatRmk
    getSeatTree
    setDefaultFirstMap
    'getPax
    'getAir
  
End Sub
Private Sub loadSeatData()
Dim intI As Integer
Dim intFlexRow As Integer
Dim intFlexCol As Integer
Dim intJ As Integer
Dim strTemp As String

msFlexSeat.rows = gobjPNR.SeatDataCount + 1

For intI = 1 To gobjPNR.SeatDataCount
    msFlexSeat.row = intI
    
    msFlexSeat.col = 0
    For intJ = 0 To cmbSmoking.ListCount - 1
        
        If Left(gobjPNR.SeatData(intI).SeatAttribute1, 1) = Trim(Left(cmbSmoking.List(intJ), 1)) Then
            msFlexSeat.Text = cmbSmoking.List(intJ)
            msFlexSeat.TextMatrix(intI, 5) = msFlexSeat.Text
            Exit For
        End If
    Next intJ
    
    msFlexSeat.col = 1
    
    If Len(gobjPNR.SeatData(intI).SeatAttribute1) = 2 Then
        For intJ = 0 To cmbSeat.ListCount - 1
            If Right(gobjPNR.SeatData(intI).SeatAttribute1, 1) = Trim(Left(cmbSeat.List(intJ), 1)) Then
                msFlexSeat.Text = cmbSeat.List(intJ)
                msFlexSeat.TextMatrix(intI, 6) = msFlexSeat.Text
                Exit For
            End If
        Next intJ
    Else
        For intJ = 0 To cmbSeat.ListCount - 1
            If gobjPNR.SeatData(intI).SeatAttribute2 = Trim(Left(cmbSeat.List(intJ), 1)) Then
                msFlexSeat.Text = cmbSeat.List(intJ)
                Exit For
            End If
        Next intJ
    End If
    
    msFlexSeat.col = 2
    For intJ = 0 To cmbPax.ListCount - 1
        If gobjPNR.SeatData(intI).PaxNo = Trim(Left(cmbPax.List(intJ), InStr(cmbPax.List(intJ), ".") - 1)) Then
            msFlexSeat.Text = cmbPax.List(intJ)
            Exit For
        End If
    Next intJ
    
    msFlexSeat.col = 3
    For intJ = 1 To cmbAirSegment.ListCount - 1
        If gobjPNR.SeatData(intI).SegNum = Trim(Left(cmbAirSegment.List(intJ), InStr(cmbAirSegment.List(intJ), ".") - 1)) Then
            msFlexSeat.Text = cmbAirSegment.List(intJ)
            Exit For
        End If
    Next intJ
    
    msFlexSeat.col = 4
    If Left(gobjPNR.SeatData(intI).SeatLocation, 1) = "0" Then
    strTemp = Mid(gobjPNR.SeatData(intI).SeatLocation, 2)
    Else
    strTemp = gobjPNR.SeatData(intI).SeatLocation
    End If
    msFlexSeat.Text = IIf(strTemp = "", "Pending", strTemp) & " - " & gobjPNR.SeatData(intI).Status
          

    
Next intI

End Sub
Private Sub setDefault(intRow As Integer)
    
    gintX = 1
    msFlexSeat.col = 0
    msFlexSeat.row = intRow

        If intRow > 2 Then
            cmbSmoking.Visible = True
            cmbSmoking.SetFocus
            cmbSmoking.Left = topFrame(1).Left + msFlexSeat.Left + msFlexSeat.CellLeft
            cmbSmoking.Top = topFrame(1).Top + msFlexSeat.Top + msFlexSeat.CellTop
            cmbSmoking.Width = msFlexSeat.CellWidth
        Else
            cmbSmoking.Visible = False
        End If
        
        msFlexSeat = cmbSmoking.Text
        
        
        msFlexSeat.col = 1
        cmbSeat.Left = topFrame(1).Left + msFlexSeat.Left + msFlexSeat.CellLeft
        cmbSeat.Top = topFrame(1).Top + msFlexSeat.Top + msFlexSeat.CellTop
        cmbSeat.Width = msFlexSeat.CellWidth
        
     If intRow > 2 Then
     cmbSeat.Visible = True
    Else
    cmbSeat.Visible = False
    End If
    msFlexSeat = cmbSeat.Text
   msFlexSeat.col = 2
        If cmbPax.ListCount > 0 Then
        cmbPax.listindex = 0
        cmbPax.Left = topFrame(1).Left + msFlexSeat.Left + msFlexSeat.CellLeft
        cmbPax.Top = topFrame(1).Top + msFlexSeat.Top + msFlexSeat.CellTop
        cmbPax.Width = msFlexSeat.CellWidth
         If intRow > 2 Then
         cmbPax.Visible = True
        Else
        cmbPax.Visible = False
        End If
        msFlexSeat = cmbPax.Text
End If
   msFlexSeat.col = 3
        If cmbAirSegment.ListCount > 0 Then
        cmbAirSegment.listindex = 0
        cmbAirSegment.Left = topFrame(1).Left + msFlexSeat.Left + msFlexSeat.CellLeft
        cmbAirSegment.Top = topFrame(1).Top + msFlexSeat.Top + msFlexSeat.CellTop
        cmbAirSegment.Width = msFlexSeat.CellWidth
         If intRow > 2 Then
         cmbAirSegment.Visible = True
        Else
        cmbAirSegment.Visible = False
        End If
        msFlexSeat = cmbAirSegment.Text
        End If

End Sub
Private Sub setFlexValue(row As Integer)
     
        msFlexSeat.col = 0
        msFlexSeat.row = row
        msFlexSeat.Text = cmbSmoking.Text
        cmbSmoking.Visible = False

       
        msFlexSeat.col = 1
        msFlexSeat.Text = cmbSeat.Text
        cmbSeat.Visible = False
    
        msFlexSeat.col = 2
        msFlexSeat.Text = cmbPax.Text
        cmbPax.Visible = False

        msFlexSeat.col = 3
        msFlexSeat.Text = cmbAirSegment.Text
        cmbAirSegment.Visible = False
        
End Sub
Private Sub getSeatRmk()
'NP.G*SEAT-NA

Dim intI As Integer
Dim strTemp As String
Dim intJ As Integer


For intI = 1 To gobjPNR.GeneralRemarkCount
    If gobjPNR.GeneralRemark(intI).Qualifier = "*G" Then
        If InStr(gobjPNR.GeneralRemark(intI).RemarkText, "SEAT-") > 0 Then
            strTemp = Trim(Mid(gobjPNR.GeneralRemark(intI).RemarkText, InStr(gobjPNR.GeneralRemark(intI).RemarkText, "-") + 1))
            
            If Len(strTemp) <= 2 Then
               If Len(strTemp) > 0 Then
                    mstrSmoke = Left(strTemp, 1)
                    For intJ = 0 To cmbSmoking.ListCount
                        If Left(cmbSmoking.List(intJ), 1) = mstrSmoke Then
                            cmbSmoking.listindex = intJ
                            mstrSmoke = Trim(Mid(cmbSmoking, InStr(cmbSmoking, "-") + 1))
                            Exit For
                        End If
                    Next
               End If
               If Len(strTemp) > 1 Then
                    mstrSeat = Mid(strTemp, 2, 1)
                    For intJ = 0 To cmbSmoking.ListCount
                        If Left(cmbSeat.List(intJ), 1) = mstrSeat Then
                            cmbSeat.listindex = intJ
                            mstrSeat = Trim(Mid(cmbSeat, InStr(cmbSeat, "-") + 1))
                            Exit For
                        End If
                    Next
                
               End If
               lblSeatRmks = strTemp & "-" & mstrSmoke & " " & mstrSeat
            Else
               If lblSeatRmks.Caption = "" Then
                  lblSeatRmks = initCap(strTemp)
               Else
                  lblSeatRmks = lblSeatRmks & ", " & initCap(strTemp)
               End If
            End If
        End If
    End If
Next

End Sub
Private Sub getSeatTree()
Dim intI As Integer
Dim intJ As Integer
'Dim node As MSComctlLib.node
Dim strTemp As String
Dim blnGrey As Boolean
With tvSeat
         .Nodes.Clear
      
         For intI = 1 To gobjPNR.AirSegCount
         'blnGrey = False
             .Nodes.Add , , "S" & CStr(gobjPNR.AirSeg(intI).segnumber), gobjPNR.AirSeg(intI).TextAirSeg & " "
         '      strTemp = Trim(Mid(node.Text, InStrRev(Trim(node.Text), " ") + 1))
               
         '      If InStr(strTemp, "HK") Or _
         '           InStr(strTemp, "RR") Or _
         '           InStr(strTemp, "SS") Or _
         '           InStr(strTemp, "TK") Or _
         '           InStr(strTemp, "KK") Then
         '              node.ForeColor = "grey"
         '              blnGrey = True
         '              HK , HS, KK, KL, RR, SS, TK
                       
         '       End If
               
            For intJ = 1 To gobjPNR.PassengerCount
                .Nodes.Add "S" & CStr(gobjPNR.AirSeg(intI).segnumber), tvwChild, , "P" & gobjPNR.PassengerName(intJ).PassengerNum & ". " & gobjPNR.PassengerName(intJ).LastName & ", " & gobjPNR.PassengerName(intJ).FirstName & _
                getSeat(gobjPNR.AirSeg(intI).segnumber, gobjPNR.PassengerName(intJ).PassengerNum)
                'If blnGrey = True Then
                '    node.ForeColor = "grey"
                'End If
            Next
         Next


End With

'For Each node In tvSeat.Nodes
 
'strTemp = Trim(Mid(node.Text, InStrRev(Trim(node.Text), " ") + 1))

'If InStr(strTemp, "HK") Or _
'InStr(strTemp, "RR") Or _
'InStr(strTemp, "SS") Or _
'InStr(strTemp, "TK") Or _
'InStr(strTemp, "KK") Then
'   node.ForeColor = "grey"
   
'   If node.Children > 0 Then
'        strTemp = node.Child.Next
'   End If
'End If



 'strTemp = node.Child
 '                           For intJ = 1 To node.Children
                           
 '                               If InStr(strTemp, "Seat:") Then
 '                                   If InStr(strTemp, "Status:") > 0 Then strTemp = Mid(strTemp, 1, InStr(strTemp, "Status:") - 1)
 '                                   strSeat = Mid(strTemp, InStr(strTemp, ":") + 1)
 '                                   highLightSeat Trim(strSeat)
 '                               End If
                          
 '                              If node.Children > 1 Then strTemp = node.Child.Next

                                
'Next


End Sub
Private Function getSeat(segnumber As String, paxnumber As String) As String

Dim intI As Integer
Dim strTemp As String
Dim intJ As Integer
Dim strRow As String
Dim strCol As String
Dim strSeatChr As String
strRow = ""
strCol = ""
intJ = 1
For intI = 1 To gobjPNR.SeatDataCount

If gobjPNR.SeatData(intI).SegNum = segnumber And paxnumber = gobjPNR.SeatData(intI).PaxNo Then
    
    strTemp = gobjPNR.SeatData(intI).SeatLocation
    strSeatChr = gobjPNR.SeatData(intI).SeatAttribute1
    
    If gobjPNR.SeatData(intI).SeatAttribute2 <> "" Then
     strSeatChr = strSeatChr & gobjPNR.SeatData(intI).SeatAttribute2
    End If
    
        If strTemp <> "" Then
            While intJ <= Len(strTemp)
            
                If IsNumeric(Mid(strTemp, intJ, 1)) Then
                    
                    If Mid(strTemp, intJ, 1) = CStr(0) And intJ = 1 Then
                    Else
                        strRow = strRow & Mid(strTemp, intJ, 1)
                    End If
                Else
                    strCol = strCol & Mid(strTemp, intJ, 1)
                End If
                
                intJ = intJ + 1
                
            Wend
            If strRow <> "" And strCol <> "" Then
                getSeat = "  Seat:" & strRow & " " & strCol & "  Status:" & gobjPNR.SeatData(intI).Status & IIf(strSeatChr <> "", "  Type:", "") & Trim(strSeatChr)
                
            End If
        Else
            getSeat = "  Status: " & gobjPNR.SeatData(intI).Status & IIf(strSeatChr <> "", "  Type:", "") & Trim(strSeatChr)

        End If
    
    
End If


Next


If getSeat = "" Then
    If Len(lblSeatRmks) > 0 And InStr(lblSeatRmks, "-") > 0 Then
        getSeat = " Type:" & Trim(Mid(lblSeatRmks, 1, InStr(lblSeatRmks, "-") - 1))
    Else
        getSeat = " Type:NA"
    End If
End If

End Function

Private Sub getSeatCharacteristic()
Dim strSql As String
Dim rs As ADODB.Recordset

strSql = "Select * from tblSeatCharacteristic order by code"
Set rs = gdbConn.Execute(strSql)
cmbSeat.Clear
cmbSmoking.Clear
cmbSeat.AddItem ""
cmbSmoking.AddItem ""

While Not rs.EOF

    If Trim(rs!category) = "S" Then
        cmbSeat.AddItem Trim(rs!Code) & " - " & Trim(rs!Description)
    
        
    Else
        cmbSmoking.AddItem Trim(rs!Code) & " - " & Trim(rs!Description)
    End If

rs.MoveNext

Wend

cmbSmoking.listindex = 0
cmbSeat.listindex = 0

End Sub


Private Sub getPax()
Dim intI As Integer
Dim intPaxCount As Integer
Dim strTemp As String


intPaxCount = gobjPNR.PassengerCount
cmbPax.AddItem "All"

        If intPaxCount > 0 Then
            For intI = 1 To intPaxCount
                With gobjPNR.PassengerName(intI)
                    strTemp = intI & ". " & .LastName & "/" & .FirstName
                End With
               cmbPax.AddItem strTemp
            Next
            cmbPax.listindex = 0
        End If




End Sub
Private Sub getAir()
Dim intI As Integer
Dim strTemp As String

    lstSegments.AddItem " ALL"
    For intI = 1 To gobjPNR.AirSegCount
        lstSegments.AddItem gobjPNR.AirSeg(intI).TextAirSeg
    Next
    lstSegments.listindex = 0
'If lstAirSegment.ListCount > 0 Then lstAirSegment.Selected(0) = True

End Sub



Public Sub msFlexSeat_DblClick()
Dim strTemp As String
 If msFlexSeat.col = 1 Then
    gintX = msFlexSeat.row
    gintY = msFlexSeat.col

    cmbSeat.listindex = 1
    If msFlexSeat.Text <> "" Then cmbSeat.Text = msFlexSeat.Text
    cmbSeat.Left = topFrame(1).Left + msFlexSeat.Left + msFlexSeat.CellLeft
    cmbSeat.Top = topFrame(1).Top + msFlexSeat.Top + msFlexSeat.CellTop
    cmbSeat.Width = msFlexSeat.CellWidth
    cmbSeat.Visible = True
    cmbSeat.SetFocus
    mstrFlex = msFlexSeat.Name
    'cmbEntry.Text = msFlexEmails.Text
    'cmbEntry.Left = topFrame.Left + fraEmails.Left + msFlexEmails.Left + msFlexEmails.CellLeft
    'cmbEntry.Top = topFrame.Top + fraEmails.Top + msFlexEmails.Top + msFlexEmails.CellTop
    'cmbEntry.Width = msFlexEmails.CellWidth
    'cmbEntry.Visible = True
    'cmbEntry.SetFocus
    'mstrFlex = msFlexEmails.Name
  ElseIf msFlexSeat.col = 0 Then
        gintX = msFlexSeat.row
        gintY = msFlexSeat.col
        cmbSmoking.listindex = 0
        If msFlexSeat.Text <> "" Then cmbSmoking.Text = msFlexSeat.Text
        cmbSmoking.Left = topFrame(1).Left + msFlexSeat.Left + msFlexSeat.CellLeft
        cmbSmoking.Top = topFrame(1).Top + msFlexSeat.Top + msFlexSeat.CellTop
        cmbSmoking.Width = msFlexSeat.CellWidth
        cmbSmoking.Visible = True
        cmbSmoking.SetFocus
        mstrFlex = msFlexSeat.Name
    ElseIf msFlexSeat.col = 2 Then
        gintX = msFlexSeat.row
        gintY = msFlexSeat.col
        
        msFlexSeat.col = 4
        strTemp = msFlexSeat.Text
        If strTemp = "" Then
            msFlexSeat.col = gintY
            cmbPax.listindex = 0
            If msFlexSeat.Text <> "" Then cmbPax.Text = msFlexSeat.Text
            cmbPax.Left = topFrame(1).Left + msFlexSeat.Left + msFlexSeat.CellLeft
            cmbPax.Top = topFrame(1).Top + msFlexSeat.Top + msFlexSeat.CellTop
            cmbPax.Width = msFlexSeat.CellWidth
            cmbPax.Visible = True
            cmbPax.SetFocus
            mstrFlex = msFlexSeat.Name
        End If
    ElseIf msFlexSeat.col = 3 Then
        gintX = msFlexSeat.row
        gintY = msFlexSeat.col
        msFlexSeat.col = 4
        strTemp = msFlexSeat.Text
        
        If strTemp = "" Then
            msFlexSeat.col = gintY
            cmbAirSegment.listindex = 0
            If msFlexSeat.Text <> "" Then cmbAirSegment.Text = msFlexSeat.Text
            cmbAirSegment.Left = topFrame(1).Left + msFlexSeat.Left + msFlexSeat.CellLeft
            cmbAirSegment.Top = topFrame(1).Top + msFlexSeat.Top + msFlexSeat.CellTop
            cmbAirSegment.Width = msFlexSeat.CellWidth
            cmbAirSegment.Visible = True
            cmbAirSegment.SetFocus
            mstrFlex = msFlexSeat.Name
        End If
  End If
End Sub


Private Sub msFlexSeat_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 65 And Shift = 4 Then
       subMenuAdd_Click
    ElseIf KeyCode = 68 And Shift = 4 Then
       subMenuDelete_Click
    ElseIf KeyCode = 13 Then
       msFlexSeat_DblClick
 End If
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       gbolCancelProcess = True
    End If
End Sub

Private Sub subMenuAdd_Click()
    Dim intI As Integer
       intI = msFlexSeat.rows
       
       
       msFlexSeat.rows = msFlexSeat.rows + 1
       'setDefault msFlexSeat.Rows - 1
       setFlexValue msFlexSeat.rows - 1
End Sub
Private Sub subMenuDelete_Click()
Dim curRow As Integer
Dim strTemp As String

        'If msFlexSeat.Rows - 1 > msFlexSeat.FixedRows Then
           curRow = msFlexSeat.row
           msFlexSeat.col = 4
        
        If msFlexSeat.Text <> "" Then
            
           msFlexSeat.col = 2
                strTemp = "P" & Left(msFlexSeat.Text, InStr(msFlexSeat, ".") - 1)
           msFlexSeat.col = 3
                strTemp = strTemp & "S" & Trim(Left(msFlexSeat.Text, InStr(msFlexSeat, ".") - 1))
           
           lstDelete.AddItem strTemp
           
        End If
        
           msFlexSeat.RemoveItem (msFlexSeat.row)
           msFlexSeat.col = 1
           msFlexSeat.row = 1
        
        If cmbPax.Visible = True Then cmbPax.Visible = False
        If cmbAirSegment.Visible = True Then cmbAirSegment.Visible = False
        If cmbSeat.Visible = True Then cmbSeat.Visible = False
        If cmbSmoking.Visible = True Then cmbSmoking.Visible = False
        
        'End If
   
End Sub

Private Sub msFlexSeat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mstrFlex = msFlexSeat.Name
        PopupMenu mnuPopUpFlex
    End If
End Sub



Private Sub optQuickAssign_Click()
loadControl
End Sub

Private Sub optSeatNumber_Click()
loadControl


setDefaultFirstMap

End Sub

Private Sub loadControl()
    If optSeatNumber.value = True Then
        topFrame(0).Visible = True
        topFrame(1).Visible = False
        cmdQuickFinish.Visible = False
        cmdAssignFinish.Visible = True
        cmdQuickFinish.Enabled = False
        cmdAssignFinish.Enabled = True
        
        cmbPax.Visible = False
        cmbAirSegment.Visible = False
        cmbSeat.Visible = False
        cmbSmoking.Visible = False

        flexSeatMap(0).Visible = False
    Else
        topFrame(1).Visible = True
        topFrame(0).Visible = False
        
        cmdQuickFinish.Visible = True
        cmdAssignFinish.Visible = False
        cmdQuickFinish.Enabled = True
        cmdAssignFinish.Enabled = False
    
    End If
End Sub

Private Sub setDefaultFirstMap()
Dim node As MSComctlLib.node

Set node = tvSeat.Nodes.item(1).Root
tvSeat.Nodes.item(1).Selected = True
expandAllNode
tvSeat_NodeClick node


End Sub
Private Sub expandAllNode()
Dim node As MSComctlLib.node
    For Each node In tvSeat.Nodes
                             node.Expanded = True
    Next
End Sub
Private Sub tvSeat_NodeClick(ByVal node As MSComctlLib.node)
    
    Dim strVendor As String
    Dim strFltNum As String
    Dim strClass As String
    Dim strStartdate As String
    Dim strFrom As String
    Dim strTo As String
    Dim strKey As String
    Dim intI As Integer
    Dim strSeat As String
    Dim intJ As Integer
    Dim childnode As Integer
    Dim strTemp As String
    Dim strSeatChr As String
    Dim bolSelPax As Boolean
    Dim selPaxNode As MSComctlLib.node
    bolSelPax = False
    
GetMap:
    If InStr(node.key, "S") Then
        strCurrentSegMap = node.Text
        For intI = 0 To flexSeatMap.Count - 1
            'flexSeatMap(intI).Clear
            flexSeatMap(intI).ClearStructure
            flexSeatMap(intI).Cols = 0
            flexSeatMap(intI).rows = 0
        Next intI
        flexSeatMap(0).Visible = False
        
        blnAutoPop = True
        ClearSeatPreference
        blnAutoPop = False
        
        strKey = Replace(node.key, "S", "")
        
        If gobjPNR.AirSegCount = 0 Then
        Set gobjPNR = New CWT_GalileoPNR3.PNR
            gobjPNR.loadPNR
        End If
        
        If gobjPNR.AirSegCount = 0 Then Exit Sub
        
        For intI = 1 To gobjPNR.AirSegCount
        
            If gobjPNR.AirSeg(intI).segnumber = CInt(strKey) Then
                
                If seatNoAssignment(Trim(Mid(node.Text, InStrRev(Trim(node.Text), " ") + 1))) Then
                lblCannotAssign.Visible = False
                strVendor = gobjPNR.AirSeg(intI).Vendor
                strFltNum = gobjPNR.AirSeg(intI).FlightNumber
                strClass = gobjPNR.AirSeg(intI).Class
                strStartdate = Format(gobjPNR.AirSeg(intI).DepartDateTime, "yyyymmdd")
                strFrom = gobjPNR.AirSeg(intI).DepartAirport
                strTo = gobjPNR.AirSeg(intI).ArriveAirport
                If Not gobjSeatMaps Is Nothing Then Set gobjSeatMaps = Nothing
                Set gobjSeatMaps = New CWT_Galileo3.SeatMaps
                gobjSeatMaps.getSeatMap strVendor, strFltNum, strClass, strStartdate, strFrom, strTo
                
                If gobjSeatMaps.MapExist = True Then
                
                        DisplayMap
                        picMapIcon.Visible = True
                        lblNoMap.Visible = False
                        flexSeatMap(0).Visible = True
                        Label1.Visible = True
                        Label1.ZOrder 1
                        Label2.Visible = True
                        Label2.ZOrder 1
                        If node.children > 0 Then
                            'childNode = Node.Children
                            strTemp = node.Child
                            For intJ = 1 To node.children
                            ' For Each Node In tvSeat.Nodes
                            '    If Node.Parent.Selected = True Then
                                If InStr(strTemp, "Seat:") Then
                                    If InStr(strTemp, "Status:") > 0 Then strTemp = Mid(strTemp, 1, InStr(strTemp, "Status:") - 1)
                                    strSeat = Mid(strTemp, InStr(strTemp, ":") + 1)
                                    highLightSeat Trim(strSeat)
                                End If
                            '    End If
                               If node.children > 1 Then strTemp = node.Child.Next
                                
                            Next
                           
                           
                        End If

                    'cmbSmoking.Enabled = True
                    'cmbSeat.Enabled = True
                    'cmdSet.Enabled = False
                
                If bolSelPax = True Then
                    Set node = selPaxNode
                    GoTo ReloadPax
                End If
                Else
                    picMapIcon.Visible = False
                    flexSeatMap(0).Visible = False
                    flexSeatMap(1).Visible = False
                    lblNoMap.Visible = True
                    cmbSmoking.Enabled = True
                    cmbSeat.Enabled = True
                    cmdSet.Enabled = True
                    If cmbDeck.Visible = True Then cmbDeck.Visible = False
                    'set default chr
                    'If Len(lblSeatRmks) > 0 And InStr(lblSeatRmks, "-") > 0 Then populateSeatChrToList Mid(lblSeatRmks, 1, InStr(lblSeatRmks, "-") - 1)
                    
                End If
                
                
                
                Exit For
                
            End If
            
            Else
                    lblCannotAssign.Visible = True
                    picMapIcon.Visible = False
                    flexSeatMap(0).Visible = False
                    flexSeatMap(1).Visible = False
                    'lblNoMap.Visible = True
                    'cmbSmoking.Enabled = True
                    'cmbSeat.Enabled = True
                    'cmdSet.Enabled = True
                    If cmbDeck.Visible = True Then cmbDeck.Visible = False
                    'set default chr
                    
                    'If Len(lblSeatRmks) > 0 And InStr(lblSeatRmks, "-") > 0 Then populateSeatChrToList Mid(lblSeatRmks, 1, InStr(lblSeatRmks, "-") - 1)
                    
                
            End If
        Next
        node.Expanded = True
    Else
        If node.Parent <> strCurrentSegMap Then
            Set selPaxNode = node
            Set node = node.Parent
            
            bolSelPax = True
            GoTo GetMap
        End If

ReloadPax:
        If cmbDeck.ListCount > 0 Then
            If InStr(node.Text, "Seat:") Then
                strSeat = Mid(node.Text, InStr(node.Text, "Seat:") + 5)
                'highLightSeat strSeat
                clearfocusSeat
                focusSeat strSeat
            End If
            strSeatChr = ""
            
            If InStr(node.Text, "Type:") > 0 Then strSeatChr = Trim(Mid(node.Text, InStr(node.Text, "Type:") + 3))
            If strSeatChr <> "" Then
                blnAutoPop = True
                populateSeatChrToList strSeatChr
                blnAutoPop = False
            Else
                If Len(lblSeatRmks) > 0 And InStr(lblSeatRmks, "-") > 0 And InStr(node.Text, "Seat:") = 0 Then populateSeatChrToList Mid(lblSeatRmks, 1, InStr(lblSeatRmks, "-") - 1)
    
            End If
        Else
            strSeatChr = Trim(Mid(node.Text, InStr(node.Text, "Type:") + 3))
            If strSeatChr <> "" Then
                populateSeatChrToList strSeatChr
            Else
                If Len(lblSeatRmks) > 0 And InStr(lblSeatRmks, "-") > 0 Then populateSeatChrToList Mid(lblSeatRmks, 1, InStr(lblSeatRmks, "-") - 1)
    
            End If
            'Set node = node.Parent
            'GoTo GetMap
        End If
    End If


End Sub
Private Sub populateSeatChrToList(seatChr As String)

Dim intJ As Integer
Dim intK As Integer
Dim strTemp As String
Dim bolFound As Boolean



For intK = 1 To Len(seatChr)
 bolFound = False
    strTemp = Mid(seatChr, intK, 1)
    
  For intJ = 0 To cmbSmoking.ListCount - 1
        
        If strTemp = Trim(Left(cmbSmoking.List(intJ), 1)) Then
            cmbSmoking.listindex = intJ
            bolFound = True
            Exit For
        End If
  Next intJ
    
  If bolFound = False Then
    For intJ = 0 To cmbSeat.ListCount - 1
              If strTemp = Trim(Left(cmbSeat.List(intJ), 1)) Then
                  cmbSeat.listindex = intJ
                  Exit For
              End If
          Next intJ
  
   End If
    
Next intK

If Len(seatChr) = 1 Then
cmbSeat.listindex = 0
End If

End Sub
Private Sub clearfocusSeat()

Dim intJ As Integer
Dim intK As Integer
Dim intI As Integer
Dim bolFound As Boolean
Dim strOldSeat As String
Dim strOldRow As String
Dim strOldCol As String
Dim strSegSel As String

'bolFound = False
'For intI = 0 To flexSeatMap.Count - 1

'    For intJ = 1 To flexSeatMap(intI).Rows - 1
    
                                     
                      
'                             For intK = 1 To flexSeatMap(intI).Cols - 1
'                                    flexSeatMap(intI).HighLight = flexHighlightNever
'                                    flexSeatMap(intI).FocusRect = flexFocusNone
'                                    flexSeatMap(intI).col = intK
'                                    flexSeatMap(intI).row = intJ
'
'                                    If flexSeatMap(intI).CellBackColor = vbBlue Then
'                                        flexSeatMap(intI).CellBackColor = vbWhite
                                        
'                                        bolFound = True
'                                        Exit For
'                                    End If
                   
                                  
                                
 '                            Next intK
                             
                             
                             
 '               If bolFound = True Then Exit For
                             
 '   Next intJ
    
'    If bolFound = True Then Exit For
'Next intI
With tvSeat

    
        
For intI = 1 To .Nodes.Count
         
            If .Nodes.item(intI).Selected Then
                strSegSel = .Nodes.item(intI).Parent.Text
            
            End If
Next intI

For intI = 1 To .Nodes.Count
    If InStr(.Nodes.item(intI).key, "S") = 0 Then
         If .Nodes.item(intI).Parent = strSegSel Then
         
   
         
            If InStr(.Nodes.item(intI).Text, "Seat:") = 0 Then
                         For intL = 0 To flexSeatMap.Count - 1
                                      For intJ = 1 To flexSeatMap(intL).rows - 1
                        
                                                         
                                          
                                                 For intK = 1 To flexSeatMap(intL).Cols - 1
                                                        'flexSeatMap(intL).HighLight = flexHighlightNever
                                                        'flexSeatMap(intL).FocusRect = flexFocusNone
                                                        flexSeatMap(intL).Redraw = False
                                                        flexSeatMap(intL).col = intK
                                                        flexSeatMap(intL).row = intJ
                                                        flexSeatMap(intL).Redraw = True
                                                        If flexSeatMap(intL).CellBackColor = vbBlue Then
                                                            flexSeatMap(intL).CellBackColor = vbWhite
                                                            
                                                            bolFound = True
                                                            Exit For
                                                        End If
                                       
                                                      
                                                    
                                                 Next intK
                                        Next intJ
                         Next intL
                 Else
                
                 strOldSeat = Mid(.Nodes.item(intI).Text, InStr(.Nodes.item(intI).Text, "Seat:") + 5)

                 strTemp = Split(strOldSeat, " ")
                 strOldRow = strTemp(0)
                 strOldCol = strTemp(1)
                 
                 For intL = 0 To flexSeatMap.Count - 1
                    For intJ = 0 To flexSeatMap(intL).rows - 1
                           If strOldRow = flexSeatMap(intL).TextMatrix(intJ, 0) Then
                           
                           
                            For intK = 1 To flexSeatMap(intL).Cols - 1
                               If strOldCol = flexSeatMap(intL).ColHeaderCaption(0, intK) Then
                                   flexSeatMap(intL).col = intK
                                   flexSeatMap(intL).row = intJ
                                   If flexSeatMap(intL).CellBackColor = vbBlue Then
                                        flexSeatMap(intL).CellBackColor = vbWhite
                                        bolFound = True
                                        Exit For
                                   End If
                               End If
                            Next intK
                            If bolFound = True Then Exit For
                            
                            End If
                            
                            
                    Next intJ
                 Next intL
                 
                 If bolFound = True Then Exit For
                 End If
            End If
            
          End If
     
   Next


End With




End Sub
Private Sub focusSeat(seat As String)

Dim strTemp() As String
Dim strRow As String
Dim strCol As String
Dim intJ As Integer
Dim intK As Integer
Dim intI As Integer
Dim intTemp As Integer
Dim bolTemp As Boolean

strTemp = Split(seat, " ")
strRow = strTemp(0)
strCol = strTemp(1)
bolTemp = False

For intI = 0 To flexSeatMap.Count - 1

    For intJ = 1 To flexSeatMap(intI).rows - 1
    
                            If strRow = flexSeatMap(intI).TextMatrix(intJ, 0) Then
                            
                            
                             For intK = 1 To flexSeatMap(intI).Cols - 1
                                If strCol = flexSeatMap(intI).ColHeaderCaption(0, intK) Then
                                    flexSeatMap(intI).col = intK
                                    flexSeatMap(intI).row = intJ
                                    flexSeatMap(intI).FocusRect = flexFocusHeavy
                                    Set flexSeatMap(intI).CellPicture = ImageList.ListImages(3).Picture
                                    'LoadPicture (App.Path & "\Icons\seat\seat_sbook.gif")
                                    
                                    If bolTemp = False Then
                                        intTemp = intI
                                        bolTemp = True
                                        flexSeatMap(intI).CellBackColor = vbBlue
                                       
                   
                                    End If
                                    
                                    Exit For
                                End If
                             Next intK
                             Exit For
                             
                             End If
                             
                             
    Next intJ
    
Next intI


If bolTemp = True Then cmbDeck.listindex = intTemp
             
End Sub
Private Sub highLightSeat(seat As String)

Dim strTemp() As String
Dim strRow As String
Dim strCol As String
Dim intJ As Integer
Dim intK As Integer
Dim intI As Integer
Dim intTemp As Integer
Dim bolTemp As Boolean

strTemp = Split(seat, " ")
strRow = strTemp(0)
strCol = strTemp(1)
bolTemp = False

For intI = 0 To flexSeatMap.Count - 1

    For intJ = 1 To flexSeatMap(intI).rows - 1
    
                            If strRow = flexSeatMap(intI).TextMatrix(intJ, 0) Then
                            
                            
                             For intK = 1 To flexSeatMap(intI).Cols - 1
                                If strCol = flexSeatMap(intI).ColHeaderCaption(0, intK) Then
                                    flexSeatMap(intI).col = intK
                                    flexSeatMap(intI).row = intJ
                                    flexSeatMap(intI).FocusRect = flexFocusHeavy
                                    Set flexSeatMap(intI).CellPicture = ImageList.ListImages(3).Picture
                                    ' LoadPicture(App.Path & "\Icons\seat\seat_sbook.gif")
                                    
                                    If bolTemp = False Then
                                        intTemp = intI
                                        bolTemp = True
                                        'flexSeatMap(intI).CellBackColor = vbBlue
                                        'flexSeatMap(intI).CellPictureAlignment = 4
                                        'flexSeatMap(intI).FocusRect = flexFocusHeavy
                                        'flexSeatMap(intI).Redraw = True
                                    End If
                                    
                                    Exit For
                                End If
                             Next intK
                             Exit For
                             
                             End If
                             
                             
    Next intJ
    
Next intI

If bolTemp = True Then cmbDeck.listindex = intTemp

             
             
End Sub
Private Sub DisplayMap()
Dim intI As Integer
Dim intJ As Integer
Dim intK As Integer
Dim intL As Integer
Dim intCol As Integer
Dim blnFound As Boolean

Dim strColumn As String
Dim strTemp As String
Dim intM As Integer
Dim bolUpper As Boolean

Dim intUpperCount As Integer
Dim intLowerCount As Integer


intUpperCount = 0
intLowerCount = 0
        If gobjSeatMaps.SeatMapCount > 0 Then
        
            cmbDeck.Clear
            If gobjSeatMaps.SeatMapCount > 1 Then
                cmbDeck.Visible = True
            Else
                cmbDeck.Visible = False
            End If
            
            For intI = 1 To gobjSeatMaps.SeatMapCount
                bolUpper = False
                
                For intM = 1 To gobjSeatMaps.SeatMap(intI).SeatRow(1).RowAttributeCount
                        If gobjSeatMaps.SeatMap(intI).SeatRow(1).RowAttribute(intM).RAttribute = "U" Then
                            bolUpper = True
                            
                        End If
                Next intM
                
                If bolUpper = True Then
                    intUpperCount = intUpperCount + 1
                    cmbDeck.AddItem "Upperdeck" & IIf(intUpperCount > 1, " " & intUpperCount, "")
                'cmbDeck.AddItem gobjSeatMaps.SeatMap(intI).ColLabel
                
                Else
                    intLowerCount = intLowerCount + 1
                    cmbDeck.AddItem "Lowerdeck" & IIf(intLowerCount > 1, " " & intLowerCount, "")
                End If
                
                If gobjSeatMaps.SeatMap(intI).DisplayType = "O" Then
                    strColumn = gobjSeatMaps.SeatMap(intI).ColLabel
                    
                    
                    
                    flexSeatMap(intI - 1).Cols = Len(strColumn)
                    
                    For intJ = 1 To Len(strColumn)
                    
                        strTemp = Mid(strColumn, intJ, 1)
                        'If intJ = 1 Then flexSeatMap(intI - 1).AddItem strTemp
                        With flexSeatMap(intI - 1)
                            
                            If strTemp <> "=" Then
                            .ColHeaderCaption(0, intJ) = strTemp
                            Else
                             .ColHeaderCaption(0, intJ) = ""
                            End If
                            .ColWidth(intJ) = 500
                            
                            
                        End With
                    Next intJ
                
                End If
                
        'flexSeatMap(intI - 1).ColHeader(0) = flexColHeaderOn
        'flexSeatMap(intI - 1).GridLines = flexGridNone
        'flexSeatMap(intI - 1).Redraw = True
                flexSeatMap(intI - 1).ColHeader(0) = flexColHeaderOn
                'flexSeat.Rows = gobjSeatMaps.SeatMap(intI).SeatCountRow
                    'flexSeatMap(intI - 1).Rows = gobjSeatMaps.SeatMap(intI).SeatCountRow
                    For intK = 1 To gobjSeatMaps.SeatMap(intI).SeatCountRow
                    
                        
                        
                        If gobjSeatMaps.SeatMap(intI).DisplayType = "O" Then
                               With gobjSeatMaps.SeatMap(intI).SeatRow(intK)
                                   flexSeatMap(intI - 1).BackColorBkg = vbWhite
                                   flexSeatMap(intI - 1).rows = intK + 1
                                   flexSeatMap(intI - 1).RowHeight(intK) = 430
                                   flexSeatMap(intI - 1).CellAlignment = flexAlignCenterCenter
                                   flexSeatMap(intI - 1).CellPictureAlignment = 4
                                   flexSeatMap(intI - 1).TextMatrix(intK, 0) = .RowNumber
                                   
                                   For intCol = 1 To flexSeatMap(intI - 1).Cols - 1
                                   
                                 
                                       
                                     If flexSeatMap(intI - 1).ColHeaderCaption(0, intCol) = "" Then
                                                flexSeatMap(intI - 1).row = intK
                                                flexSeatMap(intI - 1).col = intCol
                                                flexSeatMap(intI - 1).ColWidth(intCol) = 380
                                                flexSeatMap(intI - 1).Text = gobjSeatMaps.SeatMap(intI).SeatRow(intK).RowNumber
                                                flexSeatMap(intI - 1).CellAlignment = flexAlignCenterCenter
                                                
                                                For intM = 1 To gobjSeatMaps.SeatMap(intI).SeatRow(intK).RowAttributeCount
                                                
                                                        If gobjSeatMaps.SeatMap(intI).SeatRow(intK).RowAttribute(intM).RAttribute = "ER" Then
                                                            
                                                            Set flexSeatMap(intI - 1).CellPicture = ImageList.ListImages(5).Picture
                                                            'LoadPicture (App.Path & "\Icons\seat\seat_sexit.gif")
                                                            flexSeatMap(intI - 1).CellPictureAlignment = flexAlignCenterCenter
                                                           Exit For
                                                        End If
                                                        
                                                Next intM
                                                
                                                
                                     Else
                                   
                          
                                   
                                        blnFound = False
                                            For intL = 1 To .SeatCount
                                                If flexSeatMap(intI - 1).ColHeaderCaption(0, intCol) = .seat(intL).ColumnID Then
                                                    blnFound = True
                                                   
                                                    Exit For
                                                End If
                                            Next intL
                                            
                                            
                                          If blnFound = True Then
                                                flexSeatMap(intI - 1).row = intK
                                                flexSeatMap(intI - 1).col = intCol
                                              
                                                If .seat(intL).Status = "A" Or .seat(intL).Status = "C" Then
                                                
                                                    
                                                
                                                    Set flexSeatMap(intI - 1).CellPicture = ImageList.ListImages(1).Picture
                                                    'LoadPicture (App.Path & "\Icons\seat\seat_savail.gif")
                                                    flexSeatMap(intI - 1).Text = flexSeatMap(intI - 1).ColHeaderCaption(0, intCol)
                                                    flexSeatMap(intI - 1).CellAlignment = flexAlignCenterCenter
                                                    flexSeatMap(intI - 1).CellPictureAlignment = 4
                                                    
                                                    
                                                    For intM = 1 To .seat(intL).SeatAttributeCount
                                                
                                                        If .seat(intL).seatAttribute(intM).seatAttribute = "H" Then
                                                            
                                                            Set flexSeatMap(intI - 1).CellPicture = ImageList.ListImages(4).Picture
                                                            'LoadPicture(App.Path & "\Icons\seat\seat_shandicapped.gif")
                                                            flexSeatMap(intI - 1).CellPictureAlignment = flexAlignCenterCenter
                                            
                                                            Exit For
                                                        End If
                                                        
                                                        If .seat(intL).seatAttribute(intM).seatAttribute = "I" Then
                                                            
                                                            Set flexSeatMap(intI - 1).CellPicture = ImageList.ListImages(6).Picture
                                                            'LoadPicture(App.Path & "\Icons\seat\seat_shandicapped.gif")
                                                            flexSeatMap(intI - 1).CellPictureAlignment = flexAlignCenterCenter
                                            
                                                            Exit For
                                                        End If
                                                        
                                                        
                                                        
                                                    Next intM
                                                    
                                                    
                                                    
                                                End If
                                                
                                                If .seat(intL).Status = "O" Then
                                                    Set flexSeatMap(intI - 1).CellPicture = ImageList.ListImages(2).Picture
                                                    'LoadPicture (App.Path & "\Icons\seat\seat_snotavail.gif")
                                                    flexSeatMap(intI - 1).CellPictureAlignment = 4
                                                    'flexSeatMap(intI - 1).Text = flexSeatMap(intI - 1).ColHeaderCaption(0, intCol)
                                                End If
                                                
                                                'flexSeat.Text = flexSeat.ColHeaderCaption(0, intCol)
                                                
                                          Else
                                               
                                                
                                                flexSeatMap(intI - 1).row = intK
                                                flexSeatMap(intI - 1).col = intCol
                                               
                                                Set flexSeatMap(intI - 1).CellPicture = ImageList.ListImages(2).Picture
                                                'LoadPicture(App.Path & "\Icons\seat\seat_snotavail.gif")
                                                flexSeatMap(intI - 1).CellPictureAlignment = 4
                                                'flexSeatMap(intI - 1).Text = flexSeatMap(intI - 1).ColHeaderCaption(0, intCol)
                                          End If
                                           
                                      End If
                                    
                                   Next
                                   
                                     
        
                                   
                               End With
                       End If
                    Next intK
        'flexSeatMap(intI - 1).Col = 0
        flexSeatMap(intI - 1).ColWidth(0) = 0
        
        'flexSeatMap(intI - 1).ColHeader(0) = flexColHeaderOn
        flexSeatMap(intI - 1).GridLines = flexGridNone
        flexSeatMap(intI - 1).Redraw = True
                
            Next intI
        Else
            
        End If

        
        If cmbDeck.ListCount > 0 Then cmbDeck.listindex = 0
End Sub

Private Sub changeCelltext(ByRef Form As Form, ByRef strFunction As String, ByRef msFlex As MSFlexGrid, strStyle As String, bolSkipFirstCol As Boolean)
    
    If strStyle = "H" Then  'Horizontal
        If msFlex.col < msFlex.Cols - 1 Then
           msFlex.col = msFlex.col + 1
           CallByName Form, strFunction, VbMethod
        ElseIf msFlex.col = msFlex.Cols - 1 Then
           If msFlex.row < msFlex.rows - 1 Then
              msFlex.col = 0
              If msFlex.col < msFlex.FixedCols Or bolSkipFirstCol Then msFlex.col = msFlex.col + 1
              msFlex.row = msFlex.row + 1
              CallByName Form, strFunction, VbMethod
           End If
        End If
    End If
End Sub


Private Function seatNoAssignment(strTemp As String) As Boolean
seatNoAssignment = False
               If InStr(strTemp, "HK") Or _
                    InStr(strTemp, "HS") Or _
                    InStr(strTemp, "RR") Or _
                    InStr(strTemp, "SS") Or _
                    InStr(strTemp, "TK") Or _
                    InStr(strTemp, "KL") Or _
                    InStr(strTemp, "KK") Then
                       
                        seatNoAssignment = True
                
                       
                End If
End Function

