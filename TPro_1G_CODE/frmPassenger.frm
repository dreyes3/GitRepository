VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmPassenger 
   BackColor       =   &H00FAF6EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CWT Desktop - Passenger"
   ClientHeight    =   3420
   ClientLeft      =   1590
   ClientTop       =   3075
   ClientWidth     =   11370
   Icon            =   "frmPassenger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPassenger.frx":038A
   ScaleHeight     =   3420
   ScaleWidth      =   11370
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   3900
      Left            =   0
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   6879
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
      Begin VB.CheckBox chkEntry 
         Height          =   200
         Left            =   4440
         TabIndex        =   10
         Top             =   3120
         Visible         =   0   'False
         Width           =   200
      End
      Begin MyFramePanel.MyFrame topFrame 
         Height          =   2865
         Left            =   80
         Top             =   120
         Width           =   11200
         _ExtentX        =   19764
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
         Begin VB.OptionButton optHBT 
            BackColor       =   &H00DADAB6&
            Caption         =   "Skip Hotel"
            Height          =   255
            Index           =   1
            Left            =   9360
            TabIndex        =   12
            Top             =   2600
            Width           =   1455
         End
         Begin VB.OptionButton optHBT 
            BackColor       =   &H00DADAB6&
            Caption         =   "Book Hotel"
            Height          =   255
            Index           =   0
            Left            =   9360
            TabIndex        =   11
            Top             =   2340
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Frame fraPassenger 
            BackColor       =   &H00DADAB6&
            Caption         =   " Passengers "
            ForeColor       =   &H00000000&
            Height          =   2175
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   10900
            Begin MSFlexGridLib.MSFlexGrid msFlexPassenger 
               Height          =   1815
               Left            =   120
               TabIndex        =   1
               Top             =   240
               Width           =   10700
               _ExtentX        =   18865
               _ExtentY        =   3201
               _Version        =   393216
               Cols            =   6
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   16777215
               BackColorFixed  =   6973442
               ForeColorFixed  =   16777215
               BackColorSel    =   -2147483643
               BackColorBkg    =   14342838
               HighLight       =   0
               AllowUserResizing=   3
               BorderStyle     =   0
               Appearance      =   0
            End
         End
         Begin MSComCtl2.DTPicker dtpTktDate 
            Height          =   315
            Left            =   1320
            TabIndex        =   2
            Top             =   2400
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            OLEDropMode     =   1
            CalendarTitleBackColor=   -2147483647
            CustomFormat    =   "ddMMM"
            Format          =   17039363
            CurrentDate     =   38285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ticketing Date:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   2450
            Width           =   1095
         End
      End
      Begin VB.PictureBox cmbContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1800
         ScaleHeight     =   375
         ScaleWidth      =   855
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
         Begin MSForms.ComboBox cmbEntry 
            Height          =   375
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   855
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1508;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   3120
         Visible         =   0   'False
         Width           =   1545
      End
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
         LcK2            =   $"frmPassenger.frx":0714
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
      Begin MSComCtl2.DTPicker dtpEntry 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16751157
         CustomFormat    =   "ddMMMyy"
         Format          =   17039363
         CurrentDate     =   36161
         MaxDate         =   109574
         MinDate         =   21916
      End
      Begin MyCommandButton.MyButton cmdFinish 
         Height          =   360
         Left            =   9120
         TabIndex        =   3
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
         AppearanceThemes=   1
         BackColorDown   =   3968251
         BackColorOver   =   6805503
         BackColorFocus  =   16765357
         BackColorDisabled=   16765357
         TransparentColor=   16447215
         Caption         =   "&Finish"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdCancel 
         Height          =   360
         Left            =   10200
         TabIndex        =   4
         Top             =   3105
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
         TransparentColor=   16447215
         Caption         =   "&Cancel"
         Depth           =   1
         GradientType    =   2
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
Attribute VB_Name = "frmPassenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mbolClickBelowRow As Boolean
Dim mstrFlex As String
Dim datFormLoadEnd As Date
Dim datFormLoadStart As Date
Dim datTouchEnd As Date

Private Sub chkEntry_Click()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexPassenger.Name Then
       Set msFlex = msFlexPassenger
    End If
    checkedRow msFlex
    
End Sub

Private Sub chkEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexPassenger.Name Then
       Set msFlex = msFlexPassenger
       'Clement - 20080812
       'msFlexPassenger.SetFocus
    End If
    control_LostFocus msFlex, Me, chkEntry
End Sub


Private Sub cmbEntry_GotFocus()
    cmbGetFocus cmbEntry
End Sub

Private Sub cmbEntry_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    control_KeyDown CInt(KeyCode), Shift, Me, cmbEntry.Container
End Sub

Private Sub cmbEntry_LostFocus()
    'Clement - 20080812
    'msFlexPassenger.SetFocus
    control_LostFocus msFlexPassenger, Me, cmbEntry
End Sub

Private Sub cmbEntry_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = fAllowAlphaNumeric(CInt(KeyAscii), " -")
End Sub

Private Sub cmdCancel_Click()
    gbolCancelProcess = True
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    
    'JY - V1.2.9 20120104 - CR117 - EM Prompt for Restricted Countries and Airlines
    Dim strMsg As String
    Dim intAns As Integer
    
    If validData Then
       datTouchEnd = Now
       writeDatatoGDS
       'If gbolWritingtoPNR = False Then Exit Sub
       Set gobjPNR = New CWT_GalileoPNR3.PNR
       gobjPNR.loadPNR
       'displayPNRinBar
       
       
       'Log formload
       
       'Back up on 26 Sept - Jeremy
'       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
'       gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModTAU), _
'        IIf(gbolCreatPNR = True, gconSModTAU, ""), Me.Name, gconFormLoad, gstrProcessGrpID, _
'       datFormLoadEnd, datFormLoadStart
'
'       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
'         gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModTAU), _
'        IIf(gbolCreatPNR = True, gconSModTAU, ""), Me.Name, gconTouch, gstrProcessGrpID, _
'       datTouchEnd, datFormLoadEnd
'
'       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
'       gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModTAU), _
'       IIf(gbolCreatPNR = True, gconSModTAU, ""), Me.Name, gconProcessing, gstrProcessGrpID, _
'        , datTouchEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModTAU, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModTAU, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModTAU, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
       
       If gbolCreatPNR Then
        
            UpdatePNR gstrProcessGrpID, gobjPNR.RecLoc
            
        
       End If
              
       Unload Me
        
       'JY - V1.2.9 20120104 - CR117 - EM Prompt for Restricted Countries and Airlines
       If gbolCreatPNR = False Then
            If gobjPNR.CompInfo.AquaItin = True Then
            
               strMsg = checkRestrictedRules
               If strMsg <> "" Then
             
                  modMsgBox.YESMsg = "Yes"
                  modMsgBox.NOMsg = "No"
                  intAns = modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbExclamation + vbApplicationModal + vbYesNo, "CWT Desktop")
                  If intAns = vbYes Then
                     Load frmAquaItinRmk
                     frmAquaItinRmk.Show
                     Do
                       DoEvents
                     Loop Until isLoaded("frmAquaItinRmk") = False
                  End If
               End If
               
             End If
        End If
    End If
End Sub

Private Sub dtpEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, dtpEntry
End Sub

Private Sub dtpEntry_LostFocus()
    Dim preX As Integer
    Dim preY As Integer
    
    preX = gintX
    preY = gintY
    'Clement - 20080812
    'msFlexPassenger.SetFocus
    control_LostFocus msFlexPassenger, Me, dtpEntry, , False
    With msFlexPassenger
         .TextMatrix(preX, preY) = Format(.TextMatrix(preX, preY), dtpEntry.CustomFormat)
    End With
End Sub

Private Sub dtpTktDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdFinish_Click
    End If
End Sub

Private Sub Form_Load()
   Dim oldParent As Long
   Dim strTemp As String
   Dim i As Integer
   Dim bolHBUUser As Boolean
   
   datFormLoadStart = Now

    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0
   
   pDisplayToFP "*VR"
   'Set the header caption and width of msFlexPassenger
   With msFlexPassenger
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
                .ColWidth(i) = 250
            ElseIf i = 1 Then
               .ColWidth(i) = 0
            ElseIf i = 2 Then
               .Text = " Last Name (Surname) *"
               .ColWidth(i) = 3000
            ElseIf i = 3 Then
               .Text = " First Name (Given Name) & Title *"
               .ColWidth(i) = 3000
            ElseIf i = 4 Then
               .Text = " Type *"
               .ColWidth(i) = 1200
            ElseIf i = 5 Then
               .Text = " Name Remarks - Optional for Adult"
               .ColWidth(i) = 2900
            End If
            .ColAlignment(i) = 1
        Next
        setText msFlexPassenger, 0, 0, 0
        .row = 1
        .col = 1
   End With
   
   populatePsgr
   
   If gobjPNR.TktDate = "" Or gobjPNR.RecLoc = "" Then
      dtpTktDate.value = Date
   Else
      strTemp = Mid(gobjPNR.TktDate, 1, 2) & "/" & Mid(gobjPNR.TktDate, 3) & "/" & Format(Now, "YY")
      dtpTktDate.value = CDate(strTemp)
      If dtpTktDate.value < Now Then
         dtpTktDate.value = DateAdd("yyyy", 1, dtpTktDate.value)
      End If
   End If
   
   'Find Ticketing deadline in RI
   For i = 1 To gobjPNR.ItinRemarkCount
        With gobjPNR.ItinRemark(i)
            If InStr(1, .RemarkText, "PLEASE ISSUE TICKET BY") > 0 Then
               dtpTktDate.Tag = i
               Exit For
            End If
        End With
   Next
   
   
   datFormLoadEnd = Now
   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
   
   'CC - 20110816 - HBU
   bolHBUUser = ApplicationUser(gobjHost.AgentPCC, gobjHost.AgentSine, "HBU")
   
   'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
    If gbolCreatPNR = True And bolHBUUser = True Then
       'Show the hotel booking buttons if this is under the Create PNR process
       optHBT(0).Visible = True
       optHBT(1).Visible = True
       optHBT(0).value = True
       optHBT(1).value = False
    Else
       'Do not show the hotel booking buttons if the consultant calls this module manually
       optHBT(0).Visible = False
       optHBT(1).Visible = False
       optHBT(0).value = False
       optHBT(1).value = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       gbolCancelProcess = True
    End If
End Sub

Public Sub msFlexPassenger_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
    
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
    
    intTop = topFrame.Top + fraPassenger.Top + msFlexPassenger.Top + msFlexPassenger.CellTop
    intLeft = topFrame.Left + fraPassenger.Left + msFlexPassenger.Left + msFlexPassenger.CellLeft
    
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    If chkEntry.Visible = True Then chkEntry.Visible = False
    
    mstrFlex = msFlexPassenger.Name
    
    If msFlexPassenger.col = 0 And msFlexPassenger.row = 0 Then
       optSelectAll msFlexPassenger
    ElseIf msFlexPassenger.col = 0 And msFlexPassenger.row > 0 Then
       If msFlexPassenger.TextMatrix(msFlexPassenger.row, msFlexPassenger.col) <> "" Then
          setControlPosition msFlexPassenger, chkEntry, intTop, intLeft
          checkedRow msFlexPassenger
       End If
    ElseIf msFlexPassenger.TextMatrix(msFlexPassenger.row, 1) = "" Or gobjPNR.RecLoc = "" Or _
           (msFlexPassenger.TextMatrix(msFlexPassenger.row, 1) <> "" And msFlexPassenger.col = 5 And msFlexPassenger.row > 0) Then
        If (msFlexPassenger.col = 2 Or msFlexPassenger.col = 3) And msFlexPassenger.row > 0 Then
           setControlPosition msFlexPassenger, txtEntry, intTop, intLeft
        ElseIf msFlexPassenger.col = 4 And msFlexPassenger.row > 0 Then
          cmbEntry.Clear
          If gobjPNR.RecLoc = "" Then
             cmbEntry.AddItem "ADULT"
          End If
          cmbEntry.AddItem "INFANT"
          setControlPosition msFlexPassenger, cmbEntry.Container, intTop, intLeft, cmbEntry
        ElseIf msFlexPassenger.col = 5 And msFlexPassenger.row > 0 Then
          If UCase(Trim(msFlexPassenger.TextMatrix(msFlexPassenger.row, 4))) = "ADULT" Then
             setControlPosition msFlexPassenger, txtEntry, intTop, intLeft
          ElseIf UCase(Trim(msFlexPassenger.TextMatrix(msFlexPassenger.row, 4))) = "INFANT" Then
             'Must input birth date in DDMMMYY for infant
              If msFlexPassenger.Text <> "" Then msFlexPassenger.Text = CDate(Mid(msFlexPassenger.Text, 1, 2) & "-" & Mid(msFlexPassenger.Text, 3, 3) & "-" & Mid(msFlexPassenger.Text, 6, 2))
              setControlPosition msFlexPassenger, dtpEntry, intTop, intLeft
          End If
        End If
    End If
End Sub

Private Sub msFlexPassenger_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, msFlexPassenger
End Sub
Private Sub msFlexPassenger_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseDown msFlexPassenger, Button, Y
End Sub

Public Sub subMenuAdd_Click()
    
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    If chkEntry.value = True Then chkEntry.Visible = False
    
    If mstrFlex = msFlexPassenger.Name Then
       msFlexPassenger.rows = msFlexPassenger.rows + 1
       setText msFlexPassenger, msFlexPassenger.rows - 1, 0, 0
    End If
End Sub

Public Sub subMenuDelete_Click()
    Dim i As Integer
    Dim msFlex As MSFlexGrid
    Dim strTemp As String
    Dim bolAddNewRow As Boolean
    Dim preRow As Integer
    
    bolAddNewRow = False
    
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    If chkEntry.value = True Then chkEntry.Visible = False
    
    If mstrFlex = msFlexPassenger.Name Then
       Set msFlex = msFlexPassenger
    End If
    
    With msFlex
         'Checked how many rows selected
         i = rowSelected(msFlex)
         If i = 0 Then
            If .rows = 2 Then bolAddNewRow = True
         ElseIf i = .rows - 1 Then
            bolAddNewRow = True
         End If
         If bolAddNewRow = True Then
            'Must add new row since fixedRow need at least 1 row
            preRow = .row
            subMenuAdd_Click
            .row = preRow
         End If
         If i = 0 Then
            If msFlex.TextMatrix(.row, 0) <> "" Then
               deleteRow msFlex, .row
            End If
         Else
             For i = 1 To .rows - 1
                 If i <= .rows - 1 Then
                    If .TextMatrix(i, 0) = gstrChecked Then
                       deleteRow msFlex, i
                       i = i - 1
                    End If
                 End If
             Next
         End If
    End With
End Sub

Private Sub txtEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    control_KeyDown KeyCode, Shift, Me, txtEntry
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
   If gintY = 1 Or gintY = 2 Then
      KeyAscii = fAllowAlpha(CInt(KeyAscii), " -")
   Else
      KeyAscii = fAllowAlphaNumeric(CInt(KeyAscii), " -")
   End If
End Sub

Private Sub txtEntry_LostFocus()
    'Clement - 20080812
    'msFlexPassenger.SetFocus
    control_LostFocus msFlexPassenger, Me, txtEntry, , False
End Sub

Private Sub populatePsgr()
    Dim i As Integer
    
    'Populate passengers from PNR
    For i = 1 To gobjPNR.PassengerCount
        If i > msFlexPassenger.rows - 1 Then
           msFlexPassenger.rows = msFlexPassenger.rows + 1
        End If
        msFlexPassenger.TextMatrix(i, 1) = gobjPNR.PassengerName(i).GDSNum
        msFlexPassenger.TextMatrix(i, 2) = gobjPNR.PassengerName(i).LastName
        msFlexPassenger.TextMatrix(i, 3) = gobjPNR.PassengerName(i).FirstName
        msFlexPassenger.TextMatrix(i, 4) = IIf(gobjPNR.PassengerName(i).PassengerType = "", "ADULT", "INFANT")
        msFlexPassenger.TextMatrix(i, 5) = gobjPNR.PassengerName(i).Remark
    Next
End Sub

Private Sub writeDatatoGDS()
    
    Dim strResponse As String
    Dim strCmd As String
    Dim strTemp As String
    Dim bolDeleteAll As Boolean
    Dim strCmp As String
    Dim strQueue As String
    Dim strMsg As String
    Dim i As Integer
    
    If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
    gbolWritingtoPNR = True
    
    
    'Preethi - V1.2.2 20110303 - IR7 - To append canned remarks for newly created PNR only
    If gobjPNR.RecLoc = "" Then
       AddCRFromTbl strCmd
       gobjHost.terminalEntry (strCmd)
       strCmd = ""
    End If
    With msFlexPassenger
        For i = 1 To msFlexPassenger.rows - 1
            If (.TextMatrix(i, 1) = "" And gobjPNR.RecLoc <> "") Or gobjPNR.RecLoc = "" Then
                strTemp = Trim(.TextMatrix(i, 2)) & "/" & Trim(.TextMatrix(i, 3)) & _
                          IIf(Trim(.TextMatrix(i, 5)) = "", "", "*" & IIf(Trim(.TextMatrix(i, 4)) = "INFANT", Format(Trim(.TextMatrix(i, 5)), "ddMMMyy"), Trim(.TextMatrix(i, 5))))
               
                If bolDeleteAll = True Or i > gobjPNR.PassengerCount Then
                    strCmd = strCmd & IIf(strCmd = "", "", "+") & "N." & IIf(Trim(.TextMatrix(i, 4)) = "INFANT", "I/", "") & UCase(strTemp)
                Else
                    If IIf(Trim(.TextMatrix(i, 4)) = "INFANT", "I", "") <> gobjPNR.PassengerName(i).PassengerType Then
                       'If passenger type is different, must delete all subsequent passengers first
                       bolDeleteAll = True
                       strResponse = gobjHost.terminalEntry("N.P" & IIf(i = gobjPNR.PassengerCount, i, i & "-" & gobjPNR.PassengerCount) & "@")
                       If strResponse <> "*" Then GoTo errorWriting
                       strCmd = strCmd & IIf(strCmd = "", "", "+") & "N." & IIf(Trim(.TextMatrix(i, 4)) = "INFANT", "I/", "") & UCase(strTemp)
                    Else
                       strCmp = gobjPNR.PassengerName(i).LastName & "/" & gobjPNR.PassengerName(i).FirstName & _
                                IIf(gobjPNR.PassengerName(i).Remark = "", "", "*" & gobjPNR.PassengerName(i).Remark)
                       If UCase(strTemp) <> UCase(Trim(strCmp)) Then
                          strCmd = strCmd & IIf(strCmd = "", "", "+") & "N.P" & i & "@" & IIf(Trim(.TextMatrix(i, 4)) = "INFANT", "I/", "") & UCase(strTemp)
                       End If
                    End If
                End If
            ElseIf .TextMatrix(i, 1) <> "" Then
                strTemp = Trim(.TextMatrix(i, 5))
                strCmp = gobjPNR.PassengerName(i).Remark
                If UCase(strTemp) <> UCase(strCmp) Then
                   strCmd = strCmd & IIf(strCmd = "", "", "+") & "N.P" & i & "@*" & UCase(strTemp)
                End If
            End If
        Next
        
        If bolDeleteAll = False Then
           i = msFlexPassenger.rows - 1
           If i < gobjPNR.PassengerCount Then
              strResponse = gobjHost.terminalEntry("N.P" & IIf((gobjPNR.PassengerCount - i) = 1, i + 1, (i + 1) & "-" & gobjPNR.PassengerCount) & "@")
              If strResponse <> "*" Then GoTo errorWriting
           End If
        End If
    End With
    strTemp = ""
    
    If gobjPNR.RecLoc = "" Then
        'Run following steps if this is a new booking
        'Change CN number
        For i = 1 To gobjPNR.AcctRemarkCount
           With gobjPNR.AcctRemark(i)
                If .RemarkType = "FT" And Mid(.RemarkText, 1, 3) = "CN/" Then
                   If Trim(Mid(.RemarkText, 4)) <> frmSideBar.cmbBar.List(frmSideBar.cmbBar.listindex, 1) Then
                      strCmd = strCmd & IIf(strCmd = "", "", "+") & "DI." & .ItemNum & "@FT-CN/" & Trim(frmSideBar.cmbBar.List(frmSideBar.cmbBar.listindex, 1))
                   End If
                   Exit For
                End If
           End With
        Next i
                
        strCmd = strCmd & IIf(strCmd = "", "", "+") & "T.TAU/" & Format(dtpTktDate.value, "ddMMM")
        strCmd = strCmd & IIf(strCmd = "", "", "+") & "NP.S*APMOV SCRIPT COMPLETED+NP.SS*VBIPM+DI.FT-FF37/" & gTrxnType
            
        If gstrAgcyCountryCode = "SG" Then
           strResponse = gobjHost.terminalEntry("IMUCR1-12")
           If InStr(1, strResponse, "IMU UPDATED") = 0 Then GoTo errorWriting
           If Trim(gobjHost.AgentPCC) <> "" And Trim(gobjHost.AgentQueue) <> "" Then
              strCmd = strCmd & IIf(strCmd = "", "", "+") & "RB." & gobjHost.AgentPCC & "/" & Format(Date, "DDMMM") & "/Q" & gobjHost.AgentQueue & ""
           End If
           If Trim(gobjHost.AgentGACode) <> "" Then
              strCmd = strCmd & IIf(strCmd = "", "", "+") & "DI.FT-BA/" & gobjHost.AgentGACode
           End If
        End If
        'Preethi - V1.2.2 20110223 - CR36 - Removal Of TAU Date In RI Remarks
        strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI.PRONAME." & Trim(frmSideBar.cmbBar.Text)
        '& IIf(gobjPNR.AirSegCount > 0, "+RI./0*PLEASE ISSUE TICKET BY " & Format(dtpTktDate.value, "ddMMMyy"), "")
        
    Else
        If gobjPNR.TktDate <> "" And gobjPNR.TktDate <> Format(dtpTktDate.value, "ddMMM") Then
           strCmd = strCmd & IIf(strCmd = "", "", "+") & "T.@"
           strCmd = strCmd & IIf(strCmd = "", "", "+") & "T.TAU/" & Format(dtpTktDate.value, "ddMMM")
           'Preethi - V1.2.2 20110223 - CR36 - Removal Of TAU date In RI Remarks
           'If dtpTktDate.Tag <> "" And gobjPNR.AirSegCount > 0 Then
              'strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & dtpTktDate.Tag & "@PLEASE ISSUE TICKET BY " & Format(dtpTktDate.value, "ddMMMyy")
           'ElseIf dtpTktDate.Tag = "" And gobjPNR.AirSegCount > 0 Then
              'strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & IIf(gobjPNR.ItinRemarkCount > 0, "/0*", "") & "PLEASE ISSUE TICKET BY " & Format(dtpTktDate.value, "ddMMMyy")
           'End If
        ElseIf gobjPNR.TktDate = "" Then
           strCmd = strCmd & IIf(strCmd = "", "", "+") & "T.TAU/" & Format(dtpTktDate.value, "ddMMM")
           'Preethi - V1.2.2 20110223 - CR36 - Removal Of TAU date In RI Remarks
           'If dtpTktDate.Tag <> "" And gobjPNR.AirSegCount > 0 Then
              'strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & dtpTktDate.Tag & "@PLEASE ISSUE TICKET BY " & Format(dtpTktDate.value, "ddMMMyy")
          ' ElseIf dtpTktDate.Tag = "" And gobjPNR.AirSegCount > 0 Then
              'strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & IIf(gobjPNR.ItinRemarkCount > 0, "/0*", "") & "PLEASE ISSUE TICKET BY " & Format(dtpTktDate.value, "ddMMMyy")
           'End If
        End If
        'Preethi - V1.2.2 20110223 - CR36 - Removal Of TAU Date In RI Remarks
        If dtpTktDate.Tag <> "" Then
           strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & dtpTktDate.Tag & "@"
        End If
    End If
    
    
    'send entries, received & end the PNR
    strCmd = strCmd & IIf(strCmd = "", "", "+") & "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine) & "+ER+ER+ER"
    strResponse = gobjHost.terminalEntry(strCmd)
    strTemp = strResponse
    'Preethi - V1.2.4 20110712 - CR 76 - Change Validation Logic For ENDPNR
    'If InStr(strTemp, "1.1") = 0 Then
    If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = False Then
       For i = 0 To 1
           strTemp = gobjHost.terminalEntry("ER")
           'If InStr(strTemp, "1.1") > 0 Then
           If CheckResponse(strTemp, gstrPNRExpression, gintCheckERLineNum) = True Then
               Exit For
           End If
           If i = 1 Then GoTo errorWriting
       Next
    End If


'    strTemp = gobjHost.EndPNR2(IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine), True, 2)
'    If strTemp <> "True" Then GoTo errorWriting
    
    'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
    If optHBT(0).value = True Then
       gbolStartHBT = True
    End If
    
    Exit Sub
    
errorWriting:
    'Prompt error message if failed to write to PNR
    gbolWritingtoPNR = False
    strMsg = "Unable to end PNR. Response from GDS is " & Chr(13) & strResponse
    strMsg = strMsg & Chr(13) & "System will continue without ending this booking."
    
    'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
    If optHBT(0).value = True Then
       ' Do not invoke the HBT because PNR locator and PNR Profile are the mandatory fields
       strMsg = strMsg & Chr(13) & "Desktop will proceed without invoking HBT."
    End If
    
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
    
    'gobjHost.TerminalEntry "IR"
    
End Sub
Private Sub txtTAURemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdFinish_Click
    Else
       KeyAscii = fAllowAlphaNumeric(KeyAscii, " ")
    End If
End Sub

Private Sub AddCRFromTbl(ByRef strCmd As String)
    Dim strSQL As String
    Dim rsCR As New ADODB.Recordset
    Dim strProName As String
 
    gobjLog.LineTextToLog "Add Cann Remarks"
    strProName = UCase(frmSideBar.cmbBar.Text)
    strSQL = "Select CRNum,CRText from tblClientCR CR,tblClients C where CR.ClientID=C.ClientID and C.ProName= '" & strProName & "' and CR.RI=1 order by CRNum"

    Set rsCR = New ADODB.Recordset
    rsCR.Open strSQL, gdbConn, adOpenKeyset, adLockReadOnly

    gobjLog.LineTextToLog "EOF=" & rsCR.EOF

    If Not rsCR.EOF Then
       While Not rsCR.EOF
            If Trim(rsCR![crText]) <> "" Then
               strCmd = strCmd & IIf(strCmd = "", "", "+") & "RI." & Trim(rsCR![crText])
            End If
            rsCR.MoveNext
        Wend
    End If

    rsCR.Close
    Set rsCR = Nothing
    
End Sub

Private Function validData() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim strMsg As String
    Dim bolMultipleName As Boolean
    Dim strTemp As String
    Dim strCmp As String
    Dim rsRecord As ADODB.Recordset
    Dim strSQL As String
        
    bolMultipleName = False
    validData = True
    
    With msFlexPassenger
        For i = 1 To .rows - 1
            If i = 1 Then
               '1st Passenger must be an adult
               If Trim(.TextMatrix(i, 4)) <> "ADULT" Then
                  strMsg = "First passenger must be an adult ..." & Chr(13)
                  Exit For
               End If
            End If

            If Trim(.TextMatrix(i, 2)) = "" Then
               strMsg = strMsg & "Missing last name for passenger " & i & Chr(13)
            End If
            If Trim(.TextMatrix(i, 3)) = "" Then
               strMsg = strMsg & "Missing first name for passenger " & i & Chr(13)
            End If
            If Trim(.TextMatrix(i, 4)) <> "ADULT" And Trim(.TextMatrix(i, 4)) <> "INFANT" Then
               strMsg = strMsg & "Missing passenger type for passenger " & i & Chr(13)
            End If
            If Trim(.TextMatrix(i, 4)) = "INFANT" And Trim(.TextMatrix(i, 5)) = "" Then
               strMsg = strMsg & "Missing infant birth date in name remarks for passenger " & i & Chr(13)
            End If

            'Detect same names
            If bolMultipleName = False Then
                If Trim(.TextMatrix(i, 2)) <> "" And Trim(.TextMatrix(i, 3)) <> "" Then
                    strTemp = Trim(.TextMatrix(i, 2)) & Trim(.TextMatrix(i, 3))
                    For j = i + 1 To .rows - 1
                        strCmp = Trim(.TextMatrix(j, 2)) & Trim(.TextMatrix(j, 3))
                        If UCase(strTemp) = UCase(strCmp) Then
                           strMsg = strMsg & "Multiple names - " & strTemp & Chr(13)
                           bolMultipleName = True
                        End If
                    Next
                End If
            End If
        Next
        'Velidate date
        If dtpTktDate.value < Date Then
           strMsg = strMsg & "Invalid ticketing date " & Chr(13)
        Else
           'Check Sat/Sunday
           If dtpTktDate.DayOfWeek = 7 Or dtpTktDate.DayOfWeek = 1 Then
              strMsg = strMsg & "Ticketing date falls on weekend " & Chr(13)
           Else
              'Check Public Holidays
              strSQL = "Select * from tblCWTHolidays where HolidayDate = '" & dtpTktDate.value & "' AND CountryCode='" & gstrAgcyCountryCode & "'"
              Set rsRecord = gdbEitinConn.Execute(strSQL)
              If rsRecord.EOF = False Then
                 strMsg = strMsg & "Ticketing date falls on public holiday " & Chr(13)
              End If
              rsRecord.Close
              Set rsRecord = Nothing
           End If
        End If
    End With
            
    If strMsg <> "" Then
       validData = False
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    End If
End Function

Private Sub checkedRow(ByRef msFlex As MSFlexGrid)
    If chkEntry.value = vbChecked And gintY = 0 Then
       HighlightRow msFlex, gintX
    ElseIf chkEntry.value = vbUnchecked And gintY = 0 Then
       HighlightRow msFlex, gintX, False
    End If
End Sub

Private Sub mouseDown(ByRef msFlex As MSFlexGrid, ByVal Button As Integer, ByVal Y As Single)
         
    With msFlex
         If txtEntry.Visible = True Then txtEntry.Visible = False
         If cmbContainer.Visible = True Then cmbContainer.Visible = False
         If dtpEntry.Visible = True Then dtpEntry.Visible = False
         If chkEntry.Visible = True Then chkEntry.Visible = False
         
         mstrFlex = .Name
         .row = .MouseRow
         .col = .MouseCol
        
         If Button = vbRightButton Then
             PopupMenu mnuPopUpFlex
         Else
             If Y > .RowPos(.rows - 1) + .RowHeight(.rows - 1) Then
                mbolClickBelowRow = True
             Else
                mbolClickBelowRow = False
             End If
         End If
    End With
End Sub

Private Sub deleteRow(ByRef msFlex As MSFlexGrid, ByVal i As Integer)
           
   Dim strTemp As String
       
   With msFlex
       msFlex.RemoveItem (i)
       msFlex.col = 1
       If i <= .rows - 1 Then
          msFlex.row = i
       Else
          msFlex.row = i - 1
       End If
   End With

End Sub



