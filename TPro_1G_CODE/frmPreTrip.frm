VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.2#0"; "MyCommandButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4832871B-0993-461C-B983-0EAAA4A43E5C}#5.0#0"; "SftTabs_IX86_U_50.ocx"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{94F7E282-F78A-11D1-9587-0000B43369D3}#1.1#0"; "argradient1.ocx"
Begin VB.Form frmPreTrip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CWT Desktop - PreTrip Reporting"
   ClientHeight    =   3420
   ClientLeft      =   1680
   ClientTop       =   4710
   ClientWidth     =   11355
   Icon            =   "frmPreTrip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   11355
   Begin ARGradientControl.ARGradient ARGradient1 
      Height          =   3795
      Left            =   0
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   6694
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
      Begin SftTabsLib.SftTabs SftTabs 
         Height          =   3000
         Left            =   0
         TabIndex        =   5
         Top             =   60
         Width           =   11145
         PropVer         =   50
         xcx             =   19659
         xcy             =   5292
         PropFile        =   ""
         PropDesignTime  =   1
         DeletePropFile  =   0
         IntVal          =   55
         xBfStyle1       =   63747964
         xBfStyle2       =   -1328965009
         xBfStyle3       =   -1375935757
         xBfStyle4       =   1375935757
         TabCount        =   1
         CurrentTab      =   0
         FlatProperties  =   0   'False
         BeginProperty Tab(0) {48328725-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            Text            =   "Pre-trip Reporting Field"
            ToolTip         =   ""
            Object.Align           =   0
            BackColor       =   -1
            BackColorActive =   -1
            ClientAreaColor =   -1
            Enabled         =   1
            FlyByColor      =   0
            ForeColor       =   -1
            ForeColorActive =   -1
            Name            =   ""
            Hidden          =   0
            BackColorStart  =   -1
            BackColorEnd    =   -1
            BackColorActiveStart=   -1
            BackColorActiveEnd=   -1
            ClientAreaColorStart=   -1
            ClientAreaColorEnd=   -1
         EndProperty
         BeginProperty Tabs {48328721-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
         EndProperty
         BeginProperty Scrolling {48328724-0993-461C-B983-0EAAA4A43E5C} 
            PropVer         =   0
            ToolTipCloseButton=   "Close"
            ToolTipLeftButton=   "Scroll Left/Up"
            ToolTipRightButton=   "Scroll Right/Down"
            ToolTipRestoreButton=   "Restore"
            ToolTipMinimizeButton=   "Minimize"
         EndProperty
         Appearance      =   1
         ClientArea      =   1
         Enabled         =   1
         FlatProperties  =   0
         MousePointer    =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         Style           =   21
         UseClientAreaColors=   0
         BackColor       =   -2147483628
         BorderColor     =   -2147483628
         BrightHighlightColor=   -2147483628
         ForeColor       =   0
         HighlightColor  =   6973442
         ShadowColor     =   -2147483632
         ScrollingButtonBackColor=   -2147483633
         ScrollingButtonBorderColor=   -2147483627
         ScrollingButtonForeColor=   -2147483630
         ScrollingButtonForeColorGrayed=   -2147483630
         ScrollingButtonHighlightColor=   -2147483628
         ScrollingButtonShadowColor=   -2147483632
         TabsBoldActive  =   1
         TabsDropText    =   0
         TabsEachRow     =   0
         TabsFillComplete=   0
         TabsFixed       =   0
         TabsFlyby       =   1
         TabsRows        =   1
         TabsRowIndent   =   -1
         TabsTextOnly    =   0
         TabsLeftMargin  =   0
         TabsRightMargin =   0
         ScrollButtonStyle=   5
         ScrollCondScrollButtons=   0
         ScrollFullSize  =   1
         ScrollHideScrollButtons=   0
         ScrollNoTruncate=   0
         ScrollScrollable=   0
         ScrollScrollOnLeft=   0
         AlwaysShowAccel =   0
         UseExactRegion  =   1
         UseThemes       =   -1  'True
         ShowFocusRectangle=   1
         ButtonAlignment =   0
         CloseButton     =   0
         CloseButtonDisabled=   0
         CloseButtonWMCLOSE=   1
         CloseButtonFullSize=   0
         CloseButtonAlignment=   0
         MinimizeButton  =   0
         RestoreButton   =   0
         CustomCode      =   0
         xDesign         =   15
         yDesign         =   345
         List(0)Count    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin MyFramePanel.MyFrame sharedFra 
            Height          =   2505
            Index           =   5
            Left            =   90
            Top             =   405
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4419
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
            Begin VB.Frame fraPretripMI 
               BackColor       =   &H00DADAB6&
               Caption         =   " Pre-Trip Reporting Field"
               Height          =   2350
               Left            =   60
               TabIndex        =   6
               Top             =   60
               Width           =   10740
               Begin MSFlexGridLib.MSFlexGrid msFlexPretripMI 
                  Height          =   2055
                  Left            =   120
                  TabIndex        =   7
                  Top             =   240
                  Width           =   10455
                  _ExtentX        =   18441
                  _ExtentY        =   3625
                  _Version        =   393216
                  Cols            =   5
                  FixedCols       =   0
                  RowHeightMin    =   280
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
         End
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   3120
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.PictureBox cmbContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   1005
         TabIndex        =   1
         Top             =   3120
         Visible         =   0   'False
         Width           =   1000
         Begin MSForms.ComboBox cmbEntry 
            Height          =   375
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   855
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1508;661"
            ListWidth       =   7055
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   34
         End
      End
      Begin VB.CheckBox chkEntry 
         Height          =   200
         Left            =   4680
         TabIndex        =   0
         Top             =   3120
         Visible         =   0   'False
         Width           =   200
      End
      Begin MSComCtl2.DTPicker dtpEntry 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
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
         CalendarTitleBackColor=   -2147483647
         CheckBox        =   -1  'True
         CustomFormat    =   "ddMMMyy"
         Format          =   63832065
         CurrentDate     =   36161
         MaxDate         =   109574
         MinDate         =   21916
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
         LcK2            =   $"frmPreTrip.frx":038A
         AmbientB        =   "<:?:;7B=@<7;;=:7?=;;"
      End
      Begin MyCommandButton.MyButton cmdFinish 
         Height          =   360
         Left            =   9060
         TabIndex        =   8
         Top             =   3135
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
         Left            =   10140
         TabIndex        =   9
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
         TransparentColor=   14215660
         Caption         =   "&Cancel"
         Depth           =   1
         GradientType    =   2
      End
      Begin MyCommandButton.MyButton cmdPrevious 
         Height          =   360
         Left            =   7440
         TabIndex        =   10
         Top             =   3120
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
   End
End
Attribute VB_Name = "frmPreTrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrFlex As String
Dim mstrEmailLines As String
Dim mbolClickBelowRow As Boolean
Dim promptReminderBefore As Boolean
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date
'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
Dim colFFValue As Collection

Private Sub setMIHeader(ByRef msFlex As MSFlexGrid)
   Dim i As Integer
   With msFlex
        .row = 0
        For i = 0 To .Cols - 1
            .col = i
            If i = 0 Then
              .Text = " MIS Description"
              .ColWidth(i) = 3500
            ElseIf i = 1 Then
              .Text = " MIS Value"
              .ColWidth(i) = 2000
            ElseIf i = 2 Then
              .Text = " MIS Field"
              .ColWidth(i) = 1500
            ElseIf i = 3 Then
              .Text = " MIS Datatype"
              .ColWidth(i) = 1700
            ElseIf i = 4 Then
              .Text = " MIS Length"
              .ColWidth(i) = 1000
            End If
            .ColAlignment(i) = 0
        Next
   End With
End Sub

Private Sub getReportingField(ByRef msFlex As MSFlexGrid, strLocation As String, Optional bolFromDB As Boolean)
   Dim rsMI As ADODB.Recordset
   Dim strSql As String
   Dim i As Integer
   Dim j As Integer
   Dim intLocation() As Integer
   Dim strTemp() As String
   Dim strMsg As String
   
   strTemp = Split(strLocation, ",")
   
   For j = LBound(strTemp) To UBound(strTemp)
    ReDim Preserve intLocation(j)
    intLocation(j) = CInt(strTemp(j))
   Next
   
   strSql = "Select a.*, b.Format from tblClientMI a, tblMICategory b Where a.CN='" & gobjPNR.CN & "' AND a.location = b.code AND "
   
   strSql = strSql & "("
   For j = LBound(intLocation) To UBound(intLocation)
   strSql = strSql & IIf(j > 0, " or ", "") & "a.location='" & intLocation(j) & "'"
   Next
   strSql = strSql & ")"
   strSql = strSql & " Order by FF"
   Set rsMI = gdbConn.Execute(strSql)
   
    Do Until rsMI.EOF
       i = i + 1
       If i = 1 Then
          msFlex.row = 1
       Else
          msFlex.rows = msFlex.rows + 1
          msFlex.row = msFlex.rows - 1
       End If
       msFlex.TextMatrix(msFlex.row, 2) = rsMI!FF & ""
       msFlex.TextMatrix(msFlex.row, 0) = rsMI!Description & ""
       msFlex.TextMatrix(msFlex.row, 1) = IIf(bolFromDB = False, pGetMIFromGDS(rsMI!FF & "", , rsMI!Format & ""), "")
       msFlex.TextMatrix(msFlex.row, 3) = rsMI!dataType & ""
       msFlex.TextMatrix(msFlex.row, 4) = rsMI!length & ""
       rsMI.MoveNext
    Loop
    rsMI.Close
    Set rsMI = Nothing
   
End Sub

Private Sub cmdCancel_Click()
    gbolCancelProcess = True
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    If validData Then
  
       datTouchEnd = Now
       writeDatatoGDS
       'If gbolWritingtoPNR = False Then Exit Sub
       Set gobjPNR = New CWT_GalileoPNR3.PNR
       gobjPNR.loadPNR
       'displayPNRinBar
       
       'Log formload
       'Back up on 26 Sept - Jeremy
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModRecap), _
'       IIf(gbolCreatPNR = True, gconSModRecap, ""), Me.Name, gconFormLoad, gstrProcessGrpID, _
'       datFormLoadEnd, datFormLoadStart
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModRecap), _
'       IIf(gbolCreatPNR = True, gconSModRecap, ""), Me.Name, gconTouch, gstrProcessGrpID, _
'       datTouchEnd, datFormLoadEnd
'
'       pEndProcessTimeLog gconModAir, IIf(gbolCreatPNR = True, gconSModCreatePNR, gconSModRecap), _
'       IIf(gbolCreatPNR = True, gconSModRecap, ""), Me.Name, gconProcessing, gstrProcessGrpID, _
'        , datTouchEnd
       
        pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
        gconModAir, frmSideBar.cmbSelectType.Text, gconSModPreTrip, _
        Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
        
        pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
        gconModAir, frmSideBar.cmbSelectType.Text, gconSModPreTrip, _
        Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
        
        pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
        gconModAir, frmSideBar.cmbSelectType.Text, gconSModPreTrip, _
        Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
              
       Unload Me
    End If

End Sub

Private Sub cmdPrevious_Click()
    gbolBack = True
    Unload Me
End Sub

Private Sub Form_Load()
   Dim oldParent As Long
   Dim i As Integer
   
   'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
   Dim intC As Integer
   Dim strFF As String
   
   datFormLoadStart = Now

    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
   Set Skinner1.SkinPicture = LoadPicture(App.Path & "\Icons\Skin.bmp")
   Me.Move 0, 0
   Me.Move frmSideBar.Width, 0

   'Set the header caption and width of msFlexPretripMI and msFlexPosttripMI
   setMIHeader msFlexPretripMI

   getReportingField msFlexPretripMI, "3"    'Pretrip MI with location 3
   
   datFormLoadEnd = Now
   If gbolCreatPNR = False Then If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

  'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
  Set colFFValue = New Collection
    strFF = ""
    With msFlexPretripMI
            For intC = 1 To .rows - 1
                If strFF = "" Then
                   strFF = "'" & Trim(.TextMatrix(intC, 2)) & "'"
                Else
                   strFF = strFF & ",'" & Trim(.TextMatrix(intC, 2)) & "'"
                End If
            Next
    End With
    
    'post trip MI location 3
    Set colFFValue = GetClientMIValue(gobjPNR.CN, strFF)

End Sub

Private Sub msFlexPretripMI_Click()
    Dim intTop As Integer
    Dim intLeft As Integer
    
    If mbolClickBelowRow = True Then
       mbolClickBelowRow = False
       Exit Sub
    End If
        
    intTop = sharedFra(5).Top + fraPretripMI.Top + SftTabs.Top + msFlexPretripMI.Top + msFlexPretripMI.CellTop
    intLeft = sharedFra(5).Left + fraPretripMI.Left + SftTabs.Left + msFlexPretripMI.Left + msFlexPretripMI.CellLeft
    If txtEntry.Visible = True Then txtEntry.Visible = False
    If cmbContainer.Visible = True Then cmbContainer.Visible = False
    If dtpEntry.Visible = True Then dtpEntry.Visible = False
    mstrFlex = msFlexPretripMI.Name
  
    If msFlexPretripMI.col = 1 Then
       cmbEntry.Clear
       
       'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
'       If colFFValue.Count > 0 Then
'          cmbEntry.style = fmStyleDropDownList
'       Else
'          cmbEntry.style = fmStyleDropDownCombo
'       End If
      ' If cmbEntry.style <> 2 Then
         cmbEntry.style = fmStyleDropDownCombo
         cmbEntry.Text = ""
        'End If
      
       PopulatecmbMI cmbEntry, colFFValue, CStr(msFlexPretripMI.TextMatrix(msFlexPretripMI.row, 2)), msFlexPretripMI.Text
       
       setControlPosition msFlexPretripMI, cmbContainer, intTop, intLeft, cmbEntry
       
       If cmbEntry.ListCount > 0 Then
          cmbEntry.style = fmStyleDropDownList
       Else
          cmbEntry.style = fmStyleDropDownCombo
       End If
       
    End If
End Sub

Private Sub writeDatatoGDS()
    
    Dim i
    Dim strCmd As String
    Dim strCmd2 As String
    Dim strResponse As String
    Dim strField() As String
    Dim strField2() As String
    Dim strMsg As String
    Dim strTemp As String
    
    strCmd = ""
    strCmd2 = ""
    
    If gobjPNR.CheckPNRStatus = 3 Then gobjHost.terminalEntry "IR"
    gbolWritingtoPNR = True
                  
    'MI for Pre-Trip
    strCmd2 = addMI(msFlexPretripMI)
    If strCmd2 <> "" Then
       strCmd = strCmd & IIf(strCmd <> "", "+", "") & strCmd2
       strCmd2 = ""
    End If
    
    'send entries, received & end the PNR
     'Preethi - V1.2.4 20110614 - CR 76 - Change Validation Logic For ENDPNR
    If gbolCreatPNR = True And gobjPNR.RecLoc = "" Then
    Else
    strCmd = strCmd & IIf(strCmd = "", "", "+") & "R." & IIf(frmSideBar.txtRequestor.Text = "", gobjHost.AgentSine, frmSideBar.txtRequestor.Text & "/" & gobjHost.AgentSine)
    End If
    strCmd = strCmd & "+ER+ER+ER"
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
    
    Exit Sub
    
errorWriting:
    'Prompt error message if failed to write to PNR
    gbolWritingtoPNR = False
    strMsg = "Unable to write to PNR. Response from GDS is " & Chr(13) & strResponse
    strMsg = strMsg & Chr(13) & "System will continue without ending this booking."
 
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Write to PNR"
    'gobjHost.TerminalEntry "IR"

End Sub

Private Function addMI(ByRef msFlex As MSFlexGrid) As String
    Dim i As Integer
    Dim strFF As String
    Dim strFFValue As String
    Dim intDILine As Integer
    Dim bolFound As Boolean
    
    With msFlex
        For i = 1 To msFlex.rows - 1
            strFF = Trim(.TextMatrix(i, 2))
            If strFF <> "" Then
               If IsNumeric(strFF) Then
                  strFF = "FF" & strFF
               End If
               strFFValue = Trim(.TextMatrix(i, 1))
               intDILine = 0
               bolFound = False
               bolFound = updateMI(strFF, strFFValue, intDILine)
               If bolFound And intDILine > 0 Then
                  addMI = addMI & IIf(addMI = "", "", "+") & "DI." & intDILine & "@FT-" & strFF & "/" & strFFValue
               ElseIf bolFound = False And intDILine = 0 Then
                  addMI = addMI & IIf(addMI = "", "", "+") & "DI.FT-" & strFF & "/" & strFFValue
               End If
            End If
        Next
    End With
End Function

Private Function validData() As Boolean

    Dim strMsg As String
    
    validData = True
    
    'Validate all the tabs
    strMsg = validatePretripMI
    
    If strMsg <> "" Then
       validData = False
       modMsgBox.OKMsg = "OK"
       modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
    End If
    
End Function

Private Function validatePretripMI() As String
    Dim strMsg As String
        
    If cmbEntry.Visible Then
        cmbEntry_LostFocus
    End If
    
    strMsg = strMsg & incompleteMI(msFlexPretripMI)
    validatePretripMI = strMsg
    
End Function

Private Function incompleteMI(ByRef msFlex As MSFlexGrid) As String
    Dim i As Integer
    Dim strMsg As String
        
    incompleteMI = ""
    With msFlex
         For i = 1 To msFlex.rows - 1
             If Trim(.TextMatrix(i, 1)) = "" Then
                strMsg = strMsg & "Incomplete MI for " & IIf(IsNumeric(.TextMatrix(i, 2)), "FF " & .TextMatrix(i, 2), .TextMatrix(i, 2)) & " ..." & Chr(13)
             Else
                If IsNumeric(.TextMatrix(i, 4)) Then
                    If Len(Trim(.TextMatrix(i, 1))) <> CInt(.TextMatrix(i, 4)) Then
                       strMsg = strMsg & "Incompatible length detected for " & IIf(IsNumeric(.TextMatrix(i, 2)), "FF " & .TextMatrix(i, 2), .TextMatrix(i, 2)) & " ..." & Chr(13)
                    Else
                       If UCase(.TextMatrix(i, 3)) = "NUMERIC" And Not IsNumeric(fConvertZero(.TextMatrix(i, 1))) Then
                          strMsg = strMsg & "Invalid data detected for " & IIf(IsNumeric(.TextMatrix(i, 2)), "FF " & .TextMatrix(i, 2), .TextMatrix(i, 2)) & " ..." & Chr(13)
                       End If
                    End If
                Else
                    If UCase(.TextMatrix(i, 3)) = "NUMERIC" And Not IsNumeric(fConvertZero(.TextMatrix(i, 1))) Then
                       strMsg = strMsg & "Invalid data detected for " & IIf(IsNumeric(.TextMatrix(i, 2)), "FF " & .TextMatrix(i, 2), .TextMatrix(i, 2)) & " ..." & Chr(13)
                    End If
                End If
             End If
         Next
    End With
    
    incompleteMI = strMsg
    
End Function

Private Sub cmbEntry_LostFocus()
    Dim msFlex As MSFlexGrid
    
    If mstrFlex = msFlexPretripMI.Name Then
       Set msFlex = msFlexPretripMI
       'Clement - 20080812
       'If SftTabs.Tabs.Current = 5 Then msFlexPretripMI.SetFocus
       control_LostFocus msFlex, Me, cmbEntry, "V"
       Exit Sub
    End If
    control_LostFocus msFlex, Me, cmbEntry
     'Preethi - V1.2.13 20120703 - CR169 - Populate dropdown list with preset values for Client MI
    cmbEntry.style = fmStyleDropDownCombo
End Sub

Private Function updateMI(strFF As String, strFFValue As String, intDILine As Integer) As Boolean
    Dim i As Integer
    Dim strTemp() As String
    
    For i = 1 To gobjPNR.AcctRemarkCount
        strTemp = Split(gobjPNR.AcctRemark(i).RemarkText, "/")
        If UBound(strTemp) >= 1 Then
           If UCase(Trim(strTemp(0))) = UCase(strFF) Then
              If UCase(Trim(strTemp(1))) <> UCase(strFFValue) Then
                 updateMI = True
                 intDILine = i
              Else
                 updateMI = False
                 intDILine = i
              End If
              Exit For
           End If
        End If
    Next
End Function

'Private Sub msFlexPretripMI_KeyDown(KeyCode As Integer, Shift As Integer)
'    control_KeyDown KeyCode, Shift, Me, msFlexPretripMI, False
'End Sub
'
'Private Sub msFlexPretripMI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    mouseDown msFlexPretripMI, vbLeftButton, Y
'End Sub
'
'Private Sub mouseDown(ByRef msFlex As MSFlexGrid, ByVal Button As Integer, ByVal Y As Single)
'
'    With msFlex
'         If txtEntry.Visible = True Then txtEntry.Visible = False
'         If cmbContainer.Visible = True Then cmbContainer.Visible = False
'         If dtpEntry.Visible = True Then dtpEntry.Visible = False
'         If chkEntry.Visible = True Then chkEntry.Visible = False
'
'         mstrFlex = .Name
'         .row = .MouseRow
'         .col = .MouseCol
'
'         If Button = vbRightButton Then
'             PopupMenu mnuPopUpFlex
'         ElseIf Button = vbLeftButton Then
'             If Y > .RowPos(.rows - 1) + .RowHeight(.rows - 1) Then
'                mbolClickBelowRow = True
'             Else
'                mbolClickBelowRow = False
'             End If
'         End If
'    End With
'End Sub

