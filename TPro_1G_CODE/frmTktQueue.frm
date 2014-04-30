VERSION 5.00
Begin VB.Form frmTktQueue 
   Caption         =   "Queue"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   7455
   Begin VB.CommandButton cmdMI 
      Caption         =   "Client MI"
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton OKBUtton 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Queue Details"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6855
      Begin VB.TextBox txtQCat 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtQNum 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtPCC 
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cmbQueueNo 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   5175
      End
      Begin VB.ComboBox cmbQueueOption 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label7 
         Caption         =   "Category :"
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Number :"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "PCC :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Select Queue :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Queue To :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAddRemark 
      Caption         =   "Add Remark"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblClientMI 
      Caption         =   "Click to add/amend Client MI :"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblReminder 
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   120
      TabIndex        =   14
      Top             =   3045
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "Click to add remark :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Queue Booking File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmTktQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intQueueNo As Integer
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date

Private Sub CancelButton_Click()
    disableCtrl
Set gobjPNR = New CWT_GalileoPNR3.PNR
   gobjPNR.loadPNR
   
    AddVFF
    pMoveBottomMI
    enableCtrl
    Unload Me
End Sub

Private Sub cmbQueueNo_Click()
     intQueueNo = cmbQueueNo.ItemData(cmbQueueNo.listindex)
End Sub

Private Sub cmbQueueOption_Click()
Dim bolRobotic As Boolean

If cmbQueueOption.Text = "Issue Ticket" Then
   'Use manual ticketing if cannot auto ticketing
   'bolRobotic = checkEligible
   'If bolRobotic = True Then
      'lblReminder.Caption = "Your booking will be queued to aqua QC check, then issue ticket from robotic"
      'enableButton False, False, False, False
   'Else
      'lblReminder.Caption = "Your booking will be queued to aqua QC check, then issue ticket from ticketing agent"
      'cmbQueueOption.Text = "Issue Ticket Manually"
    lblReminder.Caption = ""
    enableButton True, False, False, False
    If gstrAgcyCountryCode = "SG" Then
       txtPCC.Text = "781P"
    ElseIf gstrAgcyCountryCode = "HK" Then
       txtPCC.Text = "1IW"
    End If
   'End If
   
ElseIf cmbQueueOption.Text = "Issue Ticket Manually" Then
   lblReminder.Caption = "Your booking will be queued to manual ticketing directly"
   enableButton True, False, False, False
   
ElseIf cmbQueueOption.Text = "Auto Wait List Monitoring" Then
   lblReminder.Caption = "Your booking will be queued to aqua to monitor wait list, agent must check your queue to follow up your booking"
   enableButton True, False, False, False
   If gstrAgcyCountryCode = "SG" Then
      txtPCC.Text = "781P"
   ElseIf gstrAgcyCountryCode = "HK" Then
      txtPCC.Text = "5E4P"
   End If
ElseIf cmbQueueOption.Text = "Others" Then
   lblReminder.Caption = ""
   enableButton False, True, True, True
   txtPCC.Text = ""
Else
   enableButton False, False, False, False
End If
loadQueueNo
End Sub
Private Sub cmdAddRemark_Click()
'Unload Me
'Me.Hide
frmAddLine.Show
End Sub

Private Sub cmdMI_Click()
    Call loadClientMI(True)
End Sub
Private Sub loadClientMI(bolShow As Boolean)

   If isLoaded("frmClientMI") Then
       'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
       If bolShow = True Then
         frmClientMI.Show
       End If
    Else
        Load frmClientMI
        frmClientMI.intLocation = 7
        frmClientMI.intProdCode = frmOthSvcs.dbcProducts.BoundText
        frmClientMI.cmbMICat.Enabled = False
        frmClientMI.pGetClientMI (gobjPNR.CN)
        '230108
        'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
        If bolShow = True Then
         frmClientMI.Show
       End If
    End If
End Sub

Private Sub Form_Load()
Dim oldParent As Long

datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

Me.Move 0, 0
Me.Move frmSideBar.Width, 0


'Add Queue Option
cmbQueueOption.AddItem "Issue Ticket"
'cmbQueueOption.AddItem "Issue Ticket Manually"
cmbQueueOption.AddItem "Auto Wait List Monitoring"
cmbQueueOption.AddItem "Others"

'Add Queue Number
loadQueueNo

'Enable Client MI
If isRequireClientMI(gobjPNR.CN, 7) Then
  cmdMI.Visible = True
  lblClientMI.Visible = True
  
Else
  cmdMI.Visible = False
  lblClientMI.Visible = False
End If

'ZhiSam - V1.2.23 20130829 - CR 231 - Desktop SGHK To Disable the X function in Queue Module
    EnableCloseButton Me.hwnd, False

'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
'frmClientMI.bolCheck = False

   datFormLoadEnd = Now
   If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

End Sub
Private Function ClientMIExist() As Boolean
Dim strSQL As String
Dim rs As ADODB.Recordset

strSQL = "select * from tblclientMI WHERE CN='" & gobjPNR.CN & "' and Location='" & 7 & "'"
Set rs = gdbConn.Execute(strSQL)

If Not rs.EOF Then
    ClientMIExist = True
Else
    ClientMIExist = False
End If

End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
    Unload frmClientMI
    If UnloadMode = 0 Then
        AddVFF
        pMoveBottomMI
        Unload Me

    End If
    
End Sub

Private Sub OKButton_Click()
    Dim strResponse As String
    Dim strKey As String
    Dim strMsg As String
    Dim strRecLoc As String
    Dim intMsg As Integer
    
    
   datTouchEnd = Now
   
'CC - V1.2.24 20140129 - CR 304 - JTB Integration
If UCase(cmbQueueOption.Text) = UCase("Issue Ticket") Then
    'intMsg = MsgBox("Has the E Invoice recipient email address been updated in RI.INV Field?", vbYesNo, "CWT Desktop - Queue")
    strMsg = "Has the E Invoice recipient email address been updated in RI.INV Field?"
    modMsgBox.YESMsg = "Yes"
    modMsgBox.NOMsg = "No"
    intMsg = modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbYesNo + vbDefaultButton1, "CWT Desktop - Queue")
    
    If intMsg = vbNo Then
        Exit Sub
    End If
End If

   disableCtrl
   
   
'Preethi - V1.2.14 20120817  - CR161 - Aqua Itn - Validation in EItin and Queue Screen
Set gobjPNR = New CWT_GalileoPNR3.PNR
gobjPNR.loadPNR

If cmbQueueOption.Text = "Issue Ticket" Then

   
   If gobjPNR.CompInfo.AquaItin = True Then
      If PreLaunchAquaItinValidate = True Then
      End If
   End If
End If
     
'230108
If isRequireClientMI(gobjPNR.CN, 7) Then
        'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
        'If frmClientMI.bolCheck = False Then
        loadClientMI (False)
        If frmClientMI.incompleteMI <> "" Then
        enableCtrl
        'MsgBox "Client MI data is incomplete", vbCritical
        strMsg = "Client MI data is incomplete"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
        'Preethi - V1.2.2 20110301 -  IR8  - To solve the issue where the consultants can bypass the MI input
        loadClientMI (True)
        Exit Sub
        End If
        'Preethi - V1.2.14 20120817  - CR161 - Aqua Itn - Validation in EItin and Queue Screen
        Set gobjPNR = New CWT_GalileoPNR3.PNR
        gobjPNR.loadPNR
End If
     
     'load PNR

   
    If cmbQueueOption.Text <> "Others" Then
       If intQueueNo <> 0 And Len(txtPCC.Text) > 0 Then
       'Preethi - V1.2.14 20120817  - CR161 - Aqua Itn - Validation in EItin and Queue Screen
          If cmbQueueOption.Text = "Issue Ticket" Then
             AddVFF
             Call pMoveBottomMI
             strKey = pAddToAQQueueLog
             gobjHost.terminalEntry "NP.TQ*" & intQueueNo
             If strKey <> "" Then
                AddQKeytoNP strKey
             End If
             strRecLoc = gobjPNR.RecLoc
             gobjHost.terminalEntry ("R.TPRO QC")
             gobjHost.terminalEntry "ER"
             gobjHost.terminalEntry "ER"
             gobjHost.terminalEntry "ER"

             strResponse = gobjHost.terminalEntry("QEB/5E4P/81")

             
             If InStr(1, strResponse, "ON QUEUE") = 0 Then
                'MsgBox "Unable to queue PNR to 5E4P/81" & vbCrLf & strResponse
                strMsg = "Unable to queue PNR to 5E4P/81"
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                enableCtrl
                Exit Sub
             Else
                'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
                If gobjPNR.CompInfo.AquaItin = True Then
                'Preethi - V1.2.14 20120817  - CR161 - Aqua Itn - Validation in EItin and Queue Screen
                    'If PreLaunchAquaItinValidate = True Then
                        gobjHost.terminalEntry ("*" & strRecLoc)
                        Load frmEItin_Aqua
                        'CC - V1.2.16 20121018 - CR163 - EM trigger generation removal in QUEUE module
                        frmEItin_Aqua.mstrActivateFrom = "QueueScreen"
                        frmEItin_Aqua.Show
                    'End If
                End If
             End If
          ElseIf cmbQueueOption.Text = "Auto Wait List Monitoring" Then
             gobjHost.terminalEntry ("NP.AQ*ZZ-Queue For Waitlist")
             gobjHost.terminalEntry ("R.TPRO WL")
             gobjHost.terminalEntry "ER"
             gobjHost.terminalEntry "ER"
             gobjHost.terminalEntry "ER"
             strResponse = gobjHost.terminalEntry("QEB/" & txtPCC.Text & "/" & intQueueNo)
             If InStr(1, strResponse, "ON QUEUE") = 0 Then
                'MsgBox "Unable to queue PNR to " & txtPCC.Text & "/" & intQueueNo & vbCrLf & strResponse
                strMsg = "Unable to queue PNR to " & txtPCC.Text & "/" & intQueueNo & vbCrLf & strResponse
                modMsgBox.OKMsg = "OK"
                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                enableCtrl
                Exit Sub
             End If
          End If
       End If
       
    ElseIf cmbQueueOption.Text = "Others" Then
        If Len(Trim(txtPCC.Text)) > 0 And Len(Trim(txtQNum.Text)) > 0 Then
           gobjHost.terminalEntry ("R.TPRO OTHERS")
           gobjHost.terminalEntry "ER"
           gobjHost.terminalEntry "ER"
           gobjHost.terminalEntry "ER"
           strResponse = gobjHost.terminalEntry("QEB/" & txtPCC.Text & "/" & txtQNum & IIf(Len(Trim(txtQCat)) > 0, "*C" & txtQCat, ""))
           If InStr(1, strResponse, "ON QUEUE") = 0 Then
              'MsgBox "Unable to queue PNR to " & txtPCC.Text & "/" & txtQNum & "*C" & txtQCat & vbCrLf & strResponse
              strMsg = "Unable to queue PNR to " & txtPCC.Text & "/" & txtQNum & "*C" & txtQCat & vbCrLf & strResponse
              modMsgBox.OKMsg = "OK"
              modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
              enableCtrl
              Exit Sub
           End If
           
        Else
           'MsgBox "Missing value in PCC or Queue Number ...."
           strMsg = "Missing value in PCC or Queue Number ...."
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
           enableCtrl
           Exit Sub
        End If
        
    End If
    enableCtrl
    
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModQueue, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModQueue, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModAir, frmSideBar.cmbSelectType.Text, gconSModQueue, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd

    
    Unload Me
End Sub
Private Sub UnloadWithOption(Optional moveMI As Boolean = True)
    If moveMI Then pMoveBottomMI
    Unload Me
End Sub

Private Sub AddVFF()
   Dim intI As Integer
   Dim intCount As Integer
   Dim strAddVFF As String
   Dim strCurVFF As String
   Dim strCurVFFLoc As String
   Dim strCmd As String
   Dim strTmp() As String
   strAddVFF = ""
   strCurVFF = ""
   strCurVFFLoc = ""
   
   With gobjPNR
   
   If .HotelSegCount = 0 And .CarSegCount = 0 Then Exit Sub
   
    For intI = 1 To .AcctRemarkCount
            If Left(.AcctRemark(intI).RemarkText, 3) = "VFF" Then
            strTmp = Split(.AcctRemark(intI).RemarkText, "/")
            If UBound(strTmp) = 1 Then
           
               strCurVFF = strCurVFF & IIf(strCurVFF = "", "", "+") & "DI.FT-" & .AcctRemark(intI).RemarkText
               strCurVFFLoc = strCurVFFLoc & "." & .AcctRemark(intI).ItemNum
            End If
            End If
    Next
    
      For intI = 1 To .AcctRemarkCount
         If Left(.AcctRemark(intI).RemarkText, 2) = "FF" And InStr(strCurVFF, .AcctRemark(intI).RemarkText) = 0 Then
            strTmp = Split(.AcctRemark(intI).RemarkText, "/")
            If UBound(strTmp) = 1 Then
               strAddVFF = strAddVFF & IIf(strAddVFF = "" And strCurVFF = "", "", "+") & "DI.FT-V" & .AcctRemark(intI).RemarkText
            End If
         End If
      Next
            
      If strCurVFFLoc <> "" Then
         gobjHost.terminalEntry "DI" & strCurVFFLoc & "@"
      End If
      If strCurVFF & strAddVFF <> "" Then
         strTmp = Split(strCurVFF & strAddVFF, "+")
         If UBound(strTmp) <= 24 Then
            gobjHost.terminalEntry strCurVFF & strAddVFF
         Else
            intCount = 0
            strCmd = ""
            For intI = 0 To UBound(strTmp)
               intCount = intCount + 1
               strCmd = strCmd & IIf(strCmd = "", "", "+") & strTmp(intCount)
               If intCount = 25 Then
                  gobjHost.terminalEntry strCmd
                  strCmd = ""
                  intCount = 0
               End If
            Next
            If strCmd <> "" Then
               gobjHost.terminalEntry strCmd
            End If
         End If
      End If
   End With
End Sub

Private Function pCheckQKey(ByVal strQKey As String) As Boolean
Dim rsQueueLog As ADODB.Recordset
Dim strSQL As String

strSQL = "Select * from lookup.dbo.AQTktLog Where PNRQKey='" & strQKey & "'"
Set rsQueueLog = gdbConn.Execute(strSQL)
If rsQueueLog.EOF Then
   pCheckQKey = False
Else
   pCheckQKey = True
End If
End Function


Private Sub enableButton(ByVal bolCmbQNum As Boolean, ByVal bolTxtPcc As Boolean, ByVal bolTxtQNum As Boolean, ByVal bolTxtQCat As Boolean)
cmbQueueNo.Enabled = bolCmbQNum
txtPCC.Enabled = bolTxtPcc
txtQNum.Enabled = bolTxtQNum
txtQCat.Enabled = bolTxtQCat
End Sub

Private Function checkDueLine() As Boolean
Dim lngC As Long
Dim bolExist As Boolean
bolExist = False
'check whether due line remark contains "Invoice Total Due"
For lngC = 1 To gobjPNR.PaidDueCount
    With gobjPNR.PaidDue(lngC)
        Select Case .SegType
            Case "D"
                If .FreeText = "**INVOICE TOTAL DUE**" Then
                   bolExist = True
                   Exit For
                End If
        End Select
    End With
Next
checkDueLine = bolExist
End Function

Private Sub loadQueueNo()
cmbQueueNo.Clear
'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration

If cmbQueueOption.Text = "Issue Ticket" Then
    'If gstrAgcyCountryCode = "SG" Then
    If gstrAgcyCountryCode = "SG" And gobjPNR.CompInfo.AgencyName = "CWT" Then
        'txtPCC.Text = "781P"
        'Helena Lim- inactive queues to be closed --> confirmed now, pls close Q93/94/98/99
        cmbQueueNo.AddItem "75 - Insurance"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 75
        'Preethi - V1.2.13 20120703 - CR170 - Change the description for Queue 85
        'cmbQueueNo.AddItem "85 - Visa"
        cmbQueueNo.AddItem "85 - Super Urgent (DEP/TTL Within 4 Hours)"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 85
        cmbQueueNo.AddItem "87 - By 6:00PM"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 87
        cmbQueueNo.AddItem "88 - E-TKT"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 88
        cmbQueueNo.AddItem "90 - Non-Air PNR"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 90
        cmbQueueNo.AddItem "92 - GE Implant"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 92
        'cmbQueueNo.AddItem "93 - Implant"
        'cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 93
        'cmbQueueNo.AddItem "94 - By 5:30PM"
        'cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 94
        cmbQueueNo.AddItem "95 - Amendment"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 95
        cmbQueueNo.AddItem "96 - No-Air PNR Invoicing"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 96
        cmbQueueNo.AddItem "97 - MCO/MPD Invoice Only"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 97
        'cmbQueueNo.AddItem "98 - By 1:30PM"
        'cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 98
        'cmbQueueNo.AddItem "99 - By 9:30AM"
        'cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 99
    'ElseIf gstrAgcyCountryCode = "HK" Then
    ElseIf gstrAgcyCountryCode = "SG" And gobjPNR.CompInfo.AgencyName = "JTB" Then
        cmbQueueNo.AddItem "88 - E-TKT"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 88
       
    ElseIf gstrAgcyCountryCode = "HK" And gobjPNR.CompInfo.AgencyName = "CWT" Then
        'txtPCC.Text = "1IW"
        cmbQueueNo.AddItem "71 - JPMC"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 71
        cmbQueueNo.AddItem "72 - BP"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 72
        cmbQueueNo.AddItem "77 - UBS"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 77
        cmbQueueNo.AddItem "78 - HEAD OFFICE"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 78
    Else
        cmbQueueNo.AddItem "N/A"
        cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 0
    End If
ElseIf cmbQueueOption.Text = "Auto Wait List Monitoring" Then
    cmbQueueNo.AddItem "84 - WAITLIST"
    cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 84
Else
    cmbQueueNo.AddItem "N/A"
    cmbQueueNo.ItemData(cmbQueueNo.NewIndex) = 0
End If
cmbQueueNo.listindex = 0
'---
End Sub

Private Sub txtPCC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtQCat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtQNum_KeyPress(KeyAscii As Integer)
KeyAscii = fAllowNumeric(KeyAscii)
End Sub
Private Sub disableCtrl()
    Me.MousePointer = 11
    OKButton.Enabled = False
    cancelButton.Enabled = False
End Sub
Private Sub enableCtrl()
    Me.MousePointer = 0
    OKButton.Enabled = True
    cancelButton.Enabled = True
End Sub
