VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmExchangeOrder 
   BorderStyle     =   0  'None
   Caption         =   "CWT Travel Pro - Other Services"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4800
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5535
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3480
         Width           =   1275
      End
      Begin VB.CheckBox chkEOCR 
         Caption         =   "Print EO using crystal report"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Reprint"
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   915
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3480
         Width           =   1275
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3480
         Width           =   1275
      End
      Begin VB.Frame fraAction 
         Caption         =   "Requested Action"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   960
         TabIndex        =   5
         Top             =   1380
         Width           =   2895
         Begin VB.CheckBox chkAction 
            Caption         =   "Email Exchange Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   1140
            Width           =   2235
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Fax Exchange Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   780
            Width           =   2235
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Print Exchange Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   420
            Width           =   2235
         End
         Begin VB.CheckBox chkAction 
            Caption         =   "Request Cheque"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Top             =   1500
            Width           =   2235
         End
      End
      Begin VB.TextBox txtEONum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "has been assigned."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   4
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Exchange Order "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Caption         =   "Exchange Order"
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
         Height          =   435
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmExchangeOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbolPreview As Boolean
Public mstrEOType As String
Dim EOTxtPath As String
Public startTime As Date
Dim strPax() As String
Dim mbolReportPath As Boolean
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date
Dim datProGDSStart As Date



Private Sub chkEOCR_Click()
   Dim strMsg As String
   
   SetPreviewButton
   If chkEOCR.value = 1 Then
      If mbolReportPath Then
      
      Else
          chkEOCR.value = 0
          'MsgBox "Report Path not found."
          strMsg = "Report Path not found."
          modMsgBox.OKMsg = "OK"
          modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      End If
   End If

End Sub

'Dim wrdApp      As Word.Application
'Dim wrdDoc      As Word.Document

Private Sub cmdClose_Click()
   'If cmdOK(0).Enabled = False Or cmdOK(1).Enabled = False Then
   If CmdOk(0).Enabled = False Then
      gbolIgnoreEO = False
   Else
      gbolIgnoreEO = True
   End If
   Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim EOType As String
Dim i As Integer
Dim freefields As String
Dim strProductType As String
Dim strProduct As String
Dim blnNumberSet As Boolean
Dim strTmp As String
Dim strMsg As String

datTouchEnd = Now
gbolRaiseEOReport = False
gbolPreviewEO = False
gbolPreviewCancel = False

If Index = 0 Then
   mbolPreview = False
Else
   mbolPreview = True
   gbolPreviewEO = True
End If

If chkAction(0).value = 0 And chkAction(1).value = 0 _
   And chkAction(2).value = 0 And chkAction(3).value = 0 Then
   gbolIgnoreEO = True
   'cmdOK.Enabled = True
     Exit Sub
End If
CmdOk(0).Enabled = False
CmdOk(1).Enabled = False
cmdClose.Enabled = False

If gbolEOAmend = False And txtEONum = "" Then
    blnNumberSet = modOthSvcs.SetEONumber
        If blnNumberSet = False Then
            cmdClose.Enabled = False
            Exit Sub
        End If
    txtEONum = gobjEO.EONumber
End If
'Call modOthSvcs.SetEONumber
'If gbolEOAmend = False Then txtEONum = gobjEO.EONumber
'Added on 8/3/2005: To end PNR by writing TUR line to get RecLoc

    strProductType = frmOthSvcs.datProducts.Recordset![Type]
    strProduct = frmOthSvcs.dbcProducts.BoundText
    '230108 remove if
    'If strProductType = "CT" Or strProductType = "BT" Or (strProductType = "MS" And (strProduct = "35" Or strProduct = "50")) Then
        freefields = ""
        If gobjEO.FF7 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "7-" & gobjEO.FF7
        If gobjEO.FF8 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "8-" & gobjEO.FF8
        If gobjEO.FF81 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "81-" & gobjEO.FF81
        'CS Change EC
        'If gobjEO.FF26 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "26-" & gobjEO.FF26
        'If gobjEO.RS <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "30-" & gobjEO.RS
        
        If gobjEO.FF38 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "38-" & gobjEO.FF38
        ''CS Add FF41
        'If gobjEO.FF41 <> "" Then freefields = freefields & IIf(freefields <> "", "/", "") & "41-" & gobjEO.FF41
        
        If frmClientMI.MSXfreefields <> "" Then
            freefields = freefields & "/" & frmClientMI.MSXfreefields
        End If
    'End If

'Call modOthSvcs.WriteOSToGDS(gobjEO, strProductType, gStartOthSvcsTime, freefields)
'If gobjEO.PNRRecLoc = "" Then Call modOthSvcs.SetRecLoc


'If chkAction(0).Value = 1 Then
'   EOType = EOType & "Print;"
'   PrintEO
'End If
If chkAction(0).value = 1 Then
   EOType = EOType & "Print;"
   If chkEOCR.value = 0 Then
      PrintEO
   End If
End If

If chkAction(1).value = 1 Then
    EOType = EOType & "Fax;"
    'FaxEO
End If
If chkAction(2).value = 1 Then
    EOType = EOType & "Email;"
    'EmailEO
End If
If chkAction(3).value = 1 Then
    EOType = EOType & "Cheque;"
End If
EOType = Left(EOType, Len(EOType) - 1)

Call UpdateExchOrdDB(gobjEO, EOType)


       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, "SYSTEM PROCESSING-UPDATE DB", gstrProcessGrpID, , datTouchEnd



'20070515
strTmp = ""
mstrEOType = ""

'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    If bolAgencyNameCheck(gobjPNR.CompInfo.AgencyName) = False Then
        Exit Sub
    End If
'--

If chkAction(2).value = 1 Or chkAction(1).value = 1 Or _
   (chkAction(0).value = 1 And chkEOCR.value = 1) Then
      gEOID = "'" & txtEONum & "'"
      'gEOReportName = "EO" & "-" & gstrAgcyCountryCode & ".rpt"
      'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
      gEOReportName = "EO" & "-" & gstrAgcyCountryCode & "-" & gobjPNR.CompInfo.AgencyName & ".rpt"
      '--
      If chkAction(2).value = 1 Then
         strTmp = strTmp & IIf(strTmp = "", "", ",") & "Email"
      End If
      If chkAction(1).value = 1 Then
         strTmp = strTmp & IIf(strTmp = "", "", ",") & "E-Fax"
      End If
      If chkAction(0).value = 1 And chkEOCR.value = 1 Then
         strTmp = strTmp & IIf(strTmp = "", "", ",") & "Print"
      End If
      gbolIndEO = True
      mstrEOType = strTmp
      'frmEOCrystalReport.Visible = False
      
      Load frmEOCrystalReport
      
      If Index = 1 Then
         frmEOCrystalReport.Caption = "EO Preview - " & strTmp
         If strTmp = "Print" Then
            frmEOCrystalReport.mnuAccept.Caption = "Print"
         End If
         frmEOCrystalReport.CR1.EnablePrintButton = False
         'frmEOCrystalReport.Visible = True
         frmEOCrystalReport.Show '1 , Me
        
        Do
            DoEvents
        Loop Until isLoaded("frmEOCrystalReport") = False
        
      ElseIf Index = 0 Then
         
         Unload frmEOCrystalReport
      End If
      gbolIndEO = False
Else
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, "SYSTEM PROCESSING-UPDATE DB", gstrProcessGrpID, , datTouchEnd

End If

If gbolPreviewCancel Then
   CmdOk(0).Enabled = True
   CmdOk(1).Enabled = True
   cmdClose.Enabled = True
   gbolPreviewCancel = False
   Exit Sub
End If
datProGDSStart = Now
Call modOthSvcs.WriteOSToGDS(gobjEO, strProductType, gStartOthSvcsTime, freefields)
If gobjEO.PNRRecLoc = "" Then Call modOthSvcs.SetRecLoc
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, "SYSTEM PROCESSING-WRITE GDS", gstrProcessGrpID, , datProGDSStart





'If chkAction(0).value = 1 And chkEOCR.value = 1 Then
'      gEOID = "'" & txtEONum & "'"
'      '20070515
'      'gEOReportName = "EO.rpt"
'      gEOReportName = "EO" & "-" & gstrAgcyCountryCode & ".rpt"
'      strTmp = strTmp & IIf(strTmp = "", "", ",") & "Print"
'      gbolIndEO = True
'
'      mstrEOType = strTmp
'      Load frmEOCrystalReport
'      If Index = 1 Then
'         frmEOCrystalReport.Caption = "EO Preview - " & strTmp
'         If strTmp = "Print" Then
'            frmEOCrystalReport.mnuAccept.Caption = "Print"
'         End If
'         frmEOCrystalReport.Show 1, Me
'      ElseIf Index = 0 Then
'         Unload frmEOCrystalReport
'      End If
'      'Load frmEOCrystalReport
'      'frmEOCrystalReport.Caption = "EO Preview - " & strTmp
'      'If strTmp = "Print" Then
'      '   frmEOCrystalReport.mnuAccept.Caption = "Print"
'      'End If
'      'frmEOCrystalReport.Show 1, Me
'      'gbolIndEO = False
'End If

'If gbolEOAmend = False Then
'   If UCase(gstrAgcyCountryCode) = "SG" Then
'      gobjHost.TerminalEntry "R.TPRO XO"
'      gobjHost.TerminalEntry "ER"
'      gobjHost.TerminalEntry "ER"
'      gobjHost.TerminalEntry "QEB/781P/78"
'   Else
'
'   End If
'End If

    If chkAction(0).value = 1 Then
       cmdPrint.Visible = True
    End If
    
    fraAction.Enabled = False
    cmdClose.Enabled = True
    'Me.Hide
    
End Sub

Private Sub SetPreviewButton()
    If chkAction(0).value = 1 And chkAction(1).value = 0 And _
       chkAction(2).value = 0 And chkEOCR.value = 0 Then
       CmdOk(1).Enabled = False
    Else
       CmdOk(1).Enabled = True
    End If
End Sub

Private Sub chkAction_Click(Index As Integer)

'the following is temporary - the SDK for the fax/email is with the Singapore Develepor
' He will need to add the code and additional controls to handle these options

'If gstrAgcyCountryCode = "HK" Then
' If Index = 1 And chkAction(Index).Value = vbChecked Then
'    MsgBox "This option is not currently available", vbApplicationModal + vbExclamation
'    chkAction(Index).Value = vbUnchecked
'  End If
'End If

Dim strAgentEmail As String
Dim strMsg As String
Dim strTmp1() As String
Dim intTmpI As Integer

'strAgentEmail = GetAgentEmail(IIf(gobjPNR.Agent = "", gobjHost.AgentSine, gobjPNR.Agent), gobjPNR.PCCOwner)
'strAgentEmail = GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjHost.AgentSine, gobjHost.AgentPcc, , True)
strAgentEmail = gobjEO.ReplyEmail

SetPreviewButton



Select Case Index
   Case 0
      
   Case 1
      If chkAction(Index).value = vbChecked Then
         'If gstrAgcyCountryCode = "HK" Then
         '   MsgBox "This option is not currently available", vbApplicationModal + vbExclamation
         '   chkAction(Index).Value = vbUnchecked
         'Else
        If gstrAgcyCountryCode = "HK" And _
           InStr(1, strAgentEmail, "@") = 0 Then
           strMsg = strMsg & "Invalid agent's email." & vbCrLf
           'chkAction(Index).value = vbUnchecked
           'Exit Sub
        End If
        
        If InStr(1, gobjEO.FaxNo, " ") > 0 Then
           strMsg = strMsg & "Fax number cannot accept space." & vbCrLf
        ElseIf Trim(gobjEO.FaxNo) = "" Then
           strMsg = strMsg & "Invalid fax number." & vbCrLf
        Else
           strTmp1 = Split(gobjEO.FaxNo, ",")
           For intTmpI = 0 To UBound(strTmp1)
              If isNumber(strTmp1(intTmpI)) = False Then
                 strMsg = strMsg & "Invalid fax number." & vbCrLf
              End If
           Next
        End If
        'If IsNumeric(gobjEO.FaxNo) = False Then
        '   strMsg = strMsg & "Invalid vendor's fax number." & vbCrLf
        '   'chkAction(Index).value = vbUnchecked
        '   'Exit Sub
        'End If
         'End If
        If strMsg <> "" Then
           chkAction(Index).value = vbUnchecked
           'strMsg = strMsg & "Please contact your operation manager."
           'MsgBox strMsg
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
           GoTo ToEnd
        End If
      End If
   Case 2
      If chkAction(Index).value = vbChecked Then
        
        If Trim(strAgentEmail) = "" Then
           strMsg = strMsg & "Reply Email in EO is required." & vbCrLf
        End If
        
        If InStr(strAgentEmail, ";") > 0 Then
           strMsg = strMsg & "Only 1 email address is allowed in reply email in EO." & vbCrLf
        End If
        
        If InStr(1, strAgentEmail, "@") = 0 And Trim(strAgentEmail) <> "" Then
           strMsg = strMsg & "Invalid reply email in EO." & vbCrLf
           'chkAction(Index).value = vbUnchecked
           'Exit Sub
        End If
        
        If InStr(1, Trim(gobjEO.Email), " ") > 0 Then
           strMsg = strMsg & "Email cannot accept space..." & vbCrLf
        ElseIf Trim(gobjEO.Email) = "" Then
           strMsg = strMsg & "Invalid vendor's email." & vbCrLf
        Else
           strTmp1 = Split(gobjEO.Email, ";")
           For intTmpI = 0 To UBound(strTmp1)
              If InStr(1, strTmp1(intTmpI), "@") = 0 Then
                 strMsg = strMsg & "Invalid vendor's email." & vbCrLf
              End If
           Next
        End If
                
        'If InStr(1, gobjEO.Email, "@") = 0 Then
        '   strMsg = strMsg & "Invalid vendor's email." & vbCrLf
        '   'chkAction(Index).value = vbUnchecked
        '   'Exit Sub
        'End If
        If strMsg <> "" Then
           chkAction(Index).value = vbUnchecked
           'strMsg = strMsg & "Please contact your operation manager."
           'MsgBox strMsg
           modMsgBox.OKMsg = "OK"
           modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Data Required"
           GoTo ToEnd
        End If
      End If
End Select

ToEnd:
If chkAction(0).value = 0 And chkAction(1).value = 0 And _
   chkAction(2).value = 0 Then
   CmdOk(1).Enabled = False
Else
   CmdOk(1).Enabled = True
End If
   
'If Index = 3 And chkAction(Index).Value = vbChecked Then
'    MsgBox "This option is not currently available", vbApplicationModal + vbExclamation
'    chkAction(Index).Value = vbUnchecked
'End If

'If Index = 2 And chkAction(Index).Value = vbChecked Then
'    MsgBox "This option is not currently available", vbApplicationModal + vbExclamation
'    chkAction(Index).Value = vbUnchecked
'End If

End Sub

Private Function isNumber(Text As String) As Boolean
   Dim i As Integer
   
   isNumber = True
   If Text = "" Then
      isNumber = False
      Exit Function
   End If
   
   For i = 1 To Len(Text)
      If IsNumeric(Mid(Text, i, 1)) = False Then
         isNumber = False
         Exit Function
      End If
   Next
End Function

Private Sub cmdPrint_Click()
  
   If chkEOCR.value = 0 Then
      PrintEO
   Else
      gEOID = txtEONum
      gbolPreviewEO = False
      gbolRaiseEOReport = True
      gbolIndEO = True
      Load frmEOCrystalReport
      frmEOCrystalReport.mnuAccept.Caption = "Print"
      frmEOCrystalReport.CR1.EnablePrintButton = False
      frmEOCrystalReport.Show
      Do
            DoEvents
      Loop Until isLoaded("frmEOCrystalReport") = False
      gbolRaiseEOReport = False
      gbolIndEO = False
   End If
End Sub


Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim oldParent As Long
    
    datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
    
    
    
    gbolIgnoreEO = True
    
    CmdOk(1).Enabled = False

    If gstrReportPath <> "" Then mbolReportPath = True
    'getDReportPath
    'If gbolEOAmend = False Then
    'Remove on 24/02/05: Visa ticket no to follow EO number
       'If Not Visa Or UCase(gstrAgcyCountryCode) = "HK" Then
     '     txtEONum = AssignEONum
     '    txtEONum = gobjEO.EONumber
     '     If gobjEO.TicketNumber = "0000" Then
             'tktlen
     '        gobjEO.TicketNumber = frmOthSvcs.datProducts.Recordset![TktPrefix] & "0" & Right(gobjEO.EONumber, Len(gobjEO.EONumber) - Len(frmOthSvcs.datProducts.Recordset![TktPrefix] & Format(Now, "yymm")))
     '     End If
       'Else
        'txtEONum = gobjEO.EONumber
       'End If
    'Else
    '   txtEONum = gobjEO.EONumber
    'End If
    'If gbolEOAmend = True Then
        txtEONum = gobjEO.EONumber
    'End If
    
    
    'added on 14062005: auto raise cheque for external vendor
    If UCase(gstrAgcyCountryCode) = "SG" Then
        If IsNull(frmOthSvcs.datSelectedVendor.Recordset!RaiseType) Then
            chkAction(3).Enabled = False
        Else
            chkAction(3).Enabled = True
            chkAction(3).value = 1
        End If
        
       
       'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
        If gobjPNR.CompInfo.AgencyName = "JTB" Then
            chkAction(1).Visible = False
            chkAction(1).Enabled = False
        End If
        '--
        'chkEOCR.Visible = True
    Else
        'chkEOCR.Visible = False
    End If
    
    cmdPrint.Visible = False
    '20070515
    'Set rs = gdbConn.Execute("Select EOTxtPath from tblPath")
    'EOTxtPath = rs!EOTxtPath
    Set rs = gdbConn.Execute("Select EOTxtPath,EMailEOFilePath, FaxEOFilePath from tblPath")
    EOTxtPath = rs!EOTxtPath
    gstrEOEmailPath = rs!EMailEOFilePath & ""
    gstrEOFaxPath = rs!FaxEOFilePath & ""
    rs.Close
    Set rs = Nothing
    
    datFormLoadEnd = Now
    If gstrCurrentPNR <> gobjPNR.RecLoc Or gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey
End Sub

'if visa, don't need to set EONumber, EONumber is already set in Visa form
Private Function Visa() As Boolean


If gobjEO.ProductCode = "06" Or gobjEO.ProductCode = "37" Then
Visa = True
End If


End Function
Private Sub PrintEO()
   Dim i As Long
   Dim intC As Integer
   
   Dialog.ShowPrinter
   
   
   Printer.FontSize = 12
   Printer.FontBold = False
   Printer.FontName = "Courier new"
   Printer.Print ""
   Printer.Print ""
   For intC = 0 To 7
       Printer.Print ""
   Next
   If gbolEOAmend Then
      Printer.FontBold = True
      Printer.Print Space(4) & UCase("Amendment")
      Printer.FontBold = False
   End If
   
   If gobjEO.ContactPerson <> "" Then
      Printer.Print Space(4) & "Attn: " & gobjEO.ContactPerson
   End If
   Printer.Print Space(4) & gobjEO.VendorName
   Printer.Print Space(4) & gobjEO.Address1 ' frmOthSvcs.datSelectedVendor.Recordset!Address1
   Printer.Print Space(4) & gobjEO.Address2 ' frmOthSvcs.datSelectedVendor.Recordset!Address2
   Printer.Print Space(4) & gobjEO.City ' frmOthSvcs.datSelectedVendor.Recordset!City
   Printer.Print Space(4) & gobjEO.Country ' frmOthSvcs.datSelectedVendor.Recordset!Country
   Printer.Print ""
   Printer.FontBold = True
   Printer.Print Space(4) & "EXCHANGE ORDER"
   Printer.Print Space(4) & "--------------"
   Printer.FontBold = False
   Printer.Print Space(4) & "EO Number      : " & gobjEO.EONumber
   
 
   If InStr(gobjEO.PaxName, vbCrLf) <> 0 Then
    strPax = Split(gobjEO.PaxName, vbCrLf)
   Else
    strPax = Split(gobjEO.PaxName, ",")
   End If
   If UBound(strPax) >= 0 Then
        For intC = 0 To UBound(strPax)
            If intC = 0 Then
                Printer.Print Space(4) & "Passenger Name : " & strPax(intC)
            Else
                Printer.Print Space(4) & "                 " & strPax(intC)
            End If
        Next
   End If
   

   
   Printer.Print Space(4) & "Agent Name     : " & gobjEO.CreatedByName
   Printer.Print Space(4) & "Agent ID       : " & gobjEO.CreatedBy
   Printer.Print Space(4) & "Record Locator : " & gobjEO.PNRRecLoc
   Printer.Print Space(4) & "TEL            : " & gstrAgcyPhone
   'Printer.Print Space(4) & "To             : " & gobjEO.VendorName
   'Printer.Print Space(4) & "Fax            : " & gobjEO.FaxNo
   'Printer.Print Space(4) & "Email          : " & gobjEO.Email
   Printer.Print Space(4) & "Date           : " & Format(Date, "Medium Date")
   Printer.Print ""
   Printer.FontBold = True
   Printer.Print Space(4) & "Service Info  "
   Printer.FontBold = False
   Printer.Print Space(4) & "Nett Cost      : " & Format(gobjEO.Cost, gstrAgcyCurrFormat)
   'Printer.Print Space(4) & "Commission     : " & gobjEO.CommissionAmt
     'If gobjEO.TaxCount = 2 Then
     ' If UCase(gobjEO.Tax(1).Code) = "GST" Then
     '    Printer.Print Space(4) & "GST            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat)
     ' ElseIf UCase(gobjEO.Tax(2).Code) = "GST" Then
     '    Printer.Print Space(4) & "GST            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat)
     ' End If
     ' If UCase(gobjEO.Tax(1).Code) <> "GST" Then
     '    Printer.Print Space(4) & "Tax            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(1).Code
     ' End If
     ' If UCase(gobjEO.Tax(2).Code) <> "GST" Then
     '    Printer.Print Space(4) & "Tax            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(2).Code
     ' End If
     
      If gobjEO.TaxCount > 0 Then
      If UCase(gobjEO.Tax(1).Code) = "GST" Then
         Printer.Print Space(4) & "GST            : " & Format(gobjEO.NettGST, gstrAgcyCurrFormat)
      'ElseIf UCase(gobjEO.Tax(2).Code) = "GST" Then
      '   Printer.Print Space(4) & "GST            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat)
      End If
      If UCase(gobjEO.Tax(1).Code) <> "GST" Then
         Printer.Print Space(4) & "Tax            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(1).Code
      End If
      If gobjEO.TaxCount > 1 Then
        If UCase(gobjEO.Tax(2).Code) <> "GST" Then
           Printer.Print Space(4) & "Tax            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(2).Code
        End If
      End If
     
     
  'remove on 21/4/05: requested by helena
  
  ' ElseIf gobjEO.TaxCount = 1 Then
  '    If UCase(gobjEO.Tax(1).Code) = "GST" Then
  '       Printer.Print Space(4) & "GST            : " & gobjEO.Tax(1).Amount
  '    Else
  '       Printer.Print Space(4) & "Tax            : " & gobjEO.Tax(1).Amount & gobjEO.Tax(1).Code
  '    End If
   'Printer.Print Space(4) & "Total          : " & Format(gobjEO.Cost + gobjEO.EOTaxTotal, gstrAgcyCurrFormat)
   If UCase(gobjEO.Tax(1).Code) = "GST" Then
        Printer.Print Space(4) & "Total          : " & Format(gobjEO.Cost + gobjEO.NettGST, gstrAgcyCurrFormat)
   Else
        Printer.Print Space(4) & "Total          : " & Format(gobjEO.Cost + gobjEO.EOTaxTotal, gstrAgcyCurrFormat)
   End If
   End If
   'modified on 21/4/05: requested by helena
   'added on 6/4/2004
   'Printer.Print Space(4) & "Total          : " & gobjEO.Cost + gobjEO.EOTaxTotal
   
     
   For i = 1 To gobjEO.DescriptionLinesCount
      If i = 1 Then
         Printer.Print Space(4) & "Description    : " & gobjEO.DescriptionLine(i)
      Else
         Printer.Print Space(4) & "               " & gobjEO.DescriptionLine(i)
      End If
   Next
   If gobjEO.RemarkCount <> 0 Then
      For i = 1 To gobjEO.RemarkCount
         If i = 1 Then
            Printer.Print Space(4) & "Remark         : "
            Printer.Print Space(4) & gobjEO.Remark(i)
         Else
            Printer.Print Space(4) & gobjEO.Remark(i)
         End If
      Next
   End If
   Printer.Print ""
   Printer.Print Space(4) & "Please prepare document for our collection today."
   Printer.Print Space(4) & "Thank you"
   Printer.EndDoc
   'MsgBox "Print Successful!", vbOKOnly, "Print Successful"

End Sub

Private Sub FaxEO()
   Dim i As Long
   Dim tmpText As String
   Dim intC As Integer
   tmpText = EOTxtPath & gstrAgcyCountryCode & "FAX" & Format(Now, "DDMMhhmmss") & ".txt"
   
   Open tmpText For Output As #1
   If gbolEOAmend Then
      Print #1, UCase("Amendment")
   End If
   If gobjEO.ContactPerson <> "" Then
      Print #1, "Attn: " & gobjEO.ContactPerson
   End If
   Print #1, gobjEO.VendorName
   Print #1, gobjEO.Address1 ' frmOthSvcs.datSelectedVendor.Recordset!Address1
   Print #1, gobjEO.Address2 ' frmOthSvcs.datSelectedVendor.Recordset!Address2
   Print #1, gobjEO.City ' frmOthSvcs.datSelectedVendor.Recordset!City
   Print #1, gobjEO.Country ' frmOthSvcs.datSelectedVendor.Recordset!Country
   Print #1, ""
   Print #1, "EXCHANGE ORDER"
   Print #1, "--------------"
   Print #1, "EO Number      : " & gobjEO.EONumber
   'Print #1, "Passenger Name : " & gobjEO.PaxName
  'strPax = Split(gobjEO.PaxName, ",")
   If InStr(gobjEO.PaxName, vbCrLf) <> 0 Then
    strPax = Split(gobjEO.PaxName, vbCrLf)
   Else
    strPax = Split(gobjEO.PaxName, ",")
   End If
   
   If UBound(strPax) >= 0 Then
        For intC = 0 To UBound(strPax)
            If intC = 0 Then
                Print #1, "Passenger Name : " & strPax(intC)
            Else
                Print #1, "                 " & strPax(intC)
            End If
        Next
   End If
   Print #1, "Agent Name     : " & gobjEO.CreatedByName
   Print #1, "Agent ID       : " & gobjEO.CreatedBy
   Print #1, "Record Locator : " & gobjEO.PNRRecLoc
   Print #1, "TEL            : " & gstrAgcyPhone
   'Print #1, "To             : " & gobjEO.VendorName
   Print #1, "Fax            : " & gobjEO.FaxNo
   'Print #1, "Email          : " & gobjEO.Email
   Print #1, "Date           : " & Format(Date, "Medium Date")
   Print #1, ""
   Print #1, "Service Info  "
   Print #1, "Nett Cost      : " & gobjEO.Cost
    'If gobjEO.TaxCount = 2 Then
    '  If UCase(gobjEO.Tax(1).Code) = "GST" Then
    '     Print #1, "GST            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat)
    '  ElseIf UCase(gobjEO.Tax(2).Code) = "GST" Then
    '     Print #1, "GST            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat)
    '  End If
    '  If UCase(gobjEO.Tax(1).Code) <> "GST" Then
    '     Print #1, "Tax            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(1).Code
    '  End If
    '  If UCase(gobjEO.Tax(2).Code) <> "GST" Then
    '     Print #1, "Tax            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(2).Code
    '  End If
     If gobjEO.TaxCount > 0 Then
      If UCase(gobjEO.Tax(1).Code) = "GST" Then
         Print #1, "GST            : " & Format(gobjEO.NettGST, gstrAgcyCurrFormat)
      End If
      If UCase(gobjEO.Tax(1).Code) <> "GST" Then
         Print #1, "Tax            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(1).Code
      End If
      If gobjEO.TaxCount > 1 Then
        If UCase(gobjEO.Tax(2).Code) <> "GST" Then
           Print #1, "Tax            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(2).Code
        End If
      End If
    'Added on 6/4/2005
   ' Print #1, "Total          : " & Format(gobjEO.Cost + gobjEO.EOTaxTotal, gstrAgcyCurrFormat)
   'ElseIf gobjEO.TaxCount = 1 Then
   '   If UCase(gobjEO.Tax(1).Code) = "GST" Then
   '      Print #1, "GST            : " & gobjEO.Tax(1).Amount
   '   Else
   '      Print #1, "Tax            : " & gobjEO.Tax(1).Amount & gobjEO.EONumber
   '   End If
    If UCase(gobjEO.Tax(1).Code) = "GST" Then
        Print #1, "Total          : " & Format(gobjEO.Cost + gobjEO.NettGST, gstrAgcyCurrFormat)
    Else
        Print #1, "Total          : " & Format(gobjEO.Cost + gobjEO.EOTaxTotal, gstrAgcyCurrFormat)
    End If

   End If
   For i = 1 To gobjEO.DescriptionLinesCount
      If i = 1 Then
         Print #1, "Description    : " & gobjEO.DescriptionLine(i)
      Else
         Print #1, "                 " & gobjEO.DescriptionLine(i)
      End If
   Next
   If gobjEO.RemarkCount <> 0 Then
      For i = 1 To gobjEO.RemarkCount
         If i = 1 Then
            Print #1, "Remark         : "
            Print #1, gobjEO.Remark(i)
         Else
            Print #1, gobjEO.Remark(i)
         End If
      Next
   End If
   Print #1, ""
   Print #1, "Please prepare document for our collection today."
   Print #1, "Thank you"
   Close #1
   'MsgBox "Print Successful!", vbOKOnly, "Print Successful"

End Sub

Private Sub EmailEO()
   Dim i As Long
   Dim tmpText As String
   Dim intC As Integer
   Dim strAgentEmail As String
   
   'strAgentEmail = GetAgentEmail(gobjPNR.Agent, gobjPNR.PCCOwner)
   'strAgentEmail = GetAgentEmail(IIf(gobjPNR.Agent = "", gobjHost.AgentSine, gobjPNR.Agent), gobjPNR.PCCOwner)
   'strAgentEmail = GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjHost.AgentSine, gobjHost.AgentPcc, , True)
   strAgentEmail = gobjEO.ReplyEmail
   tmpText = EOTxtPath & gstrAgcyCountryCode & "EMAIL" & Format(Now, "DDMMhhmmss") & ".txt"
   
   Open tmpText For Output As #1
   If gbolEOAmend Then
      Print #1, UCase("Amendment")
   End If
   If gobjEO.ContactPerson <> "" Then
      Print #1, "Attn: " & gobjEO.ContactPerson
   End If
   Print #1, gobjEO.VendorName
   Print #1, gobjEO.Address1 ' frmOthSvcs.datSelectedVendor.Recordset!Address1
   Print #1, gobjEO.Address2 ' frmOthSvcs.datSelectedVendor.Recordset!Address2
   Print #1, gobjEO.City 'frmOthSvcs.datSelectedVendor.Recordset!City
   Print #1, gobjEO.Country ' frmOthSvcs.datSelectedVendor.Recordset!Country
   Print #1, ""
   Print #1, "EXCHANGE ORDER"
   Print #1, "--------------"
   Print #1, "EO Number      : " & gobjEO.EONumber
   'Print #1, "Passenger Name : " & gobjEO.PaxName
   'strPax = Split(gobjEO.PaxName, ",")
   If InStr(gobjEO.PaxName, vbCrLf) <> 0 Then
    strPax = Split(gobjEO.PaxName, vbCrLf)
   Else
    strPax = Split(gobjEO.PaxName, ",")
   End If
   If UBound(strPax) >= 0 Then
        For intC = 0 To UBound(strPax)
            If intC = 0 Then
                Print #1, "Passenger Name : " & strPax(intC)
            Else
                Print #1, "                 " & strPax(intC)
            End If
        Next
   End If
   Print #1, "Agent Name     : " & gobjEO.CreatedByName
   Print #1, "Agent ID       : " & gobjEO.CreatedBy
   Print #1, "Agent Email    : " & strAgentEmail   '20070329
   Print #1, "Record Locator : " & gobjEO.PNRRecLoc
   Print #1, "TEL            : " & gstrAgcyPhone
   'Print #1, "To             : " & gobjEO.VendorName
   'Print #1, "Fax            : " & gobjEO.FaxNo
   Print #1, "Email          : " & gobjEO.Email
   Print #1, "Date           : " & Format(Date, "Medium Date")
   Print #1, ""
   Print #1, "Service Info  "
   Print #1, "Nett Cost      : " & gobjEO.Cost
   'Print #1, "Commission     : " & gobjEO.CommissionAmt
    'If gobjEO.TaxCount = 2 Then
    '  If UCase(gobjEO.Tax(1).Code) = "GST" Then
    '     Print #1, "GST            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat)
    '  ElseIf UCase(gobjEO.Tax(2).Code) = "GST" Then
    '     Print #1, "GST            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat)
    '  End If
    '  If UCase(gobjEO.Tax(1).Code) <> "GST" Then
    '     Print #1, "Tax            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(1).Code
    '  End If
    '  If UCase(gobjEO.Tax(2).Code) <> "GST" Then
    '     Print #1, "Tax            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(2).Code
    '  End If
      If gobjEO.TaxCount > 0 Then
      If UCase(gobjEO.Tax(1).Code) = "GST" Then
         Print #1, "GST            : " & Format(gobjEO.NettGST, gstrAgcyCurrFormat)
      
      'ElseIf UCase(gobjEO.Tax(2).Code) = "GST" Then
      '   Print #1, "GST            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat)
      End If
      If UCase(gobjEO.Tax(1).Code) <> "GST" Then
         Print #1, "Tax            : " & Format(gobjEO.Tax(1).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(1).Code
      End If
      If gobjEO.TaxCount > 1 Then
        If UCase(gobjEO.Tax(2).Code) <> "GST" Then
           Print #1, "Tax            : " & Format(gobjEO.Tax(2).Amount, gstrAgcyCurrFormat) & gobjEO.Tax(2).Code
        End If
      End If
  ' ElseIf gobjEO.TaxCount = 1 Then
  '    If UCase(gobjEO.Tax(1).Code) = "GST" Then
  '       Print #1, "GST            : " & gobjEO.Tax(1).Amount
  '    Else
  '       Print #1, "Tax            : " & gobjEO.Tax(1).Amount & gobjEO.Tax(1).Code
  '    End If
    'Added on 6/4/2005
   'Print #1, "Total          : " & Format(gobjEO.Cost + gobjEO.EOTaxTotal, gstrAgcyCurrFormat)
    If UCase(gobjEO.Tax(1).Code) = "GST" Then
        Print #1, "Total          : " & Format(gobjEO.Cost + gobjEO.NettGST, gstrAgcyCurrFormat)
    Else
        Print #1, "Total          : " & Format(gobjEO.Cost + gobjEO.EOTaxTotal, gstrAgcyCurrFormat)
    End If
   End If
   
   For i = 1 To gobjEO.DescriptionLinesCount
      If i = 1 Then
         Print #1, "Description    : " & gobjEO.DescriptionLine(i)
      Else
         Print #1, "                 " & gobjEO.DescriptionLine(i)
      End If
   Next
   If gobjEO.RemarkCount <> 0 Then
      For i = 1 To gobjEO.RemarkCount
         If i = 1 Then
            Print #1, "Remark         : "
            Print #1, gobjEO.Remark(i)
         Else
            Print #1, gobjEO.Remark(i)
         End If
      Next
   End If
   Print #1, ""
   Print #1, "Please prepare document for our collection today."
   Print #1, "Thank you"
   Close #1
   'MsgBox "Print Successful!", vbOKOnly, "Print Successful"

End Sub




