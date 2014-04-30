VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmEOCrystalReport 
   Caption         =   "CWT Travel Pro - EO Preview"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11400
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CR1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   11895
      lastProp        =   600
      _cx             =   20981
      _cy             =   12726
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.Menu mnuAccept 
      Caption         =   "Confirm"
   End
   Begin VB.Menu mnuCancel 
      Caption         =   "Cancel"
   End
End
Attribute VB_Name = "frmEOCrystalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Application As New CRAXDRT.Application
Dim m_DB As CRAXDRT.Database
Dim m_Report As CRAXDRT.Report
Dim dbTable As CRAXDRT.DatabaseTable
Dim mbolReportPath As Boolean

Dim mstrAgencyName As String
Dim mstrAgencyPhone As String
Dim mstrAgencyFax As String
Dim mstrAgencyAddr As String
Dim mstrNow As String
Dim datFormLoadStart As Date
Dim datFormLoadEnd As Date
Dim datTouchEnd As Date

Private Sub Form_Resize()
   If Me.Height >= 1200 And Me.Width >= 600 Then
      CR1.Height = Me.Height - 1000
      CR1.Width = Me.Width - 400
   End If
End Sub

Private Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    hMenu = GetSystemMenu(hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub


Private Sub Form_Load()
Dim rs_EO As ADODB.Recordset
Dim strSQL As String
Dim intI As Integer
Dim strEONum() As String
Dim strTmp As String
Dim oldParent As Long
    datFormLoadStart = Now
    ' ZhiSam - V1.2.17 20121015 - Smart Point and SyEx server selection
    ' Top most the window so that it would not influence by those form window that set as top most
    SwitchWinSetting (Me.hwnd)

    
    Me.Move 0, 0
    Me.Move frmSideBar.Width, 0
'strID = Split(gEOID, "or")

'strSql = "select * from tblExchangeOrder where "
'For intI = LBound(strID) To UBound(strID)
'    strSql = strSql & IIf(intI > LBound(strID), " or ", "") & "exchangeid =" & strID(intI)
'Next
Set m_Report = Nothing

'Set rs_EO = gdbConn.Execute(strSql)

If gstrReportPath <> "" Then mbolReportPath = True

'm_Application.LogOnServer "P2SODBC.DLL", "wwsng1ujxl299", "tprosg", "travelpro", "travelprodb"

'preethi test quickwins
'If InStr(1, gEOReportName, "SG") Then
'   gEOReportName = "EO-SG_UAT.rpt"
'Else
'   gEOReportName = "EO-HK_UAT.rpt"
'End If

Set m_Report = m_Application.OpenReport(gstrReportPath & gEOReportName, 1)

getReportLogin
'For Each dbTable In m_Report.Database.Tables
'    dbTable.SetLogOnInfo strServer, gstrEODB, strLogin, strPswd
'Next
   For i = 1 To m_Report.Database.Tables.Count
      m_Report.Database.Tables(i).ConnectionProperties("Data Source") = gstrEORptServer
      m_Report.Database.Tables(i).ConnectionProperties("password") = gstrEOPrtPwd
      m_Report.Database.Tables(i).ConnectionProperties("Initial Catalog") = gstrEORptDB
   Next
'Set m_DB = m_Report.Database
'm_DB.SetDataSource rs_EO, 3, 1


'm_Report.SQLQueryString = strSql
strEONum = Split(gEOID, ",")
mstrAgencyName = ""
mstrAgencyPhone = ""
mstrAgencyFax = ""
mstrAgencyAddr = ""
GetAgencyInfo mstrAgencyName, mstrAgencyPhone, mstrAgencyFax, mstrAgencyAddr

For i = 0 To UBound(strEONum)
   If Left(strEONum(i), 1) = "'" Then
      strTmp = Mid(strEONum(i), 2, Len(strEONum(i)) - 2)
   Else
      strTmp = strEONum(i)
   End If
   m_Report.ParameterFields.GetItemByName("EONum").AddCurrentValue (strTmp)
Next
If gbolIndEO Then
   m_Report.ParameterFields.GetItemByName("AgencyName").AddCurrentValue (mstrAgencyName)
   m_Report.ParameterFields.GetItemByName("AgencyAddress").AddCurrentValue (mstrAgencyAddr)
   m_Report.ParameterFields.GetItemByName("AgencyPhone").AddCurrentValue (mstrAgencyPhone)
   m_Report.ParameterFields.GetItemByName("AgencyFax").AddCurrentValue (mstrAgencyFax)
End If

If gbolRaiseEOReport = False Then
   If frmExchangeOrder.mbolPreview = False Then
      mnuAccept_Click
      
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datFormLoadStart

      
      
      
      Exit Sub
   End If
End If

CR1.ReportSource = m_Report
'If gEOID <> "" Then
'
'    m_Report.RecordSelectionFormula = "{tblExchangeOrder.ExchangeID} in [" & gEOID & "]"
'End If

 

CR1.ViewReport

    RemoveMenus Me, False, False, _
        False, False, False, True, True
        
    datFormLoadEnd = Now
    If gstrProcessGrpID = "" Then gstrProcessGrpID = pGetProcessKey

End Sub

Public Sub getReportLogin()
    'Dim INITFILE As String
    'Dim fnum As Integer
    'Dim textline As String
        
    
    'INITFILE = gstrDBInitPath
    
    'fnum = FreeFile
    'Open INITFILE For Input As #fnum
    'Do While Not EOF(fnum)
    '    Line Input #fnum, textline
    '    Select Case UCase(Left(textline, InStr(textline, "=") - 1))
    '        Case "EOREPORTLOGIN":
    '            strLogin = Mid(Trim(textline), InStr(Trim(textline), "=") + 1)
    '        Case "EOREPORTPSWD"
    '            strPswd = Mid(Trim(textline), InStr(Trim(textline), "=") + 1)
    '        Case "DATABASE"
    '            strDb = Mid(Trim(textline), InStr(Trim(textline), "=") + 1)
    '        Case "SERVER"
    '            strServer = Mid(Trim(textline), InStr(Trim(textline), "=") + 1)
    '    End Select
    'Loop
    'Close #fnum
    
   Dim strPath As String
   Dim strINIFile As String
   
   'strINIFile = GetSetting("CWTAPP", "APPINI", "FileLoc", "")
   'If strINIFile = "" Then
   '   MsgBox "INI file location info not in registry"
   '   End
   'End If
   strPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
   strINIFile = strPath & "cwtapplication.ini"
   If Dir(strINIFile) = "" Then
      MsgBox "INI file: " & strINIFile & " not found."
      End
   End If
   

   'EO Report Details

   
   'strLogin = GetFromINI("EOReport", "EOReportLogin", strINIFile)
   'strLogin = decrypt(strLogin, gintKey)
 
   'strPswd = GetFromINI("EOReport", "EOReportPswd", strINIFile)
   'strPswd = decrypt(strPswd, gintKey)
   
   'strDb = GetFromINI("EOReport", "EODatabase", strINIFile)
   'strDb = decrypt(strDb, gintKey)
   
   'strServer = GetFromINI("EOReport", "EOServer", strINIFile)
   'strServer = decrypt(strServer, gintKey)
End Sub

Private Sub mnuAccept_Click()
   Dim strEOFile As String
   Dim strEmailEOFile As String
   Dim strSQL As String
   'Dim strSubject As String
   'Dim strBody As String
   
   'Dim strAgentEmail As String
   datTouchEnd = Now
   bolAcceptEO = True
   mstrNow = ""
   
   'If UCase(Right(Me.Caption, 5)) = "EMAIL" Then
   '   mstrNow = Format(Now, "DDMMhhmmss")
   '   strEOFile = gstrEOEmailPath & gobjPNR.RecLoc & "_" & gobjEO.EONumber & "_" & mstrNow & ".pdf"
   'ElseIf UCase(Right(Me.Caption, 5)) = "E-FAX" Then
   '   mstrNow = Format(Now, "hhmmss")
   '   strEOFile = gstrEOFaxPath & gobjPNR.RecLoc & "_" & gobjEO.EONumber & "_" & mstrNow & ".pdf"
   'End If
   '
   ''m_Report.RecordSelectionFormula = "{tblExchangeOrder.ExchangeID} in [" & gEOID & "]"
   'm_Report.ExportOptions.DestinationType = crEDTDiskFile
   'm_Report.ExportOptions.FormatType = crEFTPortableDocFormat
   'm_Report.ExportOptions.DiskFileName = strEOFile
   'm_Report.Export False
   
   'If gbolPreviewEO Then
   If gbolRaiseEOReport = False Then
      Me.Caption = "EO Preview - " & frmExchangeOrder.mstrEOType
   
   
       mstrNow = Format(Now, "DDMMhhmmss")
       If InStr(1, UCase(Me.Caption), "E-FAX") > 0 Then
          strEOFile = gstrEOFaxPath & gobjPNR.RecLoc & "_" & gobjEO.EONumber & "_" & mstrNow & ".pdf"
          m_Report.ExportOptions.DestinationType = crEDTDiskFile
          m_Report.ExportOptions.FormatType = crEFTPortableDocFormat
          m_Report.ExportOptions.DiskFileName = strEOFile
          m_Report.Export False
       End If
       If InStr(1, UCase(Me.Caption), "EMAIL") > 0 Then
          strEmailEOFile = gstrEOEmailPath & gobjPNR.RecLoc & "_" & gobjEO.EONumber & "_" & mstrNow & ".pdf"
          m_Report.ExportOptions.DestinationType = crEDTDiskFile
          m_Report.ExportOptions.FormatType = crEFTPortableDocFormat
          m_Report.ExportOptions.DiskFileName = strEmailEOFile
          m_Report.Export False
       End If
    
       
       If InStr(1, UCase(Me.Caption), "EMAIL") > 0 Then
          SendEOEmail strEmailEOFile
       End If
       If InStr(1, UCase(Me.Caption), "E-FAX") > 0 Then
          'If UCase(gstrAgcyCountryCode) = "SG" Then
             'SendEOFax
          'Else
             SendEOEFax strEOFile
          'End If
       End If
   End If
   
   
   
   If InStr(1, UCase(Me.Caption), "PRINT") > 0 Or mnuAccept.Caption = "Print" Then
      m_Report.PrinterSetup 0
      m_Report.PrintOut False
   End If
   'strAgentEmail = GetAgentEmail(gobjPNR.Agent, gobjPNR.PCCOwner)
   'GetEOSubjectBody strSubject, strBody
   '
   'strBody = Replace(strBody, "{agentname}", gobjEO.CreatedByName)
   'strBody = Replace(strBody, "{directline}", gstrAgcyPhone)
   'strBody = Replace(strBody, "{agentemail}", strAgentEmail)
   'strBody = Replace(strBody, "{companyphone}", mstrAgencyPhone)
   'strBody = Replace(strBody, "{companyfax}", mstrAgencyFax)
   'strBody = Replace(strBody, "{companyaddress}", mstrAgencyAddr)
   '
   'With gobjEO
   '   SendEmail .Email, "", "", .CreatedByName, strAgentEmail, _
   '             .CreatedByName, strSubject, strBody, True, Date, "EO", _
   '             strEOFile, Time, gstrAgcyCountryCode, gobjPNR.RecLoc
   '   'strSql = "Insert into tblEmail (STo, CC, BCC, SFrom, SenderEmail, "
   '   'strSql = strSql & "SenderName, Subject, Body, HTML, SDate, Type, "
   '   'strSql = strSql & "Attachment, STime, Country, PNR) "
   '   'strSql = strSql & "Values('" & Replace(.Email, "'", "''") & "' "
   '   'strSql = strSql & ",'' "
   '   'strSql = strSql & ",'' "
   '   'strSql = strSql & ",'" & Replace(gobjEO.CreatedByName, "'", "''") & "' "
   '   'strSql = strSql & ",'" & Replace(strAgentEmail, "'", "''") & "' "
   '   'strSql = strSql & ",'" & Replace(gobjEO.CreatedByName, "'", "''") & "' "
   '   'strSql = strSql & ",'" & Replace(strSubject, "'", "''") & "' "
   '   'strSql = strSql & ",'" & Replace(strBody, "'", "''") & "' "
   '   'strSql = strSql & "," & "1" & " "
   '   'strSql = strSql & ",'" & Format(Date, "dd/MMM/yyyy") & "' "
   '   'strSql = strSql & ",'" & "EO" & "' "
   '   'strSql = strSql & ",'" & strEOFile & "' "
   '   'strSql = strSql & ",'" & Format(Time, "HH:mm:ss") & "' "
   '   'strSql = strSql & ",'" & gstrAgcyCountryCode & "' "
   '   'strSql = strSql & ",'" & gobjPNR.RecLoc & "' "
   '   'strSql = strSql & ")"
   'End With
   ''gdbEmailConn.Execute strSql
   If frmExchangeOrder.mbolPreview Then
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconFormLoad, gstrProcessGrpID, datFormLoadEnd, datFormLoadStart
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconTouch, gstrProcessGrpID, datTouchEnd, datFormLoadEnd
       
       pEndProcessTimeLog IIf(gobjPNR.CN <> "", gobjPNR.CN, ""), frmSideBar.cmbSelectReq.Text, _
       gconModOthServ, frmSideBar.cmbSelectType.Text, gconSModAux, _
       Me.Name, gconProcessing, gstrProcessGrpID, , datTouchEnd
      Unload Me
   End If
   
End Sub

Private Sub SendEOEmail(EOFile As String)
   Dim strSubject As String
   Dim strBody As String
   Dim strAgentEmail As String
   Dim objName As PNRName
   Dim i As Integer
   'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
   Dim strMsg As String
   Dim StrEmailSenderName As String
   Dim strEmailType As String
   '--
   
   'strAgentEmail = GetAgentEmail(gobjPNR.Agent, gobjPNR.PCCOwner)
   'strAgentEmail = GetAgentEmail(IIf(gobjPNR.Agent = "", gobjHost.AgentSine, gobjPNR.Agent), gobjPNR.PCCOwner)
   'strAgentEmail = GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjHost.AgentSine, gobjHost.AgentPcc, , True)
   strAgentEmail = gobjEO.ReplyEmail
   
   GetEOSubjectBody strSubject, strBody
   
   
   strSubject = Replace(strSubject, "{PNR}", UCase(gobjPNR.RecLoc))
   strSubject = Replace(strSubject, "{EONum}", UCase(gobjEO.EONumber))
   Set objName = gobjPNR.PassengerName(1)
   strSubject = Replace(strSubject, "{PxName}", UCase(objName.LastName & "/" & objName.FirstName))
   If gobjPNR.AirSegCount > 0 Then
      strSubject = Replace(strSubject, "{SegInfo}", UCase(", " & gobjPNR.AirSeg(1).DepartAirport & "-" & gobjPNR.AirSeg(1).ArriveAirport))
      If (IsDate(gobjPNR.AirSeg(1).DepartDateTime)) And (gobjPNR.AirSeg(1).FlightNumber <> "OPEN") Then
         strSubject = Replace(strSubject, "{strFirstDepDate}", ", " & UCase(Format(gobjPNR.AirSeg(1).DepartDateTime, "DDMMM")))
      Else
         strSubject = Replace(strSubject, "{strFirstDepDate}", "")
      End If
   Else
      strSubject = Replace(strSubject, "{SegInfo}", "")
      strSubject = Replace(strSubject, "{strFirstDepDate}", "")
   End If
   
   strBody = Replace(strBody, "{agentname}", gobjEO.CreatedByName)
   strBody = Replace(strBody, "{directline}", gstrAgcyPhone)
   strBody = Replace(strBody, "{agentemail}", strAgentEmail)
   strBody = Replace(strBody, "{companyphone}", mstrAgencyPhone)
   strBody = Replace(strBody, "{companyfax}", mstrAgencyFax)
   strBody = Replace(strBody, "{companyaddress}", mstrAgencyAddr)
   strBody = Replace(strBody, "{replysubject}", Mid(strSubject, 5) & " - ATTN: " & gobjEO.CreatedByName)
   

   
   'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
     If bolAgencyNameCheck(gobjPNR.CompInfo.AgencyName) = False Then
        Exit Sub
     Else
        StrEmailSenderName = getDefaultEmailSenderName(gobjPNR.CompInfo.AgencyName)
       
        'Production - use this when deployment
        strEmailType = "EO"
        'UAT - use this when UAT
        'strEmailType = "EO_HKSG"
    '--
        With gobjEO
           'SendEmail .Email, "", "", .CreatedByName, strAgentEmail, _
                     .CreatedByName, strSubject, strBody, True, Date, "EO", _
                     EOFile, Time, gstrAgcyCountryCode, gobjPNR.RecLoc
         'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
            SendEmail .Email, "", "", .CreatedByName, strAgentEmail, _
             StrEmailSenderName, strSubject, strBody, True, Date, strEmailType, _
             EOFile, Time, gstrAgcyCountryCode, gobjPNR.RecLoc
          '--
        End With
     End If
End Sub

Private Sub SendEOEFax(EOFile As String)
    Dim strAgentEmail As String
    Dim strEFaxBody As String
    Dim strEFaxEmail As String
    Dim strEFaxSubject As String
    Dim strEFaxSender As String
    Dim strRecipientName As String
   EFaxInfo strEFaxEmail, strEFaxSubject, strEFaxBody, strEFaxSender
   'strAgentEmail = GetAgentEmail(gobjPNR.Agent, gobjPNR.PCCOwner)
   'strAgentEmail = GetAgentEmail(IIf(gobjPNR.Agent = "", gobjHost.AgentSine, gobjPNR.Agent), gobjPNR.PCCOwner)
   'strAgentEmail = GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjHost.AgentSine, gobjHost.AgentPcc, , True)
   strAgentEmail = gobjEO.ReplyEmail
   If UCase(gstrAgcyCountryCode) = "HK" Then
   
        strEFaxSubject = Replace(strEFaxSubject, "{faxnum}", gobjEO.FaxNo)
   Else
         strEFaxEmail = Replace(strEFaxEmail, "{faxnum}", gobjEO.FaxNo)
         strEFaxSubject = Replace(strEFaxSubject, "{email}", gobjEO.ReplyEmail)
         strRecipientName = gobjEO.ContactPerson
         
         If strRecipientName <> "" Then
            
            If Len(strRecipientName) > 40 Then
                    strRecipientName = Left(strRecipientName, 40)
            End If
            strRecipientName = Replace(strRecipientName, " ", "_")
            strEFaxEmail = Replace(strEFaxEmail, "{name}", strRecipientName)
         Else
            strEFaxEmail = Replace(strEFaxEmail, "{name}.", strRecipientName)
         End If
         
   End If
   With gobjEO
      SendEmail strEFaxEmail, "", "", .CreatedByName, strEFaxSender, _
                .CreatedByName, strEFaxSubject, strEFaxBody, False, Date, "EOFax", _
                EOFile, Time, gstrAgcyCountryCode, gobjPNR.RecLoc
   End With
End Sub

Private Sub SendEOFax()
   Dim strFaxTextName As String
   Dim strFaxNum() As String
   Dim i As Integer
   
   If gstrAgcyCountryCode = "SG" Then
            
      strFaxNum = Split(gobjEO.FaxNo, ",")
      For i = 0 To UBound(strFaxNum)
            
        strFaxTextName = gstrEOFaxPath & gobjPNR.RecLoc & "_" & gobjEO.EONumber & "_" & mstrNow & "_" & i + 1 & ".txt"
        Open strFaxTextName For Output As #1
        
        Print #1, gobjPNR.RecLoc & "_" & gobjEO.EONumber & "_" & mstrNow & ".pdf"
        Print #1, strFaxNum(i)
        Print #1, "ExOrder"
        Print #1, "ExOrder"
        Print #1, "ExOrder"
        Close #1
      Next
   Else
   
   End If
End Sub

Private Sub GetEOSubjectBody(ByRef Subject As String, ByRef Body As String)
   Dim strSQL As String
   Dim rs As ADODB.Recordset
   
   'strSQL = "Select * from tblEmailFormat "
   'strSQL = strSQL & "Where DocType = '" & "EO_Email" & "' "
   
   'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
    strSQL = "Select * from tblEmailFormat "
    strSQL = strSQL & "Where DocType = '" & "EO_Email_HKSG" & "' "
    strSQL = strSQL & "and AgencyName = '" & gobjPNR.CompInfo.AgencyName & "' "
   '--

   
   Set rs = gdbEitinConn.Execute(strSQL)
   
   If Not rs.EOF Then
      Subject = rs!Subject & ""
      Body = rs!Body & ""
   End If
   rs.Close
   Set rs = Nothing
End Sub

Private Sub GetAgencyInfo(ByRef FullName As String, ByRef Phone As String, ByRef Fax As String, ByRef Address As String)
   Dim strSQL As String
   Dim rs As ADODB.Recordset
   Dim strPCC As String
   
   strPCC = Trim(gobjPNR.PCCOwner)
   If Len(strPCC) = 3 Then
      strPCC = "0" & strPCC
   End If
   
   strSQL = "Select * from tblAgency "
   strSQL = strSQL & "Where AgencyPCC = '" & strPCC & "' "
   
   Set rs = gdbEitinConn.Execute(strSQL)
   
   If Not rs.EOF Then
      FullName = rs!AgencyFullName & ""
      Phone = rs!AgencyPhone & ""
      Fax = rs!AgencyFax & ""
      Address = rs!AgencyAddr & ""
   End If
   rs.Close
   Set rs = Nothing
End Sub

Private Sub mnuCancel_Click()
   gbolAcceptEO = False
   If gbolPreviewEO Then
      If gbolEOAmend Then
         UpdateExchOrdDB gobjPreEO, gstrPreEOType
      Else
         RemoveEO gEOID
      End If
      gbolPreviewCancel = True
   End If
   Unload Me
End Sub

Private Sub RemoveEO(EOID As String)
   Dim rs As ADODB.Recordset
   Dim strSQL As String
   
   strSQL = "Delete from tblExchangeOrder "
   strSQL = strSQL & "Where ExchangeID = " & EOID & " "
   
   gdbConn.Execute strSQL
   
End Sub

'20070522
Private Sub EFaxInfo(ByRef EFaxEmail As String, ByRef EFaxSubject As String, _
                     ByRef EFaxBody As String, ByRef EFaxSender As String)
   Dim rs As ADODB.Recordset
   Dim strSQL As String
   
   EFaxEmail = ""
   EFaxSubject = ""
   EFaxBody = ""
   EFaxSender = ""
   
   strSQL = "Select OptionValue from tblModOptions "
   strSQL = strSQL & "Where OptionCode = '" & "EFaxEmail" & "' "
   Set rs = gdbConn.Execute(strSQL)
   If rs.EOF = False Then
      EFaxEmail = rs!optionvalue & ""
   End If
   rs.Close
   Set rs = Nothing
   
   strSQL = "Select OptionValue from tblModOptions "
   strSQL = strSQL & "Where OptionCode = '" & "EFaxSubject" & "' "
   Set rs = gdbConn.Execute(strSQL)
   If rs.EOF = False Then
      EFaxSubject = rs!optionvalue & ""
   End If
   rs.Close
   Set rs = Nothing
   
   strSQL = "Select OptionValue from tblModOptions "
   strSQL = strSQL & "Where OptionCode = '" & "EFaxBody" & "' "
   Set rs = gdbConn.Execute(strSQL)
   If rs.EOF = False Then
      EFaxBody = rs!optionvalue & ""
   End If
   rs.Close
   Set rs = Nothing
   
   strSQL = "Select OptionValue from tblModOptions "
   strSQL = strSQL & "Where OptionCode = '" & "EFaxSender" & "' "
   strSQL = strSQL & "and OptionSecCode = '" & gobjHost.AgentDIV & "' "
   Set rs = gdbConn.Execute(strSQL)
   If rs.EOF = False Then
      EFaxSender = rs!optionvalue & ""
   End If
   rs.Close
   Set rs = Nothing
   
   If EFaxSender = "" Then
      strSQL = "Select OptionValue from tblModOptions "
      strSQL = strSQL & "Where OptionCode = '" & "EFaxSender" & "' "
      strSQL = strSQL & "and OptionKey = '" & "1" & "' "
      Set rs = gdbConn.Execute(strSQL)
      If rs.EOF = False Then
         EFaxSender = rs!optionvalue & ""
      End If
   End If
   rs.Close
   Set rs = Nothing
   
End Sub

