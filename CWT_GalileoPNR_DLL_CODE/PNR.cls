VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PNR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SimiliarName"
Attribute VB_Ext_KEY = "Member1" ,"OSI"
Attribute VB_Ext_KEY = "Member2" ,"SSR"
Attribute VB_Ext_KEY = "Member3" ,"CarSeg"
Attribute VB_Ext_KEY = "Member4" ,"AirFaresFOP"
Attribute VB_Ext_KEY = "Member5" ,"Phone"
Attribute VB_Ext_KEY = "Member6" ,"PaidDue"
Attribute VB_Ext_KEY = "Member7" ,"ItinRemark"
Attribute VB_Ext_KEY = "Member8" ,"AcctRemark"
Attribute VB_Ext_KEY = "Member9" ,"TurSeg"
Attribute VB_Ext_KEY = "Member10" ,"HotelSeg"
Attribute VB_Ext_KEY = "Member11" ,"FiledFare"
Attribute VB_Ext_KEY = "Member12" ,"GASalesRecord"
Option Explicit

Dim mPaxFF As PxFF
Private mxmldomPNR As MSXML2.DOMDocument
Private mxmlnl1 As MSXML2.IXMLDOMNodeList
Private mxmlnl2 As MSXML2.IXMLDOMNodeList
Private mstrHeader As String 'local copy
Private mstrRecLoc As String 'local copy
Private mdtmCreateDate As Date 'local copy
Private mstrAgent As String 'local copy
Private mstrPCCOwner As String 'local copy
Private mstrFileAddr As String 'local copy
Private mstrCodeCheck As String 'local copy
Private mbolInUse As Boolean 'local copy
Private mbolTktNumExists As Boolean 'local copy
Private mbolFareDataExists As Boolean 'local copy
Private mintFiledFareCount As Integer
Private mbolIMUDataExists As Boolean 'local copy
Private mbolETktDataExists As Boolean 'local copy
Private mstrBillAddress As String
Private mstrDelivAddress As String
Private mstrFOP As String
Private mstrCN As String
Private mcolSimiliarNames As Collection
Private mcolNames As Collection
Private mcolPhone As Collection
Private mcolAirSegs As Collection
Private mcolHtlSegs As Collection
Private mcolTurSegs As Collection 'TUR
Private mcolCarSegs As Collection
Private mcolPaidDue As Collection
Private mcolGenRmks As Collection
Private mcolAcctRmks As Collection
'added on 16/08/04 - File Fare FOP in DI lines
Private mcolAirFaresFOP As Collection
Private mcolGASalesRecords As Collection
Private mcolItinRmks As Collection
Private mcolSeatData As Collection
Private mcolFreqCustNums As Collection
Private mcolQMinders As Collection
Private mcolFiledFares As Collection
Private mcolVRecLocs As Collection
Private mcolSSRequests As Collection
Private mcolOSI As Collection
Private mvarFOPType As String 'local copy
Private mvarFOP_CCNum As String 'local copy
Private mvarFOP_CCExpireDate As Date 'local copy
Private mvarFOP_CCCode As String 'local copy
Private mvarFOP_AddText As String 'local copy
Private mvarCompInfo As CompanyInfo 'local copy
Private mlngPNRStatus As Long
Private mdteTktDate As String
Private mblnTkted As Boolean
'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Private mcolETicket As Collection

'ZhiSam - V1.2.19 20120311 - CR-203 - Add TAW xml segment
Private mstrTAWDate As String
'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
Private mvarDSInfo As DSInfo



Public Enum cwtPNRStatus
    cwtPNRNone = -1
    cwtPNRNew = 0
    cwtPNREnded = 1
    cwtPNRInUse = 2
    cwtPNRSimulChange = 3
End Enum

Public Property Get CompInfo() As CompanyInfo
    Set CompInfo = mvarCompInfo
End Property

Public Property Let FOP_AddText(ByVal vData As String)
    mvarFOP_AddText = vData
End Property

Public Property Get FOP_AddText() As String
    FOP_AddText = mvarFOP_AddText
End Property

Public Property Let FOP_CCCode(ByVal vData As String)
    mvarFOP_CCCode = vData
End Property

Public Property Get FOP_CCCode() As String
    FOP_CCCode = mvarFOP_CCCode
End Property

Public Property Let FOP_CCExpireDate(ByVal vData As Date)
    mvarFOP_CCExpireDate = vData
End Property

Public Property Get FOP_CCExpireDate() As Date
    FOP_CCExpireDate = mvarFOP_CCExpireDate
End Property

Public Property Let FOP_CCNum(ByVal vData As String)
    mvarFOP_CCNum = vData
End Property

Public Property Get FOP_CCNum() As String
    FOP_CCNum = mvarFOP_CCNum
End Property

Public Property Let FOPType(ByVal vData As String)
    mvarFOPType = vData
End Property

Public Property Get FOPType() As String
    FOPType = mvarFOPType
End Property

Public Sub GenRmkAddMultiple(RmkTextArray() As String, Optional Qualifier As String, Optional InsertAfter As Integer = -1)
End Sub

Public Sub GenRmkAdd(RmkText As String, Optional Qualifier As String, Optional InsertAfter As Integer = -1)
Dim objNewGR As GenRemark
Dim lngC1 As Long
Dim lngC2 As Long
Dim strRequest As String
Dim strResponse As String

strRequest = "NP." & IIf(InsertAfter >= 0, "/" & CStr(InsertAfter), "") _
    & IIf(Qualifier <> "", Qualifier & "*", "") & RmkText
strResponse = gobjHost.TerminalEntry(strRequest)

Set objNewGR = New GenRemark
With objNewGR
    .ItemNum = lngC1
    .Qualifier = Qualifier
    .RemarkText = RmkText
End With

mcolGenRmks.Add objNewGR, , , IIf(InsertAfter = -1, mcolGenRmks.Count, InsertAfter)

End Sub

Public Property Let ETktDataExists(ByVal vData As Boolean)
    mbolETktDataExists = vData
End Property

Public Property Get ETktDataExists() As Boolean
    ETktDataExists = mbolETktDataExists
End Property

Public Property Let IMUDataExists(ByVal vData As Boolean)
    mbolIMUDataExists = vData
End Property

Public Property Get IMUDataExists() As Boolean
    IMUDataExists = mbolIMUDataExists
End Property

Public Property Let CN(ByVal vData As String)
    mstrCN = vData
End Property

Public Property Get CN() As String
    CN = mstrCN
End Property

Public Property Let FareDataExists(ByVal vData As Boolean)
    mbolFareDataExists = vData
End Property

Public Property Get FareDataExists() As Boolean
    FareDataExists = mbolFareDataExists
End Property

Public Property Let TktNumExists(ByVal vData As Boolean)
    mbolTktNumExists = vData
End Property

Public Property Get TktNumExists() As Boolean
    TktNumExists = mbolTktNumExists
End Property

Public Property Let CodeCheck(ByVal vData As String)
    mstrCodeCheck = vData
End Property

Public Property Get CodeCheck() As String
    CodeCheck = mstrCodeCheck
End Property

Public Property Let FileAddr(ByVal vData As String)
    mstrFileAddr = vData
End Property

Public Property Get FileAddr() As String
    FileAddr = mstrFileAddr
End Property

Public Property Let PCCOwner(ByVal vData As String)
    mstrPCCOwner = vData
End Property

Public Property Get PCCOwner() As String
    PCCOwner = mstrPCCOwner
End Property

Public Property Let Agent(ByVal vData As String)
    mstrAgent = vData
End Property

Public Property Get Agent() As String
    Agent = mstrAgent
End Property

Public Property Let CreateDate(ByVal vData As Date)
    mdtmCreateDate = vData
End Property

Public Property Get CreateDate() As Date
    CreateDate = mdtmCreateDate
End Property

Public Property Get RecLoc() As String
    RecLoc = mstrRecLoc
End Property

Public Property Get VendorRecLoc(Index As Long) As VRecLoc
    Set VendorRecLoc = mcolVRecLocs(Index)
End Property

Public Property Get VendorRecLocCount() As Long
     VendorRecLocCount = mcolVRecLocs.Count
End Property

Public Property Get GASaleRecordCount() As Long
    GASaleRecordCount = mcolGASalesRecords.Count
    
End Property

Public Function GASalesRecord(ByVal Index As Long) As GASalesRecord
     Set GASalesRecord = mcolGASalesRecords.Item(Index)
End Function

Public Property Get AirSegCount() As Long
    If Not mcolAirSegs Is Nothing Then
        AirSegCount = mcolAirSegs.Count
    Else
        AirSegCount = 0
    End If
    
End Property

Public Function AirSegAdd(AirSegString() As String, Optional NumberOfSeats As Integer = 1, _
    Optional AvailType As String = "G", Optional Status As String = "NN") As String
Dim strSegment() As String
Dim strRequest As String
Dim strResponse As String
Dim strResult As String
Dim strResp As String
Dim lngC1 As Long
Dim lngC2 As Long
Dim intJ As Integer 'used to track journey number for connecting flights
Dim xmlResponse As MSXML2.DOMDocument
Dim xmlnlTemp1 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp2 As MSXML2.IXMLDOMNodeList

intJ = 0 'defalult value

strRequest = "<AirSegmentSell_6_0>"
strRequest = strRequest & "    <AirSegSellMods>"
For lngC1 = 0 To UBound(AirSegString)
    If AirSegString(lngC1) = "" Then Exit For
    strRequest = strRequest & "            <AirSegSell>"
    strSegment = Split(AirSegString(lngC1), "/")
    strRequest = strRequest & "                <Vnd>" & strSegment(0) & "</Vnd>"
    strRequest = strRequest & "                <FltNum>" & Format(strSegment(1), "0000") & "</FltNum>"
    strRequest = strRequest & "                <Class>" & strSegment(2) & "</Class>"
    strRequest = strRequest & "                <StartDt>" & strSegment(3) & "</StartDt>"
    strRequest = strRequest & "                <StartAirp>" & strSegment(4) & "</StartAirp>"
    strRequest = strRequest & "                <EndAirp>" & strSegment(5) & "</EndAirp>"
    strRequest = strRequest & "                <Status>" & Status & "</Status>"
    strRequest = strRequest & "                <NumPsgrs>" & CStr(NumberOfSeats) & "</NumPsgrs>"
    strRequest = strRequest & "                <AvailDispType>" & AvailType & "</AvailDispType>"
    
    If strSegment(6) = "N" Then intJ = intJ + 1
    strRequest = strRequest & "                <AvailJrnyNum>" & CStr(intJ) & "</AvailJrnyNum>"
    strRequest = strRequest & "            </AirSegSell>"
Next lngC1
strRequest = strRequest & "    </AirSegSellMods>"
strRequest = strRequest & "</AirSegmentSell_6_0>"

Set xmlResponse = New MSXML2.DOMDocument

strResp = gobjHost.SendQuery(strRequest, "AirSegmentSell_6_0", "GalileoPNR.PNR", "AirSegAdd")
xmlResponse.loadXML strResp
Set xmlnlTemp1 = xmlResponse.selectNodes("//AirSegSell/AirSell|//AirSegSell/TextMessage/Text")

For lngC1 = 0 To xmlnlTemp1.length - 1
    Select Case xmlnlTemp1(lngC1).nodeName
        Case "AirSell"
            strResult = strResult & Chr(13)
            Set xmlnlTemp2 = xmlnlTemp1(lngC1).childNodes
                For lngC2 = 0 To xmlnlTemp2.length - 1
                    With xmlnlTemp2(lngC2)
                        Select Case .nodeName
                            Case "Vnd", "FltNum", "Class", "StartAirp", "EndAirp"
                                strResult = strResult & .Text
                            Case "StartDt"
                                strResult = strResult & UCase(Format(ConvertDate(.Text), "ddmmm")) & " "
                            Case "StartTm", "EndTm"
                                strResult = strResult & " " & .Text
                            Case "SuccessInd"
                                strResult = strResult & IIf(.Text = "N", " NOT ", " ") & "SOLD"
                        End Select
                    End With
                Next lngC2
        Case "Text"
            strResult = strResult & Chr(13) & "   " & xmlnlTemp1(lngC1).Text
    End Select
Next

AirSegAdd = strResult

End Function

Public Property Let Header(ByVal vData As String)
    mstrHeader = vData
End Property

Public Property Get Header() As String
    Header = mstrHeader
End Property

Private Sub Class_Initialize()

Set mvarCompInfo = New CompanyInfo
'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Set mcolETicket = New Collection
'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
Set mvarDSInfo = New DSInfo

End Sub
Public Sub pGetStartupValues(ByRef Conn As ADODB.Connection, ByRef objLog As CWT_AppLog.AppLog, ByRef host As CWT_Galileo3.GalileoHost, ByVal VPHwnd As Long)

'gstrDBInitPath = UCase(GetSetting("TPro", "Startup", "DBInit", "NOT FOUND"))
'gstrAgcyCountryCode = GetSetting("TPro", "Startup", "CountryCode")
'gstrAgcyCurrencyCode = GetSetting("TPro", "Startup", "CurrencyCode")
'GstrAgcyCityCode = GetSetting("TPro", "Startup", "CityCode")
'gstrAgcyAirportCode = GetSetting("TPro", "Startup", "CityCode")
'gstrAgcyPhone = GetSetting("TPro", "Startup", "AgcyPhone")
'gstrPFCode = GetSetting("TPro", "Startup", "PFCode")
'GstrPCC = GetSetting("TPro", "Startup", "PCC")
'gbolSvcBur = GetSetting("TPro", "Startup", "SvcBur", "False")
'gstrHQPCC = GetSetting("TPro", "Startup", "HQ_PCC")
'gstrPFPCC = GetSetting("TPro", "Startup", "PF_PCC")
'gbolSwitchAcc = GetSetting("TPro", "Startup", "SwitchAcc", "False")
    'Added on 21/03/2005
     
    'Set gdbConn = New ADODB.Connection
    'GetALLPath
    
    'If gstrConn <> "" Then
    If Conn.State = 1 Then
        Set gdbConn = Conn
    Else
        Conn.open
        Set gdbConn = Conn
        gdbConn.open
    End If
    
    If objLog Is Nothing Then
       Set gobjLog = New CWT_AppLog.AppLog
       gobjLog.OpenLog App.Path, App.EXEName, App.Title, App.Major, App.Minor, App.Revision
    Else
       Set gobjLog = objLog
    End If
    
    If host Is Nothing Then
       Set gobjHost = New CWT_Galileo3.GalileoHost
    Else
       Set gobjHost = host
    End If
    gVPMDIHwnd = VPHwnd
    pGetNewStartupValues
    
    
        'gdbConn.ConnectionString = "DRIVER={SQL SERVER};SERVER=CWTSIN03\CWTSINSQL01;" & _
        '                            "Database=TProSG;UID=travelpro;PWD=travelprodb;"
        'gdbConn.open
    'If getConnStr(gdbConn) Then
    '    gdbConn.open
    'Call getErrorConfig
            'Else
            '    MsgBox "Error in Database Connection String!", vbCritical
    'End If
    
    
    
End Sub

Private Sub Class_Terminate()

Set mcolNames = Nothing
Set mcolPhone = Nothing
Set mcolAirSegs = Nothing
Set mcolHtlSegs = Nothing
Set mcolTurSegs = Nothing   'Tur
Set mcolCarSegs = Nothing
Set mcolPaidDue = Nothing
Set mcolGenRmks = Nothing
Set mcolAcctRmks = Nothing
Set mcolGASalesRecords = Nothing
Set mcolItinRmks = Nothing
Set mcolFreqCustNums = Nothing
Set mcolQMinders = Nothing
Set mcolVRecLocs = Nothing
Set mcolFiledFares = Nothing
Set mcolAirFaresFOP = Nothing
Set mvarCompInfo = Nothing
Set mcolSeatData = Nothing
'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Set mcolETicket = Nothing
'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
Set mvarDSInfo = Nothing
End Sub

Public Function CheckPNRStatus() As cwtPNRStatus
Dim strRequest As String
Dim xmldomResponse As MSXML2.DOMDocument

    strRequest = ""
    strRequest = strRequest & "<AgencyPNRBFDisplay_7_5>"
    strRequest = strRequest & "    <PNRBFRetrieveMods>"
    strRequest = strRequest & "        <CurrentPNR/>"
    strRequest = strRequest & "    <RequiredData>"
    strRequest = strRequest & "        <PNRDataIDAry>"
    strRequest = strRequest & "            <PNRDataID>GI</PNRDataID>"
    strRequest = strRequest & "        </PNRDataIDAry>"
    strRequest = strRequest & "    </RequiredData>"
    strRequest = strRequest & "    </PNRBFRetrieveMods>"
    strRequest = strRequest & "</AgencyPNRBFDisplay_7_5>"
    
    Set xmldomResponse = New MSXML2.DOMDocument
    
    xmldomResponse.loadXML gobjHost.SendQuery(strRequest, "AgencyPNRBFDisplay_7_5", "GalileoPNR.PNR", "CheckPNRStatus")

If xmldomResponse.getElementsByTagName("SimultaneousUpdInd").Item(0).Text = "Y" Then
    CheckPNRStatus = 3
ElseIf xmldomResponse.getElementsByTagName("InUseInd").Item(0).Text = "Y" Then
    CheckPNRStatus = 2
ElseIf xmldomResponse.getElementsByTagName("RecLoc").Item(0).Text <> "" Then
    CheckPNRStatus = 1
ElseIf InStr(gobjHost.TerminalEntry("*R"), "NO B.F. TO DISPLAY - ") Then
    CheckPNRStatus = -1
Else
    CheckPNRStatus = 0
End If

mlngPNRStatus = CheckPNRStatus

End Function

Public Property Get PNRStatus() As cwtPNRStatus
    PNRStatus = mlngPNRStatus
End Property

Public Function LoadPNR(Optional RecLoc As String, Optional strName As String, Optional strPCC As String, Optional bolAllPCC As Boolean, Optional strAirline As String, Optional strFlightNum As String, Optional strDepartureDt As String, Optional strBoardPoint As String) As String
Dim intC As Integer
Dim intD As Integer
Dim strTemp As String
Dim xmlnlSNL As MSXML2.IXMLDOMNodeList
Dim xmlnlSNLChilds As MSXML2.IXMLDOMNodeList
Dim objSimiliarName As SimiliarName
Dim strMsg As String

strTemp = ""
strTemp = strTemp & "<PNRBFManagement_11>"
strTemp = strTemp & " <PNRBFRetrieveMods>"
If RecLoc = "" And strName = "" Then
    strTemp = strTemp & "  <CurrentPNR />"
End If
If strPCC <> "" Or bolAllPCC = True Then
   strTemp = strTemp & "<TargetCRSInfo>"
   strTemp = strTemp & "<TargetCRS>1G</TargetCRS>"
   strTemp = strTemp & "<BranchInd>Y</BranchInd>"
   If strPCC <> "" Then strPCC = exactPCC(strPCC)
   If strPCC <> "" Then
      strTemp = strTemp & "<BranchPCC>" & "<![CDATA[" & strPCC & "]]>" & "</BranchPCC>"
   ElseIf bolAllPCC = True Then
      strTemp = strTemp & "<BranchPCC></BranchPCC>"
   End If
   strTemp = strTemp & "<AffiliateName/>"
   strTemp = strTemp & "</TargetCRSInfo>"
End If
If strAirline <> "" Then
   strTemp = strTemp & "<PNRAirIndex>"
   strTemp = strTemp & "<AirV> " & strAirline & "</AirV>"
   strTemp = strTemp & "<FltNum> " & strFlightNum & "</FltNum>"
   strTemp = strTemp & "<OpSuf/>"
   If strBoardPoint <> "" Then
      strTemp = strTemp & "<StartAirp> " & strBoardPoint & "</StartAirp>"
   Else
      strTemp = strTemp & "<StartAirp/>"
   End If
   strTemp = strTemp & "<EndAirp/>"
   strTemp = strTemp & "<StartDt> " & strDepartureDt & "</StartDt>"

   strTemp = strTemp & "</PNRAirIndex>"
End If
If RecLoc <> "" Then
    strTemp = strTemp & "<PNRAddr>"
    strTemp = strTemp & "  <FileAddr />"
    strTemp = strTemp & "  <CodeCheck />"
    strTemp = strTemp & "  <RecLoc>" & UCase(RecLoc) & "</RecLoc>"
    strTemp = strTemp & "  </PNRAddr>"
End If
If strName <> "" Then
    strTemp = strTemp & "<CustNameInfo>"
    strTemp = strTemp & "<NameType/>"
    strTemp = strTemp & "<CustName>" & UCase(strName) & "</CustName>"
    strTemp = strTemp & " </CustNameInfo>"
End If
strTemp = strTemp & "  </PNRBFRetrieveMods>"
strTemp = strTemp & "  </PNRBFManagement_11>"

Set mxmldomPNR = New MSXML2.DOMDocument

Set mxmldomPNR = CreateObject("microsoft.xmldom")
mxmldomPNR.async = False
If mxmldomPNR.loadXML(gobjHost.SendQuery(strTemp, "AgencyPNRBFDisplay_7_5", "GalileoPNR.PNR", "LoadPNR")) = False Then Exit Function
' mxmldomPNR.save "C:\AAA.xml"
If InStr(mxmldomPNR.xml, "TransactionErrorCode") <> 0 Then
   If InStr(mxmldomPNR.xml, "ErrText") <> 0 Then
      If Trim(mxmldomPNR.selectNodes("//ErrText/Text").Item(0).Text) <> "" Then
         'MsgBox mxmldomPNR.selectNodes("//ErrText/Text").Item(0).Text, vbApplicationModal + vbExclamation + vbOKOnly, "CWT Desktop"
         strMsg = mxmldomPNR.selectNodes("//ErrText/Text").Item(0).Text
         modMsgBox.OKMsg = "OK"
         modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      Else
         'MsgBox "UNABLE TO RETRIEVE RECORD LOCATOR", vbApplicationModal + vbExclamation + vbOKOnly, "CWT Desktop"
          strMsg = "UNABLE TO RETRIEVE RECORD LOCATOR"
          modMsgBox.OKMsg = "OK"
          modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      End If
   Else
      'MsgBox "UNABLE TO RETRIEVE RECORD LOCATOR", vbApplicationModal + vbExclamation + vbOKOnly, "CWT Desktop"
      strMsg = "UNABLE TO RETRIEVE RECORD LOCATOR"
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
   End If
   Exit Function
End If
If Not mcolSimiliarNames Is Nothing Then Set mcolSimiliarNames = Nothing
If InStr(mxmldomPNR.xml, "AgncySimlrNameMatch") <> 0 Then
   'Similiar PNRs
   Set xmlnlSNL = mxmldomPNR.selectNodes("//AgncySimlrNameMatch")
   Set mcolSimiliarNames = New Collection
   For intC = 0 To xmlnlSNL.length - 1
       Set objSimiliarName = New SimiliarName
       Set xmlnlSNLChilds = xmlnlSNL.Item(intC).childNodes
       For intD = 0 To xmlnlSNLChilds.length - 1
            With xmlnlSNLChilds.Item(intD)
                Select Case .nodeName
                       Case "CustName"
                          objSimiliarName.PaxName = .Text
                       Case "FirstSegStartDt"
                          objSimiliarName.FirstSegDate = CDate(Mid(.Text, 1, 4) & "-" & Mid(.Text, 5, 2) & "-" & Mid(.Text, 7, 2))
                       Case "NumSeats"
                          objSimiliarName.Seat = .Text
                       Case "RecLoc"
                          objSimiliarName.PNR = .Text
                       Case "CancelledInd"
                          objSimiliarName.Status = IIf(.Text = "Y", True, False)
                End Select
            End With
       Next
       mcolSimiliarNames.Add objSimiliarName
   Next
   LoadPNR = "SNL"
   Exit Function
End If
LoadPNR = "PRO"
Call LoadGenPNRInfo(mxmldomPNR.getElementsByTagName("GenPNRInfo").Item(0).childNodes)
Call LoadNames(mxmldomPNR.selectNodes("//LNameInfo"), mxmldomPNR.selectNodes("//FNameInfo"))
Call LoadNameRemarks(mxmldomPNR.selectNodes("//NameRmkInfo"))
Call LoadAirSegs(mxmldomPNR.selectNodes("//AirSeg|//OpenAirSeg|//AirSegOpAirV"))
Call LoadVRecLoc(mxmldomPNR.selectNodes("//VndRecLocs/RecLocInfoAry/RecLocInfo"))
Call LoadHtlSegs(mxmldomPNR.selectNodes("//HtlSeg|//HtlSegOptFlds|//NonAirSeg[Type='HTL']"))
Call LoadTurSegs(mxmldomPNR.selectNodes("//NonAirSeg[Type='TUR']"))  'TUR
Call LoadAcctRmks(mxmldomPNR.selectNodes("//InvoiceRmk"))
Call LoadClientInfo(mxmldomPNR.selectNodes("//ProfileClientFileAssoc"))
Call LoadGenRmks(mxmldomPNR.selectNodes("//GenRmkInfo"))
Call LoadItinRmks(mxmldomPNR.selectNodes("//ItinRmk"))
Call LoadFreqCust(mxmldomPNR.selectNodes("//FreqCustInfo|//CrossAccrual/CrossAccrualAry/CrossAccrualItem/AirV"))
Call LoadPaidDue(mxmldomPNR.selectNodes("//DuePaidInfo"))
Call LoadPhone(mxmldomPNR.selectNodes("//PhoneInfo"))
Call LoadQMinder(mxmldomPNR.selectNodes("//QMinder"))
Call LoadFOP(mxmldomPNR.selectNodes("//CreditCardFOP|//CheckFOP|//OtherFOP"))
Call LoadAddr(mxmldomPNR.selectNodes("//DeliveryAddrInfo/DeliveryAddr|//AddrInfo/Addr"))
Call LoadCarSegs(mxmldomPNR.selectNodes("//CarSeg|//CarSegOptFlds|//NonAirSeg[Type='CAR']"))
Call LoadSSRequests(mxmldomPNR.selectNodes("//ProgramaticSSR|//ProgramaticSSRText"))
Call LoadOSI(mxmldomPNR.selectNodes("//OSI"))
Call LoadTktDate(mxmldomPNR.selectNodes("//TkArrangement/Text|//TAUTkArrangement/QTAUDt"))
'ZhiSam - V1.2.19 20130405 - CR-203 - Add LoadTAWDate for TAW xml segment
Call LoadTAWDate(mxmldomPNR.selectNodes("//TAWTkArrangement/QTAWDt"))
Call LoadSeatData(mxmldomPNR.selectNodes("//SeatSeg|//SeatAssignment"))
'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
Call LoadDSInfo

If Not mcolFiledFares Is Nothing Then Set mcolFiledFares = Nothing
Set mcolFiledFares = New Collection

If Me.FareDataExists = True Then
    mintFiledFareCount = GetNumberFF
    For intC = 1 To mintFiledFareCount
        LoadFiledFare (CStr(intC))
    Next
End If

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Dim strTknum As String
Dim TkNumAry() As String
Dim lngC As Long

strTknum = captureTktNum()

If strTknum <> "" Then
   TkNumAry = Split(strTknum, ";")
   
     If Not mcolETicket Is Nothing Then Set mcolETicket = Nothing
     Set mcolETicket = New Collection
        
     For lngC = LBound(TkNumAry) To UBound(TkNumAry)
       Call LoadETicket(TkNumAry(lngC))
     Next
End If





End Function

Public Function FiledFare(Index As Variant) As PxFF ' FiledFare
    Set FiledFare = mcolFiledFares.Item(Index)
End Function

Public Property Get FiledFareCount() As Long
    FiledFareCount = mcolFiledFares.Count
End Property
Public Function PassengerName(Index As Variant) As PNRName
    Set PassengerName = mcolNames.Item(Index)
    
End Function

Public Property Get PassengerCount() As Long
    If Not mcolNames Is Nothing Then
        PassengerCount = mcolNames.Count
    Else
        PassengerCount = 0
    End If
End Property

Public Function AirSeg(Index As Variant) As AirSegment
    Set AirSeg = mcolAirSegs.Item(Index)
End Function

Public Function HotelSeg(Index As Variant) As HotelSeg
    Set HotelSeg = mcolHtlSegs.Item(Index)
End Function

Public Property Get HotelSegCount() As Long
    HotelSegCount = mcolHtlSegs.Count
End Property

'TUR
Public Function TurSeg(Index As Variant) As TurSeg
    Set TurSeg = mcolTurSegs.Item(Index)
End Function
'TUR
Public Property Get TurSegCount() As Long
    TurSegCount = mcolTurSegs.Count
End Property

Public Property Get AcctRemarkCount() As Long
    AcctRemarkCount = mcolAcctRmks.Count
End Property

Public Function AcctRemark(Index As Variant) As AcctRemark
    Set AcctRemark = mcolAcctRmks.Item(Index)
End Function
Public Function GeneralRemarkCount() As Long
    GeneralRemarkCount = mcolGenRmks.Count
End Function
Public Function GeneralRemark(Index As Variant) As GenRemark
    Set GeneralRemark = mcolGenRmks.Item(Index)
End Function
Public Function GeneralRemarks() As Collection
    Set GeneralRemarks = mcolGenRmks
End Function

Public Function ItinRemark(Index As Variant) As ItinRemark
    Set ItinRemark = mcolItinRmks.Item(Index)
End Function

Public Function ItinRemarks() As Collection
    Set ItinRemarks = mcolItinRmks
End Function
'TUR
Public Property Get ItinRemarkCount() As Long
    ItinRemarkCount = mcolItinRmks.Count
End Property
Public Property Get FreqCustCount() As Long
    FreqCustCount = mcolFreqCustNums.Count
End Property

Public Function FreqCustNumber(Index As Variant) As FreqCustNum
    Set FreqCustNumber = mcolFreqCustNums.Item(Index)
End Function

Public Function PaidDue(Index As Variant) As PaidDue
    Set PaidDue = mcolPaidDue.Item(Index)
End Function

Public Property Get PaidDueCount() As Long
    PaidDueCount = mcolPaidDue.Count
End Property
Public Function Phone(Index As Variant) As Phone
    Set Phone = mcolPhone.Item(Index)
End Function
Public Function PhoneCount() As Long
    PhoneCount = mcolPhone.Count
End Function


Public Function QueueMinder(Index As Variant) As QMinder
    Set QueueMinder = mcolQMinders.Item(Index)
End Function

Public Function Names() As Collection
    Set Names = mcolNames
    
End Function

Public Function EnumNames() As IUnknown
Attribute EnumNames.VB_UserMemId = -4
    Set EnumNames = Me.Names.[_NewEnum]
    
End Function

Private Function DisplayCurrent() As MSXML2.DOMDocument

End Function

Private Function DisplayByRecLoc(RecordLocator As String, Optional PCC As String, Optional GDS As String = "1G") As MSXML2.DOMDocument

End Function

Private Function DisplayByName(LastName, Optional FirstName As String, Optional PCC As String, Optional GDS As String = "1G") As Variant

End Function

Private Function ConvertDate(DateString As String, Optional TimeString As String = "0001") As Date
    ConvertDate = CDate(Format(DateString & Format(TimeString, "0000"), "0000-00-00 00:00"))
End Function
Private Function ConvertNumber(NumberAmt As String) As Single
    If IsNumeric(NumberAmt) Then
        ConvertNumber = CSng(NumberAmt)
    Else
        ConvertNumber = 0
    End If
End Function

Private Function PlaceDecimal(ByVal NumDecimals As Integer, Amount As Long) As Single

Dim sngDecFactor As Single
Dim intC As Integer

sngDecFactor = 1
For intC = 1 To NumDecimals
    sngDecFactor = sngDecFactor * 0.1
Next
PlaceDecimal = Amount * sngDecFactor

End Function
Public Property Get AirFaresFOPCount() As Long
    AirFaresFOPCount = mcolAirFaresFOP.Count
End Property

Public Function AirFaresFOP(Index As Variant) As AirFaresFOP
    Set AirFaresFOP = mcolAirFaresFOP.Item(Index)
End Function
'Added on 5/1/2005: Getting company information by Profile Name instead on CN
Private Sub LoadClientInfo(ByVal NodeList As MSXML2.IXMLDOMNodeList)

Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlClient As MSXML2.IXMLDOMNodeList
Dim strBar As String
'Dim dbComp As dao.Database
Dim rsComp As ADODB.Recordset
Dim intLength As Integer
'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
Dim strBarPCC As String

If NodeList.length > 0 Then
lngC1 = 0
    Set xmlnlClient = NodeList.Item(lngC1).childNodes
            For lngC2 = 0 To xmlnlClient.length - 1
                With xmlnlClient.Item(lngC2)
                    Select Case .nodeName
                        Case "BAR"
                            strBar = .Text
                            Exit For
                        Case "MAR"
                            'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
                            'Store the BAR PCC
                            strBarPCC = .Text
                    End Select
                End With
            Next

'Else
'    MsgBox "No Client File in PNR", vbOKOnly, "Client Info Warning"
End If

Set xmlnlClient = Nothing
                                If strBar <> "" Then
                                    'Set dbComp = dao.OpenDatabase(gstrTProDBSource)

                                    Set rsComp = gdbConn.Execute("SELECT * FROM tblClients WHERE [ProName] = '" & strBar & "'")
                                    If Not rsComp.EOF Then
                                        With mvarCompInfo
                                        'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
                                            .AgencyName = rsComp![AgencyName]
                                        '--
                                            .ClientID = rsComp![ClientID]
                                            .ClientType = rsComp![ClientType] & ""
                                            .CompanyName = rsComp![ClientName]
                                            .Contracts = rsComp![PFCodes]
                                            .CreditStatus = rsComp![CreditStatus]
                                            .MerchFeePct = rsComp![MerchFeePct]
                                            If Not IsNull(rsComp![CreditTerms]) Then
                                                .CreditTerms = rsComp![CreditTerms]
                                            Else
                                                .CreditTerms = 0
                                            End If
                                            .DiscountDomestic = rsComp![PubDiscDom]
                                            .DiscountInternational = rsComp![PubDiscIntl]
                                            .DIV = rsComp![DIV] & ""
                                            .EInvoice = rsComp![Einv]
                                            .HotelRateCode = rsComp![HtlGDSRateCode] & ""
                                            .MarkUp = rsComp![MarkUpIntl]
                                            .ProfileName = rsComp![ProName] & ""
                                            If Not IsNull(rsComp![TFIncMF]) Then
                                                .TFIncMF = rsComp![TFIncMF]
                                            Else
                                                .TFIncMF = False
                                            End If
                                            If Not IsNull(rsComp![CPGClient]) Then
                                                .CPG = rsComp![CPGClient]
                                            Else
                                                .CPG = False
                                            End If
                                            If Not IsNull(rsComp![TransactionFeesGroup]) Then
                                                .TransactionFeeGroup = rsComp![TransactionFeesGroup]
                                            Else
                                                .TransactionFeeGroup = 0
                                            End If
                                            If Not IsNull(rsComp![MarkUpGroup]) Then
                                                .MarkUpGroup = rsComp![MarkUpGroup]
                                            Else
                                                .MarkUpGroup = 0
                                            End If
                                            
                                            If Not IsNull(rsComp!CalcBy) Then
                                               .TFCalcBy = Trim(rsComp!CalcBy)
                                            End If
                                            .WONum = rsComp![WONum] & ""
                                            .AltFQPCC = rsComp![AltPCC] & ""
                                            If Not IsNull(rsComp![MI]) Then
                                                .MI = rsComp![MI]
                                            Else
                                                .MI = False
                                            End If
                                            
                                            'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
                                            'Store the BAR PCC
                                            .ProfilePCC = strBarPCC
                                            'Preethi - V1.2.4 20110622  - ER01 - Tracking of Touches to CWT Booking
                                            If Not IsNull(rsComp![TrackingTouches]) Then
                                               .PNRTrackingTouches = rsComp![TrackingTouches]
                                            Else
                                               .PNRTrackingTouches = False
                                            End If
                                            'CC - V1.2.8 20111028  - CR110 - Aqua EItn Module
                                            If IsNull(rsComp![AquaItin]) Then
                                               .AquaItin = False
                                            Else
                                               .AquaItin = rsComp![AquaItin]
                                            End If

                                        End With
                                    End If
                                    rsComp.Close
                                    'dbComp.Close
                                    Set rsComp = Nothing
                                    'Set dbComp = Nothing
                                'Else
                                '    MsgBox "No Client File in PNR", vbOKOnly, "Client Info Warning"
                                End If
End Sub
Private Sub LoadVRecLoc(ByVal NodeList As MSXML2.IXMLDOMNodeList)
Dim objVRL As VRecLoc
Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlVRL As MSXML2.IXMLDOMNodeList
Dim strDt As String
Dim strTm As String

If Not mcolVRecLocs Is Nothing Then Set mcolVRecLocs = Nothing
Set mcolVRecLocs = New Collection

For lngC1 = 0 To NodeList.length - 1
    Set xmlnlVRL = NodeList.Item(lngC1).childNodes
    Set objVRL = New VRecLoc
    strDt = ""
    strTm = ""
            For lngC2 = 0 To xmlnlVRL.length - 1
                With xmlnlVRL.Item(lngC2)
                    Select Case .nodeName
                        Case "TmStamp"
                            strTm = .Text
                        Case "DtStamp"
                            strDt = .Text
                        Case "Vnd"
                            objVRL.Vendor = .Text
                        Case "RecLoc"
                            objVRL.RecLoc = .Text
                    End Select
                End With
            Next
        objVRL.DateTime = ConvertDate(strDt, strTm)
        mcolVRecLocs.Add objVRL
    Next

Set xmlnlVRL = Nothing
Set objVRL = Nothing

End Sub
Private Sub LoadGenPNRInfo(ByVal NodeList As MSXML2.IXMLDOMNodeList)
Dim intC As Integer

For intC = 0 To NodeList.length - 1
        With NodeList.Item(intC)
        Select Case .nodeName
            Case "FileAddr"
                mstrFileAddr = .Text
            Case "CodeCheck"
                mstrCodeCheck = .Text
            Case "RecLoc"
                mstrRecLoc = .Text
                If mstrRecLoc <> "" Then mlngPNRStatus = 1
            Case "OwningAgncyPCC"
                mstrPCCOwner = .Text
            Case "CreationDt"
               mdtmCreateDate = ConvertDate(IIf(.Text <> "", .Text, Format(Date, "yyyymmdd")))
            Case "CreatingAgntSignOn"
                'mstrAgent = Mid(.Text, 5, 2)
                mstrAgent = .Text
                If IsNumeric(.Text) Then
                   mstrAgent = Int(Mid(.Text, 2, 5))
                Else
                   mstrAgent = Mid(.Text, 5)
                End If
            Case "InUseInd"
                If .Text = "Y" And mlngPNRStatus < 2 Then mlngPNRStatus = 2
            Case "SimultaneousUpdInd"
                If .Text = "Y" Then mlngPNRStatus = 3
            Case "FareDataExistsInd"
                mbolFareDataExists = IIf(.Text = "Y", True, False)
            Case "TkNumExistInd"
                mbolTktNumExists = IIf(.Text = "Y", True, False)
            Case "IMUdataexists"
                mbolIMUDataExists = IIf(.Text = "Y", True, False)
            Case "ETkDataExistInd"
                mbolETktDataExists = IIf(.Text = "Y", True, False)
            Case Else
                ' nothing
        End Select
    End With
Next

If mlngPNRStatus = 0 Then
    If InStr(gobjHost.TerminalEntry("*R"), "NO B.F. TO DISPLAY - ") Then mlngPNRStatus = -1
End If

End Sub

Private Sub LoadNames(ByVal LastNames As MSXML2.IXMLDOMNodeList, ByVal FirstNames As MSXML2.IXMLDOMNodeList)
Dim intLN As Integer
Dim intFN As Integer
Dim xmlnlLN As MSXML2.IXMLDOMNodeList
Dim xmlnlFN As MSXML2.IXMLDOMNodeList
Dim lngC1 As Long
Dim lngC2 As Long
Dim objName As PNRName
Dim strLName As String
Dim strNameNum As String
Dim strNameType As String

If Not mcolNames Is Nothing Then Set mcolNames = Nothing
Set mcolNames = New Collection

For intLN = 0 To LastNames.length - 1
    Set xmlnlLN = LastNames.Item(intLN).childNodes
        With xmlnlLN
            For lngC1 = 0 To .length - 1
                With .Item(lngC1)
                Select Case .nodeName
                    Case "LNameNum"
                        strNameNum = .Text
                    Case "NumPsgrs"
                        intFN = CLng(.Text)
                    Case "LName"
                        strLName = .Text
                    Case "NameType"
                        strNameType = .Text
                    Case Else
                        'nothing
                End Select
            
            End With
            Next
        End With
        
        
        For intFN = mcolNames.Count To (intFN - 1) + mcolNames.Count
        Set objName = New PNRName
        Set xmlnlFN = FirstNames.Item(intFN).childNodes
            For lngC2 = 0 To xmlnlFN.length - 1

                With xmlnlFN.Item(lngC2)
                    Select Case .nodeName
                        Case "PsgrNum"
                            objName.GDSNum = strNameNum & "." & .Text
                        Case "AbsNameNum"
                            objName.PassengerNum = .Text
                        Case "FName"
                            objName.FirstName = .Text
                    End Select
                End With
            Next lngC2
        objName.LastName = strLName
        objName.PassengerType = strNameType
        mcolNames.Add objName
        Next intFN
    
    Next intLN
                        
                
End Sub
Private Sub LoadNameRemarks(ByVal NameRemarks As MSXML2.IXMLDOMNodeList)
    Dim i As Integer
    Dim j As Integer
    Dim strGDSNum As String
    Dim strNameRemark As String
    Dim objName As PNRName
    Dim xmlnlNR As MSXML2.IXMLDOMNodeList
    
    For i = 0 To NameRemarks.length - 1
        Set xmlnlNR = NameRemarks.Item(i).childNodes
        strNameRemark = ""
        strGDSNum = ""
        With xmlnlNR
            For j = 0 To .length - 1
                With .Item(j)
                     Select Case .nodeName
                            Case "LNameNum"
                                 strGDSNum = .Text & "."
                            Case "PsgrNum"
                                 strGDSNum = strGDSNum & .Text
                            Case "NameRmk"
                                 strNameRemark = .Text
                     End Select
                End With
            Next
            For j = 1 To mcolNames.Count
                Set objName = mcolNames.Item(j)
                If objName.GDSNum = strGDSNum Then
                   objName.Remark = strNameRemark
                   Exit For
                End If
            Next
            
        End With
    Next
End Sub



Private Sub LoadAcctRmks(ByVal AcctRmksNL As MSXML2.IXMLDOMNodeList)
Dim objAR As AcctRemark
Dim objGASR As GASalesRecord
Dim objAirFaresFOP As AirFaresFOP
Dim FOPFields() As String
'Dim intX As Integer
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim strTemp As String
Dim sngTax As Single
Dim xmlnlRmk As MSXML2.IXMLDOMNodeList
Dim strGASR As String
Dim strFields() As String
Dim bolEOR As Boolean
Dim lngBegLine As Long
'Dim dbComp As dao.Database
'Dim rsComp As dao.Recordset
Dim intDisplayNo As Integer
Dim strTmp() As String

If Not mcolAcctRmks Is Nothing Then Set mcolAcctRmks = Nothing
If Not mcolGASalesRecords Is Nothing Then Set mcolGASalesRecords = Nothing
If Not mcolAirFaresFOP Is Nothing Then Set mcolAirFaresFOP = Nothing
Set mcolAcctRmks = New Collection
Set mcolGASalesRecords = New Collection
Set mcolAirFaresFOP = New Collection

For lngC1 = 0 To AcctRmksNL.length - 1
    Set xmlnlRmk = AcctRmksNL.Item(lngC1).childNodes
    Set objAR = New AcctRemark
    Set objAirFaresFOP = New AirFaresFOP
            For lngC2 = 0 To xmlnlRmk.length - 1
                With xmlnlRmk.Item(lngC2)
                    Select Case .nodeName
                        Case "ItemNum"
                            objAR.ItemNum = CLng(.Text)
                        Case "Keyword"
                            Select Case .Text
                                Case "3000"
                                    objAR.RemarkType = "FT="
                                Case "3001"
                                    objAR.RemarkType = "DYO"
                                Case "3002"
                                    objAR.RemarkType = "FS"
                                Case "3003"
                                    objAR.RemarkType = "CR"
                                Case "3006"
                                    objAR.RemarkType = "TK"
                                Case "3007"
                                    objAR.RemarkType = "CA"
                                Case "3009"
                                    objAR.RemarkType = "AC"
                                Case "3010"
                                    objAR.RemarkType = "FT"
                                Case "3011"
                                    objAR.RemarkType = "DEF"
                                Case "3018"
                                    objAR.RemarkType = "FA"
                                Case "3019"
                                    objAR.RemarkType = "PH"
                                Case Else
                                    objAR.RemarkType = "??"
                            End Select
    
                        Case "Rmk"
                            objAR.RemarkText = IIf(InStr(.Text, "@") > 0, Replace(.Text, "@", "."), .Text)
                            If Left(.Text, 3) = "CN/" Then
                                mstrCN = Mid(.Text, 4)
                               ' If mstrCN <> "" Then
                               '     Set dbComp = dao.OpenDatabase(gstrTProDBSource)
                               '     Set rsComp = dbComp.OpenRecordset("SELECT * FROM tblClients WHERE [CN] = '" & mstrCN & "'")
                               '     If Not rsComp.EOF Then
                               '         With mvarCompInfo
                               '             .ClientType = rsComp![ClientType] & ""
                               '             .CompanyName = rsComp![ClientName]
                               '             .Contracts = rsComp![PFCodes]
                               '             .CreditStatus = rsComp![CreditStatus]
                               '             .MerchFeePct = rsComp![MerchFeePct]
                               '             If Not IsNull(rsComp![CreditTerms]) Then
                               '                 .CreditTerms = rsComp![CreditTerms]
                               '             Else
                               '                 .CreditTerms = 0
                               '             End If
                               '             .DiscountDomestic = rsComp![PubDiscDom]
                               '             .DiscountInternational = rsComp![PubDiscIntl]
                               '             .DIV = rsComp![DIV] & ""
                               '             .EInvoice = rsComp![Einv]
                               '             .HotelRateCode = rsComp![HtlGDSRateCode] & ""
                               '             .MarkUp = rsComp![MarkUpIntl]
                               '             .ProfileName = rsComp![ProName] & ""
                               '             If Not IsNull(rsComp![TransactionFeesGroup]) Then
                               '                 .TransactionFeeGroup = rsComp![TransactionFeesGroup]
                               '             Else
                               '                 .TransactionFeeGroup = 0
                               '             End If
                               '             .WONum = rsComp![WONum] & ""
                                            
                               '         End With
                               '     End If
                               '     rsComp.Close
                               '     dbComp.Close
                               '     Set rsComp = Nothing
                               '     Set dbComp = Nothing
                               ' End If
                            'Added on 16/08/04
                            ElseIf Left(.Text, 4) = "FOP/" Then
                                FOPFields = Split(Mid(.Text, 5), "/")
                                'For intX = 0 To UBound(FOPFields)
                                If UBound(FOPFields) >= 1 Then
                                    If UCase(Left(FOPFields(1), 2)) = "CC" Or UCase(Left(FOPFields(1), 2)) = "CX" Then
                                        objAirFaresFOP.FFNumber = Mid(FOPFields(0), 2)
                                        objAirFaresFOP.FOPType = FOPFields(1)
                                        objAirFaresFOP.FOP_CCCode = Mid(FOPFields(2), 1, 2)
                                        objAirFaresFOP.FOP_CCNum = Mid(FOPFields(2), 3)
                                        objAirFaresFOP.FOPAmount = IIf(IsNumeric(Replace(FOPFields(3), "@", ".")), Replace(FOPFields(3), "@", "."), 0)
                                    Else
                                        objAirFaresFOP.FFNumber = ""
                                        objAirFaresFOP.FOPType = ""
                                        objAirFaresFOP.FOP_CCCode = ""
                                        objAirFaresFOP.FOP_CCNum = ""
                                        objAirFaresFOP.FOPAmount = 0
                                    End If
                                Else
                                    objAirFaresFOP.FFNumber = ""
                                    objAirFaresFOP.FOPType = ""
                                    objAirFaresFOP.FOP_CCCode = ""
                                    objAirFaresFOP.FOP_CCNum = ""
                                    objAirFaresFOP.FOPAmount = 0
                                End If
                                'Next
                                mcolAirFaresFOP.Add objAirFaresFOP
                            End If
                    End Select
                End With
            Next
        mcolAcctRmks.Add objAR
    Next

' Now parsing Account Remarks that are a GA Sales record

strGASR = ""
lngC2 = 0
intDisplayNo = 0
For Each objAR In mcolAcctRmks
    With objAR
        lngC2 = lngC2 + 1
        If Left(.RemarkText, 2) = "SF" And InStr(1, .RemarkText, "/*") > 0 Then
           strTmp = Split(.RemarkText, "/")
           intDisplayNo = Mid(strTmp(1), 2)
        End If
        If .RemarkType = "FT" And Left(.RemarkText, 3) = "MS/" Then
            strGASR = Mid(.RemarkText, 4)
            lngBegLine = lngC2
        End If
        If .RemarkType = "FT" And Left(.RemarkText, 4) = "MSX/" Then strGASR = strGASR & Mid(.RemarkText, 4)
        If strGASR <> "" Then
            If lngC2 = mcolAcctRmks.Count Then
                bolEOR = True
            ElseIf Left(Me.AcctRemark(lngC2 + 1).RemarkText, 4) <> "MSX/" Then
                bolEOR = True
            End If
        End If
        
        If bolEOR Then
            
            Do While InStr(1, strGASR, "@")
                strGASR = Left(strGASR, InStr(1, strGASR, "@") - 1) & "." & Mid(strGASR, InStr(1, strGASR, "@") + 1)
            Loop
            
            strFields = Split(strGASR, "/")
            Set objGASR = New GASalesRecord
            With objGASR
                .DisplanyNo = intDisplayNo
                .BegLine = lngBegLine
                .EndLine = lngC2
                For lngC1 = 0 To UBound(strFields)
                    Select Case Left(strFields(lngC1), 1)
                        Case "A"
                            If IsNumeric(Mid(strFields(lngC1), 2, 1)) Then
                                .BaseAmount = CSng(Mid(strFields(lngC1), 2))
                            'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
                            ElseIf Mid(strFields(lngC1), 2, 1) = "C" Then
                                'For AC Code
                            Else
                                .BaseCurrency = Mid(strFields(lngC1), 2, 3)
                                .BaseAmount = CSng(Mid(strFields(lngC1), 3))
                            End If
                        Case "C"
                            If Mid(strFields(lngC1), 2, 2) = "CN" Then
                                .CCNumber = Mid(strFields(lngC1), 4)
                            ElseIf IsNumeric(Mid(strFields(lngC1), 2, 1)) Then
                                .CommissionAmount = Mid(strFields(lngC1), 2)
                            End If
                        Case "D"
                            .CollectedAmount = CSng(Mid(strFields(lngC1), 2))
                        Case "F"
                            If Mid(strFields(lngC1), 2, 1) = "F" Then
                                .FreeFieldAdd strFields(lngC1), .BegLine, .EndLine
                            Else
                                .FOP = Mid(strFields(lngC1), 2)
                            End If
                        Case "P"
                            If Mid(strFields(lngC1), 2, 1) = "C" Then
                                .ProductCode = Mid(strFields(lngC1), 3)
                            ElseIf Mid(strFields(lngC1), 2, 1) = "O" Then
                                .PONumber = Mid(strFields(lngC1), 3)
                            ElseIf Mid(strFields(lngC1), 2, 1) = "X" Then
                                'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
                                .PaxNum = Mid(strFields(lngC1), 3)
                            End If
                        Case "S"
                            If Mid(strFields(lngC1), 2, 1) = "F" Then
                                .SellAmount = Mid(strFields(lngC1), 3)
                            ElseIf IsNumeric(Mid(strFields(lngC1), 2, 1)) Then
                                .BaseAmount = Mid(strFields(lngC1), 2)
                            End If
                        Case "G"
                            If IsNumeric(Mid(strFields(lngC1), 2, 1)) Then
                                .GSTAmount = Mid(strFields(lngC1), 2)
                            End If
                        Case "T"
                            If Mid(strFields(lngC1), 2, 1) = "K" Then
                                .TicketNumber = Mid(strFields(lngC1), 3)
                            ElseIf Mid(strFields(lngC1), 2, 1) = "X" Then
                                strTemp = Mid(strFields(lngC1), 3)
                                If IsNumeric(strTemp) Then
                                    sngTax = CSng(strTemp)
                                Else
                                For lngC3 = 1 To Len(strTemp)
                                    If Not IsNumeric(Left(strTemp, lngC3)) Then Exit For
                                Next lngC3
                                sngTax = CSng(Left(strTemp, lngC3 - 1))
                                strTemp = Mid(strTemp, lngC3 + 2)
                                If Len(strTemp) > 0 Then
                                    For lngC3 = 1 To Len(strTemp)
                                        If Not IsNumeric(Left(strTemp, lngC3)) Then Exit For
                                    Next
                                    sngTax = sngTax + CSng(Left(strTemp, lngC3 - 1))
                                    End If
                                End If
                                .Tax = sngTax
                            End If
                        Case "V"
                            .VendorCode = Mid(strFields(lngC1), 2)
                    End Select
                Next
            End With
         mcolGASalesRecords.Add objGASR
         strGASR = ""
         bolEOR = False
         End If
    End With
Next

Set objAR = Nothing
Set objGASR = Nothing

End Sub

Private Sub LoadGenRmks(ByVal GenRmksNL As MSXML2.IXMLDOMNodeList)
Dim objGR As GenRemark
Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlRmk As MSXML2.IXMLDOMNodeList

If Not mcolGenRmks Is Nothing Then Set mcolGenRmks = Nothing
Set mcolGenRmks = New Collection

For lngC1 = 0 To GenRmksNL.length - 1
    Set xmlnlRmk = GenRmksNL.Item(lngC1).childNodes
    Set objGR = New GenRemark
            For lngC2 = 0 To xmlnlRmk.length - 1
                With xmlnlRmk.Item(lngC2)
                    Select Case .nodeName
                        Case "GenRmkNum"
                            objGR.ItemNum = CLng(.Text)
                        Case "GenlRmkQual"
                            objGR.Qualifier = .Text
                        Case "GenRmk"
                            objGR.RemarkText = .Text
                    End Select
                End With
            Next
        mcolGenRmks.Add objGR
    Next

End Sub

Private Sub LoadItinRmks(ByVal ItinRmksNL As MSXML2.IXMLDOMNodeList)
Dim objIR As ItinRemark
Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlRmk As MSXML2.IXMLDOMNodeList

If Not mcolItinRmks Is Nothing Then Set mcolItinRmks = Nothing
Set mcolItinRmks = New Collection

For lngC1 = 0 To ItinRmksNL.length - 1
    Set xmlnlRmk = ItinRmksNL.Item(lngC1).childNodes
    Set objIR = New ItinRemark
            For lngC2 = 0 To xmlnlRmk.length - 1
                With xmlnlRmk.Item(lngC2)
                    Select Case .nodeName
                        Case "SegNum"
                            If .Text <> "" Then objIR.SegNum = .Text
                        Case "RmkNum"
                            objIR.ItemNum = CLng(.Text)
                        Case "Rmk"
                            objIR.RemarkText = .Text
                    End Select
                End With
            Next
        mcolItinRmks.Add objIR
    Next

End Sub
Private Sub LoadSeatData(ByVal SeatNL As MSXML2.IXMLDOMNodeList)
Dim objSD As SeatData
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim lngC4 As Long
Dim lngC5 As Long
Dim xmlnlAir As MSXML2.IXMLDOMNodeList
Dim xmlnlSA As MSXML2.IXMLDOMNodeList
Dim xmlnlSeat As MSXML2.IXMLDOMNodeList
Dim xmlnlAttribute As MSXML2.IXMLDOMNodeList
Dim strSegNumTemp As String

If Not mcolSeatData Is Nothing Then Set mcolSeatData = Nothing
Set mcolSeatData = New Collection

For lngC1 = 0 To SeatNL.length - 1
    If SeatNL.Item(lngC1).nodeName = "SeatSeg" Then
    
     Set xmlnlSeat = SeatNL.Item(lngC1).childNodes
        For lngC2 = 0 To xmlnlSeat.length - 1
            With xmlnlSeat.Item(lngC2)
                Select Case .nodeName
                    Case "FltSegNum"
                      'Set xmlnlAir = SeatNL.Item(lngC2).childNodes
                      'For lngC3 = 0 To xmlnlAir.length - 1
                      '      If xmlnlAir.Item(lngC3).nodeName = "FltSegNum" Then
                                strSegNumTemp = xmlnlSeat.Item(lngC2).Text
                      '          Exit For
                      '      End If
                      'Next
                End Select
            End With
        Next lngC2
        End If
        
        
     If SeatNL.Item(lngC1).nodeName = "SeatAssignment" Then
                     
                     Set objSD = New SeatData
                     Set xmlnlSA = SeatNL.Item(lngC1).childNodes
                     For lngC4 = 0 To xmlnlSA.length - 1
                        With xmlnlSA.Item(lngC4)
                            
                            objSD.SegNum = strSegNumTemp
                             Select Case .nodeName
                                Case "AbsNameNum"
                                    objSD.PaxNo = .Text
                                Case "Status"
                                    objSD.Status = .Text
                                Case "Locn"
                                    objSD.SeatLocation = .Text
                                Case "AttribAry"
                                    Set xmlnlAttribute = .childNodes
                                    For lngC5 = 0 To xmlnlAttribute.length - 1
                                        If xmlnlAttribute.Item(lngC5).nodeName = "Attrib" Then
                                            If lngC5 = 0 Then objSD.SeatAttribute1 = .Text
                                            'If lngC5 = 1 Then objSD.SeatAttribute2 = .Text
                                        End If
                                    Next
                               
                             End Select
                             
                             
                        End With
                     
                     Next
                     
                     mcolSeatData.Add objSD
                     
    End If
                
                
                
        

    Next lngC1


End Sub
Public Function SeatData(Index As Variant) As SeatData
    Set SeatData = mcolSeatData.Item(Index)
End Function
Public Property Get SeatDataCount() As Long
    SeatDataCount = mcolSeatData.Count
End Property
Private Sub LoadFreqCust(ByVal FreqCustNL As MSXML2.IXMLDOMNodeList)
Dim objFC As FreqCustNum
Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlFC As MSXML2.IXMLDOMNodeList

If Not mcolFreqCustNums Is Nothing Then Set mcolFreqCustNums = Nothing
Set mcolFreqCustNums = New Collection

For lngC1 = 0 To FreqCustNL.length - 1
    Set objFC = New FreqCustNum
    Set xmlnlFC = FreqCustNL.Item(lngC1).childNodes
        For lngC2 = 0 To xmlnlFC.length - 1
            With xmlnlFC.Item(lngC2)
                Select Case .nodeName
                    Case "AbsNameNum"
                        objFC.PassengerNum = .Text
                    Case "FreqCustV"
                        objFC.Vendor = .Text
                    Case "FreqCustStatus"
                        objFC.Status = .Text
                    Case "FreqCustNum"
                        objFC.FreqCustNum = .Text
                    Case Else
                        'nothing
                End Select
            End With
        Next lngC2
CheckAgain:
        'Check next node is Cross Accrual or not
        If lngC1 + 1 <= FreqCustNL.length - 1 Then
           If FreqCustNL.Item(lngC1 + 1).nodeName = "AirV" Then
              objFC.CrossAccrual = objFC.CrossAccrual & IIf(objFC.CrossAccrual = "", "", "/") & FreqCustNL.Item(lngC1 + 1).Text
              lngC1 = lngC1 + 1
              GoTo CheckAgain
           End If
        End If
        mcolFreqCustNums.Add objFC
    Next lngC1
End Sub

Private Sub LoadPaidDue(ByVal PaidDueNL As MSXML2.IXMLDOMNodeList)
Dim objPD As PaidDue
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngTempAmt As Long
Dim intNumDec As Integer
Dim xmlnlPD As MSXML2.IXMLDOMNodeList

If Not mcolPaidDue Is Nothing Then Set mcolPaidDue = Nothing
Set mcolPaidDue = New Collection

For lngC1 = 0 To PaidDueNL.length - 1
    Set objPD = New PaidDue
    Set xmlnlPD = PaidDueNL.Item(lngC1).childNodes
        For lngC2 = 0 To xmlnlPD.length - 1
            With xmlnlPD.Item(lngC2)
                Select Case .nodeName
                    Case "SegNum"
                        objPD.SegNum = CLng(.Text)
                    Case "Type"
                        objPD.ProductType = .Text
                    Case "Dt"
                        objPD.SegDate = ConvertDate(.Text)
                    Case "DuePaidTextInd"
                        objPD.SegType = .Text
                    Case "Price"
                        lngTempAmt = CLng(.Text)
                    Case "DecPos"
                        intNumDec = IIf(.Text <> "", .Text, 0)
                    Case "Text"
                        objPD.FreeText = .Text
                     'TUR
                    Case "Currency"
                        objPD.CurrencyCode = .Text
                    Case Else
                        ' nothing
                End Select
            End With
        Next lngC2
        objPD.Amount = PlaceDecimal(intNumDec, lngTempAmt)
        lngTempAmt = 0
        mcolPaidDue.Add objPD
    Next lngC1
                    
End Sub

Private Sub LoadPhone(ByVal PhoneNL As MSXML2.IXMLDOMNodeList)
Dim objP As Phone
Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlP As MSXML2.IXMLDOMNodeList

If Not mcolPhone Is Nothing Then Set mcolPhone = Nothing
Set mcolPhone = New Collection

For lngC1 = 0 To PhoneNL.length - 1
    Set objP = New Phone
    Set xmlnlP = PhoneNL.Item(lngC1).childNodes
        For lngC2 = 0 To xmlnlP.length - 1
            With xmlnlP.Item(lngC2)
                Select Case .nodeName
                    Case "PhoneFldNum"
                        objP.ItemNum = .Text
                    Case "Pt"
                        objP.CityCode = .Text
                    Case "Type"
                        objP.PhoneType = .Text
                    Case "Phone"
                        objP.PhoneNum = .Text
                End Select
            End With
        Next lngC2
        mcolPhone.Add objP
    Next lngC1
                    
End Sub

Private Sub LoadQMinder(ByVal QMinderNL As MSXML2.IXMLDOMNodeList)
Dim objQM As QMinder
Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlQM As MSXML2.IXMLDOMNodeList
Dim strTime As String
Dim strDate As String

If Not mcolQMinders Is Nothing Then Set mcolQMinders = Nothing
Set mcolQMinders = New Collection

For lngC1 = 0 To QMinderNL.length - 1
    Set objQM = New QMinder
    Set xmlnlQM = QMinderNL.Item(lngC1).childNodes
    For lngC2 = 0 To xmlnlQM.length - 1
        With xmlnlQM.Item(lngC2)
            Select Case .nodeName
                Case "ItemNum"
                    objQM.ItemNum = .Text
                 Case "Dt"
                    strDate = .Text
                Case "QNum"
                    objQM.QueueNum = .Text
                Case "QCat"
                    objQM.QueueCat = .Text
                Case "PCC"
                    objQM.PCC = .Text
                Case "Tm"
                    strTime = .Text
                Case "Text"
                    objQM.FreeText = .Text
            End Select
        End With
        Next lngC2
    objQM.DateTime = ConvertDate(strDate, strTime)
    mcolQMinders.Add objQM
Next lngC1
End Sub

Private Sub LoadAirSegs(ByVal AirSegNL As MSXML2.IXMLDOMNodeList)
Dim objAS As AirSegment
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngTempAmt As Long
Dim intNumDec As Integer
Dim strDepTime As String
Dim strArrTime As String
Dim strDate As String
Dim strDC As String
Dim strDepCity As String
Dim strArrCity As String
Dim strSegNum As String
Dim xmlnlAS As MSXML2.IXMLDOMNodeList
'Dim dbCity As dao.Database
Dim rsCity As New ADODB.Recordset
Dim strMsg As String

If Not mcolAirSegs Is Nothing Then Set mcolAirSegs = Nothing
Set mcolAirSegs = New Collection

For lngC1 = 0 To AirSegNL.length - 1
    Set objAS = New AirSegment
    Set xmlnlAS = AirSegNL.Item(lngC1).childNodes
        For lngC2 = 0 To xmlnlAS.length - 1
            With xmlnlAS.Item(lngC2)
                Select Case .nodeName
                    Case "SegNum"
                        objAS.SegNumber = CLng(.Text)
                        strSegNum = .Text
                    Case "Status"
                        objAS.Status = .Text
                    Case "Dt"
                        strDate = .Text
                    Case "DayChg"
                        strDC = .Text
                    Case "AirV"
                        objAS.Vendor = .Text
                    Case "NumPsgrs"
                        objAS.NumberSeats = .Text
                    Case "StartAirp"
                        objAS.DepartAirport = .Text
                        strDepCity = .Text
                    Case "EndAirp"
                        objAS.ArriveAirport = .Text
                        strArrCity = .Text
                    Case "StartTm"
                        strDepTime = .Text
                    Case "EndTm"
                        strArrTime = .Text
                    Case "BIC"
                        objAS.Class = .Text
                    Case "FltNum"
                        objAS.FlightNumber = .Text
                    Case "TklessInd"
                        objAS.ETicketEligible = IIf(.Text = "Y", True, False)
                    Case "FltFlownInd"
                        objAS.Flown = IIf(.Text = "Y", True, False)
                    
                        
                    Case Else
                        ' nothing
                End Select
            End With
        Next lngC2
        
       ' check for code share flight
       If lngC1 <> AirSegNL.length - 1 Then
            With AirSegNL.Item(lngC1 + 1)
                If .nodeName = "AirSegOpAirV" Then
                    objAS.CodeShareFlight = True
                    objAS.OperatedBy = .selectSingleNode("OpAirVInfoAry/OpAirVInfo/AirV").Text
                    If objAS.OperatedBy = "" Then objAS.OperatedBy = .selectSingleNode("OpAirVInfoAry/OpAirVInfo/AirVName").Text
                    lngC1 = lngC1 + 1 ' Skipping next node since I got what I wanted
                End If
            End With
        End If
        
        If AirSegNL.Item(lngC1).nodeName = "OpenAirSeg" Then objAS.FlightNumber = "OPEN"
        If strDate <> "" Then objAS.DepartDateTime = ConvertDate(strDate, strDepTime)
        If strDate <> "" And objAS.FlightNumber <> "OPEN" Then objAS.ArriveDateTime = DateAdd("d", CDbl(strDC), ConvertDate(strDate, strArrTime))
        'Modified on 21/03/2005
        'Set dbCity = OpenDatabase(gstrTProLookupDBSource)
        
        Set rsCity = gdbConn.Execute("SELECT * FROM tblCityCodes WHERE AirportCode = '" & strArrCity & "'")
        
        With rsCity
        If Not .EOF Then
            objAS.ArriveCityCode = ![CityCode]
            objAS.ArriveCityName = ![City]
            objAS.ArriveCountry = ![CountryCode]
            objAS.ArriveDiffGMT = ![DiffGMT]
            objAS.ArriveIATAArea = ![IATAArea] & ""
            objAS.ArriveRegion = ![Region] & ""
            Set rsCity = Nothing
        Else
            'MsgBox "Arrival City " & strArrCity & " not found in Lookup table."
            strMsg = "Arrival City " & strArrCity & " not found in Lookup table."
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        End If
        End With
         
        Set rsCity = gdbConn.Execute("SELECT * FROM tblCityCodes WHERE AirportCode = '" & strDepCity & "'")
         
       With rsCity
       If Not .EOF Then
            objAS.DepartCityCode = ![CityCode]
            objAS.DepartCityName = ![City]
            objAS.DepartCountry = ![CountryCode]
            objAS.DepartDiffGMT = ![DiffGMT]
            objAS.DepartIATAArea = ![IATAArea] & ""
            objAS.DepartRegion = ![Region] & ""
            Set rsCity = Nothing
        Else
            'MsgBox "Departure City " & strDepCity & " not found in Lookup table."
            strMsg = "Departure City " & strDepCity & " not found in Lookup table."
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        End If
        End With
        
   

        mcolAirSegs.Add objAS, "S" & strSegNum
    Next lngC1

End Sub

Private Function GetNumberFF() As Long
Dim strRequest As String
Dim strResponse As String
Dim xmldomFFL As MSXML2.DOMDocument
Dim xmlnlTemp1 As MSXML2.IXMLDOMNodeList

strRequest = strRequest & "<DocProdFareManipulation_3_0>"
   strRequest = strRequest & "<FareRedisplayMods>"
      strRequest = strRequest & "<DisplayAction>"
         strRequest = strRequest & "<Action>I</Action>"
      strRequest = strRequest & "</DisplayAction>"
   strRequest = strRequest & "</FareRedisplayMods>"
strRequest = strRequest & "</DocProdFareManipulation_3_0>"

Set xmldomFFL = New MSXML2.DOMDocument
strResponse = gobjHost.SendQuery(strRequest, "DocProdFareManipulation_3_0", "GalileoPNR.PNR", "GetNumberFF")
With xmldomFFL
    '.loadXML strRequest
    ' .save "c:\My Projects\XML Examples\FF\FF List Request_" & Format(Now(), "ddmmmyy_HhNn") & ".xml"
    .loadXML strResponse
   ' .save "c:\My Projects\XML Examples\FF\FF List Response_" & Format(Now(), "ddmmmyy_HhNn") & ".xml"
End With

GetNumberFF = xmldomFFL.getElementsByTagName("FareNumInfo").length

End Function
Public Sub LoadFiledFare(ByVal FiledFareNumber As Byte, Optional bolReplace As Boolean)
Dim strRequest As String
Dim strResponse As String
Dim strErrMsg As String
Dim xmldomFF As MSXML2.DOMDocument
Dim xmlnlTemp0 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp1 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp2 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp3 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp4 As MSXML2.IXMLDOMNodeList
Dim lngBaseAmt As Long
Dim intBaseDecPos As Integer
Dim lngEquivAmt As Long
Dim intEquivDecPos As Integer
Dim lngTotAmt As Long
Dim intTotDecPos As Integer
Dim lngSurAmt As Long
Dim intSurDecPos As Integer
Dim lngC0 As Long
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim lngC4 As Long
Dim objFF As FiledFare
Dim objTax As FiledFareTax
Dim objFFSeg As FiledFareSegment
Dim objPFC As PFC
Dim objFFC As FiledFareComponent
Dim objSC As FiledFareSurcharge
Dim bolPF As Boolean
Dim intPxNum As Integer
Dim i As Integer


'On Error GoTo ErrProc
strRequest = ""
strRequest = strRequest & "<DocProdFareManipulation_7>"
   strRequest = strRequest & "<FareRedisplayMods>"
      strRequest = strRequest & "<DisplayAction>"
         strRequest = strRequest & "<Action>D</Action>"
      strRequest = strRequest & "</DisplayAction>"
     strRequest = strRequest & "<FareNumInfo>"
        strRequest = strRequest & "<FareNumAry>"
           strRequest = strRequest & "<FareNum>" & FiledFareNumber & "</FareNum>"
        strRequest = strRequest & "</FareNumAry>"
     strRequest = strRequest & "</FareNumInfo>"
   strRequest = strRequest & "</FareRedisplayMods>"
strRequest = strRequest & "</DocProdFareManipulation_7>"

strResponse = gobjHost.SendQuery(strRequest, "DocProdFareManipulation_3_0", "GalileoPNR.PNR", "LoadFiledFare")

Set mPaxFF = New PxFF

Set xmldomFF = New MSXML2.DOMDocument

With xmldomFF
    .loadXML strRequest
   '  .save "c:\My Projects\XML Examples\FF\FF Request_" & Format(Now(), "ddmmmyy_HhNn") & ".xml"
    .loadXML strResponse
   ' .save "c:\My Projects\XML Examples\FF\FF Response_" & Format(Now(), "ddmmmyy_HhNn") & ".xml"
End With

If InStr(xmldomFF.xml, "TransactionErrorCode") <> 0 Then
    strErrMsg = "Transaction Error " & xmldomFF.getElementsByTagName("ErrorCode").Item(0).Text & vbCrLf _
        & xmldomFF.documentElement.nodeName & vbCrLf _
        & xmldomFF.getElementsByTagName("Text").Item(0).Text
    Err.Raise -100001, "CWT_GalileoPNR3.GetFiledFare", strErrMsg
    Exit Sub
ElseIf InStr(xmldomFF.xml, "UNABLE TO FARE QUOTE - FAILED RULES VALIDATION") Then
    strErrMsg = "UNABLE TO FARE QUOTE - FAILED RULES VALIDATION"
    Err.Raise -100001, "CWT_Galileo3.GetFareQuote", strErrMsg
    Exit Sub

End If
 
'Added on 23 Dec
Dim intPxAryNum As Integer
Dim intKey() As Integer
Dim intUKey As Integer
Dim intAIndex As Integer

Set xmlnlTemp1 = xmldomFF.selectNodes("//AgntEnteredPsgrDescInfo//ApplesToAry//AppliesTo//AbsNameNum") 'pxnum
Set xmlnlTemp2 = xmldomFF.selectNodes("//AgntEnteredPsgrDescInfo")
intPxNum = xmlnlTemp1.length
i = 0
ReDim intKey(intPxNum - 1)

ReDim intpxnamenum(intPxNum - 1)
    
lngC1 = 0
    
    For lngC2 = 0 To xmlnlTemp2.length - 1
        
    intUKey = CInt(xmlnlTemp2(lngC2).selectSingleNode("UniqueKey").Text)
        
        Set xmlnlTemp3 = xmldomFF.selectNodes("//AgntEnteredPsgrDescInfo//ApplesToAry").Item(lngC2).childNodes
            For lngC3 = 0 To xmlnlTemp3.length - 1
        
            Set xmlnlTemp4 = xmldomFF.selectNodes("//AgntEnteredPsgrDescInfo//ApplesToAry").Item(lngC2).childNodes.Item(lngC3).childNodes
                For lngC4 = 0 To xmlnlTemp4.length - 1
                    With xmlnlTemp4.Item(lngC4)
                        Select Case .nodeName
                            Case "AbsNameNum"
                                intpxnamenum(lngC1) = CInt(.Text)
                                intKey(lngC1) = intUKey
                                lngC1 = lngC1 + 1
                        End Select
                    End With
                Next lngC4
            Next lngC3
    Next lngC2

'For lngC2 = 0 To UBound(intKey)
'Debug.Print intPxNameNum(lngC2) & intKey(lngC2)
'Next lngC2

Dim intLB As Integer
Dim intUB As Integer
Dim intMaxIndex As Integer
Dim intTemp As Integer

    intLB = LBound(intpxnamenum)
    intUB = UBound(intpxnamenum)
    
    For lngC2 = intUB To intLB + 1 Step -1
        
        intMaxIndex = 0

        'Find the largest value
        For lngC3 = intLB To lngC2
            If intpxnamenum(lngC3) > intpxnamenum(intMaxIndex) Then intMaxIndex = lngC3
        Next lngC3
        
        
        intTemp = intpxnamenum(intMaxIndex)
        intpxnamenum(intMaxIndex) = intpxnamenum(lngC2)
        intpxnamenum(lngC2) = intTemp
       
        intTemp = intKey(intMaxIndex)
        intKey(intMaxIndex) = intKey(lngC2)
        intKey(lngC2) = intTemp
    Next lngC2

'For lngC2 = 0 To UBound(intKey)
'Debug.Print intPxNameNum(lngC2) & intKey(lngC2)
'Next lngC2

Set xmlnlTemp1 = Nothing
Set xmlnlTemp2 = Nothing
Set xmlnlTemp3 = Nothing
Set xmlnlTemp4 = Nothing

 'Set objFF = New FiledFare

'reset all variables
lngBaseAmt = 0
intBaseDecPos = 0
lngEquivAmt = 0
intEquivDecPos = 0
lngTotAmt = 0
intTotDecPos = 0

'1 Px 1 GenQuoteDetails, if 2 FF 2 Px then 4 GenQuoteDetails
Set xmlnlTemp0 = xmldomFF.selectNodes("//GenQuoteDetails")
'intPxNum = xmlnlTemp0.length

For i = 0 To intPxNum - 1
intAIndex = i
i = intKey(i) - 1

 Set objFF = New FiledFare
For lngC0 = 0 To xmlnlTemp0.length - 1
    If xmlnlTemp0(lngC0).selectSingleNode("UniqueKey").Text = i + 1 Then
       Set xmlnlTemp1 = xmlnlTemp0(lngC0).childNodes
      
       For lngC1 = 0 To xmlnlTemp1.length - 1
         With xmlnlTemp1.Item(lngC1)
            Select Case .nodeName
               Case "QuoteType"
                   objFF.QuoteType = .Text
               Case "LastTkDt"
                   If .Text <> "" Then objFF.LastTicketDate = ConvertDate(.Text)
               Case "QuoteDt"
                   If .Text <> "" Then objFF.CreatedDate = ConvertDate(.Text)
               Case "IntlSaleInd"
                   objFF.ISO = .Text
               Case "BaseFareCurrency"
                   objFF.BaseCurrency = .Text
               Case "BaseFareAmt"
                   lngBaseAmt = CLng(.Text)
               Case "BaseDecPos"
                   intBaseDecPos = CLng(.Text)
               Case "EquivCurrency"
                   objFF.EquivCurrency = .Text
               Case "EquivAmt"
                   lngEquivAmt = CLng(.Text)
               Case "EquivDecPos"
                   intEquivDecPos = CLng(.Text)
               Case "TotCurrency"
                   objFF.TotalCurrency = .Text
               Case "TotAmt"
                   lngTotAmt = CLng(.Text)
               Case "TotDecPos"
                   intTotDecPos = CLng(.Text)
               Case "ITNum"
                   objFF.ITNum = .Text
               Case "PrivFQd"
                   objFF.PrivateFare = IIf(.Text = "Y", True, False)
                   bolPF = IIf(.Text = "Y", True, False)
               Case Else
                   'nothing
            End Select
         End With
       Next lngC1
       objFF.BaseAmount = PlaceDecimal(intBaseDecPos, lngBaseAmt)
       objFF.EquivAmount = PlaceDecimal(intEquivDecPos, lngEquivAmt)
       objFF.TotAmount = PlaceDecimal(intTotDecPos, lngTotAmt)
       
       
       Set xmlnlTemp2 = xmlnlTemp0(lngC0).selectNodes("TaxDataAry//TaxData")
    For lngC1 = 0 To xmlnlTemp2.length - 1
        Set objTax = New FiledFareTax
        Set xmlnlTemp3 = xmlnlTemp2.Item(lngC1).childNodes
        For lngC2 = 0 To xmlnlTemp3.length - 1
                With xmlnlTemp3.Item(lngC2)
                    Select Case .nodeName
                        Case "Country"
                            objTax.TaxCode = .Text
                        Case "Amt"
                            objTax.Amount = CSng(IIf(IsNumeric(.Text), .Text, 0))
                    End Select
                End With
        Next lngC2
        objFF.TaxAdd objTax
        Set objTax = Nothing
    Next lngC1
       
    End If
Next


'Added by Ji Yong on 06 Dec to determine is this a Cat35 fare
lngBaseAmt = 0
intBaseDecPos = 0
lngEquivAmt = 0
intEquivDecPos = 0
lngTotAmt = 0
intTotDecPos = 0

Set xmlnlTemp1 = xmldomFF.selectNodes("//ExtendedQuoteInformation")
For lngC1 = 0 To xmlnlTemp1.length - 1
    If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
       Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
       For lngC2 = 0 To xmlnlTemp2.length - 1
           With xmlnlTemp2.Item(lngC2)
                Select Case .nodeName
                       Case "DataRetQuote"
                             objFF.Cat35 = IIf(.Text = "Y", True, False)
                       Case "NTypeComm"
                             objFF.CommType = IIf(.Text = "Y", True, False)
                       Case "TkFareCrcy"
                             objFF.TktBaseCurrency = .Text
                       Case "TkFareAmt"
                             lngBaseAmt = CLng(.Text)
                       Case "TkNumDecs"
                             intBaseDecPos = CInt(.Text)
                       Case "EquivTkCrncy"
                             objFF.TktEquivCurrency = .Text
                       Case "EquivTkAmt"
                             lngEquivAmt = CLng(.Text)
                       Case "EquivTkNumDecs"
                             intEquivDecPos = CInt(.Text)
                       Case "TotTkCrncy"
                             objFF.TktTotalCurrency = .Text
                       Case "TotTkAmt"
                             lngTotAmt = CLng(.Text)
                       Case "TotTkNumDecs"
                             intTotDecPos = CInt(.Text)
                End Select
           End With
        Next lngC2
    End If
Next lngC1
objFF.TktBaseAmount = PlaceDecimal(intBaseDecPos, lngBaseAmt)
objFF.TktEquivAmount = PlaceDecimal(intEquivDecPos, lngEquivAmt)
objFF.TktTotAmount = PlaceDecimal(intTotDecPos, lngTotAmt)

Set xmlnlTemp1 = xmldomFF.selectNodes("//TkPsgrFareConstructTaxInfo")
For lngC1 = 0 To xmlnlTemp1.length - 1
  If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
          Set xmlnlTemp2 = xmlnlTemp1(lngC1).selectNodes("TaxDataAry//TaxData")
          For lngC2 = 0 To xmlnlTemp2.length - 1
              Set objTax = New FiledFareTax
              Set xmlnlTemp3 = xmlnlTemp2.Item(lngC2).childNodes
              For lngC3 = 0 To xmlnlTemp3.length - 1
                      With xmlnlTemp3.Item(lngC3)
                          Select Case .nodeName
                              Case "Country"
                                  objTax.TaxCode = .Text
                              Case "TaxAmt"
                                  If IsNumeric(.Text) = True Then
                                      objTax.Amount = CSng(.Text)
                                  Else
                                      objTax.Amount = 0
                                  End If
                                  'objTax.Amount = CSng(.Text)'decimal placement is same as total fare
                          End Select
                      End With
              Next lngC3
              objFF.TktTaxAdd objTax
              Set objTax = Nothing
          Next lngC2
  End If
Next lngC1



'1 only
Set xmlnlTemp1 = xmldomFF.selectNodes("//AdditionalFareInfo").Item(0).childNodes
    For lngC1 = 0 To xmlnlTemp1.length - 1
        With xmlnlTemp1.Item(lngC1)
            Select Case .nodeName
                Case "AgntSine"
                    objFF.CreatedBy = .Text
            End Select
        End With
    Next lngC1
    
'Change on 23Dec due to order of passenger for //AdditionalPsgrFareInfo
'1 Px 1 GenQuoteDetails, if 2 FF 2 Px then 4 GenQuoteDetails
Set xmlnlTemp4 = xmldomFF.selectNodes("//AdditionalPsgrFareInfo")
For lngC4 = 0 To xmlnlTemp4.length - 1
    If xmlnlTemp4(lngC4).selectSingleNode("AbsNameNum").Text = intpxnamenum(intAIndex) Then
        'modified on 11/05/05: Assign Pax ID
        objFF.PassengerNum = intpxnamenum(intAIndex)
        Set xmlnlTemp1 = xmlnlTemp4(lngC4).childNodes
        'Set xmlnlTemp1 = xmldomFF.selectNodes("//AdditionalPsgrFareInfo").Item(i).childNodes
        For lngC1 = 0 To xmlnlTemp1.length - 1
            With xmlnlTemp1.Item(lngC1)
                Select Case .nodeName
                    Case "Status"
                        objFF.Ticketed = IIf(.Text = "T", True, False)
                    Case "TkNum"
                        objFF.TicketNumber = .Text
                End Select
            End With
        Next lngC1
        Exit For
    End If
Next lngC4

'1 only
Set xmlnlTemp1 = xmldomFF.selectNodes("//CreditCardFOP|//CheckFOP|//OtherFOP")
For lngC1 = 0 To xmlnlTemp1.length - 1
    Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
    For lngC2 = 0 To xmlnlTemp2.length - 1
        With xmlnlTemp2.Item(lngC2)
            Select Case .nodeName
                'ZhiSam - 15 Nov 2012 - V1.3.3 - Smart Point Filed Fare Auto Cancellation - Add FOPID case for objFF
                Case "ID", "FOPID"
                    Select Case .Text
                        Case "1"
                            objFF.FOPType = "S"
                        Case "2"
                            objFF.FOPType = "CK"
                        Case "3"
                            objFF.FOPType = "NR"
                        Case "4"
                            objFF.FOPType = "MS"
                        Case "5"
                            objFF.FOPType = "INV"
                        Case "6"
                            objFF.FOPType = "CC"
                        Case "10"
                            objFF.FOPType = "CC"
                        Case "13"
                            objFF.FOPType = "EX"
                        Case "17"
                            objFF.FOPType = "NR"
                        Case "99"
                            objFF.FOPType = "FF"
                    End Select
                 Case "ExpDt"
                    objFF.FOP_CCExpireDate = CDate(Left(.Text, IIf(Len(.Text) = 3, 1, 2)) & "/01/" & Right(.Text, 2))
                Case "Vnd"
                    objFF.FOP_CCCode = .Text
                Case "Acct"
                    objFF.FOP_CCNum = Trim(.Text)
            End Select
        End With
        Next lngC2
Next lngC1
        
''1 Px 1 GenQuoteDetails, if 2 FF 2 Px then 4 TaxDataAry
'Set xmlnlTemp0 = xmldomFF.selectNodes("//TaxDataAry")
'Set xmlnlTemp1 = xmldomFF.selectNodes("//TaxDataAry//TaxData")
'Set xmlnlTemp1 = xmlnlTemp0..selectNodes("//TaxDataAry//TaxData")
'    For lngC1 = 0 To xmlnlTemp1.length - 1
'        Set objTax = New FiledFareTax
'        Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
'        For lngC2 = 0 To xmlnlTemp2.length - 1
'                With xmlnlTemp2.Item(lngC2)
'                    Select Case .nodeName
'                        Case "Country"
'                            objTax.TaxCode = .Text
'                        Case "Amt"
'                            objTax.Amount = CSng(IIf(IsNumeric(.Text), .Text, 0))
'                    End Select
'                End With
'        Next lngC2
'        objFF.TaxAdd objTax
'        Set objTax = Nothing
'    Next lngC1
 
'1 Px 1 GenQuoteDetails, if 2 FF 2 Px then 4 GenQuoteDetails
Set xmlnlTemp1 = xmldomFF.selectNodes("//SegRelatedInfo")

'Set xmlnlTemp1 = xmlnlTemp0.selectNodes("//SegRelatedInfo")

'For lngC0 = 0 To xmlnlTemp0.length - 1
'    If xmlnlTemp0(lngC0).selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
'       'Set xmlnlTemp1 = xmlnlTemp0(lngC0).childNodes

For lngC1 = 0 To xmlnlTemp1.length - 1
    If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
    Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
    Set objFFSeg = New FiledFareSegment
    For lngC2 = 0 To xmlnlTemp2.length - 1
    With xmlnlTemp2.Item(lngC2)
            Select Case .nodeName
                Case "NotValidBeforeDt"
                    If .Text <> "" Then objFFSeg.NVA = ConvertDate(.Text)
                Case "NotValidAfterDt"
                    If .Text <> "" Then objFFSeg.NVA = ConvertDate(.Text)
                Case "Stopover"
                    'objFFSeg.Stopover = IIf(.Text = "X", False, True)
                Case "FIC"
                    objFFSeg.FBC = .Text
                Case "TkDesignator"
                    objFFSeg.TD = .Text
                Case "BagInfo"
                    objFFSeg.BagInfo = .Text
            End Select
        End With
    Next lngC2
    objFF.FareSegAdd objFFSeg
    End If
Next lngC1
    'objFF.FareSegAdd objFFSeg
    'End If
'Next lngC0

'Seg Num
Set xmlnlTemp1 = xmldomFF.selectNodes("//AssocSegs//SegNumAry//SegNum")
For lngC1 = 0 To xmlnlTemp1.length - 1
    objFF.FareSeg(lngC1 + 1).AssocSeg = xmlnlTemp1.Item(lngC1).Text
Next lngC1

'???
Set xmlnlTemp1 = xmldomFF.selectNodes("//PsgrFacilityCharge//PFCAry")
For lngC1 = 0 To xmlnlTemp1.length - 1
    If xmlnlTemp1(lngC1).parentNode.selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
    Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
    Set objPFC = New PFC
    For lngC2 = 0 To xmlnlTemp2.length - 1
    With xmlnlTemp2.Item(lngC2)
            Select Case .nodeName
                Case "Airp"
                    objPFC.Airport = .Text
                Case "Amt"
                   objPFC.Amount = .Text
                Case "Currency"
                    objPFC.CurrencyCode = .Text
            End Select
        End With
    Next lngC2
    objFF.PFCAdd objPFC
    Set objPFC = Nothing
    End If
Next lngC1


Set xmlnlTemp3 = Nothing
Set xmlnlTemp4 = Nothing

'???
Set xmlnlTemp1 = xmldomFF.selectNodes("//InfoMsg//Text")
With xmlnlTemp1
For lngC1 = 0 To .length - 1
    If xmlnlTemp1(lngC1).parentNode.selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
       objFF.InfoMsgAdd (.Item(lngC1).Text)
    End If
Next
End With

'Pax number
Set xmlnlTemp1 = xmldomFF.selectNodes("//FareConstruction/FareConstructText")
With xmlnlTemp1
For lngC1 = 0 To .length - 1
    If xmlnlTemp1(lngC1).parentNode.selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
       objFF.FareConstructText = .Item(lngC1).Text
    End If
Next
End With
'objFF.FareConstructText = xmldomFF.selectSingleNode("//FareConstruction/FareConstructText").Text

'???'not in XML
Set xmlnlTemp1 = xmldomFF.selectNodes("//Surcharge")
For lngC1 = 0 To xmlnlTemp1.length - 1
    Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
    Set objSC = New FiledFareSurcharge
    For lngC2 = 0 To xmlnlTemp2.length - 1
        With xmlnlTemp2.Item(lngC2)
            Select Case .nodeName
                Case "ISGRIRelatedSecItinSegNum"
                    objSC.RelatedSegment = .Text
                Case "Amt"
                    lngTotAmt = CLng(.Text)
                Case "Currency"
                    objSC.CurrencyCode = .Text
                Case "DecPos"
                    intTotDecPos = CLng(.Text)
            End Select
        End With
    Next lngC2
    objSC.Amount = PlaceDecimal(intTotDecPos, lngTotAmt)
    objFF.SurchargeAdd objSC
    Set objSC = Nothing
Next lngC1
            
'mcolFiledFares.Add objFF


mPaxFF.PXAdd objFF


Set xmlnlTemp1 = Nothing
Set xmlnlTemp2 = Nothing
Set xmlnlTemp3 = Nothing
Set xmlnlTemp4 = Nothing
Set objFF = Nothing

i = intAIndex
Next i
If bolReplace = False Then
   mcolFiledFares.Add mPaxFF
Else
   mcolFiledFares.Remove FiledFareNumber
   If FiledFareNumber = 1 Then
      If mcolFiledFares.Count > 0 Then
         Set mcolFiledFares = Nothing
         Set mcolFiledFares = New Collection
      End If
      mcolFiledFares.Add mPaxFF
   Else
      mcolFiledFares.Add mPaxFF, , , FiledFareNumber - 1
   End If
End If
Exit Sub

ErrProc:
    Err.Raise Err.Number, "CWT_Galileo3.GetFareQuote", Err.Description
End Sub

Private Sub LoadFOP(ByVal FopNL As MSXML2.IXMLDOMNodeList)
Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlFOP As MSXML2.IXMLDOMNodeList
Dim strDate As String

Dim strMM As String
Dim strYY As String
Dim strFullDate As String

If FopNL.length = 0 Then
    mvarFOPType = ""
    mvarFOP_CCExpireDate = Now
    mvarFOP_CCCode = ""
    mvarFOP_CCNum = ""
End If
For lngC1 = 0 To FopNL.length - 1
    Set xmlnlFOP = FopNL.Item(lngC1).childNodes
    For lngC2 = 0 To xmlnlFOP.length - 1
        With xmlnlFOP.Item(lngC2)
            Select Case .nodeName
                Case "ID"
                    Select Case .Text
                        Case "1"
                            mvarFOPType = "S"
                        Case "2"
                            mvarFOPType = "CK"
                        Case "3"
                            mvarFOPType = "NR"
                        Case "4"
                            mvarFOPType = "MS"
                        Case "5"
                            mvarFOPType = "INV"
                        Case "6"
                            mvarFOPType = "CC"
                        Case "10"
                            mvarFOPType = "CC"
                        Case "13"
                            mvarFOPType = "EX"
                        Case "17"
                            mvarFOPType = "NR"
                        Case "99"
                            mvarFOPType = "FF"
                    End Select
                 Case "ExpDt"
                    If .Text <> "" Then
                        strDate = Format(.Text, "0000")
                        
                        strMM = Mid(strDate, 1, 2)
                        strYY = Right(strDate, 2)
                        strFullDate = "01/" & MMtoMMM(strMM) & "/" & strYY
                        If IsDate(strFullDate) = False Then
                            modMsgBox.OKMsg = "OK"
                            modMsgBox.sMsgBox gVPMDIHwnd, "Invalid FOP Expiry Date! Please update PNR (in *FOP).", vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                            strDate = "0101"
                        End If
                        
                        'If CLng(Left(strDate, 2)) > 12 Then
                        '     modMsgBox.OKMsg = "OK"
                        '     modMsgBox.sMsgBox gVPMDIHwnd, "Invalid FOP Expiry Date! Please update PNR (in *FOP).", vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                        '
                        'End If
                        
                        strDate = MonthName(CLng(Left(strDate, 2))) & "/25/" & Right(strDate, 2)
                        mvarFOP_CCExpireDate = DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", CDate(strDate)) + 1, CDate(strDate))))
                        
                    End If
                Case "Vnd"
                    mvarFOP_CCCode = .Text
                Case "Acct"
                    mvarFOP_CCNum = Trim(.Text)
            End Select
        End With
        Next lngC2
Next lngC1
End Sub

Private Function MMtoMMM(MM As String) As String
    Dim intMM As Integer
    
    If IsNumeric(MM) Then
        intMM = Int(MM)
    Else
        MMtoMMM = ""
        Exit Function
    End If
    
    Select Case intMM
        Case 1: MMtoMMM = "Jan"
        Case 2: MMtoMMM = "Feb"
        Case 3: MMtoMMM = "Mar"
        Case 4: MMtoMMM = "Apr"
        Case 5: MMtoMMM = "May"
        Case 6: MMtoMMM = "Jun"
        Case 7: MMtoMMM = "Jul"
        Case 8: MMtoMMM = "Aug"
        Case 9: MMtoMMM = "Sep"
        Case 10: MMtoMMM = "Oct"
        Case 11: MMtoMMM = "Nov"
        Case 12: MMtoMMM = "Dec"
        Case Else
            MMtoMMM = ""
    End Select
End Function

Private Sub LoadAddr(ByVal AddrNL As MSXML2.IXMLDOMNodeList)
Dim lngC1 As Long

For lngC1 = 0 To AddrNL.length - 1
    With AddrNL.Item(lngC1)
        Select Case .nodeName
            Case "Addr"
                mstrBillAddress = .Text
            Case "DeliveryAddr"
                mstrDelivAddress = .Text
        End Select
    End With
Next lngC1
End Sub

Public Property Get DeliveryAddress() As String
    DeliveryAddress = mstrDelivAddress
End Property

Public Property Get BillingAddress() As String
    BillingAddress = mstrBillAddress
End Property


Private Sub LoadHtlSegs(ByVal HtlSegsNL As MSXML2.IXMLDOMNodeList)
Dim objHS As HotelSeg
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim lngNext As Long
Dim xmlnlHS As MSXML2.IXMLDOMNodeList
Dim xmlnlOF As MSXML2.IXMLDOMNodeList
Dim xmlnlFlds As MSXML2.IXMLDOMNodeList
Dim lngTempAmt As Long
Dim intNumDec As Integer
Dim strOF As String
Dim strText() As String
Dim strAdd() As String
Dim lngPos As Long
Dim strTemp As String
Dim blnErr As Boolean
Dim strMsg As String
'CC - V1.2.7 20111011 - CR114 - Enable Hotel MI for HBU (Fix bug)
Dim xmlHS As MSXML2.DOMDocument

If Not mcolHtlSegs Is Nothing Then Set mcolHtlSegs = Nothing
Set mcolHtlSegs = New Collection

For lngC1 = 0 To HtlSegsNL.length - 1

If HtlSegsNL.Item(lngC1).nodeName = "HtlSeg" Then
    Set xmlnlHS = HtlSegsNL.Item(lngC1).childNodes
    Set objHS = New HotelSeg
            For lngC2 = 0 To xmlnlHS.length - 1
                With xmlnlHS.Item(lngC2)
                    Select Case .nodeName
                        Case "SegNum"
                            objHS.SegNum = ConvertNumber(.Text)
                        Case "Status"
                            objHS.Status = .Text
                        Case "StartDt"
                            objHS.CheckInDate = ConvertDate(.Text)
                        Case "EndDt"
                            objHS.CheckOutDate = ConvertDate(.Text)
                        Case "NumNights"
                            objHS.NumNights = ConvertNumber(.Text)
                        Case "Pt"
                            objHS.CityCode = .Text
                        Case "HtlV"
                            objHS.Vendor = .Text
                        Case "RateCode"
                            objHS.RoomType = .Text
                        Case "PropName"
                            objHS.PropertyName = .Text
                        Case "HtlPropNum"
                            objHS.PropertyNum = .Text
                        Case "NumPersons"
                            objHS.NumPersons = ConvertNumber(.Text)
                        Case "NumRooms"
                            objHS.NumRooms = ConvertNumber(.Text)
                        Case "ConfNum"
                            objHS.ConfNum = .Text
                        Case "RateTypeSold"
                            objHS.RateStatus = .Text
                        Case "Currency"
                            objHS.RateCurrency = .Text
                        Case "RateDecPos"
                            intNumDec = CLng(.Text)
                        Case "RateAmt"
                            lngTempAmt = CLng(.Text)
                        Case "MultilevelRateCode"
                            objHS.NegRateCode = .Text
                        
                    End Select
                End With
            Next lngC2
        objHS.RateAmount = PlaceDecimal(intNumDec, lngTempAmt)
    
        lngNext = lngC1 + 1
        If lngNext < HtlSegsNL.length Then
        If HtlSegsNL.Item(lngNext).nodeName = "HtlSegOptFlds" Then
            lngC1 = lngNext
            'CC - V1.2.7 20111011 - CR114 - Enable Hotel MI for HBU (Fix bug)
            'Set xmlnlOF = HtlSegsNL.Item(lngNext).selectNodes("//Fld")
            Set xmlHS = New MSXML2.DOMDocument
            xmlHS.loadXML (HtlSegsNL.Item(lngNext).xml)
            Set xmlnlOF = xmlHS.selectNodes("//Fld")
            'Set xmlnlOF = HtlSegsNL.Item(lngC1).childNodes
            
                For lngC2 = 0 To xmlnlOF.length - 1
                    Set xmlnlFlds = xmlnlOF(lngC2).childNodes
                    For lngC3 = 0 To xmlnlFlds.length - 1
                    With xmlnlFlds.Item(lngC3)
                        Select Case .nodeName
                            Case "ID"
                                strOF = .Text
                            Case "Contents"
                                Select Case strOF
                                    Case "RT"
                                        objHS.RateStatus = "RT"
                                        objHS.RateAmount = .Text
                                    Case "GT"
                                        objHS.Guar = .Text
                                    Case "DP"
                                        objHS.Deposit = .Text
                                    Case "RA"
                                        objHS.Rollaway = objHS.Rollaway + CLng(.Text)
                                    Case "RC"
                                        objHS.Rollaway = objHS.Rollaway + CLng(.Text)
                                    Case "CR"
                                        objHS.Crib = .Text
                                    Case "EX"
                                        objHS.ExtraPerson = objHS.ExtraPerson + CLng(.Text)
                                    Case "EC"
                                        objHS.ExtraPerson = objHS.ExtraPerson + CLng(.Text)
                                    Case "CD"
                                        objHS.CorpDiscNum = .Text
                                    Case "FG"
                                        objHS.FreqGuest = .Text
                                    Case "FT"
                                        objHS.FreqTvlr = .Text
                                    Case "SI"
                                        objHS.ServInfo = .Text
                                    Case "NF"
                                        objHS.GuestName = .Text
                                    Case "NL"
                                        objHS.GuestName = objHS.GuestName & " " & .Text
                                End Select
                        End Select
                    End With
                    Next lngC3
                Next lngC2
            End If
        End If
        
    
        mcolHtlSegs.Add objHS
    

Else
    Set xmlnlHS = HtlSegsNL.Item(lngC1).childNodes
    Set objHS = New HotelSeg
            For lngC2 = 0 To xmlnlHS.length - 1
                With xmlnlHS.Item(lngC2)
                    Select Case .nodeName
                        Case "SegNum"
                            objHS.SegNum = ConvertNumber(.Text)
                        Case "Status"
                            objHS.Status = .Text
                        Case "StartDt"
                            objHS.CheckInDate = ConvertDate(.Text)
                        Case "EndDt"
                            objHS.CheckOutDate = ConvertDate(.Text)
                        Case "NumNights"
                            objHS.NumNights = ConvertNumber(.Text)
                        Case "StartPt"
                            objHS.CityCode = .Text
                        Case "Vnd"
                            objHS.Vendor = .Text
                        Case "NumPersons"
                            objHS.NumPersons = ConvertNumber(.Text)
                        Case "Text"
                            strText = Split(.Text, "/")
                            objHS.RoomType = strText(0) ' room type is always first item
                            For lngC3 = 1 To UBound(strText)
                                lngPos = InStr(1, strText(lngC3), "-")
                                If lngPos Then
                                    Select Case Left(strText(lngC3), lngPos - 1)
                                        Case "W"
                                            strAdd = Split(Mid(strText(lngC3), lngPos + 1), "*")
                                            objHS.PropertyName = strAdd(0)
                                        Case "CD"
                                            objHS.CorpDiscNum = Mid(strText(lngC3), lngPos + 1)
                                        Case "RA"
                                            objHS.Rollaway = objHS.Rollaway + CLng(Mid(strText(lngC3), lngPos + 1))
                                        Case "RC"
                                            objHS.Rollaway = objHS.Rollaway + CLng(Mid(strText(lngC3), lngPos + 1))
                                        Case "CR"
                                            objHS.Crib = CLng(Mid(strText(lngC3), lngPos + 1))
                                        Case "EX"
                                            objHS.ExtraPerson = objHS.ExtraPerson + CLng(Mid(strText(lngC3), lngPos + 1))
                                        Case "EC"
                                            objHS.ExtraPerson = objHS.ExtraPerson + CLng(Mid(strText(lngC3), lngPos + 1))
                                        Case "FT"
                                            objHS.FreqTvlr = Mid(strText(lngC3), lngPos + 1)
                                        Case "FG"
                                            objHS.FreqGuest = Mid(strText(lngC3), lngPos + 1)
                                        Case "SI"
                                            objHS.ServInfo = Mid(strText(lngC3), lngPos + 1)
                                        Case "CF"
                                            objHS.ConfNum = Mid(strText(lngC3), lngPos + 1)
                                        Case "G"
                                            objHS.Guar = Mid(strText(lngC3), lngPos + 1)
                                        Case "NM"
                                            objHS.GuestName = Mid(strText(lngC3), lngPos + 1)
                                        Case "RT", "RQ", "RG"
                                            objHS.RateStatus = "RT"
                                            blnErr = False
                                            If IsNumeric(Mid(strText(lngC3), lngPos + 1)) Then
                                                objHS.RateAmount = Mid(strText(lngC3), lngPos + 1)
                                            'JY - V1.2.3 20110318 - CR52 - Change to include Hotel Booking Tool in the process
                                            'Change the sequence because it is causing error if rate is less than 100. Example RT-SGD10
                                            ElseIf IsNumeric(Mid(strText(lngC3), lngPos + 4)) = True Then
                                                objHS.RateCurrency = Mid(strText(lngC3), lngPos + 1, 3)
                                                objHS.RateAmount = Mid(strText(lngC3), lngPos + 4)
                                            ElseIf IsNumeric(Mid(strText(lngC3), Len(strText(lngC3)) - 2)) = False Then
                                                objHS.RateCurrency = Mid(strText(lngC3), Len(strText(lngC3)) - 2)
                                                strTemp = Left(strText(lngC3), Len(strText(lngC3)) - 3)
                                                If IsNumeric(Mid(strTemp, lngPos + 1)) Then
                                                    objHS.RateAmount = Mid(strTemp, lngPos + 1)
                                                Else
                                                    blnErr = True
                                                End If
                                            Else
                                               blnErr = True
                                            End If
                                            
                                            If blnErr = True Then
                                                'MsgBox "Currency format " & strText(lngC3) & " is not correct. The currency must be enter before the rate. ", vbOKOnly, "Hotel Segment Warning"
                                                strMsg = "Currency format " & strText(lngC3) & " is not correct. The currency must be enter before the rate. "
                                                modMsgBox.OKMsg = "OK"
                                                modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                                                objHS.RateAmount = 0
                                            End If
                                    End Select
                                End If
                            Next
                        End Select
                    End With
                Next
        mcolHtlSegs.Add objHS

End If
Next
End Sub

'TUR
Private Sub LoadTurSegs(ByVal TURSegsNL As MSXML2.IXMLDOMNodeList)
Dim objTS As TurSeg
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim lngNext As Long
Dim xmlnlTS As MSXML2.IXMLDOMNodeList
Dim xmlnlOF As MSXML2.IXMLDOMNodeList
Dim xmlnlFlds As MSXML2.IXMLDOMNodeList
Dim lngTempAmt As Long
Dim intNumDec As Integer
Dim strOF As String
Dim strText() As String
Dim strAdd() As String
Dim lngPos As Long
Dim strTemp As String
Dim blnErr As Boolean

If Not mcolTurSegs Is Nothing Then Set mcolTurSegs = Nothing
Set mcolTurSegs = New Collection

For lngC1 = 0 To TURSegsNL.length - 1

If TURSegsNL.Item(lngC1).nodeName = "NonAirSeg" Then
    Set xmlnlTS = TURSegsNL.Item(lngC1).childNodes
    Set objTS = New TurSeg
            For lngC2 = 0 To xmlnlTS.length - 1
                With xmlnlTS.Item(lngC2)
                    Select Case .nodeName
                        Case "SegNum"
                            objTS.SegNum = .Text
                        Case "Status"
                            objTS.Status = .Text
                        Case "Type"
                            objTS.SegType = .Text
                        Case "StartDt"
                            objTS.StartDt = ConvertDate(.Text)
                        Case "EndDt"
                            'objTS.EndDt = ConvertDate(.Text)
                        Case "Vnd"
                            objTS.Vnd = .Text
                        Case "NumPersons"
                            objTS.NumPersons = IIf(IsNumeric(.Text), .Text, 0)
                        Case "NumNights"
                            objTS.NumNight = IIf(IsNumeric(.Text), .Text, 0)
                        Case "SatrtPt"
                            objTS.StartPt = .Text
                        Case "EndPt"
                            objTS.EndPt = .Text
                        Case "SellType"
                            objTS.SellType = .Text
                        Case "Text"
                            objTS.Text = .Text
                        
                    End Select
                End With
            Next lngC2
            mcolTurSegs.Add objTS
Else

End If
Next

End Sub

Public Function CarSeg(Index As Variant) As CarSeg
    Set CarSeg = mcolCarSegs.Item(Index)
End Function

Private Sub LoadCarSegs(ByVal CarSegsNL As MSXML2.IXMLDOMNodeList)
Dim objCS As CarSeg
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim lngNext As Long
Dim xmlnlCS As MSXML2.IXMLDOMNodeList
Dim xmlnlOF As MSXML2.IXMLDOMNodeList
Dim xmlnlFlds As MSXML2.IXMLDOMNodeList
Dim strOF As String
Dim strStartDate As String
Dim strStartTime As String
Dim strEndDate As String
Dim strEndTime As String
Dim strFirstName As String
Dim strLastName As String
Dim lngDecimal As Long
Dim lngRateAmt As Long
Dim strTemp As String
If Not mcolCarSegs Is Nothing Then Set mcolCarSegs = Nothing
Set mcolCarSegs = New Collection
For lngC1 = 0 To CarSegsNL.length - 1

If CarSegsNL.Item(lngC1).nodeName = "NonAirSeg" Then
    Set xmlnlCS = CarSegsNL.Item(lngC1).childNodes
    Set objCS = New CarSeg
            For lngC2 = 0 To xmlnlCS.length - 1
                With xmlnlCS.Item(lngC2)
                    Select Case .nodeName
                        Case "SegNum"
                            objCS.SegNum = .Text
                        Case "Status"
                            objCS.Status = .Text
                        Case "Type"
                            objCS.SegType = .Text
                        Case "StartDt"
                            If .Text <> "" Then objCS.StartDtTime = ConvertDate(.Text)
                        Case "EndDt"
                            If .Text <> "" Then objCS.EndDtTime = ConvertDate(.Text)
                        Case "Vnd"
                            objCS.Vendor = .Text
                        Case "NumPersons"
                            objCS.NumPersons = IIf(IsNumeric(.Text), .Text, 0)
                        Case "NumNights"
                            objCS.NumNight = IIf(IsNumeric(.Text), .Text, 0)
                        Case "StartPt"
                            objCS.StartPt = .Text
                        Case "EndPt"
                            objCS.EndPt = .Text
                        Case "SellType"
                            objCS.SellType = .Text
                        Case "Text"
                            objCS.OtherText = .Text
                            If InStr(.Text, "RT-") Then
                                'If InStr(Mid(.Text, InStr(.Text, "CF-")), "/") - 1 < 0 Then
                                'strTemp = Mid(.Text, InStr(.Text, "RT-"))
                                'Else
                                strTemp = Mid(.Text, InStr(.Text, "RT-"), InStr(Mid(.Text, InStr(.Text, "RT-")), "/") - 1)
                                'End If
                                
                                If IsNumeric(Mid(strTemp, InStr(strTemp, "-") + 1 + 3)) = True Then
                                    objCS.RateAmt = Mid(strTemp, InStr(strTemp, "-") + 1 + 3)
                                    objCS.RateCurrency = Left(Mid(strTemp, InStr(strTemp, "-") + 1), 3)
                                End If
                            End If

                            If InStr(.Text, "CF-") Then
                            If InStr(Mid(.Text, InStr(.Text, "CF-")), "/") - 1 < 0 Then
                                strTemp = Mid(.Text, InStr(.Text, "CF-"))
                            Else
                                strTemp = Mid(.Text, InStr(.Text, "CF-"), InStr(Mid(.Text, InStr(.Text, "CF-")), "/") - 1)
                            End If
                            objCS.ConfNum = Mid(strTemp, InStr(strTemp, "-") + 1)
                            End If
                            If InStr(.Text, "/") > 0 Then
                            strTemp = Left(.Text, InStr(.Text, "/") - 1)
                            If Len(strTemp) > 5 Then
                            
                            If IsDate(Mid(Mid(strTemp, 1, 5), 1, 2) & "-" & Mid(Mid(strTemp, 1, 5), 3, 3)) Then objCS.EndDtTime = ConvertDate(Format(Mid(Mid(strTemp, 1, 5), 1, 2) & "-" & Mid(Mid(strTemp, 1, 5), 3, 3), "yyyymmdd"))
                            End If
                            If Len(strTemp) = 9 Then objCS.CarType = Mid(strTemp, 6, 4)
                            End If
                            
                    End Select
                End With
            Next lngC2
            
            mcolCarSegs.Add objCS
            
ElseIf CarSegsNL.Item(lngC1).nodeName = "CarSeg" Then
    Set xmlnlCS = CarSegsNL.Item(lngC1).childNodes
    Set objCS = New CarSeg
            For lngC2 = 0 To xmlnlCS.length - 1
                With xmlnlCS.Item(lngC2)
                    Select Case .nodeName
                        Case "SegNum"
                            objCS.SegNum = .Text
                        Case "Status"
                            objCS.Status = .Text
                        Case "StartDt"
                            strStartDate = .Text
                        Case "EndDt"
                            strEndDate = .Text
                        Case "StartTm"
                            strStartTime = .Text
                        Case "EndTm"
                            strEndTime = .Text
                        Case "CarV"
                            objCS.Vendor = .Text
                        Case "NumCars"
                            objCS.NumCars = IIf(IsNumeric(.Text), .Text, 0)
                        Case "Airp"
                            objCS.StartPt = .Text
                        Case "ActualStartPt"
                            objCS.ActualStartPt = .Text
                        Case "LocnCat"
                            objCS.LocationCat = .Text
                        Case "LocnNum"
                            objCS.LocationNum = .Text
                        Case "EndPt"
                            objCS.EndPt = .Text
                        Case "ConfNum"
                            objCS.ConfNum = .Text
                        Case "RateType"
                            objCS.RateType = .Text
                        Case "RateCode"
                            objCS.RateCode = .Text
                        Case "RateCat"
                            objCS.RateCat = .Text
                        Case "CarType"
                            objCS.CarType = .Text
                        Case "Currency"
                            objCS.RateCurrency = .Text
                        Case "DecPos"
                            lngDecimal = IIf(.Text = "", 0, .Text)
                        Case "RateAmt"
                            lngRateAmt = IIf(.Text = "", 0, .Text)
                        Case "RateGuarInd"
                            objCS.RateGuarInd = .Text
                        Case "ExtraDayRateAmt"
                            objCS.ExtraDayRateAmt = .Text
                        
                    End Select
                End With
            Next lngC2
        
        
        lngNext = lngC1 + 1
        If lngNext < CarSegsNL.length Then
        If CarSegsNL.Item(lngNext).nodeName = "CarSegOptFlds" Then
            lngC1 = lngNext
            Set xmlnlOF = CarSegsNL.Item(lngNext).selectNodes("//Fld")
                For lngC2 = 0 To xmlnlOF.length - 1
                    Set xmlnlFlds = xmlnlOF(lngC2).childNodes
                    For lngC3 = 0 To xmlnlFlds.length - 1
                    With xmlnlFlds.Item(lngC3)
                        Select Case .nodeName
                            Case "ID"
                                strOF = .Text
                            Case "Contents"
                                Select Case strOF
                                    Case "RG"
                                        objCS.RG = .Text
                                    Case "BS"
                                        objCS.BS = .Text
                                    Case "NF"
                                        strFirstName = .Text
                                    Case "NL"
                                        strLastName = .Text
                                End Select
                        End Select
                    End With
                    Next lngC3
                Next lngC2
            End If
        End If
            
            objCS.PaxName = strFirstName & strLastName
            If strEndDate <> "" Then objCS.EndDtTime = ConvertDate(strEndDate, IIf(strEndTime = "", "0001", strEndTime))
        
            If strStartDate <> "" Then objCS.StartDtTime = ConvertDate(strStartDate, IIf(strStartTime = "", "0001", strStartTime))
            objCS.RateAmt = PlaceDecimal(lngDecimal, lngRateAmt)
            mcolCarSegs.Add objCS

End If
Next

End Sub
Public Property Get CarSegCount() As Long
    If Not mcolCarSegs Is Nothing Then
        CarSegCount = mcolCarSegs.Count
    Else
        CarSegCount = 0
    End If
End Property
Private Sub LoadSSRequests(ByVal NodeList As MSXML2.IXMLDOMNodeList)
Dim objSSR As SSR
Dim lngC1 As Long
Dim lngC2 As Long
Dim xmlnlSSR As MSXML2.IXMLDOMNodeList
Dim xmlnlSSRPsgr As MSXML2.IXMLDOMNodeList
Dim temp As MSXML2.IXMLDOMNode

If Not mcolSSRequests Is Nothing Then Set mcolSSRequests = Nothing
Set mcolSSRequests = New Collection

For lngC1 = 0 To NodeList.length - 1
    Set xmlnlSSR = NodeList.Item(lngC1).childNodes
    Set objSSR = New SSR
    For lngC2 = 0 To xmlnlSSR.length - 1
        With xmlnlSSR.Item(lngC2)
            Select Case .nodeName
                Case "GFAXNum"
                    objSSR.GFax = .Text
                Case "SSRCode"
                    objSSR.SSCode = .Text
                Case "Status"
                    objSSR.Status = .Text
                Case "SegNum"
                    objSSR.SegNum = .Text
                Case "AppliesToAry"
                    Set xmlnlSSRPsgr = .selectNodes("AppliesTo/AbsNameNum")
                    objSSR.PsgNum = xmlnlSSRPsgr.Item(0).Text
                    Set xmlnlSSRPsgr = Nothing
            End Select
        End With
    Next lngC2
    'Check next node is ProgramaticSSRText or not
    If lngC1 + 1 < NodeList.length Then
        If NodeList.Item(lngC1 + 1).nodeName = "ProgramaticSSRText" Then
           objSSR.Text = NodeList.Item(lngC1 + 1).childNodes(0).Text
          lngC1 = lngC1 + 1
        End If
    End If
    mcolSSRequests.Add objSSR
Next lngC1

Set xmlnlSSR = Nothing
Set objSSR = Nothing
End Sub
Public Function SSR(Index As Variant) As SSR
    Set SSR = mcolSSRequests.Item(Index)
End Function
Public Property Get SSRCount() As Long
    SSRCount = mcolSSRequests.Count
End Property
Private Sub LoadOSI(ByVal NodeList As MSXML2.IXMLDOMNodeList)
Dim objOSI As OSI
Dim lngC1 As Long
Dim lngC2 As Long

Dim xmlnlOSI As MSXML2.IXMLDOMNodeList
If Not mcolOSI Is Nothing Then Set mcolOSI = Nothing
Set mcolOSI = New Collection

For lngC1 = 0 To NodeList.length - 1
    Set xmlnlOSI = NodeList.Item(lngC1).childNodes
    Set objOSI = New OSI
    For lngC2 = 0 To xmlnlOSI.length - 1
        With xmlnlOSI.Item(lngC2)
            Select Case .nodeName
                Case "GFAXNum"
                    objOSI.GFax = .Text
                Case "OSIV"
                    objOSI.Vendor = .Text
                Case "OSIMsg"
                    objOSI.Text = .Text
            End Select
        End With
    Next
        mcolOSI.Add objOSI
Next

Set xmlnlOSI = Nothing
Set objOSI = Nothing
End Sub
Public Function OSI(Index As Variant) As OSI
    Set OSI = mcolOSI.Item(Index)
End Function
Public Property Get OSICount() As Long
    OSICount = mcolOSI.Count
End Property

Public Property Get SimiliarNamesCount() As Long
    SimiliarNamesCount = mcolSimiliarNames.Count
End Property

Public Function SimiliarName(Index As Variant) As SimiliarName
    Set SimiliarName = mcolSimiliarNames.Item(Index)
End Function

Private Function exactPCC(strPCC As String) As String
   Dim i As Integer
   
   For i = 1 To 4 - Len(strPCC)
       strPCC = "0" & strPCC
   Next
   exactPCC = strPCC
End Function

Private Sub LoadTktDate(ByVal TktDateNL As MSXML2.IXMLDOMNodeList)
Dim lngC1 As Long
Dim strTmp As String

For lngC1 = 0 To TktDateNL.length - 1
    With TktDateNL.Item(lngC1)
        Select Case .nodeName
            Case "Text"
              strTmp = .Text
              If Left(strTmp, 4) = "QSB " Then
                 mdteTktDate = Mid(.Text, 5, 5)
                 mblnTkted = True
              End If
            Case "QTAUDt"
              mdteTktDate = .Text
              mblnTkted = False
        End Select
    End With
Next lngC1
End Sub
'ZhiSam - V1.2.19 20130405 - CR-203 - Add LoadTAWDate for TAW xml segment
Private Sub LoadTAWDate(ByVal TktDateNL As MSXML2.IXMLDOMNodeList)
Dim lngC1 As Long
Dim strTmp As String

For lngC1 = 0 To TktDateNL.length - 1
    With TktDateNL.Item(lngC1)
        Select Case .nodeName
            Case "QTAWDt"
              mstrTAWDate = .Text
              mblnTkted = False
        End Select
    End With
Next lngC1
End Sub
'ZhiSam - V1.2.19 20130405 - CR-203 - Add LoadTAWDate for TAW xml segment
Public Property Let TAWDate(ByVal vData As String)
    mstrTAWDate = vData
End Property
'ZhiSam - V1.2.19 20130405 - CR-203 - Add LoadTAWDate for TAW xml segment
Public Property Get TAWDate() As String
    TAWDate = mstrTAWDate
End Property

Public Property Let TktDate(ByVal vData As String)
    mdteTktDate = vData
End Property

Public Property Get TktDate() As String
    TktDate = mdteTktDate
End Property
Public Property Let Tkted(ByVal vData As Boolean)
    mblnTkted = vData
End Property

Public Property Get Tkted() As Boolean
    Tkted = mblnTkted
End Property

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Public Function captureTktNum() As String

Dim mxmldomETk As MSXML2.DOMDocument
Dim ETkNL As MSXML2.IXMLDOMNodeList
Dim strTemp As String
Dim result As String
Dim mobjSQ As HostAccess.StructuredQuery
Dim mstrID As String
mstrID = "<Application><VendorId>XMDL</VendorId><VendorType>G</VendorType><SourceId>CARWPR</SourceId><SourceType>G</SourceType></Application>"


strTemp = ""
strTemp = strTemp & "  <DocProdFareManipulation_7>"
strTemp = strTemp & "   <TicketNumbersMods />"
strTemp = strTemp & "  </DocProdFareManipulation_7>"

Set mxmldomETk = New MSXML2.DOMDocument

Set mxmldomETk = CreateObject("microsoft.xmldom")
mxmldomETk.async = False

Set mobjSQ = New HostAccess.StructuredQuery
result = mobjSQ.ExecuteXMLQuery(strTemp, mstrID)
mxmldomETk.loadXML (result)

Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim ETkNL1 As MSXML2.IXMLDOMNodeList
Dim ETkNL2 As MSXML2.IXMLDOMNodeList
Dim ETkNL3 As MSXML2.IXMLDOMNodeList

Dim strAirV As String
Dim strTktNum As String
Dim strTkType As String
Dim strFTktNum As String

Set ETkNL1 = mxmldomETk.childNodes.Item(0).childNodes
strFTktNum = ""

For lngC1 = 0 To ETkNL1.length - 1
   Set ETkNL2 = ETkNL1.Item(lngC1).childNodes
      For lngC2 = 0 To ETkNL2.length - 1
         With ETkNL2.Item(lngC2)
            Select Case .nodeName
               Case "ETicketNum"
                  Set ETkNL3 = .childNodes
                  For lngC3 = 0 To ETkNL3.length - 1
                     Select Case ETkNL3.Item(lngC3).nodeName
                        Case "AirV"
                           strAirV = ETkNL3.Item(lngC3).Text
                        Case "FirstTkNum"
                           strTktNum = ETkNL3.Item(lngC3).Text
                        Case "TkType"
                           strTkType = ETkNL3.Item(lngC3).Text
                     End Select
                     If strTkType = "E" Then
                        strFTktNum = strFTktNum & IIf(strFTktNum = "", "", ";") & strAirV & strTktNum
                        strTkType = ""
                     End If
                  Next
            End Select
         End With
      Next
Next

captureTktNum = strFTktNum
End Function

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Private Sub LoadETicket(ByVal TkNum As String)
Dim mxmldomETk As MSXML2.DOMDocument
Dim ETkNL As MSXML2.IXMLDOMNodeList
Dim intC As Integer
Dim strTemp As String
'Dim objError As ErrorInfo
Dim result As String
Dim mobjSQ As HostAccess.StructuredQuery
Dim mstrID As String
mstrID = "<Application><VendorId>XMDL</VendorId><VendorType>G</VendorType><SourceId>CARWPR</SourceId><SourceType>G</SourceType></Application>"
'TkNum = "1769900890797"

Dim objETk As ETicket
Dim objCoupon As Coupon
Dim xmlnlETk As MSXML2.IXMLDOMNodeList
Dim ETkNLs As MSXML2.IXMLDOMNodeList
Dim baseDataNode As MSXML2.IXMLDOMNodeList
Dim baseDataNodes As MSXML2.IXMLDOMNodeList

Dim strTot As String


strTemp = ""
strTemp = strTemp & "  <ETicketRetrieval_1_0>"
strTemp = strTemp & "   <EticketRetrievalMods>"
strTemp = strTemp & "    <TicketNumber>"
strTemp = strTemp & "     <TktNum>" & TkNum & "</TktNum>"
strTemp = strTemp & "    </TicketNumber>"
strTemp = strTemp & "   </EticketRetrievalMods>"
strTemp = strTemp & "  </ETicketRetrieval_1_0>"

Set mxmldomETk = New MSXML2.DOMDocument

Set mxmldomETk = CreateObject("microsoft.xmldom")
mxmldomETk.async = False

Set mobjSQ = New HostAccess.StructuredQuery
result = mobjSQ.ExecuteXMLQuery(strTemp, mstrID) ' BrdnWs:Problem occured here Object variable or With block variable not set
mxmldomETk.loadXML (result)

Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim lngC4 As Long
Dim lngC5 As Long
Dim ETkNL1 As MSXML2.IXMLDOMNodeList
Dim ETkNL2 As MSXML2.IXMLDOMNodeList
Dim ETkNL3 As MSXML2.IXMLDOMNodeList
Dim ETkNL4 As MSXML2.IXMLDOMNodeList
Dim ETkNL5 As MSXML2.IXMLDOMNodeList
Dim RouteStart As String
Dim RouteEnd As String
Dim FlightNumber As String
Dim FareBasis As String
Dim TourCode As String
Dim strPaxName As String
Dim strPaxTitle As String
Dim strVnd As String
Dim strAcct As String
Dim strChrX As String
Dim i As Long

Set ETkNL1 = mxmldomETk.childNodes.Item(0).childNodes
For lngC1 = 0 To ETkNL1.length - 1
    Set ETkNL2 = ETkNL1.Item(lngC1).childNodes
    Set objETk = New ETicket
            For lngC2 = 0 To ETkNL2.length - 1
                With ETkNL2.Item(lngC2)
                    Select Case .nodeName
                        Case "ETicketNum"
                            Set ETkNL3 = .childNodes
                            For lngC3 = 0 To ETkNL3.length - 1
                               If ETkNL3.Item(lngC3).nodeName = "VndIssueDt" Then
                                   objETk.IssueDate = Mid(ETkNL3.Item(lngC3).Text, 7, 2) + " " + MonthName(Val(Mid(ETkNL3.Item(lngC3).Text, 5, 2))) + " " + Mid(ETkNL3.Item(lngC3).Text, 1, 4)
                               End If
                            Next
                        Case "AirSegInfo" 'BaggageAllow only take air segment info
                            Set ETkNL3 = .childNodes
                            Set objCoupon = New Coupon
                            For lngC3 = 0 To ETkNL3.length - 1
                               Select Case ETkNL3.Item(lngC3).nodeName
                                   Case "CouponStatus"
                                        'objETk.CouponStatus = ETkNL3.Item(lngC3).Text
                                        'objBA.CouponStatus = ETkNL3.Item(lngC3).Text
                                        objCoupon.CouponStatus = ETkNL3.Item(lngC3).Text
                                   Case "BIC" 'BaggageAllow
                                        objCoupon.BIC = ETkNL3.Item(lngC3).Text
                                   Case "Dt"
                                        objCoupon.Dt = ETkNL3.Item(lngC3).Text
                                   Case "StartTm" 'BaggageAllow
                                        objCoupon.StartTime = ETkNL3.Item(lngC3).Text
                                   Case "StartAirp" 'BaggageAllow
                                        'objETk.RouteStart = ETkNL3.Item(lngC3).Text
                                        objCoupon.StartAirport = ETkNL3.Item(lngC3).Text
                                   Case "EndAirp" 'BaggageAllow
                                        'objETk.RouteEnd = ETkNL3.Item(lngC3).Text
                                        objCoupon.EndAirport = ETkNL3.Item(lngC3).Text
                                   Case "FltNum"
                                        'objETk.FlightNumber = ETkNL3.Item(lngC3).Text
                                        objCoupon.FlightNumber = ETkNL3.Item(lngC3).Text
                                   Case "Status" 'BaggageAllow
                                        'objBA.PaxStatus = ETkNL3.Item(lngC3).Text
                                        objETk.PaxType = ETkNL3.Item(lngC3).Text
                                   Case "FIC" 'BaggageAllow
                                        'objETk.FareBasis = ETkNL3.Item(lngC3).Text
                                        objCoupon.FareBasis = ETkNL3.Item(lngC3).Text
                                   Case "AllowCnt" 'BaggageAllow
                                        objCoupon.AllowCnt = ETkNL3.Item(lngC3).Text
                                   Case "AllowInd" 'BaggageAllow
                                        objCoupon.AllowInd = ETkNL3.Item(lngC3).Text
                                   Case "NotValidAfterDt" 'BaggageAllow
                                         objCoupon.FareExpire = ETkNL3.Item(lngC3).Text
                                   Case "AirV"
                                         objCoupon.AirVendor = ETkNL3.Item(lngC3).Text
                               End Select
                            Next
                            objETk.Route = RouteEnd + " - " + RouteStart + " - " + RouteEnd
                            objETk.AddCoupon objCoupon
                        Case "FareConstruction"
                            Set ETkNL3 = .childNodes
                            For lngC3 = 0 To ETkNL3.length - 1
                               If ETkNL3.Item(lngC3).nodeName = "FareConstructText" Then
                                   objETk.FareCalc = ETkNL3.Item(lngC3).Text
                               End If
                            Next
                            'If Mid(TkNumOri, 14, 1) = "-" Then
                            '    objETk.TiktNo = TkNumOri
                            'Else
                                objETk.TiktNo = TkNum
                            'End If

                        Case "GenQuoteDetails"
                            Set ETkNL3 = .childNodes
                            For lngC3 = 0 To ETkNL3.length - 1
                               Select Case ETkNL3.Item(lngC3).nodeName
                                    Case "BaseFareAmt"
                                    Dim strBaseDecPos As String
                                        'To obtain decimal value
                                        If IsNumeric(ETkNL3.Item(lngC3 + 2).Text) Then
                                            strBaseDecPos = ETkNL3.Item(lngC3 + 2).Text
                                        Else
                                            strBaseDecPos = "0"
                                        End If
                                        Select Case (strBaseDecPos)
                                        Case "2"
                                            objETk.Fare = Format(ETkNL3.Item(lngC3).Text / 100, "Fixed") & " " & ETkNL3.Item(lngC3 - 1).Text
                                        Case "3"
                                            objETk.Fare = Format(ETkNL3.Item(lngC3).Text / 1000, "####0.000") & " " & ETkNL3.Item(lngC3 - 1).Text
                                        Case Else
                                            objETk.Fare = Format(ETkNL3.Item(lngC3).Text, "Fixed") + " " + ETkNL3.Item(lngC3 - 1).Text
                                        End Select
                                    Case "Tot"
                                        Dim strTotDecPos As String
                                        'To obtain decimal value
                                        If IsNumeric(ETkNL3.Item(lngC3 + 1).Text) Then
                                            strTotDecPos = ETkNL3.Item(lngC3 + 1).Text
                                        Else
                                            strTotDecPos = "0"
                                        End If
                                        Select Case (strTotDecPos)
                                        Case "2"
                                            objETk.Total = Format(ETkNL3.Item(lngC3).Text / 100, "Fixed") & " " & ETkNL3.Item(lngC3 - 1).Text
                                        Case "3"
                                            objETk.Total = Format(ETkNL3.Item(lngC3).Text / 1000, "####0.000") & " " & ETkNL3.Item(lngC3 - 1).Text
                                        Case Else
                                            objETk.Total = Format(ETkNL3.Item(lngC3).Text, "Fixed") + " " + ETkNL3.Item(lngC3 - 1).Text
                                        End Select
                                    Case "Taxdata"
                                       Set ETkNL4 = ETkNL3.Item(lngC3).childNodes
                                        For lngC4 = 0 To ETkNL4.length - 1
                                            Select Case lngC4
                                                Case 0
                                                objETk.Charges = ETkNL4.Item(lngC4).childNodes.Item(1).Text + " " + ETkNL4.Item(lngC4).childNodes.Item(0).Text
                                                Case 1
                                                objETk.Taxes = ETkNL4.Item(lngC4).childNodes.Item(1).Text + " " + ETkNL4.Item(lngC4).childNodes.Item(0).Text
                                                Case 2
                                                objETk.Fees = ETkNL4.Item(lngC4).childNodes.Item(1).Text + " " + ETkNL4.Item(lngC4).childNodes.Item(0).Text
                                            End Select
                                        Next
                                   End Select
                            Next
                        Case "PassengerInfo"
                            Set ETkNL3 = .childNodes
                            For lngC3 = 0 To ETkNL3.length - 1
                                Select Case ETkNL3.Item(lngC3).nodeName
                                Case "Name"
                                    objETk.PaxLName = ETkNL3.Item(lngC3).Text
                                Case "Title"
                                    objETk.PaxFName = ETkNL3.Item(lngC3).Text
                                End Select
                            Next
                            
                        Case "TourCode"
                            Set ETkNL3 = .childNodes
                            objETk.TourCode = ETkNL3.Item(0).Text
                        Case "CreditCardFOP" 'cc FOP
                           Set ETkNL3 = .childNodes
                           For lngC3 = 0 To ETkNL3.length - 1
                              Select Case ETkNL3.Item(lngC3).nodeName
                                 Case "Vnd"
                                    strVnd = ETkNL3.Item(lngC3).Text
                                 Case "Acct"
                                    strChrX = ""
                                    strAcct = Trim(ETkNL3.Item(lngC3).Text)
                                    If strVnd <> "" And strAcct <> "" Then
'                                       For i = 1 To Len(strAcct) - 4
'                                          strChrX = strChrX & "x"
'                                       Next
'                                       strAcct = Replace(strAcct, Left(strAcct, Len(strAcct) - 4), strChrX)
                                       'CC - V1.2.20 20130617 - IRXXX - it will generate error if XML return card number's less than 4 digits
                                       'Return full card number to Desktop user
                                       objETk.FOP = strVnd & strAcct
                                    End If
                              End Select
                           Next
                        Case "OtherFOP"
                           Set ETkNL3 = .childNodes
                           For lngC3 = 0 To ETkNL3.length - 1
                              Select Case ETkNL3.Item(lngC3).nodeName
                                 Case "AdditionalInfoAry"
                                 If ETkNL3.Item(lngC3).childNodes.length <> 0 Then 'None data validate
                                    Set ETkNL4 = ETkNL3.Item(lngC3).childNodes
                                    For lngC4 = 0 To ETkNL4.length - 1
                                       Set ETkNL5 = ETkNL4.Item(0).childNodes
                                       For lngC5 = 0 To ETkNL5.length - 1
                                          Select Case ETkNL5.Item(lngC5).nodeName
                                             Case "Dt"
                                                If ETkNL5.Item(lngC5).Text = "NVAGT" Then
                                                   objETk.FOP = "I" & ETkNL5.Item(lngC5).Text 'INVAGT
                                                Else
                                                   objETk.FOP = ETkNL5.Item(lngC5).Text
                                                End If
                                             End Select
                                       Next
                                    Next
                                 End If
                              End Select
                           Next
                    End Select
                End With
            Next
    Next
    mcolETicket.Add objETk
End Sub

'CC - V1.2.20 20130416 - CR212 - TF Per Passenger Logic
Public Function ETicket(Index As Integer) As ETicket
    Set ETicket = mcolETicket.Item(Index)
End Function

Public Property Get ETicketCount() As Integer
    ETicketCount = mcolETicket.Count
End Property

Public Property Get DSInfo() As DSInfo
    Set DSInfo = mvarDSInfo
End Property

'ZhiSam - V1.2.23 20130829 - CR-229 - Data Standardization Phase 1
Private Sub LoadDSInfo()
    
    Dim objGR As GenRemark
    Dim strRmkText As String
    Dim intI As Integer
    Dim dblPhaseNum As Double

    
    If Not mcolGenRmks Is Nothing Then
        For intI = 1 To mcolGenRmks.Count
           Set objGR = mcolGenRmks.Item(intI)
            If UCase(objGR.Qualifier) = "HJ" Then
                strRmkText = objGR.RemarkText
                If InStr(1, strRmkText, CStr(CstrDSFormat), 1) <> 0 Then
                    strRmkText = Trim(Replace$(strRmkText, CStr(CstrDSFormat), ""))
                    dblPhaseNum = CDbl(strRmkText)
                        With mvarDSInfo
                            .ItemNum = objGR.ItemNum
                            .Qualifier = objGR.Qualifier
                            .RemarkText = objGR.RemarkText
                            If Val(dblPhaseNum) > 0 Then
                                .DSPhaseNum = FormatNumber(dblPhaseNum, 2)
                            Else
                                .DSPhaseNum = 0
                            End If
                        End With
                    Exit For
                End If
                
            End If
        Next
    End If

     Set objGR = Nothing

End Sub
