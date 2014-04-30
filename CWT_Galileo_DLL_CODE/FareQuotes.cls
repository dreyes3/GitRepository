VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FareQuotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mcolFQ As FQ
Private mcolPX As Collection

'Private mvarPxCount As Integer

Private Sub Class_Initialize()
Set mcolFQ = New FQ
Set mcolPX = New Collection

End Sub

Private Sub Class_Terminate()
Set mcolFQ = Nothing
Set mcolPX = Nothing
End Sub

Public Property Get PxCount() As Long
    PxCount = mcolPX.Count
    
End Property
'Public Property Get PxCount() As String
'    PxCount = mvarPxCount
'End Property

'Public Property Let PxCount(ByVal vData As Date)
'    mvarLastTktDate = vData
'End Property

Public Function Item(Index As Variant) As FQ
Attribute Item.VB_UserMemId = 0
    'Set Item = mcolFQ.Item(Index)
    Set Item = mcolPX.Item(Index)
End Function

'Public Function AddFareQuote(NewFareQuote As FareQuote) As FareQuote
'mcolFQ.Add NewFareQuote
'Set NewFareQuote = Nothing
'
'End Function
'29122004
Public Function GetFareQuote(ByVal Passenger As String, ByVal SegSelected As Boolean, ByVal Segments As String, _
    CN As String, CarrierCodes As String, PlatCarrier As String, Optional PIC As String = "AD", _
    Optional OverrideClass As String, Optional OverrideSeg As String, Optional SkipAdult As Boolean, Optional Stopover As String, _
    Optional ExcTaxCtry As String, Optional PCC As String, Optional tktType As String) As FQ
    
Dim strRequest As String
Dim strResponse As String
Dim AryStrResponse() As String
Dim strSegs() As String
Dim strErrMsg As String
Dim strErrMsg2 As String
Dim xmldomFQ As MSXML2.DOMDocument
Dim xmldomFCRules As MSXML2.DOMDocument
Dim xmlnlTemp0 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp1 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp2 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp3 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp4 As MSXML2.IXMLDOMNodeList
Dim xmlnlTemp5 As MSXML2.IXMLDOMNodeList
Dim lngBaseAmt As Long
Dim intBaseDecPos As Integer
Dim lngEquivAmt As Long
Dim intEquivDecPos As Integer
Dim lngTotAmt As Long
Dim intTotDecPos As Integer
Dim lngSurAmt As Long
Dim intSurDecPos As Integer
Dim lngRuleFC As Long
Dim lngC0 As Long
Dim lngC1 As Long
Dim lngC2 As Long
Dim lngC3 As Long
Dim objPx As FQ
Dim objFQ As FareQuote
Dim objTax As Tax
Dim objFQSeg As FareSegment
Dim objFC As FareComponent
Dim objSC As Surcharge
Dim bolPF As Boolean
Dim rsRules As New ADODB.Recordset
Dim rsCity As New ADODB.Recordset
Dim rsContracts As New ADODB.Recordset
Dim STRSQL As String
Dim strPFCodes() As String
Dim lngPF As Long
Dim PFRec As Integer
Dim failNum As Integer
Dim strCarriers() As String
Dim strOthNotesText As String
Dim str1stEndorseText As String
Dim AryStrPx() As String
Dim intAcctNo As Integer
Dim i As Integer
Dim intTemp As Integer
Dim strFicPt As String
Dim strPFType() As String
Dim multiPFBC() As String
Dim strOverrideClass() As String
Dim intClassCount As Integer
Dim strStopOver() As String
Dim bolAfOthNote As Boolean
'Added by JiYong to capture actual net fare
Dim lngNetFareAmt As Long
Dim intNetFarePos As Integer
Dim strNetFareCrcy As String
Dim lngEquivNetFareAmt As Long
Dim intEquivNetFareDecPos As Integer
Dim strEquivNetFareCrcy As String
Dim strTotalNetFareCrcy As String

AryStrPx = Split(Passenger, ";")

If OverrideClass <> "" Then strOverrideClass = Split(OverrideClass, ";")
If Stopover <> "" Then strStopOver = Split(Stopover, ";")
If Segments <> "" Then strSegs = Split(Segments, ".")
strCarriers = Split(CarrierCodes, "/")

'29122004
If SkipAdult Then
   ReDim strPFCodes(0)
   ReDim strPFType(0)
   strPFCodes(0) = gstrPFCode
   strPFType(0) = "APF"
   PFRec = 0
   'Modified on 21/03/2005
   
   
   STRSQL = "SELECT PFCODE,FARETYPE FROM tblClientAirContracts WHERE [CN] = '" & CN & "' AND ("
   For lngC1 = 0 To UBound(strCarriers)
       STRSQL = STRSQL & IIf(InStr(STRSQL, "[Carrier]") > 0, "OR ", "") & " [Carrier] = '" & strCarriers(lngC1) & "' "
   Next
   STRSQL = STRSQL & ") and ((Faretype='APF' AND PFCode<>'" & gstrPFCode & "') or FareType='SQC') GROUP BY PFCODE,FARETYPE"
   'Set rsContracts = gdbConn.Execute(strsql)
   rsContracts.open STRSQL, gdbConn, adOpenKeyset, adLockReadOnly
Else

   
   STRSQL = "SELECT PFCODE,FARETYPE FROM tblClientAirContracts WHERE [CN] = '" & CN & "' AND ("
   For lngC1 = 0 To UBound(strCarriers)
       STRSQL = STRSQL & IIf(InStr(STRSQL, "[Carrier]") > 0, "OR ", "") & " [Carrier] = '" & strCarriers(lngC1) & "' "
   Next
   STRSQL = STRSQL & ") and ((Faretype='APF' AND PFCode<>'" & gstrPFCode & "') or FareType='SQC') GROUP BY PFCODE,FARETYPE"
   'Set rsContracts = gdbConn.Execute(strsql)
   rsContracts.open STRSQL, gdbConn, adOpenKeyset, adLockReadOnly
   
   With rsContracts

       PFRec = .RecordCount 'getCount(CN, strCarriers)
       ReDim strPFCodes(PFRec)
       ReDim strPFType(PFRec)
       strPFCodes(0) = gstrPFCode
       strPFType(0) = "APF" 'SGCWT is an APF
       If PFRec > 0 Then
          '.MoveLast
          '.MoveFirst
           For lngC1 = 1 To PFRec
               strPFCodes(lngC1) = ![PFCode] & ""
               strPFType(lngC1) = ![faretype] & ""
               .MoveNext
           Next
         '.MoveFirst
       End If
   End With
End If

ReDim AryStrResponse(PFRec)

failNum = 0



'For lngPF = 0 To PFRec

'strRequest = "<FareQuoteStandard_8_0>"
'strRequest = strRequest & "   <FairQuoteMods>"
'strRequest = strRequest & "     <PsgrMods>"
'strRequest = strRequest & "<PsgrAry>"

'For lngC1 = 0 To UBound(AryStrPx)
'    strRequest = strRequest & "<Psgr>"
'    strRequest = strRequest & "<LNameNum>" & Left(AryStrPx(lngC1), InStr(1, AryStrPx(lngC1), "*") - 1) & "</LNameNum>"
'    strRequest = strRequest & "<PsgrNum>1</PsgrNum>"
'    strRequest = strRequest & "<AbsNameNum>" & Left(AryStrPx(lngC1), InStr(1, AryStrPx(lngC1), "*") - 1) & "</AbsNameNum>"
'    strRequest = strRequest & "<PIC>" & IIf(Mid(AryStrPx(lngC1), InStr(1, AryStrPx(lngC1), "*") + 1) = "AD", "", Mid(AryStrPx(lngC1), InStr(1, AryStrPx(lngC1), "*") + 1)) & "</PIC>"
'    strRequest = strRequest & "<TIC></TIC>"
'    strRequest = strRequest & "</Psgr>"
'Next

            'strRequest = strRequest & "<Psgr>"
            'strRequest = strRequest & "<LNameNum></LNameNum>"
            'strRequest = strRequest & "<PsgrNum></PsgrNum>"
            'strRequest = strRequest & "<AbsNameNum>1</AbsNameNum>"
            'strRequest = strRequest & "<PIC>" & PIC & "</PIC>"
            'strRequest = strRequest & "<TIC></TIC>"
            'strRequest = strRequest & "</Psgr>"
            '
            'strRequest = strRequest & "<Psgr>"
            'strRequest = strRequest & "<LNameNum></LNameNum>"
            'strRequest = strRequest & "<PsgrNum></PsgrNum>"
            'strRequest = strRequest & "<AbsNameNum>1</AbsNameNum>"
            ''strRequest = strRequest & "<PIC>AD</PIC>"
            'strRequest = strRequest & "<TIC></TIC>"
            'strRequest = strRequest & "</Psgr>"
            'strRequest = strRequest & "<Psgr>"
            'strRequest = strRequest & "<LNameNum></LNameNum>"
            'strRequest = strRequest & "<PsgrNum></PsgrNum>"
            'strRequest = strRequest & "<AbsNameNum>2</AbsNameNum>"
            'strRequest = strRequest & "<PIC>C08</PIC>"
            'strRequest = strRequest & "<TIC></TIC>"
            'strRequest = strRequest & "</Psgr>"
            'strRequest = strRequest & "<Psgr>"
            'strRequest = strRequest & "<LNameNum></LNameNum>"
            'strRequest = strRequest & "<PsgrNum></PsgrNum>"
            'strRequest = strRequest & "<AbsNameNum>3</AbsNameNum>"
            'strRequest = strRequest & "<PIC>INFANT</PIC>"
            'strRequest = strRequest & "<TIC></TIC>"
            'strRequest = strRequest & "</Psgr>"

'strRequest = strRequest & "</PsgrAry>"
'strRequest = strRequest & "</PsgrMods>"



'If Len(PlatCarrier) > 0 Then
'    strRequest = strRequest & "<PlatingAirVMods>"
'    strRequest = strRequest & "   <PlatingAirV>" & PlatCarrier & "</PlatingAirV>"
'    strRequest = strRequest & "</PlatingAirVMods>"
'End If
'strRequest = strRequest & "        <DocProdFareType>"
'strRequest = strRequest & "            <Type>I</Type>"
'strRequest = strRequest & "        </DocProdFareType>"
'strRequest = strRequest & "        <SegSelection>"
'strRequest = strRequest & "            <ReqAirVPFs>Y</ReqAirVPFs>"
'strRequest = strRequest & "            <SegRangeAry>"
'If OverrideClass = "" Then
'If SegSelected = True Then

'    For lngC1 = 0 To UBound(strSegs)
'        strRequest = strRequest & "<SegRange>"
'        strRequest = strRequest & "<StartSeg>" & Format(strSegs(lngC1), 0) & "</StartSeg>"
'        strRequest = strRequest & "<EndSeg>" & Format(strSegs(lngC1), 0) & "</EndSeg>"
'        strRequest = strRequest & "<FareType>P</FareType>"
'        strRequest = strRequest & "<PFQual>"
'        strRequest = strRequest & "<CRSInd>1G</CRSInd>"
'        strRequest = strRequest & "<PCC>" & gstrHQPCC & "</PCC>"
'        strRequest = strRequest & "<AirV></AirV>"
'        strRequest = strRequest & "<Acct>" & strPFCodes(lngPF) & "</Acct>"
'        strRequest = strRequest & "<Contract></Contract>"
'        strRequest = strRequest & "<PublishedFaresInd>Y</PublishedFaresInd>"
'        strRequest = strRequest & "<Type>A</Type>"
'        strRequest = strRequest & "</PFQual>"
'        strRequest = strRequest & "</SegRange>"
'    Next lngC1
'Else
'    strRequest = strRequest & "                <SegRange>"
'    strRequest = strRequest & "                    <StartSeg>00</StartSeg>"
'    strRequest = strRequest & "                    <EndSeg>00</EndSeg>"
'    strRequest = strRequest & "                    <FareType>P</FareType>"
'    strRequest = strRequest & "                    <PFQual>"
'    strRequest = strRequest & "                        <CRSInd>1G</CRSInd>"
'    strRequest = strRequest & "                        <PCC>" & gstrHQPCC & "</PCC>"
'    strRequest = strRequest & "                        <AirV></AirV>"
'    strRequest = strRequest & "                        <Acct>" & strPFCodes(lngPF) & "</Acct>"
'    strRequest = strRequest & "                        <PublishedFaresInd>Y</PublishedFaresInd>"
'    strRequest = strRequest & "                        <Type>A</Type>"
'    strRequest = strRequest & "                    </PFQual>"
'    strRequest = strRequest & "                </SegRange>"
'End If
'Else
'    strSegs = Split(Segments, ".")
'    For lngC1 = 0 To UBound(strSegs)
'        strRequest = strRequest & "<SegRange>"
'        strRequest = strRequest & "<StartSeg>" & Format(strSegs(lngC1), 0) & "</StartSeg>"
'        strRequest = strRequest & "<EndSeg>" & Format(strSegs(lngC1), 0) & "</EndSeg>"
'        strRequest = strRequest & "<FareType>B</FareType>"
'        strRequest = strRequest & "<FIC>" & strOverrideClass(lngC1) & "</FIC>"
'        strRequest = strRequest & "<PFQual>"
'        strRequest = strRequest & "<CRSInd>1G</CRSInd>"
'        strRequest = strRequest & "<PCC>" & gstrHQPCC & "</PCC>"
'        strRequest = strRequest & "<AirV></AirV>"
'        strRequest = strRequest & "<Acct>" & strPFCodes(lngPF) & "</Acct>"
'        strRequest = strRequest & "<Contract></Contract>"
'        strRequest = strRequest & "<PublishedFaresInd>Y</PublishedFaresInd>"
'        strRequest = strRequest & "<Type>A</Type>"
'        strRequest = strRequest & "</PFQual>"
'        strRequest = strRequest & "</SegRange>"
'    Next lngC1
'End If



'strRequest = strRequest & "            </SegRangeAry>"
'strRequest = strRequest & "        </SegSelection>"

'If Stopover <> "" Then
'   strRequest = strRequest & "   <SegMods>"
'   strRequest = strRequest & "     <SegRangeAry>"
   
            'If SegSelected = True Then
'      For lngC1 = 0 To UBound(strSegs)
          
'          strRequest = strRequest & "        <SegRange>"
             'strRequest = strRequest & "          <BICOverride>" & strOverrideClass(lngC1) & "</BICOverride>"
'          strRequest = strRequest & "          <StartSeg>" & Format(strSegs(lngC1), 0) & "</StartSeg>"
'          strRequest = strRequest & "          <EndSeg>" & Format(strSegs(lngC1), 0) & "</EndSeg>"
'           strRequest = strRequest & "          <StopoverConx>" & strStopOver(lngC1) & "</StopoverConx>"
'          strRequest = strRequest & "        </SegRange>"
          
'      Next
'   strRequest = strRequest & "     </SegRangeAry>"
'   strRequest = strRequest & "   </SegMods>"
            'Else
'End If
   
            'strRequest = strRequest & "   <SegMods>"
            'strRequest = strRequest & "     <SegRangeAry>"
            'strRequest = strRequest & "        <SegRange>"
            'strRequest = strRequest & "          <BICOverride>" & OverrideClass & "</BICOverride>"
            'strRequest = strRequest & "          <StartSeg>1</StartSeg>"
            'strRequest = strRequest & "          <EndSeg>1</EndSeg>"
            'strRequest = strRequest & "        </SegRange>"
            'strRequest = strRequest & "        <SegRange>"
            'strRequest = strRequest & "          <BICOverride>" & OverrideClass & "</BICOverride>"
            'strRequest = strRequest & "          <StartSeg>2</StartSeg>"
            'strRequest = strRequest & "          <EndSeg>2</EndSeg>"
            'strRequest = strRequest & "        </SegRange>"
            'strRequest = strRequest & "     </SegRangeAry>"
            'strRequest = strRequest & "   </SegMods>"
            
               
            'If OverrideClass <> "" Then
            '
            'strRequest = strRequest & "   <SegMods>"
            'strRequest = strRequest & "     <SegRangeAry>"
            'strRequest = strRequest & "        <SegRange>"
            'strRequest = strRequest & "          <BICOverride>" & OverrideClass & "</BICOverride>"
            'strRequest = strRequest & "          <StartSeg>" & OverrideStartSeg & "</StartSeg>"
            'strRequest = strRequest & "          <EndSeg>" & OverrideEndSeg & "</EndSeg>"
            'strRequest = strRequest & "        </SegRange>"
            'strRequest = strRequest & "     </SegRangeAry>"
            'strRequest = strRequest & "   </SegMods>"
            'End If

'If ExcTaxCtry <> "" Then
'    strRequest = strRequest & "   <TaxExemptMods>"
'    strRequest = strRequest & "   <ExemptCountryAry>"
'    strRequest = strRequest & "   <ExemptCountry>"
'    strRequest = strRequest & "   <ExemptAllTaxes> N </ExemptAllTaxes>"
'    strRequest = strRequest & "   <Country>" & ExcTaxCtry & "</Country>"
'    strRequest = strRequest & "  </ExemptCountry>"
'    strRequest = strRequest & "  </ExemptCountryAry>"
'    strRequest = strRequest & " </TaxExemptMods>"
'End If

                'If UCase(gstrAgcyCountryCode) = "SG" Then
                '      strRequest = strRequest & "  <GenQuoteInfo>"
                '      strRequest = strRequest & "       <SellCity/>"
                '      strRequest = strRequest & "       <TkCity/>"
                '      strRequest = strRequest & "       <AltCurrency>" & gstrAgcyCurrencyCode & "</AltCurrency>"
                '      strRequest = strRequest & "       <EquivCurrency/>"
                '      strRequest = strRequest & "       <TkDt/>"
                '      strRequest = strRequest & "       <BkDtOverride/>"
                '      strRequest = strRequest & "       <EUROverride/>"
                '      strRequest = strRequest & "       <LCUOverride/>"
                '      strRequest = strRequest & "  </GenQuoteInfo>"
                'End If
'strRequest = strRequest & "    </FairQuoteMods>"
'strRequest = strRequest & "</FareQuoteStandard_8_0>"

'Set xmldomFQ = New MSXML2.DOMDocument
                'strResponse = gobjHost.SendQuery(strRequest, "FareQuoteStandard_8_0", "FareQuotes", "GetFareQuote")
'AryStrResponse(lngPF) = gobjHost.SendQuery(strRequest, "FareQuoteStandard_8_0", "FareQuotes", "GetFareQuote")

'Next lngPF

'added 21 Dec for same passenger type use back same fare
Dim distinctAryStrPx() As String
Dim curPsger As String
Dim prevPsger As String
Dim j As Integer
Dim k As Integer
Dim q As Integer
Dim intKey() As Integer
Dim intPNum As Integer
Dim blnNoMatch As Boolean
Dim blnMatch As Boolean

'Preethi- V1.2.1 201001013 - IR 6 - Pick Up Rule ID in Fare Rules
Dim intI As Integer
'JiYong - V1.2.1 (Urgent Fix) 201001102 - IR 6 - Pick Up Rule ID in Fare Rules
'Change RuleID to string
'Dim intRuleID As Integer
Dim strRuleID As String

ReDim intKey(UBound(AryStrPx))

For i = 0 To UBound(AryStrPx)

    curPsger = Mid(AryStrPx(i), (InStr(1, AryStrPx(i), "*") + 1), 2)
    If i = 0 Then
            j = j + 1
            ReDim Preserve distinctAryStrPx(j - 1)
            distinctAryStrPx(j - 1) = curPsger
    Else
    
    For k = 0 To UBound(distinctAryStrPx)
        blnNoMatch = False
        blnMatch = False
        For q = 0 To UBound(distinctAryStrPx)
            If curPsger <> distinctAryStrPx(q) Then
              blnNoMatch = True
            Else
              blnMatch = True
            End If
        Next q
        
        If blnNoMatch = True And blnMatch = False Then
               j = j + 1
               ReDim Preserve distinctAryStrPx(j - 1)
               distinctAryStrPx(j - 1) = curPsger
               
        End If
        
    Next k
    
    End If
Next i

For i = 0 To UBound(AryStrPx)
    For j = 0 To UBound(distinctAryStrPx)
    
    If Mid(AryStrPx(i), (InStr(1, AryStrPx(i), "*") + 1), 2) = distinctAryStrPx(j) Then
        intKey(i) = j
        Exit For
    End If
    
    Next j
Next i




Set mcolPX = New Collection
For i = 0 To UBound(AryStrPx)
   
'Added 211204
   intPNum = i
   
   Set mcolFQ = New FQ
   For lngPF = 0 To PFRec
'xmldomFQ.loadXML AryStrResponse(lngPF)
'Modified on 170506: load fare quote and fare rules continuously to get the correct fare rules
 '------------------------------------------------------------------------------------------------
strRequest = "<FareQuoteStandard_11>"
strRequest = strRequest & "   <FareQuoteMods>"
strRequest = strRequest & "     <PsgrMods>"
strRequest = strRequest & "<PsgrAry>"

For lngC1 = 0 To UBound(AryStrPx)
    strRequest = strRequest & "<Psgr>"
    strRequest = strRequest & "<LNameNum>" & Left(AryStrPx(lngC1), InStr(1, AryStrPx(lngC1), "*") - 1) & "</LNameNum>"
    strRequest = strRequest & "<PsgrNum>1</PsgrNum>"
    strRequest = strRequest & "<AbsNameNum>" & Left(AryStrPx(lngC1), InStr(1, AryStrPx(lngC1), "*") - 1) & "</AbsNameNum>"
    strRequest = strRequest & "<PIC>" & IIf(Mid(AryStrPx(lngC1), InStr(1, AryStrPx(lngC1), "*") + 1) = "AD", "", Mid(AryStrPx(lngC1), InStr(1, AryStrPx(lngC1), "*") + 1)) & "</PIC>"
    strRequest = strRequest & "<TIC></TIC>"
    strRequest = strRequest & "</Psgr>"
Next


strRequest = strRequest & "</PsgrAry>"
strRequest = strRequest & "</PsgrMods>"



If Len(PlatCarrier) > 0 Then
    strRequest = strRequest & "<PlatingAirVMods>"
    strRequest = strRequest & "   <PlatingAirV>" & PlatCarrier & "</PlatingAirV>"
    strRequest = strRequest & "</PlatingAirVMods>"
End If
strRequest = strRequest & "   <GenQuoteInfo>"
strRequest = strRequest & "   <TkType>" & tktType & "</TkType>"
strRequest = strRequest & "   </GenQuoteInfo>"
strRequest = strRequest & "        <DocProdFareType>"
strRequest = strRequest & "            <Type>I</Type>"
strRequest = strRequest & "        </DocProdFareType>"
strRequest = strRequest & "        <SegSelection>"
strRequest = strRequest & "            <ReqAirVPFs>Y</ReqAirVPFs>"
strRequest = strRequest & "            <SegRangeAry>"
If OverrideClass = "" Then
If SegSelected = True Then
    
    For lngC1 = 0 To UBound(strSegs)
        strRequest = strRequest & "<SegRange>"
        strRequest = strRequest & "<StartSeg>" & Format(strSegs(lngC1), 0) & "</StartSeg>"
        strRequest = strRequest & "<EndSeg>" & Format(strSegs(lngC1), 0) & "</EndSeg>"
        strRequest = strRequest & "<FareType>P</FareType>"
        strRequest = strRequest & "<PFQual>"
        strRequest = strRequest & "<CRSInd>1G</CRSInd>"
        'added on 240806 allow FQ in other PCC
        If PCC = "" Then
            strRequest = strRequest & "<PCC>" & gstrHQPCC & "</PCC>"
        Else
            strRequest = strRequest & "<PCC>" & PCC & "</PCC>"
        End If
        strRequest = strRequest & "<AirV></AirV>"
        strRequest = strRequest & "<Acct>" & strPFCodes(lngPF) & "</Acct>"
        strRequest = strRequest & "<Contract></Contract>"
        strRequest = strRequest & "<PublishedFaresInd>Y</PublishedFaresInd>"
        strRequest = strRequest & "<Type>A</Type>"
        If strPFCodes(lngPF) <> "" Then
           strRequest = strRequest & "<AcctCodeRestrict>Y</AcctCodeRestrict>"
        End If
        strRequest = strRequest & "</PFQual>"
        strRequest = strRequest & "</SegRange>"
    Next lngC1
Else
    strRequest = strRequest & "                <SegRange>"
    strRequest = strRequest & "                    <StartSeg>00</StartSeg>"
    strRequest = strRequest & "                    <EndSeg>00</EndSeg>"
    strRequest = strRequest & "                    <FareType>P</FareType>"
    strRequest = strRequest & "                    <PFQual>"
    strRequest = strRequest & "                        <CRSInd>1G</CRSInd>"
    'strRequest = strRequest & "                        <PCC>" & gstrHQPCC & "</PCC>"
    If PCC = "" Then
            strRequest = strRequest & "<PCC>" & gstrHQPCC & "</PCC>"
        Else
            strRequest = strRequest & "<PCC>" & PCC & "</PCC>"
        End If
    strRequest = strRequest & "                        <AirV></AirV>"
    strRequest = strRequest & "                        <Acct>" & strPFCodes(lngPF) & "</Acct>"
    strRequest = strRequest & "                        <PublishedFaresInd>Y</PublishedFaresInd>"
    strRequest = strRequest & "                        <Type>A</Type>"
    If strPFCodes(lngPF) <> "" Then
       strRequest = strRequest & "<AcctCodeRestrict>Y</AcctCodeRestrict>"
    End If
    strRequest = strRequest & "                    </PFQual>"
    strRequest = strRequest & "                </SegRange>"
End If
Else
    strSegs = Split(Segments, ".")
    For lngC1 = 0 To UBound(strSegs)
        strRequest = strRequest & "<SegRange>"
        strRequest = strRequest & "<StartSeg>" & Format(strSegs(lngC1), 0) & "</StartSeg>"
        strRequest = strRequest & "<EndSeg>" & Format(strSegs(lngC1), 0) & "</EndSeg>"
        strRequest = strRequest & "<FareType>B</FareType>"
        strRequest = strRequest & "<FIC>" & strOverrideClass(lngC1) & "</FIC>"
        strRequest = strRequest & "<PFQual>"
        strRequest = strRequest & "<CRSInd>1G</CRSInd>"
        If PCC = "" Then
            strRequest = strRequest & "<PCC>" & gstrHQPCC & "</PCC>"
        Else
            strRequest = strRequest & "<PCC>" & PCC & "</PCC>"
        End If
        strRequest = strRequest & "<AirV></AirV>"
        strRequest = strRequest & "<Acct>" & strPFCodes(lngPF) & "</Acct>"
        strRequest = strRequest & "<Contract></Contract>"
        strRequest = strRequest & "<PublishedFaresInd>Y</PublishedFaresInd>"
        strRequest = strRequest & "<Type>A</Type>"
        If strPFCodes(lngPF) <> "" Then
           strRequest = strRequest & "<AcctCodeRestrict>Y</AcctCodeRestrict>"
        End If
        strRequest = strRequest & "</PFQual>"
        strRequest = strRequest & "</SegRange>"
    Next lngC1
End If



strRequest = strRequest & "            </SegRangeAry>"
strRequest = strRequest & "        </SegSelection>"

If Stopover <> "" Then
   strRequest = strRequest & "   <SegMods>"
   strRequest = strRequest & "     <SegRangeAry>"
   
      For lngC1 = 0 To UBound(strSegs)
          
          strRequest = strRequest & "        <SegRange>"
          strRequest = strRequest & "          <StartSeg>" & Format(strSegs(lngC1), 0) & "</StartSeg>"
          strRequest = strRequest & "          <EndSeg>" & Format(strSegs(lngC1), 0) & "</EndSeg>"
           strRequest = strRequest & "          <StopoverConx>" & strStopOver(lngC1) & "</StopoverConx>"
          strRequest = strRequest & "        </SegRange>"
          
      Next
   strRequest = strRequest & "     </SegRangeAry>"
   strRequest = strRequest & "   </SegMods>"
  
End If
   

If ExcTaxCtry <> "" Then
    strRequest = strRequest & "   <TaxExemptMods>"
    strRequest = strRequest & "   <ExemptCountryAry>"
    strRequest = strRequest & "   <ExemptCountry>"
    strRequest = strRequest & "   <ExemptAllTaxes> N </ExemptAllTaxes>"
    strRequest = strRequest & "   <Country>" & ExcTaxCtry & "</Country>"
    strRequest = strRequest & "  </ExemptCountry>"
    strRequest = strRequest & "  </ExemptCountryAry>"
    strRequest = strRequest & " </TaxExemptMods>"
End If
                                         
strRequest = strRequest & "    </FareQuoteMods>"
strRequest = strRequest & "</FareQuoteStandard_11>"

Set xmldomFQ = New MSXML2.DOMDocument

'AryStrResponse(lngPF) = gobjHost.SendQuery(strRequest, "FareQuoteStandard_8_0", "FareQuotes", "GetFareQuote")
'xmldomFQ.loadXML AryStrResponse(lngPF)
      
 xmldomFQ.loadXML gobjHost.SendQuery(strRequest, "FareQuoteStandard_11", "FareQuotes", "GetFareQuote")
   
'------------------------------------------------------------------------------------------------------
      
      Set xmlnlTemp5 = xmldomFQ.selectNodes("//FareInfo/InfoMsg")

      Set objFQ = New FareQuote
      
      i = intPNum
      
      objFQ.PIC = IIf(Mid(AryStrPx(i), InStr(1, AryStrPx(i), "*") + 1) = "AD", "", Mid(AryStrPx(i), InStr(1, AryStrPx(i), "*") + 1))
      objFQ.PxNum = Mid(AryStrPx(i), 1, InStr(1, AryStrPx(i), "*") - 1)
      objFQ.SegmentSelected = SegSelected
      objFQ.SegmentSelectString = Segments
      objFQ.FQPCC = PCC
      If Stopover <> "" Then objFQ.OverrideConx = True
      
      'added 21 Dec
      i = intKey(i)
      strErrMsg = ""
      
If InStr(xmldomFQ.xml, "TransactionErrorCode") <> 0 Then
    strErrMsg = "Transaction Error " & xmldomFQ.getElementsByTagName("ErrorCode").Item(0).Text & vbCrLf _
        & xmldomFQ.documentElement.nodeName & vbCrLf _
        & xmldomFQ.getElementsByTagName("Text").Item(0).Text

    
    
    objFQ.UnableToQuote = True

    mcolFQ.FQAdd objFQ
    
    'modMsgBox.OKMsg = "OK"
    'modMsgBox.sMsgBox gVPMDIHwnd, strErrMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    GoTo skip_the_rest
    'GoTo NextPx
'Preethi - V1.2.2 20110302 - IR9 - Error is prompted at Fare Quote Request Screen but still be able to proceed to Fare Quote Screen

ElseIf InStr(xmldomFQ.xml, "UNABLE TO FARE QUOTE - FAILED RULES VALIDATION") Or _
        InStr(xmldomFQ.xml, "NO VALID FARE FOR INPUT CRITERIA") Then
    failNum = failNum + 1
    If failNum > PFRec Then

        objFQ.UnableToQuote = True
        objFQ.NoFare = True

        mcolFQ.FQAdd objFQ
     
        
        If PFRec > 0 Then
            strErrMsg = "UNABLE TO QUOTE ON PRIVATE FARES"
      
            'modMsgBox.OKMsg = "OK"
            'modMsgBox.sMsgBox gVPMDIHwnd, strErrMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        Else
            strErrMsg = "UNABLE TO FARE QUOTE - FAILED RULES VALIDATION"
 
            'modMsgBox.OKMsg = "OK"
            'modMsgBox.sMsgBox gVPMDIHwnd, strErrMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        End If
        GoTo skip_the_rest
        'GoTo NextPx
    Else
        GoTo skip_the_rest
    End If

ElseIf InStr(xmldomFQ.xml, "<ErrMsg>Y</ErrMsg>") <> 0 Then
    strErrMsg = "GDS Error " & vbCrLf _
        & xmldomFQ.documentElement.nodeName & vbCrLf _
        & xmldomFQ.getElementsByTagName("Text").Item(0).Text

    objFQ.UnableToQuote = True

    mcolFQ.FQAdd objFQ


    'modMsgBox.OKMsg = "OK"
    'modMsgBox.sMsgBox gVPMDIHwnd, strErrMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    GoTo skip_the_rest
    'GoTo NextPx

End If



If xmldomFQ.selectNodes("//GenQuoteDetails").length = 0 Then
   If xmldomFQ.selectNodes("//OutputMsg//Text").length <> 0 Then
      Set xmlnlTemp1 = xmldomFQ.selectNodes("//OutputMsg//Text")
      objFQ.UnableToQuote = True

      mcolFQ.FQAdd objFQ
      strErrMsg = xmlnlTemp1.Item(0).Text
      'modMsgBox.OKMsg = "OK"
      'modMsgBox.sMsgBox gVPMDIHwnd, xmlnlTemp1.Item(0).Text, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      'GoTo NextPx
      'Exit Function
      GoTo skip_the_rest
   Else
      objFQ.UnableToQuote = True

      mcolFQ.FQAdd objFQ
      strErrMsg = "UNABLE TO FARE QUOTE"
      'modMsgBox.OKMsg = "OK"
      'modMsgBox.sMsgBox gVPMDIHwnd, "UNABLE TO FARE QUOTE", vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      'GoTo NextPx
      GoTo skip_the_rest
      'Exit Function
   End If
   
   Exit Function
End If

'modified on 09/11/04: to capture Endorsement #1
'Set xmlnlTemp5 = New MSXML2.IXMLDOMNodeList
Set xmlnlTemp5 = xmldomFQ.selectNodes("//FareInfo/InfoMsg")
str1stEndorseText = ""
If xmlnlTemp5.length > 0 Then
    For lngC1 = 0 To xmlnlTemp5.length - 1
        
        If xmlnlTemp5(lngC1).childNodes.length > 0 Then
            '29122004
            'If xmlnlTemp5(lngC1).selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") And xmlnlTemp5(lngC1).selectNodes("//MsgType")(lngC1).Text = "1" Then
            If xmlnlTemp5(lngC1).selectSingleNode("QuoteNum").Text = intPNum + 1 And xmlnlTemp5(lngC1).selectNodes("//MsgType")(lngC1).Text = "1" Then
                str1stEndorseText = str1stEndorseText & xmlnlTemp5(lngC1).selectNodes("//Text")(lngC1).Text & "^"
            End If
        End If
    Next lngC1
    If Len(str1stEndorseText) > 0 Then
        str1stEndorseText = Mid(str1stEndorseText, 1, Len(str1stEndorseText) - 1)
    End If
End If

objFQ.UnableToQuote = False
objFQ.PlatCarrier = PlatCarrier
'Set xmlnlTemp1 = xmldomFQ.selectSingleNode("//GenQuoteDetails").childNodes
'reset all variables
lngBaseAmt = 0
intBaseDecPos = 0
lngEquivAmt = 0
intEquivDecPos = 0
lngTotAmt = 0
intTotDecPos = 0

Set xmlnlTemp1 = xmldomFQ.selectNodes("//GenQuoteDetails")
For lngC1 = 0 To xmlnlTemp1.length - 1
 
  If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
  Set xmlnlTemp2 = xmlnlTemp1(lngC1).selectNodes("TaxDataAry//TaxData")
    For lngC2 = 0 To xmlnlTemp2.length - 1
        Set objTax = New Tax
        Set xmlnlTemp3 = xmlnlTemp2.Item(lngC2).childNodes
        For lngC3 = 0 To xmlnlTemp3.length - 1
                With xmlnlTemp3.Item(lngC3)
                    Select Case .nodeName
                        Case "Country"
                            objTax.TaxCode = .Text
                        Case "Amt"
                            If IsNumeric(.Text) = True Then
                                objTax.Amount = CSng(.Text)
                            Else
                                objTax.Amount = 0
                            End If
                            'objTax.Amount = CSng(.Text)'decimal placement is same as total fare
                    End Select
                End With
        Next lngC3
        objFQ.TaxAdd objTax
        Set objTax = Nothing
    Next lngC2
    
    
   For lngC2 = 0 To xmlnlTemp1(lngC1).childNodes.length - 1
    With xmlnlTemp1(lngC1).childNodes.Item(lngC2)
        Select Case .nodeName
            Case "QuoteType"
                objFQ.QuoteType = .Text
            Case "LastTkDt"
                If .Text <> "" Then objFQ.LastTicketDate = ConvertDate(.Text)
            Case "IntlSaleInd"
                objFQ.ISO = .Text
            Case "BaseFareCurrency"
                objFQ.BaseCurrency = .Text
            Case "BaseFareAmt"
                lngBaseAmt = CLng(.Text)
            Case "BaseDecPos"
                intBaseDecPos = CInt(.Text)
            Case "EquivCurrency"
                objFQ.EquivCurrency = .Text
            Case "EquivAmt"
                lngEquivAmt = CLng(.Text)
            Case "EquivDecPos"
                intEquivDecPos = CInt(.Text)
            Case "TotCurrency"
                objFQ.TotalCurrency = .Text
            Case "TotAmt"
                lngTotAmt = CLng(.Text)
            Case "TotDecPos"
                intTotDecPos = CInt(.Text)
            Case "ITNum"
                objFQ.ITNum = .Text
            Case "PrivFQd"
                If .Text = "Y" Then
                    objFQ.PrivateFare = True
                    bolPF = True
                    objFQ.PFAccountCode = strPFCodes(lngPF)
                    objFQ.PFFareType = strPFType(lngPF)
                Else
                    objFQ.PrivateFare = False
                    bolPF = False
                End If
           
            Case Else
            
                'nothing
        End Select
    End With
   Next lngC2
  End If
 
Next lngC1
objFQ.BaseAmount = PlaceDecimal(intBaseDecPos, lngBaseAmt)
objFQ.EquivAmount = PlaceDecimal(intEquivDecPos, lngEquivAmt)
objFQ.TotAmount = PlaceDecimal(intTotDecPos, lngTotAmt)
If tktType = "P" Then objFQ.PTkt = True
        
'Set xmlnlTemp1 = xmldomFQ.selectNodes("//TaxDataAry//TaxData")
'    For lngC1 = 0 To xmlnlTemp1.length - 1
'        Set objTax = New Tax
'        Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
'        For lngC2 = 0 To xmlnlTemp2.length - 1
'                With xmlnlTemp2.Item(lngC2)
'                    Select Case .nodeName
'                        Case "Country"
'                            objTax.TaxCode = .Text
'                        Case "Amt"
'                            objTax.Amount = CSng(.Text) 'decimal placement is same as total fare
'                    End Select
'                End With
'        Next lngC2
'        objFQ.TaxAdd objTax
'        Set objTax = Nothing
'    Next lngC1

'Modified on 14/1/2005: skip surface segment which is added by galileo
intClassCount = 0
Set xmlnlTemp1 = xmldomFQ.selectNodes("//ItinSeg")

For lngC1 = 0 To xmlnlTemp1.length - 1

  '''UniqueKey is not Px number
  'If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
    Set objFQSeg = New FareSegment
    Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
    For lngC2 = 0 To xmlnlTemp2.length - 1
        With xmlnlTemp2.Item(lngC2)
            Select Case .nodeName
                Case "SegType"
                    objFQSeg.SegType = .Text
                Case "AirV"
                    objFQSeg.Vendor = .Text
                Case "FltNum"
                    objFQSeg.FlightNum = .Text
                Case "StartPt"
                    objFQSeg.DepCityCode = .Text
                Case "EndPt"
                    objFQSeg.ArrCityCode = .Text
                Case "BIC"
                    objFQSeg.COS = .Text
                Case "FQSGenFictitiousPt"
                    strFicPt = .Text

            End Select
        End With
    Next lngC2
    
    'Added on 6/10/2005: Add OverridePFBC
    
        
    If Trim(UCase(strFicPt)) <> "Y" And objFQSeg.SegType <> "3" Then
        If Trim(OverrideClass) <> "" Then objFQSeg.OverridePFBC = strOverrideClass(intClassCount)
        objFQ.FareSegAdd objFQSeg
        intClassCount = intClassCount + 1
        Set objFQSeg = Nothing
    ElseIf objFQSeg.SegType = "3" Then
        For q = objFQ.FareSegCount To 1 Step -1
            If objFQ.FareSeg(q).SegType = "0" Then
                objFQ.FareSeg(q).ArrCityCode = objFQSeg.ArrCityCode
                Exit For
            End If
        Next
        Set objFQSeg = Nothing

    
    End If
    
Next lngC1

'Set xmlnlTemp1 = xmldomFQ.selectNodes("//SegRelatedInfo")
'For lngC1 = 0 To xmlnlTemp1.length - 1
'  If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
  
Set xmlnlTemp1 = xmldomFQ.selectNodes("//SegRelatedInfo")
For lngC1 = 0 To xmlnlTemp1.length - 1
  If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
    Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
    For lngC2 = 0 To xmlnlTemp2.length - 1
    With xmlnlTemp2.Item(lngC2)
            Select Case .nodeName
                Case "RelSegNum"
                    intTemp = .Text   '''Air SegNum or Iti SegNum???
                Case "NotValidBeforeDt"
                    If .Text <> "" Then objFQ.FareSeg(intTemp).NVB = ConvertDate(.Text)
                Case "NotValidAfterDt"
                    If .Text <> "" Then objFQ.FareSeg(intTemp).NVA = ConvertDate(.Text)
                Case "Stopover"
                    objFQ.FareSeg(intTemp).Stopover = IIf(.Text = "X", False, True)
                Case "FIC"
                    objFQ.FareSeg(intTemp).FBC = .Text
                Case "TkDesignator"
                    objFQ.FareSeg(intTemp).TD = .Text
                Case "BagInfo"
                    objFQ.FareSeg(intTemp).BagInfo = .Text
            End Select
        End With
    Next lngC2
  End If
Next lngC1

Set xmlnlTemp3 = Nothing
Set xmlnlTemp4 = Nothing

Set xmlnlTemp1 = xmldomFQ.selectNodes("//InfoMsg//Text")

With xmlnlTemp1
For lngC1 = 0 To .length - 1
   If xmlnlTemp1(lngC1).parentNode.selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
      objFQ.InfoMsgAdd (.Item(lngC1).Text)
   End If
Next
End With

Set xmlnlTemp1 = xmldomFQ.selectNodes("//FareInfo/FareConstruction")
For lngC1 = 0 To xmlnlTemp1.length - 1
   If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
      objFQ.FareConstructText = xmlnlTemp1(lngC1).selectSingleNode("FareConstructText").Text
   End If
Next lngC1

'objFQ.FareConstructText = xmldomFQ.selectSingleNode("//FareConstruction/FareConstructText").Text



Set xmlnlTemp1 = xmldomFQ.selectNodes("//JrnyType")
For lngC1 = 0 To xmlnlTemp1.length - 1
   If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = Format(i + 1, "0000") Then
      Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
       For lngC2 = 0 To xmlnlTemp2.length - 1
       With xmlnlTemp2.Item(lngC2)
   
        'With xmlnlTemp1.Item(lngC1)
            Select Case .nodeName
                Case "CircTrip"
                    If .Text = "Y" Then objFQ.JourneyType = "CT"
                Case "RT"
                    If .Text = "Y" Then objFQ.JourneyType = "RT"
                Case "OW"
                    If .Text = "Y" Then objFQ.JourneyType = "OW"
                Case "OpenJaw"
                    If .Text = "Y" Then objFQ.JourneyType = "OJ"
                Case "RndWorld"
                    If .Text = "Y" Then objFQ.JourneyType = "RW"
            End Select
        End With
        Next lngC2
    End If
Next lngC1

'Added by Ji Yong on 06 Dec to determine is this a Cat35 fare
lngBaseAmt = 0
intBaseDecPos = 0
lngEquivAmt = 0
intEquivDecPos = 0
lngTotAmt = 0
intTotDecPos = 0

'Added by JiYong to get the actual net fare
lngNetFareAmt = 0
intNetFarePos = 0
strNetFareCrcy = ""
lngEquivNetFareAmt = 0
intEquivNetFareDecPos = 0
strEquivNetFareCrcy = ""
strTotalNetFareCrcy = ""

Set xmlnlTemp1 = xmldomFQ.selectNodes("//ExtendedQuoteInformation")
For lngC1 = 0 To xmlnlTemp1.length - 1
    If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
       Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
       For lngC2 = 0 To xmlnlTemp2.length - 1
           With xmlnlTemp2.Item(lngC2)
                Select Case .nodeName
                       Case "DataRetQuote"
                             objFQ.Cat35 = IIf(.Text = "Y", True, False)
                       'JiYong - V1.2.15 20121010 - IR19 - Get the correct comm indicator
                       'Case "NTypeComm"
                       '      objFQ.CommType = IIf(.Text = "Y", True, False)
                       Case "TkFareCrcy"
                             objFQ.TktBaseCurrency = .Text
                       Case "TkFareAmt"
                             lngBaseAmt = CLng(.Text)
                       Case "TkNumDecs"
                             intBaseDecPos = CInt(.Text)
                       Case "EquivTkCrncy"
                             objFQ.TktEquivCurrency = .Text
                       Case "EquivTkAmt"
                             lngEquivAmt = CLng(.Text)
                       Case "EquivTkNumDecs"
                             intEquivDecPos = CInt(.Text)
                       Case "TotTkCrncy"
                             objFQ.TktTotalCurrency = .Text
                       Case "TotTkAmt"
                             lngTotAmt = CLng(.Text)
                       Case "TotTkNumDecs"
                             intTotDecPos = CInt(.Text)
                       Case "NetFareCrncy"
                            strNetFareCrcy = .Text
                       Case "NetFareAmt"
                            lngNetFareAmt = CLng(.Text)
                       Case "NetFareNumDecs"
                            intNetFarePos = CInt(.Text)
                       Case "EquivNetFareCrcy"
                            strEquivNetFareCrcy = .Text
                       Case "EquivNetFareAmt"
                            lngEquivNetFareAmt = CLng(.Text)
                       Case "EquivNumDecs"
                            intEquivNetFareDecPos = CInt(.Text)
                       Case "TotNetCrcy"
                            strTotalNetFareCrcy = .Text
                End Select
           End With
        Next lngC2
    End If
Next lngC1
objFQ.TktBaseAmount = PlaceDecimal(intBaseDecPos, lngBaseAmt)
objFQ.TktEquivAmount = PlaceDecimal(intEquivDecPos, lngEquivAmt)
objFQ.TktTotAmount = PlaceDecimal(intTotDecPos, lngTotAmt)

'Added by JiYong to get the actual Net Fare
If lngNetFareAmt > 0 Then
'Preethi - V1.2.1 20101011 - CR21 - Nett Fare Mark Up
   'If strNetFareCrcy = strTotalNetFareCrcy Then objFQ.ActualNetAmount = PlaceDecimal(intNetFarePos, lngNetFareAmt)
   'Else: objFQ.ActualNetAmount = PlaceDecimal(intEquivNetFareDecPos, lngEquivNetFareAmt)
   If strNetFareCrcy = strTotalNetFareCrcy Then
     objFQ.ActualNetAmount = PlaceDecimal(intNetFarePos, lngNetFareAmt)
   Else
     objFQ.ActualNetAmount = PlaceDecimal(intEquivNetFareDecPos, lngEquivNetFareAmt)
   End If
   
End If

'JiYong - V1.2.15 20121010 - IR19 - Get the correct comm indicator
'Get from the 1st CommNetGrossInfo section
Set xmlnlTemp1 = xmldomFQ.selectNodes("//CommNetGrossInfo")
For lngC1 = 0 To xmlnlTemp1.length - 1
    If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
       Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
       For lngC2 = 0 To xmlnlTemp2.length - 1
           With xmlnlTemp2.Item(lngC2)
                Select Case .nodeName
                       Case "CommPercentage"
                            objFQ.CommType = IIf(CLng(.Text) <> 9999999, True, False)
                End Select
           End With
        Next lngC2
        'Get from the 1st CommNetGrossInfo section
        Exit For
    End If
Next lngC1


Set xmlnlTemp1 = xmldomFQ.selectNodes("//TkPsgrFareConstructTaxInfo")
For lngC1 = 0 To xmlnlTemp1.length - 1
  If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
          Set xmlnlTemp2 = xmlnlTemp1(lngC1).selectNodes("TaxDataAry//TaxData")
          For lngC2 = 0 To xmlnlTemp2.length - 1
              Set objTax = New Tax
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
              objFQ.TktTaxAdd objTax
              Set objTax = Nothing
          Next lngC2
  End If
Next lngC1


'If xmldomFQ.getElementsByTagName("JrnyType").length <> 0 Then
'    Set xmlnlTemp1 = xmldomFQ.getElementsByTagName("JrnyType").Item(0).childNodes
'    For lngC1 = 0 To xmlnlTemp1.length - 1
'        With xmlnlTemp1.Item(lngC1)
'            Select Case .nodeName
'                Case "CircTrip"
'                    If .Text = "Y" Then objFQ.JourneyType = "CT"
'                Case "RT"
'                    If .Text = "Y" Then objFQ.JourneyType = "RT"
'                Case "OW"
'                    If .Text = "Y" Then objFQ.JourneyType = "OW"
'                Case "OpenJaw"
'                    If .Text = "Y" Then objFQ.JourneyType = "OJ"
'                Case "RndWorld"
'                    If .Text = "Y" Then objFQ.JourneyType = "RW"
'            End Select
'        End With
'    Next lngC1
'End If

'Set dbRules = DAO.OpenDatabase(gstrFaresDBSource)
'Set rsRules = dbRules.OpenRecordset("SELECT * FROM tblRules")

'Set dbCity = DAO.OpenDatabase(gstrTProLookupDBSource)
'Set rsCity = dbCity.OpenRecordset("SELECT * FROM tblCityCodes")

'strSQL = "SELECT * FROM tblRules"
'rsRules.open strSQL, gdbConn, 3, 4

'Set xmlnlTemp0 = xmldomFQ.selectNodes("//RulesInfo")
'For lngC0 = 0 To xmlnlTemp0.length - 1
'  If xmlnlTemp0(lngC0).selectSingleNode("UniqueKey").Text = i + 1 Then
     
  
intTemp = -1
Set xmlnlTemp1 = xmldomFQ.selectNodes("//RulesInfo")
For lngC1 = 0 To xmlnlTemp1.length - 1
  If xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text = i + 1 Then
    intTemp = intTemp + 1
    Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
    Set objFC = New FareComponent
    For lngC2 = 0 To xmlnlTemp2.length - 1
        With xmlnlTemp2.Item(lngC2)
            Select Case .nodeName
                Case "StartPt"
                    objFC.Origin = .Text
                    'strSQL = "AirportCode = '" & .Text & "'"
                    STRSQL = "Select * from tblCityCodes where AirportCode = '" & .Text & "'"
                    Set rsCity = gdbConn.Execute(STRSQL)
                    With rsCity
                        '.FindFirst strSQL
                        If .EOF = False Then
                            objFC.OriginCountry = IIf(IsNull(![CountryCode]), "", ![CountryCode])
                            objFC.OriginRegion = IIf(IsNull(![Region]), "", ![Region])
                        End If
                    End With
               
                Case "EndPt"
                    objFC.Destinantion = .Text
                    'strSQL = "AirportCode = '" & .Text & "'"
                    STRSQL = "Select * from tblCityCodes where AirportCode = '" & .Text & "'"
                    Set rsCity = gdbConn.Execute(STRSQL)
                    With rsCity
                        '.FindFirst strSQL
                        If .EOF = False Then
                            objFC.DestinationCountry = IIf(IsNull(![CountryCode]), "", ![CountryCode])
                            objFC.DestinationRegion = IIf(IsNull(![Region]), "", ![Region])
                        End If
                    End With
                
                Case "AirV"
                    objFC.Vendor = .Text
                Case "FIC"
                    objFC.FBC = .Text
                Case "Currency"
                    objFC.CurrencyCode = .Text
                Case "DecPos"
                    intTotDecPos = CInt(.Text)
                Case "FareAmt"
                    lngTotAmt = CLng(.Text)
                Case "PFQuoted"
                    objFC.Net = True
                Case "FareRuleInfoYQual"
                    Set xmlnlTemp3 = .childNodes
                        For lngC3 = 0 To xmlnlTemp3.length - 1
                            With xmlnlTemp3.Item(lngC3)
                                Select Case .nodeName
                                    Case "GlobDirOrdinal"
                                        objFC.DirectionalInd = .Text
                                    Case "FlownMileComponent"
                                        objFC.TPM = CLng(.Text)
                                    Case "MPMComponent"
                                        objFC.MPM = CLng(.Text)
                                    Case "OTWTransportingAirV"
                                        objFC.OTWCarrier = .Text
                                End Select
                            End With
                        Next lngC3
                Case "FareRuleInfoNQual"
                    Set xmlnlTemp3 = .childNodes
                        For lngC3 = 0 To xmlnlTemp3.length - 1
                            With xmlnlTemp3.Item(lngC3)
                                Select Case .nodeName
                                    Case "RuleNum"
                                        objFC.RuleNum = .Text
                                    
                                End Select
                            End With
                        Next lngC3
            End Select
        End With
    Next lngC2
    objFC.Amount = PlaceDecimal(intTotDecPos, lngTotAmt)
    
    rsCity.Close
    Set rsCity = Nothing
    
    strErrMsg2 = gobjHost.TerminalEntry("FQN1") 'setting up host for rule parsing by fare component
    If strErrMsg2 = CONNECTION_FAIL Then
      objFQ.UnableToQuote = True
'''      mcolFQ.Add objFQ, strPFCodes(lngPF)
      mcolFQ.FQAdd objFQ
      'MsgBox strErrMsg, vbCritical
      modMsgBox.OKMsg = "OK"
      modMsgBox.sMsgBox gVPMDIHwnd, strErrMsg2, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
      Exit Function
    End If
        ' Need to determine if need to look up and/or parse ticketing rules data
    'lngRuleFC = lngC1 + 1 'setting a default value
    lngRuleFC = intTemp + 1 'setting a default value
    
'    If lngC1 = 0 Then
'        lngRuleFC = 1
'    Else
'        For lngC2 = 1 To lngC1
'            If (objFC.RuleNum = objFQ.FareComponent(lngC2).RuleNum) And (objFC.Vendor = objFQ.FareComponent(lngC2).Vendor) Then
'                lngRuleFC = lngC2
'                Exit For
'            End If
'        Next lngC2
'    End If
    If intTemp = 0 Then
        lngRuleFC = 1
    Else
        For lngC2 = 1 To intTemp
            If (objFC.RuleNum = objFQ.FareComponent(lngC2).RuleNum) And (objFC.Vendor = objFQ.FareComponent(lngC2).Vendor) Then
                lngRuleFC = lngC2
                Exit For
            End If
        Next lngC2
    End If
    
    'If lngRuleFC = lngC1 + 1 Then
    If lngRuleFC = intTemp + 1 Then
    
    'strsql = " Select * from tblRules where Carrier = '" & objFC.Vendor & "' AND RuleNum = '" & objFC.RuleNum & "'"
    'rsRules.open strsql, gdbConn, 3, 4
'Remove on 17082007: Not required to check rules for DB anymore
'    Set rsRules = gdbConn.Execute("Select * from tblRules where Carrier = '" & objFC.Vendor & "' AND RuleNum = '" & objFC.RuleNum & "'")
    
'    With rsRules
    
        '.FindFirst "Carrier = '" & objFC.Vendor & "' AND RuleNum = '" & objFC.RuleNum & "'"
'            If .EOF = False Then
'            objFC.TktInfoApplies = True
'            objFC.PriceFBC = ![PriceFBC] & ""
'            objFC.TicketFBC = ![TicketFBC] & ""
'            objFC.TktDesignator = ![TktDesig] & ""
'            objFC.ValueCode = ![ValueCode] & ""
'            objFC.CommissionOnTicket = ![TktComm]
'            objFC.NRCC = ![NRCC]
'            objFC.FOPCode = ![FOPCode] & ""
'            objFC.HPFApplies = ![HPF]
            ' Higher Point Fare (HPF) only applies to OJ and RW Journey Types
'            If ![HPF] = True And (objFQ.JourneyType = "OJ" Or objFQ.JourneyType = "RW") Then objFQ.HPFApplies = True
'            objFC.WLAllow = ![AllowWL]
'            objFC.OpenSegAllow = ![AllowOpen]
'            objFC.ETRequired = ![ETRequired]
'            objFC.PaperTktSurcharge = ![PTSurcharge]
'            objFC.ManualProcRequired = ![ManualProc]
'            objFC.FareOnTicket = ![FareOnTkt] & ""
            ' Pick up IT Number (Tour Code) if it was not in fare quote from Galileo
'            If objFQ.ITNum = "" Then objFQ.ITNum = ![TourCode] & ""
'            objFC.EndorsementAdd (![Endors1] & "")
'            If ![Endors2] & "" <> "" Then objFC.EndorsementAdd ![Endors2]
'            If ![Endors3] & "" <> "" Then objFC.EndorsementAdd ![Endors3]
'            If ![Endors4] & "" <> "" Then objFC.EndorsementAdd ![Endors4]
'            If ![Endors5] & "" <> "" Then objFC.EndorsementAdd ![Endors5]
'        Else
            objFC.TktInfoApplies = False
'        End If
'    End With
    
'    rsRules.Close
'    Set rsRules = Nothing
    'LOAD RULES TEXT AND PARSE TICKETING DATA FROM RULES TEXT
    
'    strRequest = ""
'strRequest = strRequest & "<FareQuoteRulesDisplay_8_0>"
'strRequest = strRequest & "    <FareDisplayMods>"
'strRequest = strRequest & "        <QueryHeader>"
'strRequest = strRequest & "            <UniqueKey>0000</UniqueKey>"
'strRequest = strRequest & "            <LangNum>0</LangNum>"
'strRequest = strRequest & "            <Action>31</Action>"
'strRequest = strRequest & "            <RetCRTOutput>N</RetCRTOutput>"
'strRequest = strRequest & "            <NoMsg>N</NoMsg>"
'strRequest = strRequest & "            <NoTrunc>Y</NoTrunc>"
'strRequest = strRequest & "            <IMInd>N</IMInd>"
'strRequest = strRequest & "            <FIPlus>N</FIPlus>"
'strRequest = strRequest & "            <PEInd>N</PEInd>"
'strRequest = strRequest & "            <HostUse16>N</HostUse16>"
'strRequest = strRequest & "            <NBInd>N</NBInd>"
'strRequest = strRequest & "            <ActionOnlyInd>N</ActionOnlyInd>"
'strRequest = strRequest & "            <TranslatePeriod>N</TranslatePeriod>"
'strRequest = strRequest & "            <GFYInd>N</GFYInd>"
'strRequest = strRequest & "            <IntFrame1>Y</IntFrame1>"
'strRequest = strRequest & "            <SmartParsed>Y</SmartParsed>"
'strRequest = strRequest & "            <PDCodes>N</PDCodes>"
'strRequest = strRequest & "            <BkDtOverride>N</BkDtOverride>"
'strRequest = strRequest & "            <HostUse25>N</HostUse25>"
'strRequest = strRequest & "            <DefCurrency />"
'strRequest = strRequest & "            <PFPWInd>N</PFPWInd>"
'strRequest = strRequest & "            <HostUse28>N</HostUse28>"
'strRequest = strRequest & "            <HostUse29>N</HostUse29>"
'strRequest = strRequest & "            <HostUse30>N</HostUse30>"
'strRequest = strRequest & "            <HostUse31>N</HostUse31>"
'strRequest = strRequest & "            <DefCurrencyLocInd />"
'strRequest = strRequest & "            <HostUse33 />"
'strRequest = strRequest & "            <Spare14>NNNNNNNNNNNNNNNN</Spare14>"
'strRequest = strRequest & "        </QueryHeader>"
'strRequest = strRequest & "        <FollowUpEntries>"
'strRequest = strRequest & "            <UniqueKey>0000</UniqueKey>"
''strRequest = strRequest & "            <QuoteNum>" & CStr(lngC1 + 1) & "</QuoteNum>"
'strRequest = strRequest & "            <QuoteNum>" & CStr(intTemp + 1) & "</QuoteNum>"
'strRequest = strRequest & "            <Spare1>N</Spare1>"
'strRequest = strRequest & "            <AllParaReqind>Y</AllParaReqind>"
'strRequest = strRequest & "            <SumRuleReqInd>N</SumRuleReqInd>"
'strRequest = strRequest & "            <FulltextoptInd>N</FulltextoptInd>"
'strRequest = strRequest & "            <Spare2>NNNN</Spare2>"
'strRequest = strRequest & "        </FollowUpEntries>"
'strRequest = strRequest & "    </FareDisplayMods>"
'strRequest = strRequest & "</FareQuoteRulesDisplay_8_0>"

'CC - V1.2.19 20130404 - CR211 - Change Fare Rules XML request from FareQuoteRulesDisplay_8_0 to FareQuoteMultiDisplay_8_3
strRequest = ""
strRequest = strRequest & "<FareQuoteMultiDisplay_8_3>"
strRequest = strRequest & "    <FareDisplayMods>"
strRequest = strRequest & "        <QueryHeader>"
strRequest = strRequest & "            <UniqueKey>0000</UniqueKey>"
strRequest = strRequest & "            <LangNum>0</LangNum>"
strRequest = strRequest & "            <Action>31</Action>"
strRequest = strRequest & "            <RetCRTOutput>N</RetCRTOutput>"
strRequest = strRequest & "            <NoMsg>N</NoMsg>"
strRequest = strRequest & "            <NoTrunc>Y</NoTrunc>"
strRequest = strRequest & "            <IMInd>N</IMInd>"
strRequest = strRequest & "            <FIPlus>N</FIPlus>"
strRequest = strRequest & "            <PEInd>N</PEInd>"
'strRequest = strRequest & "            <HostUse16>N</HostUse16>"
'PVInd      - New in FareQuoteMultiDisplay_8_3
strRequest = strRequest & "            <PVInd>N</PVInd>"

strRequest = strRequest & "            <NBInd>N</NBInd>"
strRequest = strRequest & "            <ActionOnlyInd>N</ActionOnlyInd>"
strRequest = strRequest & "            <TranslatePeriod>N</TranslatePeriod>"
'strRequest = strRequest & "            <GFYInd>N</GFYInd>"
'PIInd      - New in FareQuoteMultiDisplay_8_3
strRequest = strRequest & "            <PIInd>N</PIInd>"

strRequest = strRequest & "            <IntFrame1>Y</IntFrame1>"
strRequest = strRequest & "            <SmartParsed>Y</SmartParsed>"
strRequest = strRequest & "            <PDCodes>N</PDCodes>"
strRequest = strRequest & "            <BkDtOverride>N</BkDtOverride>"
strRequest = strRequest & "            <HostUse25>N</HostUse25>"
strRequest = strRequest & "            <DefCurrency />"
strRequest = strRequest & "            <PFPWInd>N</PFPWInd>"
'PFPQInd    - New in FareQuoteMultiDisplay_8_3
strRequest = strRequest & "            <PFPQInd>N</PFPQInd>"

'strRequest = strRequest & "            <HostUse28>N</HostUse28>"
strRequest = strRequest & "            <HostUse29>N</HostUse29>"
strRequest = strRequest & "            <HostUse30>N</HostUse30>"
strRequest = strRequest & "            <HostUse31>N</HostUse31>"
'strRequest = strRequest & "            <DefCurrencyLocInd />"
strRequest = strRequest & "            <HostUse32 />"
strRequest = strRequest & "            <HostUse33 />"
'strRequest = strRequest & "            <Spare14>NNNNNNNNNNNNNNNN</Spare14>"
strRequest = strRequest & "        </QueryHeader>"
strRequest = strRequest & "        <FollowUpEntries>"
strRequest = strRequest & "            <UniqueKey>0000</UniqueKey>"
strRequest = strRequest & "            <QuoteNum>" & CStr(intTemp + 1) & "</QuoteNum>"
strRequest = strRequest & "            <Spare1>N</Spare1>"
strRequest = strRequest & "            <AllParaReqind>Y</AllParaReqind>"
strRequest = strRequest & "            <SumRuleReqInd>N</SumRuleReqInd>"
strRequest = strRequest & "            <FulltextoptInd>N</FulltextoptInd>"
strRequest = strRequest & "            <Spare2>NNNN</Spare2>"
strRequest = strRequest & "        </FollowUpEntries>"
strRequest = strRequest & "    </FareDisplayMods>"
strRequest = strRequest & "</FareQuoteMultiDisplay_8_3>"


Set xmldomFCRules = New MSXML2.DOMDocument
strResponse = gobjHost.SendQuery(strRequest, "FareQuoteRulesDisplay_8_0", "FareQuotes", "GetFareQuote")
xmldomFCRules.loadXML strResponse
bolAfOthNote = False
'modified on 09/11/04: adding endorsement 1 from //FareInfo/InfoMsg
If Len(str1stEndorseText) > 0 Then
    objFC.EndorsementAdd str1stEndorseText, 1
End If

Set xmlnlTemp2 = xmldomFCRules.selectNodes("//FareInfo/RulesData/RulesText")

For lngC2 = 0 To xmlnlTemp2.length - 1
    objFC.RuleTextAdd xmlnlTemp2(lngC2).Text
    'Added on 031104: Endorsement#1 retrieved from Account Code Page
    'If InStr(xmlnlTemp2(lngC2).Text, ":") = 0 Then
    '    If InStr(UCase(xmlnlTemp2(lngC2).Text), "TICKET ENDORSEMENT") > 0 Then
    '        If InStr(UCase(Trim(xmlnlTemp5(lngC2).nextSibling.childNodes(2).Text)), "NO RESTRICTIONS") = 0 Then
    '            strOthNotesText = Trim(xmlnlTemp5(lngC2).nextSibling.childNodes(2).Text)
    '            If InStr(UCase(strOthNotesText), "NOTE") > 0 Then
    '                str1stEndorseText = Trim(xmlnlTemp5(lngC2).nextSibling.nextSibling.childNodes(2).Text)
    '            Else
    '                str1stEndorseText = strOthNotesText
    '            End If
    '         End If
    '    End If
    'Else
  'Debug.Print xmlnlTemp2(lngC2).Text
  'Added on 070207: Only take other notes parameter
 'Preethi- V1.2.1 201001013 - IR 6 - Pick Up Rule ID in Fare Rules
    'If InStr(xmlnlTemp2(lngC2).Text, "RULE IDENTIFIER:") > 0 Then
       'If objFQ.PFAccountCode <> "" Then
          'objFQ.RuleID = Trim(Replace(xmlnlTemp2(lngC2).Text, "RULE IDENTIFIER:", ""))
      ' End If
   ' End If
    'Extract the Rule ID from Fare Rules
    intI = 0
    'JiYong - V1.2.1 (Urgent Fix) 201001102 - IR 6 - Pick Up Rule ID in Fare Rules
    'Change RuleID to string
    'intRuleID = 0
    strRuleID = ""
    If InStr(xmlnlTemp2(lngC2).Text, "RULE -") > 0 Then
       If objFQ.PFAccountCode <> "" Then
          If InStr(1, xmlnlTemp2(lngC2).Text, "APF/") > 0 Then
             For intI = InStr(1, xmlnlTemp2(lngC2).Text, "APF/") + 4 To Len(xmlnlTemp2(lngC2).Text)
                 'JiYong - V1.2.1 (Urgent Fix) 201001102 - IR 6 - Pick Up Rule ID in Fare Rules
                 'Change RuleID to string
                 'If IsNumeric(Mid$(xmlnlTemp2(lngC2).Text, intI, 1)) Then
                 If IsAlphaNum(Mid$(xmlnlTemp2(lngC2).Text, intI, 1)) Then
                    'intRuleID = intRuleID & Mid$(xmlnlTemp2(lngC2).Text, intI, 1)
                    strRuleID = strRuleID & Mid$(xmlnlTemp2(lngC2).Text, intI, 1)
                 Else
                    Exit For
                 End If
             Next
          Else
             'JiYong - V1.2.1 (Urgent Fix) 201001102 - IR 6 - Pick Up Rule ID in Fare Rules
             'Use RuleID after / (Example: RULE - 901/SCB3, Pick up RuleID as SCB3)
             'Change Rule ID to string
             'For intI = InStr(1, xmlnlTemp2(lngC2).Text, "RULE -") + 7 To Len(xmlnlTemp2(lngC2).Text)
             If InStr(1, xmlnlTemp2(lngC2).Text, "/") > 0 Then
                For intI = InStr(1, xmlnlTemp2(lngC2).Text, "/") + 1 To Len(xmlnlTemp2(lngC2).Text)
                    'If IsNumeric(Mid$(xmlnlTemp2(lngC2).Text, intI, 1)) Then
                    If IsAlphaNum(Mid$(xmlnlTemp2(lngC2).Text, intI, 1)) Then
                       'intRuleID = intRuleID & Mid$(xmlnlTemp2(lngC2).Text, intI, 1)
                       strRuleID = strRuleID & Mid$(xmlnlTemp2(lngC2).Text, intI, 1)
                    Else
                       Exit For
                    End If
                Next
            End If
          End If
          objFQ.RuleID = strRuleID
       End If
    End If
    'JiYong - V1.2.1 (Urgent Fix) 201001102 - IR 6 - Pick Up Rule ID in Fare Rules
    'Move this code back to if else above
    'objFQ.RuleID = intRuleID
    
    
    If InStr(xmlnlTemp2(lngC2).Text, "3.  SEASONALITY") > 0 Then bolAfOthNote = True
    If bolAfOthNote = False Then
    If InStr(xmlnlTemp2(lngC2).Text, ":") <> 0 Then
        
        strOthNotesText = Trim(Mid(xmlnlTemp2(lngC2).Text, InStr(xmlnlTemp2(lngC2).Text, ":") + 1))
        If Len(strOthNotesText) > 0 Then
            If Mid(strOthNotesText, Len(strOthNotesText), 1) = "*" Then
                strOthNotesText = Mid(strOthNotesText, 1, Len(strOthNotesText) - 1)
            End If
        End If
        Select Case Trim(Left(xmlnlTemp2(lngC2).Text, InStr(xmlnlTemp2(lngC2).Text, ":") - 1))
            'Case "ENDORSEMENTS"
            '   objFC.EndorsementAdd strothnotestext, 1
    
            Case "PFBC"
                'added on 07/02/05: check for multiple PFBC in a single contract
                '''''
                'If InStr(strOthNotesText, "/") > 0 Then
                '    multiPFBC = Split(strOthNotesText, "/")
                '    If UBound(multiPFBC) = xmlnlTemp2.length Then
                '        objFC.PriceFBC = multiPFBC(lngC2)
                '    Else
                '        objFC.PriceFBC = ""
                '    End If
                'Else
                objFC.PriceFBC = strOthNotesText
                'End If
            Case "TFBC"
                objFC.TicketFBC = strOthNotesText
            Case "VC"
                objFC.ValueCode = strOthNotesText
            Case "TCOM"
                If IsNumeric(strOthNotesText) Then
                    objFC.CommissionOnTicket = strOthNotesText
                Else
                    objFC.CommissionOnTicket = 0
                End If
            Case "NRCC"
                If UCase(strOthNotesText) = "YES" Then
                    objFC.NRCC = True
                Else
                    objFC.NRCC = False
                End If
            Case "FOPC", "FOP"
                If UCase(strOthNotesText) <> "INVAGT" Then
                    objFC.FOPCode = strOthNotesText
                End If
            Case "HPF", "HIP"
                If UCase(strOthNotesText) = "YES" _
                    And (objFQ.JourneyType = "OJ" Or objFQ.JourneyType = "RW") Then
                    objFC.HPFApplies = True
                Else
                    objFC.NRCC = False
                End If
            Case "AWL"
                If UCase(strOthNotesText) = "YES" Then
                    objFC.WLAllow = True
                Else
                    objFC.WLAllow = False
                End If
            Case "ETR", "ET"
                If UCase(strOthNotesText) = "YES" Then
                    objFC.ETRequired = True
                Else
                    objFC.ETRequired = False
                End If
            Case "PTS", "PTSURCHARGE"
                If IsNumeric(strOthNotesText) Then
                    objFC.PaperTktSurcharge = strOthNotesText
                Else
                    If IsNumeric(Mid(strOthNotesText, 4)) Then
                        objFC.PaperTktSurcharge = Mid(strOthNotesText, 4)
                    Else
                        objFC.PaperTktSurcharge = 0
                    End If
                End If
            Case "MPROC"
                If UCase(strOthNotesText) = "YES" Then
                    objFC.ManualProcRequired = True
                Else
                    objFC.ManualProcRequired = False
                End If
            Case "NFAR", "TFAR"
                Select Case UCase(strOthNotesText)
                    Case "APF", "ITC", "ITN", "BTC", "BTN"
                        objFC.FareOnTicket = strOthNotesText
                    Case Else
                        objFC.FareOnTicket = "PUB"
                End Select
            Case "EB2"
                'strEndorseText.Add strOthNotesText
                objFC.EndorsementAdd strOthNotesText, 2
            Case "EB3", "FC1"
                'strEndorseText.Add strOthNotesText
                objFC.EndorsementAdd strOthNotesText, 3
            Case "EB4", "FC2"
                'strEndorseText.Add strOthNotesText
                objFC.EndorsementAdd strOthNotesText, 4
            Case "EB5", "FC3"
                'strEndorseText.Add strOthNotesText
                objFC.EndorsementAdd strOthNotesText, 5
            Case "TC"
                objFC.TourCode = strOthNotesText
                objFQ.ITNum = objFC.TourCode
        End Select
    End If
    End If
Next lngC2

'If Len(str1stEndorseText) > 0 Then
'    objFC.EndorsementAdd str1stEndorseText, 1
'End If
'For lngC2 = 1 To strEndorseText.Count
'    objFC.EndorsementAdd strEndorseText(lngC2), lngC2 + objFC.EndorsementCount
'Next

    
    ' NEXT HANDLES CORPORATE CONTRACT INFO
    Set rsContracts = Nothing
    'Modified on 170807: Retrieve contract base on faretype else tourcode will appear even in Publish fare
    Set rsContracts = gdbConn.Execute("SELECT * FROM tblClientAirContracts WHERE [CN] = '" & CN & "' AND Carrier = '" & objFC.Vendor & "'" & IIf(gstrAgcyCountryCode = "SG" And objFC.Vendor <> "SQ", " And faretype='" & strPFType(lngPF) & "' ", ""))
    'Set rsContracts = gdbConn.Execute("SELECT * FROM tbllngPFClientAirContracts WHERE [CN] = '" & CN & "' AND Carrier = '" & objFC.Vendor & "' )
 
    With rsContracts
        'strsql = "Carrier = '" & objFC.Vendor & "'"
        'If Not .BOF Then .MoveFirst
        '.Find strsql
    If Not .EOF Then
            Select Case ![Action]
                Case 1
                    'If lngC1 = 0 And ![TourCode] <> "" Then objFQ.ITNum = ![TourCode] & ""
                    If objFQ.Cat35 = False Then
                       If intTemp = 0 And ![TourCode] <> "" Then objFQ.ITNum = ![TourCode] & ""
                    End If
                    If ![FOPCode] <> "" Then objFC.FOPCode = ![FOPCode] & ""
                    'If ![Endors] <> "" Then objFC.Endorsement(1) = ![Endors]
                    If ![TktDesig] <> "" Then objFC.TktDesignator = ![TktDesig] & ""
                    If objFQ.PFAccountCode = ![PFCode] And ![FareOnTkt] <> "" Then objFC.FareOnTicket = ![FareOnTkt] & ""
                Case 2
                    'If lngC1 = 0 Then objFQ.ITNum = ![TourCode] & "" & objFQ.ITNum
                    If objFQ.Cat35 = False Then
                       If intTemp = 0 Then objFQ.ITNum = ![TourCode] & "" & objFQ.ITNum
                    End If
                    objFC.FOPCode = ![FOPCode] & "" & objFC.FOPCode
                    'objFC.Endorsement (1) = ![Endors] & "" & objfc.Endorsement(1)
                    objFC.TktDesignator = ![TktDesig] & "" & objFC.TktDesignator
                    If objFQ.PFAccountCode = ![PFCode] And ![FareOnTkt] <> "" Then objFC.FareOnTicket = ![FareOnTkt] & "" 'This field will always override
                Case 3
                    'If lngC1 = 0 Then objFQ.ITNum = objFQ.ITNum & ![TourCode] & ""
                    If objFQ.Cat35 = False Then
                       If intTemp = 0 Then objFQ.ITNum = objFQ.ITNum & ![TourCode] & ""
                    End If
                    objFC.FOPCode = objFC.FOPCode & ![FOPCode] & ""
                    'objFC.Endorsement (1) = objfc.Endorsement(1) & ![Endors] & ""
                    objFC.TktDesignator = objFC.TktDesignator & ![TktDesig] & ""
                    If ![FareOnTkt] <> "" Then objFC.FareOnTicket = ![FareOnTkt] & ""
                Case Else
                    If objFQ.Cat35 = False Then
                       If objFQ.ITNum = "" Then objFQ.ITNum = ![TourCode] & ""
                    End If
                    If objFC.FOPCode = "" Then objFC.FOPCode = ![FOPCode] & ""
                    'If ![Endors] <> "" Then objFC.Endorsement(1) = ![Endors]
                    If objFC.TktDesignator = "" Then objFC.TktDesignator = ![TktDesig] & ""
                    
            End Select
        End If
    End With
Else
    objFC.RuleTextAdd "SAME RULE AS COMPONENT " & lngRuleFC
    objFC.TktInfoApplies = objFQ.FareComponent(lngRuleFC).TktInfoApplies
    objFC.PriceFBC = objFQ.FareComponent(lngRuleFC).PriceFBC
    objFC.TicketFBC = objFQ.FareComponent(lngRuleFC).TicketFBC
    objFC.TktDesignator = objFQ.FareComponent(lngRuleFC).TktDesignator
    objFC.ValueCode = objFQ.FareComponent(lngRuleFC).ValueCode
    objFC.CommissionOnTicket = objFQ.FareComponent(lngRuleFC).CommissionOnTicket
    objFC.NRCC = objFQ.FareComponent(lngRuleFC).NRCC
    objFC.FOPCode = objFQ.FareComponent(lngRuleFC).FOPCode
    objFC.HPFApplies = objFQ.FareComponent(lngRuleFC).HPFApplies
    objFC.WLAllow = objFQ.FareComponent(lngRuleFC).WLAllow
    objFC.OpenSegAllow = objFQ.FareComponent(lngRuleFC).OpenSegAllow
    objFC.ETRequired = objFQ.FareComponent(lngRuleFC).ETRequired
    objFC.PaperTktSurcharge = objFQ.FareComponent(lngRuleFC).PaperTktSurcharge
    objFC.ManualProcRequired = objFQ.FareComponent(lngRuleFC).ManualProcRequired
    objFC.FareOnTicket = objFQ.FareComponent(lngRuleFC).FareOnTicket

    If objFQ.FareComponent(lngRuleFC).EndorsementCount > 0 Then objFC.EndorsementAdd objFQ.FareComponent(lngRuleFC).Endorsement(1)
    If objFQ.FareComponent(lngRuleFC).EndorsementCount > 1 Then objFC.EndorsementAdd objFQ.FareComponent(lngRuleFC).Endorsement(2)
    If objFQ.FareComponent(lngRuleFC).EndorsementCount > 2 Then objFC.EndorsementAdd objFQ.FareComponent(lngRuleFC).Endorsement(3)
    If objFQ.FareComponent(lngRuleFC).EndorsementCount > 3 Then objFC.EndorsementAdd objFQ.FareComponent(lngRuleFC).Endorsement(4)
    If objFQ.FareComponent(lngRuleFC).EndorsementCount > 4 Then objFC.EndorsementAdd objFQ.FareComponent(lngRuleFC).Endorsement(5)

End If

    objFQ.FareComponentAdd objFC
    Set objFC = Nothing
End If
Next lngC1


Set xmlnlTemp1 = xmldomFQ.selectNodes("//Surcharge")

For lngC1 = 0 To xmlnlTemp1.length - 1
  'Hasn't test
  If Format(xmlnlTemp1(lngC1).selectSingleNode("UniqueKey").Text, "0000") = Format(i + 1, "0000") Then
    Set xmlnlTemp2 = xmlnlTemp1.Item(lngC1).childNodes
    Set objSC = New Surcharge
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
                    intTotDecPos = CInt(.Text)
            End Select
        End With
    Next lngC2
    objSC.Amount = PlaceDecimal(intTotDecPos, lngTotAmt)
    objFQ.SurchargeAdd objSC
    Set objSC = Nothing
  End If
Next lngC1

If lngPF > 0 Then
    
    If failNum = 0 Then
    'mcolFQ.FQ(1).BaseAmount
        If objFQ.BaseAmount < mcolFQ.FQ(1).BaseAmount Then
        
            '''mcolFQ.Add objFQ, strPFCodes(lngPF), 1
            mcolFQ.FQAdd objFQ, 1
        'Added on 01 Nov 07 by JY ANG,  fix to pick corporate fare instead of market fare when the fare amount is the same
        ElseIf objFQ.BaseAmount = mcolFQ.FQ(1).BaseAmount Then
            If objFQ.PFAccountCode <> "" And mcolFQ.FQ(1).PFAccountCode = "" Then
               mcolFQ.FQAdd objFQ, 1
            Else
               mcolFQ.FQAdd objFQ
            End If
        ElseIf mcolFQ.FQ(1).UnableToQuote = True Then
            '''mcolFQ.Add objFQ, strPFCodes(lngPF)
            mcolFQ.FQAdd objFQ, 1
        Else
            mcolFQ.FQAdd objFQ
        End If
    Else
        '''mcolFQ.Add objFQ, strPFCodes(lngPF)
        mcolFQ.FQAdd objFQ
    End If
Else
    '''mcolFQ.Add objFQ, strPFCodes(lngPF)
    mcolFQ.FQAdd objFQ
End If

Set xmlnlTemp1 = Nothing
Set xmlnlTemp2 = Nothing
Set xmlnlTemp3 = Nothing
Set xmlnlTemp4 = Nothing
Set objFQ = Nothing

skip_the_rest:

Next lngPF

'Set objPX = New FQ
'objPX.FQAdd mcolFQ
'Set mcolPX = New Collection

'21 Dec Assign back the actual Passenger index
'FareOverride '12012005
NextPx:
    i = intPNum

mcolPX.Add mcolFQ, i + 1 & "Px"
Set mcolFQ = Nothing
Next i


If strErrMsg <> "" Then
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strErrMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End If


'GoTo startagain

rsContracts.Close
Set rsContracts = Nothing

Exit Function

ErrProc:
    Err.Raise Err.Number, "CWT_Galileo.GetFareQuote", Err.Description
End Function


Private Function ConvertDate(DateString As String, Optional TimeString As String = "0001") As Date
    ConvertDate = CDate(Format(DateString & Format(TimeString, "0000"), "0000-00-00 00:00"))
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

Public Function getFQFromTbl(Count As Long, PxNum() As Integer, strRecLoc As String, intSegID As Integer) As FQ

Dim objFQ As FareQuote
Dim objTax As Tax
Dim objFQSeg As FareSegment
Dim objFC As FareComponent
Dim objSC As Surcharge
Dim STRSQL As String
Dim rsRec As New ADODB.Recordset
Dim rsRec2 As New ADODB.Recordset
Dim rsRec3 As New ADODB.Recordset

Dim lngC As Long
Dim lngD As Long
Dim lngI As Long
Dim intPxNo As Integer

Set mcolPX = New Collection
For lngI = 1 To Count
    
    intPxNo = PxNum(lngI)
    
    STRSQL = "Select * from tblFareQuote where [RecLoc] = '" & strRecLoc & "'"
    STRSQL = STRSQL & " and [PxID]= " & intPxNo & " and [SegID]=" & intSegID & ""
     
    Set rsRec3 = gdbConn.Execute(STRSQL)
    Set mcolFQ = New FQ
    While Not rsRec3.EOF
    Set objFQ = New FareQuote
    
       With objFQ
            .PxNum = rsRec3![PxID]
            .BaseAmount = rsRec3![BaseAmount]
            .BaseCurrency = rsRec3![BaseCurrency]
            .Commission = rsRec3![Commission]
            .DiscountAmt = rsRec3![DiscountAmt]
            .EquivAmount = rsRec3![EquivAmount]
            .EquivCurrency = rsRec3![EquivCurrency]
            .FareConstructText = rsRec3![FareConstructText]
            .HighFareComponent = rsRec3![HighFareComponent]
            .HPFApplies = rsRec3![HPFApplies]
            .ISO = rsRec3![ISO]
            .ITNum = rsRec3![ITNum]
            .JourneyNum = rsRec3![JourneyNum]
            .JourneyType = rsRec3![JourneyType]
             If rsRec3![LastTktDate] <> "" Then .LastTktDate = rsRec3![LastTktDate]
            .LowFare = rsRec3![LowFare]
            .MerchAmt = rsRec3![MerchAmt]
            .NetAmount = rsRec3![NetAmount]
            .PFAccountCode = rsRec3![PFAccountCode]
            If rsRec3![PFType] <> "" Then
                .PFFareType = rsRec3![PFType]
            End If
            .PIC = rsRec3![PIC]
            .PrivateFare = rsRec3![PrivateFare]
            .QuoteType = rsRec3![QuoteType]
            '.ROE = rsRec![ROE]
            .SegmentSelected = rsRec3![SegmentSelected]
            .SegmentSelectString = rsRec3![SegmentSelectText]
            .SellAmount = rsRec3![SellAmount]
            .StoreFare = rsRec3![StoreFare]
            .TCNum = rsRec3![TCNum]
            .TotalCurrency = rsRec3![TotalCurrency]
            .TotAmount = rsRec3![TotAmount]
            .TransactionFee = rsRec3![TransactionFee]
            .UnableToQuote = rsRec3![UnableToQuote]
            If IsNull(rsRec3![NRCC]) Then
                .NRCC = False
            Else
                .NRCC = rsRec3![NRCC]
            End If
            If Not IsNull(rsRec3![ClientType]) Then
              .ClientType = rsRec3![ClientType]
            End If
            .ComPct = IIf(IsNull(rsRec3![ComPct]), 0, rsRec3![ComPct])
            .DiscountPct = IIf(IsNull(rsRec3![DiscountPct]), 0, rsRec3![DiscountPct])
            .MerchPct = IIf(IsNull(rsRec3![MerchPct]), 0, rsRec3![MerchPct])
            If Not IsNull(rsRec3![OverrideTF]) Then
                .OverrideTF = rsRec3![OverrideTF]
            End If
            If Not IsNull(rsRec3![International]) Then
                .International = rsRec3![International]
            End If
            If Not IsNull(rsRec3![OverrideConx]) Then
                .OverrideConx = rsRec3![OverrideConx]
            End If
            If Not IsNull(rsRec3![FQPCC]) Then
              .FQPCC = rsRec3![FQPCC]
            End If
            If Not IsNull(rsRec3![PlatCarrier]) Then
              .PlatCarrier = rsRec3![PlatCarrier]
            End If
            If Not IsNull(rsRec3![Cat35]) Then
              .Cat35 = rsRec3![Cat35]
            End If
            If Not IsNull(rsRec3![Cat35CommType]) Then
              .CommType = rsRec3![Cat35CommType]
            End If
            If Not IsNull(rsRec3![PTkt]) Then
              .PTkt = rsRec3![PTkt]
            End If
            If Not IsNull(rsRec3![RuleID]) Then
              .RuleID = rsRec3![RuleID]
            End If
            .HiddenComm = rsRec3![HiddenComm]
            If Not IsNull(rsRec3![ActualNetAmount]) Then
              .ActualNetAmount = rsRec3![ActualNetAmount]
            End If
            
        End With


STRSQL = "select * from tblFareComponent where RecLoc = '" & strRecLoc & "' and PxID = " & intPxNo & " and [SegID]=" & intSegID & " order by FCSeq"

'Assigning FC obj

Set rsRec = gdbConn.Execute(STRSQL)

While Not rsRec.EOF
    Set objFC = New FareComponent
    With objFC
        .Amount = rsRec![Amount]
        .CommissionOnTicket = rsRec![CommissionOnTicket]
        .CurrencyCode = rsRec![CurrencyCode]
        .Destinantion = rsRec![Destination]
        .DestinationCountry = rsRec![DestinationCountry]
        .DestinationRegion = rsRec![DestinationRegion]
        .DirectionalInd = rsRec![DirectionalInd]
        .ETRequired = rsRec![ETRequired]
        .FareOnTicket = rsRec![FareOnTicket]
        .FBC = rsRec![FBC]
        .FOPCode = rsRec![FOPCode]
        .HPFApplies = rsRec![HPFApplies]
        .ManualProcRequired = rsRec![ManualProcRequired]
        .MPM = rsRec![MPM]
        .Net = rsRec![Net]
        .NRCC = rsRec![NRCC]
        .OpenSegAllow = rsRec![OpenSegAllow]
        .Origin = rsRec![Origin]
        .OriginCountry = rsRec![OriginCountry]
        .OriginRegion = rsRec![OriginRegion]
        .OTWCarrier = rsRec![OTWCarrier]
        .PaperTktSurcharge = rsRec![PaperTktSurcharge]
        .PriceFBC = rsRec![PriceFBC]
        .RuleNum = rsRec![RuleNum]
        .RuleTextAdd rsRec![RuleText]
        .TicketFBC = rsRec![TicketFBC]
        .TktDesignator = rsRec![TktDesignator]
        .TktInfoApplies = rsRec![TktInfoApplies]
        .TPM = rsRec![TPM]
        .ValueCode = rsRec![ValueCode]
        .Vendor = rsRec![Vendor]
        .WLAllow = rsRec![WLAllow]
    End With

    STRSQL = "select * from tblFCEndorse where RecLoc = '" & strRecLoc & "' And PxID = " & intPxNo & " and [SegID]=" & intSegID & " and [FCSeq]=" & rsRec![FCSeq] & " order by EndorseSeq"
    
  'Assign back FCEndorse
 
    Set rsRec2 = gdbConn.Execute(STRSQL)
    lngD = 0
    While Not rsRec2.EOF
    lngD = lngD + 1
       objFC.EndorsementAdd rsRec2![EndorseText], lngD
    rsRec2.MoveNext
    Wend
    
    objFQ.FareComponentAdd objFC
    Set objFC = Nothing
    
rsRec.MoveNext
Wend

STRSQL = "select * from tblFareSeg where RecLoc = '" & strRecLoc & "' and PxID = " & intPxNo & " and [SegID]=" & intSegID & " order by SegSeq"

''Adding tblFareSeg

Set rsRec = gdbConn.Execute(STRSQL)
    While Not rsRec.EOF
            Set objFQSeg = New FareSegment
            With objFQSeg
                .ArrCityCode = rsRec![ArrCityCode]
                .BagInfo = rsRec![BagInfo]
                .COS = rsRec![COS]
                .DepCityCode = rsRec![DepCityCode]
                .FBC = rsRec![FBC]
                .FlightNum = rsRec![FlightNum]
                .InfoText = rsRec![InfoText]
                If rsRec![NVA] <> "" Then .NVA = rsRec![NVA]
                If rsRec![NVB] <> "" Then .NVB = rsRec![NVB]
                .Stopover = rsRec![Stopover]
                .OverridePFBC = rsRec![OverridePFBC]
                .TD = rsRec![TD]
                .Vendor = rsRec![Vendor]
                If rsRec![DepDate] <> "" Then .DepDate = rsRec![DepDate]
                objFQ.FareSegAdd objFQSeg
            End With
            Set objFQSeg = Nothing
    rsRec.MoveNext
    Wend

STRSQL = "select * from tblFareSurcharge where RecLoc = '" & strRecLoc & "' and PxID = " & intPxNo & " and [SegID]=" & intSegID & " order by SurchNum"

''Adding tblFareSurcharge

Set rsRec = gdbConn.Execute(STRSQL)
    While Not rsRec.EOF
            Set objSC = New Surcharge
            With objSC
                .Amount = rsRec![Amount]
                .CurrencyCode = rsRec![CurrencyCode]
                .RelatedSegment = rsRec![RelatedSegment]
                objFQ.SurchargeAdd objSC
            End With
            Set objSC = Nothing
    rsRec.MoveNext
    Wend
STRSQL = "select * from tblFareTax where RecLoc = '" & strRecLoc & "' and PxID = " & intPxNo & " and [SegID]=" & intSegID & " order by TaxNum"

'Assign tax obj

    Set rsRec = gdbConn.Execute(STRSQL)
    While Not rsRec.EOF
        Set objTax = New Tax
        With objTax
            .Amount = rsRec![Amount]
            .TaxCode = rsRec![TaxCode]
            objFQ.TaxAdd objTax
        End With
        Set objTax = Nothing
    rsRec.MoveNext
    Wend
    
    mcolFQ.FQAdd objFQ
    Set objFQ = Nothing
    rsRec3.MoveNext
    Wend


mcolPX.Add mcolFQ, lngI & "Px"
Set mcolFQ = Nothing

Next lngI

rsRec.Close
rsRec2.Close
rsRec3.Close


Set rsRec = Nothing
Set rsRec2 = Nothing
Set rsRec3 = Nothing
End Function
Private Function getCount(CN As String, strCarriers() As String) As Integer
Dim rsCount As New ADODB.Recordset
Dim STRSQL As String
Dim lngC1 As Long
   STRSQL = "SELECT count([CN]) AS PFCount FROM tblClientAirContracts WHERE [CN] = '" & CN & "' AND ("
   For lngC1 = 0 To UBound(strCarriers)
       STRSQL = STRSQL & IIf(InStr(STRSQL, "[Carrier]") > 0, "OR ", "") & " [Carrier] = '" & strCarriers(lngC1) & "' "
   Next
   STRSQL = STRSQL & ")"
   Set rsCount = gdbConn.Execute(STRSQL)
   getCount = rsCount!PFCount
   
End Function
