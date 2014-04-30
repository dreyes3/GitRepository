Attribute VB_Name = "modOthSvcs"
Public Declare Function FindExecutable _
    Lib "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, _
    ByVal lpDirectory As String, _
    ByVal lpResult As String) As Long
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
Public gobjEO As EO
Public gobjPreEO As EO
Public gstrPreEOType As String
Public gbolEOAmend As Boolean
Public gbolIgnoreEO As Boolean
Public gStartOthSvcsTime As Date
Public gSysStartOthSvcsTime As Date


Public Function ValidIR(Optional strTemp As String) As Boolean
Dim strMsg As String
Dim strFChar As String
Dim strRmk As String


If IsNumeric(strTemp) = True Then
 strMsg = strMsg & "Whole number is not allow is GDS." & Chr(13)
End If
strFChar = Mid(strTemp, 1, 1)

If IsNumeric(strTemp) = False And IsNumeric(strFChar) = True Then
  strMsg = strMsg & "First character cannot be numeric." & Chr(13)
End If

If strMsg = "" Then
    ValidIR = True
Else
    'MsgBox strMsg, vbApplicationModal + vbExclamation, "Travel Pro"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
End If
End Function
    

'Modified on 21/03/2005
Public Sub UpdateExchOrdDB(ExchangeOrderObj As EO, EOType As String)

Dim lngC As Long
Dim strTemp As String
Dim strSQL As String
Dim rsEO As New ADODB.Recordset
Dim strMsg As String

'Preethi - V1.2.12 20120416 - CR127 - Quick Wins - Consolidator Ticket (HKSG)
Dim strTemp2 As String
Dim strAirVendor As String
Dim strAirVendorLocatar As String


On Error GoTo EOUpdateError

If gbolEOAmend Then
   gdbConn.BeginTrans
   gbolBeginTrans = True
   strSQL = "SELECT * FROM tblExchangeOrder where exchangeid = '" & ExchangeOrderObj.EONumber & "'"
Else
   strSQL = "tblExchangeOrder"
End If
'strSQL = "SELECT * FROM tblExchangeOrder where exchangeid = '" & ExchangeOrderObj.EONumber & "'"
'strSQL = "tblExchangeOrder"
'Set rsEO = gdbExchOrder.OpenRecordset(strSQL)
'Insert record here
'If rsEO.EOF Then
'    strSQL = "SELECT * FROM tblExchangeOrder"
'    Set rsEO = gdbExchOrder.OpenRecordset(strSQL)
rsEO.Open strSQL, gdbConn, adOpenDynamic, adLockOptimistic


    With rsEO
      If Not gbolEOAmend Then
        .AddNew
        ![ExchangeID] = ExchangeOrderObj.EONumber
      End If
        ![ProductCode] = ExchangeOrderObj.ProductCode
        ![VendorCode] = ExchangeOrderObj.VendorCode
        ![PNR] = ExchangeOrderObj.PNRRecLoc
        ![CN] = ExchangeOrderObj.CN
        
        ![Name] = ExchangeOrderObj.PaxName
        ![TransDate] = Date 'DateAdd("M", -3, ExchangeOrderObj.ServiceDate)
        ![Cost] = ExchangeOrderObj.Cost
        ![Tax1] = ExchangeOrderObj.Tax(1).Amount
        If UCase(ExchangeOrderObj.Tax(1).Code) = "GST" Then
           ![TAXCODE1] = "G*"
        ElseIf UCase(ExchangeOrderObj.Tax(1).Code) = "VAT" Then
           ![TAXCODE1] = "V*"
        ElseIf Trim(ExchangeOrderObj.Tax(1).Code) <> "" Then
            'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
            ![TAXCODE1] = Trim(UCase(ExchangeOrderObj.Tax(1).Code))
        End If
        '![TaxCode1] = UCase(ExchangeOrderObj.Tax(1).Code)
        If ExchangeOrderObj.TaxCount = 2 Then
           ![Tax2] = ExchangeOrderObj.Tax(2).Amount
           '![TaxCode2] = ExchangeOrderObj.Tax(2).Code
           If UCase(ExchangeOrderObj.Tax(2).Code) = "GST" Then
              ![TaxCode2] = "G*"
           ElseIf UCase(ExchangeOrderObj.Tax(2).Code) = "VAT" Then
              ![TaxCode2] = "V*"
           ElseIf Trim(ExchangeOrderObj.Tax(2).Code) <> "" Then
            'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
            ![TaxCode2] = Trim(UCase(ExchangeOrderObj.Tax(2).Code))
           End If
        End If
        ![SellPrice] = ExchangeOrderObj.SellPrice
        ![Commission] = ExchangeOrderObj.CommissionAmt
        ![NETTCOSTGST] = ExchangeOrderObj.NettGST
        'ZhiSam - V1.2.11.20120516 - CR 128 - Quick Wins - VisaCost (HKSG Desktop)
        If UCase(gstrAgcyCountryCode) = "HK" Then
            ![VendorHandling] = ExchangeOrderObj.VendorHandling
        End If
        
        strTemp = ""
        For lngC = 1 To ExchangeOrderObj.DescriptionLinesCount
        '02062005
            'strTemp = strTemp & IIf(strTemp = "", "", ";") & ExchangeOrderObj.DescriptionLine(lngC)
            strTemp = strTemp & IIf(strTemp = "", "", vbCrLf) & ExchangeOrderObj.DescriptionLine(lngC)
        Next
        ![Description] = strTemp
        
        strTemp = ""
        For lngC = 1 To ExchangeOrderObj.RemarkCount
            '02062005
            'strTemp = strTemp & IIf(strTemp = "", "", ";") & ExchangeOrderObj.Remark(lngC)
            strTemp = strTemp & IIf(strTemp = "", "", vbCrLf) & ExchangeOrderObj.Remark(lngC)
        Next
        ![Remarks] = strTemp

        
        ![FOP] = ExchangeOrderObj.FOP
        
        ![BillingDescription] = ExchangeOrderObj.BillingDescription
        If gbolEOAmend = False Then
           ![CreateDtTm] = ExchangeOrderObj.CreateDtTm
           ![CreatedBy] = ExchangeOrderObj.CreatedBy
           !CreatedByPCC = ExchangeOrderObj.CreatedByPCC
        End If
        !LastAmendDtTm = ExchangeOrderObj.CreateDtTm
        !LastAmendBy = ExchangeOrderObj.CreatedBy
        !LastAmendByPCC = ExchangeOrderObj.CreatedByPCC
        !EOType = EOType
        !ContactPerson = ExchangeOrderObj.ContactPerson
        If InStr(1, UCase(EOType), "CHEQUE") <> 0 Then
           !Finance = "True"
        Else
           !Finance = "False"
        End If
        !AgentPhone = gstrAgcyPhone
        
        !ClientType = ExchangeOrderObj.ClientType
        !NettFare = ExchangeOrderObj.NettFare
        !PubFare = ExchangeOrderObj.PublishedFare
        !GrossFare = ExchangeOrderObj.GrossFare
        !Discount = ExchangeOrderObj.Discount
        !MerchantFee = ExchangeOrderObj.MerchFee
        !CWTAbsorb = ExchangeOrderObj.CWTAbsorb
        !TransactionFee = ExchangeOrderObj.TranxFee
        !TktNum = ExchangeOrderObj.TicketNumber & IIf(ExchangeOrderObj.ConjunctTicket <> "", "-" & ExchangeOrderObj.ConjunctTicket, "") ' ExchangeOrderObj.TktNum
        !ListBoxRemark = ExchangeOrderObj.ListBoxRem
        'CS Change EC
        '!MIData = ExchangeOrderObj.RF & ";" & ExchangeOrderObj.LF & ";" & ExchangeOrderObj.EC & ";" & ExchangeOrderObj.FF7 & ";" & ExchangeOrderObj.FF8 & ";" & ExchangeOrderObj.FF26 & ";" & ExchangeOrderObj.FF19 & ";" & ExchangeOrderObj.FF38
        '!MIData = ExchangeOrderObj.RF & ";" & ExchangeOrderObj.LF & ";" & ExchangeOrderObj.MS & ";" & ExchangeOrderObj.FF7 & ";" & ExchangeOrderObj.FF8 & ";" & "" & ";" & ExchangeOrderObj.FF19 & ";" & ExchangeOrderObj.FF38 & ";" & ExchangeOrderObj.rs & ";" & ExchangeOrderObj.FF41
        !MIData = ExchangeOrderObj.RF & ";" & ExchangeOrderObj.LF & ";" & ExchangeOrderObj.MS & ";" & ExchangeOrderObj.FF7 & ";" & ExchangeOrderObj.FF8 & ";" & "" & ";" & ExchangeOrderObj.FF81 & ";" & ExchangeOrderObj.FF38 & ";" & ExchangeOrderObj.rs
        !PickUpFrom = ExchangeOrderObj.PickUpFrom
        !PickUpTo = ExchangeOrderObj.PickUpTo
        !PickUpTime = ExchangeOrderObj.PickUpTime
        !PickUpFlight = ExchangeOrderObj.PickUpFlight
        !ReturnFrom = ExchangeOrderObj.ReturnFrom
        !ReturnTo = ExchangeOrderObj.ReturnTo
        !ReturnTime = ExchangeOrderObj.ReturnTime
        !ReturnFlight = ExchangeOrderObj.ReturnFlight
        !SegmentSelect = ExchangeOrderObj.SegSelect
        !AdditionalInfo = ExchangeOrderObj.AdditionalInfo
        !VisaInfo = ExchangeOrderObj.VisaInfo
        !VendorFax = ExchangeOrderObj.FaxNo
        !VendorEmail = ExchangeOrderObj.Email
        'preethi – V1.2.6 20110905 – CR99 - Add Option for Fare Type in EO
        If ExchangeOrderObj.FareType <> 0 Then
           !FareType = ExchangeOrderObj.FareType
        End If
        'preethi – V1.2.6 20110905 – CR98 - Reissue Ticket Box in EO
        If ExchangeOrderObj.TktNumber = "" Then
           !OriTktNum = "NULL"
        Else
           !OriTktNum = ExchangeOrderObj.TktNumber
        End If
        'If ExchangeOrderObj.VendorCode = "999999" Then
        If ExchangeOrderObj.Misc = True Then
           !VendorInfo = ExchangeOrderObj.VendorName & vbCrLf & _
                         ExchangeOrderObj.Address1 & vbCrLf & _
                         ExchangeOrderObj.Address2 & vbCrLf & _
                         ExchangeOrderObj.City & vbCrLf & _
                         ExchangeOrderObj.Country & vbCrLf & _
                         ExchangeOrderObj.Email & vbCrLf & _
                         ExchangeOrderObj.FaxNo & vbCrLf & _
                         ExchangeOrderObj.ContactNum
          
        End If
        
        'Preethi - V1.2.12 20120416 - CR127 - Quick Wins - Consolidator Ticket (HKSG)
        strTemp = ""
        strTemp2 = ""

        Set gobjPNR = New CWT_GalileoPNR3.PNR
        With gobjPNR
            Call .loadPNR
            For lngC = 1 To .VendorRecLocCount
                strTemp = strTemp & IIf(strTemp = "", "", ";") & .VendorRecLoc(lngC).Vendor
                strTemp2 = strTemp2 & IIf(strTemp2 = "", "", ";") & .VendorRecLoc(lngC).RecLoc
            Next
        End With

        ![AirlineVendor] = strTemp
        ![AirlineVendorLocator] = strTemp2
        If frmOthSvcs.datProducts.Recordset![Type] = "CT" Then
           strTemp = ExchangeOrderObj.SegSelect
        Else
           strTemp = ""
        End If
        
        strAirVendor = ""
        strAirVendorLocatar = ""
        Call GetAirlineVendor(strTemp, strAirVendor, strAirVendorLocatar)

        ![AirlineVendorSelected] = strAirVendor
        ![AirlineVendorLocatorSelected] = strAirVendorLocatar
        
        .Update
    End With
'Update Record here
'Else
'    With rsEO
'        .Edit
'        '![ExchangeID] = ExchangeOrderObj.EONumber
'        ![ProductCode] = ExchangeOrderObj.ProductCode
'        ![VendorCode] = ExchangeOrderObj.VendorCode
'        ![PNR] = ExchangeOrderObj.PNRRecLoc
'        ![CN] = ExchangeOrderObj.CN
'        ![Name] = ExchangeOrderObj.PaxName
'        ![Date] = Date 'DateAdd("M", -3, ExchangeOrderObj.ServiceDate)
'        ![Cost] = ExchangeOrderObj.Cost
'        ![Tax1] = ExchangeOrderObj.Tax(1).Amount
'        If UCase(ExchangeOrderObj.Tax(1).Code) = "GST" Then
'           ![TaxCode1] = "G*"
'        ElseIf UCase(ExchangeOrderObj.Tax(1).Code) = "VAT" Then
'           ![TaxCode1] = "V*"
'        End If
'        '![TaxCode1] = UCase(ExchangeOrderObj.Tax(1).Code)
'        ![Tax2] = ExchangeOrderObj.Tax(2).Amount
'        '![TaxCode2] = ExchangeOrderObj.Tax(2).Code
'        If UCase(ExchangeOrderObj.Tax(2).Code) = "GST" Then
'           ![TaxCode2] = "G*"
'        ElseIf UCase(ExchangeOrderObj.Tax(2).Code) = "VAT" Then
'           ![TaxCode2] = "V*"
'        End If
'        ![SellPrice] = ExchangeOrderObj.SellPrice
'        ![Commission] = ExchangeOrderObj.CommissionAmt
'
'        strTemp = ""
'        For lngC = 1 To ExchangeOrderObj.DescriptionLinesCount
'            strTemp = strTemp & IIf(strTemp = "", "", ";") & ExchangeOrderObj.DescriptionLine(lngC)
'        Next
'        ![Description] = strTemp
'
'        strTemp = ""
'        For lngC = 1 To ExchangeOrderObj.RemarkCount
'            strTemp = strTemp & IIf(strTemp = "", "", ";") & ExchangeOrderObj.Remark(lngC)
'        Next
'        ![Remarks] = strTemp
'
'        ![FOP] = ExchangeOrderObj.FOP
'
'        ![BillingDescription] = ExchangeOrderObj.BillingDescription
'        ![CreateDtTm] = ExchangeOrderObj.CreateDtTm
'        ![CreatedBy] = ExchangeOrderObj.CreatedBy
'        !EOType = EOType
'
'        .Update
'    End With

'End If
rsEO.Close
If gbolBeginTrans Then
   gdbConn.CommitTrans
   gbolBeginTrans = False
End If
Exit Sub
EOUpdateError:
'MsgBox Err.Number & ": " & Err.Description & Chr(13) & "SQL STRING: (" & strSql & ")", vbCritical, "EO Update Error"
strMsg = Err.Number & ": " & Err.Description & Chr(13) & "SQL STRING: (" & strSQL & ")"
modMsgBox.OKMsg = "OK"
modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
gdbConn.RollbackTrans
rsEO.Close
End Sub
'Timer
Public Sub WriteOSToGDS(objEO As EO, ByVal productType As String, startTime As Date, Optional ByVal freefields As String)
    'format for FreeFields is {FF Number}-{data} (i.e '22-CHINA')  Multiple entries separated by '/'
Dim strTemp As String
Dim strTemp2 As String
Dim strMSLine As String
Dim strFOP() As String
Dim strFF() As String
Dim lngC As Long
Dim lngLen As Long
Dim strPNR As String
Dim mbytFFNum As Integer
Dim bolIsCC As Boolean
Dim strGST As String
Dim strMIData As String
Dim strSegNum As String
Dim rs As ADODB.Recordset
Dim strTrxnFeeCode As String
Dim strProduct As String
Dim strFOPTemp As String
Dim curDiscAmt As String
Dim strMIFF40 As String
Dim strMSFirstLine As String
Dim strDiscCode As String
Dim intLength As Integer
Dim strZero As String
'CS SBT Indicator
Dim strFFSBT As String
'CS Transaction Fee
Dim strFFTrans As String
'CS Change EC
Dim strRS As String
'CS Add FF41
'Dim strFF41 As String
Dim strMIOther As String
Dim strTFFOPTemp As String
Dim strSecLine As String

'Preethi - V1.2.6 20110907 - CR 90 - Change OBT Tool Code in FF35
Dim strBookingTool As String

strProduct = frmOthSvcs.dbcProducts.BoundText

Set rs = gdbConn.Execute("Select sortkey from tblProductCodes where productcode = '35'")
If rs.EOF Then
    strTrxnFeeCode = ""
Else
    strTrxnFeeCode = rs!SortKey & ""
End If
rs.Close
Set rs = gdbConn.Execute("Select sortkey from tblProductCodes where productcode = '50'")
If rs.EOF Then
    strDiscCode = ""
Else
    strDiscCode = rs!SortKey & ""
End If

rs.Close
'modified on 15Jun
If productType = "CT" Then 'And strProduct = "02" Then
    Set rs = gdbConn.Execute("Select Length from tbldocstruct where StructID='PO'")
    While Not rs.EOF
    
    intLength = intLength + rs![length]
    
    rs.MoveNext
    Wend
    
    For lngC = 1 To intLength
    strZero = strZero & "0"
    Next
End If
Set rs = Nothing




With objEO

strMSFirstLine = "/PC" & .ProductSortKey _
                & "/V" & .VendorCode _
                & getTicketNo(productType, strProduct) & "PX" & .PassengerID

'If UCase(gstrAgcyCountryCode) = "HK" Then
'strMSLine = "/PC" & .ProductCode
        If productType = "MS" And strProduct = "50" Then
            strSecLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & "-" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
                & "/SF" & "-" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
                & "/C" & "-" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
                & getSegmentSelected(strProduct, productType)
                'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
                'Pass in productType also into getSegmentSelected function
                'Move the modified code to the top due to the constraint of the _ symbol
                '& getSegmentSelected(strProduct)
                
                strSecLine = splitLongMSX(strSecLine)
                '& "/PO" & .EONumber '02062005
            '02062005
            'modified in 15Jun:
            'If frmOSMisc.optConsol = True Then
            '   strMSLine = strMSLine & "/PO" & .TicketNumber
            'End If
        '02062005
        'Modified on 15Jun
        ElseIf productType = "CT" Then 'And strProduct = "02" Then
            'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
            'Set to two decimal point if web fare is selected (For Hong Kong only)
            If gstrAgcyCountryCode = "HK" And .WebFareApplied = True Then
                strSecLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & Format(CStr(.SellPrice - .EOTaxTotal + .Discount), gstrHKWebFareDecimalPoint) _
                    & "/SF" & Format(CStr(.SellPrice - .EOTaxTotal + .Discount), gstrHKWebFareDecimalPoint) _
                    & "/C" & Format(CStr(.CommissionAmt), gstrHKWebFareDecimalPoint) _
                    & getSegmentSelected(strProduct, productType)
                    'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
                    'Pass in productType also into getSegmentSelected function
                    'Move the modified code to the top due to the constraint of the _ symbol
                    '& getSegmentSelected(strProduct)
            Else
                strSecLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & Format(CStr(.SellPrice - .EOTaxTotal + .Discount), gstrAgcyCurrFormat) _
                    & "/SF" & Format(CStr(.SellPrice - .EOTaxTotal + .Discount), gstrAgcyCurrFormat) _
                    & "/C" & Format(CStr(.CommissionAmt), gstrAgcyCurrFormat) _
                    & getSegmentSelected(strProduct, productType)
                    'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
                    'Pass in productType also into getSegmentSelected function
                    'Move the modified code to the top due to the constraint of the _ symbol
                    '& getSegmentSelected(strProduct)
            End If
             strSecLine = splitLongMSX(strSecLine)
             'Added by JiYong to add NF to MSX line if commission > 0 (for HK only)
             If gstrAgcyCountryCode = "HK" Then
                 If .NettFare > 0 And .CommissionAmt > 0 And (.CommissionAmt - .MerchFee) > 0 Then
                    strSecLine = strSecLine + "+DI.FT-MSX/NF" + CStr(.NettFare + .CommissionAmt - .MerchFee)
                 Else
                    strSecLine = strSecLine + "+DI.FT-MSX/NF" + CStr(.NettFare)
                 End If
             Else
                 'Preethi - V1.2.1 20101011 - CR21 - Nett Fare Mark Up
                     strSecLine = strSecLine + "+DI.FT-MSX/NF" + CStr(.PublishedFare)
             End If
             
             
            strMSLine = "/PO" & IIf(Len(.TicketNumber) > 3, Right(.TicketNumber, intLength) & IIf(.ConjunctTicket <> "", "-" & .ConjunctTicket, ""), strZero)
                '& "/PO" & IIf(Len(.TicketNumber) > 3, .TicketNumber, "0000000000000") '"/PO" & IIf(.TicketNumber = "", .EONumber, .TicketNumber) '.EONumber 02062005
                
        Else
            If productType = "BT" And (strProduct = "00" Or strProduct = "01") Then
                'JiYong – V1.2.6 20111003 – CR99 - Add Option for Fare Type in EO
                'Additional logic for CR99 - Set the commission field to be zero if marked up net fare is selected
                strSecLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & Format(IIf(UCase(gstrAgcyCountryCode) = "HK", CStr(.SellPrice - .EOTaxTotal + .Discount), CStr(.PublishedFare)), gstrAgcyCurrFormat) _
                    & "/SF" & Format(IIf(UCase(gstrAgcyCountryCode) = "HK", CStr(.SellPrice - .EOTaxTotal + .Discount), .SF + .MerchFee), gstrAgcyCurrFormat) _
                    & "/C" & IIf(.FareType = 2, Format(0, "0.00"), Format(IIf(UCase(gstrAgcyCountryCode) = "HK", CStr(.CommissionAmt), CStr(.CommissionAmt - .MerchFee)), "0.00")) _
                    & getSegmentSelected(strProduct, productType)
                    'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
                    'Pass in productType also into getSegmentSelected function
                    'Move the modified code to the top due to the constraint of the _ symbol
                    '& getSegmentSelected(strProduct)
                strSecLine = splitLongMSX(strSecLine)
                'Added by JiYong to add NF to MSX line if commission > 0 (for HK only)
                If gstrAgcyCountryCode = "HK" Then
                'preethi – V1.2.6 20110905 – CR99 - Add Option for Fare Type in EO
                   'If .NettFare > 0 And .CommissionAmt > 0 And (.CommissionAmt - .MerchFee) > 0 Then
                       'strSecLine = strSecLine + "+DI.FT-MSX/NF" + CStr(.NettFare + .CommissionAmt - .MerchFee)
                   'Else
                   If .FareType = 2 Then
                       strSecLine = strSecLine + "+DI.FT-MSX/NF" + CStr(.NettFare)
                   End If
                Else
                   'Preethi - V1.2.1 20101011 - CR21 - Nett Fare Mark Up
                   If .FareType = 2 Then
                       strSecLine = strSecLine + "+DI.FT-MSX/NF" + CStr(.SF)
                   End If
                End If
                
                
            ElseIf strProduct = "70" Then
                strMSLine = "/C" & Format(CStr(.CommissionAmt), gstrAgcyCurrFormat)
            ElseIf strProduct = "35" Then
            
                strMSLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
                    & "/SF" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
                    & "/C" & Format(CStr(.CommissionAmt), gstrAgcyCurrFormat) _
                    & getSegmentSelected(strProduct, productType)
                    'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
                    'Pass in productType also into getSegmentSelected function
                    'Move the modified code to the top due to the constraint of the _ symbol
                    '& getSegmentSelected(strProduct)
                       
                    strSecLine = splitLongMSX(strSecLine)
            Else
                 strMSLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
                    & "/SF" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
                    & "/C" & Format(CStr(.CommissionAmt), gstrAgcyCurrFormat)
            End If
                '& "/PO" & .EONumber '02062005
            ''02062005
            'If productType = "MS" And strProduct = "35" Then
            '   If frmOSMisc.optConsol = True Then
            '      strMSLine = strMSLine & "/PO" & .TicketNumber
            '   End If
            'ElseIf productType = "CT" And (strProduct = "35" Or strProduct = "50") Then
            '     strMSLine = strMSLine & "/PO" & .TicketNumber
            'End If
        End If
'Else
'        If productType = "MS" And strProduct = "50" Then
'            strMSLine = "/S" & "-" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
'                & "/SF" & "-" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
'                & "/C" & "-" & Format(CStr(.SellPrice), gstrAgcyCurrFormat)
            '02062005
'            If frmOSMisc.optBSPConsol = True Then
'               strMSLine = strMSLine & "/PO" & .TicketNumber
'            End If
        '02062005
'        ElseIf productType = "CT" And strProduct = "02" Then
'            strMSLine = "/S" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
'                & "/SF" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
'                & "/C" & Format(CStr(.CommissionAmt), gstrAgcyCurrFormat) _
'                & "/PO" & IIf(.TicketNumber = "", .EONumber, .TicketNumber)
'        Else
'            strMSLine = "/S" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
'                & "/SF" & Format(CStr(.SellPrice), gstrAgcyCurrFormat) _
'                & "/C" & Format(CStr(.CommissionAmt), gstrAgcyCurrFormat)
            '02062005
'            If productType = "MS" And strProduct = "35" Then
'               If frmOSMisc.optBSPConsol = True Then
'                  strMSLine = strMSLine & "/PO" & .TicketNumber
'               End If
'            End If
'        End If
'End If
    
    If .Tax(1).Amount > 0 Then
        Select Case .Tax(1).Code
            Case "GST"
                strTemp = "/G" & Format(CStr(.Tax(1).Amount), gstrAgcyCurrFormat)
                If .Tax(1).Amount <> 0 Then
                   strGST = Format(CStr(.Tax(1).Amount), gstrAgcyCurrFormat)
                End If
            Case "VAT"
                ' add code to process VAT (when needed)
            
            Case Else
                'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
                'Set to two decimal point if web fare is selected (For Hong Kong only)
                If gstrAgcyCountryCode = "HK" And .WebFareApplied = True Then
                    strTemp = "/TX" & Format(CStr(.Tax(1).Amount), gstrHKWebFareDecimalPoint) & .Tax(1).Code _
                        & IIf(.Tax(2).Amount > 0, Format(CStr(.Tax(2).Amount), gstrHKWebFareDecimalPoint) & .Tax(2).Code, "")
                Else
                    strTemp = "/TX" & Format(CStr(.Tax(1).Amount), gstrAgcyCurrFormat) & .Tax(1).Code _
                        & IIf(.Tax(2).Amount > 0, Format(CStr(.Tax(2).Amount), gstrAgcyCurrFormat) & .Tax(2).Code, "")
                End If
        End Select
    End If
    
    strMSLine = strMSLine & strTemp
    
    
    
    strTemp = ""
    strFOP = Split(objEO.FOP, "/")
    'Added on 25/08/04: Check if CC payment
    bolIsCC = False
    If strFOP(0) = "CC" Or strFOP(0) = "CX" Then
        'If strFOP(1) = "DC" And Left(strFOP(2), 7) = "3644033" Then
         'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration
        If IsTMPCard(strFOP(1), strFOP(2)) Then
            bolIsCC = False
        Else
            bolIsCC = True
        End If
    End If
    
    If bolIsCC Then
    Select Case strFOP(0)
        Case "CX"
            Select Case strFOP(1)
                Case "AX"
                    strFOPTemp = strFOP(0) & "2"
                Case "DC"
                    strFOPTemp = strFOP(0) & "3"
                Case "VI", "CA"
                    strFOPTemp = strFOP(0) & "4"
                Case "TP"
                    strFOPTemp = strFOP(0) & "5"
            End Select
            'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
            'Set to two decimal point if web fare is selected (For Hong Kong only)
            If gstrAgcyCountryCode = "HK" And .WebFareApplied = True Then
                strTemp = "/F" & strFOPTemp & "/CCN" & strFOP(1) & strFOP(2) & "EXP" & strFOP(3) _
                        & "/D" & Format(CStr(.SellPrice + IIf(.Tax(1).Code = "GST", .Tax(1).Amount, 0)), gstrHKWebFareDecimalPoint)
            Else
                strTemp = "/F" & strFOPTemp & "/CCN" & strFOP(1) & strFOP(2) & "EXP" & strFOP(3) _
                        & "/D" & Format(CStr(.SellPrice + IIf(.Tax(1).Code = "GST", .Tax(1).Amount, 0)), gstrAgcyCurrFormat)
            End If
            'strTemp = "/F" & strFOP(0) & "/CCN" & strFOP(1) & strFOP(2) _
            '       & "/D" & Format(CStr(.SellPrice), gstrAgcyCurrFormat)

        Case "CC"
            'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
            'Set to two decimal point if web fare is selected (For Hong Kong only)
            If gstrAgcyCountryCode = "HK" And .WebFareApplied = True Then
                strTemp = "/FCC/CCN" & strFOP(1) & strFOP(2) & "EXP" & strFOP(3) _
                        & "/D" & Format(CStr(.SellPrice + IIf(.Tax(1).Code = "GST", .Tax(1).Amount, 0)), gstrHKWebFareDecimalPoint)
            Else
                strTemp = "/FCC/CCN" & strFOP(1) & strFOP(2) & "EXP" & strFOP(3) _
                        & "/D" & Format(CStr(.SellPrice + IIf(.Tax(1).Code = "GST", .Tax(1).Amount, 0)), gstrAgcyCurrFormat)
            End If
            'strTemp = "/FCC/CCN" & strFOP(1) & strFOP(2) _
            '        & "/D" & Format(CStr(.SellPrice), gstrAgcyCurrFormat)
       End Select
    Else
            strTemp = strTemp & "/FS"
    End If

    strMSLine = strMSLine & strTemp
    
    'ADD FF40 for HK
    'If (strProduct = "19" Or strProduct = "16" Or strProduct = "70" Or strProduct = "50") Then
    If (strProduct = "19" Or strProduct = "16" Or strProduct = "70") Then
        strMSLine = strMSLine & "/TCW"
    End If
    If (productType = "CT" Or productType = "BT") Then
        'CS Change EC
        'strMSLine = strMSLine & "/R" & .RF & "/L" & .LF & "/E" & .EC
        'strMSLine = strMSLine & "/R" & .RF & "/L" & .LF & "/EC" & .MS
        If gobjPNR.CompInfo.MI = True Then strMSLine = strMSLine & "/R" & .RF & "/L" & .LF & "/E" & .MS
    End If

    strMIData = ""
'If UCase(gstrAgcyCountryCode) = "HK" Then
        
'    If productType = "CX" Then
'        If .PickUpTime <> CdatDefaultDate Then
'            strMIFF40 = "/FF40-PU-" _
'                        & IIf(frmOSCarTxfr.cmbLocation(0).listindex < 3, frmOSCarTxfr.cmbLocation(0).Text, IIf(frmOSCarTxfr.txtLocation(0).Text = "", frmOSCarTxfr.cmbLocation(0), frmOSCarTxfr.txtLocation(0).Text)) _
'                        & "*DO-" _
'                        & IIf(frmOSCarTxfr.cmbLocation(1).listindex < 3, frmOSCarTxfr.cmbLocation(1).Text, IIf(frmOSCarTxfr.txtLocation(1).Text = "", frmOSCarTxfr.cmbLocation(1), frmOSCarTxfr.txtLocation(1).Text))'

'        ElseIf .ReturnTime <> CdatDefaultDate Then
'            strMIFF40 = "/FF40-PU-" _
'                        & IIf(frmOSCarTxfr.cmbLocation(2).listindex < 3, frmOSCarTxfr.cmbLocation(2).Text, IIf(frmOSCarTxfr.txtLocation(2).Text = "", frmOSCarTxfr.cmbLocation(2), frmOSCarTxfr.txtLocation(2).Text)) _
'                        & "*DO-" _
'                        & IIf(frmOSCarTxfr.cmbLocation(3).listindex < 3, frmOSCarTxfr.cmbLocation(3).Text, IIf(frmOSCarTxfr.txtLocation(3).Text = "", frmOSCarTxfr.cmbLocation(3), frmOSCarTxfr.txtLocation(3).Text))
'        End If
'    ElseIf productType = "TR" Then
'        strMIFF40 = "/FF40-" & Trim(Replace(.DescriptionLine(3), Chr(9), ""))
'       End If
'    End If
'    If productType = "VI" Then
'         strMIFF40 = "/FF40-" & .VisaCountry & " " & .VisaEntries & " " & .VisaProcess
'    End If
    
 strMIData = ""

        
    'Added on 2303: Amex statement FF40/41
        
        
    If productType = "CX" Then
        If .PickUpTime <> CdatDefaultDate Then
            strMIFF40 = "/FF40-" _
                        & IIf(frmOSCarTxfr.cmbLocation(0).listindex < 3, frmOSCarTxfr.cmbLocation(0).Text, IIf(frmOSCarTxfr.txtLocation(0).Text = "", frmOSCarTxfr.cmbLocation(0), frmOSCarTxfr.txtLocation(0).Text)) _
                        & "-" _
                        & IIf(frmOSCarTxfr.cmbLocation(1).listindex < 3, frmOSCarTxfr.cmbLocation(1).Text, IIf(frmOSCarTxfr.txtLocation(1).Text = "", frmOSCarTxfr.cmbLocation(1), frmOSCarTxfr.txtLocation(1).Text))
            strMIFF40 = Left(strMIFF40, 26)
            strMIFF40 = strMIFF40 & "/FF41-" & Format(frmOSCarTxfr.dtpPUDateTime, "ddmmmyy hhmm")
            
        'ElseIf .ReturnTime <> CdatDefaultDate Then
        '    strMIFF40 = "/FF40" _
        '                & IIf(frmOSCarTxfr.cmbLocation(2).listindex < 3, frmOSCarTxfr.cmbLocation(2).Text, IIf(frmOSCarTxfr.txtLocation(2).Text = "", frmOSCarTxfr.cmbLocation(2), frmOSCarTxfr.txtLocation(2).Text)) _
        '                & " TO " _
        '                & IIf(frmOSCarTxfr.cmbLocation(3).listindex < 3, frmOSCarTxfr.cmbLocation(3).Text, IIf(frmOSCarTxfr.txtLocation(3).Text = "", frmOSCarTxfr.cmbLocation(3), frmOSCarTxfr.txtLocation(3).Text))
        '
        '    strMIFF40 = strMIFF40 & "/FF41 " & Format(frmOSCarTxfr.dtpRtnDateTime, "ddmmmyy hhmm")
        Else
            'If no pick up time, take return time
            If .ReturnTime <> CdatDefaultDate Then
                strMIFF40 = "/FF40-" _
                            & IIf(frmOSCarTxfr.cmbLocation(2).listindex < 3, frmOSCarTxfr.cmbLocation(2).Text, IIf(frmOSCarTxfr.txtLocation(2).Text = "", frmOSCarTxfr.cmbLocation(2), frmOSCarTxfr.txtLocation(2).Text)) _
                            & "-" _
                            & IIf(frmOSCarTxfr.cmbLocation(3).listindex < 3, frmOSCarTxfr.cmbLocation(3).Text, IIf(frmOSCarTxfr.txtLocation(3).Text = "", frmOSCarTxfr.cmbLocation(3), frmOSCarTxfr.txtLocation(3).Text))
                strMIFF40 = Left(strMIFF40, 26)
                strMIFF40 = strMIFF40 & "/FF41-" & Format(frmOSCarTxfr.dtpRtnDateTime, "ddmmmyy hhmm")
            End If
        End If
        
        
 ElseIf productType = "TR" Then
        'strMIFF40 = "/FF40-" & Trim(Replace(.DescriptionLine(3), Chr(9), ""))
        'added on 191207: TP Card statement Request
        If Trim(frmOSOthTkt.txtFrom(0)) <> "" And frmOSOthTkt.txtFrom(0).Visible = True Then
        
            strMIFF40 = "/FF40-" & frmOSOthTkt.txtFrom(0) & "-" & frmOSOthTkt.txtTo(0)
            
            If Trim(frmOSOthTkt.txtFrom(1)) <> "" Then
                strMIFF40 = strMIFF40 & " " & frmOSOthTkt.txtFrom(1) & "-" & frmOSOthTkt.txtTo(1)
            End If
            strMIFF40 = Left(strMIFF40, 26)
            
            strMIFF40 = strMIFF40 & "/FF41-" & Format(frmOSOthTkt.dtpDepDateTime, "DDMMMYY")
            If Trim(frmOSOthTkt.txtFrom(1)) <> "" Then
                strMIFF40 = strMIFF40 & "-" & Format(frmOSOthTkt.dtpRtnDateTime, "DDMMMYY")
            End If
           
        ElseIf frmOSOthTkt.txtTo(0).Visible = True Then
            
            strMIFF40 = "/FF40-" & frmOSOthTkt.txtTo(0) & "/FF41-" & Format(frmOSOthTkt.dtpDepDateTime, "DDMMMYY")
            If Trim(frmOSOthTkt.txtRtnRoute(0)) <> "" Then
                strMIFF40 = strMIFF40 & "-" & Format(frmOSOthTkt.dtpRtnDateTime, "DDMMMYY")
            End If
        End If
   
    ElseIf productType = "VI" Then
    
        ' FF40/[country] FF41/[Business/tourist/resident/other] [SGL/MUL] [URG/NOR]
         
         strMIFF40 = "/FF40-" & Left(Trim(.VisaCountry), 20) & "/FF41-"
         strMIFF40 = strMIFF40 & .VisaType & " "
         If UCase(.VisaEntries) = "SINGLE" Then strMIFF40 = strMIFF40 & "SGL" & " "
         If UCase(.VisaEntries) = "DOUBLE" Then strMIFF40 = strMIFF40 & "DBL" & " "
         If UCase(.VisaEntries) = "MULTIPLE" Then strMIFF40 = strMIFF40 & "MUL" & " "
         If UCase(.VisaProcess) = "NORMAL" Then strMIFF40 = strMIFF40 & "NOR"
         If UCase(.VisaProcess) = "EXPRESS" Then strMIFF40 = strMIFF40 & "URG"
       
    ElseIf productType = "MS" Then
        If Trim(frmOSMisc.txtBTADescription) <> "" Then
            strMIFF40 = "/FF40-" & frmOSMisc.txtBTADescription & "/FF41-" & Format(frmOSMisc.dtpDate, "ddmmmyy")
        End If
    
    End If
    If freefields <> "" Then
        strFF = Split(freefields, "/")
        For lngC = LBound(strFF) To UBound(strFF)
            'strMIData = strMIData & "/FF" & strFF(lngC)
            If strFF(lngC) <> "" Then
               strTemp = Mid(strFF(lngC), 1, InStr(1, strFF(lngC), "-") - 1)
               strMIData = strMIData & IIf(IsNumeric(strTemp), "/FF", "/") & strFF(lngC)
            End If
        Next
    End If
    
    'CS Transaction Fee, FF30, FF41
    strFFTrans = ""
    strRS = ""
    'strFF41 = ""
    If gobjPNR.CompInfo.MI = True Then
        If (productType = "CT" Or productType = "BT") Then
            strRS = "/FF30-" & .rs
           'strFF41 = "/FF41-" & .FF41
           If Val(frmOSAirTkt.txtTrxnFee) <> 0 Then
              strFFTrans = "/FF31-Y" & "/FF32-" & Format(frmOSAirTkt.txtTrxnFee, gstrAgcyCurrFormat)
           Else
              strFFTrans = "/FF31-N"
           End If
        End If
    End If
    'CS SBT Indicator
    
    If gobjPNR.CompInfo.MI = True Then
        
    If productType <> "BT" Then
           strFFSBT = "/FF34-AB"
           'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
           If productType = "CT" And .WebFareApplied = True Then
              strFFSBT = strFFSBT & "/FF35-AGW"
              strFFSBT = strFFSBT & "/FF36-G"
           Else
              strFFSBT = strFFSBT & "/FF35-OTH"
              strFFSBT = strFFSBT & "/FF36-M"
           End If
    Else
                '---------
            If .BookingAction <> "" Then
               Select Case .BookingAction
                  Case "AB - Agent Booked"
                     strFFSBT = "/FF34-AB"
                  Case "EB - Self Booked"
                     strFFSBT = "/FF34-EB"
                  Case "AA - Air Modified"
                     strFFSBT = "/FF34-AA"
                  Case "AM - Multiple Modification"
                     strFFSBT = "/FF34-AM"
               End Select
            End If
            
            'CS Add Booking Tool
            If .BookingTool <> "" Then
               'Preethi - V1.2.6 20110907 - CR 90 - Change OBT Tool Code in FF35
               strBookingTool = getFF35OBT(Mid(.BookingAction, 1, 2), .BookingTool)
               strFFSBT = strFFSBT & "/FF35-" & strBookingTool
            Else
               strFFSBT = strFFSBT & "/FF35-GAL"
            End If
            
             If .BookingTool <> "" Then
               strFFSBT = strFFSBT & "/FF36-S"
            Else
               strFFSBT = strFFSBT & "/FF36-G"
            End If
            '---------
                
                   'strFFSBT = strFFSBT & "/FF35-GAL"
                   'strFFSBT = strFFSBT & "/FF36-G"
            End If
    End If
    
      'CS SBT Indicator
      'strMSLine = strMSLine & strMIData & strMIFF40 & strFF41 & strRS & strFFTrans & strFFSBT
      strMSLine = strMSLine & strMIData & strMIFF40 & strRS & strFFTrans & strFFSBT
      If productType <> "BT" And productType <> "CT" Then
         strMSLine = strMSLine & "/FF47-CWT"
      End If
      
      'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
      If productType = "CT" And .WebFareApplied = True Then
         'Generate plating carrier into DI line if webfare is applied
         strMSLine = strMSLine & "/AC" & .PlatingCarrier
      End If
      
      
  '    If productType = "MS" And (strProduct = "35" Or strProduct = "50") Then
  '      If frmOSMisc.optBSPConsol = True Then
  '          strMSLine = strMSLine & "/TK" & IIf(.TicketNumber = "0000", Format(.TicketNumber, "0000000000" & "/"), fConvertTkTNo(.TicketNumber)) & "PX" & .PassengerID
  '      Else
  '          strMSLine = strMSLine & "/TKFF" & IIf(.TicketNumber = "0000", Format(.TicketNumber, "0000000000"), Format(.TicketNumber, "00")) & "/PX" & .PassengerID
  '      End If
      
  '    Else
  '      strMSLine = strMSLine & IIf(.TicketNumber = "", "", "/TK" & Format(.TicketNumber, "0000000000")) & "/PX" & .PassengerID
  '    End If
'This routine will make sure that entry does not exceed max char and will add the least amount of lines to PNR based on Max length of 45 char after FT-
'modified on 6/7/2005: remove logic to add DI line on the top due to VFF which always required to stay on top
strTemp = "DI.FT-MS" & strMSFirstLine & strSecLine
'strTemp = "DI./0" & "+DI.FT-MS" & strMSFirstLine

lngC = 0
Do Until Len(strMSLine) = 0
    If Len(strMSLine) <= 42 Then
        lngLen = Len(strMSLine)
    Else
        lngLen = InstrLast(Left(strMSLine, 42), "/") - 1
    End If
    strTemp = strTemp & "+DI.FT-MSX" & Left(strMSLine, lngLen)
    'strTemp = strTemp & IIf(lngC = 0, "+DI.FT-MS", "+DI.FT-MSX") & Left(strMSLine, lngLen)
    strMSLine = Mid(strMSLine, lngLen + 1)
    lngC = lngC + 1
Loop
'preethi – V1.2.6 20110905 – CR98 - Reissue Ticket Box in EO
If productType = "BT" Then
   If .TktNumber <> "" Then
      strTemp = strTemp & "+DI.FT-MSX/EX" & .TktNumber
   End If
End If
strTemp = strTemp & "+DI.FT-MSX/FF " & frmOthSvcs.datProducts.Recordset![Description]

gobjHost.terminalEntry strTemp



Select Case productType
    Case "HL", "BT", "CT"
        'ignore these
    Case "TR"
        strTemp = "0TURZZBK1" & gstrAgcyCityCode & Format(.ServiceDate, "ddmmm") & "-" & .DescriptionLine(1)
        gobjHost.terminalEntry UCase(strTemp)
        
        strTemp = ""
        For lngC = 2 To .DescriptionLinesCount
            strTemp = "RT.T/" & Format(.ServiceDate, "ddmmm") & "*" & .DescriptionLine(lngC)
            gobjHost.terminalEntry UCase(strTemp)
 
        Next
    'add on 1/4/2005
    Case "CX"
        For lngC = 2 To .DescriptionLinesCount
            strTemp = "0TURZZBK1" & gstrAgcyCityCode & .DescriptionLine(lngC)
            
            strTemp = gobjHost.terminalEntry(UCase(strTemp))
        Next
    
    
    
    Case Else
        
        If UCase(gstrAgcyCountryCode) = "SG" And (strProduct = "50" Or strProduct = "70") Then
           strTemp = ""
        Else
           strTemp = "0TURZZBK1" & gstrAgcyCityCode & Format(.ServiceDate, "ddmmm") & "-"      '& .DescriptionLine(1)
           For lngC = 1 To .DescriptionLinesCount
              strTemp2 = strTemp2 & "*" & .DescriptionLine(lngC)
           Next
        End If
        'added on 10/12: split TUR line for invoice display
        
        lngC = 0
        Do Until Len(strTemp2) = 0
           If Len(strTemp2) <= 42 Then
              lngLen = Len(strTemp2)
           Else
              lngLen = InstrLast(Left(strTemp2, 42), " ") - 1
              If lngLen <= 0 Then lngLen = 42
           End If
         
           gobjHost.terminalEntry UCase(strTemp & Left(strTemp2, lngLen))
           strTemp2 = Mid(strTemp2, lngLen + 1)
           lngC = lngC + 1
        Loop
End Select
'Add RD line
'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
'Do not generate RD/RP line if web fare is selected (For Hong Kong only)
If Not (UCase(gstrAgcyCountryCode) = "HK" And .WebFareApplied = True) Then

    If .Discount > 0 Then
        If (productType = "CT" Or productType = "BT") Then
             strTemp = "RD.T/" & Format(.ServiceDate, "ddmmm") & "*" & "AIR TICKET " & Format(.SellPrice + IIf(UCase(gstrAgcyCountryCode) = "HK", .Discount, 0) - .EOTaxTotal, gstrAgcyCurrFormat) & _
                        " TAX " & Format(.EOTaxTotal, gstrAgcyCurrFormat) & "*" & _
                        Format(.SellPrice + IIf(UCase(gstrAgcyCountryCode) = "HK", .Discount, 0), gstrAgcyCurrFormat)
        Else
             strTemp = "RD.T/" & Format(.ServiceDate, "ddmmm") & "*" & .DescriptionLine(1) & "*" & Format(.SellPrice + .Discount, gstrAgcyCurrFormat)
        End If
    Else
        If (productType = "CT" Or productType = "BT") Then
             strTemp = "RD.T/" & Format(.ServiceDate, "ddmmm") & "*" & "AIR TICKET " & Format(.SellPrice - .EOTaxTotal, gstrAgcyCurrFormat) & _
                        " TAX " & Format(.EOTaxTotal, gstrAgcyCurrFormat) & "*" & _
                        Format(.SellPrice, gstrAgcyCurrFormat)
        Else
            If UCase(gstrAgcyCountryCode) = "HK" Then
               strTemp = "RD.T/" & Format(.ServiceDate, "ddmmm") & "*" & .DescriptionLine(1) & "*" & Format(.SellPrice, gstrAgcyCurrFormat)
            Else
               If strProduct = "50" Or strProduct = "70" Then
                  strTemp = ""
               Else
                  strTemp = "RD.T/" & Format(.ServiceDate, "ddmmm") & "*" & .DescriptionLine(1) & "*" & Format(.SellPrice, gstrAgcyCurrFormat)
               End If
            End If
        End If
    End If

    gobjHost.terminalEntry strTemp

    If strGST <> "" Then
       strTemp = "RD.T/" & Format(.ServiceDate, "ddmmm") & "*" & CStr(frmOthSvcs.datProducts.Recordset![GST]) & " PERCENT GST*" & strGST
       gobjHost.terminalEntry strTemp
    End If
    'Added on 25/08/04: CC Paid lines
    If bolIsCC Then
    
        strTemp = "RP.T/" & Format(.ServiceDate, "ddmmm") & "*" & strFOP(1) & "XXXXXXXXXXX" & Right(strFOP(2), 4) & "*" & Format(.SellPrice, gstrAgcyCurrFormat)
        gobjHost.terminalEntry strTemp
    
        If strGST <> "" Then
           strTemp = "RP.T/" & Format(.ServiceDate, "ddmmm") & "*" & strFOP(1) & "XXXXXXXXXXX" & Right(strFOP(2), 4) & "*" & strGST
           gobjHost.terminalEntry strTemp
        End If
        
    End If

    If .Discount > 0 Then
       If UCase(gstrAgcyCountryCode) = "HK" Then
            strTemp = "RP.T/" & Format(.ServiceDate, "ddmmm") & "*" & "CLIENT DISCOUNT" & "*" & Format(.Discount, gstrAgcyCurrFormat)
            gobjHost.terminalEntry strTemp
       End If
    End If

End If

Dim strVendorNum As String
'modified on 050906
If .FF81 <> "" And gobjPNR.CompInfo.MI = False Then
    strMIOther = "/FF81-" & .FF81
End If

'EO
'If (ProductType = "CT" Or ProductType = "BT") And frmOSAirTkt.txtTrxnFee <> "0" Then
If (productType = "CT" Or productType = "BT") Then

'If UCase(gstrAgcyCountryCode) = "HK" Then
'    strVendorNum = VendorNum("35", "CWT")
'Else
    strVendorNum = VendorNum(strTrxnFeeCode, IIf(frmOSAirTkt.chkTFNRCC.value = False, "CWT", "MER"))
'End If

  If Val(frmOSAirTkt.txtTrxnFee) <> 0 Then
  strMSFirstLine = "/PC35" _
        & "/" & strVendorNum _
        & getTicketNo(productType, 35) & "PX" & .PassengerID
    strSecLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat) _
        & "/SF" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat) _
        & "/C" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat) _
        & getSegmentSelected(strProduct, productType)
        'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
        'Pass in productType also into getSegmentSelected function
        'Move the modified code to the top due to the constraint of the _ symbol
        '& getSegmentSelected(strProduct)
    strSecLine = splitLongMSX(strSecLine)
    
        ''02062005
        'strMSLine = strMSLine & "/PO" & .TicketNumber
    strTemp = ""
    'strFOP = Split(objEO.FOP, "/")
    
    'Select Case strFOP(0)
    '    Case "CX"
    '        strTemp = "/F" & strFOPTemp & "/CCN" & strFOP(1) & strFOP(2) _
    '                & "/D" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat)
    '    Case "CC"
    '        strTemp = "/FCC/CCN" & strFOP(1) & strFOP(2) _
    '                & "/D" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat)
    'End Select
    
    'Else
    '        strTemp = strTemp & "/FS"
    'End If
    'End Select
    
     'modified on 30032006
     If bolIsCC Then
     'JY – V1.2.3 20110418 – IR10 - Wrong FOP code is generated in DI line if CC is selected as FOP
    'If UCase(gstrAgcyCountryCode) = "SG" Or productType = "BT" Then
    If UCase(gstrAgcyCountryCode) = "SG" Or productType = "BT" Or (productType = "CT" And strFOP(0) = "CC") Then
            If .TFNRCC = False Then
            
                Select Case strFOP(1)
                    Case "AX"
                        strTFFOPTemp = "CX2"
                    Case "DC"
                        strTFFOPTemp = "CX3"
                    Case "VI", "CA"
                        strTFFOPTemp = "CX4"
                    Case "TP"
                        strTFFOPTemp = "CX5"
                End Select
            
            End If
            'If UCase(gstrAgcyCountryCode) = "HK" And .TFNRCC = True Then
            '    strTemp = "/FS"
            'Else
                strTemp = "/F" & IIf(.TFNRCC = True, "CCN", strTFFOPTemp) & "/CCN" & strFOP(1) & strFOP(2) _
                          & "/D" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat)
            
            'End If

    Else
            
            Select Case strFOP(0)
                Case "CX"
                    strTemp = "/F" & strFOPTemp & "/CCN" & strFOP(1) & strFOP(2) & "EXP" & strFOP(3) _
                            & "/D" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat)
                Case "CC"
                    strTemp = "/FCC/CCN" & strFOP(1) & strFOP(2) & "EXP" & strFOP(3) _
                            & "/D" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat)
            End Select
          
    End If
    Else
            strTemp = strTemp & "/FS"
    End If
    
    
    'If ProductType = "BT" Then
        strMSLine = strMSLine & strTemp & strMIOther & getMSLineforMI()
        
        'CS Add FF34,35,36
        strMSLine = strMSLine & strFFSBT & "/FF47-CWT"
    'Else
    '    strMSLine = strMSLine & strTemp & strMIData & IIf(.TicketNumber = "", "", "/TK" & Format(.TicketNumber, 0)) & "/PX" & .PassengerID
    'End If
    
    'strTemp = "DI./0"
    'strTemp = strTemp & "+DI.FT-MS" & strMSFirstLine
    'modified on 6/7/2005: remove logic to add DI line on the top due to VFF which always required to stay on top
    
    'JY – V1.2.3 20110419 – CR62 - Generate AC Code for LCC Booking
    If productType = "CT" And .WebFareApplied = True Then
       'Generate plating carrier into DI line if webfare is applied
        strMSLine = strMSLine & "/AC" & .PlatingCarrier
    End If

    strTemp = "DI.FT-MS" & strMSFirstLine & strSecLine
    
    lngC = 0
    Do Until Len(strMSLine) = 0
       If Len(strMSLine) <= 42 Then
          lngLen = Len(strMSLine)
       Else
          lngLen = InstrLast(Left(strMSLine, 42), "/") - 1
       End If
       strTemp = strTemp & "+DI.FT-MSX" & Left(strMSLine, lngLen)
       'strTemp = strTemp & IIf(lngC = 0, "+DI.FT-MS", "+DI.FT-MSX") & Left(strMSLine, lngLen)
       strMSLine = Mid(strMSLine, lngLen + 1)
       lngC = lngC + 1
    Loop
    strTemp = strTemp & "+DI.FT-MSX/FF " & UCase(getProductDesc("35"))
    
    gobjHost.terminalEntry strTemp
    'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
    'Do not generate RD/RP line if web fare is selected (For Hong Kong only)
    If Not (UCase(gstrAgcyCountryCode) = "HK" And .WebFareApplied = True) Then
        strTemp = "RD.T/" & Format(.ServiceDate, "ddmmm") & "*TRANSACTION FEE*" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat)
        gobjHost.terminalEntry strTemp
        If bolIsCC Then
            strTemp = "RP.T/" & Format(.ServiceDate, "ddmmm") & "*" & strFOP(1) & "XXXXXXXXXXX" & Right(strFOP(2), 4) & "*" & Format(CStr(IIf(frmOSAirTkt.txtTrxnFee = "", 0, frmOSAirTkt.txtTrxnFee)), gstrAgcyCurrFormat)
            gobjHost.terminalEntry strTemp
        End If
    End If
  End If
'---------------------
'fuel surcharge
Dim strFuelChargeCode As String

Set rs = gdbConn.Execute("Select sortkey from tblProductCodes where productcode = '41'")
If rs.EOF Then
    strFuelChargeCode = ""
Else
    strFuelChargeCode = rs!SortKey & ""
End If
rs.Close
Set rs = Nothing


  If Val(frmOSAirTkt.txtFuelSurcharge) <> 0 Then
  strMSFirstLine = "/PC" & strFuelChargeCode _
        & "/" & VendorNum(strFuelChargeCode, "CWT") _
        & getTicketNo(productType, strFuelChargeCode) & "PX" & .PassengerID
    strMSLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & Format(CStr(IIf(frmOSAirTkt.txtFuelSurcharge = "", 0, frmOSAirTkt.txtFuelSurcharge)), gstrAgcyCurrFormat) _
        & "/SF" & Format(CStr(IIf(frmOSAirTkt.txtFuelSurcharge = "", 0, frmOSAirTkt.txtFuelSurcharge)), gstrAgcyCurrFormat) _
        & "/C" & Format(CStr(IIf(frmOSAirTkt.txtFuelSurcharge = "", 0, frmOSAirTkt.txtFuelSurcharge)), gstrAgcyCurrFormat)
     
    strTemp = ""
   
     If bolIsCC Then
     Dim strFOPFuel As String
     
            Select Case strFOP(1)
                Case "AX"
                    strFOPFuel = "CX2"
                Case "DC"
                    strFOPFuel = "CX3"
                Case "VI", "CA"
                   strFOPFuel = "CX4"
                Case "TP"
                    strFOPFuel = "CX5"
             End Select
            
     
                    strTemp = "/F" & strFOPFuel & "/CCN" & strFOP(1) & strFOP(2) & "EXP" & strFOP(3) _
                            & "/D" & Format(CStr(IIf(frmOSAirTkt.txtFuelSurcharge = "", 0, frmOSAirTkt.txtFuelSurcharge)), gstrAgcyCurrFormat)
     Else
                    strTemp = strTemp & "/FS"
     End If
    
        strMSLine = strMSLine & strTemp & strMIOther & getMSLineforMI()
        strMSLine = strMSLine & strFFSBT & "/FF47-CWT"
        'JY – V1.2.3 20110419 – CR62 - Generate AC Code for LCC Booking
        If productType = "CT" And .WebFareApplied = True Then
           'Generate plating carrier into DI line if webfare is applied
            strMSLine = strMSLine & "/AC" & .PlatingCarrier
        End If
        strTemp = "DI.FT-MS" & strMSFirstLine
    
    lngC = 0
    Do Until Len(strMSLine) = 0
       If Len(strMSLine) <= 42 Then
          lngLen = Len(strMSLine)
       Else
          lngLen = InstrLast(Left(strMSLine, 42), "/") - 1
       End If
       strTemp = strTemp & "+DI.FT-MSX" & Left(strMSLine, lngLen)
       strMSLine = Mid(strMSLine, lngLen + 1)
       lngC = lngC + 1
    Loop
    strTemp = strTemp & "+DI.FT-MSX/FF " & UCase(getProductDesc("41"))
    
    gobjHost.terminalEntry strTemp
    'JY – V1.2.3 20110418 – CR61 - Remove Rounding Logic if LCC web fare is selected
    'Do not generate RD/RP line if web fare is selected (For Hong Kong only)
    If Not (UCase(gstrAgcyCountryCode) = "HK" And .WebFareApplied = True) Then
        strTemp = "RD.T/" & Format(.ServiceDate, "ddmmm") & "*FUEL CHARGE SVC FEE*" & Format(CStr(IIf(frmOSAirTkt.txtFuelSurcharge = "", 0, frmOSAirTkt.txtFuelSurcharge)), gstrAgcyCurrFormat)
        gobjHost.terminalEntry strTemp
        If bolIsCC Then
            strTemp = "RP.T/" & Format(.ServiceDate, "ddmmm") & "*" & strFOP(1) & "XXXXXXXXXXX" & Right(strFOP(2), 4) & "*" & Format(CStr(IIf(frmOSAirTkt.txtFuelSurcharge = "", 0, frmOSAirTkt.txtFuelSurcharge)), gstrAgcyCurrFormat)
            gobjHost.terminalEntry strTemp
        End If
    End If
End If
  




'---------------------
If .Discount > 0 Then

If UCase(gstrAgcyCountryCode) = "HK" Then
    strVendorNum = VendorNum("50", "CWT")
Else
    strVendorNum = VendorNum(strDiscCode, "REBATE")
End If

    curDiscAmt = -Format(CStr(.Discount), gstrAgcyCurrFormat)
        
        strMSFirstLine = "/PC50" _
            & "/" & strVendorNum _
            & getTicketNo(productType, 50) & "PX" & .PassengerID
            
        strSecLine = IIf(UCase(gstrAgcyCountryCode) = "HK", "/A", "/S") & curDiscAmt _
            & "/SF" & curDiscAmt _
            & "/C" & curDiscAmt _
            & getSegmentSelected(strProduct, productType)
            'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
            'Pass in productType also into getSegmentSelected function
            'Move the modified code to the top due to the constraint of the _ symbol
            '& getSegmentSelected(strProduct)
            
        strSecLine = splitLongMSX(strSecLine)
        'strMSLine = strMSLine & "/PO" & .TicketNumber
        strTemp = ""
         'Select Case strFOP(0)
         '   Case "CX"
         '       strTemp = "/F" & strFOPTemp & "/CCN" & strFOP(1) & strFOP(2) _
         '               & "/D" & Format(CStr(.Discount), gstrAgcyCurrFormat)
         '   Case "CC"
         '       strTemp = "/FCC/CCN" & strFOP(1) & strFOP(2) _
         '               & "/D" & Format(CStr(.Discount), gstrAgcyCurrFormat)
         '    Case Else
                strTemp = strTemp & "/FS"
       
       'End Select
        
        
        'If ProductType = "BT" Then
            'strMSLine = strMSLine & strTemp & strMIOther & getMSLineforMI() & strFFSBT & "/TCW" & "/FF47-#CWT"
            strMSLine = strMSLine & strTemp & strMIOther & getMSLineforMI() & strFFSBT & "/FF47-CWT"
        'Else
        '    strMSLine = strMSLine & strTemp & strMIData & IIf(.TicketNumber = "", "", "/TK" & Format(.TicketNumber, 0)) & "/PX" & .PassengerID
        'End If
        
        'strTemp = "DI./0"
        'strTemp = strTemp & "+DI.FT-MS" & strMSFirstLine
        'modified on 6/7/2005: remove logic to add DI line on the top due to VFF which always required to stay on top
        'JY – V1.2.3 20110419 – CR62 - Generate AC Code for LCC Booking
        If productType = "CT" And .WebFareApplied = True Then
           'Generate plating carrier into DI line if webfare is applied
            strMSLine = strMSLine & "/AC" & .PlatingCarrier
        End If
        strTemp = "DI.FT-MS" & strMSFirstLine & strSecLine
        lngC = 0
        Do Until Len(strMSLine) = 0
           If Len(strMSLine) <= 42 Then
              lngLen = Len(strMSLine)
           Else
              lngLen = InstrLast(Left(strMSLine, 42), "/") - 1
           End If
           strTemp = strTemp & "+DI.FT-MSX" & Left(strMSLine, lngLen)
           'strTemp = strTemp & IIf(lngC = 0, "+DI.FT-MS", "+DI.FT-MSX") & Left(strMSLine, lngLen)
           strMSLine = Mid(strMSLine, lngLen + 1)
           lngC = lngC + 1
        Loop
        strTemp = strTemp & "+DI.FT-MSX/FF " & UCase(getProductDesc("50"))
        
        gobjHost.terminalEntry strTemp
    End If
End If


'Added for Insurance
'Added for Insurance
If .ProductCode = "09" Then
    If frmOSMisc.lsvInsPax.ListItems.Count > 0 Then
        strTemp = "NP.II***** INSURANCE INFORMATION ******"
        strTemp = strTemp & "+NP.I*PLAN SELECTED-" & frmOSMisc.cmbInsPlan.Text
        strTemp = strTemp & "+NP.II******* INSURED PERSONS *********"
        For lngC = 1 To frmOSMisc.lsvInsPax.ListItems.Count
            strTemp = strTemp & "+NP.I* " & frmOSMisc.lsvInsPax.ListItems(lngC)
            strTemp = strTemp & "+NP.I*RELATION - " & frmOSMisc.lsvInsPax.ListItems(lngC).SubItems(1) & "/PREMIUM - " & Format(frmOSMisc.lsvInsPax.ListItems(lngC).SubItems(2), "0.00")
        Next
        strTemp = strTemp & "+NP.II***********************************"
        strTemp2 = ""
        For lngC = 0 To frmOSMisc.txtInsAdd.Count - 1
            If frmOSMisc.txtInsAdd(lngC) <> "" Then
                strTemp2 = strTemp2 & IIf(strTemp2 = "", "+NP.I*INSURED ADDRESS -", "+NP.I* ") & frmOSMisc.txtInsAdd(lngC)
            End If
        Next
        strTemp = strTemp & strTemp2
        
        strTemp = strTemp & "+NP.I*GEOGRAPHICAL AREA- " & frmOSMisc.cmbGeoArea.Text
        strTemp = strTemp & "+NP.I*INSURANCE PERIOD/DAYS- " & frmOSMisc.txtInsDays & "/FROM- " & Format(frmOSMisc.dtpInsFromDate, "DDMMM")
        If frmOSMisc.cmbFOPType = "CX" Then
        strTemp = strTemp & "+NP.I*FOP FROM CLIENT- " & frmOSMisc.cmbCCType
        Else
        strTemp = strTemp & "+NP.I*FOP FROM CLIENT- " & frmOSMisc.cmbFOPType
        End If
        strTemp = strTemp & "+NP.I*SELLING PRICE- " & Format(.SellPrice, "0.00")
        strTemp = strTemp & "+NP.I*COST PRICE- " & Format(.Cost, "0.00")
        strTemp = strTemp & "+NP.II***********************************"
        
        gobjHost.terminalEntry strTemp
    End If
End If

If .ProductCode = "00" Then
    If frmOSAirTkt.chkRequestMCO Then
    strTemp = "NP.M* *********** MCO REQUEST ***********"
    
    strTemp = strTemp & "+NP.M*TRAVELLER- "
        For lngC = 1 To frmOSAirTkt.lsvTraveller.ListItems.Count
            strTemp = strTemp & " " & frmOSAirTkt.lsvTraveller.ListItems(lngC)
        Next
        strTemp = strTemp & "+NP.M*RECLOC- " & frmOSAirTkt.txtRecLoc
        strTemp = strTemp & "+NP.M*SERVICE- " & frmOSAirTkt.txtTypeOfService
        strTemp = strTemp & "+NP.M*LOCATION OF ISSUANCE- " & frmOSAirTkt.txtLoc
        strTemp = strTemp & "+NP.M*CONTACT- " & frmOSAirTkt.txtContactPerson
        strTemp = strTemp & "+NP.M*FOP- " & frmOSAirTkt.txtFOP
        strTemp = strTemp & "+NP.M*EQUIV AMT PAID- " & frmOSAirTkt.txtEquiAmt
        strTemp = strTemp & "+NP.M*RATE OF EXCHANGE- " & frmOSAirTkt.txtROE
        strTemp = strTemp & "+NP.M*HEADLINE CURRENCY- " & frmOSAirTkt.txtHeadlineCurrency
        strTemp = strTemp & "+NP.M*TAXES- " & frmOSAirTkt.txtMCOTaxes
        strTemp = strTemp & "+NP.M*ISSUED IN EXCH FOR- " & frmOSAirTkt.txtExchangeFor
        strTemp = strTemp & "+NP.M*IN CONJUNCTION WITH- " & frmOSAirTkt.txtConjunction
        strTemp = strTemp & "+NP.M*IN ORIGINAL FOP- " & frmOSAirTkt.txtOrginalFOP
        strTemp = strTemp & "+NP.M*IN ORIGINAL PLACE OF ISSUE- " & frmOSAirTkt.txtOrginalPOI
        strTemp = strTemp & "+NP.M*REMARKS- "
        For lngC = 0 To frmOSAirTkt.lstRmks.ListCount - 1
            strTemp = strTemp & IIf(lngC = 0, "", "+NP.M*") & frmOSAirTkt.lstRmks.List(lngC)
        Next
        strTemp = strTemp & "+NP.M* **********************************"
        
        gobjHost.terminalEntry strTemp
    End If
End If


For lngC = 1 To .RIRemarkCount
   gobjHost.terminalEntry "RI." & .RIRemark(lngC)
Next




'If productType = "CT" Or productType = "BT" Then
    'mbytFFNum = gobjPNR.FiledFareCount + 1
    
    
    'mbytFFNum = getCurrDispNum() + 1
    
    'strEntry = "DI.FT-LF/*" & mbytFFNum & "/" & .LF _
    '    & "+DI.FT-RF/*" & mbytFFNum & "/" & .RF _
    '    & "+DI.FT-EC/*" & mbytFFNum & "/" & .EC
    'gobjHost.TerminalEntry strEntry
   
'End If
    
   'all FreeFields are stored in MSX line
   'If Trim(.FF8) <> "" Then
   '   strEntry = strEntry & "+DI.FT-FF8/*" & mbytFFNum & "/" & Trim(.FF8)
   'End If
   'If Trim(.FF10) <> "" Then
   '    strEntry = strEntry & "+DI.FT-FF10/*" & mbytFFNum & "/" & Trim(.FF10)
   'End If
   'If Trim(.FF11) <> "" Then
   '   strEntry = strEntry & "+DI.FT-FF11/*" & mbytFFNum & "/" & Trim(.FF11)
   'End If
   'If Trim(.FF19) <> "" Then
   '    strEntry = strEntry & "+DI.FT-FF19/*" & mbytFFNum & "/" & Trim(.FF19)
   'End If
   'If .FF26 <> "" Then
   '   Select Case UCase(.FF26)
   '      Case "ROUND"
   '          strEntry = strEntry & "+DI.FT-FF26/*" & mbytFFNum & "/" & "R"
   '      Case "ONE WAY"
   '          strEntry = strEntry & "+DI.FT-FF26/*" & mbytFFNum & "/" & "O"
   '   End Select
   'End If


'If .ProductCode = "07" Then
'   strSegNum = Trim(Left(.PickUpFlight, 2))
'   If IsNumeric(strSegNum) Then
'      gobjHost.TerminalEntry "RI.S" & strSegNum & "****************CAR TRANSFER***************"
'      gobjHost.TerminalEntry "RI.S" & strSegNum & "*VENDOR NAME   : " & .VendorName
'      gobjHost.TerminalEntry "RI.S" & strSegNum & "*CONTACT NUMBER: " & .ContactNum
'   Else
'      gobjHost.TerminalEntry "RI." & "***************CAR TRANSFER***************"
'      gobjHost.TerminalEntry "RI." & "VENDOR NAME   : " & .VendorName
'      gobjHost.TerminalEntry "RI." & "CONTACT NUMBER: " & .ContactNum
'   End If
'End If

'Added on 13/08/04: add notepad lines for XO
gobjHost.terminalEntry "NP.XO*XO NUMBER " & gobjEO.EONumber & " ISSUED" & _
                       "+NP.XO*FOR SERVICE TYPE: " & frmOthSvcs.dbcProducts.Text & _
                       "+NP.SS*VBIXO"
'Added on 1/2/05 - indicate XO completed by VBI
'gobjHost.TerminalEntry "NP.SS*VBIXO"
If gobjEO.ReplyEmail <> Trim(UCase(GetAgentEmail(gobjPNR.CompInfo.ProfileName, gobjPNR.Agent, gobjPNR.PCCOwner, True, True))) Then
   frmEitin.updateReplyEmail gobjEO.ReplyEmail, "EMO.", "", ""
Else

   frmEitin.updateReplyEmail "", "EMO.", "", ""
End If


strTemp = gobjHost.terminalEntry("R.XO+ER")
gobjHost.terminalEntry "ER"
gobjHost.terminalEntry "ER"
strPNR = gobjPNR.RecLoc
'gobjHost.TerminalEntry "TKPDID"
gobjHost.terminalEntry "*" & strPNR
End With
 
'Added on 14/10/04: add to VBI log table
'Timer

Call pAddToVBILog(gobjPNR.RecLoc, "Other Services", startTime, gSysStartOthSvcsTime, productType, , gSysStartOthSvcsTime)
 
End Sub
'Modified on 21/03/2005
Private Function getProductDesc(prodCode As String) As String
Dim rsProduct As New ADODB.Recordset
'Set rsProduct = gdbTPro.OpenRecordset("select description from tblProductCodes where productCode = '" & prodCode & "'")
Set rsProduct = gdbConn.Execute("select description from tblProductCodes where productCode = '" & prodCode & "'")
If Not rsProduct.EOF Then
    getProductDesc = rsProduct!Description
Else
    getProductDesc = ""
End If
rsProduct.Close
Set rsProduct = Nothing
End Function

Public Function SetEONumber() As Boolean
'Dim rs As Recordset
Dim strNum As String

    If gbolEOAmend = False Then
    'Remove on 24/02/05: Visa ticket no to follow EO number
       'If Not Visa Or UCase(gstrAgcyCountryCode) = "HK" Then
          'Modified on 16/03/05: retrieve EO, TK format from tblDocStruct
          'gobjEO.EONumber = AssignEONum
          SetEONumber = True
          Call AssignEONum
          strNum = gobjEO.EONumber
          
          If strNum = "" Then
            SetEONumber = False
            Exit Function
          End If
          gobjEO.EONumber = genNumFormat("EO", strNum)
          
          If gobjEO.TicketNumber = "0000" Then
            'Modified on 16/03/05: retrieve EO, TK format from tblDocStruct
             'tktlen
             'gobjEO.TicketNumber = frmOthSvcs.datProducts.Recordset![TktPrefix] & "0" & Right(gobjEO.EONumber, Len(gobjEO.EONumber) - Len(frmOthSvcs.datProducts.Recordset![TktPrefix] & Format(Now, "yymm")))
             gobjEO.TicketNumber = genNumFormat("TK", strNum)
          End If
       'Else
        'txtEONum = gobjEO.EONumber
       'End If
    'Else
       'txtEONum = gobjEO.EONumber
    End If
End Function
Public Sub SetRecLoc()
    Dim lngC As Integer
    Dim strPaxName As String
    'If Not gobjPNR Is Nothing Then Set gobjPNR = Nothing
    'Set gobjPNR = New CWT_GalileoPNR.PNR
    'If gobjLog.LogOpen Then gobjPNR.OpenLog gobjLog
    Set gobjPNR = New CWT_GalileoPNR3.PNR
    gobjPNR.loadPNR
    gobjEO.PNRRecLoc = gobjPNR.RecLoc
    For lngC = 1 To gobjPNR.PassengerCount
        strPaxName = strPaxName & gobjPNR.PassengerName(lngC).LastName & "/" & gobjPNR.PassengerName(lngC).FirstName & IIf(lngC = gobjPNR.PassengerCount, "", vbCrLf)
    Next
    gobjEO.PaxName = strPaxName
End Sub

Public Sub pTktQueue()
    
    
    frmTktQueue.Show
    Do
         DoEvents
    Loop Until isLoaded("frmTktQueue") = False
    
End Sub
Public Function AssignEONum(Optional Count As Integer = 0) As String
Dim rsEO As New ADODB.Recordset
Dim strEONO As String
Dim strMsg As String

On Error GoTo SQLError:
'Set rsEO = gdbConn.Execute("SELECT * FROM tblCount")
'With rsEO
'    gdbConn.Execute ("update tblCount set ExchangeID ='" & ![ExchangeID] + 1 & "'")
    'Modified on 16/03/05: retrieve EO, TK format from tblDocStruct
    'AssignEONum = Format(![ExchangeID], "000000")
'    AssignEONum = ![ExchangeID]
'    .Close
'End With
    'Modified on 20/1/2005: change EO Number to prefix+yymm+running#
    'AssignEONum = frmOthSvcs.datProducts.Recordset![TktPrefix] & Format(Now, "yymm") & AssignEONum
'Set rsEO = Nothing

'Modified on 27/4/05: Roll back tblcount if insert tblExchangeOrder is unsuccessful

 gdbConn.BeginTrans
 gbolBeginTrans = True
 
 'Set rsEO = gdbConn.Execute("SELECT * FROM tblCount")
Set rsEO = New ADODB.Recordset
rsEO.Open "SELECT * FROM tblCount", gdbConn, adOpenDynamic, adLockPessimistic
'adLockOptimistic

With rsEO
    '.MoveFirst
    strEONO = ![ExchangeID] + 1
    .Close
End With


    strSQL = "update tblCount set ExchangeID ='" & strEONO & "'"
    gdbConn.Execute strSQL
    'rsEO.Open strSql, gdbConn, adOpenDynamic, adLockOptimistic
    
    gobjEO.EONumber = strEONO

   
 Set rsEO = Nothing
Exit Function
SQLError:
    
    
    Set rsEO = Nothing
    
    gdbConn.RollbackTrans

    If Err.Number = -2147217871 Then
        gobjLog.LineTextToLog Err.Source & Err.Number & Err.Description
        Count = Count + 1
        If Count < 4 Then
            Call AssignEONum(Count)
        Else
            'MsgBox Err.Number & ": " & Err.Description, vbCritical, "EO Update Error"
            strMsg = Err.Number & ": " & Err.Description
            modMsgBox.OKMsg = "OK"
            modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        End If
    Else
        gobjLog.LineTextToLog Err.Source & Err.Number & Err.Description
        'MsgBox Err.Number & ": " & Err.Description, vbCritical, "EO Update Error"
        strMsg = Err.Number & ": " & Err.Description
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    End If
End Function

Public Function genNumFormat(NumberType As String, Num As String) As String
Dim rsEOStruct As New ADODB.Recordset
Dim strNumber As String
Dim strConcatNum As String
Dim sngLen As Single

strConcatNum = ""
Set rsEOStruct = gdbConn.Execute("SELECT * FROM tblDocStruct where StructID = '" & NumberType & "' order by seq")

With rsEOStruct
While Not .EOF
    strNumber = ""
    sngLen = IIf(IsNull(![length]), 0, ![length])
    If ![Type] = "P" Then
        If IsNull(frmOthSvcs.datProducts.Recordset![TktPrefix]) Or frmOthSvcs.datProducts.Recordset![TktPrefix] = "" Then
            If Not IsNull(![DefValue]) Then
                strNumber = strNumber & Mid(![DefValue], 1, sngLen)
            End If
        Else
            strNumber = strNumber & Mid(frmOthSvcs.datProducts.Recordset![TktPrefix], 1, sngLen)
        End If
    ElseIf ![Type] = "D" Then
        If Not IsNull(![Format]) Then
            strNumber = strNumber & Mid(Format(Now, ![Format]), 1, sngLen)
        End If
    ElseIf ![Type] = "N" Then
        strNumber = strNumber & Mid(IIf(IsNumeric(Num), Num, IIf(IsNull(![DefValue]), 0, ![DefValue])), 1, sngLen)
    End If
    
    If ![PadZero] = True Then
        strNumber = Format(strNumber, String(sngLen, "0"))
    End If
    strConcatNum = strConcatNum & strNumber
    .MoveNext
Wend
.Close
End With

Set rsEOStruct = Nothing
genNumFormat = strConcatNum
End Function
Private Function getTicketNo(productType As String, ProductCode As String) As String
Dim rsLength As ADODB.Recordset
Dim intLength As Integer
Dim strZero As String
Dim intI As Integer
Set rsLength = gdbConn.Execute("Select Length from tbldocstruct where StructID='EO'")
While Not rsLength.EOF

intLength = intLength + rsLength![length]

rsLength.MoveNext
Wend
With gobjEO

For intI = 1 To intLength
strZero = strZero & "0"
Next

    'If productType = "MS" And (ProductCode = "35" Or ProductCode = "50") Then
    'JY – V1.2.2 20110322 – CR54 - Agent Ware Integration
    If productType = "CT" And .WebFareApplied = True Then
       getTicketNo = "/TK" & .EONumber & "/"
    ElseIf productType = "MS" And .TktType = True Then
            If frmOSMisc.optBSPConsol = True Or frmOSMisc.optConsol = True Then
                '02062005
                'getTicketNo = "/TK" & IIf(.TicketNumber = "0000", Format(.TicketNumber, strZero & "/"), fConvertTkTNo(.TicketNumber))
                'modified on 27/1/06: CS Changes
                'getTicketNo = "/TK" & fConvertTkTNo(.TicketNumber)
                'getTicketNo = "/TK" & .TicketNumber
            'ElseIf frmOSMisc.optConsol = True Then
                'Modified 15Jun
                'getTicketNo = "/TK" & Format(.TicketNumber, strZero) & "/" ', 3), "000") '& Format(IIf(.EONumber = "", strZero, .EONumber), strZero & "/")
                getTicketNo = "/TK" & .TicketNumber & "/"
            Else
                getTicketNo = "/TK" & IIf(.TicketNumber = "0000", Format(.TicketNumber, strZero), "FF" & Format(.TicketNumber, "00")) & "/"
            End If
    '02062005
    ElseIf productType = "CT" And (ProductCode = "35" Or ProductCode = "50") Then
       getTicketNo = "/TK" & Format(Left(.TicketNumber, 3), "000") & Format(IIf(.EONumber = "", strZero, .EONumber), strZero & "/")
    '02062005
    ElseIf productType = "BT" And (ProductCode = "35" Or ProductCode = "50") Then
           'getTicketNo = "/TK" & IIf(.TicketNumber = "", "", fConvertTkTNo(.TicketNumber))
            getTicketNo = "/TK" & IIf(.TicketNumber = "", "", .TicketNumber & IIf(.ConjunctTicket <> "", "-" & .ConjunctTicket, "") & "/")
    ElseIf ProductCode = "35" Or ProductCode = "50" Or ProductCode = "70" Then
            '02062005
            'If .TicketNumber = "0000" Then
            '    getTicketNo = "/TK" & Format(.TicketNumber, strZero & "/")
            'ElseIf .TicketNumber = "**CT" Then
            '    getTicketNo = "/TK" & "**CT" & "/"
            'Else
            '    getTicketNo = "/TK" & fConvertTkTNo(.TicketNumber)
            'End If
            getTicketNo = "/TK" & Format(.EONumber, strZero & "/")
    '02062005
    ElseIf productType = "CT" Then
            getTicketNo = "/TK" & Format(Left(.TicketNumber, 3), "000") & Format(IIf(.EONumber = "", strZero, .EONumber), strZero & "/")
    '02062005
    ElseIf productType = "BT" Then
           'modified on 15jun
            getTicketNo = "/TK" & IIf(.TicketNumber = "", "", .TicketNumber & IIf(.ConjunctTicket <> "", "-" & .ConjunctTicket, "") & "/")
    Else
            '02062005
            'getTicketNo = "/TK" & IIf(.TicketNumber = "", "", Format(.TicketNumber, strZero)) & "/"
            getTicketNo = "/TK" & Format(IIf(.EONumber = "", strZero, .EONumber), strZero) & "/"
    End If

End With

End Function

Private Function getMSLineforMI() As String
Dim freefields As String
Dim strMIData As String
Dim strFF() As String
Dim strTemp As String
Dim lngC As Long

    'Client specific MI data
    If isLoaded("frmClientMI") Then
        freefields = freefields & "/" & frmClientMI.getMSXFreeFields()
    End If
    
    strMIData = ""
    If freefields <> "" Then
        strFF = Split(freefields, "/")
        For lngC = LBound(strFF) To UBound(strFF)
            If strFF(lngC) <> "" Then
               'strMIData = strMIData & "/FF" & strFF(lngC)
               strTemp = Mid(strFF(lngC), 1, InStr(1, strFF(lngC), "-") - 1)
               strMIData = strMIData & IIf(IsNumeric(strTemp), "/FF", "/") & strFF(lngC)
            End If
        Next
    End If
    getMSLineforMI = strMIData
End Function

'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
'Private Function getSegmentSelected(pdtcode As String) As String
Private Function getSegmentSelected(pdtcode As String, pdtType As String) As String
    Dim strTemp As String
    Dim intI As Integer
    strTemp = ""
    'JY – V1.2.3 20110420 – IR12 - Missing Segment Number /SG for Product Code 91
    'If pdtcode = "00" Or pdtcode = "02" Or pdtcode = "01" Then
    If pdtType = "CT" Or pdtType = "BT" Then
            For intI = 0 To frmOSAirTkt.lstFlights.ListCount - 1
                If frmOSAirTkt.lstFlights.Selected(intI) Then
                    strTemp = strTemp & Format(Left(frmOSAirTkt.lstFlights.List(intI), InStr(frmOSAirTkt.lstFlights.List(intI), ".") - 1), "00")
                End If
                
            Next
    ElseIf pdtcode = "35" Or pdtcode = "50" Then
    
             For intI = 0 To frmOSMisc.lstFlights.ListCount - 1
                If frmOSMisc.lstFlights.Selected(intI) Then
                    strTemp = strTemp & Format(Left(frmOSMisc.lstFlights.List(intI), InStr(frmOSMisc.lstFlights.List(intI), ".") - 1), "00")
                End If
                
            Next
    End If
    
    getSegmentSelected = "/SG" & strTemp
    
End Function
'Preethi - V1.2.12 20120416 - CR127 - Quick Wins - Consolidator Ticket (HKSG)
Private Sub GetAirlineVendor(ByVal strSelectedAirSegment As String, ByRef strVendor As String, ByRef strVendorLocator As String)

Dim strVendorLocatorCode As String
Dim strVendorCode As String
Dim intI As Integer

strVendorLocatorCode = ""
strVendor = ""
strVendorLocator = ""
strVendorCode = ""

If Len(strSelectedAirSegment) > 0 Then
  
    For intI = 1 To Len(strSelectedAirSegment)
        If intI <> Len(strSelectedAirSegment) Then
          If IsNumeric(Mid(strSelectedAirSegment, intI, 1)) And Mid(strSelectedAirSegment, intI + 1, 1) = "." Then
             If intI + 1 <> Len(strSelectedAirSegment) Then
                strVendorCode = Mid(strSelectedAirSegment, intI + 3, 2)
                If strVendorCode <> "" Then
                   strVendor = strVendor & IIf(strVendor = "", "", ";") & strVendorCode
                End If
                strVendorCode = MatchCarrierVendor(strVendorCode)
                strVendorLocatorCode = GetAirlineVendorLocator(strVendorCode)
                strVendorLocator = strVendorLocator & IIf(strVendorLocator = "", "", ";") & strVendorLocatorCode
             End If
          End If
        End If
    Next intI
End If

End Sub
'Preethi - V1.2.12 20120416 - CR127 - Quick Wins - Consolidator Ticket (HKSG)
Private Function GetAirlineVendorLocator(ByVal strAirVendor As String) As String

Dim lngC As Long
Dim strLocator As String
Dim strTemp As String

strLocator = ""

Set gobjPNR = New CWT_GalileoPNR3.PNR
        With gobjPNR
            Call .loadPNR
            For lngC = 1 To .VendorRecLocCount
                strTemp = .VendorRecLoc(lngC).Vendor
                If (strAirVendor = strTemp) Then
                    strLocator = Trim(.VendorRecLoc(lngC).RecLoc)
                End If
            Next
        End With


GetAirlineVendorLocator = strLocator

End Function

'Preethi - V1.2.12 20120416 - CR127 - Quick Wins - Consolidator Ticket (HKSG)
Private Function MatchCarrierVendor(ByVal strAirVendor As String) As String

Dim strSQL As String
Dim rsEO As New ADODB.Recordset
Dim strMsg As String

Dim strVLCarrier As String

strVLCarrier = ""
strLocator = ""

strSQL = "SELECT * FROM [tblVLCarrier] WHERE [Carrier] = '" & strAirVendor & "'"
Set rsEO = gdbConn.Execute(strSQL)

With rsEO
    If Not .EOF Then
        strVLCarrier = ![VLCarrier]
    End If
End With
rsEO.Close

strVLCarrier = Trim(strVLCarrier)
If Len(strVLCarrier) > 0 Then
    MatchCarrierVendor = strVLCarrier
Else
    MatchCarrierVendor = strAirVendor
End If


End Function

'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration - EO Email Default Sender Name
Public Function getDefaultEmailSenderName(strAgencyName As String) As String
    
    Dim rsEmail As ADODB.Recordset
    Dim strDefaultEmail As String
    Dim strSQL As String
              
    strSQL = "SELECT DefaultSenderName FROM tblDefaultEmail WHERE AgencyName ='" & strAgencyName & "'"
    Set rsEmail = gdbEitinConn.Execute(strSQL)
    
    If Not rsEmail.EOF Then
        strDefaultEmail = rsEmail!DefaultSenderName
    End If
    
    rsEmail.Close
    Set rsEmail = Nothing
    getDefaultEmailSenderName = strDefaultEmail
        
End Function

'ZhiSam - V1.2.24 20140116 - CR 304 - JTB Integration - AgencyName Error Checking - Return true if AgencyName exist
Public Function bolAgencyNameCheck(strAgencyName As String) As Boolean

    Dim bolStatus As Boolean
    Dim strMsg As String
    
     If IsNull(strAgencyName) Or strAgencyName = "" Then
        bolStatus = False
        strMsg = "Error: Client Agency Name not found!"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        
     Else
        bolStatus = True
     End If
     
    bolAgencyNameCheck = bolStatus

End Function
