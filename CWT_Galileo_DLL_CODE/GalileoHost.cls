VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GalileoHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mobjSQ As HostAccess.StructuredQuery
Private mobjTE As HostAccess.TerminalEmulation
Private Const mstrID As String = "<Application><VendorId>XMDL</VendorId><VendorType>G</VendorType><SourceId>CARWPR</SourceId><SourceType>G</SourceType></Application>"
Private mcolSimilarNames As Collection
Private mstrAgentSine As String
Private mstrAgentName As String
Private mstrProfile As String
Private mstrAgentPCC As String
Private mstrAgentQueue As String
Private mstrAgentGACode As String
Private mstrAgentDIV As String
Private mstrWorkAreas As String
Private mstrActiveArea As String
Private mstrSession As String 'local copy
Public Event GotSession()
Private mbolConnected As Boolean 'local copy
'Private gobjLog As CWT_AppLog.AppLog
Public Event ConnectedToHost()

Public Property Get AgentSine() As String

If mstrAgentSine = "" Then Call CheckWorkAreas
AgentSine = mstrAgentSine

End Property
Public Property Get AgentDIV() As String

If mstrAgentDIV = "" Then Call CheckWorkAreas
AgentDIV = mstrAgentDIV

End Property

Public Property Get AgentName() As String

If mstrAgentName = "" Then Call CheckWorkAreas
AgentName = mstrAgentName

End Property
Public Property Get AgentProfile() As String

If mstrProfile = "" Then Call CheckWorkAreas
AgentProfile = mstrProfile

End Property

Public Property Get AgentPCC() As String

If mstrAgentPCC = "" Then Call CheckWorkAreas
AgentPCC = mstrAgentPCC

End Property

Public Property Get AgentQueue() As String

If mstrAgentQueue = "" Then Call CheckWorkAreas
AgentQueue = mstrAgentQueue

End Property

Public Property Get AgentGACode() As String

If mstrAgentGACode = "" Then Call CheckWorkAreas
AgentGACode = mstrAgentGACode

End Property

Public Property Get WorkAreas() As String
    Call CheckWorkAreas
    WorkAreas = mstrWorkAreas
End Property

Public Function SendQuery(ByVal RequestText As String, ByVal Transaction As String, ByVal CallingModuleName As String, ByVal CallingProcedureName As String) As String
    Dim strReqFileName As String
    Dim strResFileName As String
    Dim objXMLStream As Scripting.TextStream
    
    With gobjLog
        .ModuleName = CallingModuleName
        .ProcedureName = CallingProcedureName
        If .LogOpen Then
            .EventToLog "Beg Transmit XML Trans " & Transaction & " in " & .ModuleName & "." & .ProcedureName
        End If
        
    SendQuery = mobjSQ.ExecuteXMLQuery(RequestText, mstrID)
        If .LogOpen Then
            .EventToLog "End Transmit XML Trans " & Transaction
            'strReqFileName = Transaction & "_" & Format(Now(), "ddmmmyyhhnnss") & ".xml"
            'strResFileName = gobjLog.FilePath & "\" & "RES_" & strReqFileName
            'strReqFileName = gobjLog.FilePath & "\" & "REQ_" & strReqFileName
            '.XMLFileToLog strReqFileName, RequestText
            '.XMLFileToLog strResFileName, SendQuery
        End If
    End With
            
End Function

Public Property Get ActiveArea() As String
    Call CheckWorkAreas
    ActiveArea = mstrActiveArea
End Property

Public Function TerminalEntry(ByVal Entry As String, Optional ByVal FullResponse As Boolean = False) As String
Dim strTemp() As String
Dim strResp As String
Dim lngC As Long
Dim intRetry As Integer
Dim strActualEntry As String
Dim intResponse As Integer
Dim intCount As Integer
Dim STRSQL As String
Dim strMsg As String
On Error GoTo ErrMakeEntry
If Trim(Entry) = "" Then Exit Function
strActualEntry = ""
strActualEntry = Entry
Entry = "<FORMAT><![CDATA[" & Entry & "]]></FORMAT>"
'Entry = "<FORMAT>" & Entry & "</FORMAT>"
'If mobjTE Is Nothing Then Set mobjTE = New HostAccess.TerminalEmulation
    If Not mobjTE Is Nothing Then Set mobjTE = Nothing
    Set mobjTE = New HostAccess.TerminalEmulation

With gobjLog
    If .LogOpen Then
        .EventToLog "BEG TERMINAL ENTRY ", 1
        .LineTextToLog "ENTRY>" & Entry
        .EventToLog "TRANSMIT"
    End If
End With
mobjTE.MakeEntry Entry

Call Sleep(300)

'modified on 07/09/04: wait until response received
intRetry = 0
While mobjTE.NumResponseLines = 0 And intRetry < 5
    intRetry = intRetry + 1
    With gobjLog
        If .LogOpen Then
            .LineTextToLog "WAITING FOR GDS RESPONSE... " & intRetry & " TRY"
        End If
    End With
    
    Sleep (1000)
Wend

'Call Sleep(500)

With gobjLog
    If .LogOpen Then
        .EventToLog "END TRANSMIT"
    End If
End With


If FullResponse Then
    Do While (mobjTE.more)
        mobjTE.GetMore 1, 1
    Loop
End If

strResp = ""
For lngC = 0 To mobjTE.NumResponseLines - 2
    strResp = strResp & mobjTE.ResponseLine(lngC) _
        & IIf(InStr(mobjTE.ResponseLine(lngC), "<CARRIAGE_RETURN/>"), "", vbCrLf)
Next
strResp = strResp & mobjTE.ResponseLine(lngC)


If InStr(strResp, ")<SOM/>") Then
    strTemp = Split(strResp, ")<SOM/>")
    strResp = ""
    For lngC = LBound(strTemp) To UBound(strTemp)
        strResp = strResp & strTemp(lngC)
    Next
End If

If InStr(strResp, "<SOM/>") Then
    strTemp = Split(strResp, "<SOM/>")
    strResp = ""
    For lngC = LBound(strTemp) To UBound(strTemp)
        strResp = strResp & strTemp(lngC)  '& Chr(187)
    Next
End If

If InStr(strResp, "<EOM/>") Then
    strTemp = Split(strResp, "<EOM/>")
    strResp = ""
    For lngC = LBound(strTemp) To UBound(strTemp) - 1
        strResp = strResp & strTemp(lngC) & Chr(171)
    Next
    strResp = strResp & strTemp(lngC)
End If

If InStr(strResp, "<PILLOW/>") Then
    strTemp = Split(strResp, "<PILLOW/>")
    strResp = ""
    For lngC = LBound(strTemp) To UBound(strTemp) - 1
        strResp = strResp & strTemp(lngC) & "@"
    Next
    strResp = strResp & strTemp(lngC)
End If

If InStr(strResp, "<TABSTOP/>") Then
    strTemp = Split(strResp, "<TABSTOP/>")
    strResp = ""
    For lngC = LBound(strTemp) To UBound(strTemp) - 1
        strResp = strResp & strTemp(lngC) & Chr(183)
    Next
    strResp = strResp & strTemp(lngC)
End If


If InStr(strResp, "<CARRIAGE_RETURN/>") Then
    strTemp = Split(strResp, "<CARRIAGE_RETURN/>")
    strResp = ""
    For lngC = LBound(strTemp) To UBound(strTemp) - 2
        strResp = strResp & strTemp(lngC) & vbCrLf
    Next
    
    strResp = strResp & strTemp(lngC)
End If

With gobjLog
    If .LogOpen Then
    .EventToLog "BEG FORMATTED RESPONSE> ", 1, 0
    .LineTextToLog strResp
    .LineTextToLog ">>>END FORMATTED RESPONSE> ", 0, 1
    .EventToLog ">>>BEG XML RESPONSE>"
    .LineTextToLog mobjTE.responseXML
    .LineTextToLog ">>>END XML RESPONSE>", 0, 1
    .EventToLog "END TERMINAL ENTRY"
    .LineTextToLog String(64, "-"), 0, 1
End If
End With

TerminalEntry = Trim(strResp)



Exit Function

ErrMakeEntry:
   
    If Err.Number = -2147467259 Then
      
        gobjLog.LineTextToLog "Error:" & Err.Number & Err.Description
        gobjLog.LineTextToLog "ErrorEntry:" & Entry
        
        If gbolTEErrorLog = True Then
            STRSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
             "VALUES('" & Now & "','" & Err.Number & "','" & Err.Description & " Resp. Len=" & Len(mobjTE.ResponseLine(0)) & " More=" & mobjTE.more & "','" & Entry & "', " & _
             "'" & AgentName() & "')"
            gdbConn.Execute STRSQL
        End If

            TerminalEntry = CONNECTION_FAIL
            strMsg = CONNECTION_FAIL & " : " & Chr(13) & _
                     "Entry: " & """" & strActualEntry & """" & " is aborted." & Chr(13) & _
                     "Do you want to resend the entry?"
            modMsgBox.RETRYMsg = "Retry"
            modMsgBox.CANCELMsg = "Cancel"
            intResponse = modMsgBox.sMsgBox(gVPMDIHwnd, strMsg, vbQuestion + vbApplicationModal + vbRetryCancel + vbDefaultButton1, "CWT Desktop - Error")
            'intResponse = MsgBox(CONNECTION_FAIL & " : " & Chr(13) & _
            "Entry: " & """" & strActualEntry & """" & " is aborted." & Chr(13) & _
            "Do you want to resend the entry?", vbApplicationModal + vbRetryCancel + vbExclamation + vbDefaultButton1, "CWT Galileo Host")
            If intResponse = vbRetry Then
                    Sleep (1000)
                    gobjLog.LineTextToLog "RetryEntry:" & Entry
                    intCount = intCount + 1
                    gobjLog.LineTextToLog "Retry Count:" & intCount
                    Call TerminalEntry(strActualEntry, FullResponse)
                    
            Else
                    strMsg = "Entry: " & """" & strActualEntry & """" & " is aborted." & Chr(13) & _
                             "Please toggle to focalpoint to enter/paste(Crtl+V) the entry before continue."
                    modMsgBox.OKMsg = "OK"
                    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
                    'MsgBox "Entry: " & """" & strActualEntry & """" & " is aborted." & Chr(13) & _
                    "Please toggle to focalpoint to enter/paste(Crtl+V) the entry before continue.", vbOKOnly + vbExclamation, "CWT Galileo Host"
                    gobjLog.LineTextToLog "CancelEntry:" & Entry
                    gobjLog.LineTextToLog "Total Retry Count:" & intCount
                    Clipboard.Clear
                    Clipboard.SetText (strActualEntry)
            End If
       
    Else
        If gbolTEErrorLog = True Then
            STRSQL = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry,Agent) " & _
             "VALUES('" & Now & "','" & Err.Number & "','" & Err.Description & " Resp. Len=" & Len(mobjTE.ResponseLine(0)) & " More=" & mobjTE.more & "','" & Entry & "', " & _
             "'" & AgentName() & "')"
            gdbConn.Execute STRSQL
        End If
        
        Err.Raise Err.Number, "Terminal Entry", Err.Description
        gobjLog.LineTextToLog "Error:" & Err.Number & Err.Description
        gobjLog.LineTextToLog "ErrorEntry:" & Entry
        
    End If
'''
End Function

Public Function EndPNR(Optional ReceivedFrom As String, Optional ByVal ReRetrieve As Boolean = False) As Boolean

Dim strRequest As String
Dim strResponse As String

strRequest = ""
strRequest = strRequest & "<PNRBFEnd_9_0>"
strRequest = strRequest & "    <EndTransactionMods>"
strRequest = strRequest & "        <EndTransactRequest>"
strRequest = strRequest & "            <ETInd>" & IIf(ReRetrieve = True, "R", "E") & "</ETInd>"
strRequest = strRequest & "            <RcvdFrom>" & ReceivedFrom & "</RcvdFrom>"
strRequest = strRequest & "        </EndTransactRequest>"
strRequest = strRequest & "    </EndTransactionMods>"
strRequest = strRequest & "</PNRBFEnd_9_0>"

strResponse = SendQuery(strRequest, "PNRBFEnd_9_0", "GalileoHost", "EndPNR")

Select Case TestETResponse(strResponse)
    Case "E"
        EndPNR = False
    Case "W"    'try again
        strRequest = ""
        strRequest = strRequest & "<PNRBFEnd_9_0>"
        strRequest = strRequest & "    <EndTransactionMods>"
        strRequest = strRequest & "        <EndTransactRequest>"
        strRequest = strRequest & "            <ETInd>" & IIf(ReRetrieve = True, "R", "E") & "</ETInd>"
        ' Recieved from is entered on first try - entering again would cause an error
        strRequest = strRequest & "        </EndTransactRequest>"
        strRequest = strRequest & "    </EndTransactionMods>"
        strRequest = strRequest & "</PNRBFEnd_9_0>"

        strResponse = SendQuery(strRequest, "PNRBFEnd_9_0", "GalileoHost", "EndPNR")

        Select Case TestETResponse(strResponse)
            Case "E"
                EndPNR = False
            Case "W"    'try one last time
                strResponse = SendQuery(strRequest, "PNRBFEnd_9_0", "GalileoHost", "EndPNR")

                If TestETResponse(strResponse) = "S" Then
                    EndPNR = True
                Else
                    EndPNR = False
                End If
            Case "S"
                EndPNR = True
            Case Else
                EndPNR = False
        End Select

    Case "S"
        EndPNR = True
    Case Else
        EndPNR = False
    
End Select



End Function

Private Function TestETResponse(Response As String) As String
Dim xmldomResponse As MSXML2.DOMDocument

If InStr(1, Response, "<CCRuleExecute") Then
    If InStr(1, Response, "RULE NOT ATTACHED MANDATORY") > 0 _
        Or InStr(1, Response, "RULE - END MANDATORY") > 0 Then
            
        TestETResponse = "E"
    Else
        TestETResponse = "W"
    End If

ElseIf InStr(1, Response, "<PNRBFRetrieve") Then
    TestETResponse = "S"
ElseIf InStr(1, Response, "<ErrSeverityInd") > 0 Then
    Set xmldomResponse = New MSXML2.DOMDocument
    xmldomResponse.loadXML Response
    TestETResponse = Left(xmldomResponse.getElementsByTagName("ErrSeverityInd").Item(0).Text & "S", 1)
Else
    TestETResponse = "E"
End If

End Function


Public Function MoveProfile(HQPCC As String, CompanyTitle As String, Optional TravelerLastName As String = "" _
    , Optional TravelerFirstName As String = "") As String
    
Dim strRequest As String
Dim strResponse As String
Dim bolTrav As Boolean
Dim xmldomPro As MSXML2.DOMDocument
Dim xmlnlSNL As MSXML2.IXMLDOMNodeList
Dim lngC1 As Long
Dim strTemp As String

If TravelerLastName <> "" Then bolTrav = True

strRequest = "<ClientFile_1_0>"
strRequest = strRequest & "    <ClientFileMods>"
strRequest = strRequest & "        <ClientFileMoveHeaderMods>"
strRequest = strRequest & "            <CRSID>1G</CRSID>"
strRequest = strRequest & "            <PCC>" & HQPCC & "</PCC>"
strRequest = strRequest & "            <BusinessTitle>" & UCase(CompanyTitle) & "</BusinessTitle>"
If bolTrav Then strRequest = strRequest & "            <PersonalTitle>" & "<![CDATA[" & UCase(Trim(TravelerLastName & " " & TravelerFirstName)) & "]]>" & "</PersonalTitle>"
strRequest = strRequest & "            <MultipleInd>Y</MultipleInd>"
strRequest = strRequest & "            <SelectInd>N</SelectInd>"
strRequest = strRequest & "            <MergeInd>Y</MergeInd>"
strRequest = strRequest & "        </ClientFileMoveHeaderMods>"
strRequest = strRequest & "    </ClientFileMods>"
strRequest = strRequest & "</ClientFile_1_0>"

strResponse = SendQuery(strRequest, "ClientFile_1_0", "", "")
Set xmldomPro = New MSXML2.DOMDocument
xmldomPro.loadXML strResponse
If bolTrav And InStr(strResponse, "ClientFileDoubleTitleList") Then
    'Set xmldomPro = New MSXML2.DOMDocument
    
    Set xmlnlSNL = xmldomPro.selectNodes("//ClientFileDoubleTitleList[BusinessTitle='" & Trim(UCase(CompanyTitle)) & "']")
    If xmlnlSNL.length = 0 Then
        MoveProfile = "NTP"
        Exit Function
    Else
        Set mcolSimilarNames = New Collection
        MoveProfile = "SNL"
        For lngC1 = 0 To xmlnlSNL.length - 1
            strTemp = xmlnlSNL.Item(lngC1).childNodes.Item(1).Text
            mcolSimilarNames.Add strTemp, strTemp
        Next
    End If
ElseIf InStr(strResponse, "<ClientFilePositiveAck") Then
    MoveProfile = "PRO"
Else
  
    If xmldomPro.selectNodes("//ErrText/Text").length > 0 Then
        If Trim(xmldomPro.selectNodes("//ErrText/Text").Item(0).Text) <> "" Then
            MoveProfile = xmldomPro.selectNodes("//ErrText/Text").Item(0).Text
        Else
            MoveProfile = "UHR"
        End If
    Else
        MoveProfile = "UHR"
    End If
End If

End Function

Public Function SimNameList(Index As Long) As String
    SimNameList = mcolSimilarNames(Index)
End Function

Public Function SimNameListCount() As Long
    SimNameListCount = mcolSimilarNames.Count
End Function

Public Sub SignOn(Sine As String, Pwd As String)
Dim strRequest As String
Dim strResponse As String
Dim strID As String
Dim strMsg As String

strRequest = ""
strRequest = strRequest & "<SessionSignOn_1_0>"
strRequest = strRequest & "    <SessionMods>"
strRequest = strRequest & "        <SignOn>"
strRequest = strRequest & "            <UserID>" & Sine & "</UserID>"
strRequest = strRequest & "            <Pwd>" & Pwd & "</Pwd>"
strRequest = strRequest & "            <Keyword></Keyword>"
strRequest = strRequest & "        </SignOn>"
strRequest = strRequest & "    </SessionMods>"
strRequest = strRequest & "</SessionSignOn_1_0>"

strID = mstrID & "<User><UserId>" & Sine & "</UserId><Pseudo>781P</Pseudo></User>"
strResponse = mobjSQ.ExecuteXMLQuery(strRequest, mstrID)
If InStr(1, strResponse, "SIGN-ON COMPLETE/") = 0 Then
    'MsgBox "Sign on was not successful!"
    strMsg = "Sign on was not successful!"
    modMsgBox.OKMsg = "OK"
    modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
    Exit Sub
End If

'If gstrsvcbur Then
    strRequest = "<FORMAT>SEM/" & "781P" & "/AG</FORMAT>"
    mobjTE.MakeEntry (strRequest)
    If InStr(1, mobjTE.responseXML, "PROCEED/") = 0 Then
        'MsgBox "Emulation was not successful!"
        strMsg = "Emulation was not successful!"
        modMsgBox.OKMsg = "OK"
        modMsgBox.sMsgBox gVPMDIHwnd, strMsg, vbOKOnly + vbDefaultButton1, "CWT Desktop - Error"
        Exit Sub
    End If
End Sub

Public Property Let Connected(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Connected = 5
    mbolConnected = vData
End Property

Public Property Get Connected() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Connected
    Connected = mbolConnected
End Property

Public Sub ConnectToHost()
End Sub

Public Property Let Session(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Session = 5
    mstrSession = vData
End Property

Public Property Get Session() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Session
    Session = mstrSession
End Property

Public Sub RetransmitLast()
End Sub

Public Function GetSession() As String
End Function

Private Sub Class_Initialize()
Set mobjSQ = New HostAccess.StructuredQuery
Set mobjTE = New HostAccess.TerminalEmulation

'Call pGetStartupValues
'pOpenDB

End Sub

Private Sub Class_Terminate()

Set mcolSimilarNames = Nothing
mobjSQ.Close
mobjTE.Close

Set mobjSQ = Nothing
Set mobjTE = Nothing

End Sub

Private Sub CheckWorkAreas()
Dim strRequest As String
Dim lngC As Long
Dim rsAgents As New ADODB.Recordset
Dim STRSQL As String
Dim xmlResponse As MSXML2.DOMDocument
Dim xmlnl As MSXML2.IXMLDOMNodeList
Dim strsql2 As String
Dim intActiveArea As Integer
Dim xmlAreaInfo As MSXML2.IXMLDOMNodeList
Dim intActiveIndCount As Integer

Dim strAAAArea As String
Dim strSONCity As String
Dim strAgntID As String
Dim strAgntInitials As String
Dim strAgncyPCC As String
Dim bolFound As Boolean

strRequest = "<SessionWorkAreaInfo_1_0>"
strRequest = strRequest & "<SessionMods>"
strRequest = strRequest & "<AreaInfoReq />"
strRequest = strRequest & "</SessionMods>"
strRequest = strRequest & "</SessionWorkAreaInfo_1_0>"

Set xmlResponse = New MSXML2.DOMDocument

xmlResponse.loadXML SendQuery(strRequest, "SessionWorkAreaInfo_1_0", "GalileoHost", "CheckWorkAreas")

Set xmlnl = xmlResponse.selectNodes("//AAAAreaAry/AAAAreaInfo/ActiveInd|//AAAAreaAry/AAAAreaInfo/PNRDataInd")

If xmlnl.length = 0 Then
    mstrWorkAreas = "XXXXX"
    mstrAgentSine = ""
    mstrAgentPCC = ""
    Exit Sub
End If

intActiveArea = 0
mstrWorkAreas = ""
intActiveIndCount = 0
For lngC = 0 To xmlnl.length - 1
    Select Case xmlnl(lngC).nodeName
       Case "ActiveInd"
         If UCase(xmlnl(lngC).Text) = "Y" Then
            intActiveArea = intActiveIndCount
         End If
         intActiveIndCount = intActiveIndCount + 1
       Case "PNRDataInd"
         If UCase(xmlnl(lngC).Text) = "Y" Then
            mstrWorkAreas = mstrWorkAreas & "Y"
         Else
            mstrWorkAreas = mstrWorkAreas & "N"
         End If
    End Select
Next


Set xmlAreaInfo = xmlResponse.getElementsByTagName("AAAAreaInfo")

With xmlAreaInfo.Item(intActiveArea)
For lngC = 0 To .childNodes.length - 1
   Select Case .childNodes(lngC).nodeName
      Case "AAAArea"
         strAAAArea = .childNodes(lngC).Text
      Case "SONCity"
         strSONCity = .childNodes(lngC).Text
      Case "AgntID"
         strAgntID = .childNodes(lngC).Text
      Case "AgntInitials"
         strAgntInitials = .childNodes(lngC).Text
      Case "AgncyPCC"
         strAgncyPCC = .childNodes(lngC).Text
   End Select
Next
End With

'mstrAgentSine = strAgntInitials
If IsNumeric(strAgntID) Then
   mstrAgentSine = Int(strAgntID)
Else
   mstrAgentSine = Mid(strAgntID, 2)
End If
mstrActiveArea = strAAAArea

If IsNumeric(mstrAgentSine) Then
   mstrAgentPCC = strAgncyPCC
Else
   mstrAgentPCC = strSONCity
End If

STRSQL = "SELECT * FROM tblAgents WHERE [Sine] = '" & mstrAgentSine & "' "
If IsNumeric(mstrAgentSine) = False Then
   STRSQL = STRSQL & "and PCC = '" & mstrAgentPCC & "'"
End If
Set rsAgents = gdbConn.Execute(STRSQL)
If Not rsAgents.EOF Then
   bolFound = True
   mstrAgentName = UCase(rsAgents![AgentName]) & ""
   mstrProfile = UCase(rsAgents![ProfileName]) & ""
   mstrAgentQueue = rsAgents![Queue] & ""
   mstrAgentGACode = UCase(rsAgents!AgentGACode) & ""
   mstrAgentDIV = UCase(rsAgents!div) & ""
Else
   bolFound = False
End If

If bolFound = False Then
   STRSQL = "SELECT * FROM tblAgents WHERE [Sine] = '" & mstrAgentSine & "' "
   Set rsAgents = gdbConn.Execute(STRSQL)
   If Not rsAgents.EOF Then
      mstrAgentName = UCase(rsAgents![AgentName]) & ""
      mstrProfile = UCase(rsAgents![ProfileName]) & ""
      mstrAgentQueue = rsAgents![Queue] & ""
      mstrAgentGACode = UCase(rsAgents!AgentGACode) & ""
      mstrAgentDIV = UCase(rsAgents!div) & ""
   End If
End If

rsAgents.Close
Set rsAgents = Nothing

End Sub


'Private Sub CheckWorkAreas()
'Dim strRequest As String
'Dim lngC As Long
''Dim dbTPro As DAO.Database
'Dim rsAgents As New adodb.Recordset
'Dim strsql As String
'Dim xmlResponse As MSXML2.DOMDocument
'Dim xmlNL As MSXML2.IXMLDOMNodeList
'Dim strsql2 As String
'strRequest = ""
'strRequest = strRequest & "<SessionEmulation_1_0>"
'strRequest = strRequest & "    <SessionMods>"
'strRequest = strRequest & "        <AreaInfoReq />"
'strRequest = strRequest & "    </SessionMods>"
'strRequest = strRequest & "</SessionEmulation_1_0>"
'
'Set xmlResponse = New MSXML2.DOMDocument
'
'xmlResponse.loadXML SendQuery(strRequest, "SessionEmulation_1_0", "GalileoHost", "CheckWorkAreas")
'
'Set xmlNL = xmlResponse.selectNodes("//AAAAreaAry/AAAAreaInfo/AgntInitials" _
'    & "|//AAAAreaAry/AAAAreaInfo/AgncyPCC" _
'    & "|//AAAAreaAry/AAAAreaInfo/ActiveInd|//AAAAreaAry/AAAAreaInfo/PNRDataInd")
'If xmlNL.length = 0 Then
'    mstrWorkAreas = "XXXXX"
'    mstrAgentSine = ""
'    'mstrAgentPCC = ""
'    Exit Sub
'End If
'
'mstrAgentSine = xmlNL(1).Text
'
'' strsql = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry) " & _
''            "VALUES('" & Now & "','88','" & xmlNL(1).Text & " " & xmlNL(2).Text & "','LOG1')"
''            gdbConn.Execute strsql
'
'
''If Len(xmlNL(2).Text) < 5 Then mstrAgentPCC = String(4 - Len(xmlNL(2).Text), "0") & xmlNL(2).Text
'
''If mstrAgentPCC = "" And gbolAgentPCCLog Then
'' strsql = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry) " & _
''            "VALUES('" & Now & "','SessionEmulation_1_0','" & xmlNL(2).Text & "','SessionEmulation_1_0')"
''            gdbConn.Execute strsql
''End If
'
''Set dbTPro = DAO.OpenDatabase(gstrTProDBSource)
'
'
'
'strsql = "SELECT * FROM tblAgents WHERE [Sine] = '" & mstrAgentSine & "'"
'
'
'
''If mstrAgentPCC <> "0000" Then
''    strsql = strsql & " AND [PCC]='" & mstrAgentPCC & "'"
''End If
'
'
''strsql2 = "INSERT into tblTEError (ErrorDate,ErrorNum,ErrorMsg,ErrorEntry) " & _
''            "VALUES('" & Now & "','88','" & Replace(strsql, "'", "''") & "','LOG2')"
''            gdbConn.Execute strsql2
'
'Set rsAgents = gdbConn.Execute(strsql)
''If rsAgents.RecordCount > 0 Then
'If Not rsAgents.EOF Then
'   mstrAgentName = UCase(rsAgents![AgentName]) & ""
'   mstrAgentPCC = UCase(rsAgents![PCC]) & ""
'   mstrAgentQueue = rsAgents![Queue] & ""
'   mstrAgentGACode = UCase(rsAgents!AgentGACode) & ""
'End If
'rsAgents.Close
''dbTPro.Close
'Set rsAgents = Nothing
''Set dbTPro = Nothing
'
'
'mstrWorkAreas = ""
'
'For lngC = 0 To 14 Step 3
'    If xmlNL(lngC).Text = "Y" Then mstrActiveArea = Chr(65 + (lngC / 3))
'    If xmlNL(lngC + 2).Text = "Y" Then
'        mstrWorkAreas = mstrWorkAreas & "Y"
'    Else
'        mstrWorkAreas = mstrWorkAreas & "N"
'    End If
'Next
'End Sub
    
Public Sub pGetStartupValues(ByRef Conn As ADODB.Connection, ByRef objLog As CWT_AppLog.AppLog, ByRef host As CWT_Galileo3.GalileoHost, ByVal VPHwnd As Long)
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
End Sub

Public Function IgnorePNR(Optional ByVal ReRetrieve As Boolean = False, Optional pnr As String) As Boolean

Dim strRequest As String
Dim strResponse As String

strRequest = ""
strRequest = strRequest & "<PNRBFManagement_11>"
If Not ReRetrieve Then

    strRequest = strRequest & "<IgnoreMods/>"
Else
      strRequest = strRequest & "<PNRBFRetrieveMods>"
        strRequest = strRequest & "<PNRAddr>"
            strRequest = strRequest & "<FileAddr></FileAddr>"
            strRequest = strRequest & "<CodeCheck></CodeCheck>"
            strRequest = strRequest & "<RecLoc>" & pnr & "</RecLoc>"
        strRequest = strRequest & "</PNRAddr>"
    strRequest = strRequest & "</PNRBFRetrieveMods>"

End If
strRequest = strRequest & "</PNRBFManagement_11>"

strResponse = SendQuery(strRequest, "PNRBFManagement_11", "GalileoHost", "IgnorePNR")

End Function

Public Function SignOff() As String
Dim strRequest As String
Dim strResponse As String

Dim strMsg As String
Dim xmldomSignOff As MSXML2.DOMDocument
Dim xmlnlSignOff As MSXML2.IXMLDOMNodeList
Dim intI As Integer
strRequest = ""
strRequest = strRequest & "<SessionSignOff_1_0>"
strRequest = strRequest & "<SessionMods>"
strRequest = strRequest & "<SignOff />"
strRequest = strRequest & "</SessionMods>"
strRequest = strRequest & "</SessionSignOff_1_0>"

Set xmldomSignOff = New MSXML2.DOMDocument

strResponse = mobjSQ.ExecuteXMLQuery(strRequest, mstrID)

xmldomSignOff.loadXML strResponse
Set xmlnlSignOff = xmldomSignOff.selectNodes("//SessionInfo/SecurityMessage/Text")
If (InStr(1, strResponse, "SIGN-OFF COMPLETE/") > 0 Or InStr(1, strResponse, "SIGN-ON") > 0) Then
    SignOff = "True"
    Exit Function
Else
    If xmlnlSignOff.length > 0 Then
        SignOff = xmlnlSignOff.Item(0).Text
    Else
        SignOff = strResponse
    End If
End If

End Function
 'Preethi - V1.2.4 20110614 - CR 76 - Change Validation Logic For EndPNR2
Public Function EndPNR2(Optional ReceivedFrom As String, Optional ByVal ReRetrieve As Boolean = False, Optional ByVal EndCount As Integer = 1) As String
'2311 Jen change terminal entry method,due to inaccurate xml response
'sometimes xml response with <PNRBFEnd_9_0><TransactionErrorCode><Domain>Unknown</Domain><Code>95</Code></TransactionErrorCode><EndTransaction><ErrorCode>0095</ErrorCode></EndTransaction></PNRBFEnd_9_0>
'but with terminal entry, no error response

Dim strRequest As String
Dim strResponse As String
Dim strErrMsg As String
Dim resxml As MSXML2.DOMDocument
Dim resxmlnode As MSXML2.IXMLDOMNodeList
Dim intI As Integer
Dim strResp() As String

For intI = 1 To EndCount

    strRequest = ""
    strRequest = strRequest & "<PNRBFEnd_9_0>"
    strRequest = strRequest & "    <EndTransactionMods>"
    strRequest = strRequest & "        <EndTransactRequest>"
    strRequest = strRequest & "            <ETInd>" & IIf(ReRetrieve = True, "R", "E") & "</ETInd>"
    strRequest = strRequest & "            <RcvdFrom>" & ReceivedFrom & "</RcvdFrom>"
    strRequest = strRequest & "        </EndTransactRequest>"
    strRequest = strRequest & "    </EndTransactionMods>"
    strRequest = strRequest & "</PNRBFEnd_9_0>"
    
    strResponse = SendQuery(strRequest, "PNRBFEnd_9_0", "GalileoHost", "EndPNR2")
    
Next intI

Set resxml = New MSXML2.DOMDocument

resxml.loadXML (strResponse)

'If InStr(strResponse, "EndTransactMessage") > 0 Then
If InStr(strResponse, "<TransactionErrorCode>") > 0 Then
    If InStr(strResponse, "EndTransactMessage") > 0 Then
            Set resxmlnode = resxml.selectNodes("//EndTransaction/EndTransactMessage/Text")
        
            For intI = 0 To resxmlnode.length - 1
                strErrMsg = strErrMsg & resxmlnode.Item(intI).Text & vbCrLf
            Next intI
            
            EndPNR2 = IIf(strErrMsg <> "", strErrMsg, strResponse) 'False
    Else 'unknown error code use terminal entry
          TerminalEntry "R." & ReceivedFrom
          If ReRetrieve Then
            strResponse = TerminalEntry("ER")
          Else
            strResponse = TerminalEntry("E")
          End If
          
          EndPNR2 = strResponse
          If InStr(strResponse, vbCrLf) > 0 Then
            strResp = Split(strResponse, vbCrLf)
            For intI = LBound(strResp) To UBound(strResp)
              If Len(strResp(intI)) > 8 Then
                If Mid(strResp(intI), 7, 1) = "/" Then
                  EndPNR2 = "True"
                  Exit For
                End If
              End If
            Next
          End If
          
    End If
Else
    EndPNR2 = "True"
    
End If
End Function
